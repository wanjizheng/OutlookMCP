"""Resumable local mirror. Source failures never authorize index deletion."""
import datetime as dt
import hashlib
import json
import sqlite3
from contextlib import closing
from . import attachments, rag, sync_source


def now():
    return dt.datetime.now(dt.UTC).isoformat()


def initialize():
    rag.initialize()
    with rag.connect(False) as c:
        c.executescript('''
CREATE TABLE IF NOT EXISTS native_sync_source(entry_id TEXT PRIMARY KEY,store_id TEXT NOT NULL,folder_id TEXT NOT NULL,modified TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS native_sync_run(mode TEXT PRIMARY KEY,started TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS native_sync_seen(entry_id TEXT PRIMARY KEY);
''')


def legacy_boundary():
    """Upgrade only a fully completed legacy window, never MAX(received_time)."""
    candidates = []
    for path in (rag.db_path().parent/'refresh-jobs').glob('*.json'):
        try:
            job = json.loads(path.read_text(encoding='utf-8'))
            folders = job.get('folders', {})
            if (job.get('status') == 'completed' and not job.get('mode')
                    and set(folders) == {'InboxAll', 'Sent'}
                    and all(v.get('has_more') is False for v in folders.values())
                    and job.get('scope', {}).get('hours') == 72):
                stamp = job['started_at']
                parsed = dt.datetime.fromisoformat(stamp.replace('Z', '+00:00'))
                if parsed.tzinfo and parsed <= dt.datetime.now(dt.UTC):
                    candidates.append(parsed)
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return max(candidates).isoformat() if candidates else None


def inventory():
    data = sync_source.invoke('inventory')
    if data.get('complete') is not True or not data.get('store_id') or not data.get('folders'):
        raise RuntimeError('Incomplete source inventory')
    rows = data.get('items')
    if not isinstance(rows, list):
        raise RuntimeError('Missing inventory rows')
    folders = data['folders']
    if len(folders) != len(set(folders)):
        raise RuntimeError('Duplicate folders')
    indexed = {}
    for row in rows:
        eid = row.get('EntryID')
        if not eid or eid in indexed or row.get('StoreID') != data['store_id'] or row.get('FolderID') not in folders or not row.get('ModifiedTime'):
            raise RuntimeError('Invalid or ambiguous source identity')
        indexed[eid] = row
    return data, indexed


def signature(data, rows):
    # Modification changes matter too: do not complete while the source is moving.
    value = [data['store_id'], sorted(data['folders']), sorted(
        (eid, r['FolderID'], r['ModifiedTime']) for eid, r in rows.items())]
    return hashlib.sha256(json.dumps(value).encode()).hexdigest()


def erase(c, ids):
    if not ids:
        return
    # One FTS scan for the whole deletion set, rather than one full scan per ID.
    # A temp table also avoids SQLite's variable limit on large reconciliations.
    c.execute('CREATE TEMP TABLE native_sync_missing(entry_id TEXT PRIMARY KEY)')
    c.executemany('INSERT INTO native_sync_missing VALUES(?)', ((eid,) for eid in ids))
    c.execute('DELETE FROM email_chunk_vectors WHERE chunk_id IN (SELECT chunk_id FROM email_chunks JOIN native_sync_missing USING(entry_id))')
    for table in ('email_chunks_fts', 'email_chunks', 'native_attachment_state', 'emails', 'native_sync_source'):
        c.execute(f'DELETE FROM {table} WHERE entry_id IN (SELECT entry_id FROM native_sync_missing)')
    c.execute('DROP TABLE native_sync_missing')


def run(mode, progress=lambda **kw: None):
    if mode not in {'incremental', 'full'}:
        raise ValueError('Invalid sync mode')
    initialize()
    with rag.connect(False) as c:
        c.execute('INSERT OR IGNORE INTO native_sync_run VALUES(?,?)', (mode, now()))
        started = c.execute('SELECT started FROM native_sync_run WHERE mode=?', (mode,)).fetchone()[0]
        checkpoint = c.execute("SELECT value FROM meta WHERE key='native_successful_sync_at'").fetchone()
        checkpoint = checkpoint[0] if checkpoint else None
        checkpoint = checkpoint or legacy_boundary()
        seen = {r[0] for r in c.execute('SELECT entry_id FROM native_sync_seen')}
        seen_established = bool(seen) or c.execute("SELECT 1 FROM meta WHERE key='native_inventory_established'").fetchone() is not None
        store = c.execute("SELECT value FROM meta WHERE key='native_sync_store'").fetchone()
    data, rows = inventory()
    if store and store[0] != data['store_id']:
        raise RuntimeError('Primary mailbox changed; preserve index and review scope locally')
    # First successful strict scan explicitly adopts the existing primary-mailbox
    # legacy cache. Subsequent runs are pinned to this exact StoreID.
    with rag.connect(False) as c:
        c.execute("INSERT OR IGNORE INTO meta VALUES('native_sync_store',?)", (data['store_id'],))
        cached = {r['entry_id']: dict(r) for r in c.execute('''SELECT e.entry_id,e.subject,e.updated_at,
            s.modified,s.store_id,s.folder_id,a.entry_id AS attachment_cache
            FROM emails e LEFT JOIN native_sync_source s USING(entry_id)
            LEFT JOIN native_attachment_state a USING(entry_id)''')}
    pending = []
    reused = 0
    deferred = 0
    verified_versions = set()
    server_complete = data.get('server_complete') is True
    # Full historical materialization needs server coverage. In cached mode
    # still fill the incremental interval, but never claim full reconciliation.
    incremental_scope = mode == 'incremental' or not server_complete
    cutoff = dt.datetime.fromisoformat(checkpoint.replace('Z', '+00:00')) - dt.timedelta(hours=1) if checkpoint else None
    for eid, row in rows.items():
        old = cached.get(eid)
        unchanged = old and old['modified'] == row['ModifiedTime'] and old['subject'] == row.get('Subject', '')
        # Upgrade legacy records only when Outlook's last change predates the
        # cached extraction and the attachment cache is present.
        if old and not old['modified'] and old['attachment_cache'] and old['subject'] == row.get('Subject', ''):
            try:
                unchanged = dt.datetime.fromisoformat(row['ModifiedTime'].replace('Z', '+00:00')) <= dt.datetime.fromisoformat(old['updated_at'].replace('Z', '+00:00'))
            except (ValueError, TypeError):
                unchanged = False
        outside_interval = False
        if incremental_scope and cutoff:
            try:
                times = [dt.datetime.fromisoformat(row[k].replace('Z', '+00:00')) for k in ('ReceivedTime','ModifiedTime') if row.get(k)]
                outside_interval = bool(times) and max(times) < cutoff
            except (ValueError, TypeError):
                outside_interval = False
            # A newly observed ID after a completed inventory is a late arrival
            # or move even if Exchange retained old timestamps.
            if seen_established and eid not in seen:
                outside_interval = False
            if old and old['modified'] and (old['modified'] != row['ModifiedTime'] or old['subject'] != row.get('Subject','')):
                outside_interval = False
        if unchanged:
            reused += 1
            verified_versions.add(eid)
        elif outside_interval:
            deferred += 1
        else:
            pending.append(row)
    progress(stage='sync', mode=mode, resumed_from=started, previous_success=checkpoint,
             scanned=len(rows), folders_scanned=len(data['folders']), pending_messages=len(pending),
             reused_messages=reused, deferred_history=deferred, processed_messages=0)
    categories = {}
    for offset in range(0, len(pending), 25):
        batch = pending[offset:offset+25]
        result = sync_source.invoke('read_batch', batch).get('items')
        expected = {r['EntryID']: r for r in batch}
        if not isinstance(result, list) or len(result) != len(batch) or {r.get('EntryID') for r in result} != set(expected):
            raise RuntimeError('Incomplete read batch; retain checkpoint')
        for index, row in enumerate(result):
            source = expected[row['EntryID']]
            if row.get('StoreID') != data['store_id'] or row.get('ModifiedTime') != source['ModifiedTime']:
                raise RuntimeError('Message changed during read; retry safely')
            records = attachments.collect(row, rag.db_path().parent)
            for record in records:
                category = record['status']
                categories[category] = categories.get(category, 0)+1
                if category in attachments.FATAL:
                    raise RuntimeError('Attachment extraction failed; retain checkpoint')
            row['_AttachmentRecords'] = records
            # Text, attachment cache, and source checkpoint commit together.
            with rag.connect(False) as c:
                rag.ingest([row], c)
                c.execute('INSERT OR REPLACE INTO native_sync_source VALUES(?,?,?,?)',
                          (row['EntryID'], data['store_id'], source['FolderID'], source['ModifiedTime']))
            verified_versions.add(row['EntryID'])
            progress(processed_messages=offset + index + 1,
                     attachments=categories)
        progress(processed_messages=min(offset+25, len(pending)), attachments=categories)
    progress(stage='verify')
    verified, after = inventory()
    if signature(data, rows) != signature(verified, after):
        raise RuntimeError('Mailbox changed during scan; no cleanup or watermark advance')
    # Missing IDs are removed only by a completely successful full scan. A moved
    # message with a new EntryID has already been ingested before this point.
    absent = sorted(set(cached)-set(after)) if mode == 'full' else []
    server_complete = verified.get('server_complete') is True and data.get('server_complete') is True
    missing = absent if server_complete else []
    backup = None
    if missing:
        progress(stage='backup', deletion_candidates=len(missing))
        folder = rag.db_path().parent / 'sync-backups'
        folder.mkdir(exist_ok=True)
        backup = folder / (dt.datetime.now(dt.UTC).strftime('%Y%m%dT%H%M%S%f')+'.sqlite')
        with rag.connect() as source, closing(sqlite3.connect(backup)) as target:
            source.backup(target)
        # Backup can take time; repeat a full source scan before any deletion.
        final, final_rows = inventory()
        if final.get('server_complete') is not True or signature(verified, after) != signature(final, final_rows):
            raise RuntimeError('Mailbox changed during backup; no cleanup')
    with rag.connect(False) as c:
        erase(c, missing)
        for eid, row in rows.items():
            # Do not mark an unmaterialized historical row as extracted.
            if eid not in verified_versions:
                continue
            c.execute('INSERT OR REPLACE INTO native_sync_source VALUES(?,?,?,?)',
                      (eid, data['store_id'], row['FolderID'], row['ModifiedTime']))
        if mode != 'full' or server_complete:
            c.execute('DELETE FROM native_sync_seen')
            c.executemany('INSERT INTO native_sync_seen VALUES(?)', ((eid,) for eid in rows))
            c.execute("INSERT OR REPLACE INTO meta VALUES('native_inventory_established','1')")
            c.execute("INSERT OR REPLACE INTO meta VALUES('native_successful_sync_at',?)", (started,))
        if mode == 'full' and server_complete:
            c.execute("INSERT OR REPLACE INTO meta VALUES('native_full_sync_at',?)", (started,))
        if mode != 'full' or server_complete:
            c.execute('DELETE FROM native_sync_run WHERE mode=?', (mode,))
    return dict(mode=mode, coverage_complete=mode != 'full' or server_complete,
                local_scan_complete=True, server_history_verified=server_complete,
                cleanup_blocked_reason='cached_exchange_cannot_prove_server_absence' if mode == 'full' and not server_complete else None,
                unverified_absent=len(absent) if not server_complete else 0,
                successful_sync_at=started if mode != 'full' or server_complete else checkpoint,
                previous_success=checkpoint, scanned=len(rows), processed=len(pending),
                reused=reused, removed=len(missing), backup_created=bool(backup),
                deferred_history=deferred,
                attachments=categories, scope='Primary mailbox normal mail folders; excludes Deleted/Junk/Drafts/Outbox/search folders',
                mailbox_modified=False, wsl_required=False)
