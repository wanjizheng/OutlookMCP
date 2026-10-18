import copy
import datetime as dt
import os
import json
from pathlib import Path
import sqlite3
import tempfile
from contextlib import closing
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from outlook_mcp import rag, synchronizer as sync, refresh

ROOT = Path(__file__).resolve().parents[1]


class SynchronizerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=ROOT/'artifacts')
        self.env = patch.dict(os.environ, OUTLOOK_MCP_DATA_DIR=self.temp.name)
        self.env.start()
        sync.initialize()
        self.rows = [self.row('one')]
        self.reads = []
        self.source = patch.object(sync.sync_source, 'invoke', side_effect=self.invoke)
        self.source.start()
        self.extract = patch.object(sync.attachments, 'collect', return_value=[{'status':'ok','name':'fixture.txt','text':'historicalattachmentneedle'}])
        self.extract.start()

    def tearDown(self):
        self.extract.stop()
        self.source.stop()
        self.env.stop()
        self.temp.cleanup()

    def row(self, eid, folder='inbox', modified='2026-10-18T00:00:00+00:00'):
        return dict(EntryID=eid, StoreID='primary', FolderID=folder, Subject='Same subject',
                    ModifiedTime=modified, ReceivedTime='2020-01-01T00:00:00+00:00', Body='mail body')

    def inventory(self, *, complete=True, server=True, store='primary'):
        rows=copy.deepcopy(self.rows)
        for row in rows:
            row['StoreID']=store
        return dict(complete=complete, server_complete=server, store_id=store,
                    folders=['inbox','archive'], items=rows)

    def invoke(self, operation, items=None):
        if operation == 'inventory':
            return self.inventory()
        self.assertEqual(operation, 'read_batch')
        self.reads.extend(r['EntryID'] for r in items)
        return dict(items=copy.deepcopy(items))

    def ids(self):
        with rag.connect() as c:
            return {r[0] for r in c.execute('SELECT entry_id FROM emails')}

    def meta(self, key):
        with rag.connect() as c:
            r = c.execute('SELECT value FROM meta WHERE key=?',(key,)).fetchone()
            return r[0] if r else None

    def seed_absent(self):
        with rag.connect(False) as c:
            rag.ingest([dict(self.row('absent'),_AttachmentRecords=[])],c)
            cid=c.execute("SELECT chunk_id FROM email_chunks WHERE entry_id='absent'").fetchone()[0]
            c.execute('INSERT INTO email_chunk_vectors VALUES(?,?,?,?)',(cid,1,b'1234','now'))

    def test_incremental_discovers_old_arrival_and_updates_after_long_gap(self):
        result=sync.run('incremental')
        watermark=self.meta('native_successful_sync_at')
        self.assertEqual(result['processed'],1)
        self.assertTrue(watermark)
        self.rows.append(self.row('late-old-mail'))
        result=sync.run('incremental')
        self.assertEqual(result['previous_success'],watermark)
        self.assertEqual(result['processed'],1)
        self.assertEqual(set(self.reads),{'one','late-old-mail'})
        self.assertEqual(len(rag.search('historicalattachmentneedle',mode='bm25')['items']),2)

    def test_interruption_resumes_committed_messages_without_advancing_watermark(self):
        self.rows.append(self.row('two'))
        calls=0
        def collect(row, data):
            nonlocal calls
            calls+=1
            if calls==2:
                raise RuntimeError('test interrupted')
            return []
        with patch.object(sync.attachments,'collect',side_effect=collect):
            with self.assertRaises(RuntimeError): sync.run('incremental')
        self.assertEqual(self.ids(),{'one'})
        self.assertIsNone(self.meta('native_successful_sync_at'))
        with rag.connect() as c:
            boundary=c.execute("SELECT started FROM native_sync_run WHERE mode='incremental'").fetchone()[0]
        self.reads.clear()
        result=sync.run('incremental')
        self.assertEqual(self.reads,['two'])
        self.assertEqual(result['successful_sync_at'],boundary)

    def test_incremental_retains_absent_and_full_prunes_all_derived_tables_with_backup(self):
        self.seed_absent()
        sync.run('incremental')
        self.assertIn('absent',self.ids())
        result=sync.run('full')
        self.assertEqual(result['removed'],1)
        self.assertTrue(result['backup_created'])
        self.assertEqual(self.ids(),{'one'})
        with rag.connect() as c:
            for table in ('emails','email_chunks','email_chunks_fts','native_attachment_state','native_sync_source'):
                self.assertEqual(c.execute(f"SELECT count(*) FROM {table} WHERE entry_id='absent'").fetchone()[0],0)
            self.assertEqual(c.execute('SELECT count(*) FROM email_chunk_vectors').fetchone()[0],0)
        backup=next((Path(self.temp.name)/'sync-backups').glob('*.sqlite'))
        with closing(sqlite3.connect(backup)) as c:
            self.assertEqual(c.execute("SELECT count(*) FROM emails WHERE entry_id='absent'").fetchone()[0],1)

    def test_cached_exchange_never_prunes_or_claims_full_history(self):
        self.seed_absent()
        with patch.object(sync.sync_source,'invoke',side_effect=lambda op,items=None: self.inventory(server=False) if op=='inventory' else dict(items=items)):
            result=sync.run('full')
        self.assertFalse(result['coverage_complete'])
        self.assertEqual(result['removed'],0)
        self.assertIn('absent',self.ids())
        self.assertIsNone(self.meta('native_full_sync_at'))
        self.assertIsNone(self.meta('native_successful_sync_at'))

    def test_failed_or_incomplete_folder_scan_preserves_everything(self):
        self.seed_absent()
        for response in (self.inventory(complete=False),dict(complete=True,store_id='primary',items=[])):
            with patch.object(sync.sync_source,'invoke',return_value=response):
                with self.assertRaises(RuntimeError): sync.run('full')
            self.assertEqual(self.ids(),{'absent'})
            self.assertIsNone(self.meta('native_full_sync_at'))

    def test_source_change_during_verification_prevents_prune_and_success(self):
        self.seed_absent()
        original=self.invoke
        calls=0
        def changing(op,items=None):
            nonlocal calls
            if op=='inventory':
                calls+=1
                if calls==2: self.rows[0]['FolderID']='archive'
            return original(op,items)
        with patch.object(sync.sync_source,'invoke',side_effect=changing):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.assertIn('absent',self.ids())
        self.assertIsNone(self.meta('native_full_sync_at'))

    def test_source_change_during_backup_prevents_prune(self):
        self.seed_absent()
        original=self.invoke
        calls=0
        def changing(op,items=None):
            nonlocal calls
            if op=='inventory':
                calls+=1
                if calls==3: self.rows.append(self.row('duringbackup'))
            return original(op,items)
        with patch.object(sync.sync_source,'invoke',side_effect=changing):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.assertIn('absent',self.ids())

    def test_moves_same_id_preserve_chunks_new_id_keeps_content_and_duplicates(self):
        sync.run('incremental')
        with rag.connect() as c:
            before=c.execute("SELECT chunk_id FROM email_chunks WHERE entry_id='one'").fetchone()[0]
        self.rows[0]['FolderID']='archive'
        self.assertEqual(sync.run('full')['processed'],0)
        with rag.connect() as c:
            self.assertEqual(c.execute("SELECT chunk_id FROM email_chunks WHERE entry_id='one'").fetchone()[0],before)
            self.assertEqual(c.execute("SELECT folder_id FROM native_sync_source WHERE entry_id='one'").fetchone()[0],'archive')
        self.rows=[self.row('moved-new-id','archive'),self.row('same-subject-copy')]
        result=sync.run('full')
        self.assertEqual(result['removed'],1)
        self.assertEqual(self.ids(),{'moved-new-id','same-subject-copy'})
        self.assertEqual(len(rag.search('historicalattachmentneedle',mode='bm25')['items']),2)

    def test_changed_mail_is_refetched_and_store_change_fails_closed(self):
        sync.run('incremental')
        self.rows[0]['ModifiedTime']='2026-10-18T00:00:01+00:00'
        self.assertEqual(sync.run('incremental')['processed'],1)
        with patch.object(sync.sync_source,'invoke',return_value=self.inventory(store='other')):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.assertEqual(self.ids(),{'one'})

    def test_missing_batch_duplicate_ids_and_attachment_failure_block_success(self):
        with patch.object(sync.sync_source,'invoke',side_effect=lambda op,items=None: self.inventory() if op=='inventory' else dict(items=[])):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.rows.append(self.rows[0].copy())
        with self.assertRaises(RuntimeError): sync.run('full')
        self.rows=self.rows[:1]
        with patch.object(sync.attachments,'collect',return_value=[{'status':'timeout'}]):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.assertIsNone(self.meta('native_successful_sync_at'))

    def test_worker_reports_text_success_and_embedding_failure_separately(self):
        model=Path(self.temp.name)/'models'
        model.mkdir()
        (model/'embedding.gguf').touch()
        state=dict(job_id='a'*32,mode='incremental')
        with patch.object(rag,'ROOT',Path(self.temp.name)),patch.object(rag,'embed_pending',side_effect=RuntimeError('private error')):
            refresh.execute(state)
        self.assertEqual(state['status'],'partial')
        self.assertTrue(state['sync']['coverage_complete'])
        self.assertTrue(self.meta('native_successful_sync_at'))
        self.assertNotIn('private error',str(state))

    def test_unsupported_metadata_never_exports_or_opens_outlook(self):
        from outlook_mcp import attachments
        self.extract.stop()
        row={'Attachments':[{'FileName':'signature.png','Size':100,'Type':1}]}
        with patch.object(attachments.subprocess,'run',side_effect=AssertionError('must not export')):
            result=attachments.collect(row,Path(self.temp.name))
        self.assertEqual(result,[{'index':1,'name':'signature.png','status':'unsupported'}])

    def test_legacy_success_migration_uses_boundary_not_newest_received_mail(self):
        folder=Path(self.temp.name)/'refresh-jobs'
        folder.mkdir()
        boundary=(dt.datetime.now(dt.UTC)-dt.timedelta(days=2)).isoformat()
        (folder/'legacy.json').write_text(json.dumps(dict(status='completed',started_at=boundary,
            scope={'hours':72},folders={'InboxAll':{'has_more':False},'Sent':{'has_more':False}})),encoding='utf-8')
        self.rows=[self.row('history')]
        self.rows[0]['ModifiedTime']='2020-01-01T00:00:00+00:00'
        recent=self.row('after-boundary',modified=dt.datetime.now(dt.UTC).isoformat())
        self.rows.append(recent)
        result=sync.run('incremental')
        self.assertEqual(result['previous_success'],boundary)
        self.assertEqual(self.reads,['after-boundary'])
        self.assertEqual(result['deferred_history'],1)
        self.rows.append(self.row('late-delivery',modified='2020-01-01T00:00:00+00:00'))
        self.reads.clear()
        sync.run('incremental')
        self.assertEqual(self.reads,['late-delivery'])
        self.assertNotIn('history',self.ids())
        result=sync.run('full')
        self.assertTrue(result['coverage_complete'])
        self.assertIn('history',self.ids())

    def test_old_dated_arrival_after_empty_baseline_is_not_skipped(self):
        self.rows=[]
        sync.run('incremental')
        self.rows=[self.row('late-after-empty',modified='2020-01-01T00:00:00+00:00')]
        self.assertEqual(sync.run('incremental')['processed'],1)

    def test_cleanup_transaction_failure_rolls_back_every_derived_table(self):
        self.seed_absent()
        erase=sync.erase
        def fail(c, ids):
            erase(c, ids)
            raise RuntimeError('test transaction failure')
        with patch.object(sync,'erase',side_effect=fail):
            with self.assertRaises(RuntimeError): sync.run('full')
        self.assertEqual(self.ids(),{'absent','one'})
        self.assertIsNone(self.meta('native_full_sync_at'))
        with rag.connect() as c:
            self.assertEqual(c.execute('SELECT count(*) FROM email_chunk_vectors').fetchone()[0],1)

    def test_backup_failure_prevents_cleanup(self):
        self.seed_absent()
        def fail(*args,**kwargs): raise OSError('disk full')
        with patch.object(sync,'sqlite3',SimpleNamespace(connect=fail)):
            with self.assertRaises(OSError): sync.run('full')
        self.assertIn('absent',self.ids())
        self.assertIsNone(self.meta('native_full_sync_at'))

    def test_cached_full_worker_reports_blocker_without_embedding_or_deleting(self):
        self.seed_absent()
        state=dict(job_id='b'*32,mode='full')
        with patch.object(sync.sync_source,'invoke',side_effect=lambda op,items=None: self.inventory(server=False) if op=='inventory' else dict(items=items)),patch.object(rag,'embed_pending') as embed:
            refresh.execute(state)
        self.assertEqual(state['status'],'partial')
        self.assertTrue(state['has_more_mail'])
        self.assertEqual(state['sync']['cleanup_blocked_reason'],'cached_exchange_cannot_prove_server_absence')
        self.assertEqual(state['sync']['removed'],0)
        embed.assert_not_called()

    def test_manual_vector_batches_commit_all_rows_with_bounded_model_loads(self):
        with rag.connect(False) as c:
            rag.ingest([dict(self.row(str(i)),_AttachmentRecords=[]) for i in range(65)],c)
        def vectors(texts,**kwargs):
            self.assertEqual(kwargs,{'timeout':180})
            self.assertLessEqual(len(texts),16)
            return [[1.0,0.0] for _ in texts]
        with patch.object(rag,'embed',side_effect=vectors) as embed:
            result=rag.embed_pending(64,batch_size=16,embedding_timeout=180)
        self.assertEqual(result['embedded'],64)
        self.assertEqual(embed.call_count,4)
        with rag.connect() as c:
            self.assertEqual(c.execute('SELECT count(*) FROM email_chunk_vectors').fetchone()[0],64)

    def test_vector_failure_preserves_prior_batch_and_next_call_resumes(self):
        with rag.connect(False) as c:
            rag.ingest([dict(self.row(str(i)),_AttachmentRecords=[]) for i in range(33)],c)
        calls=0
        progress=[]
        def interrupted(texts,**kw):
            nonlocal calls
            calls+=1
            if calls==2: raise RuntimeError('model interrupted')
            return [[1.0,0.0] for _ in texts]
        with patch.object(rag,'embed',side_effect=interrupted):
            with self.assertRaises(RuntimeError): rag.embed_pending(64,batch_size=16,embedding_timeout=180,progress=progress.append)
        self.assertEqual(progress,[16])
        with rag.connect() as c:
            self.assertEqual(c.execute('SELECT count(*) FROM email_chunk_vectors').fetchone()[0],16)
        with patch.object(rag,'embed',side_effect=lambda texts,**kw:[[1.0,0.0] for _ in texts]):
            self.assertEqual(rag.embed_pending(64,batch_size=16,embedding_timeout=180)['embedded'],17)

    def test_bulk_reconciliation_only_removes_missing_ids(self):
        with rag.connect(False) as c:
            rag.ingest([dict(self.row('missing'+str(i)),_AttachmentRecords=[]) for i in range(1100)],c)
        result=sync.run('full')
        self.assertEqual(result['removed'],1100)
        self.assertEqual(self.ids(),{'one'})
        self.assertEqual(len(rag.search('historicalattachmentneedle',mode='bm25')['items']),1)
