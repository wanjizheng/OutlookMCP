import os
import unittest
from unittest.mock import patch
from outlook_mcp import calendar_actions as c

class CalendarActionsTests(unittest.TestCase):
    def setUp(self):
        self.env=patch.dict(os.environ, OUTLOOK_MCP_ENABLE_SEND='1'); self.env.start(); c._pending.clear()
        self.event=dict(organizer_smtp='owner@example.invalid', attendees=[dict(smtp='old@example.invalid',type=1)],body='Existing Zoom link',modified='v1')
    def tearDown(self): self.env.stop()
    def create(self, **kw):
        args=dict(action='create',sender_account='owner@example.invalid',subject='Test only',body='No sending',start='2026-10-20T14:00:00+01:00',end='2026-10-20T15:00:00+01:00',required=['a@example.invalid']);args.update(kw);return c.prepare(**args)
    def test_create_preview_no_outlook_and_no_send_test_one_use(self):
        with patch.object(c,'invoke') as invoke:
            p=self.create(); self.assertIn('STOP',p['next_step']);self.assertFalse(p['sent'])
            r=c.commit(p['preparation_id'],True);self.assertFalse(r['sent']);invoke.assert_not_called()
            with self.assertRaises(ValueError): c.commit(p['preparation_id'],True)
    def test_invalid_time_recipients_scope(self):
        for kw in [dict(start='2026-10-20T14:00:00'),dict(end='2026-10-20T13:00:00Z'),dict(required=['bad']),dict(scope='series'),dict(required=[],optional=[])]:
            with self.assertRaises(ValueError):self.create(**kw)
    def test_preserves_existing_content_and_rejects_stale(self):
        with patch.object(c,'read',return_value=self.event),patch.object(c,'invoke') as invoke:
            p=c.prepare('add_remove_attendees','owner@example.invalid',entry_id='id',store_id='store',add_attendees=['new@example.invalid'])
            self.assertEqual(p['preview']['before']['body'],'Existing Zoom link')
            self.event['modified']='v2'
            with self.assertRaises(ValueError): c.commit(p['preparation_id'])
            invoke.assert_not_called()
    def test_organizer_and_attendee_validation(self):
        with patch.object(c,'read',return_value=self.event):
            for kw in [dict(sender_account='other@example.invalid',add_attendees=['new@example.invalid']),dict(sender_account='owner@example.invalid',remove_attendees=['missing@example.invalid']),dict(sender_account='owner@example.invalid',add_attendees=['old@example.invalid'])]:
                with self.assertRaises(ValueError): c.prepare('add_remove_attendees',entry_id='id',store_id='store',**kw)
    def test_expiry_disabled_and_immutable(self):
        p=self.create();p['preview']['subject']='tampered'
        r=c.commit(p['preparation_id'],True);self.assertEqual(r['preview']['subject'],'Test only')
        p=self.create()
        with patch.object(c.time,'monotonic',return_value=10**12):
            with self.assertRaises(ValueError):c.commit(p['preparation_id'],True)
        with patch.dict(os.environ,OUTLOOK_MCP_ENABLE_SEND='0'):
            with self.assertRaises(RuntimeError):self.create()

    @unittest.skipUnless(os.name == 'nt', 'Windows PowerShell mock test')
    def test_native_mock(self):
        import subprocess
        from pathlib import Path
        root=Path(__file__).resolve().parents[1]
        exe=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
        result=subprocess.run([str(exe),'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(root/'tests/test_calendar_native.ps1')],capture_output=True,timeout=30)
        self.assertEqual(result.returncode,0,result.stdout.decode('utf-8',errors='replace')+result.stderr.decode('utf-8',errors='replace'))
