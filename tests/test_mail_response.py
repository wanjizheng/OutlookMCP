import json
import os
import unittest
from unittest.mock import patch
from outlook_mcp import sending, calendar_actions

class MailResponseTests(unittest.TestCase):
    def test_response_immutable_preview_test_never_opens_outlook(self):
        source=dict(entry_id='id',store_id='store',subject='Original',body='Original body',modified='v1')
        with patch.dict(os.environ,OUTLOOK_MCP_ENABLE_SEND='1'), patch.object(calendar_actions,'invoke',return_value=source),patch.object(sending.subprocess,'run') as run:
            p=sending.prepare_response('reply_all','id','store','sender@example.invalid',['recipient@example.invalid'],'Exact reply')
            self.assertEqual(p['message']['subject'],'RE: Original')
            source['body']='tampered'
            r=sending.send(p['preparation_id'],confirmation_test=True)
            self.assertEqual(r['message']['source']['body'],'Original body')
            self.assertEqual(r['message']['body'],'Exact reply')
            self.assertEqual(r['message']['to'],['recipient@example.invalid'])
            run.assert_not_called()
    def test_invalid_mode_and_opt_in(self):
        with patch.dict(os.environ,OUTLOOK_MCP_ENABLE_SEND='1'),patch.object(calendar_actions,'invoke') as invoke:
            with self.assertRaises(ValueError):sending.prepare_response('delete','id','store','s@example.invalid',['t@example.invalid'],'body')
            invoke.assert_not_called()

    @unittest.skipUnless(os.name == 'nt', 'Windows native mock')
    def test_native_response_mock(self):
        import subprocess
        from pathlib import Path
        root=Path(__file__).resolve().parents[1]
        exe=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
        r=subprocess.run([str(exe),'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(root/'tests/test_mail_response_native.ps1')],capture_output=True,timeout=30)
        self.assertEqual(r.returncode,0,r.stderr.decode('utf-8',errors='replace'))
