"""Fresh opt-in MCP process; only preparation and rejected send calls, no UI/email."""

import asyncio
import json
import os
from pathlib import Path
import sys
import unittest

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]


class SendStdioTests(unittest.TestCase):
    def test_opt_in_tools_and_invalid_preparation_cannot_send(self):
        async def check():
            parameters = StdioServerParameters(command=sys.executable,
                args=["-m", "outlook_mcp.tunnel_server"], cwd=str(ROOT),
                env=dict(os.environ, OUTLOOK_MCP_ENABLE_SEND="1"))
            async with stdio_client(parameters) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    tools = {t.name: t for t in (await session.list_tools()).tools}
                    self.assertEqual(len(tools), 13)
                    self.assertFalse(tools["send_prepared_email"].annotations.readOnlyHint)
                    preview = await session.call_tool("prepare_email", {
                        "sender_account": "sender@example.invalid", "to": ["recipient@example.invalid"],
                        "subject": "Non-sending protocol test", "body": "Synthetic test only."})
                    self.assertFalse(preview.isError)
                    data = preview.structuredContent or json.loads(preview.content[0].text)
                    self.assertFalse(data["sent"])
                    denied = await session.call_tool("send_prepared_email", {"preparation_id": "unknown"})
                    self.assertTrue(denied.isError)
                    self.assertNotIn("confirmed", tools["send_prepared_email"].inputSchema["properties"])
                    bypass = await session.call_tool("send_prepared_email", {"preparation_id": "invalid-even-if-confirmed", "confirmed": True})
                    self.assertTrue(bypass.isError)
                    test_result = await session.call_tool("send_prepared_email", {
                        "preparation_id": data["preparation_id"], "confirmation_test": True})
                    self.assertFalse(test_result.isError)
                    test_data = test_result.structuredContent or json.loads(test_result.content[0].text)
                    self.assertEqual(test_data["status"], "chat_confirmation_test_passed")
                    self.assertEqual((test_data["sent"], test_data["test"]), (False, True))
                    replay = await session.call_tool("send_prepared_email", {
                        "preparation_id": data["preparation_id"], "confirmation_test": True})
                    self.assertTrue(replay.isError)
                    calendar_preview = await session.call_tool("prepare_calendar_action", {
                        "action":"create", "sender_account":"sender@example.invalid", "subject":"NO SEND calendar test",
                        "body":"Synthetic only", "start":"2026-10-20T14:00:00+01:00", "end":"2026-10-20T15:00:00+01:00",
                        "required":["recipient@example.invalid"]})
                    self.assertFalse(calendar_preview.isError)
                    cp = calendar_preview.structuredContent or json.loads(calendar_preview.content[0].text)
                    self.assertIn("STOP",cp["next_step"])
                    self.assertFalse(tools["commit_calendar_action"].annotations.readOnlyHint)
                    cr = await session.call_tool("commit_calendar_action", {"preparation_id":cp["preparation_id"],"confirmation_test":True})
                    self.assertFalse(cr.isError)
                    result=cr.structuredContent or json.loads(cr.content[0].text)
                    self.assertFalse(result["sent"])
                    self.assertEqual(result["status"],"chat_confirmation_test_passed")
                    replay=await session.call_tool("commit_calendar_action", {"preparation_id":cp["preparation_id"],"confirmation_test":True})
                    self.assertTrue(replay.isError)

        asyncio.run(asyncio.wait_for(check(), timeout=30))
