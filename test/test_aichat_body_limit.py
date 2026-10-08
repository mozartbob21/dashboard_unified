"""Reject oversized chat requests before multipart parsing, without network."""
import unittest
from core.http_security import ToolsBodyLimit


class ChatBodyLimitTests(unittest.IsolatedAsyncioTestCase):
    async def exercise(self, messages, headers=(), path='/aichat/api/send'):
        forwarded, sent, reads = [], [], []
        source = iter(messages)
        async def receive():
            reads.append(True)
            return next(source)
        async def send(event):
            sent.append(event)
        async def app(scope, receive, send):
            while True:
                event = await receive()
                forwarded.append(event)
                if not event.get('more_body'):
                    break
        scope = {'type': 'http', 'path': path, 'method': 'POST', 'headers': list(headers)}
        await ToolsBodyLimit(app)(scope, receive, send)
        return forwarded, sent, reads

    async def test_declared_large_request_rejected_before_body_read(self):
        forwarded, sent, reads = await self.exercise([], [(b'content-length', str(26*1024*1024).encode())])
        self.assertFalse(forwarded)
        self.assertFalse(reads)
        self.assertEqual(sent[0]['status'], 413)

    async def test_chunked_large_request_rejected_before_parser(self):
        chunk = b'x' * (13*1024*1024)
        forwarded, sent, reads = await self.exercise([
            {'type': 'http.request', 'body': chunk, 'more_body': True},
            {'type': 'http.request', 'body': chunk, 'more_body': False}])
        self.assertFalse(forwarded)
        self.assertEqual(len(reads), 2)
        self.assertEqual(sent[0]['status'], 413)

    async def test_valid_body_replayed_without_changes(self):
        messages = [{'type': 'http.request', 'body': b'first', 'more_body': True},
                    {'type': 'http.request', 'body': b'last', 'more_body': False}]
        forwarded, sent, _ = await self.exercise(messages)
        self.assertEqual(forwarded, messages)
        self.assertFalse(sent)

    async def test_tool_upload_keeps_existing_larger_limit(self):
        messages = [{'type': 'http.request', 'body': b'ok', 'more_body': False}]
        forwarded, sent, _ = await self.exercise(messages, [(b'content-length', str(26*1024*1024).encode())], '/tools/api/run')
        self.assertEqual(forwarded, messages)
        self.assertFalse(sent)


if __name__ == '__main__':
    unittest.main()
