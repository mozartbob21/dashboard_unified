"""Download boundaries and bounded uploads for the private tools workspace."""
from pathlib import PurePosixPath
from starlette.responses import JSONResponse


def private_data_path(path):
    if not path.startswith('/data/'):
        return False
    parts=path.replace('\\','/').split('/')
    if any(p in ('.','..') for p in parts):
        return True
    low=[p.casefold() for p in parts]
    if len(low)>2 and low[2]=='tools':
        return True
    forbidden={'auth','telegram','browser-profile','profile','debug','cache','cookies','history','local state',
               'session storage','local storage','storage_state.json','session.json','integrations.key','jwt_secret.key'}
    if any(p in forbidden or p.startswith('.') or 'password' in p or 'token' in p for p in low if p):
        return True
    suffix=PurePosixPath(low[-1]).suffix
    return suffix in {'.key','.pem','.pfx','.p12','.db','.sqlite','.sqlite3','.session','.log','.html','.htm','.py','.exe','.wsb'}


class ToolsBodyLimit:
    def __init__(self,app,max_bytes=61*1024*1024):
        self.app=app;self.max_bytes=max_bytes

    async def __call__(self,scope,receive,send):
        if scope['type'] != 'http':
            return await self.app(scope, receive, send)
        path, method = scope.get('path', ''), scope.get('method')
        chat_upload = path.rstrip('/') == '/aichat/api/send' and method == 'POST'
        tool_upload = path.startswith('/tools/') and method in {'POST', 'PUT'}
        if not (chat_upload or tool_upload):
            return await self.app(scope, receive, send)
        # Bound the raw multipart envelope before Starlette can spool files.
        limit = 25 * 1024 * 1024 if chat_upload else self.max_bytes
        error = ('До 8 МБ на файл и до 24 МБ на сообщение.' if chat_upload
                 else 'Общий размер запроса превышает 60 МБ.')
        headers = dict(scope.get('headers', []))
        try:
            declared = int(headers.get(b'content-length', b'0'))
        except ValueError:
            declared = 0
        if declared > limit:
            return await JSONResponse({'ok': False, 'message': error, 'detail': error}, status_code=413)(scope, receive, send)
        messages=[];size=0
        while True:
            message=await receive()
            if message['type']=='http.disconnect':return
            size+=len(message.get('body',b''))
            if size > limit:
                return await JSONResponse({'ok': False, 'message': error, 'detail': error}, status_code=413)(scope, receive, send)
            messages.append(message)
            if not message.get('more_body'):break
        iterator=iter(messages)
        async def replay():
            try:return next(iterator)
            except StopIteration:return await receive()
        return await self.app(scope,replay,send)
