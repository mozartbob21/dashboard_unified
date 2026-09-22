"""Bounded image retrieval. Resolve and pin public IPs; no cookies, proxy or redirects to LAN."""
import http.client
import ipaddress
import re
import socket
import ssl
import time
from urllib.parse import urlsplit, urljoin
from bs4 import BeautifulSoup
from .images import image_uri
from .workspace import ToolError

MAX_IMAGE=8*1024**2
MAX_TOTAL=24*1024**2
MAX_IMAGES=20


def public_address(host,port):
    try:
        addresses=socket.getaddrinfo(host,port,type=socket.SOCK_STREAM)
        if not addresses:raise ValueError()
        for item in addresses:
            ip=ipaddress.ip_address(item[4][0])
            if not ip.is_global or ip.is_multicast or ip.is_reserved:
                raise ToolError('Внутренние и служебные адреса для картинок запрещены.')
            if getattr(ip,'ipv4_mapped',None) or getattr(ip,'sixtofour',None) or getattr(ip,'teredo',None):
                raise ToolError('Туннельные IP-адреса не поддерживаются.')
        return addresses[0][4][0]
    except (OSError,ValueError):
        raise ToolError('Не удалось определить адрес сервера изображения.')


class PinnedConnection(http.client.HTTPConnection):
    def __init__(self,host,port,ip,tls,timeout):
        super().__init__(host,port,timeout=timeout)
        self.ip=ip;self.tls=tls
    def connect(self):
        self.sock=socket.create_connection((self.ip,self.port),self.timeout)
        if self.tls:
            try:self.sock=ssl.create_default_context().wrap_socket(self.sock,server_hostname=self.host)
            except Exception:
                self.sock.close();raise


def download(url,deadline=None):
    deadline=deadline or time.monotonic()+30
    try:
        for _ in range(4):
            p=urlsplit(url)
            if p.scheme not in ('http','https') or not p.hostname or p.username or p.password or p.port not in (None,80,443):
                raise ToolError('Недопустимый адрес изображения.')
            remaining=deadline-time.monotonic()
            if remaining<=0:raise ToolError('Истекло время загрузки изображений.')
            host=p.hostname.encode('idna').decode('ascii');port=p.port or (443 if p.scheme=='https' else 80)
            ip=public_address(host,port)
            connection=PinnedConnection(host,port,ip,p.scheme=='https',min(8,remaining))
            try:
                path=p.path or '/'
                if p.query:path+='?'+p.query
                connection.request('GET',path,headers={'Accept':'image/*','User-Agent':'Neurona image converter/1.0'})
                response=connection.getresponse()
                if response.status in (301,302,303,307,308):
                    location=response.getheader('Location')
                    if not location:raise ToolError('Некорректное перенаправление изображения.')
                    target=urljoin(url,location)
                    if p.scheme=='https' and urlsplit(target).scheme!='https':
                        raise ToolError('Перенаправление на незащищённый адрес запрещено.')
                    url=target;continue
                if response.status!=200:raise ToolError('Сервер не отдал изображение.')
                if not (response.getheader('Content-Type') or '').lower().startswith('image/'):
                    raise ToolError('По ссылке находится не изображение.')
                if int(response.getheader('Content-Length') or '0')>MAX_IMAGE:
                    raise ToolError('Изображение превышает 8 МБ.')
                result=bytearray()
                while True:
                    left=deadline-time.monotonic()
                    if left<=0:raise ToolError('Истекло время загрузки изображений.')
                    if connection.sock:connection.sock.settimeout(min(8,left))
                    chunk=response.read1(65536)
                    if not chunk:break
                    result.extend(chunk)
                    if len(result)>MAX_IMAGE:raise ToolError('Изображение превышает 8 МБ.')
                return bytes(result)
            finally:connection.close()
        raise ToolError('Слишком много перенаправлений изображения.')
    except ToolError:raise
    except (OSError,ValueError,http.client.HTTPException):
        raise ToolError('Не удалось загрузить изображение по ссылке.')


def embed_remote(html):
    soup=BeautifulSoup(html,'html.parser');assets={};failed=set();size=0;deadline=time.monotonic()+60
    base=soup.find('base',href=True);base_url=base['href'] if base else ''
    def replace(src):
        nonlocal size
        if src.startswith('data:'):return src
        url=urljoin(base_url,src)
        if url.startswith('//'):url='https:'+url
        if not url.lower().startswith(('http://','https://')):return src
        if url in assets:return assets[url]
        if url in failed:return ''
        if len(assets)+len(failed)>=MAX_IMAGES:raise ToolError('В HTML больше 20 внешних изображений. Уменьшите их число или загрузите файлы вручную.')
        try:
            data=download(url,deadline)
            if size+len(data)>MAX_TOTAL:raise ToolError('Превышен общий размер изображений.')
            result=image_uri(data)
        except ToolError:
            failed.add(url);return ''
        size+=len(data);assets[url]=result;return result
    for image in soup.find_all('img'):
        image['src']=replace(str(image.get('src') or ''))
        image.attrs.pop('srcset',None)
    for source in soup.select('picture source'):source.decompose()
    pattern=re.compile(r'url\(\s*([\'"]?)(.*?)\1\s*\)',re.I)
    def css(value):return pattern.sub(lambda m:'url("'+replace(m.group(2))+'")',value)
    for tag in soup.find_all(style=True):tag['style']=css(tag['style'])
    for style in soup.find_all('style'):
        if style.string:style.string=css(str(style.string))
    return str(soup),len(failed)
