"""Only explicitly uploaded images become presentation assets."""
import base64
import io
from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit
from bs4 import BeautifulSoup
from PIL import Image, ImageOps
from .workspace import ToolError


def image_uri(content):
    try:
        with Image.open(io.BytesIO(content)) as source:
            if source.width * source.height > 16000000:
                raise ToolError('Изображение превышает 16 мегапикселей.')
            image = ImageOps.exif_transpose(source).convert('RGBA')
            image.thumbnail((2400, 2400))
            data = io.BytesIO()
            image.save(data, format='PNG')
    except (OSError, ValueError, Image.DecompressionBombError):
        raise ToolError('Не удалось прочитать изображение. Поддерживаются PNG, JPEG, WebP и GIF.')
    return 'data:image/png;base64,' + base64.b64encode(data.getvalue()).decode()


def embed_uploads(html, uploads):
    if not uploads:
        return html
    if len(uploads) > 20:
        raise ToolError('Можно выбрать до 20 изображений.')
    assets = {}
    for name, content in uploads:
        if name.casefold() in assets:
            raise ToolError('У изображений должны быть разные имена файлов.')
        assets[name.casefold()] = (name, image_uri(content))
    soup = BeautifulSoup(html, 'html.parser')
    used = set()
    for image in soup.find_all('img'):
        src = str(image.get('src') or '')
        if src.startswith('data:'):
            continue
        key = PurePosixPath(unquote(urlsplit(src).path).replace('\\', '/')).name.casefold()
        if key in assets:
            image['src'] = assets[key][1]
            image.attrs.pop('srcset', None)
            used.add(key)
    # Pictures without a matching <img src="filename"> get their own slide.
    target = soup.body or soup
    if any(key not in used for key in assets) and not soup.find('section'):
        original = soup.new_tag('section')
        for child in list(target.contents):
            original.append(child.extract())
        target.append(original)
    for key, (name, content) in assets.items():
        if key in used:
            continue
        section = soup.new_tag('section')
        title = soup.new_tag('h2'); title.string = PurePosixPath(name).stem
        image = soup.new_tag('img', src=content, alt=PurePosixPath(name).stem)
        section.extend([title, image]); target.append(section)
    return str(soup)
