"""Local text measurement and lossless wrapping for editable complaint tables."""
from functools import lru_cache
from pathlib import Path
import os

from PIL import ImageFont

FONT_SIZE = 14
LINE_HEIGHT = 17.5
PARAGRAPH_GAP = 3
CELL_MARGIN_X = .09
CELL_MARGIN_Y = .07
COLUMN_WIDTHS = (1.70, 1.10, 3.75, 5.82)
HEADER_HEIGHT = .48
BODY_HEIGHT = 5.65
MAX_ROWS_PER_SLIDE = 5


@lru_cache(maxsize=8)
def _font(bold=False, size=FONT_SIZE):
    face = 'Arial Bold.ttf' if bold else 'Arial.ttf'
    win_face = 'arialbd.ttf' if bold else 'arial.ttf'
    linux_face = 'LiberationSans-Bold.ttf' if bold else 'LiberationSans-Regular.ttf'
    candidates = [Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts' / win_face,
                  Path('/System/Library/Fonts/Supplemental') / face,
                  Path('/usr/share/fonts/truetype/msttcorefonts') / face,
                  Path('/usr/share/fonts/truetype/liberation2') / linux_face,
                  Path('/usr/share/fonts/truetype/liberation') / linux_face]
    for path in candidates:
        if path.is_file():
            try:
                return ImageFont.truetype(str(path), size * 10)
            except OSError:
                pass
    # No font download. Unknown deployment fonts use a conservative em width.
    return None


def text_width(value, bold=False, size=FONT_SIZE):
    font = _font(bold, size)
    return font.getlength(value) / 10 if font else len(value) * size * 1.05


def wrap_text(text, width, bold=False, size=FONT_SIZE):
    """Wrap on words, splitting a long token only when needed, without cutting it."""
    available = max(10, (width - 2 * CELL_MARGIN_X) * 72 * .96)
    output, line = [], ''
    for word in text.split():
        trial = (line + ' ' + word).strip()
        if text_width(trial, bold, size) <= available:
            line = trial
            continue
        if line:
            output.append(line)
            line = ''
        while word:
            hi = min(len(word), 128)
            while hi < len(word) and text_width(word[:hi], bold, size) <= available:
                hi = min(len(word), hi * 2)
            if hi == len(word) and text_width(word, bold, size) <= available:
                break
            lo = 1
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if text_width(word[:mid], bold, size) <= available:
                    lo = mid
                else:
                    hi = mid - 1
            hyphen = word.rfind('-', 0, lo + 1)
            if hyphen >= max(2, lo // 2):
                lo = hyphen + 1
            output.append(word[:lo])
            word = word[lo:]
        line = word
    if line:
        output.append(line)
    return output or ['']


def lines_for(blocks, column):
    lines = []
    for text, bold in blocks:
        for paragraph in str(text).splitlines():
            if not paragraph.strip():
                continue
            wrapped = wrap_text(paragraph, COLUMN_WIDTHS[column], bold)
            lines.extend((line, bold, PARAGRAPH_GAP if i == len(wrapped) - 1 else 0)
                         for i, line in enumerate(wrapped))
    return lines


def height(lines):
    return sum(LINE_HEIGHT + line[2] for line in lines)


def row_height(cells):
    return (max((height(cell) for cell in cells), default=0) + CELL_MARGIN_Y * 144) / 72


def fragments(cells):
    """Split an oversized appeal into explicitly marked continuation rows."""
    remaining = [list(cell) for cell in cells]
    first = True
    municipality = list(cells[0])
    signatures = list(cells[1])
    limit = BODY_HEIGHT * 72 - CELL_MARGIN_Y * 144
    while any(remaining):
        row = [[], [], [], []]
        if not first:
            row[0] = lines_for([('Продолж.', True)], 0)
            # Repeat a short municipality so continuation pages remain identifiable.
            if not remaining[0] and height(municipality) < limit / 3:
                row[0] += municipality
            if not remaining[1]:
                row[1] = list(signatures)
        for i, pending in enumerate(remaining):
            used = height(row[i])
            consumed = 0
            for line in pending:
                needed = LINE_HEIGHT + line[2]
                if used + needed > limit:
                    break
                row[i].append(line)
                used += needed
                consumed += 1
            del pending[:consumed]
        yield row
        first = False
