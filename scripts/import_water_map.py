"""Import the user's archive as JSON data only; run from the repository root.

python scripts/import_water_map.py '/path/to/archive.zip'
The same import is available to administrators on /mingkh/water-map.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from services.mingkh.water_map import MAX_BYTES, import_upload


def main():
    parser = argparse.ArgumentParser(description='Импорт архива жалоб по воде в МИНЖКХ')
    parser.add_argument('archive', type=Path)
    args = parser.parse_args()
    with args.archive.open('rb') as stream:
        content = stream.read(MAX_BYTES + 1)
    try:
        result = import_upload(content)
    except ValueError as error:
        parser.error(str(error))
    print(f"Импортировано {result['count']} жалоб: {result['from']} — {result['to']}; обновление {result['updated']}")


if __name__ == '__main__':
    main()
