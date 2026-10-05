"""Tile provider settings never accept arbitrary external endpoints."""
import json
import os
import unittest
from unittest.mock import patch

from core.map_tiles import browser_config, image_origin, script_config


class MapTileConfigTests(unittest.TestCase):
    def config(self, provider="auto", key=""):
        return patch.dict(os.environ, {"MAP_TILE_PROVIDER": provider, "YANDEX_TILES_API_KEY": key})

    def test_auto_uses_yandex_only_when_key_exists(self):
        with self.config():
            self.assertEqual(browser_config(), {"provider": "osm", "notice": ""})
            self.assertEqual(image_origin(), "https://tile.openstreetmap.org")
        with self.config(key="  synthetic-key  "):
            self.assertEqual(browser_config(), {"provider": "yandex", "api_key": "synthetic-key", "notice": ""})
            self.assertEqual(image_origin(), "https://tiles.api-maps.yandex.ru")

    def test_explicit_osm_never_exposes_unused_key(self):
        with self.config(" OSM ", "unused-key"):
            self.assertEqual(browser_config()["provider"], "osm")
            self.assertNotIn("unused-key", script_config())

    def test_missing_key_and_unknown_provider_preserve_working_map_with_notice(self):
        for provider in ("yandex", "https://untrusted.invalid; script-src *"):
            with self.subTest(provider=provider), self.config(provider):
                self.assertEqual(browser_config()["provider"], "osm")
                self.assertTrue(browser_config()["notice"])
                self.assertEqual(image_origin(), "https://tile.openstreetmap.org")
                self.assertNotIn("untrusted", script_config())

    def test_script_json_cannot_break_out_of_html(self):
        key = '</script><img src=x onerror=alert(1)>&"\u2028\u2029end'
        with self.config(key=key):
            payload = script_config()
            self.assertEqual(json.loads(payload)["api_key"], key)
            for unsafe in ("<", ">", "&", "\u2028", "\u2029"):
                self.assertNotIn(unsafe, payload)


if __name__ == "__main__":
    unittest.main()
