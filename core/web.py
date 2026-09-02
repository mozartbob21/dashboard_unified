# -*- coding: utf-8 -*-
"""Общий объект templates для app.py и всех роутеров."""
from pathlib import Path
from fastapi.templating import Jinja2Templates
from jinja2 import ChainableUndefined

BASE_DIR = Path(__file__).resolve().parents[1]
TEMPLATES_DIR = BASE_DIR / "templates"

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
templates.env.undefined = ChainableUndefined
