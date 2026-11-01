import os
import sys

# Cho Vercel tim thay cac module o thu muc goc du an (app.py, models.py...).
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app import app  # noqa: E402

# Vercel's Python runtime (@vercel/python) nhan dien bien "app" la mot WSGI app
# va tu dong goi no cho moi request.
