import models
from schema import BIEU_SCHEMAS, SAN_PHAM_SCHEMA

# Noi schema (du lieu mo ta cot/nhan) voi model SQLAlchemy thuc su,
# de app.py va export.py dung chung mot nguon.

BIEU_BY_KEY = {}
for s in BIEU_SCHEMAS:
    entry = dict(s)
    entry["model"] = getattr(models, s["model_name"])
    BIEU_BY_KEY[s["key"]] = entry

SAN_PHAM_ENTRY = dict(SAN_PHAM_SCHEMA)
SAN_PHAM_ENTRY["model"] = getattr(models, SAN_PHAM_SCHEMA["model_name"])

BIEU_LIST = [BIEU_BY_KEY[s["key"]] for s in BIEU_SCHEMAS]  # giu dung thu tu 9 bieu
