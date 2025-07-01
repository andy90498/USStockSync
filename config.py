from pathlib import Path
import os

APPDATA = Path(os.getenv("APPDATA")) / "USStockSync"
EXCEL_PATH = APPDATA / "爬蟲更新區.xlsm"
CREDENTIALS_PATH = APPDATA / "credentials.json"
BASE = Path(__file__).parent