from pathlib import Path
from order_import.import_tiktok_order import import_tiktok_order

ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / "data" / "Project"
input_csv = BASE / "Raw_Files" / "TikTok_Order" / "Allorder-2026-05-26-13_20.csv"
output_xlsx = BASE / "Data" / "order_import_result_v3.xlsx"
import_tiktok_order(str(input_csv), str(output_xlsx))
