from __future__ import annotations
from datetime import datetime
from zoneinfo import ZoneInfo

FORMATS = [
    "%m/%d/%Y %I:%M:%S %p",  # TikTok US 12-hour, e.g. 05/24/2026 9:58:16 PM
    "%m/%d/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d",
]

def clean_time_text(x: str) -> str:
    return (x or "").replace("\t", "").strip()

def parse_datetime_to_standard(raw_time: str, source_tz="America/Los_Angeles", business_tz="America/Los_Angeles") -> dict:
    txt = clean_time_text(raw_time)
    if not txt:
        return {"local": "", "utc": "", "business": "", "status": "Empty", "note": "empty time"}
    dt = None
    used_fmt = ""
    for fmt in FORMATS:
        try:
            dt = datetime.strptime(txt, fmt)
            used_fmt = fmt
            break
        except ValueError:
            pass
    if dt is None:
        return {"local": txt, "utc": "", "business": "", "status": "Failed", "note": "unrecognized time format"}
    local_dt = dt.replace(tzinfo=ZoneInfo(source_tz))
    utc_dt = local_dt.astimezone(ZoneInfo("UTC"))
    business_dt = utc_dt.astimezone(ZoneInfo(business_tz))
    fmt_out = "%Y-%m-%d %H:%M:%S"
    return {
        "local": local_dt.strftime(fmt_out),
        "utc": utc_dt.strftime(fmt_out),
        "business": business_dt.strftime(fmt_out),
        "status": "Parsed",
        "note": f"parsed by {used_fmt}; source_tz={source_tz}",
    }
