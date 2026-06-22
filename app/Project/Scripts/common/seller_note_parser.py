from __future__ import annotations
import re
from collections import defaultdict

SKU_LINE_RE = re.compile(r"^#?([A-Za-z0-9_.-]+)\s+(\d+(?:\.\d+)?)\s*$")
TAG_RE = re.compile(r"^#?([A-Za-z0-9_.-]+)\s*$")

def parse_seller_note(note: str) -> dict:
    """
    Seller Note固定规则：
    第1行 = Company_Code，例如 #VLG
    第2行 = Host_Name，例如 #Cammie
    第3行及以后 = SKU_ID + 销售数量，例如 #TCG00001 2

    v3规则：数量就是SKU数量，不再进行箱/盒/包/单卡折算。
    如果一个商品在不同平台以不同单位销售，需要建立不同SKU。
    """
    raw = note or ""
    lines = [x.strip() for x in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n") if x.strip()]
    result = {
        "company_code": "",
        "host_name": "",
        "items": [],
        "status": "Empty" if not lines else "Parsed",
        "note": "",
        "raw_lines": lines,
    }
    if not lines:
        result["note"] = "Seller Note为空"
        return result

    if len(lines) >= 1:
        m = TAG_RE.match(lines[0])
        result["company_code"] = m.group(1) if m else lines[0].lstrip("#")
    else:
        result["status"] = "Failed"
        result["note"] = "缺少公司代码"

    if len(lines) >= 2:
        m = TAG_RE.match(lines[1])
        result["host_name"] = m.group(1) if m else lines[1].lstrip("#")
    else:
        result["status"] = "Partial"
        result["note"] = "缺少主播名"
        return result

    sku_qty = defaultdict(float)
    raw_by_sku = defaultdict(list)
    failed_lines = []
    for line in lines[2:]:
        m = SKU_LINE_RE.match(line)
        if not m:
            failed_lines.append(line)
            continue
        sku, qty = m.group(1), float(m.group(2))
        sku_qty[sku] += qty
        raw_by_sku[sku].append(line)

    result["items"] = [
        {"sku_id": sku, "quantity": qty, "seller_note_line": " | ".join(raw_by_sku[sku])}
        for sku, qty in sku_qty.items()
    ]
    if failed_lines:
        result["status"] = "Partial" if result["items"] else "Failed"
        result["note"] = "以下SKU行无法解析：" + "; ".join(failed_lines)
    elif not result["items"]:
        result["status"] = "Partial"
        result["note"] = "没有解析到SKU销售行"
    else:
        result["note"] = "Parsed; Quantity = SKU quantity"
    return result
