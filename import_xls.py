"""
妍發科技 — XLS 資料匯入工具
將 2026.03急件.xls 和 2026當周出貨排程.xls 的資料匯入 SQLite。
"""

import xlrd
import os
import re
from datetime import datetime
import database as db

SELF_DIR = os.path.dirname(__file__)
BASE_DIR = os.path.dirname(SELF_DIR)
# 優先找同目錄，找不到再找上層目錄
_f1_local = os.path.join(SELF_DIR, "2026.03急件.xls")
_f1_parent = os.path.join(BASE_DIR, "2026.03急件.xls")
FILE1 = _f1_local if os.path.exists(_f1_local) else _f1_parent

_f2_local = os.path.join(SELF_DIR, "2026當周出貨排程.xls")
_f2_parent = os.path.join(BASE_DIR, "2026當周出貨排程.xls")
FILE2 = _f2_local if os.path.exists(_f2_local) else _f2_parent


def _xldate_to_str(value, book):
    """將 Excel 日期數值轉為 YYYY-MM-DD 字串。"""
    try:
        if isinstance(value, float) and value > 0:
            dt = xlrd.xldate_as_datetime(value, book.datemode)
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


def _cell_text(sheet, row, col):
    """安全取得儲存格文字。"""
    try:
        val = sheet.cell_value(row, col)
        if isinstance(val, float):
            if val == int(val):
                return str(int(val))
            return str(val)
        return str(val).strip() if val else ""
    except Exception:
        return ""


def _parse_amount(val):
    """嘗試解析金額數值。"""
    if isinstance(val, (int, float)) and val != 0:
        return float(val)
    if isinstance(val, str):
        cleaned = val.replace(",", "").replace("$", "").replace("¥", "").strip()
        try:
            return float(cleaned) if cleaned else None
        except ValueError:
            return None
    return None


def _parse_int(val):
    """嘗試解析整數數量。"""
    if isinstance(val, (int, float)) and val != 0:
        return int(val)
    if isinstance(val, str):
        cleaned = val.replace(",", "").strip()
        try:
            return int(float(cleaned)) if cleaned else None
        except ValueError:
            return None
    return None


# ── 已知製程名稱清單（依序號）──

KNOWN_PROCESSES = [
    "裁切", "CNC", "PTH", "線路", "假貼", "快壓", "壓合",
    "LPI", "NC", "化金", "化銀", "文字", "飛針", "雷切",
    "沖制", "檢大張", "加工小片", "SMT", "成檢", "包裝",
    "電測", "鑽孔", "外型", "防焊", "噴錫", "鍍金",
]


def import_file1_process_tracking():
    """
    匯入 2026.03急件.xls — 廠長製程追蹤表。
    橫向結構：每個料號佔 6 欄 (序號/制程/廠商/預訂排程/實際日期/空白)。
    """
    if not os.path.exists(FILE1):
        print(f"找不到檔案: {FILE1}")
        return

    book = xlrd.open_workbook(FILE1)
    imported_count = 0

    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        if sheet.nrows < 3 or sheet.ncols < 6:
            continue

        work_date = f"2026-03-{sheet_name.strip()}" if "." in sheet_name else sheet_name

        # 掃描 Row 0 找出料號位置（每 6 欄一組）
        part_blocks = []
        col = 0
        while col < sheet.ncols - 5:
            header = _cell_text(sheet, 0, col)
            if header and len(header) >= 3 and not header.startswith("序"):
                # 這是一個料號 header
                part_blocks.append((col, header))
                col += 6
            else:
                col += 1

        # 對每個料號區塊，逐行擷取製程步驟
        for start_col, raw_part_no in part_blocks:
            part_no = raw_part_no.split("---")[0].strip()
            if not part_no:
                continue

            # 先建立或找到對應的 order
            order_id = db.insert_order(
                part_no=part_no,
                source_sheet=f"急件_{sheet_name}",
                status="製程中",
            )

            # 從 Row 2 開始讀取製程步驟
            for row in range(2, min(sheet.nrows, 30)):
                seq = _cell_text(sheet, row, start_col)
                process_name = _cell_text(sheet, row, start_col + 1)
                vendor = _cell_text(sheet, row, start_col + 2)
                planned = sheet.cell_value(row, start_col + 3)
                actual = sheet.cell_value(row, start_col + 4)

                if not process_name:
                    continue

                planned_str = _xldate_to_str(planned, book)
                actual_str = _xldate_to_str(actual, book)

                # 推斷狀態
                if actual_str:
                    step_status = "完成"
                elif planned_str:
                    step_status = "進行中"
                else:
                    step_status = "待處理"

                step_seq = int(seq) if seq.isdigit() else row - 1

                db.insert_process_step(
                    order_id=order_id,
                    part_no=part_no,
                    step_seq=step_seq,
                    process_name=process_name,
                    vendor_name=vendor if vendor else None,
                    planned_date=planned_str,
                    actual_date=actual_str,
                    status=step_status,
                    work_date=work_date,
                )
                imported_count += 1

                # 同時記錄廠商
                if vendor:
                    db.upsert_vendor(vendor)

    print(f"[檔案一] 匯入 {imported_count} 筆製程步驟")
    return imported_count


def import_file2_shipping():
    """
    匯入 2026當周出貨排程.xls — 業務出貨排程。
    每個月份一張工作表，欄位結構不完全一致。
    """
    if not os.path.exists(FILE2):
        print(f"找不到檔案: {FILE2}")
        return

    book = xlrd.open_workbook(FILE2)
    imported_orders = 0
    imported_shipments = 0

    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        if sheet.nrows < 3:
            continue

        # 找出 header row（包含 "料號" 的那一行）
        header_row = None
        headers = {}
        for r in range(min(5, sheet.nrows)):
            row_texts = [_cell_text(sheet, r, c).lower() for c in range(sheet.ncols)]
            if any("料號" in t for t in row_texts):
                header_row = r
                for c, t in enumerate(row_texts):
                    raw = _cell_text(sheet, r, c)
                    headers[raw] = c
                break

        if header_row is None:
            continue

        # 建立欄位映射（適應不同月份的欄位差異）
        col_map = {}
        for name, idx in headers.items():
            lower = name.lower().strip()
            if "料號" in lower:
                col_map["part_no"] = idx
            elif "訂單" in lower:
                col_map["order_no"] = idx
            elif "數量" in lower and "出貨" not in lower:
                col_map["quantity"] = idx
            elif "金額" in lower:
                col_map["amount"] = idx
            elif "廠商" in lower:
                col_map["vendor"] = idx
            elif "預訂" in lower or "交期" in lower:
                col_map["due_date"] = idx
            elif "現制程" in lower or "制程" in lower:
                col_map["current_process"] = idx
            elif "實際出貨" in lower:
                col_map["ship_date"] = idx
            elif "回廠" in lower:
                col_map["return_date"] = idx
            elif "庫存" in lower:
                col_map["stock"] = idx
            elif "備註" in lower:
                col_map["note"] = idx

        if "part_no" not in col_map:
            continue

        # 推算 source_month
        source_month = sheet_name.strip()

        # 逐行讀取資料
        for row in range(header_row + 1, sheet.nrows):
            part_no = _cell_text(sheet, row, col_map["part_no"])
            if not part_no or part_no in ("小計", "合計", "總計", "total"):
                continue
            # 跳過小計行
            if any(kw in part_no for kw in ["小計", "合計", "月份", "訂單金額"]):
                continue

            order_no = _cell_text(sheet, row, col_map["order_no"]) if "order_no" in col_map else None
            quantity = _parse_int(sheet.cell_value(row, col_map["quantity"])) if "quantity" in col_map else None
            amount = _parse_amount(sheet.cell_value(row, col_map["amount"])) if "amount" in col_map else None
            vendor = _cell_text(sheet, row, col_map["vendor"]) if "vendor" in col_map else None
            due_date = _xldate_to_str(sheet.cell_value(row, col_map["due_date"]), book) if "due_date" in col_map else None
            current_process = _cell_text(sheet, row, col_map["current_process"]) if "current_process" in col_map else None
            ship_date = _xldate_to_str(sheet.cell_value(row, col_map["ship_date"]), book) if "ship_date" in col_map else None
            note = _cell_text(sheet, row, col_map["note"]) if "note" in col_map else None

            # 推斷狀態
            if ship_date:
                status = "已出貨"
            elif current_process and any(kw in (note or "") for kw in ["暫停", "暫緩"]):
                status = "客戶暫停"
            elif note and "待零件" in note:
                status = "待零件"
            else:
                status = "製程中"

            # 如果是「未出貨」工作表，標記為特殊狀態
            if sheet_name.strip() == "未出貨":
                if not note:
                    note = "未出貨清單"
                if status == "製程中":
                    status = "延遲"

            order_id = db.insert_order(
                order_no=order_no,
                part_no=part_no,
                quantity=quantity,
                amount=amount,
                vendor_name=vendor,
                due_date=due_date,
                status=status,
                source_sheet=f"出貨_{source_month}",
                note=f"{current_process or ''} {note or ''}".strip(),
            )
            imported_orders += 1

            # 如果有實際出貨日期，也建立出貨記錄
            if ship_date:
                db.insert_shipment(
                    order_id=order_id,
                    part_no=part_no,
                    ship_date=ship_date,
                    ship_quantity=quantity,
                    amount=amount,
                    note=note,
                )
                imported_shipments += 1

            # 記錄廠商
            if vendor:
                db.upsert_vendor(vendor)

    print(f"[檔案二] 匯入 {imported_orders} 筆訂單, {imported_shipments} 筆出貨記錄")
    return imported_orders, imported_shipments


def run_full_import():
    """執行完整匯入流程。"""
    print("=" * 50)
    print("妍發科技 — XLS 資料匯入")
    print("=" * 50)

    # 初始化資料庫
    db.init_db()
    print(f"✓ 資料庫已初始化: {db.DB_PATH}")

    # 匯入檔案一
    print("\n[1] 匯入製程追蹤表...")
    import_file1_process_tracking()

    # 匯入檔案二
    print("\n[2] 匯入出貨排程...")
    import_file2_shipping()

    # 顯示統計
    stats = db.get_dashboard_stats()
    print("\n" + "=" * 50)
    print("匯入完成！統計：")
    print(f"  訂單總數: {stats['total_orders']}")
    print(f"  製程中:   {stats['in_progress']}")
    print(f"  已出貨:   {stats['shipped']}")
    print(f"  延遲:     {stats['delayed']}")
    print(f"  暫停:     {stats['on_hold']}")
    print(f"  總金額:   ${stats['total_amount']:,.0f}")
    print("=" * 50)


if __name__ == "__main__":
    run_full_import()
