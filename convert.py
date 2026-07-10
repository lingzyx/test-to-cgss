#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEST.xlsx (自建睦鄰案件資料) → CGSS 批次上傳格式 轉換工具
=========================================================

用法：
    python convert.py TEST.xlsx CGSS.xlsx output.xlsx

參數：
    TEST.xlsx   自建的睦鄰案件資料（來源檔）
    CGSS.xlsx   CGSS 批次上傳範例／模板檔（決定輸出欄位與版型）
    output.xlsx 產出的 CGSS 格式檔案

這支程式把「來源欄位可以百分之百對應」的欄位自動填入，
無法從 TEST.xlsx 判斷出來的必填欄位（例如 ApplyDate 申請日期）
一律留白，並在轉換結束後印出清單，提醒需要人工補齊的資料列。

===========================================================
★ 這份對照表是根據 CGSS.xlsx 裡「已經存在的 2 筆真實範例列」
   (第11、12列) 反推比對 TEST.xlsx 對應資料列 (第19、20列) 得出，
   而不是憑空假設，所以下列常數欄位的預設值有實際依據：
     Fund_Category (經費類別)      = 公務預算
     Fund_Name (工作/業務計畫名稱) = U200教育文化
     Organization_Type (團體類型)  = 扣繳編號(統一編號)
     CountyName (所屬地區)         = 南投縣
     IsPublic (開放查詢註記)       = 是
   如果換年度／換計畫名稱，請直接修改下面 CONFIG 區塊即可，
   不需要改動其他程式碼。
===========================================================
"""

import sys
import re
import datetime
from openpyxl import load_workbook

# ============================================================
# CONFIG — 每年 / 每個計畫只需要改這裡
# ============================================================
CONFIG = {
    "IS_PUBLIC": "是",                     # *開放查詢註記
    "FUND_CATEGORY": "公務預算",            # *經費類別
    "FUND_NAME": "U200教育文化",            # *工作/業務計畫名稱
    "ORGANIZATION_TYPE": "扣繳編號(統一編號)",  # *民間團體類型
    "COUNTY_NAME": "南投縣",                # *民間團體所歸屬之地區
    "FIRST_DATA_ROW": 11,                  # CGSS 範本資料從第幾列開始寫
    "ONLY_ROWS_MARKED_CGSS": True,          # 是否只轉換 TEST 裡「CGSS」欄有標記 v 的列
}

# TEST.xlsx 欄位位置 (1-indexed)
T_SERIAL = 1          # 序號 (e.g. 115-018)
T_MONTH = 2           # 月份 (未使用，僅供參考)
T_DOC_NO = 3          # 公服處編號115公國補字
T_ORG_NAME = 4        # 申請單位
T_ORG_ID = 5          # 統一編號
T_ACTIVITY_NAME = 6   # 活動名稱
T_ACTIVITY_PLACE = 7  # 活動地點 (目前 CGSS 範例未使用此欄)
T_START = 8           # 活動時間(起)
T_END = 9             # 活動時間(迄)
T_BUDGET_TOTAL = 10   # 活動預算金額 (目前 CGSS 範例未使用此欄)
T_SELF_FUND = 11      # 自有款
T_APPLIED_AMT = 12    # 申請補助金額
T_CONFIRMED_AMT = 13  # 核定補助金額
T_CONFIRM_DATE = 14   # 核定日期
T_CGSS_FLAG = 20      # CGSS 欄（v 表示要上傳 CGSS）

# CGSS.xlsx 欄位位置 (1-indexed, 對照第 5 列英文欄名)
C_CUSTOM_ID = 1
C_IS_PUBLIC = 4
C_APPLY_DATE = 5
C_FUND_CATEGORY = 6
C_FUND_NAME = 7
C_FUND_YEAR = 8
C_ORG_NAME = 9
C_ORG_ID = 10
C_ORG_TYPE = 11
C_COUNTY_NAME = 12
C_SUBJECT = 13
C_START_DATE = 14
C_END_DATE = 15
C_SELF_AMOUNT = 16
C_APPLIED_AMOUNT = 17
C_CONFIRM_DATE = 26
C_CONFIRM_FUND_NAME = 27
C_CONFIRM_FUND_YEAR = 28
C_CONFIRMED_AMOUNT = 29


def parse_roc_date(value):
    """把 TEST.xlsx 裡各種格式的民國日期轉成 datetime.date。
    支援: 115/2/26、115.5.04、115. 5. 19、115年3月18日 13:00-17:00、
    115/04/12(日) 08:30~12:20、已經是 datetime 物件 的情況。
    無法解析時回傳 None。"""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    s = str(value).strip()
    if not s:
        return None
    m = re.search(r"(\d{2,3})\s*[/.年]\s*(\d{1,2})\s*[/.月]\s*(\d{1,2})", s)
    if not m:
        return None
    roc_year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return datetime.date(roc_year + 1911, month, day)
    except ValueError:
        return None


def normalize_org_id(value):
    """統一編號一律轉成去除空白、補滿 8 碼的文字字串。"""
    if value is None:
        return ""
    s = str(value).strip()
    if s.isdigit():
        s = s.zfill(8)
    return s


def convert(test_path, cgss_path, output_path):
    warnings = []

    test_wb = load_workbook(test_path, data_only=True)
    test_ws = test_wb.active

    rows_out = []
    r = 2
    while True:
        serial = test_ws.cell(row=r, column=T_SERIAL).value
        if serial is None:
            # 允許中間偶有空列，但連續 3 列都空就視為資料結束
            blank_ahead = all(
                test_ws.cell(row=r + i, column=T_SERIAL).value is None
                for i in range(3)
            )
            if blank_ahead:
                break
            r += 1
            continue

        flag = test_ws.cell(row=r, column=T_CGSS_FLAG).value
        flagged = str(flag).strip().lower() == "v" if flag else False
        if CONFIG["ONLY_ROWS_MARKED_CGSS"] and not flagged:
            warnings.append(f"第{r}列（序號 {serial}）CGSS 欄未標記 v，已略過未輸出。")
            r += 1
            continue

        fund_year_match = re.match(r"^(\d+)-", str(serial).strip())
        fund_year = fund_year_match.group(1) if fund_year_match else ""

        start_date = parse_roc_date(test_ws.cell(row=r, column=T_START).value)
        end_date = parse_roc_date(test_ws.cell(row=r, column=T_END).value)
        if end_date is None and start_date is not None:
            warnings.append(
                f"第{r}列（序號 {serial}）活動時間(迄) 無法解析或含多個日期，"
                f"已暫用活動時間(起) {start_date} 作為結束日，請人工確認。"
            )
            end_date = start_date
        if start_date is None:
            warnings.append(f"第{r}列（序號 {serial}）活動時間(起) 無法解析，StartDate 留白。")

        confirm_date = parse_roc_date(test_ws.cell(row=r, column=T_CONFIRM_DATE).value)
        confirmed_amt = test_ws.cell(row=r, column=T_CONFIRMED_AMT).value
        if confirmed_amt not in (None, "") and confirm_date is None:
            warnings.append(
                f"第{r}列（序號 {serial}）已有核定補助金額但無核定日期，ConfirmDate 留白，請人工補齊。"
            )

        warnings.append(
            f"第{r}列（序號 {serial}）ApplyDate（*申請日期，必填）在 TEST.xlsx 中無對應欄位，已留白，請人工補齊。"
        )

        rows_out.append({
            "custom_id": str(serial).strip(),
            "apply_date": None,  # 無法從來源判斷，見上方 warnings
            "fund_category": CONFIG["FUND_CATEGORY"],
            "fund_name": CONFIG["FUND_NAME"],
            "fund_year": fund_year,
            "org_name": str(test_ws.cell(row=r, column=T_ORG_NAME).value or "").strip(),
            "org_id": normalize_org_id(test_ws.cell(row=r, column=T_ORG_ID).value),
            "org_type": CONFIG["ORGANIZATION_TYPE"],
            "county_name": CONFIG["COUNTY_NAME"],
            "subject": str(test_ws.cell(row=r, column=T_ACTIVITY_NAME).value or "").strip(),
            "start_date": start_date,
            "end_date": end_date,
            "self_amount": test_ws.cell(row=r, column=T_SELF_FUND).value,
            "applied_amount": test_ws.cell(row=r, column=T_APPLIED_AMT).value,
            "confirm_date": confirm_date,
            "confirmed_amount": confirmed_amt,
        })
        r += 1

    cgss_wb = load_workbook(cgss_path)
    cgss_ws = cgss_wb.active

    # 清掉範本裡原本示範用的資料列，避免與新資料混在一起
    max_existing_row = cgss_ws.max_row
    first = CONFIG["FIRST_DATA_ROW"]
    if max_existing_row >= first:
        cgss_ws.delete_rows(first, max_existing_row - first + 1)

    for i, row in enumerate(rows_out):
        rr = first + i
        cgss_ws.cell(row=rr, column=C_CUSTOM_ID, value=row["custom_id"])
        cgss_ws.cell(row=rr, column=C_IS_PUBLIC, value=CONFIG["IS_PUBLIC"])
        cgss_ws.cell(row=rr, column=C_APPLY_DATE, value=row["apply_date"])
        cgss_ws.cell(row=rr, column=C_FUND_CATEGORY, value=row["fund_category"])
        cgss_ws.cell(row=rr, column=C_FUND_NAME, value=row["fund_name"])
        cgss_ws.cell(row=rr, column=C_FUND_YEAR, value=row["fund_year"])
        cgss_ws.cell(row=rr, column=C_ORG_NAME, value=row["org_name"])
        cgss_ws.cell(row=rr, column=C_ORG_ID, value=row["org_id"])
        cgss_ws.cell(row=rr, column=C_ORG_TYPE, value=row["org_type"])
        cgss_ws.cell(row=rr, column=C_COUNTY_NAME, value=row["county_name"])
        cgss_ws.cell(row=rr, column=C_SUBJECT, value=row["subject"])
        cgss_ws.cell(row=rr, column=C_START_DATE, value=row["start_date"])
        cgss_ws.cell(row=rr, column=C_END_DATE, value=row["end_date"])
        cgss_ws.cell(row=rr, column=C_SELF_AMOUNT, value=str(row["self_amount"]) if row["self_amount"] is not None else "")
        cgss_ws.cell(row=rr, column=C_APPLIED_AMOUNT, value=str(row["applied_amount"]) if row["applied_amount"] is not None else "")
        if row["confirm_date"] is not None:
            cgss_ws.cell(row=rr, column=C_CONFIRM_DATE, value=row["confirm_date"].strftime("%Y-%m-%d"))
            cgss_ws.cell(row=rr, column=C_CONFIRM_FUND_NAME, value=row["fund_name"])
            cgss_ws.cell(row=rr, column=C_CONFIRM_FUND_YEAR, value=row["fund_year"])
        if row["confirmed_amount"] not in (None, ""):
            cgss_ws.cell(row=rr, column=C_CONFIRMED_AMOUNT, value=row["confirmed_amount"])

    cgss_wb.save(output_path)

    print(f"轉換完成：共輸出 {len(rows_out)} 筆資料到 {output_path}")
    print(f"\n共有 {len(warnings)} 項提醒，請於上傳 CGSS 前逐一確認：\n")
    for w in warnings:
        print(" -", w)

    return rows_out, warnings


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("用法: python convert.py TEST.xlsx CGSS.xlsx output.xlsx")
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2], sys.argv[3])
