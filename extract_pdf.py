#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次辨識《睦鄰案件申請書》PDF 裡的「申請單位聲明事項」頁
==========================================================

這支程式把資料夾裡每一份 PDF 的每一頁轉成圖片，整份送給 Claude
（不用先猜「申請單位聲明事項」在第幾頁），請它找出那一頁並整理成
結構化 JSON，最後把所有案件彙整成一份 all_agencies.json / .csv。

用法：
    export ANTHROPIC_API_KEY=你的金鑰   (Mac/Linux)
    set ANTHROPIC_API_KEY=你的金鑰       (Windows)

    python extract_pdf.py ./pdfs ./output

參數：
    ./pdfs    放置所有睦鄰案件 PDF 的資料夾（檔名建議用序號命名，如 115-005.pdf）
    ./output  結果輸出資料夾

⚠️ 注意事項：
- 這支程式會呼叫 Anthropic API，屬於「用量計費」，一份 PDF（約 9~10 頁）
  單次辨識大約會用到數千個 token，實際費用請自行到
  https://console.anthropic.com 查閱當時的定價。
- 請絕對不要把 API 金鑰寫死在程式碼裡，也不要把 .env 或金鑰檔案上傳到
  GitHub（本專案已附 .gitignore 排除這些檔案）。
"""

import os
import sys
import re
import json
import base64
import glob

import pypdfium2 as pdfium
import anthropic

MODEL = "claude-sonnet-5"

EXTRACTION_PROMPT = """\
這是一份台灣睦鄰組織補助案件的申請書 PDF（已轉成一系列頁面圖片）。
請在所有頁面中找到標題為「申請單位聲明事項」的那一頁（通常是經費預算表
附表的最後一段文字，內容會列出統一編號、自有款、總經費，以及是否有
向政府機關、私校、國內團體或個人申請補助）。

只根據「申請單位聲明事項」這頁的內容，輸出以下 JSON（不要有任何其他文字、
不要用 ```json 包住，直接輸出純 JSON）：

{
  "organization_id": "統一編號，8碼字串",
  "self_amount": "自有款金額，純數字字串，例如 12000",
  "total_budget": "總經費金額，純數字字串",
  "taipower_amount": "本案向台灣電力股份有限公司申請的金額，純數字字串",
  "other_agencies": [
    {"name": "機關名稱", "amount": "擬申請金額，純數字字串"}
  ]
}

如果某個欄位在文件中找不到，該欄位請填 null。
如果沒有「其他機關」這一項（第2點完全空白或未填寫任何機關），
other_agencies 請回傳空陣列 []。
"""


def pdf_to_images_b64(pdf_path, max_pages=15, scale=2.0):
    """把 PDF 每一頁轉成 base64 PNG，回傳 list[str]。"""
    pdf = pdfium.PdfDocument(pdf_path)
    images = []
    n = min(len(pdf), max_pages)
    for i in range(n):
        page = pdf[i]
        bitmap = page.render(scale=scale)
        pil_image = bitmap.to_pil()
        import io
        buf = io.BytesIO()
        pil_image.save(buf, format="PNG")
        images.append(base64.standard_b64encode(buf.getvalue()).decode("utf-8"))
    return images


def extract_one(client, pdf_path):
    images_b64 = pdf_to_images_b64(pdf_path)
    content = []
    for img_b64 in images_b64:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img_b64},
        })
    content.append({"type": "text", "text": EXTRACTION_PROMPT})

    response = client.messages.create(
        model=MODEL,
        max_tokens=1024,
        messages=[{"role": "user", "content": content}],
    )
    text = "".join(b.text for b in response.content if b.type == "text").strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text.strip())
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"_raw_response": text, "_parse_error": True}


def main():
    if len(sys.argv) != 3:
        print("用法: python extract_pdf.py ./pdfs ./output")
        sys.exit(1)

    pdf_folder, out_folder = sys.argv[1], sys.argv[2]
    os.makedirs(out_folder, exist_ok=True)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("錯誤：找不到 ANTHROPIC_API_KEY 環境變數，請先設定你的 API 金鑰。")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    pdf_files = sorted(glob.glob(os.path.join(pdf_folder, "*.pdf")))
    if not pdf_files:
        print(f"在 {pdf_folder} 裡找不到任何 PDF 檔案。")
        sys.exit(1)

    results = []
    for path in pdf_files:
        serial = os.path.splitext(os.path.basename(path))[0]
        print(f"辨識中：{serial} ...", end=" ", flush=True)
        try:
            data = extract_one(client, path)
            data["serial"] = serial
            data["source_file"] = os.path.basename(path)
            results.append(data)
            print("完成")
        except Exception as e:
            print(f"失敗：{e}")
            results.append({"serial": serial, "source_file": os.path.basename(path), "_error": str(e)})

    json_path = os.path.join(out_folder, "all_agencies.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    csv_path = os.path.join(out_folder, "all_agencies.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write("序號,統一編號,自有款,總經費,台電申請金額,他機關名稱,他機關金額\n")
        for r in results:
            if r.get("_error") or r.get("_parse_error"):
                f.write(f"{r.get('serial','')},辨識失敗,,,,,\n")
                continue
            agencies = r.get("other_agencies") or []
            if not agencies:
                f.write(f"{r.get('serial','')},{r.get('organization_id','')},"
                        f"{r.get('self_amount','')},{r.get('total_budget','')},"
                        f"{r.get('taipower_amount','')},,\n")
            for a in agencies:
                f.write(f"{r.get('serial','')},{r.get('organization_id','')},"
                        f"{r.get('self_amount','')},{r.get('total_budget','')},"
                        f"{r.get('taipower_amount','')},{a.get('name','')},{a.get('amount','')}\n")

    print(f"\n共處理 {len(results)} 份 PDF")
    print(f"結果已輸出：\n  {json_path}\n  {csv_path}")
    print("\n請打開 CSV 檢查一次辨識結果是否正確，尤其是手寫金額或印章壓字的部分，")
    print("確認無誤後可以用 merge_agencies.py 把這份資料併入 CGSS 輸出檔。")


if __name__ == "__main__":
    main()
