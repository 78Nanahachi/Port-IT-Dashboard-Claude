"""
sheets_to_json.py
Google Sheets の2つのシート（LIST・Technology）を取得し、
data.json に統合して出力するスクリプト。

必要な環境変数:
  SHEETS_API_KEY  ... Google Sheets API キー（GitHub Secrets に登録）

ローカル実行:
  export SHEETS_API_KEY="your_api_key_here"
  python .github/scripts/sheets_to_json.py
"""

import os
import json
import re
import sys
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

# ─── 設定 ──────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1jGre5314DttRSk2_xXVlFc0XxgroxKJMOVonhSTaLu0"

SHEETS = {
    "list": {
        "gid": "0",
        "range": "LIST!A:G",       # 検索日|国|カテゴリ|元タイトル|発信日|Gemini要約|URL
        "source": "news",
    },
    "technology": {
        "gid": "785681052",
        "range": "Technology!A:F",  # 発行日|発行月|原文タイトル|Edition|要約|URL
        "source": "magazine",
    },
}

OUTPUT_PATH = "data.json"

# カテゴリ文字列 → cat キー・日本語ラベルのマッピング（LIST シート用）
CATEGORY_MAP = [
    (["smart port", "digital port", "port digitali", "digital twin", "modernization",
      "cảng thông minh", "dx", "smartport"],
     "smart", "スマートポート"),
    (["shore power", "cold ironing", "port microgrid", "port energy"],
     "shore", "ショアパワー"),
    (["green port", "decarboniz", "net zero", "carbon neutral", "sustainability",
      "green", "hydrogen"],
     "green", "グリーン/脱炭素"),
    (["terminal operating", "tos", "port automation", "port management", "ai terminal",
      "automation"],
     "automation", "自動化/TOS"),
    (["port development", "port expansion", "tender", "ppp", "investment", "concession"],
     "dev", "港湾開発"),
    (["port regulations", "maritime policy", "port authority"],
     "policy", "規制/政策"),
]

def resolve_category(raw: str):
    """カテゴリ文字列から (cat, catLabel) を返す。マッチしなければ other。"""
    lower = raw.lower()
    for keywords, cat, label in CATEGORY_MAP:
        if any(k in lower for k in keywords):
            return cat, label
    return "other", raw[:30] if raw else "その他"


def fetch_sheet_values(api_key: str, range_name: str) -> list[list[str]]:
    """Sheets API v4 で指定レンジの値を取得する。"""
    base = "https://sheets.googleapis.com/v4/spreadsheets"
    params = urllib.parse.urlencode({
        "range": range_name,
        "valueRenderOption": "UNFORMATTED_VALUE",
        "key": api_key,
    })
    url = f"{base}/{SPREADSHEET_ID}/values/{urllib.parse.quote(range_name)}?{params}"

    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            data = json.loads(resp.read().decode())
            return data.get("values", [])
    except Exception as e:
        print(f"[ERROR] Sheets API 取得失敗 ({range_name}): {e}", file=sys.stderr)
        sys.exit(1)


def normalize_date(raw) -> str:
    """
    様々な日付フォーマットを YYYY-MM-DD に統一する。
    Sheets の数値シリアル値（例: 46000）にも対応。
    """
    if not raw:
        return ""
    s = str(raw).strip()

    # Sheets のシリアル値（整数）
    if re.fullmatch(r"\d{4,6}", s):
        serial = int(s)
        if 40000 < serial < 60000:  # 2009〜2064年の範囲
            epoch = datetime(1899, 12, 30, tzinfo=timezone.utc)
            dt = epoch + timedelta(days=serial)
            return dt.strftime("%Y-%m-%d")

    # YYYY/MM/DD HH:MM
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"

    return s[:10]  # フォールバック：先頭10文字


def parse_list_sheet(rows: list[list]) -> list[dict]:
    """
    LIST シートの行を統一フォーマットの辞書リストに変換する。
    列順: 検索日(A) | 国(B) | カテゴリ(C) | 元タイトル(D) | 発信日(E) | 要約(F) | URL(G)
    """
    results = []
    for i, row in enumerate(rows[1:], start=2):  # 1行目はヘッダー
        # 空行スキップ
        if not row or all(str(c).strip() == "" for c in row):
            continue

        def col(n, default=""):
            return str(row[n]).strip() if n < len(row) else default

        title   = col(3)
        summary = col(5)
        url     = col(6)
        if not title and not summary:
            continue

        raw_cat = col(2)
        cat, cat_label = resolve_category(raw_cat)

        # 日付は「発信日(E)」を優先、なければ「検索日(A)」
        date_raw = col(4) or col(0)
        date = normalize_date(date_raw)

        results.append({
            "date":     date,
            "country":  col(1) or "ASEAN",
            "cat":      cat,
            "catLabel": cat_label,
            "title":    title,
            "summary":  summary,
            "url":      url,
            "source":   "news",
        })

    return results


def parse_technology_sheet(rows: list[list]) -> list[dict]:
    """
    Technology シートの行を統一フォーマットの辞書リストに変換する。
    列順: 発行日(A) | 発行月(B) | 原文タイトル(C) | Edition(D) | 要約(E) | URL(F)
    """
    results = []
    for i, row in enumerate(rows[1:], start=2):
        if not row or all(str(c).strip() == "" for c in row):
            continue

        def col(n, default=""):
            return str(row[n]).strip() if n < len(row) else default

        title   = col(2)
        summary = col(4)
        url     = col(5)
        if not title and not summary:
            continue

        edition = col(3)  # "Edition 123" などの文字列

        # 日付: 発行日(A) → 発行月(B) の順で取得
        date = normalize_date(col(0)) or normalize_date(col(1))

        results.append({
            "date":     date,
            "country":  "GLOBAL",         # 誌面なので国は GLOBAL
            "cat":      "magazine",
            "catLabel": "Port Technology誌",
            "edition":  edition,           # LIST にはないフィールドを追加
            "title":    title,
            "summary":  summary,
            "url":      url,
            "source":   "magazine",
        })

    return results


def assign_ids(records: list[dict]) -> list[dict]:
    """日付降順にソートして連番 id を振る。"""
    sorted_records = sorted(
        records,
        key=lambda r: r.get("date", ""),
        reverse=True,
    )
    for i, r in enumerate(sorted_records, start=1):
        r["id"] = i
    return sorted_records


def main():
    api_key = os.environ.get("SHEETS_API_KEY", "").strip()
    if not api_key:
        print("[ERROR] 環境変数 SHEETS_API_KEY が設定されていません。", file=sys.stderr)
        print("  export SHEETS_API_KEY='your_key_here'", file=sys.stderr)
        sys.exit(1)

    print("▶ LIST シートを取得中...")
    list_rows = fetch_sheet_values(api_key, SHEETS["list"]["range"])
    list_records = parse_list_sheet(list_rows)
    print(f"  → {len(list_records)} 件取得")

    print("▶ Technology シートを取得中...")
    tech_rows = fetch_sheet_values(api_key, SHEETS["technology"]["range"])
    tech_records = parse_technology_sheet(tech_rows)
    print(f"  → {len(tech_records)} 件取得")

    # マージ → ID 採番
    all_records = assign_ids(list_records + tech_records)
    total = len(all_records)
    print(f"▶ 合計 {total} 件を {OUTPUT_PATH} に書き出し中...")

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)

    jst = timezone(timedelta(hours=9))
    now = datetime.now(jst).strftime("%Y-%m-%d %H:%M JST")
    print(f"✅ 完了 ({now}) — {total} 件")


if __name__ == "__main__":
    main()
