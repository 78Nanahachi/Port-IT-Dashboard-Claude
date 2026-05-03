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
        "range": "List!A:G",        # 検索日|国|カテゴリ|元タイトル|発信日|Gemini要約|URL
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


def fetch_sheet_values(api_key: str, range_name: str,
                       render: str = "UNFORMATTED_VALUE") -> list[list[str]]:
    """
    Sheets API v4 で指定レンジの値を取得する。
    render: "UNFORMATTED_VALUE" 日付をシリアル値/ISO形式で取得（デフォルト）
            "FORMULA"           数式をそのまま取得（HYPERLINK等）
            "FORMATTED_VALUE"   表示テキストを取得
    """
    base = "https://sheets.googleapis.com/v4/spreadsheets"
    params = urllib.parse.urlencode({
        "range": range_name,
        "valueRenderOption": render,
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

    Sheets API UNFORMATTED_VALUE が返す形式:
      - ISO 8601文字列: "2018-09-04T07:00:00.000Z"  ← 今回のケース
      - シリアル値(小数): 46142.5328
      - シリアル値(整数): 46142
      - 通常文字列: "2026-04-18" / "2026/04/18"
    """
    if raw is None or str(raw).strip() == "":
        return ""

    s = str(raw).strip()

    # ① ISO 8601形式: "2018-09-04T07:00:00.000Z" → 先頭10文字で確定
    if re.match(r"\d{4}-\d{2}-\d{2}T", s):
        return s[:10]

    # ② 数値シリアル値（整数・小数）: 40000〜60000 の範囲 = 2009〜2064年
    try:
        serial = float(s)
        if 40000 < serial < 60000:
            epoch = datetime(1899, 12, 30, tzinfo=timezone.utc)
            dt = epoch + timedelta(days=serial)
            return dt.strftime("%Y-%m-%d")
    except ValueError:
        pass

    # ③ YYYY-MM-DD / YYYY/MM/DD
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"

    # ④ フォールバック：先頭10文字
    return s[:10]


def normalize_month(raw) -> str:
    """
    発行月セルを YYYY-MM 形式に変換する（Technology シート専用）。
    例: "2026-04-01T..." → "2026-04"
        46142.5          → "2026-04"
        "2026/04"        → "2026-04"
        "April 2026"     → "2026-04"
    """
    full = normalize_date(raw)  # まず YYYY-MM-DD に変換
    if full and len(full) >= 7:
        return full[:7]         # 先頭7文字 = YYYY-MM
    return full


def resolve_url(raw: str, fallback: str = "") -> str:
    """
    セルの値からURLを抽出して返す。
    スプレッドシートに混在する以下の形式すべてに対応:
      - 生URL:          "https://example.com"
      - HYPERLINK数式:  '=HYPERLINK("https://example.com","LINK")'
      - 表示テキストのみ: "LINK" → Sheets APIがHYPERLINK数式の表示テキストだけ返す場合
      - 空・"#":        fallback を返す（デフォルトは空文字）
    """
    if not raw:
        return fallback
    s = str(raw).strip()

    # 空・無効値
    if s in ("#", "", "LINK", "link", "Link"):
        return fallback

    # HYPERLINK数式から生URLを抽出
    # =HYPERLINK("https://...", "LINK") -> https://...
    m = re.match(r'=HYPERLINK\(\s*"([^"]+)"', s, re.IGNORECASE)
    if m:
        extracted = m.group(1)
        return extracted if extracted else fallback

    # https:// で始まる通常のURL
    if s.startswith("http"):
        return s

    return fallback


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
        url     = resolve_url(col(6))
        if not title and not summary:
            continue

        raw_cat = col(2)
        cat, cat_label = resolve_category(raw_cat)

        # 日付は「発信日(E)」を優先、なければ「検索日(A)」
        # col() は str を返すが、シリアル値は数値型で来るため raw アクセスも使う
        date_raw = row[4] if len(row) > 4 and row[4] != "" else (row[0] if row else "")
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
        edition = col(3)

        # URLの解決:
        #   Editionあり（雑誌記事）→ 常に editions 一覧ページ（個別URLは無視）
        #   Editionなし（サイト記事）→ HYPERLINKのURLを使用
        if edition:
            url = "https://www.porttechnology.org/editions/"
        else:
            url = resolve_url(col(5), fallback="")

        if not title and not summary:
            continue

        # 日付: 発行月(B) のみ使用。A列（検索日）は無視する。YYYY-MM形式で保存
        date_raw_b = row[1] if len(row) > 1 and row[1] != "" else None
        date = normalize_month(date_raw_b)

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
    # 日付等は UNFORMATTED_VALUE、URL列(F列)は FORMULA で別取得してマージ
    tech_rows       = fetch_sheet_values(api_key, SHEETS["technology"]["range"])
    tech_url_rows   = fetch_sheet_values(api_key, "Technology!F:F", render="FORMULA")
    # URL列をマージ: tech_rows の各行の6列目(index 5)を FORMULA 取得値で上書き
    for i, row in enumerate(tech_rows):
        formula_val = tech_url_rows[i][0] if i < len(tech_url_rows) and tech_url_rows[i] else ""
        if len(row) < 6:
            row.extend([""] * (6 - len(row)))
        row[5] = formula_val  # F列(index 5)を数式値で上書き
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
