import streamlit as st
import anthropic
import base64
import json
import os
import io
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from copy import copy

# ─────────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="経費精算ツール",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
.main { max-width: 720px; }
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #4f7ef7, #7c4dff);
    color: white;
    font-weight: 600;
    font-size: 16px;
    padding: 12px;
    border-radius: 10px;
    border: none;
}
.stDownloadButton > button {
    width: 100%;
    background: #22c55e;
    color: white;
    font-weight: 600;
    font-size: 16px;
    padding: 12px;
    border-radius: 10px;
    border: none;
}
.result-box {
    background: #1e1e2e;
    color: #a9b1d6;
    border-radius: 10px;
    padding: 16px;
    font-family: monospace;
    font-size: 13px;
    white-space: pre-wrap;
}
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = Path(__file__).parent / "template.xlsx"
MEDIA_TYPES = {
    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
    '.gif': 'image/gif', '.webp': 'image/webp', '.pdf': 'application/pdf',
}


# ─────────────────────────────────────────────
# Claude API で領収書を解析
# ─────────────────────────────────────────────
def extract_receipt_data(client, file_bytes: bytes, filename: str) -> dict:
    ext = Path(filename).suffix.lower()
    media_type = MEDIA_TYPES.get(ext, 'image/jpeg')
    img_data = base64.standard_b64encode(file_bytes).decode('utf-8')

    prompt = """この領収書・レシート画像から経費情報をJSONで抽出してください。

以下のJSON形式のみを返してください（コードブロック不要）:
{"type":"travel","date":"YYYY-MM-DD","amount":金額,"category":"科目","description":"摘要","route":"経路またはnull"}

typeの選び方:
- 電車・新幹線・バス・タクシーなど交通系 → "travel"
- 接待・会食・備品・宿泊など → "other"

categoryの例:
- travel: 「旅費」(新幹線・特急など)、「交通費」(電車・バス・タクシー)
- other: 「接待交際費」「会議費」「消耗品費」「宿泊費」「諸経費」

routeは交通費のみ（例:品川→熱海）。その他はnull。
日付不明はnull、金額不明は0。JSONのみ返してください。"""

    resp = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=512,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_data}},
            {"type": "text", "text": prompt}
        ]}]
    )

    text = resp.content[0].text.strip()
    if '```' in text:
        for part in text.split('```'):
            part = part.strip()
            if part.startswith('json'):
                part = part[4:].strip()
            try:
                return json.loads(part)
            except Exception:
                continue
    return json.loads(text)


# ─────────────────────────────────────────────
# Excel生成
# ─────────────────────────────────────────────
def write_sheet(ws, items: list, data_start: int, data_end: int, has_route: bool):
    for r in range(data_start, data_end + 1):
        for c in range(1, 12):
            ws.cell(row=r, column=c).value = None

    for i, item in enumerate(items):
        row = data_start + i
        ws.cell(row=row, column=1).value = '=ROW()-12'

        if item.get('date'):
            try:
                ws.cell(row=row, column=2).value = datetime.strptime(item['date'], '%Y-%m-%d')
                ws.cell(row=row, column=2).number_format = 'm/d'
            except Exception:
                ws.cell(row=row, column=2).value = item['date']

        amt = item.get('amount', 0)
        ws.cell(row=row, column=3).value = amt if amt else None
        ws.cell(row=row, column=3).number_format = '[\u00a5-411]#,##0'
        ws.cell(row=row, column=4).value = item.get('category', '')
        ws.cell(row=row, column=5).value = item.get('description', '')

        if has_route:
            ws.cell(row=row, column=8).value = item.get('route') or ''

    ws['D8'] = len(items)


def generate_excel(travel_items: list, other_items: list) -> bytes:
    wb = load_workbook(str(TEMPLATE_PATH))
    ws_t = wb['旅費交通費']
    ws_o = wb['その他']

    write_sheet(ws_t, travel_items, 13, 26, True)
    write_sheet(ws_o, other_items, 13, 19, False)

    today = datetime.now()
    ws_t['D4'] = today
    ws_o['D4'] = today

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📊 経費精算ツール")
st.caption("領収書をアップロードするだけで、Excelの経費精算書を自動作成します")

st.divider()

# APIキー入力
st.subheader("🔑 Anthropic APIキー")

# 環境変数（Streamlit Cloudでの secrets 対応）
env_key = os.environ.get('ANTHROPIC_API_KEY', '') or st.secrets.get('ANTHROPIC_API_KEY', '') if hasattr(st, 'secrets') else ''
if env_key:
    api_key = env_key
    st.success("✅ APIキーが設定済みです（環境変数）")
else:
    api_key = st.text_input(
        "APIキーを入力してください",
        type="password",
        placeholder="sk-ant-...",
        help="[Anthropic Console](https://console.anthropic.com) でAPIキーを取得できます"
    )

st.divider()

# ファイルアップロード
st.subheader("📎 領収書をアップロード")
uploaded_files = st.file_uploader(
    "領収書画像・PDFをアップロード（複数選択可）",
    type=['jpg', 'jpeg', 'png', 'gif', 'webp', 'pdf'],
    accept_multiple_files=True,
    help="対応形式: JPG, PNG, PDF, WEBP, GIF"
)

if uploaded_files:
    st.write(f"📁 **{len(uploaded_files)} ファイル**が選択されています:")
    cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with cols[i % 4]:
            ext = Path(f.name).suffix.lower()
            if ext in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
                st.image(f, caption=f.name, use_container_width=True)
            else:
                st.markdown(f"📄 `{f.name}`")

st.divider()

# 処理ボタン
if st.button("⚡ Excel経費精算書を作成する", disabled=not (api_key and uploaded_files)):

    if not TEMPLATE_PATH.exists():
        st.error("テンプレートファイル (template.xlsx) が見つかりません。")
        st.stop()

    try:
        client = anthropic.Anthropic(api_key=api_key)
    except Exception as e:
        st.error(f"APIキーが無効です: {e}")
        st.stop()

    travel_items = []
    other_items = []
    log_lines = []

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"🔍 解析中... {uploaded_file.name} ({i+1}/{len(uploaded_files)})")
        progress_bar.progress((i) / len(uploaded_files))

        try:
            file_bytes = uploaded_file.read()
            data = extract_receipt_data(client, file_bytes, uploaded_file.name)

            amt = data.get('amount', 0)
            amt_str = f"¥{int(amt):,}" if isinstance(amt, (int, float)) and amt else "¥0"
            log_line = f"✅ {uploaded_file.name}\n   → {data.get('date', '日付不明')} | {data.get('category', '不明')} | {amt_str}"
            if data.get('route'):
                log_line += f" | {data.get('route')}"
            log_lines.append(log_line)

            if data.get('type') == 'travel':
                travel_items.append(data)
            else:
                other_items.append(data)

        except json.JSONDecodeError as e:
            log_lines.append(f"⚠️ {uploaded_file.name}\n   → JSON解析エラー: {e}")
        except Exception as e:
            log_lines.append(f"❌ {uploaded_file.name}\n   → エラー: {str(e)[:80]}")

    progress_bar.progress(1.0)
    status_text.text("✨ Excel生成中...")

    # 日付ソート
    travel_items.sort(key=lambda x: x.get('date') or '9999-99-99')
    other_items.sort(key=lambda x: x.get('date') or '9999-99-99')

    try:
        excel_bytes = generate_excel(travel_items, other_items)
        status_text.empty()
        progress_bar.empty()

        # 結果表示
        st.success(f"✅ 完了！ 旅費交通費: **{len(travel_items)}件** / その他: **{len(other_items)}件**")

        # ログ
        with st.expander("📋 処理ログを見る", expanded=False):
            st.code('\n\n'.join(log_lines), language=None)

        # ダウンロードボタン
        today_str = datetime.now().strftime('%Y%m')
        st.download_button(
            label="📥 Excelファイルをダウンロード",
            data=excel_bytes,
            file_name=f"{today_str}経費精算.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        status_text.empty()
        st.error(f"Excel生成エラー: {e}")

st.divider()
st.markdown(
    '<p style="text-align:center; color:#aaa; font-size:12px;">Powered by Claude AI · 領収書の情報は外部に保存されません</p>',
    unsafe_allow_html=True
)

