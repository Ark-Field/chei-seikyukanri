import io
import os
import calendar
import requests
import pandas as pd
from datetime import datetime, timedelta, date
import streamlit as st

# ReportLab関連のインポート（PDF出力用）
from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# ==========================================
# 1. ページ基本設定 ＆ デザインCSS
# ==========================================
st.set_page_config(page_title="期間集計・締日請求書発行システム", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    h1, h2, h3 { color: #0F172A !important; font-family: 'Helvetica Neue', Arial, sans-serif; }
    div.stButton > button:first-child {
        background-color: #1E3A8A; color: white; border-radius: 6px; border: none; font-weight: bold; padding: 0.5rem 1rem;
    }
    div.stButton > button:first-child:hover { background-color: #1E40AF; color: white; }
    div.stDownloadButton > button:first-child {
        background-color: #0F172A; color: white; border-radius: 6px; border: none; font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🖨️ 期間集計・締日請求書発行（安定化対策版）")

# ==========================================
# 2. 日本語フォント登録処理（ReportLab用）
# ==========================================
def setup_japanese_font():
    local_font_path = "NotoSansJP-Regular.ttf"
    if os.path.exists(local_font_path):
        try:
            pdfmetrics.registerFont(TTFont("JPFont", local_font_path))
            return "JPFont"
        except Exception as e:
            print(f"同梱フォント登録エラー: {e}")

    win_font = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "msgothic.ttc")
    if os.path.exists(win_font):
        try:
            pdfmetrics.registerFont(TTFont("JPFont", win_font, subfontIndex=0))
            return "JPFont"
        except Exception as e:
            print(f"Windowsフォント登録エラー: {e}")

    return "Helvetica"

FONT_NAME = setup_japanese_font()

# ==========================================
# 3. Secrets設定
# ==========================================
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID = 34  # @pocket対象アプリID

# ==========================================
# 4. 便利関数
# ==========================================
def clean_val(val):
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ["none", "nan", "nat", "null"]:
        return ""
    return s

def parse_date_safe(val):
    s = clean_val(val)
    if not s:
        return None
    try:
        if isinstance(val, (date, datetime)):
            return val.date() if isinstance(val, datetime) else val
        return datetime.strptime(s[:10].replace("/", "-"), "%Y-%m-%d").date()
    except ValueError:
        return None

# --- 期間集計請求書用PDF生成（18行パディング＆罫線対応） ---
def generate_period_invoice_pdf(pub_name, items, period_str, greeting_text, customer_remark, bank_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4), leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle("Title", parent=styles["Normal"], fontName=FONT_NAME, fontSize=18, leading=20, alignment=1, textColor=colors.HexColor("#1E3A8A"))
    h_style = ParagraphStyle("H", parent=styles["Normal"], fontName=FONT_NAME, fontSize=9.5, leading=13, textColor=colors.HexColor("#333333"))
    th_style = ParagraphStyle("TH", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8, leading=10, textColor=colors.white, alignment=1)
    td_style = ParagraphStyle("TD", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8, leading=11, textColor=colors.HexColor("#1E293B"))
    td_right = ParagraphStyle("TDR", parent=td_style, alignment=2)

    elements = []
    elements.append(Paragraph("<b>請 求 書</b>", title_style))
    elements.append(Spacer(1, 10))

    meta_data = [
        [
            Paragraph(f"<b>{pub_name if pub_name else '（出版社名未設定）'} 御中</b>", ParagraphStyle("CName", parent=h_style, fontSize=13, leading=15)),
            Paragraph("<b>(株)中永製本所</b><br/>埼玉県さいたま市岩槻区真福寺1400<br/>TEL 048-798-7377", ParagraphStyle("SInfo", parent=h_style, alignment=2))
        ],
        [
            Paragraph(f"<b>対象期間:</b> {period_str}<br/>{greeting_text}", h_style),
            Paragraph(f"発行日: {date.today().strftime('%Y年%m月%d日')}", ParagraphStyle("DateInfo", parent=h_style, alignment=2))
        ]
    ]
    meta_table = Table(meta_data, colWidths=[300, 245])
    meta_table.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("BOTTOMPADDING", (0,0), (-1,-1), 4)]))
    elements.append(meta_table)
    elements.append(Spacer(1, 8))

    subtotal = 0
    for item in items:
        q_num = int(item["請求数量"]) if str(item["請求数量"]).isdigit() else 0
        p_num = float(item["単価"]) if item["単価"] else 0.0
        subtotal += q_num * p_num
    
    tax_amount = round(subtotal * 0.10)
    grand_total = subtotal + tax_amount

    summary_text = f"<b>ご請求金額合計（税込）: ￥{grand_total:,.0f}</b> （うち消費税額等: ￥{tax_amount:,.0f}）"
    summary_table = Table([[Paragraph(summary_text, h_style)]], colWidths=[545])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F8F9FA")),
        ("BOX", (0,0), (-1,-1), 1, colors.HexColor("#CED4DA")),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 10), ("RIGHTPADDING", (0,0), (-1,-1), 10),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 8))

    table_data = [[
        Paragraph("売上日", th_style),
        Paragraph("ISBN", th_style),
        Paragraph("品名", th_style),
        Paragraph("判型/版", th_style),
        Paragraph("数量", th_style),
        Paragraph("単価", th_style),
        Paragraph("金額", th_style),
        Paragraph("備考", th_style)
    ]]

    for item in items:
        q_num = int(item["請求数量"]) if str(item["請求数量"]).isdigit() else 0
        p_num = float(item["単価"]) if item["単価"] else 0.0
        line_total = q_num * p_num if (q_num > 0 and p_num > 0) else ""
        format_edition = f"{item['判型']} / {item['版']}" if (item['判型'] or item['版']) else ""

        qty_display = f"{q_num} 部" if q_num > 0 else ""
        price_display = f"￥{p_num:,.2f}" if p_num > 0 else ""
        total_display = f"￥{line_total:,.2f}" if isinstance(line_total, (int, float)) else ""

        table_data.append([
            Paragraph(str(item["売上日"]), td_style),
            Paragraph(str(item["ISBN"]), td_style),
            Paragraph(str(item["品名"]), td_style),
            Paragraph(format_edition, td_style),
            Paragraph(qty_display, td_right),
            Paragraph(price_display, td_right),
            Paragraph(total_display, td_right),
            Paragraph(str(item["備考"]), td_style)
        ])

    min_rows = 18
    current_row_count = len(items)
    if current_row_count < min_rows:
        for _ in range(min_rows - current_row_count):
            table_data.append([
                Paragraph("", td_style),
                Paragraph("", td_style),
                Paragraph("", td_style),
                Paragraph("", td_style),
                Paragraph("", td_right),
                Paragraph("", td_right),
                Paragraph("", td_right),
                Paragraph("", td_style)
            ])

    inv_table = Table(table_data, colWidths=[60, 70, 140, 55, 45, 60, 65, 50])
    inv_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1,-1), 4),
        ("TOPPADDING", (0, 0), (-1,-1), 4),
        ("GRID", (0, 0), (-1,-1), 0.5, colors.HexColor("#94A3B8")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#1E3A8A")),
    ]))
    elements.append(inv_table)
    elements.append(Spacer(1, 10))

    footer_content = []
    if customer_remark:
        remark_text = f"<b>【備考】</b><br/>{customer_remark.replace(chr(10), '<br/>')}"
        footer_content.append([Paragraph(remark_text, td_style)])
    
    bank_text = f"<b>【お振込先口座】</b><br/>{bank_info['bank_name']} {bank_info['bank_branch']} ({bank_info['bank_type']}) {bank_info['bank_account']}<br/>口座名義: {bank_info['bank_holder']}"
    footer_content.append([Paragraph(bank_text, td_style)])

    footer_table = Table(footer_content, colWidths=[545])
    footer_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F8FAFC")),
        ("BOX", (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 8), ("RIGHTPADDING", (0,0), (-1,-1), 8),
    ]))
    elements.append(footer_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. データ取得関数（タイムアウト30秒に拡張 ＆ 安全なフォールバック）
# ==========================================
def fetch_pocket_records_for_period(target_shimebi):
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}
    
    # 締日による絞り込みクエリ
    query_str = f'field-99 = "{target_shimebi}" order by field-1 desc' if target_shimebi else "order by field-1 desc"
    params = {"query": query_str}
    
    try:
        st.info(f"📡 サーバーからデータを取得中（締日: {target_shimebi} / タイムアウト30秒）...")
        # タイムアウトを 30秒 に戻して確実に受け取る
        response = requests.get(url, headers=headers, params=params, timeout=30)
        
        # クエリで万が一エラーになる場合はクエリなしで安全に再取得
        if response.status_code != 200:
            response = requests.get(url, headers=headers, timeout=30)

        if response.status_code == 200:
            result = response.json()
            raw_items = result.get("records", result.get("Data", result.get("Records", [])))
            parsed_records = []
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                rec_body = item.get("record", {})
                if not isinstance(rec_body, dict):
                    rec_body = {}
                parsed_records.append({**item, **rec_body})
            st.success(f"✨ 取得成功: 合計 {len(parsed_records)} 件のレコードを読み込みました。")
            return parsed_records
        else:
            st.error(f"❌ APIエラー ({response.status_code}): {response.text}")
    except Exception as e:
        st.error(f"🚨 通信例外エラー（タイムアウト等）: {e}")
    return []

# ==========================================
# 6. メイン画面
# ==========================================
st.info("💡 期間・締日を設定し、必要に応じて挨拶文や顧客ごとの備考を入力してPDFを作成できます。")

today_date = date.today()
c_year, c_month = today_date.year, today_date.month
auto_start = date(c_year, c_month, 1)
auto_last_day = calendar.monthrange(c_year, c_month)[1]
auto_end = date(c_year, c_month, auto_last_day)

col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
    shimebi_options = {"末日締め (99)": "99", "20日締め (20)": "20", "15日締め (15)": "15", "10日締め (10)": "10"}
    selected_shimebi_label = st.selectbox("対象の締日グループ", options=list(shimebi_options.keys()), key="inv_shime_sel")
    target_shimebi = shimebi_options[selected_shimebi_label]
with col_c2:
    inv_start_date = st.date_input("集計期間（開始）", value=auto_start, key="inv_start")
with col_c3:
    inv_end_date = st.date_input("集計期間（終了）", value=auto_end, key="inv_end")

default_greeting = "下記のとおりご請求申し上げます。"
greeting_input = st.text_input("📝 PDFヘッダーの挨拶文（編集可能）", value=default_greeting, key="pdf_greeting")

st.markdown("#### 🏦 請求書記載の銀行口座情報")
bc1, bc2, bc3 = st.columns(3)
with bc1:
    p_bank_name = st.text_input("銀行名", value="三井住友銀行", key="p_b_name")
    p_bank_branch = st.text_input("支店名", value="岩槻支店", key="p_b_branch")
with bc2:
    p_bank_type = st.selectbox("口座種別", options=["普通", "当座"], key="p_b_type")
    p_bank_account = st.text_input("口座番号", value="1234567", key="p_b_acc")
with bc3:
    p_bank_holder = st.text_input("口座名義", value="カ）チュウエイセイホンジョ", key="p_b_holder")

if st.button("🔍 条件に合致するデータを集計する", type="primary", key="btn_calc_period_inv"):
    all_recs = fetch_pocket_records_for_period(target_shimebi)
    period_invoice_rows = []

    for r in all_recs:
        pub_name = clean_val(r.get("field-8", ""))
        shimebi_val = clean_val(r.get("field-99", ""))
        book_title = clean_val(r.get("field-13", ""))
        format_val = clean_val(r.get("field-23", ""))
        edition_val = clean_val(r.get("field-27", ""))

        sub_89 = r.get("field-89", [])
        if isinstance(sub_89, list):
            for item_89 in sub_89:
                rec_89 = item_89.get("record", item_89) if isinstance(item_89, dict) else {}
                sale_date_raw = parse_date_safe(rec_89.get("field-89_1"))
                q_str = clean_val(rec_89.get("field-89_3", "0"))

                if not sale_date_raw or not q_str or q_str == "0":
                    continue

                if inv_start_date <= sale_date_raw <= inv_end_date:
                    p_str = clean_val(r.get("field-5", "0"))
                    period_invoice_rows.append({
                        "出版社名": pub_name if pub_name else "（出版社名未設定）",
                        "締日": shimebi_val,
                        "売上日": sale_date_raw,
                        "ISBN": clean_val(r.get("field-93", "")),
                        "品名": book_title,
                        "判型": format_val,
                        "版": edition_val,
                        "請求数量": q_str,
                        "単価": p_str,
                        "備考": clean_val(rec_89.get("field-89_7", ""))
                    })

    st.session_state["period_invoice_rows"] = period_invoice_rows

if "period_invoice_rows" in st.session_state and st.session_state["period_invoice_rows"]:
    df_period_inv = pd.DataFrame(st.session_state["period_invoice_rows"])
    st.success(f"✨ 期間集計対象の有効な明細が **{len(df_period_inv)} 件** 見つかりました。")

    from collections import defaultdict
    grouped_period_inv = defaultdict(list)
    for r in st.session_state["period_invoice_rows"]:
        grouped_period_inv[r["出版社名"]].append(r)

    st.markdown("### 📊 顧客別売上額集計 ＆ 個別備考の編集")
    st.info("💡 各顧客ごとにPDFへ記載する「備考」を入力できます。")

    if "customer_remarks" not in st.session_state:
        st.session_state["customer_remarks"] = {}

    for pub, items in grouped_period_inv.items():
        subtotal = sum((int(i["請求数量"]) if str(i["請求数量"]).isdigit() else 0) * (float(i["単価"]) if i["単価"] else 0.0) for i in items)
        tax = round(subtotal * 0.10)
        total = subtotal + tax
        
        if pub not in st.session_state["customer_remarks"]:
            st.session_state["customer_remarks"][pub] = ""

        col_sum1, col_sum2 = st.columns([2, 3])
        with col_sum1:
            st.markdown(f"**🏢 {pub}**")
            st.write(f"明細件数: {len(items)}件 / 合計(税込): **￥{total:,.0f}**")
        with col_sum2:
            st.session_state["customer_remarks"][pub] = st.text_area(
                f"📝 {pub} の請求書備考",
                value=st.session_state["customer_remarks"][pub],
                key=f"remark_input_{pub}",
                height=70
            )
        st.markdown("---")

    st.subheader("📥 顧客ごとの個別請求書PDFダウンロード")

    period_str_header = f"{inv_start_date.strftime('%Y年%m月%d日')} ～ {inv_end_date.strftime('%Y年%m月%d日')}"
    bank_dict = {
        "bank_name": p_bank_name, "bank_branch": p_bank_branch,
        "bank_type": p_bank_type, "bank_account": p_bank_account, "bank_holder": p_bank_holder
    }

    for pub, items in grouped_period_inv.items():
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            st.write(f"🏢 **{pub}** （明細数: {len(items)}件）")
        with col_p2:
            customer_remark_text = st.session_state["customer_remarks"].get(pub, "")
            pdf_bytes = generate_period_invoice_pdf(pub, items, period_str_header, greeting_input, customer_remark_text, bank_dict)
            st.download_button(
                label="🖨️ PDFダウンロード",
                data=pdf_bytes,
                file_name=f"請求書_{pub}_{inv_start_date.strftime('%Y%m%d')}-{inv_end_date.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key=f"dl_period_inv_{pub}"
            )
elif "period_invoice_rows" in st.session_state:
    st.warning("⚠️ 指定された条件に一致する有効な売上データが見つかりませんでした。")