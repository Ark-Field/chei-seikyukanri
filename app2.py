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
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# OpenPyXL関連のインポート（Excel出力用）
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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

st.title("🖨️ **期間集計・締日請求書発行（）**")

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

def calc_invoice_totals(items):
    subtotal = 0.0
    for item in items:
        q_num = int(item["請求数量"]) if str(item["請求数量"]).isdigit() else 0
        p_num = float(item["単価"]) if item["単価"] else 0.0
        subtotal += q_num * p_num
    tax_amount = round(subtotal * 0.10)
    grand_total = subtotal + tax_amount
    return subtotal, tax_amount, grand_total

# --- 期間集計請求書用PDF生成（A4縦・自動改ページ・モダンデザイン） ---
def generate_period_invoice_pdf(pub_name, items, period_str, greeting_text, customer_remark, bank_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A4),
        leftMargin=22,
        rightMargin=22,
        topMargin=22,
        bottomMargin=22
    )
    styles = getSampleStyleSheet()

    NAVY = colors.HexColor("#0F172A")
    ACCENT = colors.HexColor("#0284C7")
    ACCENT_SOFT = colors.HexColor("#F0F9FF")
    SLATE_BORDER = colors.HexColor("#CBD5E1")
    ROW_ALT = colors.HexColor("#F8FAFC")
    TEXT_MAIN = colors.HexColor("#1E293B")
    TEXT_MUTED = colors.HexColor("#64748B")

    banner_title = ParagraphStyle("BannerTitle", parent=styles["Normal"], fontName=FONT_NAME, fontSize=18, leading=22, textColor=colors.white)
    banner_sub = ParagraphStyle("BannerSub", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8.5, leading=11, textColor=colors.HexColor("#BAE6FD"), alignment=2)
    client_style = ParagraphStyle("Client", parent=styles["Normal"], fontName=FONT_NAME, fontSize=13, leading=17, textColor=NAVY)
    sender_style = ParagraphStyle("Sender", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8.5, leading=12, textColor=TEXT_MAIN, alignment=2)
    meta_style = ParagraphStyle("Meta", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8.5, leading=12, textColor=TEXT_MAIN)
    meta_right = ParagraphStyle("MetaR", parent=meta_style, alignment=2, textColor=TEXT_MUTED)

    kpi_value = ParagraphStyle("KPIValue", parent=styles["Normal"], fontName=FONT_NAME, fontSize=13.5, leading=17, textColor=NAVY)
    kpi_accent = ParagraphStyle("KPIAccent", parent=styles["Normal"], fontName=FONT_NAME, fontSize=10, leading=14, textColor=ACCENT, alignment=2)

    th_style = ParagraphStyle("TH", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7.5, leading=10, textColor=colors.white, alignment=1)
    td_style = ParagraphStyle("TD", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7.5, leading=10, textColor=TEXT_MAIN)
    td_center = ParagraphStyle("TDC", parent=td_style, alignment=1)
    td_right = ParagraphStyle("TDR", parent=td_style, alignment=2)
    foot_style = ParagraphStyle("Foot", parent=styles["Normal"], fontName=FONT_NAME, fontSize=8, leading=11.5, textColor=TEXT_MAIN)

    elements = []

    banner_data = [[
        Paragraph("<b>請 求 書</b>", banner_title),
        Paragraph("INVOICE / 期間集計明細", banner_sub)
    ]]
    banner_table = Table(banner_data, colWidths=[351, 200])
    banner_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("LINEBELOW", (0, 0), (-1, -1), 3, ACCENT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LEFTPADDING", (0, 0), (0, 0), 14),
        ("RIGHTPADDING", (-1, 0), (-1, 0), 14),
    ]))
    elements.append(banner_table)
    elements.append(Spacer(1, 8))

    client_display = pub_name if pub_name else "（出版社名未設定）"
    info_data = [
        [
            Paragraph(f"<b>{client_display}</b> <font size='10' color='#64748B'>御中</font>", client_style),
            Paragraph(
                "<b>(株)中永製本所</b><br/>"
                "<font color='#64748B'>〒339-0078 埼玉県さいたま市岩槻区真福寺1400<br/>TEL: 048-798-7377</font>",
                sender_style
            )
        ],
        [
            Paragraph(f"<b>対象期間：</b>{period_str}<br/>{greeting_text}", meta_style),
            Paragraph(f"発行日：{date.today().strftime('%Y年%m月%d日')}", meta_right)
        ]
    ]
    info_table = Table(info_data, colWidths=[331, 220])
    info_table.setStyle(TableStyle([
        ("LINEBEFORE", (0, 0), (0, 0), 3, ACCENT),
        ("LEFTPADDING", (0, 0), (0, 0), 10),
        ("LEFTPADDING", (0, 1), (0, 1), 10),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 1), (-1, 1), 0.75, SLATE_BORDER),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 8))

    subtotal, tax_amount, grand_total = calc_invoice_totals(items)

    summary_data = [[
        Paragraph(
            f"<font size='8' color='#64748B'>ご請求金額合計（税込）</font><br/>"
            f"<b>￥{grand_total:,.0f}</b>",
            kpi_value
        ),
        Paragraph(
            f"小計（税抜）：<b>￥{subtotal:,.0f}</b> ｜ 消費税（10%）：<b>￥{tax_amount:,.0f}</b>",
            kpi_accent
        )
    ]]
    summary_table = Table(summary_data, colWidths=[250, 301])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ACCENT_SOFT),
        ("BOX", (0, 0), (-1, -1), 0.75, ACCENT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (0, 0), 12),
        ("RIGHTPADDING", (-1, 0), (-1, 0), 12),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 8))

    col_widths = [58, 74, 144, 56, 44, 60, 65, 46]
    table_data = [[
        Paragraph("売上日", th_style),
        Paragraph("ISBN", th_style),
        Paragraph("品名", th_style),
        Paragraph("判型 / 版", th_style),
        Paragraph("数量", th_style),
        Paragraph("単価", th_style),
        Paragraph("金額", th_style),
        Paragraph("備考", th_style),
    ]]

    for item in items:
        q_num = int(item["請求数量"]) if str(item["請求数量"]).isdigit() else 0
        p_num = float(item["単価"]) if item["単価"] else 0.0
        has_calc = (q_num > 0 and p_num > 0)
        line_total = q_num * p_num if has_calc else None

        format_edition = f"{item['判型']} / {item['版']}" if (item["判型"] and item["版"]) else (item["判型"] or item["版"])
        qty_display = f"{q_num:,}" if q_num > 0 else ""
        price_display = f"￥{p_num:,.2f}" if p_num > 0 else ""
        total_display = f"￥{line_total:,.0f}" if line_total is not None else ""

        table_data.append([
            Paragraph(str(item["売上日"]), td_center),
            Paragraph(str(item["ISBN"]), td_style),
            Paragraph(str(item["品名"]), td_style),
            Paragraph(format_edition, td_center),
            Paragraph(qty_display, td_right),
            Paragraph(price_display, td_right),
            Paragraph(f"<b>{total_display}</b>" if total_display else "", td_right),
            Paragraph(str(item["備考"]), td_style),
        ])

    min_rows = 10
    if len(items) < min_rows:
        for _ in range(min_rows - len(items)):
            table_data.append([
                Paragraph("&nbsp;", td_center),
                Paragraph("", td_style),
                Paragraph("", td_style),
                Paragraph("", td_center),
                Paragraph("", td_right),
                Paragraph("", td_right),
                Paragraph("", td_right),
                Paragraph("", td_style),
            ])

    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("LINEBELOW", (0, 0), (-1, 0), 1.5, ACCENT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, SLATE_BORDER),
        ("BOX", (0, 0), (-1, -1), 0.8, NAVY),
        ("LINEBEFORE", (1, 0), (-1, -1), 0.35, colors.HexColor("#E2E8F0")),
    ]

    for row_idx in range(2, len(table_data), 2):
        table_styles.append(("BACKGROUND", (0, row_idx), (-1, row_idx), ROW_ALT))

    inv_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    inv_table.setStyle(TableStyle(table_styles))
    elements.append(inv_table)
    elements.append(Spacer(1, 10))

    footer_rows = []
    if customer_remark:
        remark_html = f"<font color='#0284C7'><b>【備考】</b></font><br/>{customer_remark.replace(chr(10), '<br/>')}"
        footer_rows.append([Paragraph(remark_html, foot_style)])

    bank_html = (
        f"<font color='#0284C7'><b>【お振込先口座】</b></font><br/>"
        f"<b>{bank_info['bank_name']} {bank_info['bank_branch']}</b>（{bank_info['bank_type']}）"
        f"口座番号：<b>{bank_info['bank_account']}</b> ／ "
        f"口座名義：{bank_info['bank_holder']}"
    )
    footer_rows.append([Paragraph(bank_html, foot_style)])

    footer_table = Table(footer_rows, colWidths=[551])
    footer_styles = [
        ("BACKGROUND", (0, 0), (-1, -1), ROW_ALT),
        ("BOX", (0, 0), (-1, -1), 0.75, SLATE_BORDER),
        ("LINEBEFORE", (0, 0), (0, -1), 3, NAVY),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]
    if len(footer_rows) == 2:
        footer_styles.append(("LINEBELOW", (0, 0), (-1, 0), 0.5, SLATE_BORDER))

    footer_table.setStyle(TableStyle(footer_styles))
    elements.append(footer_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- 期間集計請求書用Excel生成（モダンデザイン・A4縦印刷対応） ---
def generate_period_invoice_excel(pub_name, items, period_str, greeting_text, customer_remark, bank_info):
    wb = Workbook()
    ws = wb.active
    ws.title = "請求書"
    ws.views.sheetView[0].showGridLines = True

    font_family = "Yu Gothic UI" if os.name == "nt" else "Arial"

    f_banner = Font(name=font_family, size=16, bold=True, color="FFFFFF")
    f_banner_sub = Font(name=font_family, size=9, bold=True, color="BAE6FD")
    f_client = Font(name=font_family, size=13, bold=True, color="0F172A")
    f_muted = Font(name=font_family, size=9, color="64748B")
    f_body = Font(name=font_family, size=9.5, color="1E293B")
    f_body_bold = Font(name=font_family, size=9.5, bold=True, color="0F172A")
    f_kpi_label = Font(name=font_family, size=9, bold=True, color="0284C7")
    f_kpi_val = Font(name=font_family, size=14, bold=True, color="0F172A")
    f_th = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")

    fill_navy = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    fill_accent_soft = PatternFill(start_color="F0F9FF", end_color="F0F9FF", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_slate = Side(style="thin", color="CBD5E1")
    med_navy = Side(style="medium", color="0F172A")
    thick_accent = Side(style="thick", color="0284C7")

    b_header = Border(top=med_navy, bottom=thick_accent, left=med_navy, right=med_navy)
    b_kpi = Border(top=thin_slate, bottom=thin_slate, left=thick_accent, right=thin_slate)
    b_th = Border(top=med_navy, bottom=thick_accent, left=thin_slate, right=thin_slate)
    b_cell = Border(top=thin_slate, bottom=thin_slate, left=thin_slate, right=thin_slate)
    b_foot = Border(top=thin_slate, bottom=thin_slate, left=thick_accent, right=thin_slate)

    # 1. バナー (A1:H2)
    ws.merge_cells("A1:F2")
    ws.merge_cells("G1:H2")
    ws["A1"] = "請 求 書"
    ws["G1"] = "INVOICE / 期間集計明細"
    for r in (1, 2):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_navy
            cell.border = b_header
            cell.font = f_banner if c <= 6 else f_banner_sub
            cell.alignment = Alignment(
                horizontal="left" if c <= 6 else "right",
                vertical="center",
                indent=1 if c == 1 else 0
            )
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18

    # 2. 宛先・発行元・期間 (A4:H5)
    client_display = pub_name if pub_name else "（出版社名未設定）"
    ws.merge_cells("A4:E4")
    ws.merge_cells("F4:H4")
    ws["A4"] = f"{client_display} 御中"
    ws["F4"] = "(株)中永製本所"
    ws["A4"].font = f_client
    ws["F4"].font = f_body_bold
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")
    ws["F4"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[4].height = 24

    ws.merge_cells("A5:E5")
    ws.merge_cells("F5:H5")
    ws["A5"] = f"対象期間：{period_str} {greeting_text}"
    ws["F5"] = f"発行日：{date.today().strftime('%Y年%m月%d日')} TEL: 048-798-7377"
    ws["A5"].font = f_body
    ws["F5"].font = f_muted
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center")
    ws["F5"].alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 9):
        ws.cell(row=5, column=c).border = Border(bottom=thin_slate)
    ws.row_dimensions[5].height = 20

    # 3. 金額サマリー (A7:H8)
    subtotal, tax_amount, grand_total = calc_invoice_totals(items)
    ws.merge_cells("A7:C7")
    ws.merge_cells("D7:H7")
    ws["A7"] = "ご請求金額合計（税込）"
    ws["D7"] = f"小計（税抜）：￥{subtotal:,.0f} ／ 消費税（10%）：￥{tax_amount:,.0f}"
    ws["A7"].font = f_kpi_label
    ws["D7"].font = f_body_bold
    ws["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["D7"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A8:C8")
    ws.merge_cells("D8:H8")
    ws["A8"] = grand_total
    ws["A8"].number_format = "￥#,##0"
    ws["A8"].font = f_kpi_val
    ws["A8"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["D8"] = "〒339-0078 埼玉県さいたま市岩槻区真福寺1400"
    ws["D8"].font = f_muted
    ws["D8"].alignment = Alignment(horizontal="right", vertical="center")

    for r in (7, 8):
        ws.row_dimensions[r].height = 20 if r == 7 else 24
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_accent_soft
            cell.border = b_kpi
    ws.row_dimensions[9].height = 10

    # 4. 明細テーブルヘッダー (Row 10)
    headers = ["売上日", "ISBN", "品名", "判型 / 版", "数量", "単価", "金額", "備考"]
    header_row = 10
    ws.row_dimensions[header_row].height = 24
    for idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=idx, value=h_text)
        c.font = f_th
        c.fill = fill_navy
        c.border = b_th
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    current_row = header_row + 1
    for i_idx, item in enumerate(items):
        q_num = int(item["請求数量"]) if str(item["請求数量"]).isdigit() else 0
        p_num = float(item["単価"]) if item["単価"] else 0.0
        has_calc = (q_num > 0 and p_num > 0)
        line_total = q_num * p_num if has_calc else None
        format_edition = f"{item['判型']} / {item['版']}" if (item["判型"] and item["版"]) else (item["判型"] or item["版"])

        row_fill = fill_zebra if i_idx % 2 == 1 else fill_white
        ws.row_dimensions[current_row].height = 20

        vals = [
            str(item["売上日"]),
            str(item["ISBN"]),
            str(item["品名"]),
            format_edition,
            q_num if q_num > 0 else None,
            p_num if p_num > 0 else None,
            line_total if line_total is not None else None,
            str(item["備考"]),
        ]

        for col_i, val in enumerate(vals, start=1):
            cell = ws.cell(row=current_row, column=col_i, value=val)
            cell.fill = row_fill
            cell.border = b_cell
            cell.font = f_body_bold if col_i == 7 else f_body

            if col_i in (1, 4):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif col_i in (5, 6, 7):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if col_i == 5 and val is not None:
                    cell.number_format = "#,##0"
                elif col_i == 6 and val is not None:
                    cell.number_format = "￥#,##0.00"
                elif col_i == 7 and val is not None:
                    cell.number_format = "￥#,##0"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        current_row += 1

    # 最低10行確保
    min_rows = 10
    if len(items) < min_rows:
        for pad_i in range(min_rows - len(items)):
            row_fill = fill_zebra if (len(items) + pad_i) % 2 == 1 else fill_white
            ws.row_dimensions[current_row].height = 20
            for col_i in range(1, 9):
                cell = ws.cell(row=current_row, column=col_i, value="")
                cell.fill = row_fill
                cell.border = b_cell
                cell.font = f_body
            current_row += 1

    current_row += 1

    # 5. 備考・振込先フッター
    if customer_remark:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
        rm_cell = ws.cell(row=current_row, column=1, value=f"【備考】{customer_remark}")
        rm_cell.font = f_body
        rm_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[current_row].height = max(24, 16 * (customer_remark.count("\n") + 1))
        for col_i in range(1, 9):
            ws.cell(row=current_row, column=col_i).fill = fill_zebra
            ws.cell(row=current_row, column=col_i).border = b_foot
        current_row += 1

    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
    bank_text = (
        f"【お振込先口座】{bank_info['bank_name']} {bank_info['bank_branch']}（{bank_info['bank_type']}）"
        f" 口座番号：{bank_info['bank_account']} ／ 口座名義：{bank_info['bank_holder']}"
    )
    bk_cell = ws.cell(row=current_row, column=1, value=bank_text)
    bk_cell.font = f_body_bold
    bk_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[current_row].height = 24
    for col_i in range(1, 9):
        ws.cell(row=current_row, column=col_i).fill = fill_zebra
        ws.cell(row=current_row, column=col_i).border = b_foot

    # 列幅設定（A4縦に収まるバランス）
    col_widths_excel = {
        "A": 13,
        "B": 16,
        "C": 30,
        "D": 14,
        "E": 10,
        "F": 13,
        "G": 14,
        "H": 16,
    }
    for col_letter, width in col_widths_excel.items():
        ws.column_dimensions[col_letter].width = width

    # 印刷設定（A4縦・1ページ幅フィット・改ページ時ヘッダー繰返し）
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{header_row}:{header_row}"

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf

# ==========================================
# 5. データ取得関数（タイムアウト45秒に拡張 ＆ 安全なフォールバック）
# ==========================================
def fetch_pocket_records_for_period(target_shimebi):
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}
    
    query_str = f'field-99 = "{target_shimebi}" order by field-1 desc' if target_shimebi else "order by field-1 desc"
    params = {"query": query_str}
    
    try:
        st.info(f"📡 サーバーからデータを取得中（締日: **{target_shimebi}** / タイムアウト **45秒**）...")
        response = requests.get(url, headers=headers, params=params, timeout=45)
        
        if response.status_code != 200:
            response = requests.get(url, headers=headers, timeout=45)

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
            st.success(f"✨ 取得成功: 合計 **{len(parsed_records)} 件** のレコードを読み込みました。")
            return parsed_records
        else:
            st.error(f"❌ APIエラー (**{response.status_code}**): {response.text}")
    except Exception as e:
        st.error(f"🚨 通信例外エラー（タイムアウト等）: **{e}**")
    return []

# ==========================================
# 6. メイン画面
# ==========================================
st.info("💡 **期間・締日**を設定し、必要に応じて**挨拶文**や**顧客ごとの備考**を入力して**PDF / Excel**を作成できます。")

today_date = date.today()
c_year, c_month = today_date.year, today_date.month
auto_start = date(c_year, c_month, 1)
auto_last_day = calendar.monthrange(c_year, c_month)[1]
auto_end = date(c_year, c_month, auto_last_day)

col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
    shimebi_options = {"末日締め (99)": "99", "20日締め (20)": "20", "15日締め (15)": "15", "10日締め (10)": "10"}
    selected_shimebi_label = st.selectbox("**対象の締日グループ**", options=list(shimebi_options.keys()), key="inv_shime_sel")
    target_shimebi = shimebi_options[selected_shimebi_label]
with col_c2:
    inv_start_date = st.date_input("**集計期間（開始）**", value=auto_start, key="inv_start")
with col_c3:
    inv_end_date = st.date_input("**集計期間（終了）**", value=auto_end, key="inv_end")

default_greeting = "下記のとおりご請求申し上げます。"
greeting_input = st.text_input("📝 **PDF・Excelヘッダーの挨拶文（編集可能）**", value=default_greeting, key="pdf_greeting")

st.markdown("#### 🏦 **請求書記載の銀行口座情報**")
bc1, bc2, bc3 = st.columns(3)
with bc1:
    p_bank_name = st.text_input("**銀行名**", value="三井住友銀行", key="p_b_name")
    p_bank_branch = st.text_input("**支店名**", value="岩槻支店", key="p_b_branch")
with bc2:
    p_bank_type = st.selectbox("**口座種別**", options=["普通", "当座"], key="p_b_type")
    p_bank_account = st.text_input("**口座番号**", value="1234567", key="p_b_acc")
with bc3:
    p_bank_holder = st.text_input("**口座名義**", value="カ）チュウエイセイホンジョ", key="p_b_holder")

if st.button("🔍 **条件に合致するデータを集計する**", type="primary", key="btn_calc_period_inv"):
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

    st.markdown("### 📊 **顧客別売上額集計 ＆ 個別備考の編集**")
    st.info("💡 各顧客ごとに請求書へ記載する **「備考」** を入力できます。")

    if "customer_remarks" not in st.session_state:
        st.session_state["customer_remarks"] = {}

    for pub, items in grouped_period_inv.items():
        _, _, total = calc_invoice_totals(items)
        
        if pub not in st.session_state["customer_remarks"]:
            st.session_state["customer_remarks"][pub] = ""

        col_sum1, col_sum2 = st.columns([2, 3])
        with col_sum1:
            st.markdown(f"**🏢 {pub}**")
            st.write(f"明細件数: **{len(items)}件** / 合計(税込): **￥{total:,.0f}**")
        with col_sum2:
            st.session_state["customer_remarks"][pub] = st.text_area(
                f"📝 **{pub} の請求書備考**",
                value=st.session_state["customer_remarks"][pub],
                key=f"remark_input_{pub}",
                height=70
            )
        st.markdown("---")

    st.subheader("📥 **顧客ごとの個別請求書ダウンロード（PDF / Excel）**")

    period_str_header = f"{inv_start_date.strftime('%Y年%m月%d日')} ～ {inv_end_date.strftime('%Y年%m月%d日')}"
    bank_dict = {
        "bank_name": p_bank_name, "bank_branch": p_bank_branch,
        "bank_type": p_bank_type, "bank_account": p_bank_account, "bank_holder": p_bank_holder
    }

    for pub, items in grouped_period_inv.items():
        col_p1, col_p2, col_p3 = st.columns([3, 1, 1])
        with col_p1:
            st.write(f"🏢 **{pub}** （明細数: **{len(items)}件**）")
        customer_remark_text = st.session_state["customer_remarks"].get(pub, "")
        with col_p2:
            pdf_bytes = generate_period_invoice_pdf(pub, items, period_str_header, greeting_input, customer_remark_text, bank_dict)
            st.download_button(
                label="🖨️ **PDF**",
                data=pdf_bytes,
                file_name=f"請求書_{pub}_{inv_start_date.strftime('%Y%m%d')}-{inv_end_date.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key=f"dl_period_inv_{pub}"
            )
        with col_p3:
            excel_bytes = generate_period_invoice_excel(pub, items, period_str_header, greeting_input, customer_remark_text, bank_dict)
            st.download_button(
                label="📊 **Excel**",
                data=excel_bytes,
                file_name=f"請求書_{pub}_{inv_start_date.strftime('%Y%m%d')}-{inv_end_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_period_inv_xlsx_{pub}"
            )
elif "period_invoice_rows" in st.session_state:
    st.warning("⚠️ **指定された条件に一致する有効な売上データが見つかりませんでした。**")