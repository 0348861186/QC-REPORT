import streamlit as st
import pandas as pd
import io
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =========================
# HELPER: BILINGUAL TEXT (STACKED)
# =========================
def tr(en, vi):
    return f"{en}\n\n{vi}"

# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="Customer Complain Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title(tr(
    "📊 CUSTOMER COMPLAIN DASHBOARD",
    "📊 BẢNG ĐIỀU KHIỂN KHIẾU NẠI KHÁCH HÀNG"
))
st.markdown("---")

# =========================
# API KEY HANDLING (FIXED + DEBUG)
# =========================
st.sidebar.markdown(tr("### 🤖 SYSTEM CONFIG", "### 🤖 CẤU HÌNH HỆ THỐNG"))

api_key_sidebar = st.sidebar.text_input(
    tr("API Key (optional)", "API Key (không bắt buộc)"),
    type="password",
    help=tr("Enter Google AI Studio API key", "Nhập API key từ Google AI Studio")
)

api_key = None

if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("✔ API Key loaded từ secrets.toml")
elif api_key_sidebar:
    api_key = api_key_sidebar
    st.sidebar.success("✔ API Key nhập từ sidebar")
else:
    st.sidebar.warning("⚠ Chưa có API_KEY")

st.sidebar.caption(f"Debug API key status: {'OK' if api_key else 'MISSING'}")

# =========================
# UPLOAD FILE
# =========================
uploaded_files = st.sidebar.file_uploader(
    tr("Upload Excel/CSV", "Tải lên Excel/CSV"),
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

# =========================
# PROCESS FILE
# =========================
def process_file_by_position(file):
    if file.name.endswith(('.xlsx', '.xls')):
        raw_df = pd.read_excel(file, header=None)
    else:
        raw_df = pd.read_csv(file, header=None)

    start_row = 0
    for idx, row in raw_df.iterrows():
        row_str = row.astype(str).str.upper().str.strip().tolist()
        if any('COMPLAIN' in str(cell) for cell in row_str):
            start_row = idx + 1
            break

    df = raw_df.iloc[start_row:].copy()
    df = df.iloc[:, :7]

    df.columns = [
        'COMPLAIN_DATE',
        'WO_NUM',
        'ITEM',
        'ORDER_QTY',
        'NG_QTY',
        'TYPE_OF_DEFECT',
        'FACILITY'
    ]

    return df.dropna(how='all')

# =========================
# LOAD DATA
# =========================
def load_data(files):
    all_df = []

    for file in files:
        df = process_file_by_position(file)
        df['SOURCE_FILE_NAME'] = file.name
        all_df.append(df)

    if not all_df:
        return pd.DataFrame()

    df = pd.concat(all_df, ignore_index=True)

    df['COMPLAIN_DATE'] = pd.to_datetime(
        df['COMPLAIN_DATE'],
        format='%d/%m/%Y',
        errors='coerce'
    )

    df = df.dropna(subset=['COMPLAIN_DATE'])

    df['ORDER_QTY'] = pd.to_numeric(df['ORDER_QTY'], errors='coerce').fillna(0)
    df['NG_QTY'] = pd.to_numeric(df['NG_QTY'], errors='coerce').fillna(0)
    df['FACILITY'] = df['FACILITY'].astype(str).str.strip()
    df['TYPE_OF_DEFECT'] = df['TYPE_OF_DEFECT'].astype(str).str.strip()

    return df

# =========================
# PDF GENERATION HELPER (ALL 4 CHARTS MATCHING DASHBOARD)
# =========================
def generate_pdf_report(total_complains, total_ng, total_order, defect_rate, top_facility, top_defect, analysis_text, df_filtered):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    # Đăng ký font hỗ trợ tiếng Việt (DejaVuSans) từ matplotlib
    try:
        font_path = fm.findfont(fm.FontProperties(family='DejaVu Sans'))
        pdfmetrics.registerFont(TTFont('UnicodeFont', font_path))
        font_name = 'UnicodeFont'
    except Exception:
        font_name = 'Helvetica'

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'ReportTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=15,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=4,
        alignment=1 # Center
    } if 'alignment' in dir() else ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontName=font_name, fontSize=15, textColor=colors.HexColor('#1f77b4'), spaceAfter=4)
    
    subtitle_style = ParagraphStyle(
        'ReportSubTitle',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=9,
        textColor=colors.HexColor('#666666'),
        spaceAfter=10,
        alignment=1
    )

    heading_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontName=font_name,
        fontSize=11,
        textColor=colors.HexColor('#222222'),
        spaceBefore=12,
        spaceAfter=6
    )
    body_style = ParagraphStyle(
        'ReportBody',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#444444')
    )

    # Title
    story.append(Paragraph("CUSTOMER COMPLAIN & QA/QC ANALYSIS REPORT", title_style))
    story.append(Paragraph("BÁO CÁO PHÂN TÍCH KHIẾU NẠI KHÁCH HÀNG & QA/QC", subtitle_style))
    story.append(Spacer(1, 4))

    # KPI Table
    kpi_data = [
        ['Total Complains / Tổng KN', str(total_complains), 'Defect Rate / Tỷ lệ lỗi', f"{defect_rate:.2f}%"],
        ['Total NG / Tổng NG', str(int(total_ng)), 'Total Order / Tổng đặt', str(int(total_order))],
        ['Top Facility / NM chính', str(top_facility), 'Top Defect / Lỗi phổ biến', str(top_defect)]
    ]
    
    kpi_table = Table(kpi_data, colWidths=[140, 120, 140, 124])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8f9fa')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#333333')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0'))
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Visual Analytics Dashboard / Phân tích trực quan", heading_style))

    temp_files = []

    try:
        # Style chung cho matplotlib khớp Plotly White template
        plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
        
        # --- CHART 1: Trend of Complains & NG Qty over Time ---
        chart1_path = "temp_chart1.png"
        temp_files.append(chart1_path)
        df_trend = df_filtered.groupby(df_filtered['COMPLAIN_DATE'].dt.date).agg(
            Complains=('WO_NUM', 'count'),
            Total_NG=('NG_QTY', 'sum')
        ).reset_index()

        fig, ax1 = plt.subplots(figsize=(6.5, 2.2), dpi=150)
        
        color = '#1f77b4'
        ax1.set_xlabel('Date / Ngày', fontsize=7)
        ax1.set_ylabel('Complains / Số KN', color=color, fontsize=7)
        line1 = ax1.plot(df_trend['COMPLAIN_DATE'], df_trend['Complains'], color=color, marker='o', linewidth=1.5, label='Complains')
        ax1.tick_params(axis='y', labelcolor=color, labelsize=7)
        ax1.tick_params(axis='x', rotation=20, labelsize=7)

        ax2 = ax1.twinx()  
        color = '#ff7f0e'
        ax2.set_ylabel('NG Quantity / Số lượng lỗi', color=color, fontsize=7)
        line2 = ax2.plot(df_trend['COMPLAIN_DATE'], df_trend['Total_NG'], color=color, marker='s', linewidth=1.5, linestyle='--', label='NG Qty')
        ax2.tick_params(axis='y', labelcolor=color, labelsize=7)
        ax2.grid(False)

        plt.title("Trend of Complains & NG Qty over Time / Xu hướng khiếu nại & số lượng lỗi", fontsize=8, fontweight='bold', fontname='DejaVu Sans' if font_name=='UnicodeFont' else 'sans-serif')
        plt.tight_layout()
        plt.savefig(chart1_path, dpi=150, bbox_inches='tight')
        plt.close()

        story.append(Image(chart1_path, width=490, height=160))
        story.append(Spacer(1, 6))

        # --- CHART 2 & 3: Defect Share (Pie) & Facility Defect (Stacked Bar) ---
        chart2_path = "temp_chart2.png"
        chart3_path = "temp_chart3.png"
        temp_files.extend([chart2_path, chart3_path])

        # Chart 2: Defect Type Share (Pie)
        df_defect_share = df_filtered.groupby('TYPE_OF_DEFECT').size().reset_index(name='Count')
        fig, ax = plt.subplots(figsize=(3.2, 2.2), dpi=150)
        colors_list = plt.cm.Paired(np.linspace(0, 1, len(df_defect_share)))
        ax.pie(df_defect_share['Count'], labels=df_defect_share['TYPE_OF_DEFECT'], colors=colors_list, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 6})
        ax.axis('equal')
        plt.title("Defect Type Share / Tỷ lệ loại lỗi", fontsize=8, fontweight='bold')
        plt.tight_layout()
        plt.savefig(chart2_path, dpi=150, bbox_inches='tight')
        plt.close()

        # Chart 3: Complains by Facility & Defect Type (Stacked Bar)
        df_facility_defect = df_filtered.groupby(['FACILITY', 'TYPE_OF_DEFECT']).size().unstack(fill_value=0)
        fig, ax = plt.subplots(figsize=(3.2, 2.2), dpi=150)
        df_facility_defect.plot(kind='bar', stacked=True, ax=ax, colormap='tab10', width=0.6)
        ax.set_title("Complains by Facility & Defect", fontsize=8, fontweight='bold')
        ax.set_xlabel("Facility / Nhà máy", fontsize=7)
        ax.set_ylabel("Complains", fontsize=7)
        ax.tick_params(axis='x', rotation=15, labelsize=7)
        ax.tick_params(axis='y', labelsize=7)
        ax.legend(fontsize=5, title="Defect", title_fontsize=6, bbox_to_anchor=(1.02, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(chart3_path, dpi=150, bbox_inches='tight')
        plt.close()

        # Đặt 2 biểu đồ cạnh nhau trong một bảng ReportLab
        chart_table_1 = Table([
            [Image(chart2_path, width=235, height=150), Image(chart3_path, width=245, height=150)]
        ], colWidths=[245, 245])
        chart_table_1.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
        ]))
        story.append(chart_table_1)
        story.append(Spacer(1, 6))

        # --- CHART 4: NG Quantity by Facility (Horizontal Bar) ---
        chart4_path = "temp_chart4.png"
        temp_files.append(chart4_path)
        df_facility_ng = df_filtered.groupby('FACILITY')['NG_QTY'].sum().reset_index(name='Total_NG').sort_values(by='Total_NG', ascending=True)
        
        fig, ax = plt.subplots(figsize=(6.5, 2.0), dpi=150)
        ax.barh(df_facility_ng['FACILITY'], df_facility_ng['Total_NG'], color='#2ca02c', height=0.5)
        ax.set_title("NG Quantity by Facility / Tổng lỗi NG theo nhà máy", fontsize=8, fontweight='bold')
        ax.set_xlabel("Total NG Qty", fontsize=7)
        ax.set_ylabel("Facility", fontsize=7)
        ax.tick_params(axis='both', labelsize=7)
        plt.tight_layout()
        plt.savefig(chart4_path, dpi=150, bbox_inches='tight')
        plt.close()

        story.append(Image(chart4_path, width=490, height=140))
        story.append(Spacer(1, 8))

    except Exception as e:
        print(f"Chart generation error: {e}")

    # Analysis Section
    story.append(Paragraph("Root Cause & Recommendations / Nguyên nhân gốc rễ & Giải pháp đề xuất", heading_style))
    story.append(Spacer(1, 4))

    for line in analysis_text.split('\n'):
        if line.strip():
            story.append(Paragraph(line.replace('**', ''), body_style))
            story.append(Spacer(1, 3))

    doc.build(story)
    buffer.seek(0)
    
    for p in temp_files:
        if os.path.exists(p):
            try:
                os.remove(p)
            except:
                pass

    return buffer.getvalue()


# =========================
# MAIN APP
# =========================
if uploaded_files:
    df_raw = load_data(uploaded_files)

    if not df_raw.empty:

        if "view" not in st.session_state:
            st.session_state.view = "ALL"

        st.sidebar.markdown(tr("### View Mode", "### Chế độ xem"))

        if st.sidebar.button("📋 ALL"):
            st.session_state.view = "ALL"

        for f in df_raw['SOURCE_FILE_NAME'].unique():
            if st.sidebar.button(f"📄 {f}"):
                st.session_state.view = f

        df_filtered = df_raw.copy()

        facility_filter = st.sidebar.multiselect(
            tr("Facility", "Nhà máy"),
            df_raw['FACILITY'].unique()
        )

        defect_filter = st.sidebar.multiselect(
            tr("Defect Type", "Loại lỗi"),
            df_raw['TYPE_OF_DEFECT'].unique()
        )

        if facility_filter:
            df_filtered = df_filtered[df_filtered['FACILITY'].isin(facility_filter)]

        if defect_filter:
            df_filtered = df_filtered[df_filtered['TYPE_OF_DEFECT'].isin(defect_filter)]

        if st.session_state.view != "ALL":
            df_filtered = df_filtered[df_filtered['SOURCE_FILE_NAME'] == st.session_state.view]

        # =========================
        # KPI
        # =========================
        total_complains = len(df_filtered)
        total_ng = df_filtered['NG_QTY'].sum()
        total_order = df_filtered['ORDER_QTY'].sum()
        defect_rate = (total_ng / total_order * 100) if total_order > 0 else 0

        top_facility = df_filtered['FACILITY'].mode().iloc[0] if not df_filtered.empty else "N/A"
        top_defect = df_filtered['TYPE_OF_DEFECT'].mode().iloc[0] if not df_filtered.empty else "N/A"

        c1, c2, c3, c4, c5 = st.columns(5)

        c1.metric(tr("Total Complains", "Tổng khiếu nại"), total_complains)
        c2.metric(tr("Defect Rate (%)", "Tỷ lệ lỗi (%)"), round(defect_rate, 2))
        c3.metric(tr("Top Facility", "Nhà máy chính"), top_facility)
        c4.metric(tr("Top Defect", "Lỗi phổ biến"), top_defect)
        c5.metric(tr("Total NG", "Tổng NG"), int(total_ng))

        st.markdown("---")

        # =========================
        # VISUALIZATION
        # =========================
        st.subheader(tr(
            "📊 Visual Analytics Dashboard",
            "📊 Phân tích trực quan"
        ))

        if not df_filtered.empty:

            chart_col1, chart_col2 = st.columns([2, 1])

            with chart_col1:
                st.markdown(tr(
                    "**📈 Trend of Complains & NG Qty over Time**",
                    "**📈 Xu hướng khiếu nại & số lượng lỗi theo thời gian**"
                ))

                df_trend = (
                    df_filtered
                    .groupby(df_filtered['COMPLAIN_DATE'].dt.date)
                    .agg(
                        Complains=('WO_NUM', 'count'),
                        Total_NG=('NG_QTY', 'sum')
                    )
                    .reset_index()
                )

                fig_trend = make_subplots(specs=[[{"secondary_y": True}]])

                fig_trend.add_trace(
                    go.Scatter(
                        x=df_trend['COMPLAIN_DATE'],
                        y=df_trend['Complains'],
                        name="Complains",
                        mode="lines+markers"
                    ),
                    secondary_y=False
                )

                fig_trend.add_trace(
                    go.Scatter(
                        x=df_trend['COMPLAIN_DATE'],
                        y=df_trend['Total_NG'],
                        name="NG Qty",
                        mode="lines+markers"
                    ),
                    secondary_y=True
                )

                fig_trend.update_layout(
                    template="plotly_white",
                    height=350,
                    margin=dict(l=20, r=20, t=20, b=20),
                    legend=dict(orientation="h")
                )

                fig_trend.update_yaxes(title_text="Complains", secondary_y=False)
                fig_trend.update_yaxes(title_text="NG Quantity", secondary_y=True)
                fig_trend.update_xaxes(title_text="Date")

                st.plotly_chart(fig_trend, use_container_width=True)

            with chart_col2:
                st.markdown(tr(
                    "**🍕 Defect Type Share**",
                    "**🍕 Tỷ lệ loại lỗi**"
                ))

                df_defect_share = (
                    df_filtered.groupby('TYPE_OF_DEFECT')
                    .size()
                    .reset_index(name='Count')
                )

                fig_pie = px.pie(
                    df_defect_share,
                    values='Count',
                    names='TYPE_OF_DEFECT',
                    hole=0.4,
                    template="plotly_white"
                )

                fig_pie.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_pie, use_container_width=True)

            chart_col3, chart_col4 = st.columns(2)

            with chart_col3:
                st.markdown(tr(
                    "**🏢 Complains by Facility & Defect Type**",
                    "**🏢 Khiếu nại theo nhà máy & loại lỗi**"
                ))

                df_facility_defect = (
                    df_filtered.groupby(['FACILITY', 'TYPE_OF_DEFECT'])
                    .size()
                    .reset_index(name='Complains')
                )

                fig_bar = px.bar(
                    df_facility_defect,
                    x='FACILITY',
                    y='Complains',
                    color='TYPE_OF_DEFECT',
                    barmode='stack',
                    template="plotly_white"
                )

                fig_bar.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_bar, use_container_width=True)

            with chart_col4:
                st.markdown(tr(
                    "**⚠️ NG Quantity by Facility**",
                    "**⚠️ Tổng lỗi NG theo nhà máy**"
                ))

                df_facility_ng = (
                    df_filtered.groupby('FACILITY')['NG_QTY']
                    .sum()
                    .reset_index(name='Total_NG')
                    .sort_values(by='Total_NG', ascending=True)
                )

                fig_hbar = px.bar(
                    df_facility_ng,
                    x='Total_NG',
                    y='FACILITY',
                    orientation='h',
                    template="plotly_white"
                )

                fig_hbar.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_hbar, use_container_width=True)

        else:
            st.info(tr(
                "No data available for charts",
                "Không có dữ liệu để vẽ biểu đồ"
            ))

        st.markdown("---")

        # =========================
        # SYSTEM ANALYSIS SECTION
        # =========================
        st.subheader(tr(
            "🔍 System Analytics & Insights",
            "🔍 Phân tích & Thông tin chi tiết từ hệ thống"
        ))

        if st.button(tr("Run Analysis", "Chạy phân tích")):

            if not api_key:
                st.error("❌ Chưa có API_KEY")
            else:
                try:
                    from google import genai
                    client = genai.Client(api_key=api_key)

                    sample_df = df_filtered.head(120)
                    sample_data = sample_df.to_markdown(index=False)

                    prompt = f"""
Analyze the provided QA/QC dataset and metrics. Do not include any introductory remarks, greetings, or meta-commentary. Jump straight into the core findings.
For every single point, output the English version first, followed immediately by the Vietnamese translation on the next line.

SUMMARY KPI:
- Complains: {total_complains}
- NG Qty: {total_ng}
- Defect Rate: {defect_rate:.2f}%
- Top Facility: {top_facility}
- Top Defect: {top_defect}

DATA EVIDENCE:
{sample_data}
"""

                    response = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=prompt
                    )

                    st.success(tr(
                        "Analysis completed successfully",
                        "Phân tích hoàn tất thành công"
                    ))

                    st.session_state['response_text'] = response.text

                except Exception as e:
                    st.error(f"Lỗi hệ thống: {e}")

        if 'response_text' in st.session_state:
            st.markdown(st.session_state['response_text'])
            st.markdown("---")
            
            pdf_bytes = generate_pdf_report(
                total_complains, total_ng, total_order, defect_rate, 
                top_facility, top_defect, st.session_state['response_text'], df_filtered
            )
            
            st.download_button(
                label=tr("📥 Download PDF Report", "📥 Tải xuống báo cáo PDF"),
                data=pdf_bytes,
                file_name="Customer_Complain_Analysis_Report.pdf",
                mime="application/pdf"
            )

        st.markdown(tr("### 📋 Data Table", "### 📋 Bảng dữ liệu"))
        st.dataframe(df_filtered, use_container_width=True)
