import streamlit as st
import pandas as pd
import io
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

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
# PDF GENERATION HELPER (INCLUDES PLOTLY CHARTS & ANALYSIS)
# =========================
def generate_pdf_report(total_complains, total_ng, total_order, defect_rate, top_facility, top_defect, analysis_text, df_filtered):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    import os

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'ReportTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=10
    )
    heading_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#333333'),
        spaceBefore=10,
        spaceAfter=4
    )
    body_style = ParagraphStyle(
        'ReportBody',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#444444')
    )

    # Title
    story.append(Paragraph("CUSTOMER COMPLAIN & QA/QC ANALYSIS REPORT", title_style))
    story.append(Paragraph("BÁO CÁO PHÂN TÍCH KHIẾU NẠI KHÁCH HÀNG & QA/QC", body_style))
    story.append(Spacer(1, 8))

    # KPI Table
    kpi_data = [
        ['Total Complains / Tổng KN', str(total_complains), 'Defect Rate / Tỷ lệ lỗi', f"{defect_rate:.2f}%"],
        ['Total NG / Tổng NG', str(int(total_ng)), 'Total Order / Tổng đặt', str(int(total_order))],
        ['Top Facility / NM chính', str(top_facility), 'Top Defect / Lỗi phổ biến', str(top_defect)]
    ]
    
    kpi_table = Table(kpi_data, colWidths=[140, 130, 140, 130])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f5f5f5')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#333333')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dddddd'))
    ]))
    
    story.append(kpi_table)
    story.append(Spacer(1, 10))

    # Generate Chart Image for PDF
    chart_path = "temp_chart.png"
    try:
        df_trend = df_filtered.groupby(df_filtered['COMPLAIN_DATE'].dt.date).agg(Complains=('WO_NUM', 'count')).reset_index()
        fig = px.bar(df_trend, x='COMPLAIN_DATE', y='Complains', title="Complains Trend / Xu hướng khiếu nại")
        fig.update_layout(height=200, width=500, margin=dict(l=20, r=20, t=30, b=20))
        fig.write_image(chart_path)
        
        story.append(Paragraph("Analytics Chart / Biểu đồ phân tích", heading_style))
        story.append(Image(chart_path, width=450, height=180))
        story.append(Spacer(1, 10))
    except Exception:
        pass # Fallback nếu môi trường thiếu engine ảnh tĩnh của plotly

    # Analysis Section
    story.append(Paragraph("Root Cause & Recommendations / Nguyên nhân gốc rễ & Giải pháp đề xuất", heading_style))
    story.append(Spacer(1, 4))

    for line in analysis_text.split('\n'):
        if line.strip():
            story.append(Paragraph(line.replace('**', ''), body_style))
            story.append(Spacer(1, 3))

    doc.build(story)
    buffer.seek(0)
    
    if os.path.exists(chart_path):
        try:
            os.remove(chart_path)
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

                    # Prompt không chứa lời dẫn mở đầu, yêu cầu trả cấu trúc song ngữ trực tiếp từng dòng/đoạn
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
                        model="gemini-3.1pro",
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
