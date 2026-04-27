"""
Enhanced Board Dashboard - Professional KPIs and Metrics for Executive View
Complete SPS Excel Pipeline Application with Full Dashboard Download
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import sys
import time
import requests
import json
from datetime import datetime
import traceback
import io
import tempfile
import os
import base64
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

sys.path.insert(0, str(Path(__file__).parent))

from orchestrator import ExcelOrchestrator
from complete_report import CompleteReportGenerator

# Import configuration
from config import (
    BASE_DIR,
    RAW_DIR,
    PROCESSED_DIR,
    REPORTS_DIR,
    PAGE_TITLE,
    PAGE_ICON,
    PAGE_LAYOUT,
    SIDEBAR_STATE,
    PDF_ORIENTATION,
    PDF_FORMAT,
    PDF_UNIT,
    PDF_MARGIN,
    MAX_UPLOAD_SIZE_MB,
    ALLOWED_EXTENSIONS,
    LOGO_PATH,
    FALLBACK_LOGO_URL,
    HEADER_COLOR,
    CHART_HEIGHT,
    CHART_WIDTH,
    CHART_TEMPLATE,
    VALUE_COLUMN_KEYWORDS,
    STATUS_COLUMN_KEYWORDS,
    DATE_COLUMN_KEYWORDS,
    SUCCESS_KEYWORDS,
    FAILURE_KEYWORDS,
    PENDING_KEYWORDS,
    HIGH_LEVEL_METRICS_PATTERNS,
    KPI_COLORS,
    ensure_directories
)

# Ensure directories exist
ensure_directories()

# ============================================
# PAGE CONFIGURATION (Must be first Streamlit command)
# ============================================
st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon=PAGE_ICON,
    layout=PAGE_LAYOUT,
    initial_sidebar_state=SIDEBAR_STATE
)

# ============================================
# AUTHENTICATION SETUP
# ============================================

# Load configuration
config_path = Path(__file__).parent.parent / 'config.yaml'
with open(config_path, 'r', encoding='utf-8') as file:
    config = yaml.load(file, Loader=SafeLoader)

# Create authenticator object
authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config.get('pre-authorized', [])
)

# Create login widget
try:
    name, authentication_status, username = authenticator.login(
        location='main',
        fields={
            'Form name': 'SPS Dashboard Login',
            'Username': 'Username',
            'Password': 'Password',
            'Login': 'Login'
        }
    )
except Exception as e:
    st.error(f"Login form error: {e}")
    authentication_status = None
    name = None
    username = None

# ============================================
# SESSION STATE MANAGEMENT
# ============================================

# Initialize session state
if 'authentication_status' not in st.session_state:
    st.session_state['authentication_status'] = authentication_status
if 'name' not in st.session_state:
    st.session_state['name'] = name
if 'username' not in st.session_state:
    st.session_state['username'] = username

# Update session state from login attempt
if authentication_status is not None:
    st.session_state['authentication_status'] = authentication_status
    st.session_state['name'] = name
    st.session_state['username'] = username

# ============================================
# LOGOUT HANDLER
# ============================================

def logout():
    for key in ['authentication_status', 'name', 'username']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ============================================
# HELPER FUNCTIONS
# ============================================

def detect_value_column(df: pd.DataFrame) -> str:
    """Detect value/amount column using config keywords"""
    numeric_cols = df.select_dtypes(include=['number']).columns
    for col in numeric_cols:
        col_lower = col.lower()
        for keyword in VALUE_COLUMN_KEYWORDS:
            if keyword in col_lower:
                return col
    return numeric_cols[0] if len(numeric_cols) > 0 else None


def detect_status_metrics(df: pd.DataFrame) -> tuple:
    """Detect success/failure counts using config keywords"""
    success_count = 0
    failed_count = 0
    
    status_col = None
    for col in df.columns:
        col_lower = col.lower()
        for keyword in STATUS_COLUMN_KEYWORDS:
            if keyword in col_lower:
                status_col = col
                break
        if status_col:
            break
    
    if status_col:
        for status, count in df[status_col].value_counts().items():
            status_str = str(status).lower()
            if any(kw in status_str for kw in SUCCESS_KEYWORDS):
                success_count += count
            elif any(kw in status_str for kw in FAILURE_KEYWORDS):
                failed_count += count
            elif any(kw in status_str for kw in PENDING_KEYWORDS):
                pass
            else:
                success_count += count
    
    return success_count, failed_count


def is_high_level_metrics_sheet(sheet_name: str) -> bool:
    """Check if sheet is a high level metrics sheet using config patterns"""
    sheet_lower = sheet_name.lower()
    for pattern in HIGH_LEVEL_METRICS_PATTERNS:
        if pattern in sheet_lower:
            return True
    return False


def get_kpi_color(kpi_name: str) -> tuple:
    """Get color for KPI based on name using config"""
    kpi_lower = kpi_name.lower()
    
    if any(word in kpi_lower for word in ['volume', 'count', 'transaction']):
        return KPI_COLORS['volume']
    elif any(word in kpi_lower for word in ['value', 'amount', 'revenue']):
        return KPI_COLORS['value']
    elif any(word in kpi_lower for word in ['failed', 'failure', 'error']):
        return KPI_COLORS['failed']
    elif any(word in kpi_lower for word in ['rate', 'percentage', '%']):
        return KPI_COLORS['rate']
    else:
        return KPI_COLORS['default']


# ============================================
# FULL DASHBOARD DOWNLOAD FUNCTION
# ============================================

def download_full_dashboard(df, sheet_name, file_name):
    """Download the complete dashboard page with charts"""
    
    try:
        from fpdf import FPDF
        import tempfile
        import uuid
        
        st.info(" Preparing dashboard for download...")
        
        # Create PDF object using config
        pdf = FPDF(orientation=PDF_ORIENTATION, unit=PDF_UNIT, format=PDF_FORMAT)
        pdf.set_auto_page_break(auto=True, margin=PDF_MARGIN)
        
        temp_files = []
        
        def clean_text(text):
            """Remove or replace Unicode characters that cause issues"""
            replacements = {
                '\u2022': '-', '\u25CF': '-', '\u25CB': 'o', '\u2013': '-',
                '\u2014': '--', '\u2018': "'", '\u2019': "'", '\u201C': '"',
                '\u201D': '"', '\u00A0': ' ', '\u20B9': 'Rs.', '\u20AC': 'EUR',
                '\u00A3': 'GBP', '\u00A5': 'Yen', '\u2713': '✓', '\u2714': '✔', '\u274C': '✘'
            }
            for old, new in replacements.items():
                text = text.replace(old, new)
            return text.encode('ascii', errors='ignore').decode('ascii')
        
        def save_chart_to_pdf(fig, title, pdf_obj):
            """Save a chart to a new page in the PDF"""
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_{uuid.uuid4().hex}.png")
            try:
                fig.write_image(temp_path, scale=1.5, width=700, height=350)
                pdf_obj.add_page()
                pdf_obj.set_font("Arial", "B", 14)
                pdf_obj.set_fill_color(HEADER_COLOR[0], HEADER_COLOR[1], HEADER_COLOR[2])
                pdf_obj.set_text_color(255, 255, 255)
                pdf_obj.cell(0, 10, clean_text(title), ln=True, fill=True)
                pdf_obj.set_text_color(0, 0, 0)
                pdf_obj.ln(8)
                pdf_obj.image(temp_path, x=15, y=35, w=270)
                return True
            except Exception as e:
                st.warning(f"Could not generate chart: {title} - {str(e)}")
                return False
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        # ========== TITLE PAGE ==========
        pdf.add_page()
        pdf.set_fill_color(HEADER_COLOR[0], HEADER_COLOR[1], HEADER_COLOR[2])
        pdf.rect(0, 0, 297, 45, 'F')
        
        pdf.set_y(15)
        pdf.set_font("Arial", "B", 24)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, PAGE_TITLE, ln=True, align="C")
        
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 8, clean_text(sheet_name), ln=True, align="C")
        pdf.set_font("Arial", "I", 10)
        pdf.cell(0, 6, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
        pdf.cell(0, 6, f"Source File: {clean_text(file_name)}", ln=True, align="C")
        
        # ========== KPI SECTION ==========
        pdf.ln(8)
        pdf.set_font("Arial", "B", 14)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(0, 10, "Key Performance Indicators", ln=True)
        pdf.ln(5)
        
        numeric_cols = df.select_dtypes(include=['number']).columns
        
        if len(numeric_cols) > 0:
            primary_metric = detect_value_column(df) or numeric_cols[0]
            total_value = df[primary_metric].sum()
            avg_value = df[primary_metric].mean()
            max_value = df[primary_metric].max()
            min_value = df[primary_metric].min()
            completeness = (1 - df.isnull().sum().sum() / (df.shape[0] * df.shape[1])) * 100 if df.shape[0] > 0 else 0
            
            kpis = [
                {'label': f'Total {primary_metric}', 'value': f'{total_value:,.0f}', 'color': KPI_COLORS['value']},
                {'label': f'Average {primary_metric}', 'value': f'{avg_value:,.0f}', 'color': KPI_COLORS['default']},
                {'label': f'Maximum {primary_metric}', 'value': f'{max_value:,.0f}', 'color': KPI_COLORS['volume']},
                {'label': f'Minimum {primary_metric}', 'value': f'{min_value:,.0f}', 'color': KPI_COLORS['failed']},
                {'label': 'Total Records', 'value': f'{len(df):,}', 'color': KPI_COLORS['rate']},
                {'label': 'Data Quality', 'value': f'{completeness:.1f}%', 'color': KPI_COLORS['default']}
            ]
            
            card_width = 85
            card_height = 45
            start_x = 15
            start_y = pdf.get_y()
            
            for i, kpi in enumerate(kpis):
                row = i // 3
                col = i % 3
                x = start_x + col * (card_width + 8)
                y = start_y + row * (card_height + 8)
                
                pdf.set_fill_color(248, 249, 250)
                pdf.rect(x, y, card_width, card_height, 'F')
                pdf.set_draw_color(kpi['color'][0], kpi['color'][1], kpi['color'][2])
                pdf.rect(x, y, card_width, card_height, 'D')
                
                pdf.set_xy(x + 5, y + 5)
                pdf.set_font("Arial", "B", 9)
                pdf.set_text_color(100, 100, 100)
                pdf.cell(card_width - 10, 6, clean_text(kpi['label']), ln=True, align='C')
                
                pdf.set_xy(x + 5, y + 15)
                pdf.set_font("Arial", "B", 16)
                pdf.set_text_color(kpi['color'][0], kpi['color'][1], kpi['color'][2])
                pdf.cell(card_width - 10, 10, clean_text(kpi['value']), ln=True, align='C')
            
            pdf.set_y(start_y + 2 * (card_height + 8) + 10)
        
        # ========== GENERATE CHARTS ==========
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
        date_cols = [c for c in df.columns if any(kw in c.lower() for kw in DATE_COLUMN_KEYWORDS)]
        
        charts_added = 0
        
        # Chart 1: Bar Chart
        if categorical_cols and numeric_cols:
            try:
                x_axis = categorical_cols[0]
                y_axis = numeric_cols[0]
                agg_data = df.groupby(x_axis)[y_axis].sum().reset_index().sort_values(y_axis, ascending=False).head(10)
                fig = px.bar(agg_data, x=x_axis, y=y_axis, title=f"Bar Chart: {y_axis} by {x_axis}",
                            color=y_axis, color_continuous_scale='Viridis')
                fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
                if save_chart_to_pdf(fig, f"{sheet_name} - Bar Chart", pdf):
                    charts_added += 1
            except Exception as e:
                st.warning(f"Could not generate bar chart: {e}")
        
        # Chart 2: Pie Chart
        if categorical_cols and numeric_cols:
            try:
                pie_col = categorical_cols[0]
                value_col = numeric_cols[0]
                pie_data = df.groupby(pie_col)[value_col].sum().reset_index().sort_values(value_col, ascending=False).head(10)
                fig = px.pie(pie_data, values=value_col, names=pie_col, title=f"Pie Chart: Distribution of {value_col}",
                            hole=0.3, color_discrete_sequence=px.colors.qualitative.Set3)
                fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
                if save_chart_to_pdf(fig, f"{sheet_name} - Pie Chart", pdf):
                    charts_added += 1
            except Exception as e:
                st.warning(f"Could not generate pie chart: {e}")
        
        # Chart 3: Trend Chart
        if date_cols and numeric_cols:
            try:
                date_col = date_cols[0]
                value_col = numeric_cols[0]
                df_temp = df.copy()
                df_temp[date_col] = pd.to_datetime(df_temp[date_col], errors='coerce')
                df_temp = df_temp.dropna(subset=[date_col])
                if len(df_temp) > 0:
                    df_temp['period'] = df_temp[date_col].dt.to_period('M')
                    trend_data = df_temp.groupby('period')[value_col].sum().reset_index()
                    trend_data['period_str'] = trend_data['period'].astype(str)
                    fig = px.line(trend_data, x='period_str', y=value_col, title=f"Trend Chart: {value_col} Over Time",
                                markers=True, color_discrete_sequence=['#3498db'])
                    fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
                    if save_chart_to_pdf(fig, f"{sheet_name} - Trend Chart", pdf):
                        charts_added += 1
            except Exception as e:
                st.warning(f"Could not generate trend chart: {e}")
        
        # Chart 4: Histogram
        if numeric_cols:
            try:
                hist_col = numeric_cols[0]
                fig = px.histogram(df, x=hist_col, nbins=20, title=f"Histogram: Distribution of {hist_col}",
                                color_discrete_sequence=['#667eea'])
                fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
                if save_chart_to_pdf(fig, f"{sheet_name} - Histogram", pdf):
                    charts_added += 1
            except Exception as e:
                st.warning(f"Could not generate histogram: {e}")
        
        # Chart 5: Scatter Plot
        if len(numeric_cols) >= 2:
            try:
                x_col = numeric_cols[0]
                y_col = numeric_cols[1]
                fig = px.scatter(df, x=x_col, y=y_col, title=f"Scatter Plot: {y_col} vs {x_col}",
                                color_discrete_sequence=['#667eea'])
                fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
                if save_chart_to_pdf(fig, f"{sheet_name} - Scatter Plot", pdf):
                    charts_added += 1
            except Exception as e:
                st.warning(f"Could not generate scatter plot: {e}")
        
        # Chart 6: Box Plot
        # if categorical_cols and numeric_cols:
        #     try:
        #         box_x = categorical_cols[0]
        #         box_y = numeric_cols[0]
        #         top_cats = df[box_x].value_counts().head(5).index.tolist()
        #         df_box = df[df[box_x].isin(top_cats)]
        #         fig = px.box(df_box, x=box_x, y=box_y, title=f"Box Plot: Distribution of {box_y} by {box_x}",
        #                     color=box_x)
        #         fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE, showlegend=False)
        #         if save_chart_to_pdf(fig, f"{sheet_name} - Box Plot", pdf):
        #             charts_added += 1
        #     except Exception as e:
        #         st.warning(f"Could not generate box plot: {e}")
        
        # Generate PDF
        pdf_bytes = pdf.output(dest='S').encode('latin-1')
        
        if charts_added == 0:
            st.warning("No charts could be generated for this dashboard. Only KPIs are included.")
        
        st.success(f" Dashboard ready! Contains {len(kpis)} KPIs and {charts_added} charts")
        
        st.download_button(
            label=" Download Full Dashboard (PDF)",
            data=pdf_bytes,
            file_name=f"{sheet_name}_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
        return True
        
    except Exception as e:
        st.error(f"Error generating dashboard: {str(e)}")
        st.code(traceback.format_exc())
        return False


# ============================================
# BOARD-LEVEL KPI FUNCTIONS
# ============================================

def display_board_kpis(df, sheet_name):
    """Display professional KPI cards for board dashboard"""
    
    st.subheader(" Key Performance Indicators")
    
    numeric_cols = df.select_dtypes(include=['number']).columns
    
    if len(numeric_cols) == 0:
        st.warning("No numeric data available for KPIs")
        return
    
    primary_metric = detect_value_column(df) or numeric_cols[0]
    
    total_value = df[primary_metric].sum()
    avg_value = df[primary_metric].mean()
    max_value = df[primary_metric].max()
    min_value = df[primary_metric].min()
    
    date_cols = [c for c in df.columns if any(kw in c.lower() for kw in DATE_COLUMN_KEYWORDS)]
    trend_pct = 0
    if date_cols:
        try:
            df_temp = df.copy()
            df_temp['_temp_date'] = pd.to_datetime(df_temp[date_cols[0]], errors='coerce')
            df_temp['_month'] = df_temp['_temp_date'].dt.to_period('M')
            monthly_trend = df_temp.groupby('_month')[primary_metric].sum()
            if len(monthly_trend) >= 2:
                trend_pct = ((monthly_trend.iloc[-1] - monthly_trend.iloc[-2]) / monthly_trend.iloc[-2]) * 100
        except:
            pass
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label=f" Total {primary_metric.replace('_', ' ').title()}",
            value=f"{total_value:,.0f}",
            delta=f"{trend_pct:+.1f}% vs previous" if trend_pct != 0 else None,
            delta_color="normal"
        )
    
    with col2:
        st.metric(
            label=f" Average {primary_metric.replace('_', ' ').title()}",
            value=f"{avg_value:,.0f}",
            help="Average value per transaction/record"
        )
    
    with col3:
        st.metric(
            label=f" Maximum {primary_metric.replace('_', ' ').title()}",
            value=f"{max_value:,.0f}",
            help="Highest value recorded"
        )
    
    with col4:
        st.metric(
            label=f" Minimum {primary_metric.replace('_', ' ').title()}",
            value=f"{min_value:,.0f}",
            help="Lowest value recorded"
        )
    
    st.divider()


def display_visualizations(df, sheet_name):
    """Display quick visualizations"""
    
    if df.empty:
        st.info("No data available for visualizations")
        return
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
    
    if not numeric_cols:
        st.info("No numeric columns available for visualizations")
        return
    
    if categorical_cols:
        st.subheader("Bar Chart")
        col1, col2 = st.columns(2)
        with col1:
            x_axis = st.selectbox("X-Axis (Category)", categorical_cols, key="bar_x")
        with col2:
            y_axis = st.selectbox("Y-Axis (Value)", numeric_cols, key="bar_y")
        
        agg_data = df.groupby(x_axis)[y_axis].sum().reset_index().sort_values(y_axis, ascending=False)
        fig = px.bar(agg_data, x=x_axis, y=y_axis, title=f"{y_axis} by {x_axis}",
                    color=y_axis, color_continuous_scale='Viridis')
        fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
    
    date_cols = [c for c in df.columns if any(kw in c.lower() for kw in DATE_COLUMN_KEYWORDS)]
    if date_cols and numeric_cols:
        st.subheader(" Trend Chart")
        col1, col2 = st.columns(2)
        with col1:
            date_col = st.selectbox("Date Column", date_cols, key="line_date")
            try:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            except:
                pass
        with col2:
            value_col = st.selectbox("Value", numeric_cols, key="line_value")
        
        df_sorted = df.sort_values(date_col, ascending=False)
        fig = px.line(df_sorted, x=date_col, y=value_col, title=f"{value_col} Over Time", 
                    markers=True, color_discrete_sequence=['#667eea'])
        fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)


def display_charts(df, sheet_name):
    """Display interactive charts"""
    
    if df.empty:
        st.info("No data available for charts")
        return
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
    
    chart_type = st.selectbox(
        "Select Chart Type",
        [" Pie Chart", " Scatter Plot", " Histogram"],
        key="chart_type_select"
    )
    
    if chart_type == " Pie Chart" and categorical_cols:
        pie_col = st.selectbox("Category", categorical_cols, key="pie_col")
        value_col = st.selectbox("Value Column", numeric_cols, key="pie_value")
        pie_data = df.groupby(pie_col)[value_col].sum().reset_index().sort_values(value_col, ascending=False).head(12)
        pie_data.columns = [pie_col, value_col]
        fig = px.pie(pie_data, values=value_col, names=pie_col, title=f"Distribution of {value_col} by {pie_col}",
                    color_discrete_sequence=px.colors.qualitative.Set3)
        fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
        
    elif chart_type == " Scatter Plot" and len(numeric_cols) >= 2:
        col1, col2 = st.columns(2)
        with col1:
            x_col = st.selectbox("X-Axis", numeric_cols, key="scatter_x")
        with col2:
            y_col = st.selectbox("Y-Axis", numeric_cols, key="scatter_y")
        
        color_col = None
        if categorical_cols:
            color_col = st.selectbox("Color By", ['None'] + categorical_cols, key="scatter_color")
            if color_col == 'None':
                color_col = None
        
        if color_col:
            fig = px.scatter(df, x=x_col, y=y_col, color=color_col, title=f"{y_col} vs {x_col}")
        else:
            fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
        
    elif chart_type == " Histogram" and numeric_cols:
        hist_col = st.selectbox("Column", numeric_cols, key="hist_col")
        bins = st.slider("Number of Bins", 5, 50, 20)
        fig = px.histogram(df, x=hist_col, nbins=bins, title=f"Distribution of {hist_col}",
                        color_discrete_sequence=['#667eea'])
        fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
        
    # elif chart_type == " Box Plot" and numeric_cols and categorical_cols:
    #     col1, col2 = st.columns(2)
    #     with col1:
    #         y_col = st.selectbox("Value", numeric_cols, key="box_y")
    #     with col2:
    #         x_col = st.selectbox("Group By", categorical_cols, key="box_x")
    #     fig = px.box(df, x=x_col, y=y_col, title=f"Distribution of {y_col} by {x_col}",
    #                 color_discrete_sequence=['#3498db'])
    #     fig.update_layout(height=CHART_HEIGHT, width=CHART_WIDTH, template=CHART_TEMPLATE)
    #     st.plotly_chart(fig, use_container_width=True)
    else:
        st.info(f"Need more columns for {chart_type}. Try a different chart type.")


def display_statistics(df):
    """Display summary statistics"""
    
    if df.empty:
        st.info("No data available for statistics")
        return
    
    st.subheader(" Summary Statistics")
    
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        stats_df = df[numeric_cols].describe().round(2)
        st.dataframe(stats_df, use_container_width=True)
    else:
        st.info("No numeric columns for statistics")
    
    st.subheader(" Missing Values")
    
    missing_data = []
    for col in df.columns:
        missing_count = df[col].isnull().sum()
        missing_pct = (missing_count / len(df) * 100) if len(df) > 0 else 0
        if missing_count > 0:
            missing_data.append({'Column': col, 'Missing Count': missing_count, 'Missing %': round(missing_pct, 2)})
    
    if missing_data:
        missing_df = pd.DataFrame(missing_data)
        st.dataframe(missing_df, use_container_width=True)
    else:
        st.success(" No missing values found!")
    
    st.subheader(" Column Information")
    
    col_data = []
    for col in df.columns:
        try:
            sample_values = df[col].dropna().head(3).tolist()
            sample_str = ', '.join([str(v)[:30] for v in sample_values]) if sample_values else 'No data'
        except:
            sample_str = 'Error reading sample'
        
        col_data.append({'Column': col, 'Data Type': str(df[col].dtype), 'Unique Values': df[col].nunique(), 'Sample Values': sample_str})
    
    if col_data:
        col_df = pd.DataFrame(col_data)
        st.dataframe(col_df, use_container_width=True)


def display_time_series_analysis(df, sheet_name):
    """Display time series analysis and trends"""
    
    date_cols = [c for c in df.columns if any(kw in c.lower() for kw in DATE_COLUMN_KEYWORDS)]
    
    if not date_cols:
        return
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    
    if not numeric_cols:
        return
    
    st.subheader(" Trend Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        date_col = st.selectbox("Select Date Column", date_cols, key="ts_date")
        metric_col = st.selectbox("Select Metric to Track", numeric_cols, key="ts_metric")
    
    with col2:
        period = st.selectbox("Aggregation Period", ["Daily", "Weekly", "Monthly", "Quarterly"], key="ts_period")
    
    df_ts = df.copy()
    df_ts[date_col] = pd.to_datetime(df_ts[date_col], errors='coerce')
    df_ts = df_ts.dropna(subset=[date_col])
    
    if len(df_ts) == 0:
        st.info("No valid dates found for time series analysis")
        return
    
    if period == "Daily":
        df_ts['period'] = df_ts[date_col].dt.date
    elif period == "Weekly":
        df_ts['period'] = df_ts[date_col].dt.to_period('W')
    elif period == "Monthly":
        df_ts['period'] = df_ts[date_col].dt.to_period('M')
    else:
        df_ts['period'] = df_ts[date_col].dt.to_period('Q')
    
    trend_data = df_ts.groupby('period')[metric_col].sum().reset_index()
    trend_data = trend_data.sort_values('period', ascending=False)
    trend_data['period_str'] = trend_data['period'].astype(str)
    
    if len(trend_data) < 2:
        st.info("Insufficient data points for trend analysis")
        return
    
    fig = px.line(
        trend_data,
        x='period_str',
        y=metric_col,
        title=f"{metric_col} Trend ({period})",
        markers=True,
        color_discrete_sequence=['#667eea']
    )
    
    window = min(3, len(trend_data))
    fig.add_trace(
        go.Scatter(
            x=trend_data['period_str'],
            y=trend_data[metric_col].rolling(window=window, min_periods=1).mean(),
            mode='lines',
            name='Moving Average',
            line=dict(dash='dash', color='#e74c3c')
        )
    )
    
    fig.update_layout(height=CHART_HEIGHT + 150, xaxis_title="Time Period", yaxis_title=metric_col, template=CHART_TEMPLATE)
    st.plotly_chart(fig, use_container_width=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_growth = ((trend_data[metric_col].iloc[-1] - trend_data[metric_col].iloc[0]) / trend_data[metric_col].iloc[0]) * 100
        st.metric(label="Total Growth", value=f"{total_growth:+.1f}%", delta_color="normal")
    
    with col2:
        if len(trend_data) >= 3:
            periods = len(trend_data) - 1
            cagr = (pow(trend_data[metric_col].iloc[-1] / trend_data[metric_col].iloc[0], 1/periods) - 1) * 100
            st.metric(label=f"CAGR ({period})", value=f"{cagr:+.1f}%", help="Compound Annual Growth Rate")
    
    with col3:
        best_period = trend_data.loc[trend_data[metric_col].idxmax(), 'period_str']
        worst_period = trend_data.loc[trend_data[metric_col].idxmin(), 'period_str']
        st.metric(label="Best/Worst Period", value=f"{best_period} / {worst_period}")
    
    st.divider()


def display_comparative_analysis(df, sheet_name):
    """Display comparative analysis charts"""
    
    st.subheader(" Comparative Analysis")
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
    
    if not numeric_cols or not categorical_cols:
        st.info("Need both numeric and categorical columns for comparative analysis")
        return
    
    col1, col2 = st.columns(2)
    
    with col1:
        cat_col = st.selectbox("Select Category to Compare", categorical_cols, key="compare_cat")
        metric_col = st.selectbox("Select Metric", numeric_cols, key="compare_metric")
        top_n = st.slider("Show Top N Categories", 5, 20, 10)
        
        agg_data = df.groupby(cat_col)[metric_col].sum().sort_values(ascending=False).head(top_n)
        
        fig = px.bar(
            x=agg_data.values,
            y=agg_data.index,
            orientation='h',
            title=f"Top {top_n} {cat_col} by {metric_col}",
            labels={'x': metric_col, 'y': cat_col},
            color=agg_data.values,
            color_continuous_scale='Viridis'
        )
        fig.update_layout(height=CHART_HEIGHT + 100, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        if len(agg_data) <= 10:
            fig = px.pie(
                values=agg_data.values,
                names=agg_data.index,
                title=f"Distribution of {metric_col} by {cat_col}",
                hole=0.3,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=CHART_HEIGHT + 100, template=CHART_TEMPLATE)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.dataframe(agg_data.reset_index().rename(columns={cat_col: cat_col, metric_col: metric_col}), use_container_width=True)
    
    st.divider()


def display_target_vs_actual(df, sheet_name):
    """Display target vs actual comparison if target columns exist"""
    
    numeric_cols = df.select_dtypes(include=['number']).columns
    
    target_cols = [c for c in numeric_cols if 'target' in c.lower() or 'goal' in c.lower() or 'budget' in c.lower()]
    actual_cols = [c for c in numeric_cols if 'actual' in c.lower() or 'sales' in c.lower() or 'revenue' in c.lower()]
    
    if not target_cols or not actual_cols:
        return
    
    st.subheader(" Target vs Actual Performance")
    
    target_col = target_cols[0]
    actual_col = actual_cols[0]
    
    total_target = df[target_col].sum()
    total_actual = df[actual_col].sum()
    achievement = (total_actual / total_target * 100) if total_target > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(label=f"Target ({target_col})", value=f"{total_target:,.0f}")
    
    with col2:
        st.metric(
            label=f"Actual ({actual_col})",
            value=f"{total_actual:,.0f}",
            delta=f"{achievement - 100:+.1f}% vs target",
            delta_color="inverse" if achievement < 100 else "normal"
        )
    
    with col3:
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=achievement,
            title={'text': "Achievement %"},
            delta={'reference': 100},
            gauge={
                'axis': {'range': [None, 150]},
                'bar': {'color': "#27ae60"},
                'steps': [
                    {'range': [0, 50], 'color': "#e74c3c"},
                    {'range': [50, 80], 'color': "#f39c12"},
                    {'range': [80, 100], 'color': "#f1c40f"},
                    {'range': [100, 150], 'color': "#2ecc71"}
                ],
                'threshold': {'line': {'color': "#c0392b", 'width': 4}, 'thickness': 0.75, 'value': 100}
            }
        ))
        fig.update_layout(height=250, template=CHART_TEMPLATE)
        st.plotly_chart(fig, use_container_width=True)
    
    st.divider()


# ============================================
# FILE MANAGEMENT FUNCTIONS
# ============================================

def display_file_and_sheet_selector(file_structure):
    """Display file and sheet selector in sidebar"""
    
    st.sidebar.markdown("###  Reports")
    
    if st.sidebar.button(" Generate Complete Transaction Report", use_container_width=True,
                        help="Generates a comprehensive report with all transaction metrics from ALL sheets"):
        generate_complete_transaction_report()
    
    st.sidebar.divider()
    st.sidebar.markdown("###  Select File")
    
    file_names = list(file_structure.keys())
    
    current_file = st.session_state.get('selected_file', file_names[0]) if file_names else None
    
    if current_file and current_file in file_names:
        default_index = file_names.index(current_file)
    else:
        default_index = 0
    
    selected_file = st.sidebar.selectbox(
        "Choose a file", file_names, key="file_selector", index=default_index if file_names else 0
    )
    
    st.session_state['selected_file'] = selected_file
    file_info = file_structure[selected_file]
    sheets = file_info['sheets']
    
    st.sidebar.markdown("###  Select Sheet")
    
    current_sheet = st.session_state.get('selected_sheet', sheets[0]) if sheets else None
    
    if current_sheet and current_sheet in sheets:
        sheet_index = sheets.index(current_sheet)
    else:
        sheet_index = 0
    
    selected_sheet = st.sidebar.selectbox(
        "Choose a sheet", sheets, key="sheet_selector", index=sheet_index if sheets else 0
    )
    
    st.session_state['selected_sheet'] = selected_sheet
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("###  File Info")
    st.sidebar.markdown(f"**Type:** {file_info['type'].replace('_', ' ').title()}")
    st.sidebar.markdown(f"**Sheets:** {len(sheets)}")
    
    with st.sidebar.expander(" All Sheets in this File"):
        for sheet in sheets:
            if sheet == selected_sheet:
                st.markdown(f" **{sheet}** (current)")
            else:
                st.markdown(f" {sheet}")
    
    return selected_file, selected_sheet, file_info


def display_sheet_data_enhanced(file_info, selected_sheet):
    """Enhanced version with board-level KPIs and full dashboard download"""
    
    if selected_sheet not in file_info['data']:
        st.error(f"Sheet '{selected_sheet}' not found in data")
        return
    
    df = file_info['data'][selected_sheet]
    
    if df is None or df.empty:
        st.warning(f"No data found for sheet: {selected_sheet}")
        return
    
    # Header with download button
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.header(f" {selected_sheet}")
        st.caption(f"From file: {st.session_state.get('selected_file', 'Unknown')}")
    
    with col2:
        if st.button(" Download Full Dashboard", type="primary", use_container_width=True,
                    help="Download complete dashboard with KPIs and charts as PDF"):
            download_full_dashboard(df, selected_sheet, st.session_state.get('selected_file', 'Unknown'))
    
    # Board KPIs Section
    display_board_kpis(df, selected_sheet)
    
    # Quick metrics row
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Records", f"{len(df):,}")
    with col2:
        numeric_cols = df.select_dtypes(include=['number']).columns
        st.metric("Numeric Columns", len(numeric_cols))
    with col3:
        st.metric("Total Columns", df.shape[1])
    
    st.divider()
    
    # Time Series Analysis (if date columns exist)
    date_cols = [c for c in df.columns if any(kw in c.lower() for kw in DATE_COLUMN_KEYWORDS)]
    if date_cols:
        display_time_series_analysis(df, selected_sheet)
    
    # Target vs Actual (if target columns exist)
    numeric_cols = df.select_dtypes(include=['number']).columns
    target_cols = [c for c in numeric_cols if 'target' in c.lower() or 'goal' in c.lower()]
    if target_cols:
        display_target_vs_actual(df, selected_sheet)
    
    # Comparative Analysis
    display_comparative_analysis(df, selected_sheet)
    
    # Filters
    with st.expander(" Advanced Filters", expanded=False):
        filtered_df = df.copy()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
        
        if categorical_cols:
            cols = st.columns(min(3, len(categorical_cols)))
            for i, col in enumerate(categorical_cols[:3]):
                with cols[i]:
                    unique_vals = ['All'] + sorted(df[col].dropna().unique().tolist())
                    selected = st.selectbox(f"Filter by {col}", unique_vals, key=f"filter_{col}_{selected_sheet}")
                    if selected != 'All':
                        filtered_df = filtered_df[filtered_df[col] == selected]
    
    # Tabs for visualizations
    tab1, tab2, tab3, tab4 = st.tabs([" Visualizations", " Charts", " Data Table", " Statistics"])
    
    with tab1:
        display_visualizations(filtered_df, selected_sheet)
    
    with tab2:
        display_charts(filtered_df, selected_sheet)
    
    with tab3:
        st.dataframe(filtered_df, use_container_width=True, height=400)
    
    with tab4:
        display_statistics(filtered_df)


def upload_excel_files():
    """Handle file uploads"""
    st.subheader(" Upload Excel Files")
    
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=list(ALLOWED_EXTENSIONS),
        accept_multiple_files=True,
        help=f"Upload single or multiple Excel files. Max size: {MAX_UPLOAD_SIZE_MB}MB per file"
    )
    
    if uploaded_files:
        st.info(f" {len(uploaded_files)} file(s) selected")
        for file in uploaded_files:
            st.write(f"   - {file.name}")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button(" Upload & Process", type="primary", use_container_width=True):
                for uploaded_file in uploaded_files:
                    file_path = RAW_DIR / uploaded_file.name
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    st.success(f" Uploaded: {uploaded_file.name}")
                
                run_pipeline()
                time.sleep(2)
                st.rerun()
        
        with col2:
            if st.button(" Run Pipeline Only", use_container_width=True):
                run_pipeline()
                time.sleep(2)
                st.rerun()


def fetch_external_data():
    """Fetch data from external system"""
    st.subheader(" Fetch External Data")
    
    with st.expander("Configure External Data Source"):
        source_url = st.text_input("API Endpoint URL", placeholder="https://api.example.com/data")
        auth_token = st.text_input("Authentication Token (optional)", type="password")
        parameters_json = st.text_area("Parameters (JSON)", value="{}", height=100)
        
        if st.button(" Fetch and Process", type="primary"):
            if source_url:
                try:
                    params = json.loads(parameters_json) if parameters_json else {}
                    headers = {}
                    if auth_token:
                        headers['Authorization'] = f'Bearer {auth_token}'
                    
                    response = requests.get(source_url, headers=headers, params=params, timeout=30)
                    response.raise_for_status()
                    
                    external_data = response.json()
                    
                    if isinstance(external_data, list):
                        df = pd.DataFrame(external_data)
                    elif isinstance(external_data, dict) and 'data' in external_data:
                        df = pd.DataFrame(external_data['data'])
                    else:
                        df = pd.DataFrame([external_data])
                    
                    filename = f"external_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    file_path = RAW_DIR / filename
                    df.to_excel(file_path, index=False)
                    
                    st.success(f" Fetched {len(df)} rows from external source")
                    run_pipeline()
                    time.sleep(2)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f" Error: {str(e)}")
            else:
                st.warning("Please enter an API endpoint URL")


def run_pipeline():
    """Run the Excel processing pipeline"""
    with st.spinner(" Processing Excel files..."):
        try:
            orchestrator = ExcelOrchestrator(RAW_DIR, PROCESSED_DIR)
            results = orchestrator.process_all()
            st.success(" Pipeline completed successfully!")
            return results
        except Exception as e:
            st.error(f" Pipeline error: {str(e)}")
            return None


def get_file_sheet_structure():
    """Get file and sheet structure from processed data"""
    file_structure = {}
    summary_path = PROCESSED_DIR / 'dashboard_summary.csv'
    if not summary_path.exists():
        return file_structure
    
    try:
        summary_df = pd.read_csv(summary_path)
        
        for idx, row in summary_df.iterrows():
            display_name = row['display_name']
            file_type = row.get('type', 'unknown')
            
            if file_type == 'multi_sheet':
                all_outputs = row.get('all_outputs', '')
                output_files = all_outputs.split('|') if all_outputs else []
                
                sheets_data = {}
                for output_file in output_files:
                    if output_file and Path(output_file).exists():
                        try:
                            df = pd.read_csv(output_file)
                            sheet_name = Path(output_file).stem
                            if '_' in sheet_name:
                                parts = sheet_name.split('_')
                                if len(parts) > 1:
                                    sheet_name = '_'.join(parts[1:])
                            sheets_data[sheet_name] = df
                        except Exception as e:
                            st.warning(f"Could not read {output_file}: {e}")
                
                if sheets_data:
                    file_structure[display_name] = {
                        'type': 'multi_sheet',
                        'sheets': list(sheets_data.keys()),
                        'data': sheets_data,
                        'original_file': row.get('files', display_name)
                    }
            else:
                output_file = row.get('output_file', '')
                if output_file and Path(output_file).exists():
                    try:
                        df = pd.read_csv(output_file)
                        
                        if '_source_sheet' in df.columns and df['_source_sheet'].nunique() > 1:
                            sheets = df['_source_sheet'].unique().tolist()
                            sheets_data = {}
                            for sheet in sheets:
                                sheets_data[sheet] = df[df['_source_sheet'] == sheet]
                            
                            file_structure[display_name] = {
                                'type': 'multi_sheet_tracked',
                                'sheets': sheets,
                                'data': sheets_data,
                                'original_file': row.get('files', display_name)
                            }
                        else:
                            file_structure[display_name] = {
                                'type': 'single_sheet',
                                'sheets': [display_name],
                                'data': {display_name: df},
                                'original_file': row.get('files', display_name)
                            }
                    except Exception as e:
                        st.warning(f"Could not read {output_file}: {e}")
    except Exception as e:
        st.error(f"Error reading summary: {e}")
    
    return file_structure


def generate_complete_transaction_report():
    """Generate a complete report for the currently SELECTED file only"""
    try:
        with st.spinner(f" Generating complete report for '{st.session_state.get('selected_file', 'Unknown')}'..."):
            file_structure = get_file_sheet_structure()
            
            if not file_structure:
                st.error("No data available to generate report")
                return
            
            selected_file = st.session_state.get('selected_file')
            
            if not selected_file or selected_file not in file_structure:
                st.error(f"Selected file '{selected_file}' not found")
                return
            
            generator = CompleteReportGenerator(REPORTS_DIR)
            report_path = generator.generate_report_for_file(selected_file, file_structure[selected_file])
            
            with open(report_path, "rb") as f:
                pdf_bytes = f.read()
            
            total_sheets = len(file_structure[selected_file]['sheets'])
            
            st.success(f" Report generated for '{selected_file}'! Contains {total_sheets} sheet(s)")
            
            st.download_button(
                label=" Download Complete Report (PDF)",
                data=pdf_bytes,
                file_name=Path(report_path).name,
                mime="application/pdf",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")
        st.code(traceback.format_exc())


# ============================================
# MAIN APPLICATION (Protected by Authentication)
# ============================================

# Show login form if not authenticated
if st.session_state['authentication_status'] is None or st.session_state['authentication_status'] is False:
    # Already showing login form from authenticator.login()
    if st.session_state['authentication_status'] is False:
        st.error(' Username or password is incorrect')
    elif st.session_state['authentication_status'] is None:
        st.info(' Please enter your username and password to continue')
    
    # Stop execution here - don't show dashboard
    st.stop()

# If authenticated, show dashboard
if st.session_state['authentication_status']:
    
    # Add logout button to sidebar
    with st.sidebar:
        if st.button(" Logout", use_container_width=True):
            logout()
        st.divider()
        
        # Display user info
        st.success(f" Logged in as: {st.session_state.get('name', 'User')}")
        st.caption(f"Username: {st.session_state.get('username', '')}")
    
    # ========== MAIN DASHBOARD CONTENT ==========
    st.title(PAGE_TITLE)
    st.markdown("*Enterprise-grade data visualization and analytics platform*")
    
    with st.sidebar:
        try:
            if LOGO_PATH.exists():
                st.image(str(LOGO_PATH), width=100)
            else:
                st.image(FALLBACK_LOGO_URL, width=100)
        except:
            st.image(FALLBACK_LOGO_URL, width=100)
        
        st.markdown(PAGE_TITLE)
        st.markdown("*Executive Analytics Platform*")
        st.divider()
        
        summary_path = PROCESSED_DIR / 'dashboard_summary.csv'
        if summary_path.exists():
            st.success(" Pipeline Ready")
            last_updated = datetime.fromtimestamp(summary_path.stat().st_mtime)
            st.caption(f"Last updated: {last_updated.strftime('%Y-%m-%d %H:%M')}")
        else:
            st.warning(" No data processed")
        st.divider()
    
    file_structure = get_file_sheet_structure()
    
    if not file_structure:
        st.warning(" No processed data found! Please upload Excel files.")
        tab1, tab2 = st.tabs([" Upload Files", " External Data"])
        with tab1:
            upload_excel_files()
        with tab2:
            fetch_external_data()
    else:
        with st.sidebar:
            selected_file, selected_sheet, file_info = display_file_and_sheet_selector(file_structure)
            st.divider()
            
            if st.button(" Refresh Data", use_container_width=True):
                run_pipeline()
                time.sleep(2)
                st.rerun()
            
            st.divider()
            
            page = st.radio("Navigation", [" Dashboard", " Upload Files", " External Data"], index=0)
        
        if page == " Dashboard":
            if selected_file and selected_sheet:
                try:
                    display_sheet_data_enhanced(file_info, selected_sheet)
                except Exception as e:
                    st.error(f"Error displaying sheet data: {str(e)}")
                    st.code(traceback.format_exc())
        elif page == " Upload Files":
            upload_excel_files()
        elif page == " External Data":
            fetch_external_data()