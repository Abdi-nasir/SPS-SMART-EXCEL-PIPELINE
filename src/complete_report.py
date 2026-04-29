"""
Complete Report Generator - Generates report for selected file only
KPIs from high_level_metrics sheet displayed on cover page, Charts for all other sheets
No blank pages
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
import tempfile
import os
import uuid
from fpdf import FPDF
from config import HIGH_LEVEL_METRICS_PATTERNS

class CompleteReportGenerator:
    """Generates a complete report for the selected file only - No blank pages"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def _is_high_level_metrics_sheet(self, sheet_name: str) -> bool:
        """Check if sheet is a high level metrics sheet"""
        sheet_lower = sheet_name.lower()
        for pattern in HIGH_LEVEL_METRICS_PATTERNS:
            if pattern in sheet_lower:
                return True
        return False
    
    def _extract_kpis_from_sheet(self, df):
        result = {'categories': {}, 'all_metrics': []}
        
        if df.empty:
            return result

        category_col = None
        metric_col = None
        value_col = None
        payment_vol = None
        payment_val = None

        # Detect columns dynamically
        for col in df.columns:
            col_lower = str(col).lower()

            if 'category' in col_lower:
                category_col = col
            elif 'metric' in col_lower or 'kpi' in col_lower or 'name' in col_lower:
                metric_col = col
            elif 'value' in col_lower:
                value_col = col
            elif 'volume' in col_lower:
                payment_vol = col
            elif 'value' in col_lower and 'payment' in col_lower:
                payment_val = col

        if not (metric_col and value_col):
            return result

        for _, row in df.iterrows():
            category = str(row.get(category_col, 'General')).strip()
            metric = str(row.get(metric_col, '')).strip()
            value = str(row.get(value_col, '')).strip()

            payment_volume = str(row.get(payment_vol, '')).strip() if payment_vol else None
            payment_value = str(row.get(payment_val, '')).strip() if payment_val else None

            if not metric or metric.lower() in ['metric', 'kpi', 'name']:
                continue

            # Format numeric values
            def format_number(val):
                try:
                    num = float(str(val).replace(',', ''))
                    return f"{int(num):,}" if num.is_integer() else f"{num:,.2f}"
                except:
                    return val

            metric_data = {
                'name': metric,
                'value': format_number(value),
                'payment_volume': format_number(payment_volume) if payment_volume else None,
                'payment_value': format_number(payment_value) if payment_value else None,
                'color': self._get_kpi_color(metric)
            }

            result['categories'].setdefault(category, []).append(metric_data)
            result['all_metrics'].append(metric_data)

        return result
    
    def _get_kpi_color(self, kpi_name: str) -> tuple:
        """Get color for KPI based on name"""
        kpi_lower = kpi_name.lower()
        
        if 'volume' in kpi_lower or 'count' in kpi_lower or 'transaction' in kpi_lower:
            return (52, 152, 219)  # Blue
        elif 'value' in kpi_lower or 'amount' in kpi_lower or 'revenue' in kpi_lower:
            return (46, 204, 113)  # Green
        elif 'failed' in kpi_lower or 'failure' in kpi_lower or 'error' in kpi_lower:
            return (231, 76, 60)   # Red
        elif 'rate' in kpi_lower or 'percentage' in kpi_lower or '%' in kpi_lower:
            return (155, 89, 182)  # Purple
        else:
            return (52, 152, 219)  # Default Blue
    
    def _prepare_chart_data(self, df: pd.DataFrame, analysis: dict) -> dict:
        """Prepare chart data for non-KPI sheets"""
        
        chart_data = {
            'has_bar_chart': False,
            'has_pie_chart': False,
            'has_trend_chart': False,
            'has_status_chart': False,
            'bar_fig': None,
            'pie_fig': None,
            'trend_fig': None,
            'status_fig': None
        }
        
        # Find columns
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = analysis.get('categorical_columns', [])
        date_cols = analysis.get('date_columns', [])
        
        if not numeric_cols:
            return chart_data
        
        value_column = numeric_cols[0]
        
        # Bar Chart
        if categorical_cols:
            try:
                bar_data = df.groupby(categorical_cols[0])[value_column].sum().reset_index()
                bar_data = bar_data.sort_values(value_column, ascending=False).head(10)
                
                fig = px.bar(bar_data, x=categorical_cols[0], y=value_column,
                            title=f"{categorical_cols[0]} by {value_column}",
                            color=value_column, color_continuous_scale='Viridis')
                fig.update_layout(
                    height=300, width=700, template='plotly_white',
                    title_font_size=12, title_x=0.5,
                    margin=dict(t=35, l=40, r=20, b=40),
                    xaxis_tickangle=-45,
                    font=dict(size=9)
                )
                chart_data['has_bar_chart'] = True
                chart_data['bar_fig'] = fig
            except:
                pass
        
        # Pie Chart
        if categorical_cols:
            try:
                pie_data = df.groupby(categorical_cols[0])[value_column].sum().reset_index()
                pie_data = pie_data.sort_values(value_column, ascending=False).head(8)
                
                fig = px.pie(pie_data, values=value_column, names=categorical_cols[0],
                            title=f"Distribution of {value_column} by {categorical_cols[0]}",
                            hole=0.3, color_discrete_sequence=px.colors.qualitative.Set3)
                fig.update_traces(textposition='inside', textinfo='percent+label', textfont_size=9)
                fig.update_layout(
                    height=300, width=600, template='plotly_white',
                    title_font_size=12, title_x=0.5,
                    margin=dict(t=35, l=40, r=40, b=40),
                    font=dict(size=9)
                )
                chart_data['has_pie_chart'] = True
                chart_data['pie_fig'] = fig
            except:
                pass
        
        # Trend Chart
        if date_cols:
            try:
                df_temp = df.copy()
                df_temp[date_cols[0]] = pd.to_datetime(df_temp[date_cols[0]], errors='coerce')
                df_temp = df_temp.dropna(subset=[date_cols[0]])
                if len(df_temp) > 3:
                    df_temp['month'] = df_temp[date_cols[0]].dt.to_period('M')
                    trend_data = df_temp.groupby('month')[value_column].sum().reset_index()
                    trend_data['month_str'] = trend_data['month'].astype(str)
                    
                    fig = px.line(trend_data, x='month_str', y=value_column,
                                title=f"{value_column} Trend Over Time",
                                markers=True, color_discrete_sequence=['#3498db'])
                    fig.update_layout(
                        height=300, width=700, template='plotly_white',
                        title_font_size=12, title_x=0.5,
                        xaxis_title="Month", yaxis_title=value_column,
                        margin=dict(t=35, l=40, r=20, b=40),
                        font=dict(size=9)
                    )
                    chart_data['has_trend_chart'] = True
                    chart_data['trend_fig'] = fig
            except:
                pass
        
        # Status Chart
        status_col = None
        for col in df.columns:
            if 'status' in col.lower() or 'state' in col.lower():
                status_col = col
                break
        
        if status_col:
            try:
                success_count = 0
                failed_count = 0
                success_keywords = ['success', 'succeeded', 'completed', 'approved', 'done']
                failure_keywords = ['failed', 'fail', 'error', 'declined', 'rejected']
                
                for status, count in df[status_col].value_counts().items():
                    status_str = str(status).lower()
                    if any(k in status_str for k in success_keywords):
                        success_count += count
                    elif any(k in status_str for k in failure_keywords):
                        failed_count += count
                
                if success_count > 0 or failed_count > 0:
                    fig = go.Figure(data=[go.Pie(
                        labels=['Successful', 'Failed'],
                        values=[success_count, failed_count],
                        hole=0.4,
                        marker_colors=['#2ecc71', '#e74c3c'],
                        textinfo='label+percent',
                        textposition='auto'
                    )])
                    fig.update_layout(
                        height=280, width=550, template='plotly_white',
                        title="Success vs Failure",
                        title_font_size=12, title_x=0.5,
                        margin=dict(t=35, l=40, r=40, b=40),
                        font=dict(size=9)
                    )
                    chart_data['has_status_chart'] = True
                    chart_data['status_fig'] = fig
            except:
                pass
        
        return chart_data
    
    def _analyze_sheet_content(self, df: pd.DataFrame, sheet_name: str, file_name: str) -> dict:
        """Analyze sheet content"""
        
        result = {
            'sheet_name': sheet_name,
            'file_name': file_name,
            'row_count': len(df),
            'column_count': df.shape[1],
            'numeric_columns': [],
            'date_columns': [],
            'categorical_columns': [],
            'date_range': None,
            'chart_data': {},
            'is_kpi_sheet': self._is_high_level_metrics_sheet(sheet_name)
        }
        
        if df.empty:
            return result
        
        if result['is_kpi_sheet']:
            result['kpi_data'] = self._extract_kpis_from_sheet(df)
            return result
        
        # For non-KPI sheets
        numeric_cols = df.select_dtypes(include=['number']).columns
        result['numeric_columns'] = list(numeric_cols)
        
        date_keywords = ['date', 'time', 'created', 'updated', 'timestamp']
        for col in df.columns:
            col_lower = col.lower()
            if any(kw in col_lower for kw in date_keywords):
                result['date_columns'].append(col)
        
        categorical_cols = df.select_dtypes(include=['object']).columns
        result['categorical_columns'] = [c for c in categorical_cols if not c.startswith('_')]
        
        if result['date_columns']:
            try:
                df_date = df.copy()
                df_date[result['date_columns'][0]] = pd.to_datetime(df_date[result['date_columns'][0]], errors='coerce')
                df_date = df_date.dropna(subset=[result['date_columns'][0]])
                if len(df_date) > 0:
                    result['date_range'] = {
                        'start': df_date[result['date_columns'][0]].min().strftime('%Y-%m-%d'),
                        'end': df_date[result['date_columns'][0]].max().strftime('%Y-%m-%d')
                    }
            except:
                pass
        
        result['chart_data'] = self._prepare_chart_data(df, result)
        
        return result
    
    def _save_chart_safely(self, fig, filename):
        """Safely save chart to image"""
        try:
            fig.write_image(filename, scale=1.2, width=700, height=300)
            return True
        except:
            try:
                fig.write_image(filename, engine='kaleido', scale=1.2, width=700, height=300)
                return True
            except:
                return False
    
    def _clean_text(self, text):
        """Clean text for PDF encoding"""
        if text is None:
            return ""
        replacements = {
            '\u2022': '-', '\u25CF': '-', '\u25CB': 'o', '\u2013': '-',
            '\u2014': '--', '\u2018': "'", '\u2019': "'", '\u201C': '"',
            '\u201D': '"', '\u00A0': ' '
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text.encode('latin-1', errors='ignore').decode('latin-1')
    
    def generate_report_for_file(self, file_name: str, file_info: dict) -> str:
        """Generate a complete report for a single selected file - No blank pages"""
        
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        pdf.set_font("Arial", "", 10)
        
        temp_files = []
        
        try:
            sheets_analysis = []
            kpi_data = None
            
            for sheet_name, df in file_info['data'].items():
                if df is not None and not df.empty:
                    analysis = self._analyze_sheet_content(df, sheet_name, file_name)
                    sheets_analysis.append(analysis)
                    
                    if analysis.get('is_kpi_sheet', False) and 'kpi_data' in analysis:
                        kpi_data = analysis['kpi_data']
            
            if not sheets_analysis:
                raise Exception("No sheets with data found in this file")
            
            # COVER PAGE with KPIs
            self._add_cover_page_with_kpis(pdf, file_name, sheets_analysis, kpi_data)
            
            # Individual Data Sheets (ONLY for sheets that have charts)
            for sheet in sheets_analysis:
                if not sheet.get('is_kpi_sheet', False):
                    chart_data = sheet.get('chart_data', {})
                    has_any_chart = (chart_data.get('has_bar_chart') or 
                                    chart_data.get('has_pie_chart') or 
                                    chart_data.get('has_trend_chart') or 
                                    chart_data.get('has_status_chart'))
                    
                    if has_any_chart:
                        self._add_data_sheet_report(pdf, sheet, temp_files)
            
          
            pdf.set_y(250)
            pdf.set_font("Arial", "I", 8)
            pdf.set_text_color(150, 150, 150)
            pdf.cell(0, 5, "Generated by SPS Excel Pipeline - Complete Data Report", ln=True, align="C")
            pdf.cell(0, 5, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
            
            # Generate PDF
            pdf_bytes = pdf.output(dest='S').encode('latin-1')
            
            for temp_path in temp_files:
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except:
                    pass
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{file_name}_complete_report_{timestamp}.pdf"
            filename = "".join(c for c in filename if c.isalnum() or c in '._- ')
            filepath = self.output_dir / filename
            
            with open(filepath, 'wb') as f:
                f.write(pdf_bytes)
            
            return str(filepath)
            
        except Exception as e:
            for temp_path in temp_files:
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except:
                    pass
            raise e
    
    def _add_cover_page_with_kpis(self, pdf: FPDF, file_name: str, sheets_analysis: list, kpi_data: dict):
        """Add cover page with KPIs - First page of PDF"""
        
        pdf.add_page()
        
        # Header
        pdf.set_fill_color(102, 126, 234)
        pdf.rect(0, 0, 297, 45, 'F')
        
        pdf.set_y(15)
        pdf.set_font("Arial", "B", 22)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, "SPS Complete Data Report", ln=True, align="C")
        
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 6, f"File: {self._clean_text(file_name)}", ln=True, align="C")
        
        pdf.set_y(42)
        pdf.set_font("Arial", "I", 8)
        pdf.cell(0, 4, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
        
        pdf.ln(8)
        
        # KPIs Section (only if available)
        if kpi_data and kpi_data.get('categories') and len(kpi_data['categories']) > 0:
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(0, 6, "Key Performance Indicators", ln=True, align="C")
            pdf.ln(4)
            
            categories = kpi_data.get('categories', {})
          
            
            for category, metrics in categories.items():
                # Category header
                pdf.set_font("Arial", "B", 10)
                pdf.set_fill_color(52, 152, 219)
                pdf.set_text_color(255, 255, 255)
                pdf.cell(0, 6, f" {self._clean_text(category)}", ln=True, fill=True)
                pdf.set_text_color(0, 0, 0)
                pdf.ln(2)
                
                card_width = 85
                card_height = 32
                current_x = 15
                current_y = pdf.get_y()
                
                for i, metric in enumerate(metrics):
                    if i > 0 and i % 3 == 0:
                        current_y += card_height + 4
                        current_x = 15
                    
                    x = current_x
                    y = current_y
                    
                    pdf.set_fill_color(248, 249, 250)
                    pdf.rect(x, y, card_width, card_height, 'F')
                    pdf.set_draw_color(metric['color'][0], metric['color'][1], metric['color'][2])
                    pdf.rect(x, y, card_width, card_height, 'D')
                    
                    pdf.set_xy(x + 3, y + 3)
                    pdf.set_font("Arial", "", 7)
                    pdf.set_text_color(100, 100, 100)
                    pdf.cell(card_width - 6, 4, self._clean_text(metric['name'][:22]), ln=True, align='C')
                    
                    pdf.set_xy(x + 3, y + 12)
                    pdf.set_font("Arial", "B", 11)
                    pdf.set_text_color(metric['color'][0], metric['color'][1], metric['color'][2])
                    pdf.cell(card_width - 6, 7, self._clean_text(metric['value']), ln=True, align='C')
                    
                    current_x += card_width + 5
                
                pdf.set_y(current_y + card_height + 6)
                pdf.ln(2)
    
    def _add_data_sheet_report(self, pdf: FPDF, sheet: dict, temp_files: list):
        """Add data sheet report with charts - No blank pages"""
        
        chart_data = sheet.get('chart_data', {})
        
        # Sheet information page
        pdf.add_page()
        
        pdf.set_font("Arial", "B", 13)
        pdf.set_fill_color(52, 152, 219)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 9, f"Sheet: {self._clean_text(sheet['sheet_name'])}", ln=True, fill=True)
        
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "I", 8)
        pdf.cell(0, 4, f"Rows: {sheet['row_count']:,} | Columns: {sheet['column_count']}", ln=True)
        pdf.ln(2)
        
        if sheet.get('date_range'):
            pdf.set_font("Arial", "I", 7)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 4, f"Date Range: {sheet['date_range']['start']} to {sheet['date_range']['end']}", ln=True)
            pdf.ln(2)
        
        # Bar Chart
        if chart_data.get('has_bar_chart'):
            pdf.add_page()
            pdf.set_font("Arial", "B", 13)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 9, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Bar Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(6)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_bar_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['bar_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=15, y=30, w=270)
        
        # Pie Chart
        if chart_data.get('has_pie_chart'):
            pdf.add_page()
            pdf.set_font("Arial", "B", 13)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 9, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Pie Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(6)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_pie_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['pie_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=55, y=30, w=190)
        
        # Trend Chart
        if chart_data.get('has_trend_chart'):
            pdf.add_page()
            pdf.set_font("Arial", "B", 13)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 9, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Trend Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(6)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_trend_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['trend_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=15, y=30, w=270)
        
        # Status Chart
        if chart_data.get('has_status_chart'):
            pdf.add_page()
            pdf.set_font("Arial", "B", 13)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 9, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Status Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(6)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_status_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['status_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=65, y=30, w=170)