"""
Complete Report Generator - Generates report for selected file only
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

class CompleteReportGenerator:
    """Generates a complete report for the selected file only"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def _analyze_sheet_content(self, df: pd.DataFrame, sheet_name: str, file_name: str) -> dict:
        """Analyze sheet content to determine appropriate metrics and charts"""
        
        result = {
            'sheet_name': sheet_name,
            'file_name': file_name,
            'row_count': len(df),
            'column_count': df.shape[1],
            'numeric_columns': [],
            'date_columns': [],
            'categorical_columns': [],
            'total_value': 0,
            'avg_value': 0,
            'min_value': 0,
            'max_value': 0,
            'success_count': 0,
            'failed_count': 0,
            'value_column': None,
            'date_range': None,
            'sheet_type': 'general',
            'chart_data': {}
        }
        
        if df.empty:
            return result
        
        # Find numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns
        result['numeric_columns'] = list(numeric_cols)
        
        # Find date columns
        date_keywords = ['date', 'time', 'created', 'updated', 'timestamp']
        for col in df.columns:
            col_lower = col.lower()
            if any(kw in col_lower for kw in date_keywords):
                result['date_columns'].append(col)
        
        # Find categorical columns
        categorical_cols = df.select_dtypes(include=['object']).columns
        result['categorical_columns'] = [c for c in categorical_cols if not c.startswith('_')]
        
        # Check for value column
        value_keywords = ['value', 'amount', 'total', 'price', 'revenue', 'sales']
        for col in numeric_cols:
            col_lower = col.lower()
            for kw in value_keywords:
                if kw in col_lower:
                    result['value_column'] = col
                    result['total_value'] = df[col].sum()
                    result['avg_value'] = df[col].mean()
                    result['min_value'] = df[col].min()
                    result['max_value'] = df[col].max()
                    result['sheet_type'] = 'value'
                    break
            if result['value_column']:
                break
        
        # If no value column found, use first numeric column
        if result['value_column'] is None and len(numeric_cols) > 0:
            result['value_column'] = numeric_cols[0]
            result['total_value'] = df[numeric_cols[0]].sum()
        
        # Check for status data
        status_keywords = ['status', 'state', 'result']
        for col in df.columns:
            col_lower = col.lower()
            if any(kw in col_lower for kw in status_keywords):
                success_keywords = ['success', 'succeeded', 'completed', 'approved', 'done']
                failure_keywords = ['failed', 'fail', 'error', 'declined', 'rejected']
                
                for status, count in df[col].value_counts().items():
                    status_str = str(status).lower()
                    if any(k in status_str for k in success_keywords):
                        result['success_count'] += count
                    elif any(k in status_str for k in failure_keywords):
                        result['failed_count'] += count
                
                if result['success_count'] > 0 or result['failed_count'] > 0:
                    result['sheet_type'] = 'transaction'
                break
        
        # Get date range
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
        
        # Prepare chart data
        result['chart_data'] = self._prepare_chart_data(df, result)
        
        return result
    
    def _save_chart_safely(self, fig, filename):
        """Safely save chart to image with wider dimensions"""
        try:
            fig.write_image(filename, scale=1.2, width=700, height=300)
            return True
        except:
            try:
                fig.write_image(filename, engine='kaleido', scale=1.2, width=700, height=300)
                return True
            except:
                return False
    
    def _prepare_chart_data(self, df: pd.DataFrame, analysis: dict) -> dict:
        """Prepare chart data with wider dimensions"""
        
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
        
        # Bar chart
        if analysis['categorical_columns'] and analysis['value_column']:
            try:
                bar_data = df.groupby(analysis['categorical_columns'][0])[analysis['value_column']].sum().reset_index()
                bar_data = bar_data.sort_values(analysis['value_column'], ascending=False).head(10)
                
                fig = px.bar(bar_data, x=analysis['categorical_columns'][0], y=analysis['value_column'],
                            title=f"{analysis['categorical_columns'][0]} by {analysis['value_column']}",
                            color=analysis['value_column'], color_continuous_scale='Viridis')
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
        
        # Pie chart
        if analysis['categorical_columns']:
            try:
                if analysis['value_column']:
                    pie_data = df.groupby(analysis['categorical_columns'][0])[analysis['value_column']].sum().reset_index()
                    pie_data = pie_data.sort_values(analysis['value_column'], ascending=False).head(8)
                    values = pie_data[analysis['value_column']]
                    names = pie_data[analysis['categorical_columns'][0]]
                    title = f"Distribution of {analysis['categorical_columns'][0]}"
                else:
                    pie_data = df[analysis['categorical_columns'][0]].value_counts().head(8).reset_index()
                    pie_data.columns = [analysis['categorical_columns'][0], 'count']
                    values = pie_data['count']
                    names = pie_data[analysis['categorical_columns'][0]]
                    title = f"Distribution of {analysis['categorical_columns'][0]}"
                
                fig = px.pie(values=values, names=names, title=title, hole=0.3,
                            color_discrete_sequence=px.colors.qualitative.Set3)
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
        
        # Trend chart
        if analysis['date_columns'] and analysis['value_column']:
            try:
                df_temp = df.copy()
                df_temp[analysis['date_columns'][0]] = pd.to_datetime(df_temp[analysis['date_columns'][0]], errors='coerce')
                df_temp = df_temp.dropna(subset=[analysis['date_columns'][0]])
                if len(df_temp) > 3:
                    df_temp['month'] = df_temp[analysis['date_columns'][0]].dt.to_period('M')
                    trend_data = df_temp.groupby('month')[analysis['value_column']].sum().reset_index()
                    trend_data['month_str'] = trend_data['month'].astype(str)
                    
                    fig = px.line(trend_data, x='month_str', y=analysis['value_column'],
                                title=f"{analysis['value_column']} Trend",
                                markers=True, color_discrete_sequence=['#3498db'])
                    fig.update_layout(
                        height=300, width=700, template='plotly_white',
                        title_font_size=12, title_x=0.5,
                        xaxis_title="Month", yaxis_title=analysis['value_column'],
                        margin=dict(t=35, l=40, r=20, b=40),
                        font=dict(size=9)
                    )
                    chart_data['has_trend_chart'] = True
                    chart_data['trend_fig'] = fig
            except:
                pass
        
        # Status chart
        if analysis['success_count'] > 0 or analysis['failed_count'] > 0:
            try:
                fig = go.Figure(data=[go.Pie(
                    labels=['Successful', 'Failed'],
                    values=[analysis['success_count'], analysis['failed_count']],
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
        """Generate a complete report for a single selected file"""
        
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        
        temp_files = []
        
        try:
            # Analyze all sheets in this file only
            sheets_analysis = []
            
            for sheet_name, df in file_info['data'].items():
                if df is not None and not df.empty:
                    analysis = self._analyze_sheet_content(df, sheet_name, file_name)
                    sheets_analysis.append(analysis)
            
            if not sheets_analysis:
                raise Exception("No sheets with data found in this file")
            
            # Cover Page
            pdf.add_page()
            self._add_cover_page(pdf, file_name, len(sheets_analysis), sheets_analysis)
            
            # Executive Summary
            # pdf.add_page()
            # self._add_executive_summary(pdf, sheets_analysis)
            
            # Individual Sheet Reports
            for sheet in sheets_analysis:
                self._add_individual_sheet_report(pdf, sheet, temp_files)
            
            # Final Summary
            pdf.add_page()
            self._add_final_summary(pdf, sheets_analysis)
            
            # Generate PDF
            pdf_bytes = pdf.output(dest='S').encode('latin-1')
            
            # Clean up temp files
            for temp_path in temp_files:
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except:
                    pass
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{file_name}_complete_report_{timestamp}.pdf"
            # Remove invalid characters from filename
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
    
    def _add_cover_page(self, pdf: FPDF, file_name: str, sheet_count: int, sheets_analysis: list):
        """Add cover page for selected file"""
        
        pdf.set_fill_color(102, 126, 234)
        pdf.rect(0, 0, 297, 50, 'F')
        
        pdf.set_y(18)
        pdf.set_font("Arial", "B", 24)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, "SPS Complete Data Report", ln=True, align="C")
        
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 8, f"File: {self._clean_text(file_name)}", ln=True, align="C")
        
        pdf.set_y(45)
        pdf.set_font("Arial", "I", 9)
        pdf.cell(0, 5, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
        
        # pdf.ln(15)
        
        # pdf.set_text_color(0, 0, 0)
        
        # total_rows = sum(s['row_count'] for s in sheets_analysis)
        
        # box_width = 85
        # box_height = 35
        # start_x = 55
        # start_y = pdf.get_y()
        
        # # Sheets Box
        # pdf.set_fill_color(248, 249, 250)
        # pdf.rect(start_x, start_y, box_width, box_height, 'F')
        # pdf.set_draw_color(46, 204, 113)
        # pdf.rect(start_x, start_y, box_width, box_height, 'D')
    
        
        # pdf.set_xy(start_x + 5, start_y + 4)
        # pdf.set_font("Arial", "", 8)
        # pdf.set_text_color(100, 100, 100)
        # pdf.cell(box_width - 10, 5, "Sheets in File", ln=True, align='C')
        # pdf.set_xy(start_x + 5, start_y + 14)
        # pdf.set_font("Arial", "B", 16)
        # pdf.set_text_color(46, 204, 113)
        # pdf.cell(box_width - 10, 8, f"{sheet_count}", ln=True, align='C')
        
        # # Rows Box
        # x = start_x + box_width + 10
        # pdf.set_fill_color(248, 249, 250)
        # pdf.rect(x, start_y, box_width, box_height, 'F')
        # pdf.set_draw_color(241, 196, 15)
        # pdf.rect(x, start_y, box_width, box_height, 'D')
        
        # pdf.set_xy(x + 5, start_y + 4)
        # pdf.set_font("Arial", "", 8)
        # pdf.set_text_color(100, 100, 100)
        # pdf.cell(box_width - 10, 5, "Total Records", ln=True, align='C')
        # pdf.set_xy(x + 5, start_y + 14)
        # pdf.set_font("Arial", "B", 16)
        # pdf.set_text_color(241, 196, 15)
        # pdf.cell(box_width - 10, 8, f"{total_rows:,}", ln=True, align='C')
        
        # pdf.set_y(start_y + box_height + 12)
    
    def _add_executive_summary(self, pdf: FPDF, sheets_analysis: list):
        """Add executive summary page"""
        
        # pdf.set_font("Arial", "B", 16)
        # pdf.cell(0, 10, "Executive Summary", ln=True, align="C")
        # pdf.ln(5)
        
        # total_rows = sum(s['row_count'] for s in sheets_analysis)
        # total_value = sum(s['total_value'] for s in sheets_analysis)
        # total_success = sum(s['success_count'] for s in sheets_analysis)
        # total_failed = sum(s['failed_count'] for s in sheets_analysis)
        
        # card_width = 95
        # card_height = 35
        # start_x = 15
        # start_y = pdf.get_y()
        
        # metrics = [
        #     ('Total Records', f"{total_rows:,}", '#3498db'),
        #     ('Total Value', f"{total_value:,.0f}", '#2ecc71'),
        #     ('Success', f"{total_success:,}", '#27ae60'),
        #     ('Failed', f"{total_failed:,}", '#e74c3c')
        # ]
        
        # for i, (label, value, color) in enumerate(metrics):
        #     x = start_x + i * (card_width + 8)
        #     y = start_y
            
        #     pdf.set_fill_color(248, 249, 250)
        #     pdf.rect(x, y, card_width, card_height, 'F')
        #     pdf.set_draw_color(52, 152, 219)
        #     pdf.rect(x, y, card_width, card_height, 'D')
            
        #     pdf.set_xy(x + 5, y + 4)
        #     pdf.set_font("Arial", "", 8)
        #     pdf.set_text_color(100, 100, 100)
        #     pdf.cell(card_width - 10, 5, label, ln=True, align='C')
            
        #     pdf.set_xy(x + 5, y + 14)
        #     pdf.set_font("Arial", "B", 13)
        #     pdf.set_text_color(52, 152, 219)
        #     pdf.cell(card_width - 10, 8, value, ln=True, align='C')
        
        # pdf.set_y(start_y + card_height + 10)
        
        # # Sheets summary table
        # pdf.set_font("Arial", "B", 11)
        # pdf.cell(0, 7, "Sheets Summary", ln=True)
        
        # pdf.set_font("Arial", "B", 7)
        # pdf.set_fill_color(52, 152, 219)
        # pdf.set_text_color(255, 255, 255)
        
        # pdf.cell(80, 6, "Sheet Name", border=1, fill=True)
        # pdf.cell(30, 6, "Rows", border=1, fill=True, align='C')
        # pdf.cell(40, 6, "Value", border=1, fill=True, align='C')
        # pdf.cell(30, 6, "Success", border=1, fill=True, align='C')
        # pdf.cell(30, 6, "Failed", border=1, fill=True, align='C')
        # pdf.cell(35, 6, "Type", border=1, fill=True, align='C')
        # pdf.ln()
        
        # pdf.set_font("Arial", "", 6)
        # pdf.set_text_color(0, 0, 0)
        
        # for sheet in sheets_analysis:
        #     pdf.cell(80, 5, self._clean_text(sheet['sheet_name'][:35]), border=1)
        #     pdf.cell(30, 5, f"{sheet['row_count']:,}", border=1, align='C')
        #     pdf.cell(40, 5, f"{sheet['total_value']:,.0f}", border=1, align='C')
        #     pdf.cell(30, 5, f"{sheet['success_count']:,}", border=1, align='C')
        #     pdf.cell(30, 5, f"{sheet['failed_count']:,}", border=1, align='C')
        #     pdf.cell(35, 5, sheet['sheet_type'].upper(), border=1, align='C')
        #     pdf.ln()
    
    def _add_individual_sheet_report(self, pdf: FPDF, sheet: dict, temp_files: list):
        """Add individual sheet report with wider charts - each chart on its own page"""
        
        chart_data = sheet['chart_data']
        
        # PAGE 1: Sheet Header + KPIs
        pdf.add_page()
        
        # Sheet header
        pdf.set_font("Arial", "B", 14)
        pdf.set_fill_color(52, 152, 219)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, f"Sheet: {self._clean_text(sheet['sheet_name'])}", ln=True, fill=True)
        
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "I", 9)
        pdf.cell(0, 5, f"Type: {sheet['sheet_type'].upper()} | Rows: {sheet['row_count']:,} | Columns: {sheet['column_count']}", ln=True)
        pdf.ln(3)
        
        # KPIs
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 7, "Key Metrics", ln=True)
        
        card_width = 90
        card_height = 30
        start_x = 15
        start_y = pdf.get_y()
        
        kpis = [('Records', f"{sheet['row_count']:,}", '#3498db')]
        
        if sheet['total_value'] > 0:
            kpis.append(('Value', f"{sheet['total_value']:,.0f}", '#2ecc71'))
            kpis.append(('Avg', f"{sheet['avg_value']:,.0f}", '#27ae60'))
        
        if sheet['success_count'] > 0 or sheet['failed_count'] > 0:
            kpis.append(('Success', f"{sheet['success_count']:,}", '#27ae60'))
            kpis.append(('Failed', f"{sheet['failed_count']:,}", '#e74c3c'))
            total = sheet['success_count'] + sheet['failed_count']
            success_pct = (sheet['success_count'] / total * 100) if total > 0 else 0
            kpis.append(('Rate', f"{success_pct:.0f}%", '#2ecc71'))
        
        kpis = kpis[:6]
        
        for i, (label, value, color) in enumerate(kpis):
            row = i // 3
            col = i % 3
            x = start_x + col * (card_width + 8)
            y = start_y + row * (card_height + 6)
            
            pdf.set_fill_color(248, 249, 250)
            pdf.rect(x, y, card_width, card_height, 'F')
            pdf.set_draw_color(52, 152, 219)
            pdf.rect(x, y, card_width, card_height, 'D')
            
            pdf.set_xy(x + 5, y + 4)
            pdf.set_font("Arial", "", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(card_width - 10, 5, label, ln=True, align='C')
            
            pdf.set_xy(x + 5, y + 13)
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(52, 152, 219)
            pdf.cell(card_width - 10, 8, value, ln=True, align='C')
        
        # Date range at bottom of first page
        if sheet['date_range']:
            pdf.set_y(195)
            pdf.set_font("Arial", "I", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 5, f"Date Range: {sheet['date_range']['start']} to {sheet['date_range']['end']}", ln=True, align='C')
        
        # ========== EACH CHART ON ITS OWN PAGE ==========
        
        # Status Chart Page
        if chart_data['has_status_chart']:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 10, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Status Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(8)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_status_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['status_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=65, y=35, w=170)
        
        # Bar Chart Page
        if chart_data['has_bar_chart']:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 10, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Bar Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(8)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_bar_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['bar_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=15, y=35, w=270)
        
        # Pie Chart Page
        if chart_data['has_pie_chart']:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 10, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Pie Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(8)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_pie_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['pie_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=55, y=35, w=190)
        
        # Trend Chart Page
        if chart_data['has_trend_chart']:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 10, f"Sheet: {self._clean_text(sheet['sheet_name'])} - Trend Chart", ln=True, fill=True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(8)
            
            temp_path = os.path.join(tempfile.gettempdir(), f"chart_trend_{uuid.uuid4().hex}.png")
            if self._save_chart_safely(chart_data['trend_fig'], temp_path):
                temp_files.append(temp_path)
                pdf.image(temp_path, x=15, y=35, w=270)
    
    def _add_final_summary(self, pdf: FPDF, sheets_analysis: list):
        """Add final summary page"""
        
        # pdf.set_font("Arial", "B", 16)
        # pdf.cell(0, 10, "Report Summary", ln=True, align="C")
        # pdf.ln(5)
        
        # total_rows = sum(s['row_count'] for s in sheets_analysis)
        # total_value = sum(s['total_value'] for s in sheets_analysis)
        # total_success = sum(s['success_count'] for s in sheets_analysis)
        # total_failed = sum(s['failed_count'] for s in sheets_analysis)
        
        # pdf.set_font("Arial", "B", 12)
        # pdf.cell(0, 7, "Overall Statistics", ln=True)
        
        # pdf.set_font("Arial", "", 10)
        # pdf.cell(0, 6, f"Total Sheets in File: {len(sheets_analysis)}", ln=True)
        # pdf.cell(0, 6, f"Total Records: {total_rows:,}", ln=True)
        
        # if total_value > 0:
        #     pdf.cell(0, 6, f"Total Value: {total_value:,.0f}", ln=True)
        
        # if total_success > 0 or total_failed > 0:
        #     pdf.cell(0, 6, f"Successful: {total_success:,} | Failed: {total_failed:,}", ln=True)
        #     success_pct = (total_success / (total_success + total_failed) * 100) if (total_success + total_failed) > 0 else 0
        #     pdf.cell(0, 6, f"Overall Success Rate: {success_pct:.1f}%", ln=True)
        
        # pdf.ln(5)
        # pdf.set_font("Arial", "B", 11)
        # pdf.cell(0, 6, "Sheets in this File:", ln=True)
        # pdf.set_font("Arial", "", 8)
        
        # for sheet in sheets_analysis:
        #     sheet_type = sheet['sheet_type'].upper()
        #     pdf.cell(0, 4, f"  - {self._clean_text(sheet['sheet_name'])} [{sheet_type}] ({sheet['row_count']:,} rows)", ln=True)
        
        pdf.ln(5)
        pdf.set_font("Arial", "I", 8)
        pdf.set_text_color(150, 150, 150)
        pdf.cell(0, 5, "Generated by SPS Excel Pipeline - Complete Data Report", ln=True, align="C")
        pdf.cell(0, 5, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")