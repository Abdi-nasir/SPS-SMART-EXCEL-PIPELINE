"""
PDF Report Generator - Creates PDF reports with ONLY charts for each sheet
No statistics, no tables, just visualizations
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
import tempfile
import os

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# Custom PDF class with Unicode support
class ChartsOnlyPDF(FPDF):
    """Custom PDF class for charts-only reports"""
    
    def __init__(self):
        super().__init__()
        # Try to use DejaVu font if available, otherwise fallback to Arial
        try:
            self.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
            self.set_font('DejaVu', '', 10)
            self.use_unicode = True
        except:
            self.set_font('Arial', '', 10)
            self.use_unicode = False
    
    def safe_text(self, text):
        """Safely encode text for PDF"""
        if self.use_unicode:
            return text
        # Replace problematic characters for Arial
        replacements = {
            '\u2022': '-', '\u25CF': '-', '\u25CB': '-', '\u2013': '-',
            '\u2014': '--', '\u2018': "'", '\u2019': "'", '\u201C': '"',
            '\u201D': '"', '\u00A0': ' '
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text.encode('latin-1', errors='ignore').decode('latin-1')

class PDFReportGenerator:
    """Generates PDF reports with ONLY charts for each sheet"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def generate_complete_report(self, file_structure: dict) -> str:
        """Generate a PDF report with charts only for ALL sheets in ALL files"""
        if not PDF_AVAILABLE:
            raise ImportError("Please install: pip install fpdf pillow matplotlib kaleido")
        
        pdf = ChartsOnlyPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Add title page
        self._add_title_page(pdf, file_structure)
        
        # Process EACH file and EACH sheet - charts only
        for file_name, file_info in file_structure.items():
            for sheet_name, df in file_info['data'].items():
                if df is not None and not df.empty:
                    self._add_sheet_charts(pdf, file_name, sheet_name, df)
        
        # Save PDF
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"charts_report_{timestamp}.pdf"
        filepath = self.output_dir / filename
        pdf.output(str(filepath))
        
        return str(filepath)
    
    def _add_title_page(self, pdf: ChartsOnlyPDF, file_structure: dict):
        """Add simple title page"""
        pdf.add_page()
        
        pdf.set_font("Arial", "B", 24)
        pdf.cell(0, 40, "SPS Excel Pipeline", ln=True, align="C")
        pdf.set_font("Arial", "B", 18)
        pdf.cell(0, 10, "Charts Dashboard Report", ln=True, align="C")
        
        pdf.ln(20)
        
        pdf.set_font("Arial", "", 11)
        pdf.cell(0, 8, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
        
        total_files = len(file_structure)
        total_sheets = sum(len(info['sheets']) for info in file_structure.values())
        pdf.cell(0, 8, f"Files: {total_files}  |  Sheets: {total_sheets}", ln=True, align="C")
        
        pdf.ln(20)
        pdf.set_font("Arial", "I", 10)
        pdf.cell(0, 8, "This report contains visualizations only", ln=True, align="C")
    
    def _add_sheet_charts(self, pdf: ChartsOnlyPDF, file_name: str, sheet_name: str, df: pd.DataFrame):
        """Add charts for a single sheet - no statistics, just visualizations"""
        
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
        
        if not numeric_cols:
            return
        
        # Add sheet header
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.set_fill_color(52, 152, 219)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 12, f"Sheet: {sheet_name}", ln=True, fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "I", 10)
        pdf.cell(0, 6, f"Source: {file_name}", ln=True)
        pdf.ln(5)
        
        with tempfile.TemporaryDirectory() as tmpdir:
            charts_added = 0
            
            # Chart 1: Bar Chart (if categorical and numeric columns exist)
            if categorical_cols and numeric_cols:
                try:
                    agg_data = df.groupby(categorical_cols[0])[numeric_cols[0]].sum().reset_index()
                    agg_data = agg_data.sort_values(numeric_cols[0], ascending=False).head(10)
                    
                    fig = px.bar(
                        agg_data, 
                        x=categorical_cols[0], 
                        y=numeric_cols[0],
                        title=f"Top 10 {categorical_cols[0]} by {numeric_cols[0]}",
                        color=numeric_cols[0],
                        color_continuous_scale='Viridis'
                    )
                    fig.update_layout(height=400, width=700)
                    
                    chart_path = os.path.join(tmpdir, f"bar_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    pdf.image(chart_path, x=10, w=190)
                    pdf.ln(85)
                    charts_added += 1
                except Exception as e:
                    pdf.cell(0, 6, f"Bar chart not available", ln=True)
            
            # Chart 2: Pie Chart (if categorical column exists)
            if categorical_cols and charts_added < 2:
                try:
                    pie_data = df[categorical_cols[0]].value_counts().head(8)
                    
                    fig = px.pie(
                        values=pie_data.values,
                        names=pie_data.index,
                        title=f"Distribution of {categorical_cols[0]}",
                        hole=0.3
                    )
                    fig.update_layout(height=400, width=700)
                    
                    chart_path = os.path.join(tmpdir, f"pie_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    pdf.image(chart_path, x=10, w=190)
                    pdf.ln(85)
                    charts_added += 1
                except Exception as e:
                    pdf.cell(0, 6, f"Pie chart not available", ln=True)
            
            # Chart 3: Histogram (if numeric columns exist)
            if numeric_cols and charts_added < 3:
                try:
                    fig = go.Figure()
                    fig.add_trace(go.Histogram(x=df[numeric_cols[0]].dropna(), nbinsx=20))
                    fig.update_layout(
                        title=f"Distribution of {numeric_cols[0]}",
                        xaxis_title=numeric_cols[0],
                        yaxis_title="Frequency",
                        height=400,
                        width=700
                    )
                    
                    chart_path = os.path.join(tmpdir, f"hist_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    pdf.image(chart_path, x=10, w=190)
                    pdf.ln(85)
                    charts_added += 1
                except Exception as e:
                    pdf.cell(0, 6, f"Histogram not available", ln=True)
            
            # Chart 4: Box Plot (if categorical and numeric columns exist)
            if categorical_cols and numeric_cols and charts_added < 4:
                try:
                    # Limit to top 5 categories for readability
                    top_categories = df[categorical_cols[0]].value_counts().head(5).index.tolist()
                    df_filtered = df[df[categorical_cols[0]].isin(top_categories)]
                    
                    fig = px.box(
                        df_filtered,
                        x=categorical_cols[0],
                        y=numeric_cols[0],
                        title=f"Distribution of {numeric_cols[0]} by {categorical_cols[0]}"
                    )
                    fig.update_layout(height=400, width=700)
                    
                    chart_path = os.path.join(tmpdir, f"box_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    pdf.image(chart_path, x=10, w=190)
                    pdf.ln(85)
                    charts_added += 1
                except Exception as e:
                    pdf.cell(0, 6, f"Box plot not available", ln=True)
            
            # Chart 5: Line Chart (if date columns exist)
            date_cols = [c for c in df.columns if 'date' in c.lower() or 'time' in c.lower()]
            if date_cols and numeric_cols and charts_added < 5:
                try:
                    df_line = df.copy()
                    df_line[date_cols[0]] = pd.to_datetime(df_line[date_cols[0]], errors='coerce')
                    df_line = df_line.dropna(subset=[date_cols[0]])
                    
                    if len(df_line) > 0:
                        df_line['period'] = df_line[date_cols[0]].dt.to_period('M')
                        trend_data = df_line.groupby('period')[numeric_cols[0]].sum().reset_index()
                        trend_data['period_str'] = trend_data['period'].astype(str)
                        
                        fig = px.line(
                            trend_data,
                            x='period_str',
                            y=numeric_cols[0],
                            title=f"{numeric_cols[0]} Trend Over Time",
                            markers=True
                        )
                        fig.update_layout(height=400, width=700)
                        
                        chart_path = os.path.join(tmpdir, f"line_{sheet_name}.png")
                        fig.write_image(chart_path, engine="kaleido")
                        pdf.image(chart_path, x=10, w=190)
                        pdf.ln(85)
                        charts_added += 1
                except Exception as e:
                    pdf.cell(0, 6, f"Line chart not available", ln=True)
            
            if charts_added == 0:
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 6, "No charts could be generated for this sheet", ln=True, align="C")
            else:
                pdf.set_font("Arial", "I", 9)
                pdf.cell(0, 6, f"Generated {charts_added} chart(s)", ln=True, align="R")


class SimpleChartsReport:
    """Simpler version - just charts, no extra formatting"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def generate_charts_only_report(self, file_structure: dict) -> str:
        """Generate a charts-only PDF report"""
        if not PDF_AVAILABLE:
            raise ImportError("Please install: pip install fpdf pillow matplotlib kaleido")
        
        pdf = ChartsOnlyPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Title
        pdf.add_page()
        pdf.set_font("Arial", "B", 20)
        pdf.cell(0, 30, "Charts Dashboard", ln=True, align="C")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 8, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ln=True, align="C")
        pdf.ln(10)
        
        # Process each sheet
        for file_name, file_info in file_structure.items():
            for sheet_name, df in file_info['data'].items():
                if df is not None and not df.empty:
                    self._add_charts_page(pdf, sheet_name, df)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"charts_only_{timestamp}.pdf"
        filepath = self.output_dir / filename
        pdf.output(str(filepath))
        
        return str(filepath)
    
    def _add_charts_page(self, pdf: ChartsOnlyPDF, sheet_name: str, df: pd.DataFrame):
        """Add a page with charts for a sheet"""
        
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        categorical_cols = [c for c in categorical_cols if not c.startswith('_')]
        
        if not numeric_cols:
            return
        
        pdf.add_page()
        
        # Sheet title
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, sheet_name, ln=True, align="C")
        pdf.ln(5)
        
        with tempfile.TemporaryDirectory() as tmpdir:
            chart_positions = [(10, 50), (105, 50)]  # Two charts per page positions
            chart_idx = 0
            
            # Bar Chart
            if categorical_cols and numeric_cols and chart_idx < 4:
                try:
                    agg_data = df.groupby(categorical_cols[0])[numeric_cols[0]].sum().reset_index()
                    agg_data = agg_data.sort_values(numeric_cols[0], ascending=False).head(8)
                    
                    fig = px.bar(agg_data, x=categorical_cols[0], y=numeric_cols[0],
                                title=f"{numeric_cols[0]} by {categorical_cols[0]}")
                    fig.update_layout(height=250, width=250, title_font_size=10)
                    
                    chart_path = os.path.join(tmpdir, f"bar_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    
                    x, y = chart_positions[chart_idx % 2]
                    pdf.image(chart_path, x=x, y=y, w=85)
                    chart_idx += 1
                except:
                    pass
            
            # Pie Chart
            if categorical_cols and chart_idx < 4:
                try:
                    pie_data = df[categorical_cols[0]].value_counts().head(6)
                    fig = px.pie(values=pie_data.values, names=pie_data.index,
                               title=f"{categorical_cols[0]} Distribution", hole=0.3)
                    fig.update_layout(height=250, width=250, title_font_size=10)
                    
                    chart_path = os.path.join(tmpdir, f"pie_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    
                    x, y = chart_positions[chart_idx % 2]
                    pdf.image(chart_path, x=x, y=y, w=85)
                    chart_idx += 1
                except:
                    pass
            
            # Histogram
            if numeric_cols and chart_idx < 4:
                try:
                    fig = go.Figure()
                    fig.add_trace(go.Histogram(x=df[numeric_cols[0]].dropna(), nbinsx=15))
                    fig.update_layout(title=f"{numeric_cols[0]} Distribution", height=250, width=250)
                    
                    chart_path = os.path.join(tmpdir, f"hist_{sheet_name}.png")
                    fig.write_image(chart_path, engine="kaleido")
                    
                    x, y = chart_positions[chart_idx % 2]
                    pdf.image(chart_path, x=x, y=y, w=85)
                    chart_idx += 1
                except:
                    pass


def generate_charts_report(file_structure: dict, output_dir: Path) -> str:
    """
    Simple function to generate a charts-only PDF report
    
    Args:
        file_structure: Dictionary with file and sheet data
        output_dir: Directory to save the PDF
    
    Returns:
        Path to generated PDF file
    """
    generator = SimpleChartsReport(output_dir)
    return generator.generate_charts_only_report(file_structure)