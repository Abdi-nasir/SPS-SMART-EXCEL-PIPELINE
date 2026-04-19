"""
Excel File Analyzer - Understands the structure and content of Excel files
Supports multi-sheet Excel files with full sheet-by-sheet analysis
"""
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional

class ExcelFileAnalyzer:
    """Analyzes Excel files to understand their structure and content across all sheets"""
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.analysis = {}
        
    def analyze(self) -> Dict[str, Any]:
        """Perform complete analysis of the Excel file, including all sheets"""
        
        try:
            # Read ALL sheets at once using sheet_name=None
            all_sheets = pd.read_excel(self.file_path, sheet_name=None)
        except Exception as e:
            # Fallback: try reading only the first sheet
            print(f"Warning: Could not read all sheets from {self.file_path.name}: {e}")
            df = pd.read_excel(self.file_path)
            all_sheets = {'Sheet1': df}
        
        # Basic file information
        self.analysis = {
            'filename': self.file_path.name,
            'file_size_kb': self.file_path.stat().st_size / 1024,
            'sheet_count': len(all_sheets),
            'sheet_names': list(all_sheets.keys()),
            'sheets': {},
            'has_multiple_sheets': len(all_sheets) > 1
        }
        
        # Analyze each sheet individually
        for sheet_name, df in all_sheets.items():
            sheet_analysis = self._analyze_sheet(df, sheet_name)
            self.analysis['sheets'][sheet_name] = sheet_analysis
        
        # Determine overall file content type
        content_types = [s['content_type'] for s in self.analysis['sheets'].values()]
        self.analysis['content_type'] = max(set(content_types), key=content_types.count) if content_types else 'General Data'
        
        # Overall quality score
        quality_scores = [s['quality_score'] for s in self.analysis['sheets'].values()]
        self.analysis['quality_score'] = sum(quality_scores) / len(quality_scores) if quality_scores else 0
        
        return self.analysis
    
    def _analyze_sheet(self, df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Analyze a single sheet within an Excel file"""
        
        if df.empty:
            return {
                'sheet_name': sheet_name,
                'shape': (0, 0),
                'rows': 0,
                'columns': 0,
                'column_names': [],
                'column_types': {},
                'is_empty': True,
                'has_numeric_data': False,
                'has_dates': False
            }
        
        sheet_analysis = {
            'sheet_name': sheet_name,
            'shape': df.shape,
            'rows': df.shape[0],
            'columns': df.shape[1],
            'column_names': list(df.columns),
            'column_types': {},
            'null_counts': df.isnull().sum().to_dict(),
            'null_percentages': (df.isnull().sum() / len(df) * 100).to_dict(),
            'unique_counts': df.nunique().to_dict(),
            'has_numeric_data': False,
            'has_dates': False,
            'has_duplicates': df.duplicated().any(),
            'duplicate_count': df.duplicated().sum(),
            'sample_data': df.head(5).to_dict('records') if len(df) > 0 else [],
            'is_empty': False
        }
        
        # Analyze each column's data type
        for col, dtype in df.dtypes.items():
            sheet_analysis['column_types'][col] = str(dtype)
            if 'int' in str(dtype) or 'float' in str(dtype):
                sheet_analysis['has_numeric_data'] = True
            if 'datetime' in str(dtype):
                sheet_analysis['has_dates'] = True
        
        # Detect content type
        sheet_analysis['content_type'] = self._detect_content_type(df)
        sheet_analysis['quality_score'] = self._calculate_quality_score(df)
        
        return sheet_analysis
    
    def _detect_content_type(self, df: pd.DataFrame) -> str:
        """Detect what type of data this sheet contains"""
        sales_keywords = ['sales', 'revenue', 'transaction', 'order', 'invoice', 'sale', 'sell']
        hr_keywords = ['employee', 'salary', 'hire', 'department', 'manager', 'staff', 'person']
        inventory_keywords = ['stock', 'inventory', 'product', 'sku', 'quantity', 'item']
        financial_keywords = ['balance', 'account', 'payment', 'budget', 'expense', 'cost', 'profit']
        
        all_columns = ' '.join(df.columns.str.lower())
        
        if any(keyword in all_columns for keyword in sales_keywords):
            return 'Sales Data'
        elif any(keyword in all_columns for keyword in hr_keywords):
            return 'HR Data'
        elif any(keyword in all_columns for keyword in inventory_keywords):
            return 'Inventory Data'
        elif any(keyword in all_columns for keyword in financial_keywords):
            return 'Financial Data'
        else:
            return 'General Data'
    
    def _calculate_quality_score(self, df: pd.DataFrame) -> float:
        """Calculate a data quality score from 0-100 for a sheet"""
        if df.empty:
            return 0
            
        score = 100
        null_percentage = df.isnull().sum().sum() / (df.shape[0] * df.shape[1]) * 100
        score -= null_percentage * 0.5
        
        if df.shape[0] > 0:
            duplicate_percentage = (df.duplicated().sum() / len(df)) * 100
            score -= duplicate_percentage
        
        if df.shape[1] > 20:
            score -= (df.shape[1] - 20) * 0.5
        
        if df.shape[0] < 10:
            score -= (10 - df.shape[0]) * 2
        
        return max(0, min(100, score))


def compare_files(analyses: List[Dict]) -> Dict[str, Any]:
    """Compare multiple file analyses to determine compatibility"""
    if len(analyses) <= 1:
        return {'should_combine': False, 'reason': 'Only one file provided', 'confidence': 100}
    
    reference = analyses[0]
    reference_sheets = reference.get('sheets', {})
    reference_is_multi = reference.get('has_multiple_sheets', False)
    
    all_compatible = True
    differences = []
    
    for i, analysis in enumerate(analyses[1:], start=1):
        current_sheets = analysis.get('sheets', {})
        current_is_multi = analysis.get('has_multiple_sheets', False)
        
        if reference_is_multi != current_is_multi:
            all_compatible = False
            differences.append({
                'file': analysis['filename'],
                'issue': 'Sheet count mismatch',
                'reference_sheets': list(reference_sheets.keys()),
                'current_sheets': list(current_sheets.keys())
            })
            continue
        
        ref_sheet_names = set(reference_sheets.keys())
        cur_sheet_names = set(current_sheets.keys())
        
        if ref_sheet_names != cur_sheet_names:
            all_compatible = False
            differences.append({
                'file': analysis['filename'],
                'issue': 'Sheet name mismatch',
                'missing_sheets': list(ref_sheet_names - cur_sheet_names),
                'extra_sheets': list(cur_sheet_names - ref_sheet_names)
            })
            continue
        
        for sheet_name in ref_sheet_names:
            ref_columns = set(reference_sheets[sheet_name].get('column_names', []))
            cur_columns = set(current_sheets[sheet_name].get('column_names', []))
            
            if ref_columns != cur_columns:
                all_compatible = False
                differences.append({
                    'file': analysis['filename'],
                    'sheet': sheet_name,
                    'issue': 'Column mismatch',
                    'missing_columns': list(ref_columns - cur_columns),
                    'extra_columns': list(cur_columns - ref_columns)
                })
    
    if all_compatible:
        return {
            'should_combine': True,
            'reason': 'All files have identical sheet and column structures',
            'file_count': len(analyses),
            'total_rows': sum(a.get('sheets', {}).get(list(a.get('sheets', {}).keys())[0], {}).get('rows', 0) 
                             if a.get('sheets') else 0 for a in analyses),
            'confidence': 100
        }
    else:
        return {
            'should_combine': False,
            'reason': 'Files have incompatible structures',
            'differences': differences,
            'file_count': len(analyses),
            'confidence': 100
        }