"""
Multi-Sheet Excel Handler - Processes Excel files with multiple sheets intelligently
"""
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

class MultiSheetHandler:
    """Handles Excel files with multiple sheets intelligently"""
    
    def __init__(self, strategy: str = 'auto'):
        self.strategy = strategy
    
    def process_file(self, file_path: Path) -> Dict[str, Any]:
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None)
        except Exception as e:
            print(f"Warning: Could not read all sheets from {file_path.name}: {e}")
            df = pd.read_excel(file_path)
            all_sheets = {'Sheet1': df}
        
        result = {
            'filename': file_path.name,
            'sheet_count': len(all_sheets),
            'sheets': {},
            'combined_df': None,
            'processing_mode': None,
            'sheet_combination_map': {}
        }
        
        for sheet_name, df in all_sheets.items():
            if df.empty:
                print(f"  Skipping empty sheet: {sheet_name}")
                continue
                
            df = df.drop_duplicates()
            df = df.dropna(how='all')
            df['_source_file'] = file_path.name
            df['_source_sheet'] = sheet_name
            result['sheets'][sheet_name] = df
        
        if not result['sheets']:
            result['processing_mode'] = 'empty'
            return result
        
        if self.strategy == 'combine_all':
            result['combined_df'] = self._combine_sheets_vertically(result['sheets'])
            result['processing_mode'] = 'combined'
        elif self.strategy == 'keep_separate':
            result['processing_mode'] = 'separate'
        elif self.strategy == 'combine_similar':
            result = self._combine_similar_sheets(result)
        else:
            result = self._auto_decide(result)
        
        return result
    
    def _combine_sheets_vertically(self, sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
        if len(sheets) == 1:
            return list(sheets.values())[0]
        
        all_columns = set()
        for df in sheets.values():
            all_columns.update(df.columns)
        
        aligned_dfs = []
        for sheet_name, df in sheets.items():
            for col in all_columns:
                if col not in df.columns:
                    df[col] = None
            aligned_dfs.append(df)
        
        return pd.concat(aligned_dfs, ignore_index=True)
    
    def _auto_decide(self, result: Dict[str, Any]) -> Dict[str, Any]:
        sheets = result['sheets']
        
        if len(sheets) <= 1:
            result['processing_mode'] = 'single'
            return result
        
        sheet_items = list(sheets.items())
        first_sheet_df = sheet_items[0][1]
        first_sheet_columns = set(first_sheet_df.columns)
        metadata_cols = {'_source_file', '_source_sheet'}
        first_sheet_columns = first_sheet_columns - metadata_cols
        
        all_compatible = True
        compatible_sheets = []
        incompatible_sheets = []
        
        for sheet_name, df in sheet_items:
            current_columns = set(df.columns) - metadata_cols
            if current_columns == first_sheet_columns:
                compatible_sheets.append(sheet_name)
            else:
                incompatible_sheets.append(sheet_name)
                all_compatible = False
        
        if all_compatible:
            result['combined_df'] = self._combine_sheets_vertically(sheets)
            result['processing_mode'] = 'combined'
        elif len(compatible_sheets) > 1:
            compatible_sheets_dict = {name: sheets[name] for name in compatible_sheets}
            result['combined_df'] = self._combine_sheets_vertically(compatible_sheets_dict)
            result['processing_mode'] = 'mixed'
            result['separate_sheets'] = {name: sheets[name] for name in incompatible_sheets}
        else:
            result['processing_mode'] = 'separate'
        
        return result
    
    def _combine_similar_sheets(self, result: Dict[str, Any]) -> Dict[str, Any]:
        sheets = result['sheets']
        
        if len(sheets) <= 1:
            result['processing_mode'] = 'single'
            return result
        
        metadata_cols = {'_source_file', '_source_sheet'}
        groups = {}
        
        for sheet_name, df in sheets.items():
            columns = tuple(sorted([c for c in df.columns if c not in metadata_cols]))
            if columns not in groups:
                groups[columns] = []
            groups[columns].append((sheet_name, df))
        
        if len(groups) == 1:
            result['combined_df'] = self._combine_sheets_vertically(sheets)
            result['processing_mode'] = 'combined'
        else:
            combined_groups = {}
            for cols, group_sheets in groups.items():
                group_dict = {name: df for name, df in group_sheets}
                if len(group_sheets) > 1:
                    combined_groups[f"combined_{cols[:20]}"] = self._combine_sheets_vertically(group_dict)
                else:
                    combined_groups[group_sheets[0][0]] = group_sheets[0][1]
            
            result['sheets'] = combined_groups
            result['processing_mode'] = 'grouped'
        
        return result