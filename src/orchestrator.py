"""
Orchestrator - Makes decisions on combining vs separating Excel files
"""
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any
from file_analyzer import ExcelFileAnalyzer, compare_files
from multi_sheet_handler import MultiSheetHandler

class ExcelOrchestrator:
    """Orchestrates the processing of Excel files with intelligent grouping"""
    
    def __init__(self, raw_dir: str, processed_dir: str):
        self.raw_dir = Path(raw_dir)
        self.processed_dir = Path(processed_dir)
        self.processed_dir.mkdir(parents=True, exist_ok=True)
        self.sheet_handler = MultiSheetHandler(strategy='auto')
        
    def process_all(self) -> Dict[str, Any]:
        print("\n" + "="*60)
        print(" SMART EXCEL PIPELINE ORCHESTRATOR (Multi-Sheet Support)")
        print("="*60)
        
        excel_files = list(self.raw_dir.glob("*.xlsx")) + list(self.raw_dir.glob("*.xls"))
        
        if not excel_files:
            print(f"\n No Excel files found in {self.raw_dir}")
            return {'error': 'No files found'}
        
        print(f"\n Found {len(excel_files)} Excel file(s):")
        for f in excel_files:
            print(f"   - {f.name}")
        
        print("\n Analyzing each file and its sheets...")
        analyses = []
        
        for file_path in excel_files:
            analyzer = ExcelFileAnalyzer(file_path)
            analysis = analyzer.analyze()
            analyses.append(analysis)
            
            sheet_count = analysis.get('sheet_count', 1)
            sheet_names = analysis.get('sheet_names', ['Sheet1'])
            print(f"\n    {file_path.name}:")
            print(f"      Sheets: {sheet_count} ({', '.join(sheet_names[:3])}{'...' if sheet_count > 3 else ''})")
            
            for sheet_name, sheet_info in analysis.get('sheets', {}).items():
                print(f"         └─ {sheet_name}: {sheet_info['rows']} rows, {sheet_info['columns']} cols")
        
        print("\n Grouping files by compatibility...")
        groups = self._group_by_compatibility(analyses)
        print(f"   Created {len(groups)} group(s)")
        
        results = {}
        all_processed_data = []
        
        for group_id, group_analyses in groups.items():
            decision = compare_files(group_analyses)
            
            if decision['should_combine']:
                print(f"\n Group '{group_id}': COMBINING {len(group_analyses)} file(s)")
                print(f"   Reason: {decision['reason']}")
                result = self._combine_files(group_analyses, group_id)
                results[group_id] = result
                if result.get('output_file'):
                    all_processed_data.append(result)
            else:
                print(f"\n Group '{group_id}': KEEPING SEPARATE ({len(group_analyses)} file(s))")
                print(f"   Reason: {decision['reason']}")
                for analysis in group_analyses:
                    result = self._process_single_file(analysis)
                    results[analysis['filename']] = result
                    if result.get('output_file'):
                        all_processed_data.append(result)
        
        self._save_summary(results, analyses)
        
        if len(all_processed_data) > 1:
            self._create_master_view(all_processed_data)
        
        return results
    
    def _group_by_compatibility(self, analyses: List[Dict]) -> Dict[str, List[Dict]]:
        groups = {}
        for analysis in analyses:
            sheets = analysis.get('sheets', {})
            sheet_count = len(sheets)
            sheet_names_tuple = tuple(sorted(sheets.keys()))
            
            column_signatures = []
            for sheet_name, sheet_info in sheets.items():
                cols = tuple(sorted(sheet_info.get('column_names', [])))
                column_signatures.append(f"{sheet_name}:{cols}")
            
            columns_key = tuple(sorted(column_signatures))
            content_key = analysis.get('content_type', 'General')
            group_key = f"{content_key}_{sheet_count}_{columns_key}"
            
            if group_key not in groups:
                groups[group_key] = []
            groups[group_key].append(analysis)
        
        return groups
    
    def _combine_files(self, analyses: List[Dict], group_id: str) -> Dict[str, Any]:
        all_data = []
        sheet_handler = MultiSheetHandler(strategy='auto')
        
        for analysis in analyses:
            file_path = self.raw_dir / analysis['filename']
            processed = sheet_handler.process_file(file_path)
            
            if processed['processing_mode'] == 'combined':
                df = processed['combined_df']
                all_data.append(df)
            elif processed['processing_mode'] == 'separate':
                for sheet_name, sheet_df in processed['sheets'].items():
                    all_data.append(sheet_df)
            elif processed['processing_mode'] == 'mixed':
                if processed.get('combined_df') is not None:
                    all_data.append(processed['combined_df'])
                if processed.get('separate_sheets'):
                    for sheet_name, sheet_df in processed['separate_sheets'].items():
                        all_data.append(sheet_df)
            elif processed['processing_mode'] == 'grouped':
                for sheet_name, sheet_df in processed['sheets'].items():
                    all_data.append(sheet_df)
            else:
                if processed['sheets']:
                    df = list(processed['sheets'].values())[0]
                    all_data.append(df)
        
        if not all_data:
            return {'type': 'combined', 'group_id': group_id, 'files': [a['filename'] for a in analyses], 'error': 'No data to combine'}
        
        try:
            combined_df = pd.concat(all_data, ignore_index=True, sort=False)
            combined_df = combined_df.fillna('')
            safe_group_id = group_id.replace(' ', '_').replace('/', '_').replace('\\', '_')
            output_name = f"combined_{safe_group_id[:50]}.csv"
            output_path = self.processed_dir / output_name
            combined_df.to_csv(output_path, index=False)
            
            print(f"      Saved: {output_name} ({len(combined_df)} rows, {len(combined_df.columns)} cols)")
            
            return {
                'type': 'combined',
                'group_id': group_id,
                'files': [a['filename'] for a in analyses],
                'output_file': str(output_path),
                'total_rows': len(combined_df),
                'total_columns': len(combined_df.columns),
                'content_types': list(set(a.get('content_type', 'General') for a in analyses)),
                'dashboard_mode': 'unified_with_filters'
            }
        except Exception as e:
            return {'type': 'combined', 'group_id': group_id, 'files': [a['filename'] for a in analyses], 'error': str(e)}
    
    def _process_single_file(self, analysis: Dict) -> Dict[str, Any]:
        file_path = self.raw_dir / analysis['filename']
        processed = self.sheet_handler.process_file(file_path)
        
        if processed['processing_mode'] == 'combined':
            output_name = f"independent_{analysis['filename'].replace('.xlsx', '.csv').replace('.xls', '.csv')}"
            output_path = self.processed_dir / output_name
            processed['combined_df'].to_csv(output_path, index=False)
            
            return {
                'type': 'independent',
                'files': [analysis['filename']],
                'output_file': str(output_path),
                'total_rows': len(processed['combined_df']),
                'total_columns': len(processed['combined_df'].columns),
                'content_type': analysis.get('content_type', 'General'),
                'quality_score': analysis.get('quality_score', 0),
                'dashboard_mode': 'single_dashboard',
                'sheet_count': processed['sheet_count'],
                'processing_mode': processed['processing_mode']
            }
        
        elif processed['processing_mode'] == 'separate':
            output_files = []
            for sheet_name, sheet_df in processed['sheets'].items():
                safe_name = f"{analysis['filename'].replace('.xlsx', '').replace('.xls', '')}_{sheet_name}.csv"
                output_path = self.processed_dir / safe_name
                sheet_df.to_csv(output_path, index=False)
                output_files.append(str(output_path))
            
            return {
                'type': 'independent_multi_sheet',
                'files': [analysis['filename']],
                'output_files': output_files,
                'total_rows': sum(len(df) for df in processed['sheets'].values()),
                'sheet_count': len(processed['sheets']),
                'content_type': analysis.get('content_type', 'General'),
                'dashboard_mode': 'multi_sheet_selector',
                'sheets': {name: str(self.processed_dir / f"{analysis['filename'].replace('.xlsx', '').replace('.xls', '')}_{name}.csv") 
                          for name in processed['sheets'].keys()}
            }
        
        else:
            if processed['sheets']:
                sheet_df = list(processed['sheets'].values())[0]
                output_name = f"independent_{analysis['filename'].replace('.xlsx', '.csv').replace('.xls', '.csv')}"
                output_path = self.processed_dir / output_name
                sheet_df.to_csv(output_path, index=False)
                
                return {
                    'type': 'independent',
                    'files': [analysis['filename']],
                    'output_file': str(output_path),
                    'total_rows': len(sheet_df),
                    'total_columns': len(sheet_df.columns),
                    'content_type': analysis.get('content_type', 'General'),
                    'quality_score': analysis.get('quality_score', 0),
                    'dashboard_mode': 'single_dashboard'
                }
        
        return {'type': 'error', 'files': [analysis['filename']], 'error': 'Could not process file'}
    
    def _save_summary(self, results: Dict, all_analyses: List[Dict]):
        summary_data = []
        
        for key, result in results.items():
            if result.get('type') == 'independent_multi_sheet':
                summary_data.append({
                    'dashboard_id': key,
                    'display_name': key.replace('_', ' ').title().replace('.Xlsx', '').replace('.Xls', ''),
                    'type': 'multi_sheet',
                    'files': ', '.join(result.get('files', [])),
                    'file_count': len(result.get('files', [])),
                    'total_rows': result.get('total_rows', 0),
                    'sheet_count': result.get('sheet_count', 0),
                    'dashboard_mode': result.get('dashboard_mode', 'multi_sheet_selector'),
                    'output_file': result.get('output_files', [''])[0] if result.get('output_files') else '',
                    'all_outputs': '|'.join(result.get('output_files', [])),
                    'sheets': '|'.join(result.get('sheets', {}).keys())
                })
            else:
                summary_data.append({
                    'dashboard_id': key,
                    'display_name': key.replace('_', ' ').title().replace('.Xlsx', '').replace('.Xls', ''),
                    'type': result.get('type', 'unknown'),
                    'files': ', '.join(result.get('files', [])),
                    'file_count': len(result.get('files', [])),
                    'total_rows': result.get('total_rows', 0),
                    'dashboard_mode': result.get('dashboard_mode', 'standard'),
                    'output_file': result.get('output_file', ''),
                    'sheet_count': result.get('sheet_count', 1)
                })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_csv(self.processed_dir / 'dashboard_summary.csv', index=False)
        
        detailed_data = []
        for analysis in all_analyses:
            for sheet_name, sheet_info in analysis.get('sheets', {}).items():
                detailed_data.append({
                    'filename': analysis['filename'],
                    'sheet_name': sheet_name,
                    'rows': sheet_info.get('rows', 0),
                    'columns': sheet_info.get('columns', 0),
                    'content_type': sheet_info.get('content_type', 'Unknown'),
                    'quality_score': sheet_info.get('quality_score', 0),
                    'has_numeric': sheet_info.get('has_numeric_data', False),
                    'has_dates': sheet_info.get('has_dates', False)
                })
        
        if detailed_data:
            detailed_df = pd.DataFrame(detailed_data)
            detailed_df.to_csv(self.processed_dir / 'file_analysis_detailed.csv', index=False)
        
        print(f"\n Summary saved to {self.processed_dir}/dashboard_summary.csv")
        print(f" Detailed analysis saved to {self.processed_dir}/file_analysis_detailed.csv")
    
    def _create_master_view(self, all_processed_data: List[Dict]):
        try:
            master_data = []
            for data in all_processed_data:
                if data.get('output_file') and Path(data['output_file']).exists():
                    df = pd.read_csv(data['output_file'])
                    master_data.append(df)
            
            if master_data and len(master_data) > 1:
                first_cols = set(master_data[0].columns)
                compatible = all(set(df.columns) == first_cols for df in master_data)
                
                if compatible:
                    master_df = pd.concat(master_data, ignore_index=True)
                    master_df.to_csv(self.processed_dir / 'master_view.csv', index=False)
                    print(f" Master view created: master_view.csv ({len(master_df)} rows)")
        except Exception as e:
            print(f"Note: Could not create master view: {e}")