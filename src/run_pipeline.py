"""
Main Pipeline Runner - Execute the entire data processing pipeline
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from orchestrator import ExcelOrchestrator

def main():
    PROJECT_ROOT = Path(__file__).parent.parent
    RAW_DIR = PROJECT_ROOT / "data" / "raw"
    PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
    
    print("\n" + "=" * 30)
    print("STARTING SMART EXCEL DATA PIPELINE")
    print("=" * 30)
    
    if not RAW_DIR.exists():
        RAW_DIR.mkdir(parents=True)
        print(f"\n Created {RAW_DIR}")
        print(" Please add your Excel files to this folder and run again!")
        return
    
    excel_files = list(RAW_DIR.glob("*.xlsx")) + list(RAW_DIR.glob("*.xls"))
    if not excel_files:
        print(f"\n No Excel files found in {RAW_DIR}")
        print(" Please add .xlsx or .xls files to this folder")
        return
    
    print(f"\n Found {len(excel_files)} Excel file(s):")
    for f in excel_files:
        print(f"   - {f.name}")
    
    orchestrator = ExcelOrchestrator(RAW_DIR, PROCESSED_DIR)
    results = orchestrator.process_all()
    
    print("\n" + "="*60)
    print(" PIPELINE EXECUTION COMPLETE!")
    print("="*60)
    print("\n To launch the dashboard, run:")
    print("   streamlit run src/app.py")

if __name__ == "__main__":
    main()