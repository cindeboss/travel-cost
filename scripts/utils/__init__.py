# Utils package for Travel Cost Analysis
from .file_scanner import (
    scan_excel_files,
    scan_and_classify_files,
    print_scan_summary,
    get_files_to_process,
    update_processed_metadata,
    ScanResult,
    TravelFileInfo,
    RosterFileInfo
)

from .date_matcher import (
    parse_date_range_from_filename,
    extract_roster_month,
    find_matching_roster_file,
    DateRange,
    extract_date_from_string
)

# Import load_employee_index from process_roster
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from process_roster import load_employee_index

__all__ = [
    'scan_excel_files',
    'scan_and_classify_files',
    'print_scan_summary',
    'get_files_to_process',
    'update_processed_metadata',
    'ScanResult',
    'TravelFileInfo',
    'RosterFileInfo',
    'parse_date_range_from_filename',
    'extract_roster_month',
    'find_matching_roster_file',
    'DateRange',
    'extract_date_from_string',
    'load_employee_index'
]
