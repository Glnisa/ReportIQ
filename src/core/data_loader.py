"""
Data Loader Module for ReportIQ
Handles Excel file loading and column mapping
"""

import pandas as pd
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime


class DataLoader:
    """Loads and parses Excel vulnerability data"""
    
    # Expected column mappings (internal name -> possible Excel names)
    COLUMN_MAPPINGS = {
        "ticket_id": ["TICKETID", "TICKET_ID", "Ticket ID", "ticket_id", "ID"],
        "priority": ["REPORTEDPRIORITY", "PRIORITY", "Priority", "priority", "Severity"],
        "sla_status": ["SLA_Value", "SLA_STATUS", "SLA Status", "sla_status", "SLA"],
        "status": ["STATUS", "Status", "status", "State"],
        "creation_date": ["Day_of_CREATIONDATE", "CREATIONDATE", "Creation Date", "creation_date", "Created"],
        "closed_date": ["Day_of_CLOSEDDATE", "CLOSEDDATE", "Closed Date", "closed_date", "Closed"],
        "status_date": ["Day_of_STATUS/DATE", "STATUS_DATE", "Status Date", "status_date"],
        "department": ["Department", "DEPARTMENT", "department", "Dept"],
        "line_manager": ["Line Manager (group)", "LINE_MANAGER", "Line Manager", "line_manager", "Manager"],
        "assigned_owner": ["ASSIGNEDOWNER/GROUP", "ASSIGNED_OWNER", "Assigned Owner", "assigned_owner", "Owner"],
        "new_assigned_owner": ["NEW_ASSIGNEDOWNER/GROUP", "NEW_ASSIGNED_OWNER", "New Assigned Owner"],
        "plugin_id": ["PLUGINID", "PLUGIN_ID", "Plugin ID", "plugin_id", "PluginID"],
        "plugin_desc": ["PLUGINDESC", "PLUGIN_DESC", "Plugin Description", "plugin_desc", "Vulnerability"],
        "tool": ["TOOL", "Tool", "tool", "Source", "Scanner"],
        "port": ["PORT", "Port", "port"],
        "ip": ["IP", "ip", "IP Address", "Host"],
        "dns": ["DNS", "dns", "Hostname", "FQDN"],
        "sla_time": ["Negative_Value", "SLA_TIME", "SLA Time", "sla_time", "Days Overdue"],
    }
    
    # Status categories
    OPEN_STATUSES = ["PENDING", "QUEUED", "QUEUEDR", "WRISKACCPT"]
    CLOSED_STATUSES = ["CLOSED", "CANCEL", "RISKACCPT", "CANCELLED"]
    
    def __init__(self):
        self.df: Optional[pd.DataFrame] = None
        self.original_df: Optional[pd.DataFrame] = None
        self.file_path: Optional[Path] = None
        self.column_map: Dict[str, str] = {}
        self._unique_values_cache: Dict[str, List[str]] = {}
    
    def load_file(self, file_path: str) -> Tuple[bool, str]:
        """
        Load an Excel file and attempt automatic column mapping.
        Returns (success, message)
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return False, f"File not found: {file_path}"
            
            # Read Excel file
            self.df = pd.read_excel(file_path, engine='openpyxl')
            self.original_df = self.df.copy()
            self.file_path = path
            
            # Attempt automatic column mapping
            self._auto_map_columns()
            
            # Parse dates
            self._parse_dates()
            
            # Clear cache
            self._unique_values_cache.clear()
            
            return True, f"Loaded {len(self.df)} records from {path.name}"
            
        except Exception as e:
            return False, f"Error loading file: {str(e)}"
    
    def clear(self) -> None:
        """Clear the current data and state"""
        self.df = None
        self.original_df = None
        self.file_path = None
        self.column_map = {}
        self._unique_values_cache.clear()
    
    def _auto_map_columns(self) -> None:
        """Automatically map Excel columns to internal names"""
        excel_columns = list(self.df.columns)
        self.column_map = {}
        
        for internal_name, possible_names in self.COLUMN_MAPPINGS.items():
            for excel_col in excel_columns:
                # Case-insensitive comparison, strip whitespace
                excel_col_clean = str(excel_col).strip()
                for possible in possible_names:
                    if excel_col_clean.lower() == possible.lower():
                        self.column_map[internal_name] = excel_col
                        break
                if internal_name in self.column_map:
                    break
    
    def _parse_dates(self) -> None:
        """Parse date columns to datetime"""
        date_columns = ["creation_date", "closed_date", "status_date"]
        
        for col in date_columns:
            if col in self.column_map and self.column_map[col] in self.df.columns:
                excel_col = self.column_map[col]
                try:
                    self.df[excel_col] = pd.to_datetime(
                        self.df[excel_col], 
                        dayfirst=True,  # DD/MM/YYYY format
                        errors='coerce'
                    )
                except Exception:
                    pass  # Keep original if parsing fails
    
    def get_column(self, internal_name: str) -> Optional[str]:
        """Get the Excel column name for an internal name"""
        return self.column_map.get(internal_name)
    
    def get_column_data(self, internal_name: str) -> Optional[pd.Series]:
        """Get the data for a column by internal name"""
        excel_col = self.get_column(internal_name)
        if excel_col and excel_col in self.df.columns:
            return self.df[excel_col]
        return None
    
    def get_unique_values(self, internal_name: str) -> List[str]:
        """Get unique values for a column (cached)"""
        if internal_name in self._unique_values_cache:
            return self._unique_values_cache[internal_name]
        
        data = self.get_column_data(internal_name)
        if data is not None:
            unique_vals = data.dropna().unique().tolist()
            # Convert to strings and sort
            unique_vals = sorted([str(v) for v in unique_vals])
            self._unique_values_cache[internal_name] = unique_vals
            return unique_vals
        return []
    
    def get_years(self) -> List[int]:
        """Get unique years from creation_date"""
        data = self.get_column_data("creation_date")
        if data is not None:
            try:
                years = data.dropna().dt.year.unique().tolist()
                return sorted([int(y) for y in years if pd.notna(y)])
            except Exception:
                pass
        return []
    
    def get_tools(self) -> List[str]:
        """Get unique tool values"""
        return self.get_unique_values("tool")
    
    def get_departments(self) -> List[str]:
        """Get unique department values"""
        return self.get_unique_values("department")
    
    def get_line_managers(self) -> List[str]:
        """Get unique line manager values"""
        return self.get_unique_values("line_manager")
    
    def get_statuses(self) -> List[str]:
        """Get unique status values"""
        return self.get_unique_values("status")
    
    def get_priorities(self) -> List[str]:
        """Get unique priority values"""
        return self.get_unique_values("priority")
    
    def get_sla_statuses(self) -> List[str]:
        """Get unique SLA status values"""
        return self.get_unique_values("sla_status")
    
    def get_mapped_columns(self) -> Dict[str, str]:
        """Get the current column mapping"""
        return self.column_map.copy()
    
    def get_unmapped_columns(self) -> List[str]:
        """Get list of internal columns that weren't mapped"""
        return [k for k in self.COLUMN_MAPPINGS.keys() if k not in self.column_map]
    
    def get_excel_columns(self) -> List[str]:
        """Get list of all Excel column names"""
        if self.df is not None:
            return list(self.df.columns)
        return []
    
    def set_column_mapping(self, internal_name: str, excel_column: str) -> None:
        """Manually set a column mapping"""
        if self.df is not None and excel_column in self.df.columns:
            self.column_map[internal_name] = excel_column
            self._unique_values_cache.pop(internal_name, None)
    
    def get_record_count(self) -> int:
        """Get total number of records"""
        if self.df is not None:
            return len(self.df)
        return 0
    
    def get_dataframe(self) -> Optional[pd.DataFrame]:
        """Get the current dataframe"""
        return self.df
    
    def get_preview_data(self, rows: int = 10) -> List[Dict[str, Any]]:
        """Get preview data as list of dicts"""
        if self.df is not None:
            return self.df.head(rows).to_dict('records')
        return []
    
    def is_loaded(self) -> bool:
        """Check if data is loaded"""
        return self.df is not None and len(self.df) > 0
    
    def get_open_status_values(self) -> List[str]:
        """Get list of status values that indicate open tickets"""
        all_statuses = self.get_statuses()
        return [s for s in all_statuses if s.upper() in [os.upper() for os in self.OPEN_STATUSES]]
    
    def get_closed_status_values(self) -> List[str]:
        """Get list of status values that indicate closed tickets"""
        all_statuses = self.get_statuses()
        return [s for s in all_statuses if s.upper() in [os.upper() for os in self.CLOSED_STATUSES]]
