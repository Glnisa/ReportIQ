"""
Filter Engine Module for ReportIQ
Handles data filtering based on user selections
"""

import pandas as pd
from typing import Optional, List, Dict, Any
from datetime import datetime


class FilterEngine:
    """Applies filters to vulnerability data"""
    
    def __init__(self, data_loader):
        """
        Initialize with a DataLoader instance
        
        Args:
            data_loader: DataLoader instance with loaded data
        """
        self.data_loader = data_loader
        self.filters: Dict[str, Any] = {}
        self._filtered_df: Optional[pd.DataFrame] = None
    
    def set_filter(self, filter_name: str, value: Any) -> None:
        """
        Set a filter value
        
        Args:
            filter_name: Name of the filter (sla_status, status, tool, year, etc.)
            value: Filter value (single value or list for multi-select)
        """
        if value is None or (isinstance(value, list) and len(value) == 0):
            # Remove filter if empty
            self.filters.pop(filter_name, None)
        else:
            self.filters[filter_name] = value
        
        # Invalidate cache
        self._filtered_df = None
    
    def clear_filters(self) -> None:
        """Clear all filters"""
        self.filters.clear()
        self._filtered_df = None
    
    def get_filters(self) -> Dict[str, Any]:
        """Get current filter settings"""
        return self.filters.copy()
    
    def apply_filters(self) -> pd.DataFrame:
        """
        Apply all current filters and return filtered DataFrame
        
        Returns:
            Filtered pandas DataFrame
        """
        if self._filtered_df is not None:
            return self._filtered_df
        
        df = self.data_loader.get_dataframe()
        if df is None:
            return pd.DataFrame()
        
        # Start with a copy
        filtered = df.copy()
        
        # Apply each filter
        for filter_name, value in self.filters.items():
            filtered = self._apply_single_filter(filtered, filter_name, value)
        
        self._filtered_df = filtered
        return filtered
    
    def _apply_single_filter(self, df: pd.DataFrame, filter_name: str, value: Any) -> pd.DataFrame:
        """Apply a single filter to the dataframe"""
        
        # Get the Excel column name
        excel_col = self.data_loader.get_column(filter_name)
        
        if filter_name == "year":
            return self._filter_by_year(df, value)
        elif filter_name == "year_range":
            return self._filter_by_year_range(df, value)
        elif filter_name == "open_only":
            return self._filter_open_tickets(df, value)
        elif filter_name == "out_of_sla_only":
            return self._filter_out_of_sla(df, value)
        elif excel_col and excel_col in df.columns:
            return self._filter_by_column(df, excel_col, value)
        
        return df
    
    def _filter_by_column(self, df: pd.DataFrame, column: str, value: Any) -> pd.DataFrame:
        """Filter by column value(s)"""
        if isinstance(value, list):
            # Multi-select filter
            return df[df[column].astype(str).isin([str(v) for v in value])]
        else:
            # Single value filter
            return df[df[column].astype(str) == str(value)]
    
    def _filter_by_year(self, df: pd.DataFrame, years: Any) -> pd.DataFrame:
        """Filter by creation year"""
        creation_col = self.data_loader.get_column("creation_date")
        if not creation_col or creation_col not in df.columns:
            return df
        
        if not isinstance(years, list):
            years = [years]
        
        try:
            # Ensure dates are parsed
            date_series = pd.to_datetime(df[creation_col], errors='coerce')
            mask = date_series.dt.year.isin(years)
            return df[mask]
        except Exception:
            return df
    
    def _filter_by_year_range(self, df: pd.DataFrame, year_range: tuple) -> pd.DataFrame:
        """Filter by year range (start_year, end_year)"""
        creation_col = self.data_loader.get_column("creation_date")
        if not creation_col or creation_col not in df.columns:
            return df
        
        start_year, end_year = year_range
        
        try:
            date_series = pd.to_datetime(df[creation_col], errors='coerce')
            mask = (date_series.dt.year >= start_year) & (date_series.dt.year <= end_year)
            return df[mask]
        except Exception:
            return df
    
    def _filter_open_tickets(self, df: pd.DataFrame, open_only: bool) -> pd.DataFrame:
        """Filter to show only open tickets"""
        if not open_only:
            return df
        
        status_col = self.data_loader.get_column("status")
        if not status_col or status_col not in df.columns:
            return df
        
        open_statuses = [s.upper() for s in self.data_loader.OPEN_STATUSES]
        mask = df[status_col].astype(str).str.upper().isin(open_statuses)
        return df[mask]
    
    def _filter_out_of_sla(self, df: pd.DataFrame, out_of_sla_only: bool) -> pd.DataFrame:
        """Filter to show only out of SLA tickets"""
        if not out_of_sla_only:
            return df
        
        sla_col = self.data_loader.get_column("sla_status")
        if not sla_col or sla_col not in df.columns:
            return df
        
        # Check for various "out of SLA" values
        out_of_sla_values = ["out of sla", "out_of_sla", "outlier", "breached", "overdue"]
        mask = df[sla_col].astype(str).str.lower().str.contains('|'.join(out_of_sla_values), na=False)
        return df[mask]
    
    def get_filtered_count(self) -> int:
        """Get count of filtered records"""
        filtered = self.apply_filters()
        return len(filtered)
    
    def get_total_count(self) -> int:
        """Get total record count (unfiltered)"""
        return self.data_loader.get_record_count()
    
    def get_filtered_data(self) -> pd.DataFrame:
        """Get the filtered DataFrame"""
        return self.apply_filters()
    
    def get_preview_data(self, rows: int = 10) -> List[Dict[str, Any]]:
        """Get preview of filtered data"""
        filtered = self.apply_filters()
        return filtered.head(rows).to_dict('records')
    
    # Aggregation methods for charts
    
    def get_count_by_column(self, internal_name: str) -> pd.Series:
        """Get value counts for a column"""
        filtered = self.apply_filters()
        excel_col = self.data_loader.get_column(internal_name)
        
        if excel_col and excel_col in filtered.columns:
            return filtered[excel_col].value_counts()
        
        return pd.Series()
    
    def get_count_by_year(self) -> pd.Series:
        """Get vulnerability count by year"""
        filtered = self.apply_filters()
        creation_col = self.data_loader.get_column("creation_date")
        
        if creation_col and creation_col in filtered.columns:
            try:
                date_series = pd.to_datetime(filtered[creation_col], errors='coerce')
                return date_series.dt.year.value_counts().sort_index()
            except Exception:
                pass
        
        return pd.Series()
    
    def get_count_by_month(self) -> pd.Series:
        """Get vulnerability count by month (for trend analysis)"""
        filtered = self.apply_filters()
        creation_col = self.data_loader.get_column("creation_date")
        
        if creation_col and creation_col in filtered.columns:
            try:
                date_series = pd.to_datetime(filtered[creation_col], errors='coerce')
                # Group by year-month
                return date_series.dt.to_period('M').value_counts().sort_index()
            except Exception:
                pass
        
        return pd.Series()
    
    def get_top_vulnerabilities(self, n: int = 10) -> pd.DataFrame:
        """Get top N most common vulnerabilities"""
        filtered = self.apply_filters()
        plugin_desc_col = self.data_loader.get_column("plugin_desc")
        plugin_id_col = self.data_loader.get_column("plugin_id")
        
        if plugin_desc_col and plugin_desc_col in filtered.columns:
            # Group by plugin description
            counts = filtered[plugin_desc_col].value_counts().head(n)
            
            result_df = pd.DataFrame({
                'vulnerability': counts.index,
                'count': counts.values
            })
            
            # Try to add plugin IDs if available
            if plugin_id_col and plugin_id_col in filtered.columns:
                # Get first plugin_id for each plugin_desc
                id_map = filtered.groupby(plugin_desc_col)[plugin_id_col].first().to_dict()
                result_df['plugin_id'] = result_df['vulnerability'].map(id_map)
            
            return result_df
        
        return pd.DataFrame()
    
    def get_sla_summary(self) -> Dict[str, int]:
        """Get SLA status summary"""
        filtered = self.apply_filters()
        sla_col = self.data_loader.get_column("sla_status")
        
        summary = {"in_sla": 0, "out_of_sla": 0, "unknown": 0}
        
        if sla_col and sla_col in filtered.columns:
            for val in filtered[sla_col]:
                val_str = str(val).lower()
                if "out" in val_str or "breach" in val_str:
                    summary["out_of_sla"] += 1
                elif "in" in val_str and "sla" in val_str:
                    summary["in_sla"] += 1
                else:
                    summary["unknown"] += 1
        
        return summary
    
    def get_priority_summary(self) -> Dict[str, int]:
        """Get priority distribution"""
        counts = self.get_count_by_column("priority")
        return counts.to_dict()
    
    def get_resolution_time_stats(self) -> Dict[str, float]:
        """Get resolution time statistics for closed tickets"""
        filtered = self.apply_filters()
        creation_col = self.data_loader.get_column("creation_date")
        closed_col = self.data_loader.get_column("closed_date")
        status_col = self.data_loader.get_column("status")
        
        stats = {"mean": 0, "median": 0, "min": 0, "max": 0}
        
        if not all([creation_col, closed_col, status_col]):
            return stats
        
        try:
            # Filter to closed tickets only
            closed_statuses = [s.upper() for s in self.data_loader.CLOSED_STATUSES]
            closed_mask = filtered[status_col].astype(str).str.upper().isin(closed_statuses)
            closed_df = filtered[closed_mask]
            
            if len(closed_df) == 0:
                return stats
            
            # Calculate resolution time in days
            creation = pd.to_datetime(closed_df[creation_col], errors='coerce')
            closed = pd.to_datetime(closed_df[closed_col], errors='coerce')
            
            resolution_days = (closed - creation).dt.days
            resolution_days = resolution_days.dropna()
            resolution_days = resolution_days[resolution_days >= 0]  # Remove invalid values
            
            if len(resolution_days) > 0:
                stats["mean"] = round(resolution_days.mean(), 1)
                stats["median"] = round(resolution_days.median(), 1)
                stats["min"] = int(resolution_days.min())
                stats["max"] = int(resolution_days.max())
        
        except Exception:
            pass
        
        return stats
    
    def get_sla_breach_distribution(self) -> pd.Series:
        """Get distribution of SLA breach days"""
        filtered = self.apply_filters()
        sla_time_col = self.data_loader.get_column("sla_time")
        
        if sla_time_col and sla_time_col in filtered.columns:
            try:
                sla_times = pd.to_numeric(filtered[sla_time_col], errors='coerce')
                # Only negative values indicate breach (days overdue)
                breached = sla_times[sla_times < 0].abs()
                return breached
            except Exception:
                pass
        
        return pd.Series()
    
    def get_ip_density(self, top_n: int = 10) -> pd.Series:
        """Get vulnerability count by IP"""
        counts = self.get_count_by_column("ip")
        return counts.head(top_n)
