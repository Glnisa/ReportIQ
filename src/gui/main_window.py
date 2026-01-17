"""
Main Window Module for ReportIQ
The primary GUI window with all components
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Optional, Dict, List, Callable
import threading
from datetime import datetime

from .styles import Styles
from ..data.translations import Translations, t
from ..data.vulnerability_dict import VulnerabilityDictionary
from ..core.data_loader import DataLoader
from ..core.filter_engine import FilterEngine
from ..core.chart_generator import ChartGenerator
from ..core.word_generator import WordGenerator


class MainWindow(ctk.CTk):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        
        # Configure appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        
        # Window setup
        self.title(t("app_title"))
        self.geometry("1400x900")
        self.minsize(1200, 700)
        
        # Set colors
        self.configure(fg_color=Styles.COLORS["bg_primary"])
        
        # Initialize components
        self.data_loader = DataLoader()
        self.filter_engine: Optional[FilterEngine] = None
        self.chart_generator: Optional[ChartGenerator] = None
        self.word_generator: Optional[WordGenerator] = None
        
        # State variables
        self.file_path: Optional[str] = None
        self.filter_vars: Dict[str, ctk.Variable] = {}
        self.section_vars: Dict[str, ctk.BooleanVar] = {}
        self.is_generating = False
        
        # Build UI
        self._create_widgets()
        self._bind_events()
    
    def _create_widgets(self) -> None:
        """Create all UI widgets"""
        # Main container
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header
        self._create_header()
        
        # Content area (3-column layout)
        self.content_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, pady=(20, 0))
        
        # Configure grid
        self.content_frame.grid_columnconfigure(0, weight=1, minsize=280)
        self.content_frame.grid_columnconfigure(1, weight=1, minsize=300)
        self.content_frame.grid_columnconfigure(2, weight=2, minsize=400)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # Left panel - Filters
        self._create_filter_panel()
        
        # Middle panel - Report sections
        self._create_sections_panel()
        
        # Right panel - Preview & Actions
        self._create_preview_panel()
        
        # Bottom - Generate button & Progress
        self._create_action_bar()
    
    def _create_header(self) -> None:
        """Create header with title and language switch"""
        header = ctk.CTkFrame(self.main_container, fg_color="transparent", height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        # Left side - Logo and title
        left_frame = ctk.CTkFrame(header, fg_color="transparent")
        left_frame.pack(side="left", fill="y")
        
        # Shield emoji as logo
        logo_label = ctk.CTkLabel(
            left_frame, 
            text="🛡️",
            font=("Segoe UI", 36)
        )
        logo_label.pack(side="left", padx=(0, 10))
        
        # Title and subtitle
        title_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        title_frame.pack(side="left", fill="y", pady=5)
        
        self.title_label = ctk.CTkLabel(
            title_frame,
            text="ReportIQ",
            font=Styles.FONTS["title"],
            text_color=Styles.COLORS["text_primary"]
        )
        self.title_label.pack(anchor="w")
        
        self.subtitle_label = ctk.CTkLabel(
            title_frame,
            text=t("app_subtitle"),
            font=Styles.FONTS["small"],
            text_color=Styles.COLORS["text_muted"]
        )
        self.subtitle_label.pack(anchor="w")
        
        # Right side - Language switch
        right_frame = ctk.CTkFrame(header, fg_color="transparent")
        right_frame.pack(side="right", fill="y")
        
        self.lang_switch = ctk.CTkSegmentedButton(
            right_frame,
            values=["TR", "EN"],
            command=self._on_language_change,
            font=Styles.FONTS["button"],
            fg_color=Styles.COLORS["bg_tertiary"],
            selected_color=Styles.COLORS["accent_primary"],
            selected_hover_color="#00b494",
            unselected_color=Styles.COLORS["bg_tertiary"],
            unselected_hover_color=Styles.COLORS["bg_hover"],
            text_color=Styles.COLORS["text_primary"],
        )
        self.lang_switch.set(Translations.get_language())
        self.lang_switch.pack(pady=15)
    
    def _create_filter_panel(self) -> None:
        """Create the filters panel on the left"""
        filter_frame = ctk.CTkFrame(
            self.content_frame,
            fg_color=Styles.COLORS["bg_secondary"],
            corner_radius=Styles.DIMENSIONS["corner_radius"]
        )
        filter_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Header
        header = ctk.CTkFrame(filter_frame, fg_color="transparent", height=50)
        header.pack(fill="x", padx=15, pady=(15, 5))
        header.pack_propagate(False)
        
        self.filter_title = ctk.CTkLabel(
            header,
            text=t("filters"),
            font=Styles.FONTS["heading"],
            text_color=Styles.COLORS["accent_primary"]
        )
        self.filter_title.pack(side="left")
        
        # Scrollable filter content
        self.filter_scroll = ctk.CTkScrollableFrame(
            filter_frame,
            fg_color="transparent",
            scrollbar_fg_color=Styles.COLORS["bg_tertiary"],
            scrollbar_button_color=Styles.COLORS["bg_hover"]
        )
        self.filter_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # File selection section
        self._create_file_section(self.filter_scroll)
        
        # Filter options (will be populated after file load)
        self.filter_options_frame = ctk.CTkFrame(self.filter_scroll, fg_color="transparent")
        self.filter_options_frame.pack(fill="x", pady=(10, 0))
        
        # Placeholder message
        self.filter_placeholder = ctk.CTkLabel(
            self.filter_options_frame,
            text=t("no_file_selected"),
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["text_muted"]
        )
        self.filter_placeholder.pack(pady=30)
    
    def _create_file_section(self, parent) -> None:
        """Create file selection section"""
        file_frame = ctk.CTkFrame(
            parent,
            fg_color=Styles.COLORS["bg_tertiary"],
            corner_radius=Styles.DIMENSIONS["corner_radius"]
        )
        file_frame.pack(fill="x", pady=(0, 10))
        
        # File icon and label
        file_header = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_header.pack(fill="x", padx=15, pady=(15, 10))
        
        self.file_label = ctk.CTkLabel(
            file_header,
            text=t("select_file"),
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["text_primary"]
        )
        self.file_label.pack(side="left")
        
        # Browse button
        self.browse_btn = ctk.CTkButton(
            file_frame,
            text=t("browse"),
            command=self._on_browse_file,
            **Styles.BUTTON_STYLES["primary"],
            font=Styles.FONTS["button"],
            height=36
        )
        self.browse_btn.pack(fill="x", padx=15, pady=(0, 10))
        
        # File path display
        self.file_path_label = ctk.CTkLabel(
            file_frame,
            text="",
            font=Styles.FONTS["small"],
            text_color=Styles.COLORS["text_muted"],
            wraplength=230
        )
        self.file_path_label.pack(fill="x", padx=15, pady=(0, 15))
    
    def _create_sections_panel(self) -> None:
        """Create the report sections panel in the middle"""
        sections_frame = ctk.CTkFrame(
            self.content_frame,
            fg_color=Styles.COLORS["bg_secondary"],
            corner_radius=Styles.DIMENSIONS["corner_radius"]
        )
        sections_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        
        # Header
        header = ctk.CTkFrame(sections_frame, fg_color="transparent", height=50)
        header.pack(fill="x", padx=15, pady=(15, 5))
        header.pack_propagate(False)
        
        self.sections_title = ctk.CTkLabel(
            header,
            text=t("report_sections"),
            font=Styles.FONTS["heading"],
            text_color=Styles.COLORS["accent_primary"]
        )
        self.sections_title.pack(side="left")
        
        # Select/Clear all buttons
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right")
        
        self.select_all_btn = ctk.CTkButton(
            btn_frame,
            text=t("select_all"),
            command=self._select_all_sections,
            width=80,
            height=28,
            font=Styles.FONTS["small"],
            **Styles.BUTTON_STYLES["secondary"]
        )
        self.select_all_btn.pack(side="left", padx=(0, 5))
        
        self.clear_all_btn = ctk.CTkButton(
            btn_frame,
            text=t("clear_all"),
            command=self._clear_all_sections,
            width=70,
            height=28,
            font=Styles.FONTS["small"],
            **Styles.BUTTON_STYLES["ghost"]
        )
        self.clear_all_btn.pack(side="left")
        
        # Scrollable sections content
        sections_scroll = ctk.CTkScrollableFrame(
            sections_frame,
            fg_color="transparent",
            scrollbar_fg_color=Styles.COLORS["bg_tertiary"],
            scrollbar_button_color=Styles.COLORS["bg_hover"]
        )
        sections_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Create section checkboxes
        self._create_section_checkboxes(sections_scroll)
    
    def _create_section_checkboxes(self, parent) -> None:
        """Create checkboxes for report sections"""
        sections = [
            ("yearly_open", "chart_yearly_open"),
            ("priority_dist", "chart_priority_dist"),
            ("line_manager", "chart_line_manager"),
            ("department", "chart_department"),
            ("tool", "chart_tool"),
            ("sla", "chart_sla"),
            ("trend", "chart_trend"),
            ("top10", "chart_top10"),
            ("ip_density", "chart_ip_density"),
            ("resolution_time", "chart_resolution_time"),
            ("sla_breach", "chart_sla_breach"),
        ]
        
        for section_id, label_key in sections:
            self.section_vars[section_id] = ctk.BooleanVar(value=True)
            
            cb = ctk.CTkCheckBox(
                parent,
                text=t(label_key),
                variable=self.section_vars[section_id],
                font=Styles.FONTS["body"],
                fg_color=Styles.COLORS["accent_primary"],
                hover_color="#00b494",
                border_color=Styles.COLORS["border"],
                checkmark_color=Styles.COLORS["bg_primary"],
                text_color=Styles.COLORS["text_primary"],
                corner_radius=4
            )
            cb.pack(fill="x", pady=5, padx=5)
    
    def _create_preview_panel(self) -> None:
        """Create the preview panel on the right"""
        preview_frame = ctk.CTkFrame(
            self.content_frame,
            fg_color=Styles.COLORS["bg_secondary"],
            corner_radius=Styles.DIMENSIONS["corner_radius"]
        )
        preview_frame.grid(row=0, column=2, sticky="nsew")
        
        # Header
        header = ctk.CTkFrame(preview_frame, fg_color="transparent", height=50)
        header.pack(fill="x", padx=15, pady=(15, 5))
        header.pack_propagate(False)
        
        self.preview_title = ctk.CTkLabel(
            header,
            text=t("data_preview"),
            font=Styles.FONTS["heading"],
            text_color=Styles.COLORS["accent_primary"]
        )
        self.preview_title.pack(side="left")
        
        # Count labels
        count_frame = ctk.CTkFrame(header, fg_color="transparent")
        count_frame.pack(side="right")
        
        self.filtered_count_label = ctk.CTkLabel(
            count_frame,
            text="",
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["accent_primary"]
        )
        self.filtered_count_label.pack(side="right")
        
        self.total_count_label = ctk.CTkLabel(
            count_frame,
            text="",
            font=Styles.FONTS["small"],
            text_color=Styles.COLORS["text_muted"]
        )
        self.total_count_label.pack(side="right", padx=(0, 10))
        
        # Data preview table
        self.preview_container = ctk.CTkFrame(
            preview_frame,
            fg_color=Styles.COLORS["bg_tertiary"],
            corner_radius=Styles.DIMENSIONS["corner_radius"]
        )
        self.preview_container.pack(fill="both", expand=True, padx=15, pady=(5, 15))
        
        # Preview text widget
        self.preview_text = ctk.CTkTextbox(
            self.preview_container,
            font=Styles.FONTS["code"],
            fg_color=Styles.COLORS["bg_tertiary"],
            text_color=Styles.COLORS["text_primary"],
            scrollbar_button_color=Styles.COLORS["bg_hover"],
            wrap="none"
        )
        self.preview_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.preview_text.configure(state="disabled")
        
        # Placeholder
        self._update_preview_placeholder()
    
    def _create_action_bar(self) -> None:
        """Create bottom action bar with generate button and progress"""
        action_frame = ctk.CTkFrame(self.main_container, fg_color="transparent", height=80)
        action_frame.pack(fill="x", pady=(20, 0))
        action_frame.pack_propagate(False)
        
        # Progress bar
        self.progress_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame,
            fg_color=Styles.COLORS["bg_tertiary"],
            progress_color=Styles.COLORS["accent_primary"],
            height=8,
            corner_radius=4
        )
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="",
            font=Styles.FONTS["small"],
            text_color=Styles.COLORS["text_muted"]
        )
        self.progress_label.pack(pady=(5, 0))
        
        # Hide progress initially
        self.progress_frame.pack_forget()
        
        # Generate button
        self.generate_btn = ctk.CTkButton(
            action_frame,
            text=t("generate_report"),
            command=self._on_generate_report,
            font=("Segoe UI", 16, "bold"),
            height=50,
            fg_color=Styles.COLORS["accent_primary"],
            hover_color="#00b494",
            text_color=Styles.COLORS["bg_primary"],
            corner_radius=10
        )
        self.generate_btn.pack(fill="x")
    
    def _bind_events(self) -> None:
        """Bind event handlers"""
        pass  # Events are bound inline in widget creation
    
    def _on_language_change(self, value: str) -> None:
        """Handle language switch"""
        Translations.set_language(value)
        self._update_all_labels()
    
    def _update_all_labels(self) -> None:
        """Update all labels after language change"""
        # Update title
        self.title(t("app_title"))
        self.subtitle_label.configure(text=t("app_subtitle"))
        
        # Update section titles
        self.filter_title.configure(text=t("filters"))
        self.sections_title.configure(text=t("report_sections"))
        self.preview_title.configure(text=t("data_preview"))
        
        # Update buttons
        self.browse_btn.configure(text=t("browse"))
        self.file_label.configure(text=t("select_file"))
        self.select_all_btn.configure(text=t("select_all"))
        self.clear_all_btn.configure(text=t("clear_all"))
        self.generate_btn.configure(text=t("generate_report"))
        
        # Update section checkboxes
        sections = [
            ("yearly_open", "chart_yearly_open"),
            ("priority_dist", "chart_priority_dist"),
            ("line_manager", "chart_line_manager"),
            ("department", "chart_department"),
            ("tool", "chart_tool"),
            ("sla", "chart_sla"),
            ("trend", "chart_trend"),
            ("top10", "chart_top10"),
            ("ip_density", "chart_ip_density"),
            ("resolution_time", "chart_resolution_time"),
            ("sla_breach", "chart_sla_breach"),
        ]
        
        # Recreate sections panel to update labels
        # (Would need to store checkbox references for direct update)
        
        # Update counts
        self._update_counts()
        
        # Update filter labels if loaded
        if self.data_loader.is_loaded():
            self._populate_filter_options()
    
    def _on_browse_file(self) -> None:
        """Handle file browse button"""
        file_path = filedialog.askopenfilename(
            title=t("select_file"),
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self._load_file(file_path)
    
    def _load_file(self, file_path: str) -> None:
        """Load the selected Excel file"""
        self.file_path_label.configure(
            text=t("loading_file"),
            text_color=Styles.COLORS["warning"]
        )
        self.update()
        
        success, message = self.data_loader.load_file(file_path)
        
        if success:
            self.file_path = file_path
            filename = Path(file_path).name
            self.file_path_label.configure(
                text=f"✓ {filename}",
                text_color=Styles.COLORS["success"]
            )
            
            # Initialize filter engine
            self.filter_engine = FilterEngine(self.data_loader)
            
            # Populate filter options
            self._populate_filter_options()
            
            # Update preview
            self._update_preview()
            
            # Update counts
            self._update_counts()
        else:
            self.file_path_label.configure(
                text=f"✗ {message}",
                text_color=Styles.COLORS["danger"]
            )
    
    def _populate_filter_options(self) -> None:
        """Populate filter options based on loaded data"""
        # Clear placeholder
        self.filter_placeholder.pack_forget()
        
        # Clear existing filter widgets
        for widget in self.filter_options_frame.winfo_children():
            widget.destroy()
        
        self.filter_vars.clear()
        
        # SLA Status filter
        self._create_filter_group(
            "sla_status",
            t("sla_status"),
            self.data_loader.get_sla_statuses()
        )
        
        # Status filter
        self._create_filter_group(
            "status",
            t("status"),
            self.data_loader.get_statuses(),
            preselect=self.data_loader.get_open_status_values()
        )
        
        # Priority filter
        self._create_filter_group(
            "priority",
            t("priority"),
            self.data_loader.get_priorities()
        )
        
        # Tool filter
        tools = self.data_loader.get_tools()
        if tools:
            self._create_filter_group(
                "tool",
                t("tool_source"),
                tools[:15]  # Limit to 15 for display
            )
        
        # Year filter
        years = self.data_loader.get_years()
        if years:
            self._create_year_filter(years)
        
        # Department filter (dropdown for many values)
        depts = self.data_loader.get_departments()
        if depts and len(depts) > 5:
            self._create_dropdown_filter(
                "department",
                t("department"),
                ["All"] + depts
            )
        elif depts:
            self._create_filter_group("department", t("department"), depts)
    
    def _create_filter_group(self, filter_name: str, title: str, 
                              options: List[str], preselect: List[str] = None) -> None:
        """Create a group of checkboxes for a filter"""
        if not options:
            return
        
        frame = ctk.CTkFrame(
            self.filter_options_frame,
            fg_color=Styles.COLORS["bg_tertiary"],
            corner_radius=6
        )
        frame.pack(fill="x", pady=(0, 8))
        
        # Title
        title_label = ctk.CTkLabel(
            frame,
            text=title,
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["text_primary"]
        )
        title_label.pack(anchor="w", padx=10, pady=(8, 5))
        
        # Options
        self.filter_vars[filter_name] = {}
        for option in options:
            var = ctk.BooleanVar(value=preselect is None or option in (preselect or []))
            self.filter_vars[filter_name][option] = var
            
            cb = ctk.CTkCheckBox(
                frame,
                text=str(option)[:25],
                variable=var,
                command=self._on_filter_change,
                font=Styles.FONTS["small"],
                fg_color=Styles.COLORS["accent_primary"],
                hover_color="#00b494",
                border_color=Styles.COLORS["border"],
                checkmark_color=Styles.COLORS["bg_primary"],
                text_color=Styles.COLORS["text_secondary"],
                height=24,
                corner_radius=3
            )
            cb.pack(anchor="w", padx=15, pady=2)
        
        # Add padding at bottom
        ctk.CTkFrame(frame, fg_color="transparent", height=5).pack()
    
    def _create_year_filter(self, years: List[int]) -> None:
        """Create year range filter"""
        frame = ctk.CTkFrame(
            self.filter_options_frame,
            fg_color=Styles.COLORS["bg_tertiary"],
            corner_radius=6
        )
        frame.pack(fill="x", pady=(0, 8))
        
        # Title
        title_label = ctk.CTkLabel(
            frame,
            text=t("year"),
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["text_primary"]
        )
        title_label.pack(anchor="w", padx=10, pady=(8, 5))
        
        # All years checkbox
        self.all_years_var = ctk.BooleanVar(value=True)
        all_years_cb = ctk.CTkCheckBox(
            frame,
            text=t("all_years"),
            variable=self.all_years_var,
            command=self._on_filter_change,
            font=Styles.FONTS["small"],
            fg_color=Styles.COLORS["accent_primary"],
            hover_color="#00b494",
            border_color=Styles.COLORS["border"],
            text_color=Styles.COLORS["text_secondary"],
            height=24
        )
        all_years_cb.pack(anchor="w", padx=15, pady=2)
        
        # Year dropdown
        year_strs = [str(y) for y in years]
        self.year_var = ctk.StringVar(value=year_strs[-1] if year_strs else "")
        
        year_dropdown = ctk.CTkOptionMenu(
            frame,
            values=year_strs,
            variable=self.year_var,
            command=lambda x: self._on_filter_change(),
            font=Styles.FONTS["small"],
            fg_color=Styles.COLORS["bg_secondary"],
            button_color=Styles.COLORS["bg_hover"],
            button_hover_color=Styles.COLORS["accent_primary"],
            dropdown_fg_color=Styles.COLORS["bg_secondary"],
            dropdown_hover_color=Styles.COLORS["bg_hover"],
            text_color=Styles.COLORS["text_primary"],
            width=100,
            height=28
        )
        year_dropdown.pack(anchor="w", padx=15, pady=(5, 10))
    
    def _create_dropdown_filter(self, filter_name: str, title: str, options: List[str]) -> None:
        """Create a dropdown filter for many options"""
        frame = ctk.CTkFrame(
            self.filter_options_frame,
            fg_color=Styles.COLORS["bg_tertiary"],
            corner_radius=6
        )
        frame.pack(fill="x", pady=(0, 8))
        
        # Title
        title_label = ctk.CTkLabel(
            frame,
            text=title,
            font=Styles.FONTS["body"],
            text_color=Styles.COLORS["text_primary"]
        )
        title_label.pack(anchor="w", padx=10, pady=(8, 5))
        
        # Dropdown
        var = ctk.StringVar(value="All")
        self.filter_vars[filter_name] = var
        
        dropdown = ctk.CTkOptionMenu(
            frame,
            values=options,
            variable=var,
            command=lambda x: self._on_filter_change(),
            font=Styles.FONTS["small"],
            fg_color=Styles.COLORS["bg_secondary"],
            button_color=Styles.COLORS["bg_hover"],
            button_hover_color=Styles.COLORS["accent_primary"],
            dropdown_fg_color=Styles.COLORS["bg_secondary"],
            dropdown_hover_color=Styles.COLORS["bg_hover"],
            text_color=Styles.COLORS["text_primary"],
            width=220,
            height=30
        )
        dropdown.pack(anchor="w", padx=10, pady=(0, 10))
    
    def _on_filter_change(self) -> None:
        """Handle filter change"""
        if not self.filter_engine:
            return
        
        # Clear existing filters
        self.filter_engine.clear_filters()
        
        # Apply checkbox filters
        for filter_name, values in self.filter_vars.items():
            if isinstance(values, dict):
                # Checkbox group
                selected = [k for k, v in values.items() if v.get()]
                if selected:
                    self.filter_engine.set_filter(filter_name, selected)
            elif isinstance(values, ctk.StringVar):
                # Dropdown
                val = values.get()
                if val and val != "All":
                    self.filter_engine.set_filter(filter_name, [val])
        
        # Apply year filter
        if hasattr(self, 'all_years_var') and not self.all_years_var.get():
            if hasattr(self, 'year_var') and self.year_var.get():
                try:
                    year = int(self.year_var.get())
                    self.filter_engine.set_filter("year", [year])
                except ValueError:
                    pass
        
        # Update preview and counts
        self._update_preview()
        self._update_counts()
    
    def _update_preview(self) -> None:
        """Update the data preview"""
        if not self.filter_engine:
            self._update_preview_placeholder()
            return
        
        filtered = self.filter_engine.get_filtered_data()
        
        if filtered.empty:
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", t("error_no_data"))
            self.preview_text.configure(state="disabled")
            return
        
        # Format preview
        preview_df = filtered.head(50)
        preview_str = preview_df.to_string(max_cols=8, max_colwidth=20)
        
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", preview_str)
        self.preview_text.configure(state="disabled")
    
    def _update_preview_placeholder(self) -> None:
        """Show placeholder in preview"""
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", t("no_file_selected"))
        self.preview_text.configure(state="disabled")
    
    def _update_counts(self) -> None:
        """Update count labels"""
        if not self.filter_engine:
            self.total_count_label.configure(text="")
            self.filtered_count_label.configure(text="")
            return
        
        total = self.filter_engine.get_total_count()
        filtered = self.filter_engine.get_filtered_count()
        
        self.total_count_label.configure(text=t("total_count", count=f"{total:,}"))
        self.filtered_count_label.configure(text=t("filtered_count", count=f"{filtered:,}"))
    
    def _select_all_sections(self) -> None:
        """Select all report sections"""
        for var in self.section_vars.values():
            var.set(True)
    
    def _clear_all_sections(self) -> None:
        """Clear all report sections"""
        for var in self.section_vars.values():
            var.set(False)
    
    def _on_generate_report(self) -> None:
        """Handle generate report button"""
        if not self.filter_engine:
            messagebox.showerror(t("error"), t("no_file_selected"))
            return
        
        if self.filter_engine.get_filtered_count() == 0:
            messagebox.showerror(t("error"), t("error_no_data"))
            return
        
        # Get selected sections
        selected_sections = [k for k, v in self.section_vars.items() if v.get()]
        
        if not selected_sections:
            messagebox.showerror(t("error"), t("error_no_data"))
            return
        
        # Ask for save location
        output_path = filedialog.asksaveasfilename(
            title=t("generate_report"),
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=f"VulnerabilityReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )
        
        if not output_path:
            return
        
        # Start generation in background thread
        self.is_generating = True
        self.generate_btn.configure(state="disabled")
        self.progress_frame.pack(fill="x", pady=(0, 10))
        
        thread = threading.Thread(
            target=self._generate_report_thread,
            args=(selected_sections, output_path)
        )
        thread.start()
    
    def _generate_report_thread(self, sections: List[str], output_path: str) -> None:
        """Generate report in background thread"""
        try:
            language = Translations.get_language()
            
            # Update progress
            self._update_progress(0.1, t("creating_charts"))
            
            # Initialize generators
            self.chart_generator = ChartGenerator(self.filter_engine, language)
            
            # Generate charts
            total_sections = len(sections)
            for i, section in enumerate(sections):
                progress = 0.1 + (0.6 * (i + 1) / total_sections)
                self._update_progress(progress, f"{t('creating_charts')} ({i+1}/{total_sections})")
                self.chart_generator.generate_all_charts([section])
            
            # Create Word document
            self._update_progress(0.75, t("creating_document"))
            
            self.word_generator = WordGenerator(
                self.filter_engine,
                self.chart_generator,
                language
            )
            self.word_generator.set_vulnerability_dict(VulnerabilityDictionary)
            
            # Generate report
            self._update_progress(0.85, t("creating_document"))
            saved_path = self.word_generator.generate_report(sections, output_path)
            
            # Complete
            self._update_progress(1.0, t("report_saved", path=Path(saved_path).name))
            
            # Show success message
            self.after(500, lambda: self._on_generation_complete(saved_path))
            
        except Exception as e:
            self.after(0, lambda: self._on_generation_error(str(e)))
    
    def _update_progress(self, value: float, message: str) -> None:
        """Update progress bar and label"""
        self.after(0, lambda: self._set_progress(value, message))
    
    def _set_progress(self, value: float, message: str) -> None:
        """Set progress values (must run on main thread)"""
        self.progress_bar.set(value)
        self.progress_label.configure(text=message)
    
    def _on_generation_complete(self, saved_path: str) -> None:
        """Handle generation complete"""
        self.is_generating = False
        self.generate_btn.configure(state="normal")
        
        messagebox.showinfo(
            t("generate_report"),
            t("report_saved", path=saved_path)
        )
        
        # Reset progress after delay
        self.after(2000, self._reset_progress)
    
    def _on_generation_error(self, error: str) -> None:
        """Handle generation error"""
        self.is_generating = False
        self.generate_btn.configure(state="normal")
        self.progress_frame.pack_forget()
        
        messagebox.showerror(t("error"), t("error_generating", error=error))
    
    def _reset_progress(self) -> None:
        """Reset progress bar"""
        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        self.progress_frame.pack_forget()
