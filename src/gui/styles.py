"""
Styles Module for ReportIQ GUI
Defines colors, fonts, and styling constants
"""


class Styles:
    """UI Styling constants for the application"""
    
    # Color Palette - Cybersecurity Dark Theme
    COLORS = {
        # Backgrounds
        "bg_primary": "#0d1117",      # Main background (GitHub dark)
        "bg_secondary": "#161b22",    # Secondary panels
        "bg_tertiary": "#21262d",     # Cards/Elevated elements
        "bg_hover": "#30363d",        # Hover states
        
        # Accent colors
        "accent_primary": "#00d4aa",   # Teal/Cyan - Primary actions
        "accent_secondary": "#7c3aed", # Purple - Secondary elements
        "accent_gradient_start": "#00d4aa",
        "accent_gradient_end": "#00b4d8",
        
        # Status colors
        "success": "#4ade80",          # Green
        "warning": "#fbbf24",          # Yellow/Orange
        "danger": "#f87171",           # Red
        "info": "#60a5fa",             # Blue
        
        # Text colors
        "text_primary": "#f0f6fc",     # Primary text
        "text_secondary": "#8b949e",   # Secondary text
        "text_muted": "#6e7681",       # Muted text
        
        # Border colors
        "border": "#30363d",
        "border_focus": "#00d4aa",
        
        # Chart specific
        "chart_bg": "#1a1a2e",
        "chart_grid": "#4a4a6a",
    }
    
    # Font Sizes
    FONTS = {
        "title": ("Segoe UI", 24, "bold"),
        "subtitle": ("Segoe UI", 16, "normal"),
        "heading": ("Segoe UI", 14, "bold"),
        "body": ("Segoe UI", 12, "normal"),
        "small": ("Segoe UI", 10, "normal"),
        "button": ("Segoe UI", 12, "bold"),
        "code": ("Consolas", 11, "normal"),
    }
    
    # Widget Dimensions
    DIMENSIONS = {
        "corner_radius": 8,
        "border_width": 1,
        "button_height": 40,
        "input_height": 35,
        "padding_small": 5,
        "padding_medium": 10,
        "padding_large": 20,
        "sidebar_width": 280,
        "preview_height": 200,
    }
    
    # Button Styles
    BUTTON_STYLES = {
        "primary": {
            "fg_color": "#00d4aa",
            "hover_color": "#00b494",
            "text_color": "#0d1117",
            "corner_radius": 8,
        },
        "secondary": {
            "fg_color": "#30363d",
            "hover_color": "#484f58",
            "text_color": "#f0f6fc",
            "corner_radius": 8,
        },
        "danger": {
            "fg_color": "#f87171",
            "hover_color": "#dc2626",
            "text_color": "#0d1117",
            "corner_radius": 8,
        },
        "ghost": {
            "fg_color": "transparent",
            "hover_color": "#30363d",
            "text_color": "#8b949e",
            "corner_radius": 8,
        },
    }
    
    # Checkbox Styles
    CHECKBOX_STYLE = {
        "fg_color": "#00d4aa",
        "hover_color": "#00b494",
        "border_color": "#30363d",
        "checkmark_color": "#0d1117",
        "text_color": "#f0f6fc",
        "corner_radius": 4,
    }
    
    # Entry/Input Styles
    ENTRY_STYLE = {
        "fg_color": "#21262d",
        "border_color": "#30363d",
        "text_color": "#f0f6fc",
        "placeholder_text_color": "#6e7681",
        "corner_radius": 6,
    }
    
    # Dropdown Styles
    DROPDOWN_STYLE = {
        "fg_color": "#21262d",
        "button_color": "#30363d",
        "button_hover_color": "#484f58",
        "dropdown_fg_color": "#21262d",
        "dropdown_hover_color": "#30363d",
        "text_color": "#f0f6fc",
        "corner_radius": 6,
    }
    
    # Progress Bar Styles
    PROGRESSBAR_STYLE = {
        "fg_color": "#21262d",
        "progress_color": "#00d4aa",
        "corner_radius": 4,
    }
    
    # Label Styles
    @classmethod
    def get_label_style(cls, variant: str = "normal") -> dict:
        """Get label style based on variant"""
        styles = {
            "title": {
                "text_color": cls.COLORS["text_primary"],
                "font": cls.FONTS["title"],
            },
            "subtitle": {
                "text_color": cls.COLORS["text_secondary"],
                "font": cls.FONTS["subtitle"],
            },
            "heading": {
                "text_color": cls.COLORS["text_primary"],
                "font": cls.FONTS["heading"],
            },
            "normal": {
                "text_color": cls.COLORS["text_primary"],
                "font": cls.FONTS["body"],
            },
            "muted": {
                "text_color": cls.COLORS["text_muted"],
                "font": cls.FONTS["small"],
            },
            "accent": {
                "text_color": cls.COLORS["accent_primary"],
                "font": cls.FONTS["body"],
            },
        }
        return styles.get(variant, styles["normal"])
    
    # Frame Styles
    @classmethod
    def get_frame_style(cls, variant: str = "default") -> dict:
        """Get frame style based on variant"""
        styles = {
            "default": {
                "fg_color": cls.COLORS["bg_secondary"],
                "corner_radius": cls.DIMENSIONS["corner_radius"],
            },
            "card": {
                "fg_color": cls.COLORS["bg_tertiary"],
                "corner_radius": cls.DIMENSIONS["corner_radius"],
                "border_width": 1,
                "border_color": cls.COLORS["border"],
            },
            "transparent": {
                "fg_color": "transparent",
                "corner_radius": 0,
            },
        }
        return styles.get(variant, styles["default"])
    
    # Scrollable Frame Styles
    SCROLLABLE_FRAME_STYLE = {
        "fg_color": "transparent",
        "corner_radius": 0,
        "scrollbar_fg_color": "#30363d",
        "scrollbar_button_color": "#484f58",
        "scrollbar_button_hover_color": "#6e7681",
    }
    
    # Tabview Styles
    TABVIEW_STYLE = {
        "fg_color": "#161b22",
        "segmented_button_fg_color": "#21262d",
        "segmented_button_selected_color": "#00d4aa",
        "segmented_button_selected_hover_color": "#00b494",
        "segmented_button_unselected_color": "#21262d",
        "segmented_button_unselected_hover_color": "#30363d",
        "text_color": "#f0f6fc",
        "corner_radius": 8,
    }
