"""
Translations module for ReportIQ
Supports Turkish (TR) and English (EN)
"""

from typing import Dict


class Translations:
    """Multi-language support for the application"""
    
    _current_language: str = "TR"
    
    STRINGS: Dict[str, Dict[str, str]] = {
        # Application
        "app_title": {
            "TR": "ReportIQ - Zafiyet Rapor Oluşturucu",
            "EN": "ReportIQ - Vulnerability Report Generator"
        },
        "app_subtitle": {
            "TR": "Güvenlik Analiz Aracı",
            "EN": "Security Analysis Tool"
        },
        
        # File Selection
        "select_file": {
            "TR": "📂 Excel Dosyası Seç",
            "EN": "📂 Select Excel File"
        },
        "browse": {
            "TR": "Gözat",
            "EN": "Browse"
        },
        "file_loaded": {
            "TR": "✓ Dosya Yüklendi",
            "EN": "✓ File Loaded"
        },
        "no_file_selected": {
            "TR": "Dosya seçilmedi",
            "EN": "No file selected"
        },
        "loading_file": {
            "TR": "Dosya yükleniyor...",
            "EN": "Loading file..."
        },
        
        # Filters Section
        "filters": {
            "TR": "📋 FİLTRELER",
            "EN": "📋 FILTERS"
        },
        "sla_status": {
            "TR": "SLA Durumu",
            "EN": "SLA Status"
        },
        "out_of_sla": {
            "TR": "Out of SLA",
            "EN": "Out of SLA"
        },
        "in_sla": {
            "TR": "In SLA",
            "EN": "In SLA"
        },
        "status": {
            "TR": "Durum",
            "EN": "Status"
        },
        "tool_source": {
            "TR": "Kaynak/Tool",
            "EN": "Source/Tool"
        },
        "year": {
            "TR": "Yıl",
            "EN": "Year"
        },
        "all_years": {
            "TR": "Tüm Yıllar",
            "EN": "All Years"
        },
        "department": {
            "TR": "Departman",
            "EN": "Department"
        },
        "line_manager": {
            "TR": "Line Manager",
            "EN": "Line Manager"
        },
        "priority": {
            "TR": "Öncelik",
            "EN": "Priority"
        },
        "select_all": {
            "TR": "Tümünü Seç",
            "EN": "Select All"
        },
        "clear_all": {
            "TR": "Temizle",
            "EN": "Clear All"
        },
        
        # Report Sections
        "report_sections": {
            "TR": "📊 RAPOR BÖLÜMLERİ",
            "EN": "📊 REPORT SECTIONS"
        },
        "chart_yearly_open": {
            "TR": "📊 Yıllara Göre Açık Zafiyet",
            "EN": "📊 Open Vulnerabilities by Year"
        },
        "chart_priority_dist": {
            "TR": "🎯 Priority Dağılımı",
            "EN": "🎯 Priority Distribution"
        },
        "chart_line_manager": {
            "TR": "👥 Line Manager Kırılımı",
            "EN": "👥 Breakdown by Line Manager"
        },
        "chart_department": {
            "TR": "🏢 Departman Kırılımı",
            "EN": "🏢 Breakdown by Department"
        },
        "chart_tool": {
            "TR": "🔧 Tool Kırılımı",
            "EN": "🔧 Breakdown by Tool"
        },
        "chart_sla": {
            "TR": "⏰ SLA Durumu",
            "EN": "⏰ SLA Status"
        },
        "chart_trend": {
            "TR": "📈 Trend Analizi",
            "EN": "📈 Trend Analysis"
        },
        "chart_top10": {
            "TR": "🔥 Top 10 Zafiyet",
            "EN": "🔥 Top 10 Vulnerabilities"
        },
        "chart_ip_density": {
            "TR": "💻 IP Bazlı Yoğunluk",
            "EN": "💻 IP-Based Density"
        },
        "chart_resolution_time": {
            "TR": "📅 Ortalama Çözüm Süresi",
            "EN": "📅 Average Resolution Time"
        },
        "chart_sla_breach": {
            "TR": "⚠️ SLA Aşım Analizi",
            "EN": "⚠️ SLA Breach Analysis"
        },
        
        # Data Preview
        "data_preview": {
            "TR": "👁️ VERİ ÖNİZLEME",
            "EN": "👁️ DATA PREVIEW"
        },
        "filtered_count": {
            "TR": "Filtrelenen: {count} zafiyet",
            "EN": "Filtered: {count} vulnerabilities"
        },
        "total_count": {
            "TR": "Toplam: {count} kayıt",
            "EN": "Total: {count} records"
        },
        
        # Actions
        "generate_report": {
            "TR": "🚀 Rapor Oluştur",
            "EN": "🚀 Generate Report"
        },
        "generating": {
            "TR": "Rapor oluşturuluyor...",
            "EN": "Generating report..."
        },
        "report_saved": {
            "TR": "✓ Rapor kaydedildi: {path}",
            "EN": "✓ Report saved: {path}"
        },
        "creating_charts": {
            "TR": "Grafikler oluşturuluyor...",
            "EN": "Creating charts..."
        },
        "creating_document": {
            "TR": "Word belgesi oluşturuluyor...",
            "EN": "Creating Word document..."
        },
        
        # Errors
        "error": {
            "TR": "Hata",
            "EN": "Error"
        },
        "error_loading_file": {
            "TR": "Dosya yüklenirken hata oluştu: {error}",
            "EN": "Error loading file: {error}"
        },
        "error_no_data": {
            "TR": "Seçilen filtrelere uygun veri bulunamadı",
            "EN": "No data found matching selected filters"
        },
        "error_generating": {
            "TR": "Rapor oluşturulurken hata: {error}",
            "EN": "Error generating report: {error}"
        },
        
        # Column Mapping
        "column_mapping": {
            "TR": "Sütun Eşleştirme",
            "EN": "Column Mapping"
        },
        "map_columns": {
            "TR": "Lütfen Excel sütunlarını eşleştirin",
            "EN": "Please map Excel columns"
        },
        "confirm_mapping": {
            "TR": "Eşleştirmeyi Onayla",
            "EN": "Confirm Mapping"
        },
        
        # Report Title
        "report_title": {
            "TR": "Zafiyet Analiz Raporu",
            "EN": "Vulnerability Analysis Report"
        },
        "report_date": {
            "TR": "Rapor Tarihi",
            "EN": "Report Date"
        },
        "executive_summary": {
            "TR": "Yönetici Özeti",
            "EN": "Executive Summary"
        },
        
        # Status Values
        "status_pending": {
            "TR": "Beklemede",
            "EN": "Pending"
        },
        "status_queued": {
            "TR": "Sırada",
            "EN": "Queued"
        },
        "status_closed": {
            "TR": "Kapalı",
            "EN": "Closed"
        },
        "status_cancelled": {
            "TR": "İptal",
            "EN": "Cancelled"
        },
        
        # Settings
        "settings": {
            "TR": "⚙️ Ayarlar",
            "EN": "⚙️ Settings"
        },
        "language": {
            "TR": "Dil",
            "EN": "Language"
        },
        "theme": {
            "TR": "Tema",
            "EN": "Theme"
        },
        "dark_mode": {
            "TR": "Karanlık Mod",
            "EN": "Dark Mode"
        },
        
        # Open Statuses Description
        "open_statuses": {
            "TR": "Açık Durumlar (PENDING, QUEUED, QUEUEDR, WRISKACCPT)",
            "EN": "Open Statuses (PENDING, QUEUED, QUEUEDR, WRISKACCPT)"
        },
        "closed_statuses": {
            "TR": "Kapalı Durumlar (CLOSED, CANCEL, RISKACCPT)",
            "EN": "Closed Statuses (CLOSED, CANCEL, RISKACCPT)"
        },
    }
    
    @classmethod
    def set_language(cls, lang: str) -> None:
        """Set current language (TR or EN)"""
        if lang in ["TR", "EN"]:
            cls._current_language = lang
    
    @classmethod
    def get_language(cls) -> str:
        """Get current language"""
        return cls._current_language
    
    @classmethod
    def get(cls, key: str, **kwargs) -> str:
        """Get translated string by key"""
        if key in cls.STRINGS:
            text = cls.STRINGS[key].get(cls._current_language, cls.STRINGS[key].get("EN", key))
            if kwargs:
                try:
                    return text.format(**kwargs)
                except KeyError:
                    return text
            return text
        return key
    
    @classmethod
    def toggle_language(cls) -> str:
        """Toggle between TR and EN, returns new language"""
        cls._current_language = "EN" if cls._current_language == "TR" else "TR"
        return cls._current_language


# Convenience function
def t(key: str, **kwargs) -> str:
    """Shorthand for Translations.get()"""
    return Translations.get(key, **kwargs)
