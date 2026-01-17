"""
Chart Generator Module for ReportIQ
Creates matplotlib charts for vulnerability analysis
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any
from io import BytesIO
import warnings

# Suppress matplotlib warnings
warnings.filterwarnings('ignore', category=UserWarning)

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")


class ChartGenerator:
    """Generates charts for vulnerability reports"""
    
    # Color schemes
    COLORS = {
        "primary": "#00D4AA",      # Teal/Cyan
        "secondary": "#7B68EE",    # Purple
        "accent": "#FF6B6B",       # Coral Red
        "warning": "#FFB347",      # Orange
        "success": "#4ECDC4",      # Turquoise
        "danger": "#FF4757",       # Red
        "info": "#5DADE2",         # Blue
        "dark": "#2C3E50",         # Dark Blue
        "light": "#ECF0F1",        # Light Gray
    }
    
    # Priority colors
    PRIORITY_COLORS = {
        "Critical": "#FF4757",
        "High": "#FF6B6B",
        "Medium": "#FFB347",
        "Low": "#4ECDC4",
        "Info": "#5DADE2",
    }
    
    # SLA colors
    SLA_COLORS = {
        "in_sla": "#4ECDC4",
        "out_of_sla": "#FF4757",
        "unknown": "#95A5A6",
    }
    
    def __init__(self, filter_engine, language: str = "TR"):
        """
        Initialize chart generator
        
        Args:
            filter_engine: FilterEngine instance with filtered data
            language: Language code (TR or EN)
        """
        self.filter_engine = filter_engine
        self.language = language
        self.output_dir: Optional[Path] = None
        self._charts: Dict[str, BytesIO] = {}
        
        # Configure matplotlib for Turkish characters
        plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
    
    def set_output_dir(self, path: str) -> None:
        """Set output directory for saving charts"""
        self.output_dir = Path(path)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def set_language(self, language: str) -> None:
        """Set language for chart labels"""
        self.language = language
    
    def _get_label(self, tr: str, en: str) -> str:
        """Get label based on current language"""
        return tr if self.language == "TR" else en
    
    def _create_figure(self, figsize: Tuple[float, float] = (10, 6)) -> Tuple[plt.Figure, plt.Axes]:
        """Create a new figure with dark theme"""
        fig, ax = plt.subplots(figsize=figsize, facecolor='#1a1a2e')
        ax.set_facecolor('#16213e')
        
        # Style axes
        ax.spines['bottom'].set_color('#4a4a6a')
        ax.spines['top'].set_color('#4a4a6a')
        ax.spines['left'].set_color('#4a4a6a')
        ax.spines['right'].set_color('#4a4a6a')
        ax.tick_params(colors='#e0e0e0')
        ax.xaxis.label.set_color('#e0e0e0')
        ax.yaxis.label.set_color('#e0e0e0')
        ax.title.set_color('#ffffff')
        
        return fig, ax
    
    def _save_chart(self, fig: plt.Figure, name: str) -> BytesIO:
        """Save chart to BytesIO buffer"""
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', 
                    facecolor=fig.get_facecolor(), edgecolor='none')
        buf.seek(0)
        plt.close(fig)
        
        self._charts[name] = buf
        
        # Also save to file if output dir is set
        if self.output_dir:
            file_path = self.output_dir / f"{name}.png"
            with open(file_path, 'wb') as f:
                f.write(buf.getvalue())
            buf.seek(0)
        
        return buf
    
    def get_chart(self, name: str) -> Optional[BytesIO]:
        """Get a previously generated chart"""
        chart = self._charts.get(name)
        if chart:
            chart.seek(0)
        return chart
    
    def generate_all_charts(self, sections: List[str]) -> Dict[str, BytesIO]:
        """Generate all requested chart sections"""
        chart_methods = {
            "yearly_open": self.chart_yearly_open_vulnerabilities,
            "priority_dist": self.chart_priority_distribution,
            "line_manager": self.chart_by_line_manager,
            "department": self.chart_by_department,
            "tool": self.chart_by_tool,
            "sla": self.chart_sla_status,
            "trend": self.chart_trend_analysis,
            "top10": self.chart_top_vulnerabilities,
            "ip_density": self.chart_ip_density,
            "resolution_time": self.chart_resolution_time,
            "sla_breach": self.chart_sla_breach_analysis,
        }
        
        results = {}
        for section in sections:
            if section in chart_methods:
                try:
                    buf = chart_methods[section]()
                    if buf:
                        results[section] = buf
                except Exception as e:
                    print(f"Error generating chart {section}: {e}")
        
        return results
    
    def chart_yearly_open_vulnerabilities(self) -> BytesIO:
        """Generate bar chart of open vulnerabilities by year"""
        data = self.filter_engine.get_count_by_year()
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("Yıllara Göre Veri Bulunamadı", "No Data by Year")
            )
        
        fig, ax = self._create_figure()
        
        # Create gradient-like bars
        colors = plt.cm.viridis(np.linspace(0.3, 0.9, len(data)))
        bars = ax.bar(data.index.astype(str), data.values, color=colors, edgecolor='white', linewidth=0.5)
        
        # Add value labels on bars
        for bar, val in zip(bars, data.values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(data.values)*0.02,
                   str(int(val)), ha='center', va='bottom', color='white', fontweight='bold')
        
        ax.set_xlabel(self._get_label("Yıl", "Year"), fontsize=12, fontweight='bold')
        ax.set_ylabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("📊 Yıllara Göre Açık Zafiyet Sayısı", "📊 Open Vulnerabilities by Year"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "yearly_open")
    
    def chart_priority_distribution(self) -> BytesIO:
        """Generate pie chart of priority distribution"""
        data = self.filter_engine.get_priority_summary()
        
        if not data or sum(data.values()) == 0:
            return self._create_empty_chart(
                self._get_label("Priority Verisi Bulunamadı", "No Priority Data")
            )
        
        fig, ax = self._create_figure(figsize=(8, 8))
        
        labels = list(data.keys())
        sizes = list(data.values())
        colors = [self.PRIORITY_COLORS.get(label, '#888888') for label in labels]
        
        # Create pie chart
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, colors=colors, autopct='%1.1f%%',
            startangle=90, explode=[0.02]*len(sizes),
            textprops={'color': 'white', 'fontweight': 'bold'}
        )
        
        # Style autotexts
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        ax.set_title(self._get_label("🎯 Priority Dağılımı", "🎯 Priority Distribution"),
                    fontsize=14, fontweight='bold', pad=20, color='white')
        
        # Add legend
        ax.legend(wedges, [f"{l}: {s}" for l, s in zip(labels, sizes)],
                 loc="lower right", facecolor='#1a1a2e', edgecolor='#4a4a6a',
                 labelcolor='white')
        
        return self._save_chart(fig, "priority_dist")
    
    def chart_by_line_manager(self) -> BytesIO:
        """Generate horizontal bar chart by line manager"""
        data = self.filter_engine.get_count_by_column("line_manager").head(15)
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("Line Manager Verisi Bulunamadı", "No Line Manager Data")
            )
        
        fig, ax = self._create_figure(figsize=(12, 8))
        
        # Sort and create horizontal bars
        data_sorted = data.sort_values(ascending=True)
        colors = plt.cm.plasma(np.linspace(0.2, 0.8, len(data_sorted)))
        
        bars = ax.barh(range(len(data_sorted)), data_sorted.values, color=colors, edgecolor='white', linewidth=0.5)
        ax.set_yticks(range(len(data_sorted)))
        ax.set_yticklabels([str(x)[:30] + '...' if len(str(x)) > 30 else str(x) for x in data_sorted.index],
                          color='white')
        
        # Add value labels
        for bar, val in zip(bars, data_sorted.values):
            ax.text(bar.get_width() + max(data_sorted.values)*0.02, bar.get_y() + bar.get_height()/2,
                   str(int(val)), ha='left', va='center', color='white', fontweight='bold')
        
        ax.set_xlabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("👥 Line Manager Bazında Zafiyet Dağılımı", "👥 Vulnerabilities by Line Manager"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "line_manager")
    
    def chart_by_department(self) -> BytesIO:
        """Generate horizontal bar chart by department"""
        data = self.filter_engine.get_count_by_column("department").head(15)
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("Departman Verisi Bulunamadı", "No Department Data")
            )
        
        fig, ax = self._create_figure(figsize=(12, 8))
        
        data_sorted = data.sort_values(ascending=True)
        colors = plt.cm.cool(np.linspace(0.2, 0.8, len(data_sorted)))
        
        bars = ax.barh(range(len(data_sorted)), data_sorted.values, color=colors, edgecolor='white', linewidth=0.5)
        ax.set_yticks(range(len(data_sorted)))
        ax.set_yticklabels([str(x)[:35] + '...' if len(str(x)) > 35 else str(x) for x in data_sorted.index],
                          color='white')
        
        for bar, val in zip(bars, data_sorted.values):
            ax.text(bar.get_width() + max(data_sorted.values)*0.02, bar.get_y() + bar.get_height()/2,
                   str(int(val)), ha='left', va='center', color='white', fontweight='bold')
        
        ax.set_xlabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("🏢 Departman Bazında Zafiyet Dağılımı", "🏢 Vulnerabilities by Department"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "department")
    
    def chart_by_tool(self) -> BytesIO:
        """Generate pie chart by tool/source"""
        data = self.filter_engine.get_count_by_column("tool")
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("Tool Verisi Bulunamadı", "No Tool Data")
            )
        
        fig, ax = self._create_figure(figsize=(10, 8))
        
        # Limit to top 10 for readability
        if len(data) > 10:
            top_data = data.head(9)
            other_sum = data[9:].sum()
            data = pd.concat([top_data, pd.Series({'Other': other_sum})])
        
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        
        wedges, texts, autotexts = ax.pie(
            data.values, labels=None, colors=colors, autopct='%1.1f%%',
            startangle=90, pctdistance=0.8,
            textprops={'color': 'white', 'fontweight': 'bold', 'fontsize': 9}
        )
        
        # Add legend outside
        ax.legend(wedges, [f"{l[:20]}..." if len(str(l)) > 20 else str(l) for l in data.index],
                 loc="center left", bbox_to_anchor=(1, 0.5),
                 facecolor='#1a1a2e', edgecolor='#4a4a6a', labelcolor='white')
        
        ax.set_title(self._get_label("🔧 Tool/Kaynak Dağılımı", "🔧 Tool/Source Distribution"),
                    fontsize=14, fontweight='bold', pad=20, color='white')
        
        plt.tight_layout()
        return self._save_chart(fig, "tool")
    
    def chart_sla_status(self) -> BytesIO:
        """Generate SLA status donut chart"""
        data = self.filter_engine.get_sla_summary()
        
        if sum(data.values()) == 0:
            return self._create_empty_chart(
                self._get_label("SLA Verisi Bulunamadı", "No SLA Data")
            )
        
        fig, ax = self._create_figure(figsize=(8, 8))
        
        labels_tr = {"in_sla": "SLA İçinde", "out_of_sla": "SLA Dışında", "unknown": "Bilinmiyor"}
        labels_en = {"in_sla": "In SLA", "out_of_sla": "Out of SLA", "unknown": "Unknown"}
        labels_map = labels_tr if self.language == "TR" else labels_en
        
        # Filter out zero values
        filtered_data = {k: v for k, v in data.items() if v > 0}
        
        labels = [labels_map[k] for k in filtered_data.keys()]
        sizes = list(filtered_data.values())
        colors = [self.SLA_COLORS[k] for k in filtered_data.keys()]
        
        # Create donut chart
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, colors=colors, autopct='%1.1f%%',
            startangle=90, pctdistance=0.75,
            wedgeprops=dict(width=0.5, edgecolor='#1a1a2e'),
            textprops={'color': 'white', 'fontweight': 'bold'}
        )
        
        # Add center text
        total = sum(sizes)
        out_of_sla = data.get("out_of_sla", 0)
        percentage = (out_of_sla / total * 100) if total > 0 else 0
        
        ax.text(0, 0, f"{percentage:.1f}%\n{self._get_label('SLA Dışı', 'Out of SLA')}",
               ha='center', va='center', fontsize=16, fontweight='bold', color='#FF4757')
        
        ax.set_title(self._get_label("⏰ SLA Durumu", "⏰ SLA Status"),
                    fontsize=14, fontweight='bold', pad=20, color='white')
        
        return self._save_chart(fig, "sla")
    
    def chart_trend_analysis(self) -> BytesIO:
        """Generate trend line chart by month"""
        data = self.filter_engine.get_count_by_month()
        
        if data.empty or len(data) < 2:
            return self._create_empty_chart(
                self._get_label("Trend Verisi Yetersiz", "Insufficient Trend Data")
            )
        
        fig, ax = self._create_figure(figsize=(14, 6))
        
        x_labels = [str(p) for p in data.index]
        
        # Plot line with markers
        ax.plot(range(len(data)), data.values, color=self.COLORS["primary"],
               linewidth=2.5, marker='o', markersize=6, markerfacecolor='white',
               markeredgecolor=self.COLORS["primary"], markeredgewidth=2)
        
        # Fill under the line
        ax.fill_between(range(len(data)), data.values, alpha=0.3, color=self.COLORS["primary"])
        
        # Set x-axis labels (show every nth label if too many)
        step = max(1, len(x_labels) // 12)
        ax.set_xticks(range(0, len(x_labels), step))
        ax.set_xticklabels([x_labels[i] for i in range(0, len(x_labels), step)],
                          rotation=45, ha='right', color='white')
        
        ax.set_xlabel(self._get_label("Tarih", "Date"), fontsize=12, fontweight='bold')
        ax.set_ylabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("📈 Zafiyet Trend Analizi", "📈 Vulnerability Trend Analysis"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "trend")
    
    def chart_top_vulnerabilities(self) -> BytesIO:
        """Generate horizontal bar chart of top 10 vulnerabilities"""
        data = self.filter_engine.get_top_vulnerabilities(10)
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("Zafiyet Verisi Bulunamadı", "No Vulnerability Data")
            )
        
        fig, ax = self._create_figure(figsize=(14, 8))
        
        # Truncate long names
        labels = [str(x)[:50] + '...' if len(str(x)) > 50 else str(x) for x in data['vulnerability']]
        values = data['count'].values
        
        # Create gradient colors based on count
        colors = plt.cm.Reds(np.linspace(0.4, 0.9, len(values)))[::-1]
        
        bars = ax.barh(range(len(labels)), values, color=colors, edgecolor='white', linewidth=0.5)
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels[::-1], color='white', fontsize=9)
        
        # Reverse order so highest is at top
        bars_reversed = list(reversed(list(bars)))
        values_reversed = list(reversed(values))
        
        for bar, val in zip(bars, values[::-1]):
            ax.text(bar.get_width() + max(values)*0.02, bar.get_y() + bar.get_height()/2,
                   str(int(val)), ha='left', va='center', color='white', fontweight='bold')
        
        ax.set_xlabel(self._get_label("Tespit Sayısı", "Detection Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("🔥 En Çok Görülen 10 Zafiyet", "🔥 Top 10 Most Common Vulnerabilities"),
                    fontsize=14, fontweight='bold', pad=20)
        
        ax.invert_yaxis()
        plt.tight_layout()
        return self._save_chart(fig, "top10")
    
    def chart_ip_density(self) -> BytesIO:
        """Generate bar chart of vulnerability density by IP"""
        data = self.filter_engine.get_ip_density(10)
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("IP Verisi Bulunamadı", "No IP Data")
            )
        
        fig, ax = self._create_figure(figsize=(12, 6))
        
        colors = plt.cm.magma(np.linspace(0.3, 0.8, len(data)))
        
        bars = ax.bar(range(len(data)), data.values, color=colors, edgecolor='white', linewidth=0.5)
        ax.set_xticks(range(len(data)))
        ax.set_xticklabels(data.index, rotation=45, ha='right', color='white', fontsize=9)
        
        for bar, val in zip(bars, data.values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(data.values)*0.02,
                   str(int(val)), ha='center', va='bottom', color='white', fontweight='bold')
        
        ax.set_xlabel(self._get_label("IP Adresi", "IP Address"), fontsize=12, fontweight='bold')
        ax.set_ylabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("💻 IP Bazlı Zafiyet Yoğunluğu (Top 10)", "💻 Vulnerability Density by IP (Top 10)"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "ip_density")
    
    def chart_resolution_time(self) -> BytesIO:
        """Generate resolution time statistics chart"""
        stats = self.filter_engine.get_resolution_time_stats()
        
        if stats["mean"] == 0 and stats["max"] == 0:
            return self._create_empty_chart(
                self._get_label("Çözüm Süresi Verisi Bulunamadı", "No Resolution Time Data")
            )
        
        fig, ax = self._create_figure(figsize=(10, 6))
        
        labels = [
            self._get_label("Ortalama", "Average"),
            self._get_label("Medyan", "Median"),
            self._get_label("Minimum", "Minimum"),
            self._get_label("Maksimum", "Maximum")
        ]
        values = [stats["mean"], stats["median"], stats["min"], stats["max"]]
        colors = [self.COLORS["primary"], self.COLORS["secondary"], 
                  self.COLORS["success"], self.COLORS["warning"]]
        
        bars = ax.bar(labels, values, color=colors, edgecolor='white', linewidth=2)
        
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.02,
                   f"{val:.1f}", ha='center', va='bottom', color='white', fontweight='bold', fontsize=12)
        
        ax.set_ylabel(self._get_label("Gün", "Days"), fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("📅 Zafiyet Çözüm Süreleri", "📅 Vulnerability Resolution Times"),
                    fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._save_chart(fig, "resolution_time")
    
    def chart_sla_breach_analysis(self) -> BytesIO:
        """Generate histogram of SLA breach days"""
        data = self.filter_engine.get_sla_breach_distribution()
        
        if data.empty:
            return self._create_empty_chart(
                self._get_label("SLA Aşım Verisi Bulunamadı", "No SLA Breach Data")
            )
        
        fig, ax = self._create_figure(figsize=(12, 6))
        
        # Create histogram
        n, bins, patches = ax.hist(data, bins=20, color=self.COLORS["danger"],
                                   edgecolor='white', linewidth=0.5, alpha=0.8)
        
        # Color gradient based on breach severity
        max_val = max(n) if len(n) > 0 else 1
        for i, patch in enumerate(patches):
            intensity = n[i] / max_val
            patch.set_facecolor(plt.cm.Reds(0.3 + 0.6 * intensity))
        
        ax.set_xlabel(self._get_label("SLA Aşım Süresi (Gün)", "SLA Breach Duration (Days)"), 
                     fontsize=12, fontweight='bold')
        ax.set_ylabel(self._get_label("Zafiyet Sayısı", "Vulnerability Count"), 
                     fontsize=12, fontweight='bold')
        ax.set_title(self._get_label("⚠️ SLA Aşım Analizi", "⚠️ SLA Breach Analysis"),
                    fontsize=14, fontweight='bold', pad=20)
        
        # Add summary stats
        mean_breach = data.mean()
        median_breach = data.median()
        stats_text = f"{self._get_label('Ort:', 'Avg:')} {mean_breach:.1f} | {self._get_label('Medyan:', 'Median:')} {median_breach:.1f}"
        ax.text(0.95, 0.95, stats_text, transform=ax.transAxes, ha='right', va='top',
               fontsize=11, color='white', fontweight='bold',
               bbox=dict(boxstyle='round', facecolor='#2C3E50', alpha=0.8))
        
        plt.tight_layout()
        return self._save_chart(fig, "sla_breach")
    
    def _create_empty_chart(self, message: str) -> BytesIO:
        """Create a placeholder chart with a message"""
        fig, ax = self._create_figure(figsize=(8, 6))
        
        ax.text(0.5, 0.5, message, ha='center', va='center',
               fontsize=16, color='#888888', transform=ax.transAxes)
        ax.set_xticks([])
        ax.set_yticks([])
        
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                    facecolor=fig.get_facecolor())
        buf.seek(0)
        plt.close(fig)
        
        return buf
