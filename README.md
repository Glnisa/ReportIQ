# ReportIQ - Vulnerability Report Generator

🛡️ **ReportIQ** is a Python desktop application that automates vulnerability report generation from Excel data.

## Features

- 📂 **Excel Import**: Load vulnerability data from Excel files with automatic column mapping
- 🔍 **Smart Filtering**: Filter by SLA status, priority, tool/source, year, department, and more
- 📊 **Visual Analytics**: Generate 11 different chart types for comprehensive analysis
- 📝 **Word Export**: Create professional Word documents with embedded charts and tables
- 🌐 **Bilingual**: Full Turkish and English language support
- 🎨 **Modern UI**: Beautiful dark theme with cybersecurity aesthetics

## Installation

### Prerequisites
- Python 3.9 or higher

### Setup

1. Clone or download this repository
2. Navigate to the project directory:
   ```bash
   cd ReportIQ
   ```

3. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On macOS/Linux
   # or
   venv\Scripts\activate  # On Windows
   ```

4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. Click **"Gözat" / "Browse"** to select your Excel file

3. Configure filters:
   - Select SLA status (In SLA / Out of SLA)
   - Choose vulnerability statuses (PENDING, QUEUED, etc.)
   - Filter by tool/source, year, department

4. Select report sections to include (charts and analyses)

5. Click **"🚀 Rapor Oluştur" / "🚀 Generate Report"**

6. Choose where to save the Word document

## Excel File Format

The application expects an Excel file with vulnerability data. It will automatically detect columns such as:

| Column | Description |
|--------|-------------|
| TICKETID | Unique ticket identifier |
| REPORTEDPRIORITY | Priority level (High, Critical) |
| SLA_Value | SLA status (In SLA, Out of SLA) |
| STATUS | Ticket status (PENDING, QUEUED, CLOSED, etc.) |
| Day_of_CREATIONDATE | Vulnerability creation date |
| Department | Department name |
| Line Manager | Manager name |
| PLUGINID | Tenable Plugin ID |
| PLUGINDESC | Vulnerability name/description |
| TOOL | Scan source (TenableSC, NessusAgent, etc.) |
| IP | Target IP address |
| PORT | Target port |

## Report Sections

The following report sections are available:

1. 📊 **Yearly Open Vulnerabilities** - Bar chart by year
2. 🎯 **Priority Distribution** - Pie chart of High vs Critical
3. 👥 **Line Manager Breakdown** - Horizontal bar by manager
4. 🏢 **Department Breakdown** - Horizontal bar by department
5. 🔧 **Tool Distribution** - Pie chart by scan source
6. ⏰ **SLA Status** - Donut chart of SLA compliance
7. 📈 **Trend Analysis** - Line chart over time
8. 🔥 **Top 10 Vulnerabilities** - Most common vulnerabilities with descriptions
9. 💻 **IP Density** - Vulnerabilities by IP address
10. 📅 **Resolution Time** - Average time to close vulnerabilities
11. ⚠️ **SLA Breach Analysis** - Distribution of SLA overdue days

## Project Structure

```
ReportIQ/
├── main.py                     # Application entry point
├── requirements.txt            # Python dependencies
├── README.md                   # This file
│
├── src/
│   ├── core/
│   │   ├── data_loader.py      # Excel loading & parsing
│   │   ├── filter_engine.py    # Data filtering logic
│   │   ├── chart_generator.py  # Matplotlib chart generation
│   │   └── word_generator.py   # Word document creation
│   │
│   ├── gui/
│   │   ├── main_window.py      # Main application window
│   │   └── styles.py           # UI styling constants
│   │
│   └── data/
│       ├── translations.py     # TR/EN language strings
│       └── vulnerability_dict.py  # Vulnerability definitions
│
└── output/                     # Generated reports
```

## Technologies

- **CustomTkinter** - Modern GUI framework
- **pandas** - Data manipulation
- **matplotlib / seaborn** - Chart generation
- **python-docx** - Word document creation

## License

MIT License

## Author

Created with ReportIQ 🛡️
