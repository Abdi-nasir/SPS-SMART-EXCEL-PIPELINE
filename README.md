#  SPS Excel Pipeline - Enterprise Data Dashboard

An intelligent Excel data processing pipeline that automatically handles multiple Excel files with multiple sheets, provides professional KPIs, interactive visualizations, and executive-level dashboard reporting.

##  Features

### Authentication & Security
- **Secure Login** - Username/password [admin/admin123] authentication for dashboard access
- **Session Management** - Cookie-based session handling with configurable expiry
- **Logout Functionality** - Secure logout option
- **Configurable Users** - Easy user management via YAML configuration

###  File Management
- **Multi-File Support** - Upload and manage multiple Excel files simultaneously
- **Multi-Sheet Support** - Each file can have multiple sheets, displayed independently
- **Smart Processing** - Automatically detects file structures and decides whether to combine or keep separate
- **Auto-Processing** - Pipeline runs automatically after file upload

###  Dashboard & Visualization
- **Professional KPIs** - Key Performance Indicators with trend indicators and color-coded metrics
- **Interactive Charts** - Bar charts, line charts, pie charts, scatter plots, histograms, and box plots
- **Time Series Analysis** - Trend lines with moving averages and CAGR calculation
- **Comparative Analysis** - Top N categories with distribution visualization
- **Data Quality Alerts** - Automatic detection of outliers, missing data, and anomalies

###  Download & Export
- **Full Dashboard PDF** - Download complete dashboard with KPIs and all charts in full color
- **Charts Report** - Generate PDF reports containing only visualizations
- **Individual Charts** - Download any chart as PNG

###  Data Processing
- **External Data Fetch** - Pull data from external APIs directly into the pipeline
- **Statistical Summary** - Count, mean, std, min, quartiles, and max for numeric columns

###  Deployment
- **Docker Support** - Ready-to-use Docker and Docker Compose configurations
- **Volume Persistence** - Data persistence across container restarts
- **Production Ready** - Nginx reverse proxy and HTTPS support

##  Quick Start

### Prerequisites
- Python 3.8+ or Docker

### Option 1: Run with Python (No Docker)

1. **Clone the repository**
   ```
   git clone https://github.com/Abdi-nasir/SPS-SMART-EXCEL-PIPELINE.git
   cd sps-excel-pipeline
    ```

### Option 2: Run with Docker

1. **Clone and run**
    ```
    docker-compose up -d
    ```

2. **Access the dashboard**
    ```
    Open http://localhost:8501
    ```


### Folders Auto-Created

    ```

    # App automatically creates:
    📁 Created directory: /path/to/project/data
    📁 Created directory: /path/to/project/data/raw
    📁 Created directory: /path/to/project/data/processed
    📁 Created directory: /path/to/project/reports
    📁 Created directory: /path/to/project/logs

    # Now you have this structure:
    sps-excel-pipeline/
    ├── data/
    │   ├── raw/
    │   └── processed/
    ├── reports/
    ├── src/
    │   └── app.py
    └── ...




    

