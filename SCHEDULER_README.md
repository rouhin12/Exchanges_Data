# Exchange Data Download Scheduler

Automated daily scheduler for downloading Indian exchange data (NSE & BSE) every morning at 9:00 AM.

## Features

✓ **NSE F&O Segment**: Daily F&O data with notional turnover columns  
✓ **NSE CM Segment**: Daily cash market segment data  
✓ **BSE F&O**: T-2 and T-1 daily data (all 4 derivatives)  
✓ **BSE Historical Equity**: Current month equity data  

✓ Concurrent downloads where possible  
✓ Detailed logging with daily log files  
✓ Windows Task Scheduler integration  
✓ Test mode for immediate downloads  

## Installation

1. **Install Dependencies**
   ```bash
   pip install schedule selenium openpyxl
   ```

2. **Test the Scheduler**
   ```bash
   # Run all downloads immediately (test)
   python scheduler.py --test
   
   # Run NSE only (test)
   python scheduler.py --test --nse-only
   ```

## Usage

### Option 1: Windows Batch Setup (Recommended)

Double-click `setup_scheduler.bat` and follow the menu:
- Option 1: Test run (download all data now)
- Option 2: Test NSE only
- Option 3: Instructions for Windows Task Scheduler
- Option 4: Run scheduler in foreground

### Option 2: Manual Python Execution

```bash
# Run scheduler at 9 AM daily (foreground)
python scheduler.py

# Custom time (e.g., 8:30 AM)
python scheduler.py --time 08:30

# Test mode (run once immediately)
python scheduler.py --test

# NSE only mode
python scheduler.py --nse-only
```

### Option 3: Windows Task Scheduler (Production)

1. Open **Task Scheduler** (Win+R → `taskschd.msc`)
2. Click **"Create Basic Task"**
3. Configure:
   - **Name**: Exchange Data Daily Download
   - **Trigger**: Daily
   - **Start Time**: 9:00 AM
   - **Action**: Start a program
   - **Program**: `C:\path\to\.venv\Scripts\python.exe`
   - **Arguments**: `C:\path\to\scheduler.py`
   - **Start in**: `C:\path\to\Exchanges Data`

4. Optional: Under **Settings** tab:
   - ✓ Run whether user is logged on or not
   - ✓ Run with highest privileges (if needed)

## What Gets Downloaded

### Daily (Default Behavior)

| Downloader | What | Time Range | Output |
|-----------|------|------------|---------|
| **NSE F&O** | F&O segment daily data | Current month | `NSE/FnO/nse_fno_FY{year}_{year}.xlsx` |
| **NSE CM** | Cash market segment | Current month | `NSE/nse_daily_FY{year}_{year}.xlsx` |
| **BSE F&O** | All 4 derivatives | T-2 & T-1 | `BSE/FnO/{type}/{derivative}_FY{year}_{year}.xlsx` |
| **BSE Equity** | Historical equity | Current month | `BSE/bse_daily_FY{year}_{year}.xlsx` |

### Data Organization

- **Financial Year**: April-March convention (FY2025_2026 = Apr 2025 - Mar 2026)
- **Monthly Sheets**: Each workbook has sheets named YYYY-MM (e.g., "2026-01" for January 2026)
- **Automatic Consolidation**: Data appends to existing sheets

## Logs

All execution logs are saved to `logs/` directory:
- **Filename**: `scheduler_YYYY-MM-DD.log`
- **Includes**: Start time, download status, errors, duration

Example log:
```
2026-02-12 09:00:01 | INFO     | Starting: NSE F&O Daily Data
2026-02-12 09:00:45 | INFO     | ✓ Completed: NSE F&O Daily Data
2026-02-12 09:00:45 | INFO     | Starting: NSE CM Segment Daily Data
2026-02-12 09:01:12 | INFO     | ✓ Completed: NSE CM Segment Daily Data
```

## Troubleshooting

### Issue: "schedule package not installed"
```bash
pip install schedule
```

### Issue: Downloads fail
1. Check logs in `logs/` directory
2. Run test mode: `python scheduler.py --test`
3. Check individual downloaders work:
   ```bash
   python NSE/FnO/nse_fno_download.py
   python BSE/FnO/bse_fno_download.py
   ```

### Issue: Task Scheduler not running
- Ensure "Start in" directory is set correctly
- Check user has permissions to write to output directories
- Review Task Scheduler history for errors

## Advanced Usage

### Download Historical Data

Each downloader supports custom date ranges:

```bash
# NSE F&O - specific months
python NSE/FnO/nse_fno_download.py --start-month Dec-2025 --end-month Feb-2026

# NSE F&O - financial years
python NSE/FnO/nse_fno_download.py --start-year 2021 --end-year 2025

# BSE F&O - date range with specific derivatives
python BSE/FnO/bse_fno_download.py --from-date 01/01/2026 --to-date 31/01/2026 --derivatives index-futures

# NSE CM - specific months
python NSE/nse_business_growth_cm_download.py --start-month 2025-12 --end-month 2026-02

# BSE Equity - specific months
python BSE/bse_historical_equity_download.py --start-month 2025-12 --end-month 2026-02
```

### Modify Schedule Time

Edit `scheduler.py` line 218:
```python
schedule_daily(hour=9, minute=0)  # Change to desired time
```

Or use command line:
```bash
python scheduler.py --time 08:30
```

## File Structure

```
Exchanges Data/
├── scheduler.py              # Main scheduler script
├── setup_scheduler.bat       # Windows setup utility
├── logs/                     # Daily log files
│   └── scheduler_2026-02-12.log
├── NSE/
│   ├── FnO/
│   │   ├── nse_fno_download.py
│   │   └── nse_fno_FY2025_2026.xlsx
│   ├── nse_business_growth_cm_download.py
│   └── nse_daily_FY2025_2026.xlsx
└── BSE/
    ├── FnO/
    │   ├── bse_fno_download.py
    │   ├── Index/
    │   │   ├── Futures/
    │   │   └── Options/
    │   └── Equity/
    │       ├── Futures/
    │       └── Options/
    ├── bse_historical_equity_download.py
    └── bse_daily_FY2025_2026.xlsx
```

## Notes

- **Default Behavior**: Scheduler downloads current month data (T-1/T-2 for BSE F&O)
- **Runtime**: Typically 3-5 minutes for all downloaders
- **Network**: Requires stable internet connection
- **Browser**: Edge WebDriver required for BSE/NSE F&O downloads
- **Excel Files**: Auto-creates and updates workbooks organized by financial year

## Support

For issues or questions, check the individual downloader scripts for detailed usage and options.
