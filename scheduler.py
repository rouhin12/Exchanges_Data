#!/usr/bin/env python3
"""
Daily Scheduler for Exchange Data Downloads

Runs every day at 9:00 AM and executes:
1. NSE F&O Segment (Daily)
2. NSE CM Segment (Daily)
3. BSE F&O (T-2 and T-1)
4. BSE Historical Equity (Current Month)

All downloads are concurrent where possible to minimize runtime.
Logs are saved to logs/ directory with timestamp.

Installation:
    pip install schedule

Usage:
    python scheduler.py                    # Run scheduler in foreground
    python scheduler.py --daemon           # Run as background/daemon process
    python scheduler.py --test             # Run all downloads once immediately
    python scheduler.py --nse-only         # Run only NSE downloaders

Scheduling (Windows Task Scheduler):
    1. Open Task Scheduler
    2. Create Basic Task > Daily > 9:00 AM
    3. Action: Start a program
       Program: C:\\path\\to\\.venv\\Scripts\\python.exe
       Arguments: C:\\path\\to\\scheduler.py
"""

import argparse
import datetime as dt
import logging
import os
import subprocess
import sys
import time
from pathlib import Path

try:
    import schedule
except ImportError:
    print("ERROR: 'schedule' package not installed. Run: pip install schedule")
    sys.exit(1)


# Setup logging
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / f"scheduler_{dt.date.today().strftime('%Y-%m-%d')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# File paths
BASE_DIR = Path(__file__).parent
NSE_FNO = BASE_DIR / "NSE" / "nse_fno_download.py"
NSE_CM = BASE_DIR / "NSE" / "nse_business_growth_cm_download.py"
BSE_FNO = BASE_DIR / "BSE" / "bse_fno_consolidated_download.py"
BSE_EQUITY = BASE_DIR / "BSE" / "bse_historical_equity_download.py"
FII_DII_SCRAPER = BASE_DIR / "FII_DII" / "fii_dii_download.py"
EXCEL_BUILDER = BASE_DIR / "automation" / "build_exchange_database.py"

# Output directories
NSE_FNO_DIR = BASE_DIR / "NSE" / "FnO"
NSE_CM_DIR = BASE_DIR / "NSE" / "Cash segment"  # NSE Cash/equity data goes here
BSE_EQUITY_DIR = BASE_DIR / "BSE"


def get_date_ranges():
    """Calculate last month and current month for downloads."""
    today = dt.date.today()
    
    # Current month
    current_month_start = today.replace(day=1)
    
    # Last month
    last_month_end = current_month_start - dt.timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    
    return {
        # NSE F&O format: "Jan-2026"
        'nse_fno_start': last_month_start.strftime("%b-%Y"),
        'nse_fno_end': today.strftime("%b-%Y"),
        
        # NSE CM format: "2026-01"
        'nse_cm_start': last_month_start.strftime("%Y-%m"),
        'nse_cm_end': today.strftime("%Y-%m"),
        
        # BSE F&O format: "Jan-2026" (matching consolidated downloader)
        'bse_fno_start': last_month_start.strftime("%b-%Y"),
        'bse_fno_end': today.strftime("%b-%Y"),
        
        # BSE Equity format: "2026-01"
        'bse_equity_start': last_month_start.strftime("%Y-%m"),
        'bse_equity_end': today.strftime("%Y-%m"),
    }

PYTHON_EXE = sys.executable


def run_command(cmd, description):
    """Run a command and log output."""
    logger.info(f"Starting: {description}")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes timeout
        )
        
        # Log output
        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                logger.info(f"  {line}")
        
        if result.returncode == 0:
            logger.info(f"[OK] Completed: {description}")
            return True
        else:
            logger.error(f"[X] Failed: {description}")
            if result.stderr:
                for line in result.stderr.strip().split("\n"):
                    logger.error(f"  {line}")
            return False
    
    except subprocess.TimeoutExpired:
        logger.error(f"[X] Timeout: {description} (exceeded 10 minutes)")
        return False
    except Exception as e:
        logger.error(f"[X] Error running {description}: {e}")
        return False


def download_nse_fno():
    """Download NSE F&O daily data (last month + current month)."""
    if not NSE_FNO.exists():
        logger.error(f"File not found: {NSE_FNO}")
        return False
    
    dates = get_date_ranges()
    cmd = [
        PYTHON_EXE, str(NSE_FNO),
        "--start-month", dates['nse_fno_start'],
        "--end-month", dates['nse_fno_end'],
        "--output-dir", str(NSE_FNO_DIR)
    ]
    return run_command(cmd, "NSE F&O Daily Data")


def download_nse_cm():
    """Download NSE CM segment daily data (last month + current month)."""
    if not NSE_CM.exists():
        logger.error(f"File not found: {NSE_CM}")
        return False
    
    dates = get_date_ranges()
    cmd = [
        PYTHON_EXE, str(NSE_CM),
        "--start-month", dates['nse_cm_start'],
        "--end-month", dates['nse_cm_end'],
        "--output-dir", str(NSE_CM_DIR)
    ]
    return run_command(cmd, "NSE CM Segment Daily Data")


def download_bse_fno():
    """Download consolidated BSE F&O data (last month + current month, all derivatives in single file)."""
    if not BSE_FNO.exists():
        logger.error(f"File not found: {BSE_FNO}")
        return False
    
    dates = get_date_ranges()
    cmd = [
        PYTHON_EXE, str(BSE_FNO),
        "--start-month", dates['bse_fno_start'],
        "--end-month", dates['bse_fno_end'],
        "--output-dir", "BSE/FnO"
    ]
    return run_command(cmd, "BSE F&O Data (Consolidated - All Derivatives)")


def download_bse_equity():
    """Download BSE Historical Equity (last month + current month)."""
    if not BSE_EQUITY.exists():
        logger.error(f"File not found: {BSE_EQUITY}")
        return False
    
    dates = get_date_ranges()
    cmd = [
        PYTHON_EXE, str(BSE_EQUITY),
        "--start-month", dates['bse_equity_start'],
        "--end-month", dates['bse_equity_end'],
        "--output-dir", str(BSE_EQUITY_DIR)
    ]
    return run_command(cmd, "BSE Historical Equity Data")


def download_fii_dii():
    """Download FII/DII data from Moneycontrol."""
    if not FII_DII_SCRAPER.exists():
        logger.error(f"File not found: {FII_DII_SCRAPER}")
        return False

    cmd = [PYTHON_EXE, str(FII_DII_SCRAPER)]
    return run_command(cmd, "FII/DII Data from Moneycontrol")


def build_exchange_database():
    """Build the automated Exchanges Database workbook."""
    if not EXCEL_BUILDER.exists():
        logger.error(f"File not found: {EXCEL_BUILDER}")
        return False

    cmd = [
        PYTHON_EXE, str(EXCEL_BUILDER)
    ]
    return run_command(cmd, "Exchanges Database Workbook")


def run_all_downloads(nse_only=False):
    """Run all downloads and log results."""
    logger.info("=" * 70)
    logger.info("DAILY DOWNLOAD SCHEDULER")
    logger.info(f"Time: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    results = {}
    start_time = time.time()
    
    try:
        # NSE Downloads (usually faster)
        logger.info("\n>> Starting NSE downloads...")
        results["NSE F&O"] = download_nse_fno()
        results["NSE CM"] = download_nse_cm()
        
        if nse_only:
            logger.info("\n(NSE-only mode: Skipping BSE downloads)")
        else:
            # BSE Downloads
            logger.info("\n>> Starting BSE downloads...")
            results["BSE F&O"] = download_bse_fno()
            results["BSE Equity"] = download_bse_equity()

        # FII/DII Data
        logger.info("\n>> Downloading FII/DII data...")
        results["FII/DII Data"] = download_fii_dii()

        # Build Excel workbook after data downloads
        # Note: Requires automation/build_exchange_database.py to exist
        if EXCEL_BUILDER.exists():
            results["Excel Workbook"] = build_exchange_database()
        else:
            logger.info("Skipping Excel Workbook (script not found)")
        
    except Exception as e:
        logger.error(f"Scheduler error: {e}")
        import traceback
        traceback.print_exc()
    
    # Summary
    elapsed = int(time.time() - start_time)
    successful = sum(1 for v in results.values() if v)
    total = len(results)
    
    logger.info("\n" + "=" * 70)
    logger.info("SUMMARY")
    logger.info("=" * 70)
    for name, success in results.items():
        status = "[OK]" if success else "[FAILED]"
        logger.info(f"  {name}: {status}")
    
    logger.info(f"\nTotal: {successful}/{total} successful")
    logger.info(f"Duration: {elapsed} seconds")
    logger.info("=" * 70 + "\n")


def schedule_daily(hour=9, minute=0):
    """Schedule downloads to run daily at specified time."""
    schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(run_all_downloads)
    logger.info(f"Scheduler configured to run at {hour:02d}:{minute:02d} daily")


def run_scheduler(nse_only=False):
    """Main scheduler loop."""
    logger.info("Scheduler started. Press Ctrl+C to stop.\n")
    
    schedule_daily(hour=9, minute=0)
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    except KeyboardInterrupt:
        logger.info("\nScheduler stopped by user.")


def main():
    parser = argparse.ArgumentParser(
        description="Daily Exchange Data Download Scheduler",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Run scheduler at 9 AM daily
  python scheduler.py
  
  # Run all downloads immediately (test mode)
  python scheduler.py --test
  
  # Run only NSE downloaders
  python scheduler.py --nse-only
  
  # Run NSE downloads immediately
  python scheduler.py --test --nse-only
        """
    )
    
    parser.add_argument(
        "--test",
        action="store_true",
        help="Run all downloads immediately (once) instead of scheduling"
    )
    parser.add_argument(
        "--nse-only",
        action="store_true",
        help="Download only NSE data (skip BSE)"
    )
    parser.add_argument(
        "--time",
        default="09:00",
        help="Scheduled time in HH:MM format (default: 09:00)"
    )
    
    args = parser.parse_args()
    
    if args.test:
        logger.info("TEST MODE: Running all downloads immediately...\n")
        run_all_downloads(nse_only=args.nse_only)
    else:
        # Parse scheduled time
        try:
            hour, minute = map(int, args.time.split(":"))
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError("Invalid time")
        except:
            logger.error(f"Invalid time format: {args.time}. Use HH:MM format.")
            sys.exit(1)
        
        # Start scheduler
        try:
            run_scheduler(nse_only=args.nse_only)
        except KeyboardInterrupt:
            logger.info("Scheduler stopped.")


if __name__ == "__main__":
    main()
