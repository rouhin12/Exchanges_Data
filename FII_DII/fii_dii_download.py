#!/usr/bin/env python3
"""Download FII/DII data from Moneycontrol and save to Excel."""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path
from datetime import datetime

OUTPUT_FILE = Path(__file__).parent / "fii_dii_data.xlsx"
OUTPUT_FILE.parent.mkdir(exist_ok=True)

URL = "https://www.moneycontrol.com/stocks/marketstats/fii_dii_activity/index.php"

def scrape_fii_dii():
    """Scrape FII/DII data from Moneycontrol."""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(URL, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Find the table
        table = soup.find("div", class_="fidi_tbescrol")
        if not table:
            print("❌ Table not found on page")
            return None
        
        rows = []
        tbody = table.find("tbody")
        if not tbody:
            print("❌ Table body not found")
            return None
        
        for tr in tbody.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 7:
                continue
            
            try:
                date_text = tds[0].get_text(strip=True)
                # Handle duplicated date text (e.g., "16-Feb-202616-Feb-2026" -> take first match)
                # Extract just the date part using regex
                import re
                date_match = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', date_text)
                if date_match:
                    date_text = date_match.group()
                
                # Handle negative values with minus sign
                def safe_float(text):
                    text = text.replace(",", "").strip()
                    # Handle minus sign
                    if text.startswith("−") or text.startswith("-"):
                        return -float(text.lstrip("−-"))
                    return float(text)
                
                fii_purchase = safe_float(tds[1].get_text(strip=True))
                fii_sales = safe_float(tds[2].get_text(strip=True))
                fii_net = safe_float(tds[3].get_text(strip=True))
                dii_purchase = safe_float(tds[4].get_text(strip=True))
                dii_sales = safe_float(tds[5].get_text(strip=True))
                dii_net = safe_float(tds[6].get_text(strip=True))
                
                rows.append({
                    "Date": date_text,
                    "FII_Gross_Purchase": fii_purchase,
                    "FII_Gross_Sales": fii_sales,
                    "FII_Net": fii_net,
                    "DII_Gross_Purchase": dii_purchase,
                    "DII_Gross_Sales": dii_sales,
                    "DII_Net": dii_net,
                })
            except (ValueError, AttributeError, TypeError) as e:
                continue
        
        if not rows:
            print("ERROR: No data rows found")
            return None
        
        df = pd.DataFrame(rows)
        # Use mixed format parsing to handle various date formats
        df["Date"] = pd.to_datetime(df["Date"], format="mixed", dayfirst=True, errors="coerce")
        # Remove rows with invalid dates
        df = df.dropna(subset=["Date"])
        df = df.sort_values("Date")
        
        print(f"OK: Downloaded {len(df)} rows of FII/DII data")
        return df
        
    except Exception as e:
        print(f"ERROR: Error scraping: {e}")
        return None

def save_to_excel(df):
    """Save DataFrame to Excel file."""
    if df is None or df.empty:
        print("ERROR: No data to save")
        return False
    
    try:
        # Load existing data if it exists
        if OUTPUT_FILE.exists():
            existing = pd.read_excel(OUTPUT_FILE)
            existing["Date"] = pd.to_datetime(existing["Date"])
            # Remove duplicates, keeping new data
            df = pd.concat([existing, df], ignore_index=True)
            df = df.drop_duplicates(subset=["Date"], keep="last")
            df = df.sort_values("Date")
        
        # Save to Excel
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="FII_DII", index=False)
        
        print(f"OK: Saved to {OUTPUT_FILE}")
        print(f"   Total records: {len(df)}")
        print(f"   Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")
        return True
        
    except Exception as e:
        print(f"ERROR: Error saving: {e}")
        return False

if __name__ == "__main__":
    print(f"Downloading FII/DII data from Moneycontrol...")
    df = scrape_fii_dii()
    if df is not None:
        save_to_excel(df)
    else:
        print("Failed to download FII/DII data")
