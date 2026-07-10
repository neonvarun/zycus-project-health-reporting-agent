import os
import re
import pandas as pd
from datetime import datetime

def load_dotenv(path=".env"):
    """Manually loads environment variables from a .env file."""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    if "=" in line:
                        key, val = line.split("=", 1)
                        os.environ[key.strip()] = val.strip()

def clean_text(val) -> str:
    """Safely cleans text, removing extra spaces and handling nulls."""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()

def parse_date(val):
    """Safely parses dates from various formats, including Excel serial dates and string formats."""
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    
    # Handle numeric Excel serial date
    if isinstance(val, (int, float)):
        try:
            # Excel date serials offset from 1899-12-30 due to leap year bug
            return pd.to_datetime(val, unit='D', origin='1899-12-30').to_pydatetime()
        except:
            return None

    # Handle string values
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "null", "n/a", "none"]:
        return None

    # Try standard string date parsers
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str, fmt)
        except ValueError:
            continue
            
    # Try pandas fallback
    try:
        return pd.to_datetime(val_str).to_pydatetime()
    except:
        return None
