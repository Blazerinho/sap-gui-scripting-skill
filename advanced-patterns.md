# Advanced SAP GUI Scripting Patterns

## Table of Contents
1. [Batch Processing](#batch-processing)
2. [Multi-Transaction Workflows](#multi-transaction-workflows)
3. [Clipboard-Based Data Extraction](#clipboard-based-data-extraction)
4. [Transaction Navigation Helpers](#transaction-navigation-helpers)
5. [Logging and Audit Trail](#logging-and-audit-trail)

---

## Batch Processing

For processing multiple company codes, fiscal years, or document ranges:

```python
def process_batch(session, items, process_fn, delay=0.3):
    """
    Process a list of items through SAP, with error recovery.
    
    Args:
        session: SAP GUI session
        items: list of items to process
        process_fn: function(session, item) -> result
        delay: seconds between items
    """
    results = []
    errors = []
    
    for idx, item in enumerate(items):
        try:
            result = process_fn(session, item)
            results.append(result)
            print(f"[{idx+1}/{len(items)}] OK: {item}")
        except Exception as e:
            errors.append({"item": item, "error": str(e)})
            print(f"[{idx+1}/{len(items)}] ERROR: {item} — {e}")
            
            # Try to recover: go back to safe screen
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
            except:
                pass
        
        time.sleep(delay)
    
    return results, errors


# Example: query SE16H for multiple company codes
def query_company_code(session, bukrs):
    run_se16h_query(session, "BKPF", 
                    filter_fields={"BUKRS-LOW": bukrs, "GJAHR-LOW": "2024"})
    return extract_alv_grid(session)

company_codes = ["1000", "2000", "3000"]
all_results, errors = process_batch(session, company_codes, query_company_code)
```

---

## Multi-Transaction Workflows

Pattern for workflows spanning multiple transactions (e.g., query BKPF → look up vendor in FK03):

```python
def navigate_to_transaction(session, tcode):
    """Navigate to a transaction from anywhere."""
    ok_code = session.findById("wnd[0]/tbar[0]/okcd")
    ok_code.text = f"/n{tcode}"
    session.findById("wnd[0]").sendVKey(0)
    wait_for_session(session)
    
    # Verify we arrived
    msg, msg_type = get_status_message(session)
    if msg_type == "E":
        raise RuntimeError(f"Navigation to {tcode} failed: {msg}")

def get_current_transaction(session):
    """Read the current transaction code from the title bar."""
    try:
        return session.Info.Transaction
    except:
        return ""
```

---

## Clipboard-Based Data Extraction

Most reliable extraction method — uses SAP's own export rather than direct API:

```python
import win32clipboard
import pandas as pd
from io import StringIO

def copy_alv_to_clipboard(session, grid_id="wnd[0]/usr/cntlGRID1/shellcont/shell"):
    """Select all cells in ALV grid and copy to clipboard."""
    grid = session.findById(grid_id)
    
    # Select all rows
    grid.pressToolbarContextButton("&MB_EXPORT")  # Export button
    wait_for_session(session)
    # OR use Ctrl+A to select all, then Ctrl+C
    grid.setCurrentCell(0, grid.CurrentCellColumn)
    session.findById("wnd[0]").sendVKey(1)  # Ctrl+A select all — may vary
    wait_for_session(session)

def read_clipboard_as_dataframe(sep="\t"):
    """Read tab-separated clipboard content into pandas DataFrame."""
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()
    
    df = pd.read_csv(StringIO(data), sep=sep, dtype=str)
    return df.fillna("")
```

---

## Transaction Navigation Helpers

```python
# Common virtual key codes (sendVKey)
VKEYS = {
    "Enter":    0,
    "F1":       1,
    "F2":       2,
    "F3":       3,   # Back
    "F4":       4,   # F4 help / dropdown
    "F5":       5,
    "F6":       6,
    "F7":       7,
    "F8":       8,   # Execute
    "F9":       9,
    "F10":      10,
    "F11":      11,
    "F12":      12,  # Cancel
    "CtrlS":    11,  # Save (same as F11 in most transactions)
    "CtrlF":    33,  # Find
    "CtrlShF10": 37, # Local menu
}

def press_key(session, key_name):
    vkey = VKEYS.get(key_name)
    if vkey is None:
        raise ValueError(f"Unknown key: {key_name}")
    session.findById("wnd[0]").sendVKey(vkey)
    wait_for_session(session)

# Common button IDs in SAP standard toolbar
TOOLBAR_BUTTONS = {
    "save":     "wnd[0]/tbar[0]/btn[11]",
    "back":     "wnd[0]/tbar[0]/btn[3]",
    "exit":     "wnd[0]/tbar[0]/btn[15]",
    "cancel":   "wnd[0]/tbar[0]/btn[12]",
    "execute":  "wnd[0]/tbar[1]/btn[8]",   # application toolbar execute
}
```

---

## Logging and Audit Trail

For automation scripts that post or change data:

```python
import logging
from datetime import datetime

def setup_logging(script_name):
    log_file = f"{script_name}_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(script_name)

# Usage
logger = setup_logging("se16h_bkpf_extract")
logger.info(f"Starting extraction for company code {bukrs}")
logger.error(f"Failed on document {belnr}: {e}")
```
