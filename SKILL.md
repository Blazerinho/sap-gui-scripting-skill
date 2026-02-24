---
name: sap-gui-scripting
description: >
  Use this skill whenever the user wants to automate SAP GUI interactions using Python and the SAP GUI Scripting API (win32com.client). Triggers include: writing or debugging SAP GUI automation scripts, extracting data from SAP transactions via scripting, navigating SAP screens programmatically, handling SAP popups or modal dialogs in code, automating SE16H/SE16N/SM30 or any other SAP transaction, querying SAP tables via script, posting documents via SAP GUI automation, and any task involving SAPScriptingConnector, GuiApplication, GuiSession, or GuiComponent objects. Also use when the user asks how to find a control ID in SAP, how to handle SAP GUI errors in Python, or how to structure a reusable SAP automation script. Use this skill even if the user just pastes a fragment of SAP GUI Python code and asks for help — the patterns here apply broadly.
---

# SAP GUI Scripting Skill

Automation of SAP GUI using Python via the SAP GUI Scripting API (COM interface). This skill captures proven patterns, common pitfalls, and reusable structures for reliable SAP automation.

> **⚠️ Validate before use**: Control IDs, field names, and menu paths in this skill are based on standard SAP GUI patterns but **vary between SAP releases, screen variants, and client configurations**. Always verify IDs in your system using the Script Recorder (Alt+F12) before relying on them in production scripts. Sections requiring system-specific validation are marked with `[VERIFY IN YOUR SYSTEM]`.

## Quick Reference

| Task | Go to |
|------|-------|
| Session setup & connection | [Session Initialisation](#session-initialisation) |
| Finding control IDs | [Control Discovery](#control-discovery) |
| Reading table data | [Table Extraction](#table-extraction) |
| SE16H automation | [SE16H Patterns](#se16h-patterns) |
| Error handling | [Error Handling](#error-handling) |
| Popups & modal dialogs | [Dialog Handling](#dialog-handling) |
| Reusable script template | [Script Template](#script-template) |
| Advanced patterns | `references/advanced-patterns.md` |
| SE16H field reference | `references/se16h-reference.md` |

---

## Core Concepts

The SAP GUI Scripting API exposes the SAP GUI as a COM object tree. Python accesses it via `win32com.client`. The hierarchy is:

```
SapGuiNamespace (ROT entry)
  └── GuiApplication
        └── GuiConnection  [0..n]
              └── GuiSession  [0..n]
                    └── GuiFrameWindow (main window)
                          └── GuiUserArea
                                └── [screen controls]
```

**Critical behaviour**: SAP GUI throws COM exceptions rather than returning None or raising Python exceptions. Always wrap `findById()` calls in try/except. Never assume a control exists — screen state changes between steps.

---

## Session Initialisation

### Standard Pattern (single session)

```python
import win32com.client
import pywintypes

def get_sap_session():
    """Connect to an already-open SAP GUI session."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not SapGuiAuto:
            raise RuntimeError("SAP GUI not running")
        
        application = SapGuiAuto.GetScriptingEngine
        if not application:
            raise RuntimeError("SAP GUI Scripting not enabled")
        
        connection = application.Children(0)   # first connection
        session = connection.Children(0)        # first session
        return session
    
    except pywintypes.com_error as e:
        raise RuntimeError(f"Could not connect to SAP GUI: {e}")
```

### Multi-Session Pattern

```python
def get_session_by_info(system_name=None, session_index=0):
    """Get a specific session, optionally filtering by system."""
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    
    for conn_idx in range(application.Children.Count):
        connection = application.Children(conn_idx)
        for sess_idx in range(connection.Children.Count):
            session = connection.Children(sess_idx)
            if system_name is None or system_name in session.Info.SystemName:
                if sess_idx == session_index:
                    return session
    
    raise RuntimeError(f"No matching session found for system '{system_name}'")
```

### Prerequisites Check

Before running any script, verify:
1. SAP GUI is open and logged in
2. Scripting is enabled: SAP GUI Options → Scripting tab → "Enable Scripting" checked
3. The target session is not locked/busy

---

## Control Discovery

Finding the correct control ID is the most common challenge. Use these approaches:

### Method 1: SAP GUI Script Recorder (fastest)
1. In SAP GUI: Customize Local Layout (Alt+F12) → Script Recording and Playback
2. Record the manual action
3. Read the generated VBS file — copy the `findById(...)` paths directly

### Method 2: Python Discovery Script

```python
def explore_children(element, depth=0, max_depth=3):
    """Recursively print control tree for ID discovery."""
    indent = "  " * depth
    try:
        print(f"{indent}[{element.Type}] ID: {element.Id}")
        print(f"{indent}  Name: {getattr(element, 'Name', 'N/A')}")
        print(f"{indent}  Text: {getattr(element, 'Text', 'N/A')[:50]}")
    except:
        pass
    
    if depth < max_depth:
        try:
            for i in range(element.Children.Count):
                explore_children(element.Children(i), depth + 1, max_depth)
        except:
            pass

# Usage: explore from main window
session = get_sap_session()
explore_children(session.findById("wnd[0]"))
```

### Common Control ID Patterns

> **[VERIFY IN YOUR SYSTEM]** These are typical patterns — actual IDs depend on the specific transaction and screen variant. Use the Script Recorder or `explore_children()` to confirm.

```python
# Main window
session.findById("wnd[0]")

# Menu bar
session.findById("wnd[0]/mbar")

# Toolbar
session.findById("wnd[0]/tbar[0]")
session.findById("wnd[0]/tbar[1]")  # application toolbar

# User area (where screen fields live)
session.findById("wnd[0]/usr")

# Typical field patterns
session.findById("wnd[0]/usr/ctxtBUKRS")     # context text field (company code)
session.findById("wnd[0]/usr/txtBELNR")      # text field
session.findById("wnd[0]/usr/chkXBUCH")      # checkbox
session.findById("wnd[0]/usr/cmbBLART")      # combo/dropdown
session.findById("wnd[0]/usr/radKSEL")       # radio button

# Modal popup window
session.findById("wnd[1]")
session.findById("wnd[1]/usr/txtMESSAGE")

# Grid control (ALV)
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
```

---

## Error Handling

### Standard Wrapper

```python
import win32com.client
import pywintypes
import time

def safe_find(session, control_id, retries=3, delay=0.5):
    """Find a control with retry logic — handles timing issues."""
    for attempt in range(retries):
        try:
            control = session.findById(control_id)
            return control
        except pywintypes.com_error:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise RuntimeError(f"Control not found after {retries} attempts: {control_id}")
```

### Synchronisation

SAP GUI is asynchronous. After triggering actions (Enter, button click, transaction call), wait for the session to finish processing:

```python
def wait_for_session(session, timeout=30):
    """Wait until SAP GUI is ready for input."""
    start = time.time()
    while session.Busy:
        if time.time() - start > timeout:
            raise TimeoutError("SAP session still busy after timeout")
        time.sleep(0.2)

# Always call after: pressing Enter, clicking buttons, calling transactions
session.findById("wnd[0]").sendVKey(0)  # Enter
wait_for_session(session)
```

### Common COM Errors

| Symptom | Likely Cause | Fix |
|---------|-------------|-----|
| `pywintypes.com_error: -2147352567` | Control ID wrong or screen changed | Re-record with Script Recorder |
| `AttributeError: 'NoneType'` | `GetObject("SAPGUI")` returned None | SAP GUI not open or scripting disabled |
| Script runs but nothing happens | Scripting disabled in SAP GUI options | Enable in Customize Local Layout → Scripting |
| Random failures on fast machines | Missing synchronisation | Add `wait_for_session()` after each action |

---

## Dialog Handling

SAP frequently shows modal popups (confirmation dialogs, error messages, info boxes). Always handle them proactively:

```python
def dismiss_popup(session, confirm=True):
    """Handle common SAP modal dialogs. Returns True if popup was found."""
    try:
        popup = session.findById("wnd[1]")
        if confirm:
            # Try "Yes" / "OK" / "Enter" buttons in order
            for btn_id in ["wnd[1]/usr/btnSPOP-OPTION1", 
                           "wnd[1]/tbar[0]/btn[0]",
                           "wnd[1]/tbar[0]/btn[11]"]:
                try:
                    session.findById(btn_id).press()
                    wait_for_session(session)
                    return True
                except pywintypes.com_error:
                    continue
        else:
            # Cancel / No
            for btn_id in ["wnd[1]/usr/btnSPOP-OPTION2",
                           "wnd[1]/tbar[0]/btn[12]"]:
                try:
                    session.findById(btn_id).press()
                    wait_for_session(session)
                    return True
                except pywintypes.com_error:
                    continue
        return False
    except pywintypes.com_error:
        return False  # No popup present

def get_status_bar_message(session):
    """Read the status bar message (errors, info, warnings)."""
    try:
        statusbar = session.findById("wnd[0]/sbar")
        return statusbar.Text, statusbar.MessageType  # MessageType: E=error, W=warning, S=success
    except:
        return "", ""
```

---

## Table Extraction

### ALV Grid (most common in modern transactions)

```python
def extract_alv_grid(session, grid_id="wnd[0]/usr/cntlGRID1/shellcont/shell"):
    """Extract all data from an ALV grid control."""
    try:
        grid = session.findById(grid_id)
    except pywintypes.com_error:
        raise RuntimeError(f"ALV Grid not found at: {grid_id}")
    
    row_count = grid.RowCount
    col_count = grid.ColumnCount
    
    # Get column names
    columns = []
    for col_idx in range(col_count):
        col_name = grid.GetColumnDataByName("", col_idx)  # varies by version
        columns.append(col_name)
    
    # Extract rows
    data = []
    for row_idx in range(row_count):
        row = {}
        for col_idx, col_name in enumerate(columns):
            try:
                value = grid.GetCellValue(row_idx, col_name)
                row[col_name] = value
            except:
                row[col_name] = ""
        data.append(row)
    
    return data

def export_alv_to_clipboard(session):
    """
    Simpler alternative: use SAP's built-in export to clipboard.
    Works for most ALV grids without needing to know column IDs.
    [VERIFY IN YOUR SYSTEM] — menu path after sendVKey(37) varies by SAP release.
    Record the export steps manually with Script Recorder to get exact IDs.
    """
    # Ctrl+Shift+F10 or local menu → Spreadsheet
    session.findById("wnd[0]").sendVKey(37)  # Ctrl+Shift+F7 for local menu
    wait_for_session(session)
    # Then navigate local menu to export — record this step in your system
```

### Classic Table Control (older transactions)

```python
def extract_table_control(session, table_id, columns):
    """Extract from classic table control (e.g., in SM30-style screens)."""
    table = session.findById(table_id)
    rows = table.RowCount
    data = []
    
    for row in range(rows):
        table.FirstVisibleRow = row  # scroll to row
        row_data = {}
        for col_name, col_id_suffix in columns.items():
            try:
                cell_id = f"{table_id}/txtfield[{col_id_suffix},{row}]"
                row_data[col_name] = session.findById(cell_id).Text
            except:
                row_data[col_name] = ""
        data.append(row_data)
    
    return data
```

---

## SE16H Patterns

SE16H (General Table Display with Aggregation) is the primary transaction for data analytics queries. Key behaviour:

- Preferred over SE16N because it supports aggregation (SUM, COUNT, MAX, MIN, AVG)
- Fields shown depend on the table selected — always verify field names against the specific table
- Large result sets require handling the "Maximum number of hits" popup
- Results can be exported via the local menu (Ctrl+Shift+F10 → Spreadsheet)

> **[VERIFY IN YOUR SYSTEM]** The field IDs below (`ctxtGD-TAB`, `txtGD-MAX_LINES`) are standard for ECC 6.0 and S/4HANA but may differ in older releases or heavily customised systems. Use the Script Recorder on SE16H in your system to confirm.

```python
def run_se16h_query(session, table_name, fields_to_display, 
                    filter_fields=None, max_rows=500):
    """
    Execute an SE16H query and return to results screen.
    
    Args:
        table_name: SAP table name e.g. 'BKPF'
        fields_to_display: list of field names to show in output
        filter_fields: dict of {field_name: value} for WHERE conditions
        max_rows: maximum number of rows to retrieve
    """
    # Navigate to SE16H
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16H"
    session.findById("wnd[0]").sendVKey(0)
    wait_for_session(session)
    
    # Enter table name
    try:
        table_field = session.findById("wnd[0]/usr/ctxtGD-TAB")
        table_field.text = table_name
        session.findById("wnd[0]").sendVKey(0)  # Enter to confirm table
        wait_for_session(session)
    except pywintypes.com_error:
        raise RuntimeError(f"Could not enter table name: {table_name}")
    
    # Set max rows
    try:
        max_field = session.findById("wnd[0]/usr/txtGD-MAX_LINES")
        max_field.text = str(max_rows)
    except pywintypes.com_error:
        pass  # field may not be visible
    
    # Apply filters if provided
    if filter_fields:
        for field_name, value in filter_fields.items():
            try:
                field = session.findById(f"wnd[0]/usr/txt{field_name}")
                field.text = str(value)
            except pywintypes.com_error:
                # Try context text field variant
                try:
                    field = session.findById(f"wnd[0]/usr/ctxt{field_name}")
                    field.text = str(value)
                except pywintypes.com_error:
                    print(f"Warning: Could not set filter for field {field_name}")
    
    # Execute (F8)
    session.findById("wnd[0]").sendVKey(8)
    wait_for_session(session)
    
    # Handle "maximum rows exceeded" popup
    dismiss_popup(session, confirm=True)
    wait_for_session(session)
    
    # Check status bar for errors
    msg, msg_type = get_status_bar_message(session)
    if msg_type == "E":
        raise RuntimeError(f"SE16H error: {msg}")
    
    return True  # caller reads the grid

# Read SE16H reference for field-level documentation
# See: references/se16h-reference.md
```

---

## Script Template

Use this as the starting point for any new SAP automation script:

```python
"""
SAP Automation Script: [description]
Transaction: [e.g. SE16H / FB03 / etc.]
Author: [name]
"""

import win32com.client
import pywintypes
import time
import pandas as pd
from datetime import datetime


# ── helpers ──────────────────────────────────────────────────────────────────

def get_sap_session():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)

def wait_for_session(session, timeout=30):
    start = time.time()
    while session.Busy:
        if time.time() - start > timeout:
            raise TimeoutError("SAP session timeout")
        time.sleep(0.2)

def safe_find(session, control_id, retries=3, delay=0.5):
    for attempt in range(retries):
        try:
            return session.findById(control_id)
        except pywintypes.com_error:
            if attempt < retries - 1:
                time.sleep(delay)
    raise RuntimeError(f"Control not found: {control_id}")

def dismiss_popup(session, confirm=True):
    try:
        popup = session.findById("wnd[1]")
        btn = "wnd[1]/tbar[0]/btn[0]" if confirm else "wnd[1]/tbar[0]/btn[12]"
        session.findById(btn).press()
        wait_for_session(session)
        return True
    except pywintypes.com_error:
        return False

def get_status_message(session):
    try:
        bar = session.findById("wnd[0]/sbar")
        return bar.Text, bar.MessageType
    except:
        return "", ""


# ── main logic ────────────────────────────────────────────────────────────────

def main():
    print(f"Starting at {datetime.now():%Y-%m-%d %H:%M:%S}")
    
    session = get_sap_session()
    results = []
    
    try:
        # ── your automation steps here ──
        pass
        
    except Exception as e:
        msg, _ = get_status_message(session)
        print(f"Error: {e}")
        if msg:
            print(f"SAP status: {msg}")
        raise
    
    # Export results
    if results:
        df = pd.DataFrame(results)
        output_file = f"output_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"Saved {len(df)} rows to {output_file}")
    
    print("Done.")


if __name__ == "__main__":
    main()
```

---

## Key Behaviours to Remember

1. **Never assume screen state** — always navigate explicitly (e.g., `/nSE16H` instead of assuming you're already there)
2. **`findById()` throws, not returns None** — always use try/except
3. **Add `wait_for_session()` after every action** — SAP GUI is asynchronous
4. **Popups block everything** — check for `wnd[1]` after each action that might trigger a message
5. **Status bar is your friend** — always read it after F8 / Execute to detect errors
6. **Prefer F8 (sendVKey 8) over clicking Execute buttons** — more reliable across screen variants
7. **Control IDs are screen-state dependent** — an ID valid on initial screen may not exist after data entry

---

## References

- `references/advanced-patterns.md` — Batch processing, multi-session parallelism, clipboard export
- `references/se16h-reference.md` — SE16H field reference, aggregation patterns, BKPF/BSEG common queries
