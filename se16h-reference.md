# SE16H Reference

SE16H (General Table Display with Aggregation) is the primary tool for SAP FI/CO data analytics queries. This reference covers field names, patterns, and common queries for financial analytics.

## Table of Contents
1. [SE16H vs SE16N](#se16h-vs-se16n)
2. [Screen Field IDs](#screen-field-ids)
3. [Aggregation Functions](#aggregation-functions)
4. [Common BKPF Queries](#common-bkpf-queries)
5. [Common BSEG Queries](#common-bseg-queries)
6. [BSET (Tax) Queries](#bset-tax-queries)
7. [Export Patterns](#export-patterns)

---

## SE16H vs SE16N

| Feature | SE16H | SE16N |
|---------|-------|-------|
| Aggregation (SUM, COUNT) | ✅ Yes | ❌ No |
| Field selection | ✅ Flexible | ✅ Flexible |
| Performance on large tables | ⚠️ Same | ⚠️ Same |
| Scripting reliability | ✅ Consistent | ✅ Consistent |
| **Recommendation** | **Use for analytics** | Use for simple lookups |

---

## Screen Field IDs

After entering table name and pressing Enter, SE16H shows the selection screen. Common field patterns:

```python
# Table name entry (initial screen)
"wnd[0]/usr/ctxtGD-TAB"          # table name

# After table is entered — selection screen fields
# Fields are named by their SAP field name with prefix:
# txt  = text input field
# ctxt = context text field (with F4 help)
# chk  = checkbox

# Common BKPF fields in SE16H selection screen:
"wnd[0]/usr/ctxtBUKRS-LOW"       # Company Code (from)
"wnd[0]/usr/ctxtBUKRS-HIGH"      # Company Code (to)
"wnd[0]/usr/ctxtGJAHR-LOW"       # Fiscal Year (from)
"wnd[0]/usr/ctxtGJAHR-HIGH"      # Fiscal Year (to)
"wnd[0]/usr/ctxtBELNR-LOW"       # Document Number (from)
"wnd[0]/usr/ctxtBELNR-HIGH"      # Document Number (to)
"wnd[0]/usr/ctxtBLART-LOW"       # Document Type
"wnd[0]/usr/ctxtAWTYP-LOW"       # Reference Transaction Type
"wnd[0]/usr/txtBUDAT-LOW"        # Posting Date (from)
"wnd[0]/usr/txtBUDAT-HIGH"       # Posting Date (to)

# Max rows / hits
"wnd[0]/usr/txtGD-MAX_LINES"     # Maximum number of hits

# Output fields selection button
"wnd[0]/tbar[1]/btn[4]"          # Fields for output list

# Execute
session.findById("wnd[0]").sendVKey(8)  # F8
```

---

## Aggregation Functions

SE16H supports aggregation in the output field selection (accessed via toolbar button "Fields for output list" or F7):

In scripting, aggregation is set via the field properties dialog. Practical approach:
1. Script the navigation to SE16H
2. Set filters
3. For aggregation: record the steps manually first, then replicate

Common aggregation use cases:
- COUNT(*) GROUP BY AWTYP → document type distribution
- SUM(DMBTR) GROUP BY BUKRS, GJAHR → totals by company and year
- COUNT(*) GROUP BY BLART, AWTYP → cross-classification

---

## Common BKPF Queries

BKPF = FI Document Header. Always filter BUKRS + GJAHR first for performance.

### Document Origin Analysis
```python
filters = {
    "BUKRS-LOW": "1000",      # company code
    "GJAHR-LOW": "2024",
    "GJAHR-HIGH": "2024",
    "AWTYP-LOW": "RMRP",     # logistics invoice verification
}
# Output fields: BELNR, GJAHR, BLART, AWTYP, AWKEY, BUDAT, BLDAT, XBLNR
```

### Find Deferred Tax Documents
```python
filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
    "AWTYP-LOW": "DEFTAX",
    "AWTYP-HIGH": "DEFTAX",
}
# Output fields: BELNR, GJAHR, BUDAT, AWKEY, STBLG
```

### Find Reversed Documents
```python
# STBLG is not empty = document has been reversed
# In SE16H: use selection with STBLG != '' 
# (enter a wildcard in STBLG field or use NE option)
filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
}
# Then filter STBLG <> '' in output or via SE16H selection option
```

### Accrual Documents (FBS1)
```python
filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
    "BLART-LOW": "SA",        # GL document type — adjust per client
    "AWTYP-LOW": "ACCR",
}
# Output fields: BELNR, BUDAT, STODT (planned reversal date), STBLG
```

---

## Common BSEG Queries

BSEG = FI Document Line Items. Very large table — NEVER query without BUKRS + GJAHR + BELNR range.

```python
# Always include these as mandatory filters:
filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
    "GJAHR-HIGH": "2024",
    # Add document number range or GL account range
    "HKONT-LOW": "0001000000",   # GL account range
    "HKONT-HIGH": "0001999999",
}
# Output fields: BUKRS, BELNR, GJAHR, BUZEI, HKONT, DMBTR, SHKZG, KOSTL, AUFNR
```

### Line Item with Cost Object
```python
filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
    "KOSTL-LOW": "CC001",      # cost centre
}
# Output: BELNR, GJAHR, HKONT, DMBTR, SHKZG, KOSTL, AUFNR, PRCTR
```

---

## BSET (Tax) Queries

BSET = Tax Line Items per FI document. Key for VAT analytics and deferred tax.

```python
# Key BSET fields:
# BUKRS  = company code
# BELNR  = FI document number
# GJAHR  = fiscal year
# MWSKZ  = tax code
# KSCHL  = condition type
# HWSTE  = tax amount (local currency)
# HWBAS  = tax base amount (local currency)
# FWSTE  = tax amount (foreign currency)
# FWBAS  = tax base amount (foreign currency)
# TXJCD  = tax jurisdiction (US)

filters = {
    "BUKRS-LOW": "1000",
    "GJAHR-LOW": "2024",
    "MWSKZ-LOW": "V1",        # specific tax code
}
# Join to BKPF via BUKRS+BELNR+GJAHR for AWTYP context
```

---

## Export Patterns

### Via SAP Local Menu (most reliable)

```python
def export_se16h_to_excel(session, output_path):
    """Export SE16H results via SAP local menu → Spreadsheet."""
    # Results must already be displayed
    
    # Open local menu
    session.findById("wnd[0]").sendVKey(37)  # Ctrl+Shift+F10
    wait_for_session(session)
    
    # Navigate: Spreadsheet (position varies — record this step)
    # Alternative: use menu path
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
    except pywintypes.com_error:
        # Try alternate menu path
        pass
    wait_for_session(session)
    
    # In the export dialog, set file path
    try:
        path_field = session.findById("wnd[1]/usr/ctxtDY_PATH")
        file_field = session.findById("wnd[1]/usr/ctxtDY_FILENAME")
        
        import os
        path_field.text = os.path.dirname(output_path)
        file_field.text = os.path.basename(output_path)
        
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Replace/OK
        wait_for_session(session)
    except pywintypes.com_error:
        pass

### Via Clipboard (simple, works everywhere)

```python
import subprocess
import pandas as pd

def export_via_clipboard(session):
    """
    Select all + copy from ALV grid, then read clipboard.
    Works when direct grid access is unreliable.
    """
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.selectAll()
    grid.contextMenu()
    # Then navigate to copy — easier to record this step
    
    # Read clipboard into pandas
    df = pd.read_clipboard(sep='\t')
    return df
```
