INVOICE CONSOLIDATION WEB APPLICATION
========================================

Run on Windows:
1. Install Python 3.11+.
2. Double-click RUN_WEB_APP.bat.
3. Open http://127.0.0.1:5000 in the browser.

Workflow:
- Upload Invoice Excel
- Upload Master List Excel
- Upload Upload Template v1_2.xlsx
- Enter inward/header details
- Select product-code first-letter groups
- Blank Product Code = Non-Control Items
- Select No origin split or Origin-wise split
- Generate Excel; maximum 50 consolidated rows per file
- Master List page supports import, add, edit and delete

Important matching behavior:
- HS Code found + Product Code blank => Non-Control Item.
- HS Code not found in Master List => remains an exception and is NOT called Non-Control.
- Product descriptions are deduplicated and joined using " / ".
- Amount uses Total Amount including charges.
- Consolidation key = HS Code + Origin.
- IMP_BL!B12 receives "Inv no : ..." with unique Sales Invoice Numbers.

This is a local web application. It runs in a browser but the processing/data stay on the computer.
