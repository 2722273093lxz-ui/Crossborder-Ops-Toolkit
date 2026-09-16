# Competitor‑Cost‑Estimation
## Overview
This repository contains two desktop GUI helper tools built with Python for cross‑border luxury e‑commerce daily work.

### Tool 1: Competitor Cost Calculator
In specific company situations, we use the company's profit model to derive the competitor's cost from its price, enabling internal profit adjustment.

### Tool 2: eBay‑SOP‑Title‑Checker
For verifying the format accuracy of required listing titles and copy texts against internal SOP rules.
- Batch validate listing titles from Excel (.xlsx)
- Support single‑item manual checking
- Check title length limit, brand capitalization, material conflicts, duplicate‑color detection and other SOP constraints
- Export checked results to new Excel file
- Local GUI application, no network calls, can be packaged to EXE

### Tool 3: FedEx‑USPS‑Tracking‑Batch‑Checker
Batch logistics tracking query tool for FedEx and USPS tracking numbers.
> ⚠️ **Important Requirement**: This program relies on official FedEx & USPS developer API credentials to run functional queries.
> You need to apply for your own Client‑ID / Secret from FedEx Developer Portal and USPS Developer Portal.
> Without valid API keys, the GUI window can still be opened, but actual tracking data cannot be retrieved.
- Import tracking number list via Excel (.xlsx)
- Batch query logistics status and detailed scan history
- Preview query result inside GUI text panel
- Export all tracking outcomes to a new Excel file
- Graphical desktop GUI, configurable request delay to avoid API rate‑limiting
