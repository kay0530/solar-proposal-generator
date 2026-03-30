# Solar Proposal Generator

Streamlit web application for generating PPA/EPC solar proposal decks (PPTX).

Built with `python-pptx` for programmatic slide generation and `openpyxl`/`xlwings` for Excel-based calculation engine integration.

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run proposal_generator/app.py --server.port 8502
```

Or use `run_app.bat` on Windows.

## Project Structure

```
proposal_generator/
  app.py              # Streamlit UI
  generator.py        # PPTX assembly engine
  excel_runner.py     # Excel calculation engine interface
  ppa_calc.py         # PPA pricing calculations
  demand_calc.py      # Demand/load calculations
  subsidy_calc.py     # Subsidy calculation logic
  box_client.py       # Box API integration
  utils.py            # Shared slide drawing utilities
  slides/
    ppa/              # PPA proposal slide generators (pp0-pp13)
    epc/              # EPC proposal slide generators (ep0-ep8)
    new/              # Shared/new slide generators
templates/            # PPTX template files
input/                # Place Excel input files here (gitignored)
```

## Features

- PPA (Power Purchase Agreement) proposal deck generation
- EPC (Engineering, Procurement, Construction) proposal deck generation
- Drag-and-drop slide ordering
- Excel-based calculation engine integration
- CO2 reduction calculations
- Subsidy eligibility checks
- Competitor comparison slides
