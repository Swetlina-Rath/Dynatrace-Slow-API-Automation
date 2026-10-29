# Dynatrace Slow API Automation

A Python + Playwright automation script that finds slow APIs and services in Dynatrace, identifies why they are slow, and builds a ready-to-share report. Built for Aditya Birla Capital Digital.

> An API (also called an endpoint) is a single request URL or operation of an application, for example `getAccountList`.

## What it does
- Opens the Dynatrace Multidimensional Analysis (MDA) Top Web Requests view for the selected time window.
- Reads the full table, including rows that only load while scrolling.
- Keeps APIs whose average response time is above 5 seconds (top 20). `schedule` and `terminateSessionAndUpdateLead` are excluded.
- Opens each API's **Response time hotspots** view and reads the top 2 findings.
- Classifies the findings as **External** (calls to URLs, network IO) or **Internal** (active wait, suspension, SQL and database usage).
- Writes an Excel file and a Word file.

## Output
Saved in the `endpoint_response_time_report` folder next to the script:
- `Top_20_Endpoint_Response_Time_Greater_Than_5s.xlsx`
- `Top_20_Endpoint_Response_Time_Greater_Than_5s.docx`

## Prerequisites
- Python 3.9 or newer
- Chromium installed through Playwright (see Installation)
- Access to the Dynatrace tenant (sign in when the browser opens)
- A desktop session, because the script opens a visible browser window

Supported on Windows, macOS and Linux.

## Installation

**Windows**
```bash
pip install playwright python-docx openpyxl
playwright install chromium
```

**macOS / Linux**
```bash
pip3 install playwright python-docx openpyxl
python3 -m playwright install chromium
```

On some Linux distributions, pip may refuse the install. In that case use a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate
pip install playwright python-docx openpyxl
playwright install chromium
```

On Linux, if Chromium reports missing system libraries, run:
```bash
python3 -m playwright install-deps chromium
```

## Usage

**Windows**
```bash
python endpoint.py
```
If `python` is not recognised, use `py endpoint.py`.

**macOS / Linux**
```bash
python3 endpoint.py
```

## When to run it (IST)
The report window is chosen from the time you run the script:

| Run time | Report window |
|---|---|
| 11:00 - 14:59 | 08:00 - 11:00 |
| 15:00 - 17:59 | 11:00 - 15:00 |
| 18:00 - 23:59 | 15:00 - 18:00 |

Outside these times the script stops with an error. If the window has passed, collect the data manually.

## Important notes
- Each run **deletes and replaces** the previous report folder. Copy the earlier files first if you need them.
- Do not click inside the browser while the script is running.
- The Dynatrace tenant URL is set in `get_mda_url()` inside the script.
