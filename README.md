[README.md](https://github.com/user-attachments/files/31952787/README.md)
# 🔍 Lead Generator Pro — أداة توليد العملاء المحتملين

A desktop app that automatically finds business leads (name, phone, email, website) and exports them to Excel.

## Features
- **Search by keyword + location** — "restaurants Cairo", "dentists Dubai", "lawyers Alexandria"
- **Multi-source** — Google, Yelp, or both
- **Auto email extraction** — visits each website and pulls contact emails
- **Clean GUI** — no command line needed, works like any desktop app
- **Export to Excel or CSV** — formatted, ready to use
- **Real-time results** — see leads as they're found
- **Stop anytime** — interrupt the search at any point

## Quick Start

```bash
pip install customtkinter requests beautifulsoup4 openpyxl
python lead_scraper.py
```

## How to Use

1. Enter a **business type** (e.g. "real estate agents", "dentists", "restaurants")
2. Enter a **location** (e.g. "Cairo Egypt", "Dubai UAE")
3. Choose number of results and source (Google/Yelp/Both)
4. Toggle "Extract emails from websites" for deeper contact info
5. Click **Start Scraping**
6. Watch leads appear in real time
7. Click **Export Results** → save as Excel or CSV

## Output Columns
- Business Name
- Phone Number
- Email Address
- Website URL
- Address
- Source (Google/Yelp)

## Project Structure
```
├── lead_scraper.py     # Main app (GUI + scraping engine)
├── requirements.txt
└── README.md
```

## Use Cases
- Find potential clients for your services
- Build cold email lists for marketing campaigns
- Market research and competitor analysis
- Real estate lead generation
- B2B sales prospecting

## Technologies
Python 3.10+ • CustomTkinter • requests • BeautifulSoup4 • openpyxl
