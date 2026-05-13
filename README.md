# Inventory Processor

A web app to combine, classify, and split inventory Excel reports into Y&S and Non Y&S exports.

## Features

- **Drag & drop** multiple `.xlsx` inventory files
- Combines all files into one dataset
- Adds **Total Cost** column (`Quantity × Cost`)
- Adds **Main Company** column (with mappings: YS-Seatgeek/YS-Seatgeek2 → YS Tickets, YSA 2/YSA 3 → YSA)
- Splits output into two files based on company classification:
  - `Inventory - {date} (YS).xlsx` — tab named **Available**
  - `Inventory - {date} (Non YS).xlsx` — tab named **Available**
- **Clear** button resets all inputs and outputs
- Month end date is a user input field

## Local Development

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## Deploy to Railway

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub repo
3. Select your repo — Railway auto-detects and deploys
4. Your app will be live at the provided Railway URL

## Deploy to Heroku

```bash
heroku create your-app-name
git push heroku main
```

## Company Mapping

| Company | Classification |
|---|---|
| Damon and Crew | Non Y&S |
| The Ticket Guy | Non Y&S |
| YourTickets | Non Y&S |
| GK LLC, Jacks YS, Levovitz, Needle Tickets LLC, Pollak Tickets, Yoni Levine, YS Katz, YS Tickets, YS TL, YSA, YSA 2, YSA 3, YSM Tickets, YSS Tickets, YS-Seatgeek, YS-Seatgeek2, YSW | Y&S |

### Main Company overrides
- `YS-Seatgeek` / `YS-Seatgeek2` → `YS Tickets`
- `YSA 2` / `YSA 3` → `YSA`
