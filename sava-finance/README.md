# Sava Financial Model Dashboard

Live dashboard for the Sava Technologies financial model. Reads live from a Google Sheets model published to the web.

## Files
- `index.html` — page structure, tabs, layout
- `style.css` — Sava palette (cream/charcoal/clinical green)
- `dashboard.js` — data fetching from Google Sheets, rendering, scenario toggle
- `README.md` — this file

## Deployment

Mirrors the pattern of `antonschneider.com/stripe-cbfs/`:

1. Drop `index.html`, `style.css`, `dashboard.js` into a folder e.g. `sava-cbfs/` in the website repo
2. Push to GitHub Pages
3. Access at `antonschneider.com/sava-cbfs/`

## Google Sheets requirements

The dashboard reads from sheet `1fJLaVixEazLYqhqpE6dyC5g3bY3jrdzu2fo3rj2sr8k`.

For the dashboard to work, the sheet must be:
1. **Published to web** (File → Share → Publish to web → Entire document → Web page)

Sharing permissions don't need to change — the sheet can stay restricted; "Publish to web" provides a separate read-only data endpoint that doesn't expose the underlying sheet.

## Tabs fetched
- `Current Year Overview`
- `Forecast Years Overview`
- `RFE_Base`, `RFE_Bear`, `RFE_Bull`
- `Assumptions`
- `Headcount Plan`
- `Revenue Build`
- `Summary`

## Sheet → Dashboard latency

After editing the sheet, Google republishes CSV/JSON endpoints in 1-5 minutes. Click the **Refresh** button after that delay to pull new data.
