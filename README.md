# SKYLUXE — Way Forward Media Plan Studio

A complete, client-side real estate media plan calculator and 17-slide
PowerPoint deck generator. No backend, no build step — just static files.

## What's inside

```
index.html              Main app shell (three tabs)
css/style.css            Luxury navy/gold theme, tables, slide styling
js/calculator.js         Pure calculation engine (no DOM access)
js/pptx-generator.js     Builds the 17-slide .pptx with PptxGenJS
js/app.js                State, rendering, event wiring
README.md                This file
```

## Run it locally

No installation needed. Just open `index.html` in a modern browser
(Chrome, Edge, Firefox, Safari). All libraries load from CDN.

## Host it on GitHub Pages (drag & drop)

1. Create a new GitHub repository (public).
2. Drag and drop all files in this folder — keeping the `css/` and `js/`
   subfolders intact — into the repository via the GitHub web UI
   ("Add file → Upload files"), or push them with git:
   ```bash
   git init
   git add .
   git commit -m "Initial commit — media plan studio"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```
3. In the repository, go to **Settings → Pages**.
4. Under **Build and deployment → Source**, choose **Deploy from a branch**.
5. Choose the `main` branch and `/ (root)` folder, then **Save**.
6. GitHub will publish the site at
   `https://<your-username>.github.io/<your-repo>/` within a minute or two.

## Using the app

1. **Strategy Planner** — set the project name, duration, ticket size,
   monthly booking targets by source (Presales/Digital, Direct/Walk-in,
   Channel Partner), and the **Site Visits Target & Source Split**
   (directly editable — e.g. 20 bookings but 80 visits across 3 months).
   Use the quick presets to load the default 43-booking plan or the
   20-booking / 80-visit fast-track plan (6-6-8 monthly split).
   The **Historical Reporting Data** section feeds the MTD report,
   lifetime visit summary, agreement tracker and till-date advertising
   cost slides — edit any cell to reflect actuals.
2. **Presentation Deck** — browse all 17 slides live, matching the exact
   flow: cover, MTD report, lifetime summary, agreement tracker, till-date
   cost, way-forward divider, strategic targets, 90-day funnel, digital
   funnel & platform split, campaign bifurcation, audience matrix, OOH
   split, CP activation, creative showcase, festive offers, 90-day budget,
   and closing slide. Use the slide pills, arrows, or the Fullscreen
   button.
3. **Export PPTX** — click **Download Complete Presentation (.pptx)** to
   generate a genuine, styled 16:9 PowerPoint file with the same data as
   the live deck.

## Notes

- All calculations live in `js/calculator.js` and are pure functions —
  safe to unit test or reuse elsewhere.
- Numbers are formatted in Indian currency style (₹16,31,579 / ₹4.17 Cr).
- If a CDN script is blocked by your network, the app shows an in-page
  error explaining what to check, instead of a blank screen.
