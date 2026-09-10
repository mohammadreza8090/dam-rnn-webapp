# Dam Displacement Prediction — Client-Side RNN App

A fully static, client-side web app that trains a SimpleRNN model on dam
monitoring data (crest/pendulum displacement, optionally with reservoir
water level and air temperature) and predicts displacement — entirely in
the browser, with **no backend and no server to maintain**.

Built for the ICOLD 2027 paper demo.

## How it works

- **SheetJS** reads the uploaded `.xlsx` file directly in the browser.
- **TensorFlow.js** builds, trains, and runs the RNN model in the browser
  (no Python, no GPU server).
- **Plotly.js** renders the interactive charts.

Because everything runs client-side, this app can be hosted as static
files on GitHub Pages (or any static host) and will keep working
permanently, with zero ongoing maintenance.

## Files

- `index.html` — page structure and controls
- `style.css` — styling
- `app.js` — all application logic (data parsing, scaling, model, training, plotting)

## Run locally

Just open `index.html` in a browser — no build step, no install required.
(Some browsers restrict local file access for security; if the file
upload doesn't work when opening `index.html` directly, run a tiny local
server instead, e.g. `python -m http.server` in this folder, then visit
`http://localhost:8000`.)

## Deploy to GitHub Pages (permanent, free hosting)

1. Create a new GitHub repository (e.g. `dam-rnn-webapp`).
2. Upload `index.html`, `style.css`, and `app.js` to the repository root.
3. Go to **Settings → Pages**.
4. Under "Build and deployment", set **Source** to "Deploy from a branch",
   choose the `main` branch and the `/ (root)` folder, then save.
5. After a minute or two, GitHub gives you a permanent URL like:
   `https://<your-username>.github.io/dam-rnn-webapp/`
6. That link works forever, with no server to keep running — reference it
   directly in the paper or demo.

## Data format

Two supported modes (selectable in the app):

- **Univariate**: Date column + one target column (the instrument/pendulum
  reading).
- **Multivariate**: Date column + target column + water level column +
  temperature column.

Any column names are supported — pick them from the dropdowns after
uploading the file.
