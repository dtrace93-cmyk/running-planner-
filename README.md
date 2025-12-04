# Running Planner

This repo hosts a single-page running planner with Strava sync support. The app includes dashboards, training logs, and analytics powered by client-side JavaScript.

## Progressive Web App
- A web app manifest (`manifest.webmanifest`) declares metadata and icons for installability.
- A service worker (`sw.js`) caches core assets so the planner works offline while avoiding Strava API requests.
- `index.html` registers the service worker and links to the manifest so the site can be installed as a standalone experience.

## Development
Open `index.html` directly in a browser or serve the repo via any static file host. Strava integration expects the backend endpoints referenced in the code to be available.
