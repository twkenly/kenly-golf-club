# Kenly Golf Club

Trip-planning site for the annual Kenly Golf Club trip, served by GitHub Pages at
https://twkenly.github.io/kenly-golf-club/

## Spring 2027

`index.html` is the current trip: three destination options with detail pages
(`#st-george`, `#scottsdale`, `#pebble-beach`), a survey and live results.

### Survey backend

Answers are stored in the Google Sheet "KGC 2027 Survey Responses" through a
Google Apps Script web app (`apps-script/Code.gs`). The site only receives totals
and comments; names and notes stay in the Sheet.

To redeploy the script: open the Sheet, Extensions > Apps Script, paste
`apps-script/Code.gs`, then Deploy > Manage deployments > Edit > New version.
The web app URL is set in `index.html` as `API_URL`.
