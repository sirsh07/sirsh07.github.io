# PFAS Dataset Explorer

Open `index.html` directly in a browser, or serve the workspace locally:

```bash
python3 -m http.server 8000
```

Then visit `http://localhost:8000/pfas_dataset_explorer/`.

The application is a static HTML/CSS/JavaScript dashboard. It does not require a build step or API key. Leaflet, Natural Earth country boundaries, U.S. state boundaries, dataset metadata, figures, and charts are stored locally. When internet access is available, OpenStreetMap tiles provide street/place detail; if tiles fail, the local country/state boundary map and all markers continue to work.

Features:

- world and United States map views;
- filters for human/environmental datasets and access status;
- searchable, sortable dataset cards;
- participant/record counts displayed using compatible units;
- detailed PFAS coverage, outcomes, conclusions, limitations and source links;
- shareable URL state for the selected dataset and main filters.

## Local map assets

- Leaflet 1.9.4: <https://leafletjs.com/> (BSD-2-Clause)
- Natural Earth 1:110m country boundaries: <https://github.com/nvkelso/natural-earth-vector> (public domain)
- U.S. state GeoJSON: <https://github.com/PublicaMundi/MappingAPI>
- Optional online basemap: <https://www.openstreetmap.org/copyright>

The research cards, figures, outcome table, and flowchart link directly to their supporting papers or official documentation inside the app.
