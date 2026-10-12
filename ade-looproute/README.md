# ADE looproute (vrijdag 23 oktober 2026)

Interactieve kaart van één ADE-dag: klikbare pin per event, locatietags en de efficiëntste looproute binnen de tijdvensters.
Open `index.html` in een browser. Alles is ingebed behalve Leaflet (cdnjs) en de lettertypes (Google Fonts).

## Routelogica
- Je komt aan, wacht als je te vroeg bent, en moet minimaal X minuten (instelbaar: 30/45/60/90) binnen het tijdvenster kunnen blijven.
- Volgorde: eerst zoveel mogelijk events, dan zo min mogelijk loop- plus wachttijd, dan zo min mogelijk meters.
- Exacte zoektocht over (bezochte set, laatste stop) met Pareto-snoei; met 10 events gaat dat in milliseconden.

## Bronnen
- Adressen: BAG via PDOK Locatieserver.
- Basiskaart: Kadaster / PDOK TOP10NL (CC BY 4.0), vereenvoudigd in `data/basemap.json`.
- Looproutes: OSRM voetprofiel (routing.openstreetmap.de), © OpenStreetMap-bijdragers, in `data/legs.json`.

## Opnieuw bouwen
```
python3 build/fetch_pdok.py raw/            # TOP10NL-lagen ophalen (~42 MB, niet in git)
python3 build/basemap.py raw/ data/basemap.json   # vereist shapely
python3 build/routes.py data/legs.json
python3 build/assemble.py build/template.html leaflet.min.css data/basemap.json data/legs.json fragment.html index.html
```
