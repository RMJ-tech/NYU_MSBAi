# First 50 Venues — Web-Enriched

Extracted the first 50 venues from `Data_Input` of the naming-rights research
database and filled the empty venue-profile fields from public web sources.

## Files
- `first_50_venues_enriched.xlsx` — `Summary` sheet (key columns) + `Full_Enriched` sheet (all original columns with fields populated).
- `first_50_venues_enriched.csv` — Summary sheet as CSV.
- `build_output.py` — reproducible build script with all researched values and source URLs inline.

## Fields filled (were 100% empty in the source for these rows)
Country, City, State / Prefecture, Ownership Type, Opened Year, Capacity.

Two columns were added:
- `Web Source URL (added)` — the specific page used for each venue (Wikipedia preferred; official team/venue/news pages otherwise).
- `Research Notes (added)` — caveats (disambiguation, expandable capacities, under-construction venues, etc.).

## Caveats
- Capacities reflect the primary configuration (e.g. hockey vs. concert); some venues are expandable.
- A few venues are under construction or recently renamed (e.g. Amway Stadium, Atrium Health Training Facility) — `Opened Year` is blank where not yet open.
- Acrisure Bounce House (#7, UCF football) and Addition Financial Arena (#9, UCF basketball) are two distinct buildings.
- Values were collected ~June 2026 and should be re-verified against the cited source before use in analysis.
