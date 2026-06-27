# Reusable Prompt — Naming-Rights Venue Enrichment (any row range)

Use this to have Claude enrich more venues from the master dataset in the **exact
same audit style** as `first_50_venues_filled.xlsx`. The dataset has ~642 data
rows total (`Data_Input` sheet, header on row 4), so you can run it in batches:
51–100, 101–150, 151–200, … up to 642.

## What to attach (every time)
1. **The original master workbook** — `20260523_naming_rights_database_1.xlsx`
   (the file you first uploaded). This is the source of truth for the venue rows
   and the 100-field schema.
2. **One finished output as the style template** — `first_50_venues_filled.xlsx`
   (or any later batch). Claude mirrors its layout, colors, and field logic.

> If you're running in *this same repo*, the build scripts and per-venue research
> JSON are already committed (`build_filled_50.py`, `research/` notes), so Claude
> can adapt them directly instead of rebuilding from scratch.

## The prompt to paste (edit the row range)

> Using the two attached files, enrich **venues 51–100** from the `Data_Input`
> sheet of the master workbook, and produce a workbook in the **identical style
> and logic** as `first_50_venues_filled.xlsx`.
>
> For each venue, research the web and fill every blank "Populated Value" field
> (treat all missing fields as approved for external input). Follow the same
> rules used for the first 50:
> - One sheet per venue (named `<RecordID>_<VenueName>`), plus an `Index` sheet.
> - The same 8-column audit table: Col Index · Category Block · Field Name ·
>   Original Group Priority · Populated Value · Original Raw Value · Needs
>   External Input? · Source & Rationale Context.
> - Color coding: **green/"Completed"** = filled from web or derivation (put the
>   source URL or basis in the rationale column); **yellow/"Yes"** = still blank
>   after a genuine external-input attempt (state why); **white/"No"** =
>   pre-existing value from the original extract.
> - Cite a real source URL for every web-sourced fact (Wikipedia, official
>   team/venue sites, Census Reporter, reputable news).
> - For metrics that aren't publicly disclosed (social/media impressions, mobile
>   foot traffic, search volume, Google Trends index, engagement rate, broadcast
>   viewers, media mentions, social posts), give a **reasoned estimate** and
>   prefix the rationale with `Estimate:` plus a one-line method — do not invent
>   it as a hard fact, and don't leave it blank.
> - Derive the classification/calculated fields deterministically (Region,
>   Subregion, Public/Private Facility, Rights Type, Naming Scope, Value
>   Disclosure Status, Contract Start Year, Attendance per Seat, Utilization
>   Rate, local/English venue name, original language, data basis).
> - For practice/training facilities and not-yet-open venues, mark attendance/
>   event/broadcast/social fields "Not applicable" rather than estimating.
> - Note any corrections (renamings, wrong league/tenant, etc.) in each sheet.
>
> Deliver the `.xlsx` file. Commit and push to the branch.

## The 42 research fields collected per venue (web + estimates)
prev_name, metro_area, latitude, longitude, renovated_year, total_site_area_sqm,
nearest_station_airport, distance_to_cbd_km, sponsor_hq_country,
contract_announcement_date, league_competition, avg_home_attendance,
annual_venue_attendance, annual_event_days, sports_event_days, concert_days,
other_event_days, broadcast_viewers_annual, national_broadcast_events,
intl_exposure_yn, media_mentions_annual, est_media_impressions, social_followers,
annual_social_posts, social_impressions_annual, avg_engagement_rate,
google_trends_index, search_volume_venue, mobile_foot_traffic_annual,
city_population, metro_population, regional_gdp_usd, median_household_income_usd,
annual_tourists_city, num_large_corp_hq, exclusivity_yn, renewal_yn,
termination_yn, facility_governance, bundled_assets_flag, sponsor_venue_fit_notes,
special_factors

(Plus the 6 core fields filled in the first pass: Country, City, State/Prefecture,
Ownership Type, Opened Year, Capacity.)

## How Claude runs it (so you know what to expect)
- It spawns one research sub-agent per venue (each does ~6–7 web lookups and
  writes a `vNN.json`), in batches of ~10. A higher-limit account lets more run
  before hitting a cap.
- It then runs a build script that merges the JSON into the audit-styled workbook.
- If a run is interrupted by a limit, just say **"continue from where you left
  off"** — completed `vNN.json` files are reused, so only the missing venues are
  re-researched.

## Tips for the higher-limit account
- You can ask for a **bigger batch** (e.g. "do 51–150 in one go"). Claude will
  still wave it in groups of ~10 agents to stay within concurrency limits.
- If you want it fully unattended, add: *"keep going through all remaining
  batches until the whole dataset (~642 rows) is done, committing each batch."*
