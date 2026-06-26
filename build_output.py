import pandas as pd, json

SRC = "/root/.claude/uploads/25fd9b56-1752-5edc-8721-3932cd50f9d6/70eb99d7-20260523_naming_rights_database_1.xlsx"

research = [
{"n":1,"country":"USA","city":"Minneapolis","state":"Minnesota","ownership_type":"Public — University of Minnesota","opened_year":1993,"capacity":10000,"source_url":"https://en.wikipedia.org/wiki/3M_Arena_at_Mariucci"},
{"n":2,"country":"USA","city":"Windsor","state":"Colorado","ownership_type":"Private — Future Legends LLC","opened_year":2023,"capacity":6500,"source_url":"https://en.wikipedia.org/wiki/Future_Legends_Complex"},
{"n":3,"country":"USA","city":"Beloit","state":"Wisconsin","ownership_type":"Private — Hendricks family / Hendricks Properties","opened_year":2021,"capacity":3850,"source_url":"https://en.wikipedia.org/wiki/ABC_Supply_Stadium"},
{"n":4,"country":"Canada","city":"North Battleford","state":"Saskatchewan","ownership_type":"Public — City of North Battleford","opened_year":1962,"capacity":2500,"source_url":"https://en.wikipedia.org/wiki/North_Battleford_Civic_Centre"},
{"n":5,"country":"USA","city":"Kent","state":"Washington","ownership_type":"Public — City of Kent","opened_year":2009,"capacity":6500,"source_url":"https://en.wikipedia.org/wiki/Accesso_ShoWare_Center"},
{"n":6,"country":"USA","city":"Thousand Palms","state":"California","ownership_type":"Private — Oak View Group","opened_year":2022,"capacity":10500,"source_url":"https://en.wikipedia.org/wiki/Acrisure_Arena"},
{"n":7,"country":"USA","city":"Orlando","state":"Florida","ownership_type":"Public — University of Central Florida","opened_year":2007,"capacity":44206,"source_url":"https://en.wikipedia.org/wiki/Acrisure_Bounce_House"},
{"n":8,"country":"USA","city":"Pittsburgh","state":"Pennsylvania","ownership_type":"Public — Sports & Exhibition Authority of Pittsburgh and Allegheny County","opened_year":2001,"capacity":68400,"source_url":"https://en.wikipedia.org/wiki/Acrisure_Stadium"},
{"n":9,"country":"USA","city":"Orlando","state":"Florida","ownership_type":"Public — University of Central Florida","opened_year":2007,"capacity":10000,"source_url":"https://ucfknights.com/facilities/addition-financial-arena"},
{"n":10,"country":"USA","city":"Beaverton","state":"Oregon","ownership_type":"Private — Portland Timbers (Peregrine Sports LLC)","opened_year":2012,"capacity":None,"source_url":"https://www.timbers.com/news/timbers-open-new-adidas-training-center"},
{"n":11,"country":"USA","city":"Utica","state":"New York","ownership_type":"Public — Upper Mohawk Valley Memorial Auditorium Authority","opened_year":1960,"capacity":3815,"source_url":"https://en.wikipedia.org/wiki/Adirondack_Bank_Center"},
{"n":12,"country":"USA","city":"Rome","state":"Georgia","ownership_type":"Public — City of Rome","opened_year":2003,"capacity":5105,"source_url":"https://en.wikipedia.org/wiki/AdventHealth_Stadium"},
{"n":13,"country":"USA","city":"Orlando","state":"Florida","ownership_type":"Private — Orlando Magic / RDV Sports","opened_year":2022,"capacity":None,"source_url":"https://www.nba.com/magic/news/orlando-magic-and-adventhealth-unveil-state-of-the-art-adventhealth-training-center-20220831"},
{"n":14,"country":"USA","city":"Stockton","state":"California","ownership_type":"Public — City of Stockton","opened_year":2005,"capacity":12000,"source_url":"https://en.wikipedia.org/wiki/Adventist_Health_Arena"},
{"n":15,"country":"USA","city":"Chicago","state":"Illinois","ownership_type":"Private — Chicago Bulls (Reinsdorf)","opened_year":2014,"capacity":None,"source_url":"https://www.hok.com/projects/view/advocate-center-chicago-bulls-practice-facility/"},
{"n":16,"country":"USA","city":"Davis","state":"California","ownership_type":"Public — University of California, Davis","opened_year":2007,"capacity":10743,"source_url":"https://en.wikipedia.org/wiki/Aggie_Stadium_(UC_Davis)"},
{"n":17,"country":"USA","city":"Seattle","state":"Washington","ownership_type":"Public — University of Washington","opened_year":1927,"capacity":10000,"source_url":"https://en.wikipedia.org/wiki/Hec_Edmundson_Pavilion"},
{"n":18,"country":"USA","city":"Anchorage","state":"Alaska","ownership_type":"Public — University of Alaska Anchorage","opened_year":2014,"capacity":5000,"source_url":"https://en.wikipedia.org/wiki/Alaska_Airlines_Center"},
{"n":19,"country":"USA","city":"Seattle","state":"Washington","ownership_type":"Public — University of Washington","opened_year":1920,"capacity":70138,"source_url":"https://en.wikipedia.org/wiki/Husky_Stadium"},
{"n":20,"country":"USA","city":"Boise","state":"Idaho","ownership_type":"Public — Boise State University","opened_year":1970,"capacity":36363,"source_url":"https://en.wikipedia.org/wiki/Albertsons_Stadium"},
{"n":21,"country":"USA","city":"Grand Forks","state":"North Dakota","ownership_type":"Public — City of Grand Forks","opened_year":2001,"capacity":12283,"source_url":"https://en.wikipedia.org/wiki/Alerus_Center"},
{"n":22,"country":"USA","city":"Fort Myers","state":"Florida","ownership_type":"Public — Florida Gulf Coast University","opened_year":2002,"capacity":4633,"source_url":"https://en.wikipedia.org/wiki/Alico_Arena"},
{"n":23,"country":"USA","city":"Winston-Salem","state":"North Carolina","ownership_type":"Private — Wake Forest University","opened_year":1968,"capacity":31500,"source_url":"https://en.wikipedia.org/wiki/Allegacy_Federal_Credit_Union_Stadium"},
{"n":24,"country":"USA","city":"Las Vegas (Paradise)","state":"Nevada","ownership_type":"Public — Las Vegas Stadium Authority","opened_year":2020,"capacity":65000,"source_url":"https://en.wikipedia.org/wiki/Allegiant_Stadium"},
{"n":25,"country":"USA","city":"Madison","state":"Wisconsin","ownership_type":"Public — Dane County","opened_year":1967,"capacity":10231,"source_url":"https://en.wikipedia.org/wiki/Alliant_Energy_Center"},
{"n":26,"country":"USA","city":"Cedar Rapids","state":"Iowa","ownership_type":"Public — City of Cedar Rapids","opened_year":1979,"capacity":6900,"source_url":"https://en.wikipedia.org/wiki/Alliant_Energy_PowerHouse"},
{"n":27,"country":"USA","city":"Saint Paul","state":"Minnesota","ownership_type":"Private — Minnesota United FC","opened_year":2019,"capacity":19400,"source_url":"https://en.wikipedia.org/wiki/Allianz_Field"},
{"n":28,"country":"USA","city":"Rosemont","state":"Illinois","ownership_type":"Public — Village of Rosemont","opened_year":1980,"capacity":16692,"source_url":"https://en.wikipedia.org/wiki/Allstate_Arena"},
{"n":29,"country":"USA","city":"Rockford","state":"Michigan","ownership_type":"Public — West Michigan Sports Commission","opened_year":2026,"capacity":None,"source_url":"https://www.westmisports.com/news/alro-steel-foundation-donates-15-million-to-meijer-sports-complex-expansion"},
{"n":30,"country":"USA","city":"Richmond","state":"Kentucky","ownership_type":"Public — Eastern Kentucky University","opened_year":1963,"capacity":6500,"source_url":"https://en.wikipedia.org/wiki/Alumni_Coliseum"},
{"n":31,"country":"USA","city":"Sunrise","state":"Florida","ownership_type":"Public — Broward County","opened_year":1998,"capacity":19250,"source_url":"https://en.wikipedia.org/wiki/Amerant_Bank_Arena"},
{"n":32,"country":"USA","city":"South Jordan (Daybreak)","state":"Utah","ownership_type":"Private — Miller Sports + Entertainment","opened_year":2025,"capacity":8000,"source_url":"https://www.mlb.com/milb/news/featured/visit-the-ballpark-at-america-first-square-home-of-the-salt-lake-bees"},
{"n":33,"country":"USA","city":"Cedar City","state":"Utah","ownership_type":"Public — Southern Utah University","opened_year":1985,"capacity":5300,"source_url":"https://en.wikipedia.org/wiki/America_First_Event_Center"},
{"n":34,"country":"USA","city":"Sandy","state":"Utah","ownership_type":"Private — Miller Sports + Entertainment","opened_year":2008,"capacity":20213,"source_url":"https://en.wikipedia.org/wiki/America_First_Field"},
{"n":35,"country":"USA","city":"Dallas","state":"Texas","ownership_type":"Public — City of Dallas","opened_year":2001,"capacity":19200,"source_url":"https://en.wikipedia.org/wiki/American_Airlines_Center"},
{"n":36,"country":"USA","city":"Corpus Christi","state":"Texas","ownership_type":"Public — City of Corpus Christi","opened_year":2004,"capacity":10000,"source_url":"https://en.wikipedia.org/wiki/American_Bank_Center"},
{"n":37,"country":"USA","city":"Milwaukee","state":"Wisconsin","ownership_type":"Public — Southeast Wisconsin Professional Baseball Park District","opened_year":2001,"capacity":41900,"source_url":"https://en.wikipedia.org/wiki/American_Family_Field"},
{"n":38,"country":"USA","city":"Phoenix","state":"Arizona","ownership_type":"Public — City of Phoenix","opened_year":1998,"capacity":7000,"source_url":"https://en.wikipedia.org/wiki/American_Family_Fields_of_Phoenix"},
{"n":39,"country":"USA","city":"Sioux Center","state":"Iowa","ownership_type":"Public — Dordt University & City of Sioux Center","opened_year":2023,"capacity":None,"source_url":"https://www.dordt.edu/about-dordt/offices-and-services/american-state-bank-sports-complex"},
{"n":40,"country":"USA","city":"Providence","state":"Rhode Island","ownership_type":"Public — Rhode Island Convention Center Authority","opened_year":1972,"capacity":11075,"source_url":"https://en.wikipedia.org/wiki/Amica_Mutual_Pavilion"},
{"n":41,"country":"USA","city":"Grand Rapids","state":"Michigan","ownership_type":"Public — Grand Rapids-Kent County Convention/Arena Authority","opened_year":None,"capacity":8500,"source_url":"https://en.wikipedia.org/wiki/Amway_Stadium"},
{"n":42,"country":"USA","city":"Everett","state":"Washington","ownership_type":"Public — Everett Public Facilities District","opened_year":2003,"capacity":8149,"source_url":"https://en.wikipedia.org/wiki/Angel_of_the_Winds_Arena"},
{"n":43,"country":"USA","city":"Springdale","state":"Arkansas","ownership_type":"Public — City of Springdale","opened_year":2008,"capacity":7305,"source_url":"https://en.wikipedia.org/wiki/Arvest_Ballpark"},
{"n":44,"country":"USA","city":"Nashville","state":"Tennessee","ownership_type":"Private — Tennessee Titans","opened_year":1999,"capacity":None,"source_url":"https://www.tennesseetitans.com/news/titans-practice-facility-renamed-saint-thomas-sports-park-10455494"},
{"n":45,"country":"USA","city":"Chattanooga","state":"Tennessee","ownership_type":"Public — City of Chattanooga","opened_year":2000,"capacity":6382,"source_url":"https://en.wikipedia.org/wiki/AT%26T_Field"},
{"n":46,"country":"USA","city":"Arlington","state":"Texas","ownership_type":"Public — City of Arlington","opened_year":2009,"capacity":80000,"source_url":"https://en.wikipedia.org/wiki/AT%26T_Stadium"},
{"n":47,"country":"USA","city":"Florham Park","state":"New Jersey","ownership_type":"Private — New York Jets","opened_year":2008,"capacity":None,"source_url":"https://en.wikipedia.org/wiki/Atlantic_Health_Jets_Training_Center"},
{"n":48,"country":"USA","city":"Harrisonburg","state":"Virginia","ownership_type":"Public — James Madison University","opened_year":2020,"capacity":8500,"source_url":"https://en.wikipedia.org/wiki/Atlantic_Union_Bank_Center"},
{"n":49,"country":"USA","city":"Kannapolis","state":"North Carolina","ownership_type":"Public — City of Kannapolis","opened_year":2021,"capacity":4930,"source_url":"https://en.wikipedia.org/wiki/Atrium_Health_Ballpark"},
{"n":50,"country":"USA","city":"Charlotte","state":"North Carolina","ownership_type":"Private — Tepper Sports & Entertainment","opened_year":None,"capacity":None,"source_url":"https://www.panthers.com/news/carolina-panthers-unveil-atrium-health-training-facility-naming-partnership"},
]

notes = {
 7:"Acrisure Bounce House = UCF football stadium (formerly FBC Mortgage Stadium). Distinct from #9 Addition Financial Arena (basketball).",
 9:"Addition Financial Arena = UCF basketball arena (formerly CFE Arena). Distinct from #7 football stadium.",
 24:"Stadium is in Paradise, NV (Las Vegas metro). Cap 65,000, expandable to ~72,000.",
 29:"Softball field within Meijer Sports Complex; opened with 2026 expansion; no fixed seating published.",
 38:"Originally opened 1998 as Maryvale Baseball Park; renamed 2019.",
 41:"Under construction (began 2025), expected to open ~2027; capacity planned ~8,500.",
 45:"Lookouts moved to new Erlanger Park (opened Apr 2026). AT&T Field (opened 2000) now retired.",
 46:"Cap 80,000 standard, expandable to ~100,000 for marquee events.",
 50:"Naming partnership announced 2026; facility expected to open ~2027.",
}

df = pd.read_excel(SRC, sheet_name="Data_Input", header=3).dropna(how="all").reset_index(drop=True)
f = df.head(50).copy()

rmap = {r["n"]: r for r in research}
for i in range(50):
    r = rmap[i+1]
    f.at[i, "Country"] = r["country"]
    f.at[i, "City"] = r["city"]
    f.at[i, "State / Prefecture"] = r["state"]
    f.at[i, "Ownership Type"] = r["ownership_type"]
    f.at[i, "Opened Year"] = r["opened_year"]
    f.at[i, "Capacity"] = r["capacity"]

f["Web Source URL (added)"] = [rmap[i+1]["source_url"] for i in range(50)]
f["Research Notes (added)"] = [notes.get(i+1, "") for i in range(50)]

# Compact view for readability
view_cols = ["Record ID","Venue Name","Venue Type","Primary Tenant / Team","Sponsor Name",
             "Country","City","State / Prefecture","Ownership Type","Opened Year","Capacity",
             "Web Source URL (added)","Research Notes (added)"]
view = f[view_cols]

out_xlsx = "/home/user/NYU_MSBAi/first_50_venues_enriched.xlsx"
with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
    f.to_excel(w, sheet_name="Full_Enriched", index=False)
    view.to_excel(w, sheet_name="Summary", index=False)
view.to_csv("/home/user/NYU_MSBAi/first_50_venues_enriched.csv", index=False)

print("Wrote", out_xlsx)
print(view.to_string(index=False, max_colwidth=28))
