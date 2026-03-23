"""
=============================================================
SignalBleed Site Generator
=============================================================

HOW TO USE:
1. Open a terminal in VS Code or a GitHub Codespace
2. Run:   pip install pandas requests openpyxl jinja2
3. Run:   python build.py
4. Wait ~5 minutes
5. Everything goes into the /output folder
6. Copy the contents of /output into your repo root
7. Commit and push

That's it.
=============================================================
"""

import os, re, time, zipfile, io
import pandas as pd
import requests
from pathlib import Path
from string import Template

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
OUTPUT = Path("output")
SITE   = "https://signalbleed.com"
BLS_URL = "https://www.bls.gov/oes/special.requests/oesm24all.zip"
RELATED_COUNT = 5
WIKI_DELAY = 0.15

PLAUSIBLE = """<!-- Privacy-friendly analytics by Plausible -->
<script async src="https://plausible.io/js/pa-zL2LxD5lIwlbUZGyyXwfO.js"></script>
<script>
  window.plausible=window.plausible||function(){(plausible.q=plausible.q||[]).push(arguments)},plausible.init=plausible.init||function(i){plausible.o=i||{}};
  plausible.init()
</script>"""


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def slugify(text):
    t = text.lower().strip()
    t = re.sub(r"[^a-z0-9\s-]", "", t)
    t = re.sub(r"[\s]+", "-", t)
    return re.sub(r"-+", "-", t).strip("-")

def city_slug(area):
    m = re.match(r"^([^-,]+).*,\s*([A-Z]{2})", str(area))
    if m:
        c = re.sub(r"[^a-z0-9\s]", "", m.group(1).strip().lower())
        return re.sub(r"\s+", "-", c) + "-" + m.group(2).strip().lower()
    return slugify(str(area))

def city_name(area):
    m = re.match(r"^([^-,]+).*,\s*([A-Z]{2})", str(area))
    return f"{m.group(1).strip()}, {m.group(2)}" if m else str(area)

def fmt(val):
    return f"${val:,.0f}" if pd.notna(val) else "N/A"

def fmth(val):
    return f"${val:,.2f}" if pd.notna(val) else "N/A"


# ---------------------------------------------------------------------------
# STEP 1: DOWNLOAD BLS DATA
# ---------------------------------------------------------------------------
def download():
    cache = Path("oesm24all.zip")
    if cache.exists():
        print("  Using cached BLS data")
        return cache
    print("  Downloading BLS OEWS May 2024 data (~45MB)...")
    r = requests.get(BLS_URL, stream=True)
    r.raise_for_status()
    with open(cache, "wb") as f:
        for chunk in r.iter_content(8192):
            f.write(chunk)
    print("  Done")
    return cache


# ---------------------------------------------------------------------------
# STEP 2: LOAD + CLEAN
# ---------------------------------------------------------------------------
def load(zippath):
    print("  Loading data...")
    # BLS puts an xlsx inside the zip
    with zipfile.ZipFile(zippath) as z:
        # Find the xlsx or csv file
        names = z.namelist()
        datafile = None
        for n in names:
            if n.endswith(".xlsx"):
                datafile = n
                break
        if not datafile:
            for n in names:
                if n.endswith(".csv"):
                    datafile = n
                    break
        if not datafile:
            raise Exception(f"No data file found in zip. Contents: {names}")

        print(f"  Reading {datafile}...")
        with z.open(datafile) as f:
            if datafile.endswith(".xlsx"):
                df = pd.read_excel(io.BytesIO(f.read()), engine="openpyxl")
            else:
                df = pd.read_csv(f)

    df.columns = df.columns.str.strip().str.upper()

    # Filter: metro areas + detailed occupations only
    df = df[
        (df["AREA_TYPE"].astype(str) == "2") &
        (df["O_GROUP"].astype(str).str.lower() == "detailed")
    ].copy()

    for col in ["A_MEAN", "A_PCT10", "A_PCT90", "H_MEAN", "TOT_EMP"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["A_MEAN"])
    df["job_slug"] = df["OCC_TITLE"].apply(slugify)
    df["city_slug"] = df["AREA_TITLE"].apply(city_slug)
    df["city_name"] = df["AREA_TITLE"].apply(city_name)

    print(f"  {len(df):,} rows | {df['OCC_TITLE'].nunique()} jobs | {df['AREA_TITLE'].nunique()} cities")
    return df


# ---------------------------------------------------------------------------
# STEP 3: NATIONAL AVERAGES
# ---------------------------------------------------------------------------
def national_avgs(df):
    print("  Computing national averages...")
    if "TOT_EMP" in df.columns:
        def wmean(g):
            v = g.dropna(subset=["A_MEAN","TOT_EMP"])
            if len(v)==0: return g["A_MEAN"].mean()
            return (v["A_MEAN"]*v["TOT_EMP"]).sum()/v["TOT_EMP"].sum()
        na = df.groupby("OCC_CODE").apply(wmean).reset_index()
        na.columns = ["OCC_CODE","natl_mean"]
    else:
        na = df.groupby("OCC_CODE")["A_MEAN"].mean().reset_index()
        na.columns = ["OCC_CODE","natl_mean"]
    return na


# ---------------------------------------------------------------------------
# STEP 4: CITY IMAGES (Wikipedia API)
# ---------------------------------------------------------------------------
def fetch_images(df):
    imgdir = OUTPUT / "images" / "cities"
    imgdir.mkdir(parents=True, exist_ok=True)
    cities = df[["city_slug","city_name"]].drop_duplicates("city_slug")
    total = len(cities)
    imap = {}
    print(f"  Fetching Wikipedia images for {total} cities...")

    for i,(_, row) in enumerate(cities.iterrows()):
        s = row["city_slug"]
        p = imgdir / f"{s}.jpg"
        if p.exists():
            imap[s] = f"/images/cities/{s}.jpg"
            continue
        try:
            r = requests.get("https://en.wikipedia.org/w/api.php", params={
                "action":"query","titles":row["city_name"],"prop":"pageimages",
                "format":"json","pithumbsize":800,"redirects":1
            }, timeout=10)
            pages = r.json().get("query",{}).get("pages",{})
            url = None
            for pg in pages.values():
                if "thumbnail" in pg:
                    url = pg["thumbnail"]["source"]
                    break
            if url:
                img = requests.get(url, timeout=10)
                img.raise_for_status()
                with open(p,"wb") as f: f.write(img.content)
                imap[s] = f"/images/cities/{s}.jpg"
            else:
                imap[s] = None
        except:
            imap[s] = None
        if (i+1) % 50 == 0:
            print(f"    {i+1}/{total}")
        time.sleep(WIKI_DELAY)

    got = sum(1 for v in imap.values() if v)
    print(f"  Done: {got}/{total} images")
    return imap


# ---------------------------------------------------------------------------
# SHARED HTML PIECES
# ---------------------------------------------------------------------------
FONTS = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600&family=Inter+Tight:wght@600;700;800;900&display=swap" rel="stylesheet">'

CSS = """*, *::before, *::after { margin:0; padding:0; box-sizing:border-box; }
:root {
  --bg:#F4F4F5;--surface:#FFF;--surface-alt:#FAFAFA;--text:#18181B;--text-muted:#71717A;
  --accent:#DC2626;--accent-light:#FEF2F2;--accent-border:#FECACA;
  --positive:#16A34A;--negative:#DC2626;--border:#E4E4E7;--border-subtle:#F4F4F5;
  --heading:'Inter Tight',sans-serif;--body:'DM Sans',sans-serif;--max-w:720px;--radius:6px;
}
body{font-family:var(--body);background:var(--bg);color:var(--text);line-height:1.6;font-size:15px;-webkit-font-smoothing:antialiased}
a{color:var(--text);text-decoration:underline;text-decoration-color:var(--border);text-underline-offset:2px}
a:hover{text-decoration-color:var(--text)}
.nav{background:var(--surface);border-bottom:1px solid var(--border)}
.nav-inner{max-width:1080px;margin:0 auto;padding:0 24px;height:52px;display:flex;align-items:center;justify-content:space-between}
.logo{font-family:var(--heading);font-weight:900;font-size:17px;color:var(--text);text-decoration:none;letter-spacing:-0.04em;text-transform:uppercase}
.logo span{color:var(--accent)}
.nav-links{display:flex;gap:24px;list-style:none}
.nav-links a{font-size:13px;font-weight:600;color:var(--text-muted);text-decoration:none;text-transform:uppercase;letter-spacing:0.03em}
.nav-links a:hover{color:var(--text)}
.footer{border-top:1px solid var(--border);background:var(--surface);padding:28px 24px;text-align:center;font-size:13px;color:var(--text-muted)}
.footer a{color:var(--text-muted);text-decoration:none;margin:0 12px}
.footer a:hover{text-decoration:underline}"""

NAV = '<nav class="nav"><div class="nav-inner"><a href="/" class="logo">SIGNAL<span>BLEED</span></a><ul class="nav-links"><li><a href="/jobs/">Jobs</a></li><li><a href="/cities/">Cities</a></li></ul></div></nav>'

FOOTER = '<footer class="footer"><span>&copy; 2025 SignalBleed</span><a href="/about.html">About</a><a href="/methodology.html">Methodology</a><a href="/jobs/">All Jobs</a><a href="/cities/">All Cities</a></footer>'

COMMUNITY_LEADERBOARD = '''<a href="https://www.yearup.org/" target="_blank" rel="noopener" class="community-slot cs-leaderboard">
<div class="community-slot-icon"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4.26 10.147a60.438 60.438 0 0 0-.491 6.347A48.62 48.62 0 0 1 12 20.904a48.62 48.62 0 0 1 8.232-4.41 60.46 60.46 0 0 0-.491-6.347m-15.482 0a50.636 50.636 0 0 0-2.658-.813A59.906 59.906 0 0 1 12 3.493a59.903 59.903 0 0 1 10.399 5.84c-.896.248-1.783.52-2.658.814m-15.482 0A50.717 50.717 0 0 1 12 13.489a50.702 50.702 0 0 1 7.74-3.342"/></svg></div>
<div class="community-slot-text"><div class="community-slot-label">Community Spotlight</div><div class="community-slot-title">Year Up &mdash; Closing the Opportunity Divide</div><div class="community-slot-desc">Free job training and internships for young adults ages 18&ndash;29.</div></div></a>'''

COMMUNITY_RECTANGLE = '''<a href="https://www.freecodecamp.org/" target="_blank" rel="noopener" class="community-slot cs-rectangle">
<div class="community-slot-icon"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M17.25 6.75 22.5 12l-5.25 5.25m-10.5 0L1.5 12l5.25-5.25m7.5-3-4.5 16.5"/></svg></div>
<div class="community-slot-text"><div class="community-slot-label">Free Resource</div><div class="community-slot-title">freeCodeCamp</div><div class="community-slot-desc">Free coding certifications for anyone looking to break into tech.</div></div></a>'''

SALARY_CSS_EXTRA = """.page{max-width:var(--max-w);margin:0 auto;padding:0 24px}
.breadcrumb{padding:14px 0 0;font-size:13px;color:var(--text-muted)}
.breadcrumb a{color:var(--text-muted);text-decoration:none}.breadcrumb a:hover{text-decoration:underline}
.breadcrumb .sep{margin:0 5px;opacity:0.4}
.hero-image{margin:16px 0;width:100%;height:220px;border-radius:var(--radius);overflow:hidden;background:var(--text);display:flex;align-items:center;justify-content:center;color:rgba(255,255,255,0.4);font-family:var(--heading);font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.06em}
.hero-image img{width:100%;height:100%;object-fit:cover;display:block}
.hero-image.has-img{background:var(--border)}
.hero h1{font-family:var(--heading);font-weight:900;font-size:36px;line-height:1.08;margin-bottom:8px;letter-spacing:-0.035em;text-transform:uppercase}
.hero h1 .city{white-space:nowrap}
.hero-subtitle{font-size:14px;color:var(--text-muted);font-weight:500;margin-bottom:24px}
.data-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border-radius:var(--radius);overflow:hidden;margin-bottom:28px}
.data-card{background:var(--surface);padding:18px 16px}
.data-label{font-family:var(--heading);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:var(--text-muted);margin-bottom:8px}
.data-value{font-family:var(--heading);font-weight:800;font-size:26px;line-height:1;letter-spacing:-0.02em}
.data-note{font-size:13px;color:var(--text-muted);margin-top:6px;font-weight:500}
.data-note.positive{color:var(--positive);font-weight:600}
.data-note.negative{color:var(--negative);font-weight:600}
.community-slot{border-radius:var(--radius);margin:0 auto 24px;overflow:hidden;text-decoration:none;display:flex;align-items:center;gap:16px;color:var(--text);transition:opacity 0.15s}
.community-slot:hover{opacity:0.85;text-decoration:none}
.community-slot-icon{flex-shrink:0;width:48px;height:48px;border-radius:50%;display:flex;align-items:center;justify-content:center}
.community-slot-icon svg{width:22px;height:22px}
.community-slot-text{flex:1}
.community-slot-label{font-family:var(--heading);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:var(--text-muted);margin-bottom:2px}
.community-slot-title{font-weight:600;font-size:14px;line-height:1.3}
.community-slot-desc{font-size:13px;color:var(--text-muted);margin-top:2px}
.cs-leaderboard{background:#EFF6FF;border:1px solid #BFDBFE;padding:16px 20px;max-width:728px;width:100%}
.cs-leaderboard .community-slot-icon{background:#DBEAFE}.cs-leaderboard .community-slot-icon svg{color:#2563EB}
.cs-rectangle{background:#F0FDF4;border:1px solid #BBF7D0;padding:20px;max-width:300px;width:100%;flex-direction:column;text-align:center;gap:12px}
.cs-rectangle .community-slot-icon{background:#DCFCE7}.cs-rectangle .community-slot-icon svg{color:#16A34A}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:24px;margin-bottom:20px}
.card h2{font-family:var(--heading);font-weight:800;font-size:18px;margin-bottom:14px;letter-spacing:-0.02em;text-transform:uppercase}
.range-desc{font-size:15px;color:var(--text-muted);margin-bottom:28px;line-height:1.5}
.range-bar-wrap{padding-top:36px;margin-bottom:14px}
.range-bar{position:relative;height:8px;background:var(--border);border-radius:4px}
.range-fill{position:absolute;top:0;bottom:0;left:0;right:0;background:var(--accent-light);border-radius:4px;border:1px solid var(--accent-border)}
.range-marker{position:absolute;top:-10px;width:3px;height:28px;background:var(--accent);border-radius:2px;transform:translateX(-50%)}
.range-marker-label{position:absolute;bottom:100%;left:50%;transform:translateX(-50%);margin-bottom:10px;font-family:var(--heading);font-size:13px;font-weight:800;color:var(--accent);white-space:nowrap}
.range-endpoints{display:flex;justify-content:space-between;margin-top:12px}
.range-endpoints span{font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.03em}
.range-endpoints strong{display:block;font-family:var(--heading);font-size:16px;font-weight:800;color:var(--text);margin-top:3px}
.compare-table{width:100%;border-collapse:collapse}
.compare-table th,.compare-table td{text-align:left;padding:10px 0;border-bottom:1px solid var(--border-subtle);font-size:15px}
.compare-table th{font-family:var(--heading);font-weight:700;color:var(--text-muted);font-size:11px;text-transform:uppercase;letter-spacing:0.05em;padding-bottom:12px}
.compare-table td:last-child,.compare-table th:last-child{text-align:right;font-family:var(--heading);font-weight:700;letter-spacing:-0.01em}
.compare-table tr.current td{font-weight:700}
.compare-table tr.current{background:var(--accent-light)}
.compare-table tr.current td:first-child{position:relative;padding-left:12px}
.compare-table tr.current td:first-child::before{content:'';position:absolute;left:0;top:4px;bottom:4px;width:3px;background:var(--accent);border-radius:2px}
.related-grid{display:grid;grid-template-columns:1fr 1fr;gap:28px}
.related-group h3{font-family:var(--heading);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:var(--text-muted);margin-bottom:12px}
.related-group ul{list-style:none}
.related-group li{margin-bottom:8px;display:flex;justify-content:space-between;align-items:baseline}
.related-group li a{font-size:15px;text-decoration:none;color:var(--text)}.related-group li a:hover{text-decoration:underline}
.salary-hint{color:var(--text-muted);font-size:14px;font-family:var(--heading);font-weight:700;letter-spacing:-0.01em}
.source-box{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-bottom:48px}
.source-box h3{font-family:var(--heading);font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:0.05em;color:var(--text-muted);margin-bottom:8px}
.source-box p{font-size:14px;color:var(--text-muted);line-height:1.6}.source-box a{color:var(--text-muted)}
@media(max-width:600px){.hero h1{font-size:28px}.hero-image{height:160px}.data-grid{grid-template-columns:repeat(2,1fr)}.data-value{font-size:22px}.related-grid{grid-template-columns:1fr;gap:20px}}"""

INDEX_CSS_EXTRA = """.page{max-width:1080px;margin:0 auto;padding:32px 24px 64px}
.page h1{font-family:var(--heading);font-weight:900;font-size:32px;text-transform:uppercase;letter-spacing:-0.03em;margin-bottom:24px}
.item-list{list-style:none}
.item-list li{display:flex;justify-content:space-between;align-items:baseline;padding:10px 0;border-bottom:1px solid var(--border-subtle)}
.item-list li a{font-size:15px;text-decoration:none;color:var(--text)}.item-list li a:hover{text-decoration:underline}
.item-list .val{font-family:var(--heading);font-weight:700;font-size:14px;color:var(--text-muted)}
.city-list{list-style:none;columns:3;column-gap:32px}
.city-list li{padding:6px 0;break-inside:avoid}
.city-list li a{font-size:15px;text-decoration:none;color:var(--text)}.city-list li a:hover{text-decoration:underline}
@media(max-width:600px){.city-list{columns:1}}"""


# ---------------------------------------------------------------------------
# PAGE BUILDERS
# ---------------------------------------------------------------------------
def head(title, desc, canonical, extra_css=""):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{canonical}">
{FONTS}
{PLAUSIBLE}
<style>{CSS}
{extra_css}</style>
</head>
<body>
{NAV}"""


def foot():
    return f"""{FOOTER}
</body>
</html>"""


def salary_page(row, df, natl, imap):
    mean = row["A_MEAN"]
    nm = natl.get(row["OCC_CODE"], mean)
    p10 = row.get("A_PCT10")
    p90 = row.get("A_PCT90")
    hr = row.get("H_MEAN")
    diff = round(((mean-nm)/nm)*100, 1) if nm > 0 else 0
    diff_cls = "positive" if diff >= 0 else "negative"
    diff_str = f"+{diff}" if diff >= 0 else str(diff)

    if pd.notna(p10) and pd.notna(p90) and p90 > p10:
        marker = round(((mean-p10)/(p90-p10))*100, 1)
    else:
        marker = 50

    img = imap.get(row["city_slug"])
    if img:
        hero_img = f'<div class="hero-image has-img"><img src="{img}" alt="{row["city_name"]}" loading="lazy"></div>'
    else:
        hero_img = f'<div class="hero-image">{row["city_name"]}</div>'

    # Related cities (same job)
    sj = df[(df["OCC_CODE"]==row["OCC_CODE"])&(df["city_slug"]!=row["city_slug"])].nlargest(RELATED_COUNT,"A_MEAN")
    rc = "".join(f'<li><a href="/salary/{row["job_slug"]}/{r["city_slug"]}.html">{r["city_name"]}</a><span class="salary-hint">{fmt(r["A_MEAN"])}</span></li>' for _,r in sj.iterrows())

    # Related jobs (same city)
    sc = df[(df["city_slug"]==row["city_slug"])&(df["OCC_CODE"]!=row["OCC_CODE"])].nlargest(RELATED_COUNT,"A_MEAN")
    rj = "".join(f'<li><a href="/salary/{r["job_slug"]}/{row["city_slug"]}.html">{r["OCC_TITLE"]}</a><span class="salary-hint">{fmt(r["A_MEAN"])}</span></li>' for _,r in sc.iterrows())

    # Comparison table
    comp = sj.head(4)
    comp_rows = f'<tr class="current"><td>{row["city_name"]}</td><td>{fmt(mean)}</td></tr>\n<tr><td>National Average</td><td>{fmt(nm)}</td></tr>'
    comp_rows += "".join(f'\n<tr><td>{r["city_name"]}</td><td>{fmt(r["A_MEAN"])}</td></tr>' for _,r in comp.iterrows())

    cshort = row["city_name"].split(",")[0]
    cs_nbs = row["city_name"].replace(", ", ",&nbsp;")
    jt = row["OCC_TITLE"]
    js = row["job_slug"]
    cslu = row["city_slug"]
    at = row["AREA_TITLE"]

    h = head(
        f"{jt} Salary in {row['city_name']} — Average Pay &amp; Range | SignalBleed",
        f"Average salary for a {jt} in {row['city_name']} is {fmt(mean)}. BLS OEWS data.",
        f"{SITE}/salary/{js}/{cslu}.html",
        SALARY_CSS_EXTRA
    )

    schema = f'''<script type="application/ld+json">
{{"@context":"https://schema.org","@type":"WebPage","name":"{jt} Salary in {row['city_name']}","description":"Average salary for a {jt} in {row['city_name']} is {fmt(mean)} based on BLS OEWS data.","url":"{SITE}/salary/{js}/{cslu}.html","publisher":{{"@type":"Organization","name":"SignalBleed"}}}}
</script>'''

    return f"""{h}
{schema}
<div class="page">
<div class="breadcrumb"><a href="/">Home</a><span class="sep">/</span><a href="/jobs/">Jobs</a><span class="sep">/</span><a href="/jobs/{js}.html">{jt}</a><span class="sep">/</span>{row['city_name']}</div>
{hero_img}
<section class="hero">
<h1>{jt} Salary in <span class="city">{cs_nbs}</span></h1>
<p class="hero-subtitle">BLS Occupational Employment &amp; Wage Statistics &middot; May 2024</p>
</section>
<div class="data-grid">
<div class="data-card"><div class="data-label">Mean Annual</div><div class="data-value">{fmt(mean)}</div><div class="data-note {diff_cls}">{diff_str}% vs national</div></div>
<div class="data-card"><div class="data-label">Hourly Wage</div><div class="data-value">{fmth(hr)}</div><div class="data-note">&nbsp;</div></div>
<div class="data-card"><div class="data-label">10th Percentile</div><div class="data-value">{fmt(p10)}</div><div class="data-note">Entry level</div></div>
<div class="data-card"><div class="data-label">90th Percentile</div><div class="data-value">{fmt(p90)}</div><div class="data-note">Senior level</div></div>
</div>
{COMMUNITY_LEADERBOARD}
<section class="card">
<h2>Salary Range</h2>
<p class="range-desc">{jt}s in {cshort} earn between {fmt(p10)} and {fmt(p90)}, with a mean of {fmt(mean)}. The national average for this role is {fmt(nm)}.</p>
<div class="range-bar-wrap"><div class="range-bar"><div class="range-fill"></div><div class="range-marker" style="left:{marker}%;"><span class="range-marker-label">{fmt(mean)} mean</span></div></div></div>
<div class="range-endpoints"><span>10th Percentile<strong>{fmt(p10)}</strong></span><span style="text-align:right;">90th Percentile<strong>{fmt(p90)}</strong></span></div>
</section>
<section class="card">
<h2>How {cshort} Compares</h2>
<table class="compare-table"><thead><tr><th>Location</th><th>Mean Salary</th></tr></thead><tbody>
{comp_rows}
</tbody></table>
</section>
{COMMUNITY_RECTANGLE}
<section class="card">
<h2>Related Salaries</h2>
<div class="related-grid">
<div class="related-group"><h3>Same Job, Other Cities</h3><ul>{rc}</ul></div>
<div class="related-group"><h3>Same City, Other Jobs</h3><ul>{rj}</ul></div>
</div>
</section>
<div class="source-box"><h3>Data Source</h3><p><a href="https://www.bls.gov/oes/" target="_blank" rel="noopener">Bureau of Labor Statistics</a> OEWS, May 2024. {at}. Hourly wages assume 2,080 hours/year.</p></div>
</div>
{foot()}"""


def job_index_page(job_summary):
    items = "".join(f'<li><a href="/jobs/{r["job_slug"]}.html">{r["OCC_TITLE"]}</a><span class="val">{fmt(r["natl_mean"])}</span></li>' for _,r in job_summary.iterrows())
    h = head("All Jobs — Salary Data by Occupation | SignalBleed", f"Browse salary data for {len(job_summary)} occupations.", f"{SITE}/jobs/", INDEX_CSS_EXTRA)
    return f"""{h}<div class="page"><h1>All Jobs</h1><ul class="item-list">{items}</ul></div>{foot()}"""


def city_index_page(city_summary):
    items = "".join(f'<li><a href="/cities/{r["city_slug"]}.html">{r["city_name"]}</a></li>' for _,r in city_summary.iterrows())
    h = head("All Cities — Salary Data by Metro Area | SignalBleed", f"Browse salary data for {len(city_summary)} metro areas.", f"{SITE}/cities/", INDEX_CSS_EXTRA)
    return f"""{h}<div class="page"><h1>All Cities</h1><ul class="city-list">{items}</ul></div>{foot()}"""


def job_page(jt, js, nm_fmt, cities_data):
    rows = "".join(f'<tr><td><a href="/salary/{js}/{r["city_slug"]}.html">{r["city_name"]}</a></td><td>{fmt(r["A_MEAN"])}</td></tr>' for _,r in cities_data.iterrows())
    h = head(f"{jt} Salary by City | SignalBleed", f"Compare {jt} salaries across {len(cities_data)} cities. National avg: {nm_fmt}.", f"{SITE}/jobs/{js}.html", INDEX_CSS_EXTRA)
    return f"""{h}<div class="page"><h1>{jt} Salary by City</h1><p style="font-size:14px;color:var(--text-muted);margin-bottom:24px">National average: {nm_fmt} &middot; {len(cities_data)} metro areas</p><table class="compare-table"><thead><tr><th>City</th><th>Mean Salary</th></tr></thead><tbody>{rows}</tbody></table></div>{foot()}"""


def city_page_html(cn, cslu, jobs_data):
    rows = "".join(f'<tr><td><a href="/salary/{r["job_slug"]}/{cslu}.html">{r["OCC_TITLE"]}</a></td><td>{fmt(r["A_MEAN"])}</td></tr>' for _,r in jobs_data.iterrows())
    h = head(f"Salaries in {cn} | SignalBleed", f"Compare salaries for {len(jobs_data)} occupations in {cn}.", f"{SITE}/cities/{cslu}.html", INDEX_CSS_EXTRA)
    return f"""{h}<div class="page"><h1>Salaries in {cn}</h1><p style="font-size:14px;color:var(--text-muted);margin-bottom:24px">{len(jobs_data)} occupations &middot; BLS OEWS May 2024</p><table class="compare-table"><thead><tr><th>Occupation</th><th>Mean Salary</th></tr></thead><tbody>{rows}</tbody></table></div>{foot()}"""


# ---------------------------------------------------------------------------
# STEP 5: GENERATE EVERYTHING
# ---------------------------------------------------------------------------
def generate(df, natl_df, imap):
    natl = dict(zip(natl_df["OCC_CODE"], natl_df["natl_mean"]))
    total = len(df)

    # --- Salary pages ---
    print(f"  Generating {total:,} salary pages...")
    for i,(_,row) in enumerate(df.iterrows()):
        d = OUTPUT / "salary" / row["job_slug"]
        d.mkdir(parents=True, exist_ok=True)
        html = salary_page(row, df, natl, imap)
        (d / f'{row["city_slug"]}.html').write_text(html)
        if (i+1) % 5000 == 0:
            print(f"    {i+1:,}/{total:,}")
    print(f"  Done: {total:,} salary pages")

    # --- Job index ---
    print("  Generating job pages...")
    js = df.groupby(["OCC_CODE","OCC_TITLE","job_slug"]).agg(
        natl_mean=("A_MEAN","mean"), city_count=("city_slug","nunique")
    ).reset_index().sort_values("OCC_TITLE")
    # Use actual national avgs
    js["natl_mean"] = js["OCC_CODE"].map(natl)

    (OUTPUT/"jobs").mkdir(parents=True, exist_ok=True)
    (OUTPUT/"jobs"/"index.html").write_text(job_index_page(js))

    for _,jr in js.iterrows():
        cd = df[df["OCC_CODE"]==jr["OCC_CODE"]].sort_values("A_MEAN",ascending=False)
        (OUTPUT/"jobs"/f'{jr["job_slug"]}.html').write_text(
            job_page(jr["OCC_TITLE"], jr["job_slug"], fmt(jr["natl_mean"]), cd)
        )
    print(f"  Done: {len(js)} job pages + index")

    # --- City index ---
    print("  Generating city pages...")
    cs = df.groupby(["city_slug","city_name"]).agg(job_count=("OCC_CODE","nunique")).reset_index().sort_values("city_name")

    (OUTPUT/"cities").mkdir(parents=True, exist_ok=True)
    (OUTPUT/"cities"/"index.html").write_text(city_index_page(cs))

    for _,cr in cs.iterrows():
        jd = df[df["city_slug"]==cr["city_slug"]].sort_values("A_MEAN",ascending=False)
        (OUTPUT/"cities"/f'{cr["city_slug"]}.html').write_text(
            city_page_html(cr["city_name"], cr["city_slug"], jd)
        )
    print(f"  Done: {len(cs)} city pages + index")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    print("="*60)
    print("SIGNALBLEED SITE GENERATOR")
    print("="*60)

    print("\n[1/5] Downloading BLS data...")
    zp = download()

    print("\n[2/5] Loading + cleaning...")
    df = load(zp)

    print("\n[3/5] National averages...")
    na = national_avgs(df)

    print("\n[4/5] City images from Wikipedia...")
    imap = fetch_images(df)

    OUTPUT.mkdir(parents=True, exist_ok=True)

    print("\n[5/5] Generating pages...")
    generate(df, na, imap)

    # Copy homepage if it exists alongside this script
    hp = Path("index.html")
    if hp.exists():
        import shutil
        shutil.copy(hp, OUTPUT/"index.html")
        print("\n  Homepage copied")
    else:
        print("\n  NOTE: Put your index.html in the same folder as this script")
        print("  and re-run, or copy it into /output manually.")

    print("\n" + "="*60)
    print("DONE!")
    print(f"All files are in: {OUTPUT.resolve()}")
    print()
    print("Next steps:")
    print("  1. Copy everything from /output into your repo root")
    print("  2. git add -A && git commit -m 'Add salary pages' && git push")
    print("="*60)


if __name__ == "__main__":
    main()
