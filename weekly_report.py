"""
WKT Weekly Report (condensed)
=============================
Three things only:
  1. GP:CAC + throughput per channel (Facebook vs Google)
  2. Tech utilisation  (last week, and next 10 working days)
  3. Retention         (live / settled 3-month / lifetime)

Run:
    python weekly_report.py               # asks for CPL, margin, techs -> saves PDF to reports/
    python weekly_report.py --no-prompt   # reuse last week's answers

Data sources: ServiceM8 (jobs, bookings, customer contacts) + the leads inbox via Microsoft Graph.
Manual inputs (asked each run, remembered): Facebook CPL, Google CPL, gross margin, tech list.

Definitions
-----------
Lead        one email in the leads inbox. Unbounce/gclid = Google, "submitted your form" = Organic,
            "Facebook" = Facebook. Same phone/email within 30 days = one lead; across channels the
            first channel to see the person gets the credit. Organic is used only for this dedupe.
Booked      the person (matched by phone or email to a ServiceM8 customer) has a job on the
            dispatch calendar (any status) or a Completed job, with a booking date within 30 days
            of the lead (1 day grace before).
Throughput  booked / leads, as a decimal.
CAC         CPL / throughput.
GP (FB)     average value of the first earned sale (scheduled Work Order or Completed job) x margin.
GP (Google) average 6-month sum of earned sales from the lead x gross margin.
GP:CAC      GP / CAC.
Utilisation booked hours / (techs x 8h x working days). Booked hours = each tech's scheduled job time,
            plus gaps between jobs of <= 1h (drive time). Can exceed 100%.
Retention   due date = completion + 12 months for every service. Retained = any job on the calendar
            from due - 30 days, or the same service rebooked 6+ months after the last one (early return). Live = cohorts due < 3 months ago (window open). Settled = returns
            within 3 months of due, frozen. Lifetime = ever returned (settled cohorts).
"""

import os
import re
import json
import html
import argparse
import calendar
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone

import requests

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass

HERE        = os.path.dirname(os.path.abspath(__file__))
INPUTS_FILE = os.path.join(HERE, "report_inputs.json")
LEADS_CACHE = os.path.join(HERE, "leads_cache.json")
CACHE_VERSION = 3

# ── Settings ──────────────────────────────────────────────────────────────────
DEFAULT_TECHS           = ["Max Horvath", "Sam Dickinson", "Greg Edwards",
                           "Hayden Johnston", "Ethan McDougall"]
DEFAULT_GROSS_MARGIN    = 0.50
HOURS_PER_DAY           = 8.0      # 40h / tech / week
HORIZON_DAYS            = 10       # prospective: next 10 working days
DRIVE_GAP_MAX_H         = 1.0
BOOK_WINDOW_DAYS        = 30       # lead -> booking window
BOOK_GRACE_BEFORE_DAYS  = 1        # job booked the day before the email still counts
DEDUPE_DAYS             = 30
GOOGLE_CATCHMENT_MONTHS = 6
LEAD_LOOKBACK_DAYS      = 240      # first run: how far back to read the inbox
WEEKS_SHOWN             = 6
MONTHS_SHOWN            = 6
POOLED_WEEKS            = 4
RETENTION_START         = (2024, 11)
RETENTION_GRACE_BEFORE  = 30
RETENTION_SETTLE_MONTHS = 3
DUE_MONTHS              = 12

SM8_BASE   = "https://api.servicem8.com/api_1.0"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ── Small helpers ─────────────────────────────────────────────────────────────
def add_months(d, n):
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def parse_date(s):
    try:
        return date.fromisoformat((s or "")[:10])
    except ValueError:
        return None


def parse_dt(s):
    s = (s or "")[:19]
    if not s or s.startswith("0000"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s if fmt != "%Y-%m-%d" else s[:10], fmt)
        except ValueError:
            continue
    return None


def is_active(rec):
    return str(rec.get("active", "1")) != "0"


def money(v):
    return f"${v:,.0f}"


# ── Manual inputs (asked every run, remembered between runs) ──────────────────
def load_inputs():
    if os.path.exists(INPUTS_FILE):
        with open(INPUTS_FILE) as f:
            return json.load(f)
    return {"fb_cpl": None, "google_cpl": None, "gross_margin": DEFAULT_GROSS_MARGIN,
            "techs": list(DEFAULT_TECHS), "cpl_history": {}}


def save_inputs(inp):
    with open(INPUTS_FILE, "w") as f:
        json.dump(inp, f, indent=2)


def _ask_number(label, current, is_margin=False):
    shown = (f"{current:.2f}" if is_margin else money(current)) if current is not None else "none"
    while True:
        raw = input(f"{label} [Enter = keep {shown}]: ").strip().replace("$", "").replace(",", "")
        if not raw:
            if current is None:
                print("  Need a number the first time.")
                continue
            return current
        if raw.endswith("%"):
            raw = str(float(raw[:-1]) / 100)
        try:
            v = float(raw)
        except ValueError:
            print("  Not a number - try again.")
            continue
        if is_margin and v > 1:            # typed 50 meaning 50%
            v /= 100
        return v


def _split_names(raw):
    return [n.strip() for n in raw.split(",") if n.strip()]


def week_start(today):
    """Weeks run Saturday to Friday. Start of the week containing `today`."""
    return today - timedelta(days=(today.weekday() - 5) % 7)


def last_complete_month(today):
    end = today.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def prompt_inputs(inp, staff_names, today):
    last_mon = week_start(today) - timedelta(days=7)          # Saturday that started last week
    print(f"\n--- Inputs ---")
    print(f"Last week ({last_mon:%d %b} - {last_mon + timedelta(days=6):%d %b}):")
    inp["fb_cpl"]       = _ask_number("  Facebook CPL", inp.get("fb_cpl"))
    inp["google_cpl"]   = _ask_number("  Google CPL", inp.get("google_cpl"))
    inp.setdefault("cpl_history", {})[last_mon.isoformat()] = {"fb_cpl": inp["fb_cpl"],
                                                               "google_cpl": inp["google_cpl"]}
    ms, _ = last_complete_month(today)
    mk = ms.strftime("%Y-%m")
    monthly = inp.setdefault("monthly_cpl", {})
    prev = monthly.get(mk) or {}
    print(f"Last month ({ms:%B %Y}):")
    monthly[mk] = {"fb_cpl": _ask_number("  Facebook CPL", prev.get("fb_cpl", inp["fb_cpl"])),
                   "google_cpl": _ask_number("  Google CPL", prev.get("google_cpl", inp["google_cpl"]))}
    inp["gross_margin"] = _ask_number("Gross margin (e.g. 0.5)", inp.get("gross_margin"), is_margin=True)

    techs = inp.get("techs") or list(DEFAULT_TECHS)
    print("\nTechs in the utilisation calculation: " + ", ".join(techs))
    ans = input("Would you like to remove or add any techs to this calculation? (y/n): ").strip().lower()
    if ans.startswith("y"):
        lookup = {n.lower(): n for n in staff_names}
        for name in _split_names(input("  Names to ADD (comma-separated, Enter for none): ")):
            real = lookup.get(name.lower())
            if not real:
                close = [n for n in staff_names if name.lower() in n.lower()]
                real = close[0] if len(close) == 1 else None
            if real and real not in techs:
                techs.append(real)
                print(f"  + added {real}")
            elif not real:
                print(f"  ! '{name}' not found in ServiceM8 staff - skipped")
        for name in _split_names(input("  Names to REMOVE (comma-separated, Enter for none): ")):
            match = [t for t in techs if t.lower() == name.lower()] or \
                    [t for t in techs if name.lower() in t.lower()]
            if len(match) == 1:
                techs.remove(match[0])
                print(f"  - removed {match[0]}")
            else:
                print(f"  ! '{name}' not in the current list - skipped")
    inp["techs"] = techs
    return inp


# ── ServiceM8 ─────────────────────────────────────────────────────────────────
def sm8_session():
    s = requests.Session()
    key = os.getenv("SM8_API_KEY")
    if key:
        s.headers["X-API-Key"] = key
    else:
        s.auth = (os.environ["SM8_USERNAME"], os.environ["SM8_PASSWORD"])
    return s


def sm8_get(sess, endpoint):
    r = sess.get(f"{SM8_BASE}/{endpoint}", timeout=180)
    r.raise_for_status()
    return r.json()


def fetch_sm8():
    print("Fetching ServiceM8...")
    s = sm8_session()
    data = {name: sm8_get(s, f"{name}.json")
            for name in ("job", "jobactivity", "companycontact", "jobcontact", "staff", "jobmaterial")}
    for k, v in data.items():
        print(f"  {k}: {len(v)} records")
    return data


def staff_name(rec):
    first = rec.get("first") or rec.get("first_name") or ""
    last  = rec.get("last") or rec.get("last_name") or ""
    return f"{first} {last}".strip() or (rec.get("name") or "").strip()


def resolve_techs(names, staff):
    by_name = {}
    for s in staff:
        if is_active(s) and staff_name(s):
            by_name[staff_name(s).lower()] = s.get("uuid")
    out, missing = {}, []
    for n in names:
        u = by_name.get(n.lower())
        if u:
            out[u] = n
        else:
            missing.append(n)
    if missing:
        print(f"  WARNING: not found in ServiceM8 staff (left out of utilisation): {', '.join(missing)}")
    return out


# ── Lead emails (Microsoft Graph) ─────────────────────────────────────────────
PHONE_RE = re.compile(r"(?<![\w+])(?:\+?61[\s-]?|0)([2-478](?:[\s-]?\d){8})(?!\w)")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
OWN_EMAIL_HINTS = ("wekilltermite", "unbounce", "facebook", "servicem8", "google", "microsoft",
                   "noreply", "no-reply", "mailer", "zapier")


def classify_lead(subject, body):
    """Only real lead notifications. A bare mention of 'facebook' (e.g. an email signature or a
    Meta billing email) is NOT a lead."""
    t = f"{subject}\n{body}".lower()
    if "submitted your form" in t:
        return "Organic"
    if "unbounce" in t and ("gclid" in t or "new lead has been captured" in t):
        return "Google"
    if "facebook lead" in t:
        return "Facebook"
    return None


def norm_phone(digits):
    return "0" + re.sub(r"\D", "", digits)


def extract_contacts(text):
    phones = {norm_phone(m.group(1)) for m in PHONE_RE.finditer(text or "")}
    emails = {e.lower().rstrip(".") for e in EMAIL_RE.findall(text or "")
              if not any(h in e.lower() for h in OWN_EMAIL_HINTS)}
    return phones, emails


def parse_lead_email(subject, body, received_iso):
    ch = classify_lead(subject, body)
    if not ch:
        return None
    phones, emails = extract_contacts(body)
    if not phones and not emails:
        return None                                    # can't be a usable lead without contact details
    rec = datetime.fromisoformat(received_iso.replace("Z", "+00:00"))
    local = rec.astimezone(timezone(timedelta(hours=10)))          # Brisbane, no DST
    return {"channel": ch, "received": local.date().isoformat(),
            "phones": sorted(phones), "emails": sorted(emails), "subject": subject[:80]}


def graph_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}/oauth2/v2.0/token",
        data={"client_id": os.environ["MS_CLIENT_ID"], "client_secret": os.environ["MS_CLIENT_SECRET"],
              "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"},
        timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def fetch_leads(today):
    """Read new emails since the last run, keep only lead emails, cache them locally."""
    cache = {"version": CACHE_VERSION, "last_fetch": None, "leads": {}}
    if os.path.exists(LEADS_CACHE):
        with open(LEADS_CACHE) as f:
            old = json.load(f)
        if old.get("version") == CACHE_VERSION:
            cache = old
        else:
            print("Lead rules changed - re-reading the whole inbox once")
    if cache.get("last_fetch"):
        since = datetime.fromisoformat(cache["last_fetch"]) - timedelta(days=1)
    else:
        since = datetime.combine(today - timedelta(days=LEAD_LOOKBACK_DAYS), datetime.min.time())
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    mailbox = os.getenv("LEADS_MAILBOX", "admin@wekilltermites.com.au")
    print(f"Reading lead emails in {mailbox} since {since_iso[:10]}...")
    headers = {"Authorization": f"Bearer {graph_token()}",
               "Prefer": 'outlook.body-content-type="text"'}
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {"$filter": f"receivedDateTime ge {since_iso}",
              "$select": "id,subject,receivedDateTime,body,from",
              "$orderby": "receivedDateTime asc", "$top": "100"}
    scanned = new = 0
    while url:
        r = requests.get(url, headers=headers, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        for m in data.get("value", []):
            scanned += 1
            if re.match(r"\s*(re|fw|fwd)\s*:", m.get("subject") or "", re.I):
                continue                                # replies / forwards, not new leads
            lead = parse_lead_email(m.get("subject") or "", (m.get("body") or {}).get("content") or "",
                                    m["receivedDateTime"])
            if lead and m["id"] not in cache["leads"]:
                cache["leads"][m["id"]] = lead
                new += 1
        url, params = data.get("@odata.nextLink"), None
    cache["last_fetch"] = datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds")
    with open(LEADS_CACHE, "w") as f:
        json.dump(cache, f)
    print(f"  scanned {scanned} emails, {new} new lead emails, {len(cache['leads'])} cached in total")
    return list(cache["leads"].values())


# ── Leads -> episodes -> bookings ─────────────────────────────────────────────
def build_episodes(leads):
    """Merge repeat enquiries: same phone/email within DEDUPE_DAYS of an episode's first lead joins
    that episode. The episode keeps its first channel (first-touch credit)."""
    eps, by_key = [], {}
    for L in sorted(leads, key=lambda x: x["received"]):
        d = date.fromisoformat(L["received"])
        keys = {f"p:{p}" for p in L["phones"]} | {f"e:{e}" for e in L["emails"]}
        ep = None
        for k in keys:
            cand = by_key.get(k)
            if cand and 0 <= (d - cand["start"]).days <= DEDUPE_DAYS:
                ep = cand
                break
        if ep is None:
            ep = {"start": d, "channel": L["channel"], "keys": set(), "n_emails": 0}
            eps.append(ep)
        ep["keys"] |= keys
        ep["n_emails"] += 1
        for k in keys:
            by_key[k] = ep
    return eps


def contact_index(sm8):
    """key ('p:04..' / 'e:x@y') -> set of company_uuids, from client contacts and job contacts."""
    idx = defaultdict(set)
    job_company = {j.get("uuid"): j.get("company_uuid") for j in sm8["job"]}

    def add(text, cu):
        if not cu:
            return
        phones, emails = extract_contacts(text)
        for p in phones:
            idx[f"p:{p}"].add(cu)
        for e in emails:
            idx[f"e:{e}"].add(cu)

    for c in sm8["companycontact"]:
        if is_active(c):
            add(" ".join(str(c.get(k) or "") for k in ("phone", "mobile", "email")), c.get("company_uuid"))
    for c in sm8["jobcontact"]:
        if is_active(c):
            add(" ".join(str(c.get(k) or "") for k in ("phone", "mobile", "email")),
                job_company.get(c.get("job_uuid")))
    return idx


def booked_jobs_by_company(jobs, activities):
    """A job counts as booked when it is active and either has a visit on the dispatch calendar
    (any status, so a booked inspection still sitting as a Quote counts) or is Completed."""
    on_cal = {a.get("job_uuid") for a in activities if is_active(a) and a.get("job_uuid")}
    out = defaultdict(list)
    for j in jobs:
        if not is_active(j):
            continue
        if j.get("uuid") not in on_cal and (j.get("status") or "").lower() != "completed":
            continue
        d = parse_date(j.get("date"))
        if d and j.get("company_uuid"):
            st = (j.get("status") or "").lower()
            earned = st == "completed" or (st == "work order" and j.get("uuid") in on_cal)
            out[j["company_uuid"]].append((d, float(j.get("total_invoice_amount") or 0), j.get("uuid"), earned))
    for cu in out:
        out[cu].sort()
    return out


def attach_bookings(eps, idx, jobs_by_co):
    for ep in eps:
        cos = set()
        for k in ep["keys"]:
            cos |= idx.get(k, set())
        lo = ep["start"] - timedelta(days=BOOK_GRACE_BEFORE_DAYS)
        hi30 = ep["start"] + timedelta(days=BOOK_WINDOW_DAYS)
        hi6m = add_months(ep["start"], GOOGLE_CATCHMENT_MONTHS)
        jobs = sorted({j for cu in cos for j in jobs_by_co.get(cu, []) if lo <= j[0] <= hi6m})
        in30 = [j for j in jobs if j[0] <= hi30]
        ep["matched"]     = bool(cos)
        ep["booked"]      = bool(in30)
        # Earned sale = scheduled AND Work Order/Completed (a scheduled quote is booked, not sold)
        paid = [j for j in in30 if j[3] and j[1] > 0]
        ep["first_value"] = paid[0][1] if paid else None
        ep["first_job"]   = paid[0][2] if paid else None
        ep["companies"]   = cos
        earned = [j for j in jobs if j[3]]
        ep["six_month"]   = sum(j[1] for j in earned) if earned else None
    return eps


def channel_stats(eps, channel, frm, to, cpl, gm):
    pool   = [e for e in eps if e["channel"] == channel and frm <= e["start"] <= to]
    leads  = len(pool)
    booked = sum(1 for e in pool if e["booked"])
    thru   = booked / leads if leads else 0.0
    if channel == "Facebook":
        vals = [e["first_value"] for e in pool if e["booked"] and e["first_value"] is not None]
    else:
        vals = [e["six_month"] for e in pool if e["six_month"] is not None]
    asp  = sum(vals) / len(vals) if vals else 0.0
    gp   = asp * gm
    cac  = (cpl / thru) if (cpl and thru) else None
    ratio = (gp / cac) if cac else None
    return {"leads": leads, "booked": booked, "thru": thru, "asp": asp, "n_asp": len(vals),
            "gp": gp, "cpl": cpl, "cac": cac, "ratio": ratio,
            "unmatched": sum(1 for e in pool if not e["matched"])}


# ── Tech utilisation ──────────────────────────────────────────────────────────
def working_days(start, n=None, end=None):
    days, d = [], start
    while (n is not None and len(days) < n) or (end is not None and d <= end):
        if d.weekday() < 5:
            days.append(d)
        d += timedelta(days=1)
    return days


def utilisation(activities, days, techs):
    dayset = set(days)
    by_key = defaultdict(list)
    for a in activities:
        if not is_active(a) or a.get("staff_uuid") not in techs:
            continue
        s, e = parse_dt(a.get("start_date")), parse_dt(a.get("end_date"))
        if not s or not e or e <= s or s.date() not in dayset:
            continue
        by_key[(a["staff_uuid"], s.date())].append((s, e))
    per_tech = defaultdict(float)
    gap_max = timedelta(hours=DRIVE_GAP_MAX_H)
    for (tech, _), ivals in by_key.items():
        ivals.sort()
        merged = []
        for s, e in ivals:
            if merged and s <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))
            else:
                merged.append((s, e))
        secs = sum((e - s).total_seconds() for s, e in merged)
        for i in range(1, len(merged)):
            gap = merged[i][0] - merged[i - 1][1]
            if timedelta(0) < gap <= gap_max:
                secs += gap.total_seconds()
        per_tech[tech] += secs / 3600
    avail_each = HOURS_PER_DAY * len(days)
    rows = [(name, per_tech.get(u, 0.0), avail_each) for u, name in techs.items()]
    booked = sum(r[1] for r in rows)
    avail = avail_each * len(techs)
    return {"pct": booked / avail * 100 if avail else 0.0, "booked": booked, "avail": avail,
            "rows": rows, "from": days[0], "to": days[-1]}


# ── Retention ─────────────────────────────────────────────────────────────────
_SVC_OTHER = ("surcharge", "cancellation fee", "key collection", "service location fee", "weekend callout",
              "card 2%", "gst", "tax 10%", "discount", "credit", "referral", "tip")
_SVC_TREAT = ("termite barrier", "termite management system", "termite protection", "termite bait",
              "reticulation", "trelona", "atbs", "sentricon", "live termite", "termite part treatment",
              "termite treatment", "termite foaming", "termidor", "chemical treated zone", "chemical barrier",
              "termite monitoring", "dry wood termite")
_SVC_INSPECT = ("termite inspection", "termite check", "timber pest inspection", "timber pest",
                "visual pest inspection")
_SVC_PEST = ("pest control", "general pest", "premium pest", "rodent", "cockroach", "german", "flea",
             "spider", "ant treatment", "wasp", "silverfish", "pest treatment")


def job_services(job, items):
    """Service categories on a job: PC (pest control), TI (termite inspection), TT (termite
    treatment). A job with none of them is 'OT' (other)."""
    cats = set()
    for text in [job.get("job_description") or ""] + list(items):
        n = re.sub(r"\s+", " ", text.lower())
        if not n or any(k in n for k in _SVC_OTHER):
            continue
        if any(k in n for k in _SVC_TREAT):
            cats.add("TT")
        if any(k in n for k in _SVC_INSPECT):
            cats.add("TI")
        if any(k in n for k in _SVC_PEST):
            cats.add("PC")
    return cats or {"OT"}


def retention_series(jobs, activities, materials, today):
    """Cohort = month a customer's service fell due (completion + 12 months, every service).
    Retained if, for that due date D from a service on date d:
      - ANY job (any service) is put on the calendar between D - 30 days and D + 3 months
        (settled; lifetime has no upper bound), or
      - an early return: a job for the SAME service put on the calendar between d + 6 months and
        D - 30 days.
    A same-service job completed within 6 months of the previous one is a follow-up visit, not a
    new cycle, so it doesn't create its own due date."""
    items = defaultdict(list)
    for m in materials:
        if m.get("job_uuid") and is_active(m):
            items[m["job_uuid"]].append(m.get("name") or "")
    on_cal = {a.get("job_uuid") for a in activities if is_active(a) and a.get("job_uuid")}
    sched = defaultdict(list)                       # customer -> [(booking date, services)]
    done = defaultdict(list)                        # customer -> [(completion date, services)]
    for j in jobs:
        cu = j.get("company_uuid")
        if not cu or not is_active(j):
            continue
        cats = None
        if j.get("uuid") in on_cal:
            d = parse_date(j.get("date"))
            if d:
                cats = job_services(j, items.get(j.get("uuid"), []))
                sched[cu].append((d, cats))
        if j.get("status") == "Completed":
            text = f"{j.get('job_description') or ''} {j.get('work_done_description') or ''}".lower()
            cd = parse_date(j.get("completion_date"))
            if cd and "fixup" not in text:
                done[cu].append((cd, cats or job_services(j, items.get(j.get("uuid"), []))))

    by_month = defaultdict(lambda: defaultdict(list))   # (yr, mo) -> customer -> [(d, D, cats)]
    for cu, svcs in done.items():
        svcs.sort(key=lambda x: x[0])
        for i, (d, cats) in enumerate(svcs):
            if any(cats & c2 and d < add_months(d2, 6) for d2, c2 in svcs[:i] if d2 < d):
                continue                                # follow-up visit within 6 months
            D = add_months(d, DUE_MONTHS)
            by_month[(D.year, D.month)][cu].append((d, D, cats))

    def returned(cu, entries, upper):
        for d, D, cats in entries:
            lo = D - timedelta(days=RETENTION_GRACE_BEFORE)
            for sd, sc in sched.get(cu, []):
                if lo <= sd and (upper is None or sd <= add_months(D, upper)):
                    return True
                if add_months(d, 6) <= sd < lo and sc & cats:
                    return True
        return False

    first_this = today.replace(day=1)
    last_start = add_months(first_this, -1)
    settle_cut = add_months(first_this, -RETENTION_SETTLE_MONTHS)
    series = {"live": [], "settled": [], "lifetime": []}
    last = {}
    m = date(*RETENTION_START, 1)
    while m <= last_start:
        custs = by_month.get((m.year, m.month), {})
        n = len(custs)
        in_win = sum(returned(cu, e, RETENTION_SETTLE_MONTHS) for cu, e in custs.items())
        ever = sum(returned(cu, e, None) for cu, e in custs.items())
        lab = m.strftime("%b %Y")
        pct = lambda k: (k / n * 100) if n else 0.0
        if m <= settle_cut:
            series["settled"].append((lab, pct(in_win)))
            series["lifetime"].append((lab, pct(ever)))
            last = {"lab": lab, "n": n, "in_win": in_win, "ever": ever}
        else:
            series["live"].append((lab, pct(in_win)))
        m = add_months(m, 1)
    return series, last


# ── Rendering ─────────────────────────────────────────────────────────────────
CSS = """
@page { size: A4; margin: 14mm 12mm; }
body { font-family: Arial, sans-serif; font-size: 11px; color: #333; }
h1 { font-size: 19px; margin: 0 0 2px; color: #1a1a2e; }
.sub { color: #888; font-size: 10.5px; margin-bottom: 14px; }
h2 { font-size: 13px; color: #fff; background: #1a1a2e; padding: 5px 9px; border-radius: 3px; margin: 18px 0 8px; }
h3 { font-size: 11.5px; margin: 12px 0 5px; color: #1a1a2e; }
table { border-collapse: collapse; width: 100%; }
th { font-size: 9.5px; color: #888; font-weight: 400; text-align: right; padding: 4px 6px; background: #f6f6f6; }
th:first-child, td:first-child { text-align: left; }
td { padding: 4px 6px; text-align: right; border-top: 1px solid #eee; font-variant-numeric: tabular-nums; }
tr.pool td { font-weight: 700; background: #eef3fb; }
.big { font-size: 22px; font-weight: 700; color: #1a1a2e; }
.kpis td { border: none; text-align: left; padding: 2px 14px 8px 0; vertical-align: top; }
.note { font-size: 9px; color: #999; margin-top: 4px; line-height: 1.45; }
.grid td { font-size: 9px; padding: 2px 5px; }
"""
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def fmt_ratio(v):
    return f"{v:.1f} : 1" if v is not None else "n/a"


def channel_table(name, rows, period="Week"):
    h = (f"<table><tr><th>{period}</th><th>Leads</th><th>Booked</th><th>Throughput</th><th>CPL</th>"
         "<th>CAC</th><th>Avg sale</th><th>GP</th><th>GP:CAC</th></tr>")
    for label, s, pool in rows:
        h += (f"<tr class='{'pool' if pool else ''}'><td>{label}</td><td>{s['leads']}</td><td>{s['booked']}</td>"
              f"<td>{s['thru']:.2f}</td><td>{money(s['cpl']) if s['cpl'] else 'n/a'}</td>"
              f"<td>{money(s['cac']) if s['cac'] else 'n/a'}</td><td>{money(s['asp'])}</td>"
              f"<td>{money(s['gp'])}</td><td>{fmt_ratio(s['ratio'])}</td></tr>")
    return f"<h3>{name}</h3>" + h + "</table>"


def retention_grid(months):
    if not months:
        return "<p class='note'>No data yet.</p>"
    data, years = {}, []
    for lab, v in months:
        mon, yr = lab.split()
        if yr not in years:
            years.append(yr)
        data[(yr, mon)] = v
    h = "<table class='grid'><tr><th></th>" + "".join(f"<th>{m}</th>" for m in MONTHS) + "</tr>"
    for yr in years:
        h += f"<tr><td><b>{yr}</b></td>" + "".join(
            f"<td>{data[(yr, m)]:.1f}</td>" if (yr, m) in data else "<td style='color:#ddd'>&middot;</td>"
            for m in MONTHS) + "</tr>"
    return h + "</table>"


def util_table(u):
    h = "<table><tr><th>Tech</th><th>Booked h</th><th>Available h</th><th>Utilisation</th></tr>"
    for name, b, a in u["rows"]:
        h += f"<tr><td>{html.escape(name)}</td><td>{b:.1f}</td><td>{a:.0f}</td><td>{(b / a * 100 if a else 0):.0f}%</td></tr>"
    h += (f"<tr class='pool'><td>All techs</td><td>{u['booked']:.1f}</td><td>{u['avail']:.0f}</td>"
          f"<td>{u['pct']:.0f}%</td></tr>")
    return h + "</table>"


def build_html(today, ch_rows, util_retro, util_pro, ret, ret_last, gm, notes):
    def kpi(label, val, sub=""):
        return f"<td><div class='note'>{label}</div><div class='big'>{val}</div><div class='note'>{sub}</div></td>"

    fb_pool, g_pool = ch_rows["Facebook"][-1][1], ch_rows["Google"][-1][1]
    settled_latest = ret["settled"][-1][1] if ret["settled"] else 0
    top = ("<table class='kpis'><tr>"
           + kpi("GP:CAC Facebook (last 4 wks)", fmt_ratio(fb_pool["ratio"]), f"throughput {fb_pool['thru']:.2f}")
           + kpi("GP:CAC Google (last 4 wks)", fmt_ratio(g_pool["ratio"]), f"throughput {g_pool['thru']:.2f}")
           + kpi("Utilisation next 2 wks", f"{util_pro['pct']:.0f}%",
                 f"{util_pro['from']:%d %b}&ndash;{util_pro['to']:%d %b}")
           + kpi("Utilisation last week", f"{util_retro['pct']:.0f}%",
                 f"{util_retro['from']:%d %b}&ndash;{util_retro['to']:%d %b}")
           + kpi(f"Retention settled ({ret_last.get('lab', '')})", f"{settled_latest:.1f}%",
                 f"{ret_last.get('in_win', 0)} of {ret_last.get('n', 0)} due")
           + "</tr></table>")

    body = f"""
<h1>WKT Weekly Report</h1>
<div class='sub'>Generated {today:%d %B %Y} &middot; gross margin {gm:.0%}</div>
{top}
<h2>GP:CAC by channel</h2>
{channel_table('Facebook - by week', ch_rows['Facebook'])}
{channel_table('Facebook - by month', ch_rows['Facebook_m'], 'Month')}
<p class='note'>GP = average value of the first earned sale (scheduled Work Order or Completed job, not a
quote) &times; gross margin. CAC = CPL &divide; throughput.</p>
{channel_table('Google - by week', ch_rows['Google'])}
{channel_table('Google - by month', ch_rows['Google_m'], 'Month')}
<p class='note'>GP = average 6-month sum of earned sales (scheduled Work Order or Completed jobs) from the lead
&times; gross margin. Recent weeks
are understated until their 6-month window fills in, so compare older weeks with care.</p>
<p class='note'>Leads come from the leads inbox (repeat enquiries within 30 days merged, first channel gets
credit). Booked = matched by phone/email to a ServiceM8 customer with a job put on the calendar (or completed)
within 30 days. The latest week's throughput can still rise as late bookings land. {notes}</p>

<h2>Tech utilisation</h2>
<h3>Last week ({util_retro['from']:%d %b}&ndash;{util_retro['to']:%d %b})</h3>
{util_table(util_retro)}
<h3>Next 10 working days ({util_pro['from']:%d %b}&ndash;{util_pro['to']:%d %b})</h3>
{util_table(util_pro)}
<p class='note'>Booked hours = scheduled job time plus gaps of 1 hour or less between jobs (drive time).
Available = 8h per tech per working day. Public holidays and leave are not removed.</p>

<h2>Retention</h2>
<h3>Live (window still open, will rise)</h3>
{retention_grid(ret['live'])}
<h3>Settled (returned within {RETENTION_SETTLE_MONTHS} months of due, frozen)</h3>
{retention_grid(ret['settled'])}
<h3>Lifetime (ever returned, same cohorts)</h3>
{retention_grid(ret['lifetime'])}
<p class='note'>By month the service fell due. Due = completion + 12 months for every customer. Retained = any new job
put on the calendar (no need to wait for completion) from 30 days before due, or an early return: the same
service rebooked 6+ months after the last one. Follow-up visits within 6 months don't start a new cycle.</p>
"""
    return f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>"


# ── Save to OneDrive ─────────────────────────────────────────────────────────
def save_to_onedrive(pdf_path):
    """Upload the PDF to the mailbox owner's OneDrive, folder 'WKT Reports'. Needs the app
    permission Files.ReadWrite.All (Application) with admin consent."""
    user = os.getenv("ONEDRIVE_USER") or os.getenv("LEADS_MAILBOX", "admin@wekilltermites.com.au")
    folder = os.getenv("ONEDRIVE_FOLDER", "WKT Reports")
    name = os.path.basename(pdf_path)
    try:
        with open(pdf_path, "rb") as f:
            r = requests.put(f"{GRAPH_BASE}/users/{user}/drive/root:/{folder}/{name}:/content",
                             headers={"Authorization": f"Bearer {graph_token()}",
                                      "Content-Type": "application/pdf"}, data=f.read(), timeout=120)
        if r.status_code in (200, 201):
            print(f"Saved to OneDrive ({user}): {folder}/{name}")
            print(f"  Open: {r.json().get('webUrl', '')}")
        else:
            print(f"OneDrive upload failed ({r.status_code}): {r.text[:200]}")
            if r.status_code in (401, 403):
                print("  -> add Files.ReadWrite.All (Application) to the Azure app and grant admin consent")
    except Exception as e:                                       # noqa: BLE001
        print(f"OneDrive upload failed: {e}")


# ── Main ──────────────────────────────────────────────────────────────────────
def build_report(sm8, leads, inp, today):
    gm = inp["gross_margin"]
    eps = attach_bookings(build_episodes(leads), contact_index(sm8), booked_jobs_by_company(sm8["job"], sm8["jobactivity"]))

    this_monday = week_start(today)                           # this week's Saturday
    weeks = [(this_monday - timedelta(days=7 * i), this_monday - timedelta(days=7 * i) + timedelta(days=6))
             for i in range(WEEKS_SHOWN, 0, -1)]
    hist = inp.get("cpl_history", {})
    ch_rows = {}
    for ch, key in (("Facebook", "fb_cpl"), ("Google", "google_cpl")):
        rows = []
        for ws, we in weeks:
            cpl = hist.get(ws.isoformat(), {}).get(key, inp[key])
            rows.append((f"{ws:%d %b}", channel_stats(eps, ch, ws, we, cpl, gm), False))
        pool_from = weeks[-POOLED_WEEKS][0]
        rows.append((f"Last {POOLED_WEEKS} weeks", channel_stats(eps, ch, pool_from, weeks[-1][1], inp[key], gm), True))
        ch_rows[ch] = rows
        mrows = []
        ms, _ = last_complete_month(today)
        for i in range(MONTHS_SHOWN - 1, -1, -1):
            a = add_months(ms, -i)
            b = add_months(a, 1) - timedelta(days=1)
            cpl = inp.get("monthly_cpl", {}).get(a.strftime("%Y-%m"), {}).get(key)
            mrows.append((f"{a:%b %Y}", channel_stats(eps, ch, a, b, cpl, gm), False))
        ch_rows[ch + "_m"] = mrows

    techs = resolve_techs(inp["techs"], sm8["staff"])
    last_mon = this_monday - timedelta(days=7)
    util_retro = utilisation(sm8["jobactivity"], working_days(last_mon, end=last_mon + timedelta(days=6)), techs)
    util_pro = utilisation(sm8["jobactivity"], working_days(today + timedelta(days=1), n=HORIZON_DAYS), techs)
    ret, ret_last = retention_series(sm8["job"], sm8["jobactivity"], sm8.get("jobmaterial", []), today)

    recent = [e for e in eps if e["channel"] in ("Facebook", "Google") and e["start"] >= weeks[0][0]]
    unmatched = sum(1 for e in recent if not e["matched"])
    notes = (f"{unmatched} of {len(recent)} Facebook/Google leads in this window had no phone/email matching a "
             f"ServiceM8 customer (never booked, or details entered differently)." if recent else "")
    return build_html(today, ch_rows, util_retro, util_pro, ret, ret_last, gm, notes), ch_rows, util_retro, util_pro, ret, ret_last


def check_leads(sm8, leads, today):
    """Print a sample of recent Facebook leads that didn't book, to see why."""
    idx = contact_index(sm8)
    jobs = defaultdict(list)
    for j in sm8["job"]:
        if j.get("company_uuid"):
            jobs[j["company_uuid"]].append(j)
    eps = attach_bookings(build_episodes(leads), idx, booked_jobs_by_company(sm8["job"], sm8["jobactivity"]))
    from collections import Counter
    fb60 = [L for L in leads if L["channel"] == "Facebook"
            and date.fromisoformat(L["received"]) >= today - timedelta(days=60)]
    print(f"\nFacebook lead emails, last 60 days: {len(fb60)}. Most common subjects:")
    for subj, n in Counter(L.get("subject", "") for L in fb60).most_common(8):
        print(f"  {n:5}  {subj}")
    ep60 = [e for e in eps if e["channel"] == "Facebook" and e["start"] >= today - timedelta(days=60)]
    m60 = [e for e in ep60 if e["matched"]]
    print(f"After merging repeats: {len(ep60)} leads; {len(m60)} match a ServiceM8 customer; "
          f"{sum(e['booked'] for e in ep60)} booked within 30 days")
    recent = [e for e in eps if e["channel"] == "Facebook" and e["start"] >= today - timedelta(days=60)
              and not e["booked"]]
    unmatched = [e for e in recent if not e["matched"]]
    matched = [e for e in recent if e["matched"]]
    print(f"\nFacebook leads last 60 days, not booked: {len(recent)} "
          f"({len(unmatched)} no ServiceM8 match, {len(matched)} matched but no qualifying job)")
    print("\nSample with NO match (search one of these in ServiceM8 by hand):")
    for e in unmatched[:8]:
        print(f"  {e['start']}  {', '.join(k[2:] for k in sorted(e['keys'])) or '(no phone/email in email)'}")
    print("\nSample MATCHED but not counted as booked (their jobs: date / status):")
    for e in matched[:8]:
        cos = set().union(*(idx.get(k, set()) for k in e["keys"]))
        js = sorted((j.get("date") or "")[:10] + " " + (j.get("status") or "") for cu in cos for j in jobs[cu])
        print(f"  lead {e['start']}: " + ("; ".join(js[-4:]) or "no jobs"))


def check_sales(sm8, leads, today):
    """List every booked Facebook lead in the last 6 weeks with the job used for its sale price."""
    idx = contact_index(sm8)
    eps = attach_bookings(build_episodes(leads), idx, booked_jobs_by_company(sm8["job"], sm8["jobactivity"]))
    job = {j.get("uuid"): j for j in sm8["job"]}
    njobs = defaultdict(int)
    for j in sm8["job"]:
        njobs[j.get("company_uuid")] += 1
    this_monday = week_start(today)
    for i in range(6, 0, -1):
        ws = this_monday - timedelta(days=7 * i)
        pool = [e for e in eps if e["channel"] == "Facebook" and ws <= e["start"] <= ws + timedelta(days=6)
                and e["booked"]]
        vals = [e["first_value"] for e in pool if e["first_value"] is not None]
        print(f"\nWeek {ws:%d %b}: {len(pool)} booked, avg first earned sale "
              f"{money(sum(vals) / len(vals)) if vals else '-'}")
        for e in pool:
            j = job.get(e["first_job"]) or {}
            print(f"  lead {e['start']}  job {(j.get('date') or '')[:10]}  "
                  f"{money(e['first_value']) if e['first_value'] is not None else 'no sale yet':>12}  "
                  f"#{j.get('generated_job_id', '')}  {(j.get('job_description') or '')[:45]!r}  "
                  f"[{len(e['companies'])} client(s), {sum(njobs[c] for c in e['companies'])} jobs on file]")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-prompt", action="store_true", help="reuse last answers")
    ap.add_argument("--check-leads", action="store_true", help="show why Facebook leads aren't matching")
    ap.add_argument("--check-sales", action="store_true", help="list the job behind each Facebook sale price")
    args = ap.parse_args()
    today = date.today()

    sm8 = fetch_sm8()
    if args.check_leads:
        check_leads(sm8, fetch_leads(today), today)
        return
    if args.check_sales:
        check_sales(sm8, fetch_leads(today), today)
        return
    inp = load_inputs()
    if not args.no_prompt:
        inp = prompt_inputs(inp, sorted({staff_name(s) for s in sm8["staff"] if is_active(s) and staff_name(s)}),
                            today)
    save_inputs(inp)

    leads = fetch_leads(today)
    page, ch_rows, ur, up, ret, rl = build_report(sm8, leads, inp, today)

    print("\n=== Summary ===")
    for ch in ("Facebook", "Google"):
        s = ch_rows[ch][-1][1]
        print(f"  {ch:9} last 4 wks: {s['leads']} leads, {s['booked']} booked, throughput {s['thru']:.2f}, "
              f"CAC {money(s['cac']) if s['cac'] else '-'}, avg sale {money(s['asp'])}, GP:CAC {fmt_ratio(s['ratio'])}")
        print(f"            {s['unmatched']} of {s['leads']} leads had no phone/email matching any ServiceM8 customer")
    print(f"  Utilisation last week {ur['pct']:.0f}%, next 10 working days {up['pct']:.0f}%")
    if ret["settled"]:
        print(f"  Retention settled {ret['settled'][-1][0]}: {ret['settled'][-1][1]:.1f}%")

    from weasyprint import HTML
    out = os.path.join(HERE, "reports", f"weekly_report_{today:%Y_%m_%d}.pdf")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    HTML(string=page).write_pdf(out)
    print(f"\nPDF saved: {out}")
    save_to_onedrive(out)


if __name__ == "__main__":
    main()
