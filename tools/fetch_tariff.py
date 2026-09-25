"""Download the official Ethiopian customs tariff and codification lists from the
public Customs Trade Portal (customs.erca.gov.et) into tariff-data.js for the app.

    python3 tools/fetch_tariff.py            # full refresh (about 25 minutes)

The portal only searches by 4+ digit codes, so every heading (0101..9999) is searched
and its CSV export downloaded. Each export is checked against the "N Tariff found"
count on the results page; any mismatch stops the run, so a partial tariff is never
written. Requests are spaced out to be gentle on the public server.
"""
import csv
import datetime
import http.cookiejar
import io
import json
import re
import sys
import time
import urllib.parse
import urllib.request

BASE = "https://customs.erca.gov.et/trade"
DELAY = 0.35          # seconds between requests
MISSES_TO_STOP = 15   # consecutive empty headings before moving to the next chapter

jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
opener.addheaders = [("User-Agent", "Mozilla/5.0 (tariff lookup for eSAD Item Builder)")]


def fetch(url, data=None, tries=4):
    body = urllib.parse.urlencode(data).encode() if data else None
    for i in range(tries):
        try:
            with opener.open(url, body, timeout=120) as r:
                raw = r.read()
            time.sleep(DELAY)
            return raw.decode("utf-8", errors="strict")
        except Exception as e:  # network hiccup: back off and retry
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def search_heading(code):
    page = fetch(f"{BASE}/help/tariff/search", {"keywords": "", "tariff": code, "searchMode": "ANY",
                                                "pageSize": "50", "_action_searchByTariffKeywords": "Search"})
    m = re.search(r"([\d,]+)\s+Tariff found", page)
    if not m:
        return []
    expected = int(m.group(1).replace(",", ""))
    text = fetch(f"{BASE}/help/tariff/export?tariff={code}&format=csv")
    rows = list(csv.DictReader(io.StringIO(text)))
    if len(rows) != expected:
        sys.exit(f"STOP: heading {code}: page says {expected} lines but the CSV has {len(rows)}. Nothing written.")
    for r in rows:
        c = r["Code"].strip()
        if not re.fullmatch(r"\d{11}", c) or not c.startswith(code):
            sys.exit(f"STOP: heading {code}: unexpected code {c!r}. Nothing written.")
    return rows


def lookup(ref, code):
    """One codification entry by exact code (the portal exports at most 10 rows per search)."""
    fetch(f"{BASE}/help/codification/search", {"beanDataLoad": "/trade/beanDataLoad/loadData", "locale": "en", "keywords": code,
                                               "nature": "code", "reference": ref, "pageSize": "50", "search": "Search"})
    text = fetch(f"{BASE}/help/codification/export/csv?reference={ref}&reference.operation=IN")
    return [{"code": r["Code"].strip(), "name": r["Name"].strip()} for r in csv.DictReader(io.StringIO(text)) if r["Code"].strip() == code]


ISO2 = ("AD AE AF AG AI AL AM AO AQ AR AS AT AU AW AX AZ BA BB BD BE BF BG BH BI BJ BL BM BN BO BQ BR BS BT BV BW BY BZ CA CC CD CF CG CH CI CK CL CM CN CO CR CU CV CW CX CY CZ "
        "DE DJ DK DM DO DZ EC EE EG EH ER ES ET FI FJ FK FM FO FR GA GB GD GE GF GG GH GI GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL IM IN IO IQ IR IS IT JE JM JO JP "
        "KE KG KH KI KM KN KP KR KW KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD ME MF MG MH MK ML MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI NL NO NP NR NU NZ "
        "OM PA PE PF PG PH PK PL PM PN PR PS PT PW PY QA RE RO RS RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO SR SS ST SV SX SY SZ TC TD TF TG TH TJ TK TL TM TN TO TR TT TV TW TZ "
        "UA UG UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT ZA ZM ZW").split()
PACKAGES = "PK CT BX PL PX CS BG SA DR BA RO BE CR PC NE CN TN CA BL BN RL CO JR BK TU VA VL VR VY VQ VO ZZ".split()


def main():
    fetch(f"{BASE}/help/tariff?lang=en")  # establish the anonymous session
    lines = []
    for ch in range(1, 100):
        if ch == 77:  # reserved chapter
            continue
        misses = found_any = 0
        for h in range(1, 100):
            code = f"{ch:02d}{h:02d}"
            rows = search_heading(code)
            if rows:
                found_any, misses = 1, 0
                for r in rows:
                    lines.append({"c": r["Code"][:8], "p": r["Code"][8:], "d": r["Description"].strip(),
                                  "u": (r.get("Unit(s)") or "").strip(),
                                  "r": {k: (r.get(k) or "").strip() for k in ("DR", "ER", "VAT", "WHR", "SR", "EXR", "D2R", "DSR", "DAR") if k in r}})
            else:
                misses += 1
                if misses >= MISSES_TO_STOP:
                    break
        print(f"chapter {ch:02d}: {sum(1 for l in lines if l['c'][:2] == f'{ch:02d}')} lines", flush=True)

    codes = [l["c"] + l["p"] for l in lines]
    if len(codes) != len(set(codes)):
        sys.exit("STOP: duplicate tariff codes found. Nothing written.")
    if len(lines) < 5000:
        sys.exit(f"STOP: only {len(lines)} tariff lines; expected more than 5000. Nothing written.")

    lists = {"unitOfMeasurment": [], "typeOfPackage": [], "customsRegime": [], "additionalCode": [], "country": []}
    units = sorted({u for l in lines for u in re.split(r"[\s,/]+", l["u"]) if u} | {"UNT", "NPR", "KGM", "LTR", "MTR", "MTK", "MTQ"})
    for u in units:
        lists["unitOfMeasurment"] += lookup("unitOfMeasurment", u)
    for pk in PACKAGES:
        lists["typeOfPackage"] += lookup("typeOfPackage", pk)
    for rg in ("4000", "4051", "4071", "4400", "4500", "4700", "4900"):
        lists["customsRegime"] += lookup("customsRegime", rg)
    for n in range(400, 500):
        lists["additionalCode"] += lookup("additionalCode", str(n))
    lists["additionalCode"] += lookup("additionalCode", "000")
    for c in ISO2:
        lists["country"] += lookup("country", c)
    for k, v in lists.items():
        print(f"list {k}: {len(v)}", flush=True)
    missing_units = [u for u in {u for l in lines for u in re.split(r"[\s,/]+", l["u"]) if u} if u not in {x["code"] for x in lists["unitOfMeasurment"]}]
    if missing_units:
        sys.exit(f"STOP: tariff uses units the portal list doesn't name: {missing_units}. Nothing written.")

    out = {"source": "https://customs.erca.gov.et/trade/help/tariff (public Ethiopia Customs Trade Portal)",
           "fetched": datetime.date.today().isoformat(), "count": len(lines), "lines": lines, "lists": lists}
    with open("tariff-data.js", "w", encoding="utf-8") as f:
        f.write("// Generated by tools/fetch_tariff.py. Do not edit by hand.\n")
        f.write("window.TARIFF = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n")
    print(f"Wrote tariff-data.js: {len(lines)} tariff lines, fetched {out['fetched']}")


if __name__ == "__main__":
    main()
