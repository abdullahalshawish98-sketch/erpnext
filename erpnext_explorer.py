#!/usr/bin/env python3
"""
ERPNext read-only explorer (for Claude Code)

Commands:
  schema  <DocType>              list every field (standard + custom + child tables)
  map     [--doctypes ...]       show link fields that connect the documents together
  sample  <DocType>              print one full recent document (with child tables)
  fetch   <DocType> [options]    pull all records with ALL fields -> data/*.json + *.csv
  export  [--doctypes ...]       fetch several DocTypes at once
  unpaid  [--source order|invoice] [--top N]           oldest orders/invoices not fully paid
  debtors [--source order|invoice] [--by amount|count] top customers by outstanding debt

Credentials come from environment variables or a .env file:
  ERPNEXT_URL, ERPNEXT_API_KEY, ERPNEXT_API_SECRET
Only GET requests are ever sent. No dependencies (standard library only).
"""
import argparse
import csv
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path

DEFAULT_DOCTYPES = ["Sales Order", "Purchase Invoice"]
LAYOUT_TYPES = {"Section Break", "Column Break", "Tab Break", "HTML", "Button", "Heading"}


def load_env():
    p = Path(".env")
    if not p.exists():
        return
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_env()
URL = os.environ.get("ERPNEXT_URL", "").rstrip("/")
KEY = os.environ.get("ERPNEXT_API_KEY", "")
SECRET = os.environ.get("ERPNEXT_API_SECRET", "")
if not (URL and KEY and SECRET):
    sys.exit("Missing ERPNEXT_URL / ERPNEXT_API_KEY / ERPNEXT_API_SECRET (put them in .env)")


def api(path, params=None):
    query = ""
    if params:
        query = "?" + urllib.parse.urlencode(
            {k: json.dumps(v) if isinstance(v, (list, dict)) else v for k, v in params.items()}
        )
    req = urllib.request.Request(
        URL + path + query,
        headers={"Authorization": f"token {KEY}:{SECRET}", "Accept": "application/json"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")[:300]
        sys.exit(f"HTTP {e.code} on {path}: {body}")
    except urllib.error.URLError as e:
        sys.exit(f"Connection error: {e.reason}")


def q(s):
    return urllib.parse.quote(s, safe="")


def get_schema(doctype):
    """Return {doctype_name: [fields]} for the DocType and all its child tables."""
    r = api("/api/method/frappe.desk.form.load.getdoctype", {"doctype": doctype})
    docs = r.get("docs") or r.get("message", {}).get("docs", [])
    return {d["name"]: d.get("fields", []) for d in docs}


def fetch_all(doctype, filters=None, page=500, full=False):
    rows, start = [], 0
    while True:
        batch = api(
            f"/api/resource/{q(doctype)}",
            {
                "fields": ["*"],
                "filters": filters or [],
                "limit_start": start,
                "limit_page_length": page,
                "order_by": "creation desc",
            },
        ).get("data", [])
        rows += batch
        print(f"  {doctype}: {len(rows)} records", end="\r", file=sys.stderr)
        if len(batch) < page:
            break
        start += page
    print(file=sys.stderr)
    if full:  # the list API omits child tables (e.g. items); fetch each document fully
        for i, r in enumerate(rows):
            rows[i] = api(f"/api/resource/{q(doctype)}/{q(r['name'])}")["data"]
            print(f"  full documents: {i + 1}/{len(rows)}", end="\r", file=sys.stderr)
        print(file=sys.stderr)
    return rows


def save(doctype, rows):
    out = Path("data")
    out.mkdir(exist_ok=True)
    base = out / doctype.replace(" ", "_")
    base.with_suffix(".json").write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    if rows:
        cols = list({k: None for r in rows for k in r})
        with open(base.with_suffix(".csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for r in rows:
                w.writerow({k: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
                            for k, v in r.items()})
    print(f"  saved {len(rows)} rows -> {base}.json / .csv")


def since(days):
    return [["creation", ">=", str(date.today() - timedelta(days=days))]] if days else []


def cmd_schema(a):
    for name, fields in get_schema(a.doctype).items():
        print(f"\n== {name} ==")
        for f in fields:
            if f.get("fieldtype") in LAYOUT_TYPES:
                continue
            tag = " [custom]" if str(f.get("fieldname", "")).startswith("custom_") else ""
            print(f"  {f['fieldname']:<34} {f['fieldtype']:<14} {f.get('label') or ''}"
                  f"{'  -> ' + f['options'] if f.get('options') and f['fieldtype'] in ('Link', 'Table') else ''}{tag}")


def cmd_map(a):
    names = set(a.doctypes)
    for dt in a.doctypes:
        for name, fields in get_schema(dt).items():
            rows = [f for f in fields if f.get("fieldtype") in ("Link", "Dynamic Link", "Table", "Table MultiSelect")]
            if not rows:
                continue
            print(f"\n== {name} ==")
            for f in rows:
                hit = "   <== يربط بمستند من مستنداتك" if f.get("options") in names else ""
                print(f"  {f['fieldname']:<34} {f['fieldtype']:<12} -> {f.get('options', '')}{hit}")


def cmd_sample(a):
    rows = api(f"/api/resource/{q(a.doctype)}",
               {"fields": ["name"], "limit_page_length": 1, "order_by": "creation desc"}).get("data", [])
    if not rows:
        sys.exit("No records found")
    doc = api(f"/api/resource/{q(a.doctype)}/{q(rows[0]['name'])}")["data"]
    print(json.dumps(doc, ensure_ascii=False, indent=1))


def cmd_fetch(a):
    print(f"Fetching {a.doctype} ...")
    save(a.doctype, fetch_all(a.doctype, since(a.days), full=a.full))


def cmd_export(a):
    for dt in a.doctypes:
        print(f"Fetching {dt} ...")
        save(dt, fetch_all(dt, since(a.days), full=a.full))


def outstanding_rows(source):
    """Rows of unpaid amounts. source='invoice' uses Sales Invoice outstanding_amount;
    source='order' uses Sales Order grand_total - advance_paid."""
    out = []
    if source == "invoice":
        flt = [["docstatus", "=", 1], ["outstanding_amount", ">", 0]]
        for r in fetch_all("Sales Invoice", flt):
            out.append({"doc": r["name"], "customer": r.get("customer_name") or r.get("customer"),
                        "date": r.get("posting_date"), "total": r.get("grand_total") or 0,
                        "outstanding": r.get("outstanding_amount") or 0, "currency": r.get("currency")})
    else:
        flt = [["docstatus", "!=", 2], ["status", "not in", ["Cancelled", "Closed"]]]
        for r in fetch_all("Sales Order", flt):
            due = (r.get("grand_total") or 0) - (r.get("advance_paid") or 0)
            if due > 0.005:
                out.append({"doc": r["name"], "customer": r.get("customer_name") or r.get("customer"),
                            "date": r.get("transaction_date"), "total": r.get("grand_total") or 0,
                            "outstanding": due, "currency": r.get("currency")})
    today = date.today()
    for r in out:
        r["age_days"] = (today - date.fromisoformat(str(r["date"])[:10])).days if r["date"] else None
    if not out:
        print("لا توجد مبالغ غير مسددة بهذا المصدر. جرّب --source " + ("order" if source == "invoice" else "invoice"))
    return out


def show(rows, cols):
    w = {c: max([len(c)] + [len(str(r.get(c, ""))) for r in rows]) for c in cols}
    print("  ".join(c.ljust(w[c]) for c in cols))
    for r in rows:
        print("  ".join(str(r.get(c, "")).ljust(w[c]) for c in cols))


def money(rows, keys):
    return [{**r, **{k: f"{r[k]:,.2f}" for k in keys if k in r}} for r in rows]


def cmd_unpaid(a):
    rows = outstanding_rows(a.source)
    if not rows:
        return
    rows.sort(key=lambda r: r["date"] or "9999")
    save("unpaid_" + a.source, rows)
    print(f"\nأقدم {min(a.top, len(rows))} من أصل {len(rows)} غير مسددة:\n")
    show(money(rows[:a.top], ["total", "outstanding"]),
         ["doc", "customer", "date", "age_days", "total", "outstanding", "currency"])


def cmd_debtors(a):
    rows = outstanding_rows(a.source)
    if not rows:
        return
    agg = {}
    for r in rows:
        g = agg.setdefault((r["customer"], r["currency"]),
                           {"customer": r["customer"], "currency": r["currency"], "orders": 0,
                            "outstanding": 0.0, "oldest": r["date"], "oldest_age_days": r["age_days"]})
        g["orders"] += 1
        g["outstanding"] += r["outstanding"]
        if r["date"] and (not g["oldest"] or r["date"] < g["oldest"]):
            g["oldest"], g["oldest_age_days"] = r["date"], r["age_days"]
    key = (lambda g: g["orders"]) if a.by == "count" else (lambda g: g["outstanding"])
    res = sorted(agg.values(), key=key, reverse=True)
    save("debtors_" + a.source, res)
    print(f"\nأكثر {min(a.top, len(res))} عملاء ديوناً (من أصل {len(res)}):\n")
    show(money(res[:a.top], ["outstanding"]),
         ["customer", "orders", "outstanding", "currency", "oldest", "oldest_age_days"])
    totals = {}
    for g in res:
        totals[g["currency"]] = totals.get(g["currency"], 0) + g["outstanding"]
    print("\nإجمالي الديون: " + " | ".join(f"{v:,.2f} {c}" for c, v in totals.items()))


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("schema"); s.add_argument("doctype"); s.set_defaults(fn=cmd_schema)
    s = sub.add_parser("map"); s.add_argument("--doctypes", nargs="+", default=DEFAULT_DOCTYPES); s.set_defaults(fn=cmd_map)
    s = sub.add_parser("sample"); s.add_argument("doctype"); s.set_defaults(fn=cmd_sample)

    s = sub.add_parser("fetch"); s.add_argument("doctype")
    s.add_argument("--days", type=int, default=0, help="only records created in the last N days")
    s.add_argument("--full", action="store_true", help="include child tables (slower)")
    s.set_defaults(fn=cmd_fetch)

    s = sub.add_parser("export"); s.add_argument("--doctypes", nargs="+", default=DEFAULT_DOCTYPES)
    s.add_argument("--days", type=int, default=0); s.add_argument("--full", action="store_true")
    s.set_defaults(fn=cmd_export)

    for name, fn in (("unpaid", cmd_unpaid), ("debtors", cmd_debtors)):
        s = sub.add_parser(name)
        s.add_argument("--source", choices=["invoice", "order"], default="order",
                       help="invoice = Sales Invoice outstanding | order = Sales Order total minus advance paid")
        s.add_argument("--top", type=int, default=20)
        if name == "debtors":
            s.add_argument("--by", choices=["amount", "count"], default="amount")
        s.set_defaults(fn=fn)

    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
