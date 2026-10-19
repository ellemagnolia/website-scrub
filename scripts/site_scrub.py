"""Website scrub for one or more MH Categories and/or Sub-Categories -> F&T Site Scrub <subcat> <date>.xlsx

  python site_scrub.py --msl MSL.xlsx --ds DS.xlsx --shopify shopify_variants.tsv --subcat "Uph Seating" --out-dir .
  python site_scrub.py ... --category "Upholstery" --out-dir .          (a whole MH Category)

The workbook shows its work: the data tabs (MSL, DS, Shopify, HTML) hold exactly what was used,
and the answer tabs (Scrub, Outliers, Run) are formulas that point back at them.

Should be online = MSL rows in MH Division "Furniture + Textiles", the chosen MH Category(s) / Sub-Category(s),
Stage in STAGES. Each of those SKUs is matched (by SKU) to the live site (magnolia.com/products.json).
SKUs that aren't live get a Reason, checked in this order:
  Not in Shopify -> Draft/Archived in Shopify -> No inventory (<= 0) -> Under stock buffer (<= buffer)
  -> Has inventory - investigate
Inventory = DS report Quantity for dropship SKUs, Jewell Warehouse available for the rest.

shopify_variants.tsv = Shopify admin pull (README step 2): header, then sku, status, available, on_hand.
"""
import argparse
import datetime
import json
import os
import time
import urllib.request

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DIVISION = "Furniture + Textiles"
STAGES = ["Continuing", "New", "Markdown", "Residual", "Re-Intro"]
BUFFER = 10
INVESTIGATE = "Has inventory - investigate"
REASONS = [INVESTIGATE, "Not in Shopify", "Draft in Shopify", "Archived in Shopify", "Under stock buffer",
           "No inventory"]
QUALITIES = ["Good", "OK", "Low", "No variant photo", "No image"]
MSL_COLS = ["SKU", "Display Name", "Vendor", "Stage", "MH Category", "MH Sub-Category", "Is Dropship Item",
            "Launch Season", "Exit Date"]

DARK = PatternFill("solid", fgColor="3F3F3F")
DATA = PatternFill("solid", fgColor="DDEBF7")      # tab color: raw data I used
ANSWER = "548235"                                   # tab color: formulas / answers


def read_sheet(path, need):
    """Rows as dicts from the first sheet whose header row has every column in `need`."""
    wb = openpyxl.load_workbook(path, read_only=True)
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        head = [str(h).strip() if h is not None else "" for h in next(rows, [])]
        if all(n in head for n in need):
            return [dict(zip(head, r)) for r in rows]
    raise SystemExit(f"no sheet in {path} has columns {need}")


def read_shopify(path):
    """sku -> (status, available, on_hand, variant count). Prefers the ACTIVE record when a SKU is on
    more than one variant."""
    out = {}
    with open(path, encoding="utf-8-sig") as f:
        for line in f.read().splitlines()[1:]:
            sku, status, avail, on_hand = (line.split("\t") + [""] * 4)[:4]
            sku = sku.strip()
            if not sku:
                continue
            n = out[sku][3] + 1 if sku in out else 1
            num = lambda x: int(x) if x not in ("", None) else None
            if sku in out and out[sku][0] == "ACTIVE":
                out[sku] = out[sku][:3] + (n,)
            else:
                out[sku] = (status, num(avail), num(on_hand), n)
    return out


def pull_site(log=print):
    prods, page = [], 1
    while True:
        req = urllib.request.Request(f"https://magnolia.com/products.json?limit=250&page={page}",
                                     headers={"User-Agent": "Mozilla/5.0"})
        batch = json.loads(urllib.request.urlopen(req, timeout=60).read()).get("products", [])
        if not batch:
            break
        prods += batch
        page += 1
        time.sleep(0.5)
    log(f"site: {len(prods)} products")
    return prods


def site_variants(prods):
    """sku -> the live variant, with the image a shopper sees for it."""
    out = {}
    for p in prods:
        main = p["images"][0] if p["images"] else None
        multi = len(p["variants"]) > 1
        for v in p["variants"]:
            sku = (v.get("sku") or "").strip()
            if not sku:
                continue
            own = v.get("featured_image")
            img = own or (None if multi else main)
            out[sku] = {"product": p["title"], "variant": "" if v["title"] == "Default Title" else v["title"],
                        "vendor": p.get("vendor") or "", "in_stock": "Yes" if v.get("available") else "No",
                        "variants": len(p["variants"]), "own": "Yes" if own else "No",
                        "w": img["width"] if img else None, "h": img["height"] if img else None,
                        "main_w": main["width"] if main else None, "main_h": main["height"] if main else None,
                        "url": f"https://magnolia.com/products/{p['handle']}?variant={v['id']}"}
    return out


def style(ws, widths, fill=DARK, header_row=1, freeze=True):
    for c in ws[header_row]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=header_row + 1):
        for c in row:
            c.font = Font(name="Arial", bold=c.font.bold)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if freeze:
        ws.freeze_panes = ws.cell(header_row + 1, 1)
    if ws.max_row > header_row:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"


def rng(tab, col, n):
    """Absolute range for one column of a data tab, e.g. HTML!$A$2:$A$812 (never whole-column)."""
    return f"'{tab}'!${col}$2:${col}${max(n, 1) + 1}"


def build(msl_path, ds_path, shopify_path, subcats, out_dir, prods=None, sources=None, cats=()):
    sources = sources or {}
    msl_all = read_sheet(msl_path, ["SKU", "Stage", "MH Division", "MH Sub-Category"])
    in_div = [r for r in msl_all if r["MH Division"] == DIVISION and r["SKU"]]
    subcats, cats = list(subcats or []), list(cats or [])
    for picked, col in ((cats, "MH Category"), (subcats, "MH Sub-Category")):
        known = {r[col] for r in in_div if r[col]}
        bad = [x for x in picked if x not in known]
        if bad:
            raise SystemExit(f"not an F+T {col}: {bad}. Choices: {sorted(known)}")
    in_sub = [r for r in in_div if r["MH Category"] in cats or r["MH Sub-Category"] in subcats]
    should = [r for r in in_sub if r["Stage"] in STAGES]
    scope = " + ".join(cats + subcats)
    scope_rule = " or ".join(([f"MH Category = {', '.join(cats)}"] if cats else []) +
                             ([f"MH Sub-Category = {', '.join(subcats)}"] if subcats else []))
    if not should:
        raise SystemExit(f"no {DIVISION} SKUs in {scope} at stages {STAGES}")
    sub_skus = {str(r["SKU"]).strip() for r in in_sub}
    ds_rows = read_sheet(ds_path, ["SKU", "Quantity"]) if ds_path else []
    ds = {str(r["SKU"]).strip(): r for r in ds_rows if r["SKU"] and str(r["SKU"]).strip() in sub_skus}
    shop_all = read_shopify(shopify_path)
    shop = {k: v for k, v in shop_all.items() if k in sub_skus}
    prods = prods if prods is not None else pull_site()
    site_all = site_variants(prods)
    site = {k: v for k, v in site_all.items() if k in sub_skus}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    run = wb.create_sheet("Run")
    steps = wb.create_sheet("Steps")
    outl = wb.create_sheet("Outliers")
    scrub = wb.create_sheet("Scrub")

    # ---- data tabs: exactly what was used -------------------------------------------------------
    ws = wb.create_sheet("MSL")
    ws.append(MSL_COLS + ["Should Be Online?"])
    msl_sorted = sorted(in_sub, key=lambda r: (r["MH Category"] or "", r["MH Sub-Category"] or "", r["Display Name"] or "", str(r["SKU"])))
    for i, r in enumerate(msl_sorted, 2):
        ws.append([str(r.get(c)).strip() if c == "SKU" else r.get(c) for c in MSL_COLS] +
                  [f'=IF(COUNTIF(Run!$B$4:$B${3 + len(STAGES)},D{i})>0,"Yes","No")'])
    style(ws, [18, 42, 30, 12, 16, 22, 10, 14, 12, 12], DARK)
    ws.sheet_properties.tabColor = "9BC2E6"
    n_msl = len(msl_sorted)

    ws = wb.create_sheet("HTML")
    ws.append(["SKU", "Product Name", "Variant", "Vendor", "In Stock", "Variants on Product", "Own Photo?",
               "Image Width", "Image Height", "Shorter Side", "Image Quality", "Main Image Size", "URL"])
    html_sorted = sorted(site.items(), key=lambda kv: (kv[1]["product"], kv[1]["variant"]))
    for i, (sku, v) in enumerate(html_sorted, 2):
        main = f"{v['main_w']}x{v['main_h']}" if v["main_w"] else ""
        ws.append([sku, v["product"], v["variant"], v["vendor"], v["in_stock"], v["variants"], v["own"],
                   v["w"], v["h"], f'=IF(H{i}="","",MIN(H{i},I{i}))',
                   f'=IF(AND(F{i}>1,G{i}="No"),"No variant photo",IF(J{i}="","No image",'
                   f'IF(J{i}<Run!$B$11,"Low",IF(J{i}<Run!$B$12,"OK","Good"))))', main, v["url"]])
        ws.cell(i, 13).hyperlink = v["url"]
    style(ws, [18, 42, 26, 28, 9, 10, 10, 10, 10, 10, 16, 16, 50])
    ws.sheet_properties.tabColor = "9BC2E6"
    n_html = len(html_sorted)

    ws = wb.create_sheet("Shopify")
    ws.append(["SKU", "Product Status", "Jewell Available", "Jewell On Hand", "Variants with this SKU"])
    for sku in sorted(shop):
        st, av, oh, n = shop[sku]
        ws.append([sku, st, av, oh, n])
    style(ws, [18, 14, 14, 14, 12])
    ws.sheet_properties.tabColor = "9BC2E6"
    n_shop = len(shop)

    ws = wb.create_sheet("DS")
    ds_cols = list(ds_rows[0].keys()) if ds_rows else ["SKU", "Quantity"]
    ds_cols = ["SKU", "Quantity"] + [c for c in ds_cols if c not in ("SKU", "Quantity") and c]
    ws.append(ds_cols)
    for sku in sorted(ds):
        ws.append([sku if c == "SKU" else ds[sku].get(c) for c in ds_cols])
    style(ws, [18, 10] + [16] * (len(ds_cols) - 2))
    ws.sheet_properties.tabColor = "9BC2E6"
    n_ds = len(ds)

    # ---- Scrub: one row per should-be-online SKU, every answer a formula -------------------------
    M = lambda col: rng("MSL", col, n_msl)
    H = lambda col: rng("HTML", col, n_html)
    S = lambda col: rng("Shopify", col, n_shop)
    D = lambda col: rng("DS", col, n_ds)
    scrub.append(["SKU", "Name", "Variant", "Active", "Image Quality", "Inventory", "Inventory Source",
                  "Shopify Status", "Reason", "Dropship?", "Sub-Category", "Vendor", "Stage"])
    should_sorted = sorted(should, key=lambda r: (r["MH Category"] or "", r["MH Sub-Category"] or "", r["Display Name"] or "", str(r["SKU"])))
    for i, r in enumerate(should_sorted, 2):
        x = lambda col, i=i: f'_xlfn.XLOOKUP(A{i},{M("A")},{M(col)},"")'
        scrub.append([
            str(r["SKU"]).strip(),
            f"={x('B')}",
            f'=IF(D{i}="Yes",_xlfn.XLOOKUP(A{i},{H("A")},{H("C")},""),"")',
            f'=IF(COUNTIF({H("A")},A{i})>0,"Yes","No")',
            f'=IF(D{i}="Yes",_xlfn.XLOOKUP(A{i},{H("A")},{H("K")},""),"")',
            f'=IF(J{i}="Yes",_xlfn.XLOOKUP(A{i},{D("A")},{D("B")},""),_xlfn.XLOOKUP(A{i},{S("A")},{S("C")},""))',
            f'=IF(J{i}="Yes",IF(COUNTIF({D("A")},A{i})>0,"DS","DS (not on report)"),'
            f'IF(F{i}="","Jewell (no record)","Jewell"))',
            f'=_xlfn.XLOOKUP(A{i},{S("A")},{S("B")},"Not in Shopify")',
            f'=IF(D{i}="Yes","",IF(H{i}="Not in Shopify","Not in Shopify",IF(H{i}<>"ACTIVE",PROPER(H{i})&" in Shopify",'
            f'IF(N(F{i})<=0,"No inventory",IF(F{i}<=Run!$B$10,"Under stock buffer","{INVESTIGATE}")))))',
            f"={x('G')}", f"={x('F')}", f"={x('C')}", f"={x('D')}",
        ])
    style(scrub, [18, 42, 26, 8, 16, 10, 18, 14, 26, 10, 22, 28, 12])
    scrub.sheet_properties.tabColor = ANSWER
    n_scrub = len(should_sorted)
    SC = lambda col: f"Scrub!${col}$2:${col}${n_scrub + 1}"

    # Python works out the same answers so the outlier list (and its order) can be written out
    site_skus, rows = set(site), []
    for r in should_sorted:
        sku = str(r["SKU"]).strip()
        if sku in site_skus:
            continue
        dropship = r.get("Is Dropship Item") == "Yes"
        inv = ds[sku]["Quantity"] if dropship and sku in ds else None if dropship else (shop[sku][1] if sku in shop else None)
        st = shop[sku][0] if sku in shop else None
        rsn = ("Not in Shopify" if st is None else f"{st.title()} in Shopify" if st != "ACTIVE" else
               "No inventory" if inv is None or inv <= 0 else "Under stock buffer" if inv <= BUFFER else INVESTIGATE)
        rows.append((REASONS.index(rsn) if rsn in REASONS else len(REASONS), r["Display Name"] or "", sku))
    rows.sort()

    # ---- Outliers: SKUs where Scrub Active = No; everything else looked up from Scrub ------------
    outl.append(["SKU", "Name", "Reason", "Inventory", "Inventory Source", "Shopify Status", "Dropship?",
                 "Sub-Category", "Vendor", "Stage"])
    for i, (_, _, sku) in enumerate(rows, 2):
        lk = lambda col, i=i: f'=_xlfn.XLOOKUP($A{i},{SC("A")},{SC(col)},"")'
        outl.append([sku, lk("B"), lk("I"), lk("F"), lk("G"), lk("H"), lk("J"), lk("K"), lk("L"), lk("M")])
    style(outl, [18, 42, 26, 10, 18, 14, 10, 22, 28, 12])
    outl.sheet_properties.tabColor = ANSWER

    # ---- Run: settings you can change + counts ---------------------------------------------------
    run.append(["Website Scrub", ""])
    run.append(["Scope", scope_rule])
    run.append(["Run", f"{datetime.datetime.now():%m/%d/%Y %I:%M %p}"])
    for k, s in enumerate(STAGES):
        run.append(["Stages that should be online" if k == 0 else "", s])
    run.append(["", ""])
    run.append(["Stock buffer (change me)", BUFFER])                                # B10
    run.append(["Image: Low if shorter side under", 1000])                          # B11
    run.append(["Image: Good if shorter side at least", 1500])                      # B12
    run.append(["", ""])
    run.append(["Counts", ""])
    first = run.max_row + 1
    run.append(["Should be online", f"=COUNTA({SC('A')})"])
    run.append(["Online", f'=COUNTIF({SC("D")},"Yes")'])
    run.append(["Outliers (not online)", f'=COUNTIF({SC("D")},"No")'])
    run.append(["", ""])
    run.append(["Outliers by reason", ""])
    for rsn in REASONS:
        run.append([rsn, f'=COUNTIF({SC("I")},"{rsn}")'])
    run.append(["", ""])
    run.append(["Image quality (online SKUs)", ""])
    for q in QUALITIES:
        run.append([q, f'=COUNTIF({SC("E")},"{q}")'])
    style(run, [36, 40], freeze=False)
    run.auto_filter.ref = None
    for row in (10, 11, 12):
        run.cell(row, 2).fill = PatternFill("solid", fgColor="FFF2CC")
    for row in run.iter_rows(min_row=first):
        if row[0].value in ("Counts", "Outliers by reason", "Image quality (online SKUs)"):
            row[0].font = Font(name="Arial", bold=True)
    run.sheet_properties.tabColor = ANSWER

    # ---- Steps: what was done, in order, with the numbers at each step ---------------------------
    n_dup = sum(1 for k in shop if shop[k][3] > 1)
    steps.append(["#", "What I did", "Result", "Where to see it"])
    for row in [
        [1, f"Read the MSL export ({sources.get('msl', os.path.basename(msl_path))})", f"{len(msl_all):,} rows", ""],
        [2, f'Kept MH Division = "{DIVISION}"', f"{len(in_div):,} SKUs", ""],
        [3, f"Kept {scope_rule}", f"{len(in_sub):,} SKUs (all stages)", "MSL tab"],
        [4, f"Kept Stage = {', '.join(STAGES)} (the 'should be online' list)", f"{len(should):,} SKUs",
         "MSL tab, Should Be Online? = Yes; Scrub tab"],
        [5, "Pulled every live variant from magnolia.com/products.json" + (f" ({sources['site']})" if sources.get("site") else ""),
         f"{len(prods):,} products / {len(site_all):,} SKUs on the site", ""],
        [6, "Kept site SKUs that are in this subcat on the MSL", f"{n_html:,} SKUs", "HTML tab"],
        [7, "Pulled every Shopify variant: product status + Jewell Warehouse available / on hand" +
         (f" ({sources['shopify']})" if sources.get("shopify") else ""),
         f"{len(shop_all):,} SKUs in Shopify; {n_shop:,} in this subcat"
         + (f" ({n_dup} SKU(s) on more than one variant: the Active one is used)" if n_dup else ""), "Shopify tab"],
        [8, "Read the DS report (dropship quantities)" + (f" ({sources.get('ds', os.path.basename(ds_path))})" if ds_path else ": none given"),
         f"{n_ds:,} SKUs in this subcat", "DS tab"],
        [9, "Matched each should-be-online SKU to the site by SKU (never by name)",
         "Active = Yes if the SKU is on the HTML tab", "Scrub tab, column D"],
        [10, "Image quality = the photo a shopper sees for that variant, rated on its shorter side",
         "Good 1500+, OK 1000-1499, Low under 1000; a variant with no photo of its own on a multi-variant product = No variant photo",
         "HTML tab, columns G-K; Scrub column E"],
        [11, "Inventory = DS Quantity if the MSL says dropship, otherwise Jewell available", "",
         "Scrub columns F-G"],
        [12, "Reason for each SKU that isn't online, first match wins",
         "Not in Shopify -> Draft/Archived in Shopify -> No inventory (0 or less) -> Under stock buffer "
         "(1 to the buffer on Run) -> Has inventory - investigate", "Scrub column I"],
        [13, "Listed the SKUs that aren't online, 'investigate' first", f"{len(rows):,} outliers", "Outliers tab"],
    ]:
        steps.append(row)
    style(steps, [5, 70, 60, 36])
    steps.auto_filter.ref = None
    for row in steps.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    steps.sheet_properties.tabColor = ANSWER

    wb.calculation.fullCalcOnLoad = True
    picked = cats + subcats
    name = " + ".join(picked) if len(picked) <= 3 else f"{len(picked)} groups"
    out = os.path.join(out_dir, f"F&T Site Scrub {name.replace('/', '-')} {datetime.date.today():%m.%d.%y}.xlsx")
    wb.save(out)
    counts = {rsn: sum(1 for r in rows if r[0] == REASONS.index(rsn)) for rsn in REASONS}
    print(f"should be online {len(should)}  online {len(should) - len(rows)}  outliers {len(rows)}")
    print("  " + "  ".join(f"{k}: {v}" for k, v in counts.items() if v))
    print("saved", out)
    return out


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--msl", required=True, help="workbook with the raw MSL export (any sheet with SKU/Stage/MH columns)")
    parser.add_argument("--ds", help="DS inventory report (sheet with SKU + Quantity); optional")
    parser.add_argument("--shopify", required=True, help="shopify_variants.tsv from the admin pull")
    parser.add_argument("--category", action="append", default=[], help="MH Category (e.g. Upholstery); repeat for more than one")
    parser.add_argument("--subcat", action="append", default=[], help="MH Sub-Category (e.g. Uph Seating); repeat for more than one")
    parser.add_argument("--site-json", help="reuse a saved products.json pull instead of fetching")
    parser.add_argument("--out-dir", default=".")
    args = parser.parse_args()
    prods = json.load(open(args.site_json, encoding="utf-8")) if args.site_json else None
    sources = {"shopify": f"pulled {datetime.datetime.fromtimestamp(os.path.getmtime(args.shopify)):%m/%d %I:%M %p}"}
    if args.site_json:
        sources["site"] = f"pulled {datetime.datetime.fromtimestamp(os.path.getmtime(args.site_json)):%m/%d %I:%M %p}"
    if not args.category and not args.subcat:
        parser.error("give at least one --category or --subcat")
    build(args.msl, args.ds, args.shopify, args.subcat, args.out_dir, prods, sources, args.category)


if __name__ == "__main__":
    main()
