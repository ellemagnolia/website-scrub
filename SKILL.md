---
name: website-scrub
description: Weekly website scrub for Furniture + Textiles by MH Category or MH Sub-Category - every SKU that should be online (by stage) vs what's live on magnolia.com, outliers with a reason (not in Shopify, draft, no inventory, under stock buffer, has inventory - investigate), variants and image resolution. Use when the user says site scrub, website scrub, "run the scrub for <category or subcat>", what fell off the site, or checks items/variants/pictures online.
---

Follow `README.md` in this folder; it is the SOP. Scripts are in `scripts/`.

1. Confirm the category or subcat(s) (`--category` / `--subcat`) and get the newest MSL export and DS report from the user (or the latest they point to).
2. Step 2 runs in Chrome in the Shopify admin, only when the user is already signed in (never type
   a password). Read-only. Run `scripts/shopify_pull.js`; stop if the Jewell location changed.
3. Run `scripts/site_scrub.py`. Report the counts and list the "Has inventory - investigate" SKUs.
4. The user makes the calls on what the outliers mean. Teach formulas step by step if they ask.
   Read-only everywhere; don't send the file.
