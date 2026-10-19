// Step 2: every Shopify variant -> sku, product status, Jewell available, Jewell on hand.
// Paste into the javascript tool on any admin.shopify.com/store/magnoliamom page (signed in).
// Read-only: GraphQL queries only. Runs in the background; check window.jw until done is true
// (about 3 min for ~43k variants), then copy window.tsv out through the clipboard (README step 2).
const LOCATION = "gid://shopify/Location/42661508"; // Jewell Warehouse

window.gq = async (query, variables = {}) => {
  const t = JSON.parse(document.querySelector('script[data-serialized-id="server-data"]').textContent).csrfToken;
  const r = await fetch("/api/shopify/magnoliamom?operation=Q&type=query", {
    method: "POST",
    headers: { "content-type": "application/json", accept: "application/json", "x-csrf-token": t },
    body: JSON.stringify({ query, variables, operationName: "Q" }),
  });
  return r.json();
};

window.jw = { rows: [], done: false, err: null, pages: 0 };
(async () => {
  const loc = (await gq("query Q{locations(first:100, includeInactive:true){nodes{id name}}}")).data.locations.nodes;
  if (!loc.some((n) => n.id === LOCATION && /Jewell/i.test(n.name))) {
    jw.err = "Jewell Warehouse location changed: " + JSON.stringify(loc.filter((n) => /Jewell/i.test(n.name)));
    jw.done = true;
    return;
  }
  const q = `query Q($after:String){productVariants(first:200, after:$after){pageInfo{hasNextPage endCursor}
    nodes{sku product{status} inventoryItem{inventoryLevel(locationId:"${LOCATION}"){
    quantities(names:["available","on_hand"]){name quantity}}}}}}`;
  let after = null;
  try {
    while (true) {
      let r, tries = 0;
      while (true) {
        r = await gq(q, { after });
        if (r.data) break;
        if (++tries > 4) throw new Error(JSON.stringify(r).slice(0, 300));
        await new Promise((s) => setTimeout(s, 3000));
      }
      const pv = r.data.productVariants;
      for (const n of pv.nodes) {
        const l = n.inventoryItem?.inventoryLevel;
        const g = (k) => (l ? (l.quantities.find((x) => x.name === k) || {}).quantity : "");
        jw.rows.push([n.sku || "", n.product?.status || "", g("available"), g("on_hand")].join("\t"));
      }
      jw.pages++;
      if (!pv.pageInfo.hasNextPage) break;
      after = pv.pageInfo.endCursor;
    }
  } catch (e) {
    jw.err = String(e);
  }
  window.tsv = "sku\tstatus\tavailable\ton_hand\n" + jw.rows.join("\n");
  jw.done = true;
})();
"started";
