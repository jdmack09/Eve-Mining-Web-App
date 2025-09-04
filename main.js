/* main.js — EVE Mining Tracker (browser)
   - Paste mining hold (from inventory) into <textarea id="miningHold">
   - Click Generate to build report
   - Live Jita prices (EVEMarketer region 10000002)
   - TypeID + m³ via Fuzzwork (typeid2)
   - Categorized sheets + Grand Total; Excel export
*/

// ---------- Utilities ----------
const sleep = ms => new Promise(r => setTimeout(r, ms));
const fmtISK = n => Number(n || 0).toLocaleString();
const stripCommas = s => String(s).replace(/[,\s]/g, "");

// ---------- Category detection ----------
function detectCategory(name) {
  const n = name.toLowerCase();

  // Trig ores
  if (/(bezdnacine|rakovene|talassonite)/i.test(name)) return "Triglavian";

  // Asteroid ores (HS/LS/NS + variants)
  if (/(veldspar|scordite|pyroxeres|plagioclase|omber|kernite|jaspet|hemorphite|hedbergite|gneiss|ochre|crokite|bistot|arkonor|spodumain|mercoxit)/i.test(name))
    return "Asteroid";

  // Moon ores (ubiquitous → exceptional)
  if (/(bitumens|coesite|zeolites|sylvite|cobaltite|eu|euxenite|scheelite|titanite|otavite|sperrylite|vanadinite|chromite|carnotite|zircon|pollucite|ytterbite|monazite|loparite|xenotime)/i.test(name))
    return "Moon";

  // Gas (WH gases + k-space boosters)
  if (/(fullerite|mykoserocin|cytoserocin)/i.test(name))
    return "Gas";

  // Ice products
  if (/(blue ice|white glaze|glacial mass|glare crust|dark glitter|gelidus|krystallos|clear icicle|thick blue ice)/i.test(name)
      || /ice/i.test(name))
    return "Ice";

  return "Other";
}

// ---------- Static name registry (all base + compressed) ----------
// NOTE: This is for *categorization and sanity checks only*.
// We still resolve TypeID & volume live per *exact* pasted name.
const NAME_REGISTRY = (() => {
  const asteroid = [
    // Veldspar family
    "Veldspar","Concentrated Veldspar","Dense Veldspar",
    // Scordite
    "Scordite","Condensed Scordite","Massive Scordite",
    // Pyroxeres
    "Pyroxeres","Solid Pyroxeres","Viscous Pyroxeres",
    // Plagioclase
    "Plagioclase","Azure Plagioclase","Rich Plagioclase",
    // Omber
    "Omber","Silvery Omber","Golden Omber",
    // Kernite
    "Kernite","Luminous Kernite","Fiery Kernite",
    // Jaspet
    "Jaspet","Pure Jaspet","Pristine Jaspet",
    // Hemorphite
    "Hemorphite","Vivid Hemorphite","Radiant Hemorphite",
    // Hedbergite
    "Hedbergite","Vitric Hedbergite","Glazed Hedbergite",
    // Gneiss
    "Gneiss","Iridescent Gneiss","Prismatic Gneiss",
    // Dark Ochre
    "Dark Ochre","Onyx Ochre","Obsidian Ochre",
    // Crokite
    "Crokite","Sharp Crokite","Crystalline Crokite",
    // Bistot
    "Bistot","Triclinic Bistot","Monoclinic Bistot",
    // Arkonor
    "Arkonor","Crimson Arkonor","Prime Arkonor",
    // Spodumain
    "Spodumain","Bright Spodumain","Gleaming Spodumain",
    // Mercoxit
    "Mercoxit","Magma Mercoxit","Stable Mercoxit",
  ];

  const moon = [
    // Ubiquitous
    "Bitumens","Coesite","Sylvite","Zeolites",
    // Common
    "Cobaltite","Euxenite","Scheelite","Titanite",
    // Uncommon
    "Otavite","Sperrylite","Vanadinite","Chromite",
    // Rare
    "Carnotite","Zircon","Pollucite","Monazite?",
    // Exceptional
    "Xenotime","Monazite","Loparite","Ytterbite",
  ];

  // Triglavian (Abyssal belts)
  const trig = ["Bezdnacine","Rakovene","Talassonite"];

  // Wormhole & booster gases
  const gas = [
    // Mykoserocin (k-space)
    "Amber Mykoserocin","Golden Mykoserocin","Celadon Mykoserocin",
    "Lime Mykoserocin","Malachite Mykoserocin","Vermillion Mykoserocin",
    "Teal Mykoserocin","Azure Mykoserocin",
    // Cytoserocin (k-space)
    "Azure Cytoserocin","Chartreuse Cytoserocin","Celadon Cytoserocin",
    "Lime Cytoserocin","Malachite Cytoserocin","Amber Cytoserocin",
    "Vermillion Cytoserocin","Teal Cytoserocin",
    // Fullerites (wormhole)
    "Fullerite-C28","Fullerite-C32","Fullerite-C50","Fullerite-C60",
    "Fullerite-C70","Fullerite-C72","Fullerite-C84","Fullerite-C320","Fullerite-C540",
  ];

  // Ice
  const ice = [
    "Blue Ice","Clear Icicle","Glacial Mass","White Glaze","Glare Crust",
    "Dark Glitter","Gelidus","Krystallos","Thick Blue Ice"
  ];

  // Add compressed versions
  function withCompressed(list) {
    return [...list, ...list.map(n => `Compressed ${n}`)];
  }

  return new Set([
    ...withCompressed(asteroid),
    ...withCompressed(moon),
    ...withCompressed(trig),
    ...withCompressed(gas),
    ...withCompressed(ice),
  ]);
})();

// ---------- Caches to avoid duplicate remote lookups ----------
const typeCache = new Map();   // name -> {typeID, volume, name}
const priceCache = new Map();  // typeID -> price

// Resolve exact name (including "Compressed ...") to {typeID, volume, name}
async function resolveType(name) {
  if (typeCache.has(name)) return typeCache.get(name);

  // Try exact first
  let resolved = await fetchTypeid2(name);
  // If not found, try stripping "Compressed "
  if (!resolved && /^Compressed\s+/i.test(name)) {
    const base = name.replace(/^Compressed\s+/i, "").trim();
    resolved = await fetchTypeid2(base);
  }

  typeCache.set(name, resolved || null);
  return resolved;
}

async function fetchTypeid2(exactName) {
  try {
    const url = `https://www.fuzzwork.co.uk/api/typeid2.php?typename=${encodeURIComponent(exactName)}`;
    const r = await fetch(url);
    const data = await r.json();
    if (!data || !data.typeid) return null;
    return {
      typeID: Number(data.typeid),
      volume: Number(data.volume || 0),
      name: data.name || exactName
    };
  } catch (e) {
    console.error("typeid2 lookup failed:", exactName, e);
    return null;
  }
}

// Live Jita price for a typeID (sell min)
async function getPrice(typeID) {
  if (!typeID) return 0;
  if (priceCache.has(typeID)) return priceCache.get(typeID);
  try {
    const url = `https://api.evemarketer.com/ec/marketstat/json?typeid=${typeID}&regionlimit=10000002`;
    const r = await fetch(url);
    const data = await r.json();
    const price = Number(data?.[0]?.sell?.min || 0);
    priceCache.set(typeID, price);
    return price;
  } catch (e) {
    console.error("price lookup failed:", typeID, e);
    priceCache.set(typeID, 0);
    return 0;
  }
}

// ---------- Parse input ----------
function parseHold(text) {
  const lines = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const out = [];
  for (const line of lines) {
    const parts = line.split(/\t+|\s{2,}/);
    if (parts.length >= 2) {
      const name = parts.slice(0, -1).join(" ").trim();
      const qty = Number(stripCommas(parts[parts.length - 1]));
      if (name && qty > 0) out.push({ name, qty });
    }
  }
  return out;
}

// ---------- Generate Report ----------
document.getElementById("generate").addEventListener("click", async () => {
  const input = document.getElementById("miningHold").value.trim();
  if (!input) {
    alert("Paste your mining hold first.");
    return;
  }

  const parsed = parseHold(input);
  if (!parsed.length) {
    alert("No valid lines found. Expected: '<ItemName>\\t<Qty>' per line.");
    return;
  }

  // Quick sanity: keep only things we recognize (but allow unknowns to resolve)
  // We won’t *block* unknown names; we’ll try to resolve them anyway.
  const unresolvedByName = [];

  // Resolve in batches to be gentle on APIs
  const BATCH = 20;
  const rows = [];
  for (let i = 0; i < parsed.length; i += BATCH) {
    const chunk = parsed.slice(i, i + BATCH);

    const resolvedChunk = await Promise.all(
      chunk.map(async (row) => {
        const res = await resolveType(row.name);
        if (!res) {
          unresolvedByName.push(row);
          return null;
        }
        const cat = detectCategory(row.name);
        const price = await getPrice(res.typeID);
        return {
          rawName: row.name,
          name: res.name,        // resolved proper name
          typeID: res.typeID,
          m3each: res.volume,
          qty: row.qty,
          volume: res.volume * row.qty,
          price,
          total: price * row.qty,
          category: cat
        };
      })
    );

    rows.push(...resolvedChunk.filter(Boolean));
    // tiny pause between batches
    await sleep(120);
  }

  // Group by category
  const categories = {};
  let grandTotal = 0;
  rows.forEach(r => {
    if (!categories[r.category]) categories[r.category] = [];
    categories[r.category].push(r);
    grandTotal += r.total;
  });

  // Render
  const reportDiv = document.getElementById("report");
  reportDiv.innerHTML = "";

  for (const [category, list] of Object.entries(categories)) {
    // Sort by total descending
    list.sort((a, b) => b.total - a.total);
    const catTotal = list.reduce((s, r) => s + r.total, 0);

    const section = document.createElement("div");
    section.className = "report-section category-" + category.replace(/\s+/g, "");

    const rowsHtml = list.map(r => `
      <tr>
        <td>${r.rawName}</td>
        <td>${r.qty.toLocaleString()}</td>
        <td title="${r.m3each} m³ each">${r.volume.toLocaleString()}</td>
        <td>${fmtISK(r.price)}</td>
        <td>${fmtISK(r.total)}</td>
      </tr>
    `).join("");

    section.innerHTML = `
      <h2>${category}</h2>
      <table>
        <tr>
          <th>Item</th><th>Qty</th><th>m³</th><th>ISK (Jita)</th><th>Total ISK</th>
        </tr>
        ${rowsHtml}
        <tr class="total"><td colspan="4">Category Total</td><td>${fmtISK(catTotal)}</td></tr>
      </table>
    `;
    reportDiv.appendChild(section);
  }

  // Grand total
  const grand = document.createElement("div");
  grand.className = "report-section grand-total";
  grand.innerHTML = `<h2>Grand Total: ${fmtISK(grandTotal)} ISK</h2>`;
  reportDiv.appendChild(grand);

  // Unresolved block (if any)
  if (unresolvedByName.length) {
    const unr = document.createElement("div");
    unr.className = "report-section unresolved";
    unr.innerHTML = `
      <h2>Unresolved Items</h2>
      <p>These names didn’t resolve via Fuzzwork (check spelling or in-game names):</p>
      <ul>
        ${unresolvedByName.map(u => `<li>${u.name} (${u.qty.toLocaleString()})</li>`).join("")}
      </ul>
    `;
    reportDiv.appendChild(unr);
  }

  document.getElementById("downloadExcel").disabled = false;
});

// ---------- Excel export (one sheet per section + grand) ----------
document.getElementById("downloadExcel").addEventListener("click", () => {
  const wb = XLSX.utils.book_new();
  const now = new Date();
  const filename =
    `EVE_Mining_Report_${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}_${String(now.getHours()).padStart(2,"0")}${String(now.getMinutes()).padStart(2,"0")}.xlsx`;

  const sections = document.querySelectorAll(".report-section");
  sections.forEach(section => {
    const title = section.querySelector("h2").innerText;
    const rows = Array.from(section.querySelectorAll("tr")).map(tr =>
      Array.from(tr.querySelectorAll("th,td")).map(td => td.innerText)
    );
    if (rows.length) {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, title.substring(0, 31));
    }
  });

  XLSX.writeFile(wb, filename);
});
