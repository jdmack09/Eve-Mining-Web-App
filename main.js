// ===== Helpers =====
const fmtISK = (n) => (isFinite(n) ? Math.round(n).toLocaleString() : "0");
const fmtVOL = (n) => (isFinite(n) ? n.toLocaleString() : "0");

const typeCache = new Map();
const priceCache = new Map();

// ===== Category Detection =====
function getCategory(name) {
  if (/Bezdnacine|Rakovene|Talassonite/i.test(name)) return "Triglavian";
  if (/Veldspar|Scordite|Plagioclase|Pyroxeres|Omber|Kernite|Jaspet|Hemorphite|Hedbergite|Gneiss|Ochre|Crokite|Bistot|Arkonor|Spodumain|Mercoxit/i.test(name)) return "Asteroid";
  if (/Bitumens|Coesite|Sylvite|Zeolites|Cobaltite|Euxenite|Titanite|Scheelite|Otavite|Sperrylite|Vanadinite|Chromite|Carnotite|Zircon|Pollucite|Cinnabar|Xenotime|Monazite|Loparite|Ytterbite/i.test(name)) return "Moon";
  if (/Fullerite|Mykoserocin|Cytoserocin/i.test(name)) return "Gas";
  if (/Ice|Icicle|Glacial|Glare|Glitter|Gelidus|Krystallos|Glaze/i.test(name)) return "Ice";
  return "Other";
}

// ===== Normalize name (for compressed versions) =====
function normalizeName(name) {
  return String(name || "").replace(/^Compressed\s+/i, "").trim();
}

// ===== Alias corrections (all compressed ores/ice/gas) =====
const nameAliases = {
  // --- Compressed Veldspar Variants ---
  "Compressed Veldspar": "Compressed Veldspar",
  "Compressed Concentrated Veldspar": "Compressed Concentrated Veldspar",
  "Compressed Dense Veldspar": "Compressed Dense Veldspar",

  // --- Compressed Scordite Variants ---
  "Compressed Scordite": "Compressed Scordite",
  "Compressed Condensed Scordite": "Compressed Condensed Scordite",
  "Compressed Massive Scordite": "Compressed Massive Scordite",

  // --- Compressed Plagioclase Variants ---
  "Compressed Plagioclase": "Compressed Plagioclase",
  "Compressed Azure Plagioclase": "Compressed Azure Plagioclase",
  "Compressed Rich Plagioclase": "Compressed Rich Plagioclase",

  // --- Compressed Pyroxeres Variants ---
  "Compressed Pyroxeres": "Compressed Pyroxeres",
  "Compressed Solid Pyroxeres": "Compressed Solid Pyroxeres",
  "Compressed Viscous Pyroxeres": "Compressed Viscous Pyroxeres",

  // --- Compressed Omber Variants ---
  "Compressed Omber": "Compressed Omber",
  "Compressed Golden Omber": "Compressed Golden Omber",
  "Compressed Silvery Omber": "Compressed Silvery Omber",

  // --- Compressed Kernite Variants ---
  "Compressed Kernite": "Compressed Kernite",
  "Compressed Fiery Kernite": "Compressed Fiery Kernite",
  "Compressed Luminous Kernite": "Compressed Luminous Kernite",

  // --- Compressed Jaspet Variants ---
  "Compressed Jaspet": "Compressed Jaspet",
  "Compressed Pure Jaspet": "Compressed Pure Jaspet",
  "Compressed Pristine Jaspet": "Compressed Pristine Jaspet",

  // --- Compressed Hemorphite Variants ---
  "Compressed Hemorphite": "Compressed Hemorphite",
  "Compressed Vivid Hemorphite": "Compressed Vivid Hemorphite",
  "Compressed Radiant Hemorphite": "Compressed Radiant Hemorphite",

  // --- Compressed Hedbergite Variants ---
  "Compressed Hedbergite": "Compressed Hedbergite",
  "Compressed Glazed Hedbergite": "Compressed Glazed Hedbergite",
  "Compressed Vitric Hedbergite": "Compressed Vitric Hedbergite",

  // --- Compressed Gneiss Variants ---
  "Compressed Gneiss": "Compressed Gneiss",
  "Compressed Iridescent Gneiss": "Compressed Iridescent Gneiss",
  "Compressed Prismatic Gneiss": "Compressed Prismatic Gneiss",

  // --- Compressed Dark Ochre Variants ---
  "Compressed Dark Ochre": "Compressed Dark Ochre",
  "Compressed Obsidian Ochre": "Compressed Obsidian Ochre",
  "Compressed Onyx Ochre": "Compressed Onyx Ochre",

  // --- Compressed Spodumain Variants ---
  "Compressed Spodumain": "Compressed Spodumain",
  "Compressed Bright Spodumain": "Compressed Bright Spodumain",
  "Compressed Gleaming Spodumain": "Compressed Gleaming Spodumain",

  // --- Compressed Crokite Variants ---
  "Compressed Crokite": "Compressed Crokite",
  "Compressed Sharp Crokite": "Compressed Sharp Crokite",
  "Compressed Crystalline Crokite": "Compressed Crystalline Crokite",

  // --- Compressed Bistot Variants ---
  "Compressed Bistot": "Compressed Bistot",
  "Compressed Triclinic Bistot": "Compressed Triclinic Bistot",
  "Compressed Monoclinic Bistot": "Compressed Monoclinic Bistot",

  // --- Compressed Arkonor Variants ---
  "Compressed Arkonor": "Compressed Arkonor",
  "Compressed Crimson Arkonor": "Compressed Crimson Arkonor",
  "Compressed Prime Arkonor": "Compressed Prime Arkonor",

  // --- Compressed Mercoxit Variants ---
  "Compressed Mercoxit": "Compressed Mercoxit",
  "Compressed Magma Mercoxit": "Compressed Magma Mercoxit",
  "Compressed Vitreous Mercoxit": "Compressed Vitreous Mercoxit",

  // --- Compressed Ice Variants ---
  "Compressed Blue Ice": "Compressed Blue Ice",
  "Compressed Clear Icicle": "Compressed Clear Icicle",
  "Compressed Glacial Mass": "Compressed Glacial Mass",
  "Compressed White Glaze": "Compressed White Glaze",
  "Compressed Glare Crust": "Compressed Glare Crust",
  "Compressed Dark Glitter": "Compressed Dark Glitter",
  "Compressed Gelidus": "Compressed Gelidus",
  "Compressed Krystallos": "Compressed Krystallos",

  // --- Compressed Gas (Fullerite) ---
  "Compressed Fullerite-C28": "Compressed Fullerite-C28",
  "Compressed Fullerite-C32": "Compressed Fullerite-C32",
  "Compressed Fullerite-C50": "Compressed Fullerite-C50",
  "Compressed Fullerite-C60": "Compressed Fullerite-C60",
  "Compressed Fullerite-C70": "Compressed Fullerite-C70",
  "Compressed Fullerite-C72": "Compressed Fullerite-C72",
  "Compressed Fullerite-C84": "Compressed Fullerite-C84",
  "Compressed Fullerite-C320": "Compressed Fullerite-C320",
  "Compressed Fullerite-C540": "Compressed Fullerite-C540",
};

// ===== Resolve Type (via Fuzzwork with alias fix) =====
async function resolveType(name) {
  const cleanName = name.replace(/\s+/g, " ").trim();
  const queryName = nameAliases[cleanName] || cleanName;

  if (typeCache.has(queryName)) return typeCache.get(queryName);

  try {
    const resp = await fetch(
      `https://www.fuzzwork.co.uk/api/typeid2.php?typename=${encodeURIComponent(queryName)}`
    );
    const data = await resp.json();

    if (data && data.typeid) {
      const resolved = {
        typeID: data.typeid,
        volume: Number(data.volume) || 0,
        name: cleanName, // preserve original
        category: getCategory(normalizeName(cleanName)),
      };
      typeCache.set(queryName, resolved);
      return resolved;
    }
  } catch (e) {
    console.error("Resolve error", queryName, e);
  }

  typeCache.set(queryName, null);
  return null;
}

// ===== Fetch Average Price from Fuzzwork Aggregates =====
async function getPrice(typeID) {
  if (!typeID) return 0;
  if (priceCache.has(typeID)) return priceCache.get(typeID);

  const regionID = document.getElementById("hubSelect")?.value || "10000002";
  let price = 0;

  try {
    const url = `https://market.fuzzwork.co.uk/aggregates/?region=${regionID}&types=${typeID}`;
    const resp = await fetch(url);
    const data = await resp.json();

    price = Number(data?.[typeID]?.sell?.avg) || 0;
    priceCache.set(typeID, price);
    return price;
  } catch (e) {
    console.error("Price fetch failed", typeID, e);
    return 0;
  }
}

// ===== Generate Report =====
document.getElementById("generate").addEventListener("click", async () => {
  const input = document.getElementById("miningHold").value.trim();
  if (!input) return alert("Paste your mining hold first.");

  const lines = input.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const parsed = [];

  for (const line of lines) {
    const parts = line.split(/\t+|\s{2,}/);
    if (parts.length >= 2) {
      const name = parts.slice(0, -1).join(" ").trim();
      const qty = Number(parts[parts.length - 1].replace(/[,\s]/g, ""));
      if (name && qty > 0) parsed.push({ name, qty });
    }
  }
  if (!parsed.length) return alert("No valid lines found.");

  const reportDiv = document.getElementById("report");
  reportDiv.innerHTML = "";
  const buckets = {};
  let grandTotal = 0;
  const unresolved = [];

  for (const row of parsed) {
    const resolved = await resolveType(row.name);
    if (!resolved) {
      unresolved.push(row);
      continue;
    }

    const price = await getPrice(resolved.typeID);
    const total = price * row.qty;
    const vol = (resolved.volume || 0) * row.qty;

    if (!buckets[resolved.category]) buckets[resolved.category] = [];
    buckets[resolved.category].push({
      name: resolved.name,
      qty: row.qty,
      price,
      total,
      volume: vol,
    });

    grandTotal += total;
  }

  // Render per-category tables
  const catOrder = ["Asteroid", "Moon", "Triglavian", "Gas", "Ice", "Other"];
  for (const cat of catOrder) {
    const rows = buckets[cat];
    if (!rows) continue;

    rows.sort((a, b) => b.total - a.total);
    const catTotal = rows.reduce((s, r) => s + r.total, 0);

    const section = document.createElement("div");
    section.className = "report-section";
    section.innerHTML = `
      <h2>${cat}</h2>
      <table>
        <tr><th>Ore</th><th>Qty</th><th>m³</th><th>ISK (avg sell)</th><th>Total ISK</th></tr>
        ${rows.map(r => `
          <tr>
            <td>${r.name}</td>
            <td>${r.qty.toLocaleString()}</td>
            <td>${fmtVOL(r.volume)}</td>
            <td>${fmtISK(r.price)}</td>
            <td>${fmtISK(r.total)}</td>
          </tr>`).join("")}
        <tr class="total"><td colspan="4">Category Total</td><td>${fmtISK(catTotal)}</td></tr>
      </table>
    `;
    reportDiv.appendChild(section);
  }

  // Grand total
  const grand = document.createElement("div");
  grand.className = "grand-total";
  grand.innerHTML = `<h2>Grand Total: ${fmtISK(grandTotal)} ISK</h2>`;
  reportDiv.appendChild(grand);

  // Unresolved items
  if (unresolved.length) {
    const unr = document.createElement("div");
    unr.className = "report-section";
    unr.innerHTML = `
      <h2>Unresolved Items</h2>
      <ul>${unresolved.map(u => `<li>${u.name} (${u.qty})</li>`).join("")}</ul>
    `;
    reportDiv.appendChild(unr);
  }

  document.getElementById("downloadExcel").disabled = false;
});

// ===== Excel Export =====
document.getElementById("downloadExcel").addEventListener("click", () => {
  const wb = XLSX.utils.book_new();
  const now = new Date();
  const hubName = document.getElementById("hubSelect")?.selectedOptions[0].text || "Jita";

  const timestamp = `${now.toISOString().slice(0,10)}_${String(now.getHours()).padStart(2,"0")}${String(now.getMinutes()).padStart(2,"0")}`;
  const filename = `EVE_Mining_Report_${hubName}_${timestamp}.xlsx`;

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
