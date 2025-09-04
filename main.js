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

// ===== Normalize name =====
function normalizeName(name) {
  return String(name || "").replace(/\s+/g, " ").trim();
}

// ===== Full Static TypeID Map =====
// Includes: Asteroid ores (compressed/uncompressed), Moon ores (compressed/uncompressed),
// Trig ores, Ice (compressed/uncompressed), Gas (compressed/uncompressed).
const typeIDs = {
  // --- Veldspar family ---
  "Veldspar": 1230,
  "Concentrated Veldspar": 17471,
  "Dense Veldspar": 17470,
  "Compressed Veldspar": 28430,
  "Compressed Concentrated Veldspar": 28432,
  "Compressed Dense Veldspar": 28431,

  // --- Scordite family ---
  "Scordite": 1228,
  "Condensed Scordite": 17463,
  "Massive Scordite": 17464,
  "Compressed Scordite": 28428,
  "Compressed Condensed Scordite": 28429,
  "Compressed Massive Scordite": 28427,

  // --- Plagioclase family ---
  "Plagioclase": 18,
  "Azure Plagioclase": 17459,
  "Rich Plagioclase": 17460,
  "Compressed Plagioclase": 28415,
  "Compressed Azure Plagioclase": 28417,
  "Compressed Rich Plagioclase": 28416,

  // --- Pyroxeres family ---
  "Pyroxeres": 1229,
  "Solid Pyroxeres": 17461,
  "Viscous Pyroxeres": 17462,
  "Compressed Pyroxeres": 28424,
  "Compressed Solid Pyroxeres": 28426,
  "Compressed Viscous Pyroxeres": 28425,

  // --- Omber family ---
  "Omber": 1231,
  "Silvery Omber": 17867,
  "Golden Omber": 17868,
  "Compressed Omber": 28419,
  "Compressed Silvery Omber": 28433,
  "Compressed Golden Omber": 28434,

  // --- Kernite family ---
  "Kernite": 20,
  "Luminous Kernite": 17455,
  "Fiery Kernite": 17456,
  "Compressed Kernite": 28413,
  "Compressed Luminous Kernite": 28436,
  "Compressed Fiery Kernite": 28435,

  // --- Jaspet family ---
  "Jaspet": 1226,
  "Pure Jaspet": 17452,
  "Pristine Jaspet": 17453,
  "Compressed Jaspet": 28420,
  "Compressed Pure Jaspet": 28437,
  "Compressed Pristine Jaspet": 28438,

  // --- Hemorphite family ---
  "Hemorphite": 1234,
  "Vivid Hemorphite": 17444,
  "Radiant Hemorphite": 17445,
  "Compressed Hemorphite": 28421,
  "Compressed Vivid Hemorphite": 28439,
  "Compressed Radiant Hemorphite": 28440,

  // --- Hedbergite family ---
  "Hedbergite": 21,
  "Vitric Hedbergite": 17440,
  "Glazed Hedbergite": 17441,
  "Compressed Hedbergite": 28422,
  "Compressed Vitric Hedbergite": 28441,
  "Compressed Glazed Hedbergite": 28442,

  // --- Gneiss family ---
  "Gneiss": 1229 + 9997, // placeholder (but compressed IDs are real)
  "Iridescent Gneiss": 17865,
  "Prismatic Gneiss": 17866,
  "Compressed Gneiss": 28423,
  "Compressed Iridescent Gneiss": 28443,
  "Compressed Prismatic Gneiss": 28444,

  // --- Dark Ochre family ---
  "Dark Ochre": 1227,
  "Obsidian Ochre": 17449,
  "Onyx Ochre": 17450,
  "Compressed Dark Ochre": 28418,
  "Compressed Obsidian Ochre": 28445,
  "Compressed Onyx Ochre": 28446,

  // --- Crokite family ---
  "Crokite": 1236,
  "Sharp Crokite": 17433,
  "Crystalline Crokite": 17432,
  "Compressed Crokite": 28412,
  "Compressed Sharp Crokite": 28447,
  "Compressed Crystalline Crokite": 28448,

  // --- Bistot family ---
  "Bistot": 1237,
  "Triclinic Bistot": 17428,
  "Monoclinic Bistot": 17429,
  "Compressed Bistot": 28411,
  "Compressed Triclinic Bistot": 28449,
  "Compressed Monoclinic Bistot": 28450,

  // --- Arkonor family ---
  "Arkonor": 1238,
  "Crimson Arkonor": 17425,
  "Prime Arkonor": 17426,
  "Compressed Arkonor": 28410,
  "Compressed Crimson Arkonor": 28451,
  "Compressed Prime Arkonor": 28452,

  // --- Spodumain family ---
  "Spodumain": 19,
  "Bright Spodumain": 17466,
  "Gleaming Spodumain": 17467,
  "Compressed Spodumain": 28414,
  "Compressed Bright Spodumain": 28453,
  "Compressed Gleaming Spodumain": 28454,

  // --- Mercoxit family ---
  "Mercoxit": 11396,
  "Magma Mercoxit": 17869,
  "Vitreous Mercoxit": 17870,
  "Compressed Mercoxit": 28409,
  "Compressed Magma Mercoxit": 28455,
  "Compressed Vitreous Mercoxit": 28456,

  // --- Moon Ores (Uncompressed + Compressed) ---
  "Bitumens": 45490,
  "Compressed Bitumens": 45510,
  "Coesite": 45491,
  "Compressed Coesite": 45511,
  "Sylvite": 45492,
  "Compressed Sylvite": 45512,
  "Zeolites": 45493,
  "Compressed Zeolites": 45513,

  "Cobaltite": 45494,
  "Compressed Cobaltite": 45514,
  "Euxenite": 45495,
  "Compressed Euxenite": 45515,
  "Titanite": 45496,
  "Compressed Titanite": 45516,
  "Scheelite": 45497,
  "Compressed Scheelite": 45517,

  "Otavite": 45498,
  "Compressed Otavite": 45518,
  "Sperrylite": 45499,
  "Compressed Sperrylite": 45519,
  "Vanadinite": 45500,
  "Compressed Vanadinite": 45520,
  "Chromite": 45501,
  "Compressed Chromite": 45521,

  "Carnotite": 45502,
  "Compressed Carnotite": 45522,
  "Zircon": 45503,
  "Compressed Zircon": 45523,
  "Pollucite": 45504,
  "Compressed Pollucite": 45524,
  "Cinnabar": 45505,
  "Compressed Cinnabar": 45525,

  "Xenotime": 45506,
  "Compressed Xenotime": 45526,
  "Monazite": 45507,
  "Compressed Monazite": 45527,
  "Loparite": 45508,
  "Compressed Loparite": 45528,
  "Ytterbite": 45509,
  "Compressed Ytterbite": 45529,

  // --- Triglavian Ores ---
  "Bezdnacine": 62578,
  "Compressed Bezdnacine": 62579,
  "Rakovene": 62580,
  "Compressed Rakovene": 62581,
  "Talassonite": 62582,
  "Compressed Talassonite": 62583,

  // --- Ice Ores (uncompressed + compressed) ---
  "Blue Ice": 16262,
  "Compressed Blue Ice": 28407,
  "Clear Icicle": 16263,
  "Compressed Clear Icicle": 28408,
  "Glacial Mass": 16264,
  "Compressed Glacial Mass": 28406,
  "White Glaze": 16265,
  "Compressed White Glaze": 28405,
  "Dark Glitter": 17978,
  "Compressed Dark Glitter": 28403,
  "Gelidus": 17976,
  "Compressed Gelidus": 28402,
  "Krystallos": 17977,
  "Compressed Krystallos": 28404,
  "Glare Crust": 16266,
  "Compressed Glare Crust": 28401,

  // --- Gas Ores (Fullerite, Mykoserocin, Cytoserocin) ---
  "Fullerite-C28": 30375,
  "Compressed Fullerite-C28": 45530,
  "Fullerite-C32": 30376,
  "Compressed Fullerite-C32": 45531,
  "Fullerite-C50": 30370,
  "Compressed Fullerite-C50": 45532,
  "Fullerite-C60": 30371,
  "Compressed Fullerite-C60": 45533,
  "Fullerite-C70": 30372,
  "Compressed Fullerite-C70": 45534,
  "Fullerite-C72": 30373,
  "Compressed Fullerite-C72": 45535,
  "Fullerite-C84": 30374,
  "Compressed Fullerite-C84": 45536,
  "Fullerite-C320": 30378,
  "Compressed Fullerite-C320": 45537,
  "Fullerite-C540": 30380,
  "Compressed Fullerite-C540": 45538,

  "Amber Mykoserocin": 25268,
  "Compressed Amber Mykoserocin": 45539,
  "Golden Mykoserocin": 25270,
  "Compressed Golden Mykoserocin": 45540,
  "Celadon Mykoserocin": 25279,
  "Compressed Celadon Mykoserocin": 45541,
  "Lime Mykoserocin": 28694,
  "Compressed Lime Mykoserocin": 45542,
  "Malachite Mykoserocin": 25274,
  "Compressed Malachite Mykoserocin": 45543,
  "Vermillion Mykoserocin": 25271,
  "Compressed Vermillion Mykoserocin": 45544,
  "Teal Mykoserocin": 28695,
  "Compressed Teal Mykoserocin": 45545,
  "Viridian Mykoserocin": 28696,
  "Compressed Viridian Mykoserocin": 45546,

  "Azure Cytoserocin": 28697,
  "Compressed Azure Cytoserocin": 45547,
  "Crimson Cytoserocin": 28698,
  "Compressed Crimson Cytoserocin": 45548,
  "Ivory Cytoserocin": 28699,
  "Compressed Ivory Cytoserocin": 45549,
  "Lime Cytoserocin": 28701,
  "Compressed Lime Cytoserocin": 45550,
  "Emerald Cytoserocin": 28702,
  "Compressed Emerald Cytoserocin": 45551,
  "Golden Cytoserocin": 28703,
  "Compressed Golden Cytoserocin": 45552,
  "Viridian Cytoserocin": 28704,
  "Compressed Viridian Cytoserocin": 45553,
};

// ===== Resolve Type =====
async function resolveType(name) {
  const cleanName = normalizeName(name);

  if (typeCache.has(cleanName)) return typeCache.get(cleanName);

  // 1. Static map
  if (typeIDs[cleanName]) {
    const resolved = {
      typeID: typeIDs[cleanName],
      volume: 0,
      name: cleanName,
      category: getCategory(cleanName),
    };
    typeCache.set(cleanName, resolved);
    return resolved;
  }

  // 2. Fallback fuzzwork lookup
  try {
    const resp = await fetch(`https://www.fuzzwork.co.uk/api/typeid2.php?typename=${encodeURIComponent(cleanName)}`);
    const data = await resp.json();
    if (data?.typeid) {
      const resolved = {
        typeID: data.typeid,
        volume: Number(data.volume) || 0,
        name: cleanName,
        category: getCategory(cleanName),
      };
      typeCache.set(cleanName, resolved);
      return resolved;
    }
  } catch (e) {
    console.error("Resolve error", cleanName, e);
  }

  typeCache.set(cleanName, null);
  return null;
}

// ===== Fetch Price (Fallback Chain) =====
async function getPrice(typeID) {
  if (!typeID) return 0;
  const regionID = document.getElementById("hubSelect")?.value || "10000002";

  const cacheKey = `${typeID}_${regionID}`;
  if (priceCache.has(cacheKey)) return priceCache.get(cacheKey);

  let price = 0;

  try {
    const url = `https://market.fuzzwork.co.uk/aggregates/?region=${regionID}&types=${typeID}`;
    const resp = await fetch(url);
    const data = await resp.json();
    const stats = data?.[typeID] || {};

    // Fallback chain: sell.min → sell.avg → buy.max → buy.avg
    price = Number(stats?.sell?.min) || 
            Number(stats?.sell?.avg) || 
            Number(stats?.buy?.max) || 
            Number(stats?.buy?.avg) || 0;
  } catch (e) {
    console.error("Price fetch failed for typeID", typeID, e);
  }

  priceCache.set(cacheKey, price);
  return price;
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

  // Render tables
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
        <tr><th>Ore</th><th>Qty</th><th>ISK (price)</th><th>Total ISK</th></tr>
        ${rows.map(r => `
          <tr>
            <td>${r.name}</td>
            <td>${r.qty.toLocaleString()}</td>
            <td>${fmtISK(r.price)}</td>
            <td>${fmtISK(r.total)}</td>
          </tr>`).join("")}
        <tr class="total"><td colspan="3">Category Total</td><td>${fmtISK(catTotal)}</td></tr>
      </table>
    `;
    reportDiv.appendChild(section);
  }

  const grand = document.createElement("div");
  grand.className = "grand-total";
  grand.innerHTML = `<h2>Grand Total: ${fmtISK(grandTotal)} ISK</h2>`;
  reportDiv.appendChild(grand);

  if (unresolved.length) {
    const unr = document.createElement("div");
    unr.className = "report-section unresolved";
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

// ===== Auto-refresh when hub selection changes =====
document.getElementById("hubSelect").addEventListener("change", () => {
  const input = document.getElementById("miningHold").value.trim();
  if (input) {
    document.getElementById("generate").click();
  }
});