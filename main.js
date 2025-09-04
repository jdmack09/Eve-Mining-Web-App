// ===== Helpers =====
const fmtISK = (n) => (isFinite(n) ? Math.round(n).toLocaleString() : "0");

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

// ===== Trade Hubs (region + station) =====
const hubs = {
  Jita:    { region: 10000002, station: 60003760 },
  Amarr:   { region: 10000043, station: 60008494 },
  Dodixie: { region: 10000032, station: 60011866 },
  Rens:    { region: 10000030, station: 60004588 },
  Hek:     { region: 10000042, station: 60005686 },
};

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
  "Gneiss": 1229,
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

// ===== Compression map (Compressed → Uncompressed) =====
const compressionMap = {
  // Veldspar
  28430: 1230, 28432: 17471, 28431: 17470,
  // Scordite
  28428: 1228, 28429: 17463, 28427: 17464,
  // Plagioclase
  28415: 18,   28417: 17459, 28416: 17460,
  // Pyroxeres
  28424: 1229, 28426: 17461, 28425: 17462,
  // Omber
  28419: 1231, 28433: 17867, 28434: 17868,
  // Kernite
  28413: 20,   28436: 17455, 28435: 17456,
  // Jaspet
  28420: 1226, 28437: 17452, 28438: 17453,
  // Hemorphite
  28421: 1234, 28439: 17444, 28440: 17445,
  // Hedbergite
  28422: 21,   28441: 17440, 28442: 17441,
  // Gneiss
  28423: 1229, 28443: 17865, 28444: 17866,
  // Dark Ochre
  28418: 1227, 28445: 17449, 28446: 17450,
  // Crokite
  28412: 1236, 28447: 17433, 28448: 17432,
  // Bistot
  28411: 1237, 28449: 17428, 28450: 17429,
  // Arkonor
  28410: 1238, 28451: 17425, 28452: 17426,
  // Spodumain
  28414: 19,   28453: 17466, 28454: 17467,
  // Mercoxit
  28409: 11396, 28455: 17869, 28456: 17870,
  // Ice
  28407: 16262, 28408: 16263, 28406: 16264, 28405: 16265,
  28403: 17978, 28402: 17976, 28404: 17977, 28401: 16266,
  // Trig ores
  62579: 62578, 62581: 62580, 62583: 62582,
  // Gas (Fullerites, Mykoserocin, Cytoserocin)
  45530: 30375, 45531: 30376, 45532: 30370, 45533: 30371,
  45534: 30372, 45535: 30373, 45536: 30374, 45537: 30378,
  45538: 30380,
  45539: 25268, 45540: 25270, 45541: 25279, 45542: 28694,
  45543: 25274, 45544: 25271, 45545: 28695, 45546: 28696,
  45547: 28697, 45548: 28698, 45549: 28699, 45550: 28701,
  45551: 28702, 45552: 28703, 45553: 28704,
  // Moon ores
  45510: 45490, 45511: 45491, 45512: 45492, 45513: 45493,
  45514: 45494, 45515: 45495, 45516: 45496, 45517: 45497,
  45518: 45498, 45519: 45499, 45520: 45500, 45521: 45501,
  45522: 45502, 45523: 45503, 45524: 45504, 45525: 45505,
  45526: 45506, 45527: 45507, 45528: 45508, 45529: 45509,
};

// ===== Resolve Type =====
async function resolveType(name) {
  const cleanName = normalizeName(name);
  if (typeCache.has(cleanName)) return typeCache.get(cleanName);

  if (typeIDs[cleanName]) {
    let typeID = typeIDs[cleanName];
    if (compressionMap[typeID]) typeID = compressionMap[typeID];
    const resolved = { typeID, name: cleanName, category: getCategory(cleanName) };
    typeCache.set(cleanName, resolved);
    return resolved;
  }

  try {
    const resp = await fetch(`https://www.fuzzwork.co.uk/api/typeid2.php?typename=${encodeURIComponent(cleanName)}`);
    const data = await resp.json();
    if (data?.typeid) {
      let typeID = data.typeid;
      if (compressionMap[typeID]) typeID = compressionMap[typeID];
      const resolved = { typeID, name: cleanName, category: getCategory(cleanName) };
      typeCache.set(cleanName, resolved);
      return resolved;
    }
  } catch (e) {
    console.error("Resolve error", cleanName, e);
  }

  typeCache.set(cleanName, null);
  return null;
}

// ===== Fetch Prices (ESI: highest buy @ hub station) =====
async function getPricesAllHubs(typeID) {
  if (!typeID) return {};
  if (priceCache.has(typeID)) return priceCache.get(typeID);

  const prices = {};

  for (const [hubName, { region, station }] of Object.entries(hubs)) {
    let page = 1;
    let bestBuy = 0;

    try {
      while (true) {
        const url = `https://esi.evetech.net/latest/markets/${region}/orders/?order_type=buy&type_id=${typeID}&datasource=tranquility&page=${page}`;
        const resp = await fetch(url);
        if (!resp.ok) break;

        const orders = await resp.json();
        if (!Array.isArray(orders) || orders.length === 0) break;

        const atStation = orders.filter(o => o.location_id === station);
        if (atStation.length) {
          const pageBest = Math.max(...atStation.map(o => o.price));
          if (pageBest > bestBuy) bestBuy = pageBest;
        }

        if (orders.length < 1000) break;
        page++;
      }
    } catch (e) {
      console.error(`ESI fetch failed for ${typeID} in ${hubName}`, e);
    }

    prices[hubName] = bestBuy || 0;
  }

  priceCache.set(typeID, prices);
  return prices;
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
  const unresolved = [];

  const grandTotals = {};
  for (const hub of Object.keys(hubs)) grandTotals[hub] = 0;

  for (const row of parsed) {
    const resolved = await resolveType(row.name);
    if (!resolved) {
      unresolved.push(row);
      continue;
    }

    const prices = await getPricesAllHubs(resolved.typeID);
    const totals = {};
    let bestHub = null;
    let bestVal = 0; // highest ISK is best

    for (const hub of Object.keys(hubs)) {
      const perItem = prices[hub] || 0;
      totals[hub] = perItem * row.qty;
      grandTotals[hub] += totals[hub];

      if (totals[hub] > bestVal) {
        bestVal = totals[hub];
        bestHub = hub;
      }
    }

    if (!buckets[resolved.category]) buckets[resolved.category] = [];
    buckets[resolved.category].push({ name: resolved.name, qty: row.qty, prices, totals, bestHub });
  }

  // Render per-category tables
  const catOrder = ["Asteroid", "Moon", "Triglavian", "Gas", "Ice", "Other"];
  for (const cat of catOrder) {
    const rows = buckets[cat];
    if (!rows) continue;

    const section = document.createElement("div");
    section.className = "report-section";
    section.innerHTML = `
      <h2>${cat}</h2>
      <table>
        <tr>
          <th>Ore</th><th>Qty</th>
          ${Object.keys(hubs).map(h => `<th>${h} (Buy / Total)</th>`).join("")}
          <th>Best Hub</th>
        </tr>
        ${rows.map(r => `
          <tr>
            <td>${r.name}</td>
            <td>${r.qty.toLocaleString()}</td>
            ${Object.keys(hubs).map(h => `
              <td class="${h === r.bestHub ? 'best-hub' : ''}">
                <div class="price-each">${fmtISK(r.prices[h])} ISK</div>
                <div class="price-total">${fmtISK(r.totals[h])} ISK</div>
              </td>`).join("")}
            <td class="best-hub"><strong>${r.bestHub || "-"}</strong></td>
          </tr>`).join("")}
      </table>
    `;
    reportDiv.appendChild(section);
  }

  // Grand Totals
  const bestHubOverall = Object.entries(grandTotals).reduce(
    (a, b) => (b[1] > a[1] ? b : a)
  )[0];

  const grand = document.createElement("div");
  grand.className = "grand-total";
  grand.innerHTML = `
    <h2>Grand Totals</h2>
    <table>
      <tr>${Object.keys(hubs).map(h => `<th>${h}</th>`).join("")}<th>Best Hub</th></tr>
      <tr>
        ${Object.keys(hubs).map(h => `<td class="${h === bestHubOverall ? 'best-hub' : ''}">${fmtISK(grandTotals[h])} ISK</td>`).join("")}
        <td class="best-hub"><strong>${bestHubOverall}</strong></td>
      </tr>
    </table>
  `;
  reportDiv.appendChild(grand);

  // Unresolved items
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

// ===== Clear Button =====
    document.getElementById("clearReport").addEventListener("click", () => {
      document.getElementById("miningHold").value = "";
      document.getElementById("report").innerHTML = "";
      document.getElementById("downloadExcel").disabled = true;
    });

    // ===== Download Excel (with per-category sheets + Full Report + Grand Totals) =====
document.getElementById("downloadExcel").addEventListener("click", () => {
  const reportDiv = document.getElementById("report");
  if (!reportDiv.innerHTML.trim()) return;

  const wb = XLSX.utils.book_new();
  const now = new Date();
  const timestamp = now.toISOString().replace("T", " ").slice(0, 16);
  const bestHubCell = reportDiv.querySelector(".grand-total .best-hub strong");
  const bestHub = bestHubCell ? bestHubCell.textContent.trim() : "Unknown";

  // 🎨 Theme colors
  const theme = {
    baseBg: "1A1A1A",
    altBg: "111111",
    border: "FF8800",
    headerBg: "000000",
    headerText: "FFCC66",
    text: "FFCC66",
    totalBg: "222222",
    bestHubBg: "FFC800",
    bestHubText: "000000"
  };

  const categoryColors = {
    Asteroid: { text: "FFC266", tab: "FF9933" },
    Moon:     { text: "FF99CC", tab: "FF6699" },
    Triglavian: { text: "FFFF99", tab: "FFCC00" },
    Gas:      { text: "99FF99", tab: "33CC33" },
    Ice:      { text: "99CCFF", tab: "3399FF" },
    Other:    { text: "DDDDDD", tab: "AAAAAA" }
  };

  const sections = reportDiv.querySelectorAll(".report-section");
  const fullReportRows = [];

  // ===== Category Sheets =====
  sections.forEach(section => {
    const title = section.querySelector("h2")?.textContent.trim() || "Sheet";
    const table = section.querySelector("table");
    if (!table) return;

    const cat = Object.keys(categoryColors).find(c => title.includes(c)) || "Other";
    const catColor = categoryColors[cat].text;
    const tabColor = categoryColors[cat].tab;

    const rows = [];
    table.querySelectorAll("tr").forEach((tr, rIdx) => {
      const cells = [];
      tr.querySelectorAll("th,td").forEach(td => {
        let text = td.innerText.trim();
        const clean = text.replace(/[,\s]/g, "");
        let cell = { v: text };

        if (/^-?\d+(\.\d+)?ISK$/.test(text.replace(/\s+/g, ""))) {
          const num = Number(clean.replace(/ISK/i, ""));
          cell = { v: num, t: "n", z: '#,##0" ISK"' };
        } else if (/^-?\d+(\.\d+)?$/.test(clean)) {
          cell = { v: Number(clean), t: "n" };
        }

        if (tr.querySelector("th")) {
          cell.s = { fill: { fgColor: { rgb: theme.headerBg } },
                     font: { bold: true, color: { rgb: theme.headerText } },
                     alignment: { horizontal: "center" } };
        } else if (td.classList.contains("best-hub")) {
          cell.s = { fill: { fgColor: { rgb: theme.bestHubBg } },
                     font: { bold: true, color: { rgb: theme.bestHubText } },
                     alignment: { horizontal: "center" } };
        } else if (rIdx % 2 === 1) {
          cell.s = { fill: { fgColor: { rgb: theme.altBg } },
                     font: { color: { rgb: theme.text } } };
        } else {
          cell.s = { fill: { fgColor: { rgb: theme.baseBg } },
                     font: { color: { rgb: theme.text } } };
        }

        cells.push(cell);
      });
      rows.push(cells);
    });

    // Insert header metadata
    const header = [
      [{ v: "Ore Report Export", s: { font: { bold: true, sz: 14, color: { rgb: catColor } }, alignment: { horizontal: "center" } } }],
      [{ v: `Section: ${title}`, s: { font: { bold: true, color: { rgb: catColor } }, alignment: { horizontal: "center" } } }],
      [{ v: `Best Hub: ${bestHub}`, s: { alignment: { horizontal: "center" }, font: { color: { rgb: theme.text } } } }],
      [{ v: `Generated: ${timestamp}`, s: { alignment: { horizontal: "center" }, font: { color: { rgb: theme.text } } } }],
      []
    ];

    const ws = XLSX.utils.aoa_to_sheet([...header, ...rows]);
    const colCount = rows[0] ? rows[0].length : 1;

    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: colCount - 1 } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: colCount - 1 } },
    ];
    ws['!freeze'] = { xSplit: 0, ySplit: 5, topLeftCell: "A6", activePane: "bottomLeft", state: "frozen" };
    ws['!cols'] = Array(colCount).fill({ wch: 18 });
    ws['!tabColor'] = { rgb: tabColor };

    XLSX.utils.book_append_sheet(wb, ws, title.substring(0, 31));

    // Save rows for Full Report
    fullReportRows.push([{ v: title, s: { font: { bold: true, sz: 14, color: { rgb: catColor } } } }]);
    rows.forEach(r => fullReportRows.push(r));
    fullReportRows.push([]); // spacer
  });

  // ===== Full Report Sheet =====
  if (fullReportRows.length > 0) {
    const styledFullReportRows = [];

    // Insert category-colored headers
    sections.forEach(section => {
      const title = section.querySelector("h2")?.textContent.trim() || "Sheet";
      const table = section.querySelector("table");
      if (!table) return;

      const cat = Object.keys(categoryColors).find(c => title.includes(c)) || "Other";
      const catColor = categoryColors[cat].text;

      // Category header row
      styledFullReportRows.push([
        { v: title, s: { font: { bold: true, sz: 14, color: { rgb: catColor } }, alignment: { horizontal: "left" } } }
      ]);

      // Table rows
      table.querySelectorAll("tr").forEach((tr, rIdx) => {
        const cells = [];
        tr.querySelectorAll("th,td").forEach(td => {
          let text = td.innerText.trim();
          const clean = text.replace(/[,\s]/g, "");
          let cell = { v: text };

          if (/^-?\d+(\.\d+)?ISK$/.test(text.replace(/\s+/g, ""))) {
            const num = Number(clean.replace(/ISK/i, ""));
            cell = { v: num, t: "n", z: '#,##0" ISK"' };
          } else if (/^-?\d+(\.\d+)?$/.test(clean)) {
            cell = { v: Number(clean), t: "n" };
          }

          if (tr.querySelector("th")) {
            cell.s = { fill: { fgColor: { rgb: theme.headerBg } },
                       font: { bold: true, color: { rgb: theme.headerText } },
                       alignment: { horizontal: "center" } };
          } else if (td.classList.contains("best-hub")) {
            cell.s = { fill: { fgColor: { rgb: theme.bestHubBg } },
                       font: { bold: true, color: { rgb: theme.bestHubText } },
                       alignment: { horizontal: "center" } };
          } else if (rIdx % 2 === 1) {
            cell.s = { fill: { fgColor: { rgb: theme.altBg } },
                       font: { color: { rgb: theme.text } } };
          } else {
            cell.s = { fill: { fgColor: { rgb: theme.baseBg } },
                       font: { color: { rgb: theme.text } } };
          }

          cells.push(cell);
        });
        styledFullReportRows.push(cells);
      });

      styledFullReportRows.push([]); // spacer
    });

    const ws = XLSX.utils.aoa_to_sheet([
      [{ v: "Full Ore Report", s: { font: { bold: true, sz: 16, color: { rgb: theme.headerText } }, alignment: { horizontal: "center" } } }],
      [{ v: `Best Hub Overall: ${bestHub}`, s: { font: { bold: true, color: { rgb: theme.headerText } }, alignment: { horizontal: "center" } } }],
      [{ v: `Generated: ${timestamp}`, s: { font: { color: { rgb: theme.text } }, alignment: { horizontal: "center" } } }],
      [],
      ...styledFullReportRows
    ]);

    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 6 } }
    ];
    ws['!tabColor'] = { rgb: "FF8800" }; // orange tab
    XLSX.utils.book_append_sheet(wb, ws, "Full Report");
  }

  // ===== Grand Totals Sheet =====
  const grandSection = reportDiv.querySelector(".grand-total table");
  if (grandSection) {
    const rows = [];
    grandSection.querySelectorAll("tr").forEach(tr => {
      const cells = [];
      tr.querySelectorAll("th,td").forEach(td => {
        let text = td.innerText.trim();
        const clean = text.replace(/[,\s]/g, "");
        let cell = { v: text };

        if (/^-?\d+(\.\d+)?ISK$/.test(text.replace(/\s+/g, ""))) {
          const num = Number(clean.replace(/ISK/i, ""));
          cell = { v: num, t: "n", z: '#,##0" ISK"' };
        }

        if (tr.querySelector("th")) {
          cell.s = { fill: { fgColor: { rgb: theme.headerBg } },
                     font: { bold: true, color: { rgb: theme.headerText } },
                     alignment: { horizontal: "center" } };
        } else if (td.classList.contains("best-hub")) {
          cell.s = { fill: { fgColor: { rgb: theme.bestHubBg } },
                     font: { bold: true, color: { rgb: theme.bestHubText } },
                     alignment: { horizontal: "center" } };
        }

        cells.push(cell);
      });
      rows.push(cells);
    });

    const ws = XLSX.utils.aoa_to_sheet([
      [{ v: "Ore Report - Grand Totals", s: { font: { bold: true, sz: 16, color: { rgb: theme.headerText } }, alignment: { horizontal: "center" } } }],
      [{ v: `Best Hub Overall: ${bestHub}`, s: { font: { bold: true, color: { rgb: theme.headerText } }, alignment: { horizontal: "center" } } }],
      [{ v: `Generated: ${timestamp}`, s: { font: { color: { rgb: theme.text } }, alignment: { horizontal: "center" } } }],
      [],
      ...rows
    ]);

    ws['!tabColor'] = { rgb: "FFCC00" }; // gold tab
    XLSX.utils.book_append_sheet(wb, ws, "Grand Totals");
  }

  const safeTimestamp = now.toISOString().replace("T", "_").slice(0, 16).replace(/:/g, "-");
  const filename = `OreReport_${bestHub}_${safeTimestamp}.xlsx`;
  XLSX.writeFile(wb, filename);
});
