import oreData from "./oreData.js";

// --- Category Detection ---
function getCategory(name) {
  if (/Bezdnacine|Rakovene|Talassonite/i.test(name)) return "Triglavian";
  if (/Veldspar|Scordite|Plagioclase|Pyroxeres|Omber|Kernite|Jaspet|Hemorphite|Hedbergite|Gneiss|Ochre|Crokite|Bistot|Arkonor/i.test(name)) return "Asteroid";
  if (/Bitumens|Cobaltite|Sylvite|Zeolites/i.test(name)) return "Moon";
  if (/Fullerite|Mykoserocin|Cytoserocin/i.test(name)) return "Gas";
  if (/Ice|Icicle/i.test(name)) return "Ice";
  return "Other";
}

// --- Live Price Fetch (EVEMarketer Jita Region) ---
async function getPrice(typeID) {
  try {
    const resp = await fetch(
      `https://api.evemarketer.com/ec/marketstat/json?typeid=${typeID}&regionlimit=10000002`
    );
    const data = await resp.json();
    return data[0]?.sell?.min || 0;
  } catch (e) {
    console.error("Price fetch failed for typeID", typeID, e);
    return 0;
  }
}

// --- Main Generate Report ---
document.getElementById("generate").addEventListener("click", async () => {
  const input = document.getElementById("miningHold").value.trim();
  if (!input) {
    alert("Paste your mining hold first.");
    return;
  }

  const lines = input.split(/\r?\n/);
  const ores = [];

  // Parse hold into ore objects
  for (let line of lines) {
    const parts = line.split(/\t+|\s{2,}/);
    if (parts.length >= 2) {
      const name = parts.slice(0, -1).join(" ").trim();
      const qty = Number(parts[parts.length - 1].replace(/[,\s]/g, ""));
      if (oreData[name] && qty > 0) {
        ores.push({ name, qty });
      }
    }
  }

  if (!ores.length) {
    alert("No valid ore lines found.");
    return;
  }

  // Group ores into categories
  const categories = {};
  for (const o of ores) {
    const { typeID, volume } = oreData[o.name];
    const price = await getPrice(typeID);
    const total = price * o.qty;
    const vol = volume * o.qty;

    const cat = getCategory(o.name);
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push({ ...o, price, total, volume: vol });
  }

  // Render report
  const reportDiv = document.getElementById("report");
  reportDiv.innerHTML = "";
  let grandTotal = 0;

  for (const [category, rows] of Object.entries(categories)) {
    let catTotal = rows.reduce((s, r) => s + r.total, 0);
    grandTotal += catTotal;

    const section = document.createElement("div");
    section.className = "report-section category-" + category.replace(/\s+/g, "");
    section.innerHTML = `
      <h2>${category}</h2>
      <table>
        <tr><th>Ore</th><th>Qty</th><th>m³</th><th>ISK (Jita)</th><th>Total ISK</th></tr>
        ${rows.map(r => `
          <tr>
            <td>${r.name}</td>
            <td>${r.qty.toLocaleString()}</td>
            <td>${r.volume.toLocaleString()}</td>
            <td>${r.price.toLocaleString()}</td>
            <td>${r.total.toLocaleString()}</td>
          </tr>`).join("")}
        <tr class="total"><td colspan="4">Category Total</td><td>${catTotal.toLocaleString()}</td></tr>
      </table>
    `;
    reportDiv.appendChild(section);
  }

  const grand = document.createElement("div");
  grand.className = "report-section grand-total";
  grand.innerHTML = `<h2>Grand Total: ${grandTotal.toLocaleString()} ISK</h2>`;
  reportDiv.appendChild(grand);

  document.getElementById("downloadExcel").disabled = false;
});

// --- Excel Export ---
document.getElementById("downloadExcel").addEventListener("click", () => {
  const wb = XLSX.utils.book_new();
  const now = new Date();
  const filename = `EVE_Mining_Report_${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}_${String(now.getHours()).padStart(2,"0")}${String(now.getMinutes()).padStart(2,"0")}.xlsx`;

  const sections = document.querySelectorAll(".report-section");
  sections.forEach(section => {
    const title = section.querySelector("h2").innerText;
    const rows = Array.from(section.querySelectorAll("tr")).map(tr =>
      Array.from(tr.querySelectorAll("th,td")).map(td => td.innerText)
    );
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, title.substring(0, 31));
  });

  XLSX.writeFile(wb, filename);
});
