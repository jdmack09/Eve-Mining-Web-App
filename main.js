async function fetchJSON(url) {
  const res = await fetch(url);
  return res.json();
}

async function getTypeId(name) {
  const url = "https://www.fuzzwork.co.uk/api/typeid.php?typename=" + encodeURIComponent(name.trim());
  const json = await fetchJSON(url);
  return json && json.typeID ? Number(json.typeID) : null;
}

async function getTypeInfo(typeId) {
  const url = `https://esi.evetech.net/latest/universe/types/${typeId}/?datasource=tranquility&language=en`;
  return await fetchJSON(url);
}

async function getVolume(typeId) {
  const info = await getTypeInfo(typeId);
  return info && info.volume ? Number(info.volume) : 0;
}

async function getPrices(typeId) {
  const url = `https://market.fuzzwork.co.uk/aggregates/?station=60003760&types=${typeId}`; // Jita 4-4
  const json = await fetchJSON(url);
  const rec = json[String(typeId)] || {};
  return rec.sell && rec.sell.min ? rec.sell.min : 0;
}

async function getCategory(typeId) {
  const info = await getTypeInfo(typeId);
  const groupId = info.group_id;
  if (!groupId) return "Other";
  const g = await fetchJSON(`https://esi.evetech.net/latest/universe/groups/${groupId}/?datasource=tranquility&language=en`);
  const gname = g && g.name ? g.name : "Other";
  if (/Ice/i.test(gname)) return "Ice Product";
  if (/Fullerite|Mykoserocin|Gas/i.test(gname)) return "Gas Cloud";
  if (/Moon/i.test(gname)) return "Moon Ore";
  if (/Triglavian|Abyss|Talassonite|Rakovene|Bezdnacine/i.test(gname)) return "Trig Ore";
  if (/Asteroid|Ore/i.test(gname)) return "Asteroid Ore";
  return gname;
}

let lastReportData = []; // store rows for Excel export

document.getElementById("generate").addEventListener("click", async () => {
  const pasted = document.getElementById("miningHold").value.trim();
  if (!pasted) {
    alert("Paste your mining hold first!");
    return;
  }

  const lines = pasted.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const rows = [];
  for (const line of lines) {
    const m = line.split(/\t+|\s{2,}/);
    if (m.length < 2) continue;
    const name = m.slice(0, -1).join(" ").trim();
    const qty = Number(String(m[m.length - 1]).replace(/[,\s]/g, "")) || 0;
    if (!name || !qty) continue;
    rows.push({ name, qty });
  }

  let grandTotal = 0;
  let table = `
    <table>
      <tr>
        <th>Name</th><th>Category</th><th>Compressed?</th>
        <th>Qty</th><th>m³ Each</th><th>Total m³</th>
        <th>Jita Sell ISK</th><th>Total ISK</th>
      </tr>`;

  lastReportData = [["Name", "Category", "Compressed?", "Qty", "m³ Each", "Total m³", "Jita Sell ISK", "Total ISK"]];

  for (const r of rows) {
    const typeId = await getTypeId(r.name);
    if (!typeId) continue;
    const volume = await getVolume(typeId);
    const cat = await getCategory(typeId);
    const sellMin = await getPrices(typeId);

    const totalM3 = r.qty * volume;
    const totalISK = r.qty * sellMin;
    grandTotal += totalISK;

    table += `
      <tr>
        <td>${r.name.replace(/^Compressed\s+/i, "")}</td>
        <td>${cat}</td>
        <td>${/^Compressed\s+/i.test(r.name) ? "Yes" : "No"}</td>
        <td>${r.qty}</td>
        <td>${volume}</td>
        <td>${totalM3}</td>
        <td>${sellMin}</td>
        <td>${totalISK}</td>
      </tr>`;

    lastReportData.push([
      r.name.replace(/^Compressed\s+/i, ""),
      cat,
      /^Compressed\s+/i.test(r.name) ? "Yes" : "No",
      r.qty,
      volume,
      totalM3,
      sellMin,
      totalISK
    ]);
  }

  table += `<tr class="total"><td colspan="7">Grand Total (ISK)</td><td>${grandTotal}</td></tr></table>`;
  lastReportData.push(["Grand Total", "", "", "", "", "", "", grandTotal]);

  document.getElementById("report").innerHTML = table;
  document.getElementById("downloadExcel").disabled = false;
});

// Excel Export
document.getElementById("downloadExcel").addEventListener("click", () => {
  if (!lastReportData.length) return;

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(lastReportData);
  XLSX.utils.book_append_sheet(wb, ws, "Mining Report");

  // Format date & time for filename: YYYY-MM-DD_HHMM
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;

  XLSX.writeFile(wb, `EVE_Mining_Report_${timestamp}.xlsx`);
});