// === Expanded oreData ===
const oreData = {
  // --- Triglavian ores ---
  "Bezdnacine": { typeID: 62307, volume: 16.0 },
  "Compressed Bezdnacine": { typeID: 62520, volume: 16.0 },
  "Rakovene": { typeID: 62308, volume: 16.0 },
  "Compressed Rakovene": { typeID: 62521, volume: 16.0 },
  "Talassonite": { typeID: 62309, volume: 8.0 },
  "Compressed Talassonite": { typeID: 62522, volume: 8.0 },

  // --- Asteroid ores (subset, extend as needed) ---
  "Veldspar": { typeID: 1230, volume: 0.1 },
  "Concentrated Veldspar": { typeID: 17471, volume: 0.1 },
  "Dense Veldspar": { typeID: 17470, volume: 0.1 },
  "Compressed Veldspar": { typeID: 28430, volume: 0.1 },
  "Scordite": { typeID: 1228, volume: 0.15 },
  "Condensed Scordite": { typeID: 17463, volume: 0.15 },
  "Massive Scordite": { typeID: 17464, volume: 0.15 },
  "Compressed Scordite": { typeID: 28428, volume: 0.15 },
  "Plagioclase": { typeID: 18, volume: 0.35 },
  "Compressed Plagioclase": { typeID: 28415, volume: 0.35 },
  "Pyroxeres": { typeID: 1229, volume: 0.3 },
  "Compressed Pyroxeres": { typeID: 28424, volume: 0.3 },
  "Omber": { typeID: 1231, volume: 0.6 },
  "Compressed Omber": { typeID: 28419, volume: 0.6 },
  "Kernite": { typeID: 1232, volume: 1.2 },
  "Compressed Kernite": { typeID: 28413, volume: 1.2 },
  "Jaspet": { typeID: 1233, volume: 2.0 },
  "Compressed Jaspet": { typeID: 28420, volume: 2.0 },
  "Hemorphite": { typeID: 1234, volume: 3.0 },
  "Compressed Hemorphite": { typeID: 28421, volume: 3.0 },
  "Hedbergite": { typeID: 1235, volume: 3.0 },
  "Compressed Hedbergite": { typeID: 28422, volume: 3.0 },
  "Gneiss": { typeID: 1226, volume: 5.0 },
  "Compressed Gneiss": { typeID: 28423, volume: 5.0 },
  "Dark Ochre": { typeID: 1227, volume: 8.0 },
  "Compressed Dark Ochre": { typeID: 28418, volume: 8.0 },
  "Crokite": { typeID: 1236, volume: 16.0 },
  "Compressed Crokite": { typeID: 28412, volume: 16.0 },
  "Bistot": { typeID: 1237, volume: 16.0 },
  "Compressed Bistot": { typeID: 28411, volume: 16.0 },
  "Arkonor": { typeID: 1238, volume: 16.0 },
  "Compressed Arkonor": { typeID: 28410, volume: 16.0 },

  // --- Ice ---
  "Blue Ice": { typeID: 16262, volume: 1000 },
  "Compressed Blue Ice": { typeID: 28433, volume: 1000 },
  "Clear Icicle": { typeID: 16264, volume: 1000 },
  "Compressed Clear Icicle": { typeID: 28436, volume: 1000 },

  // --- Gas ---
  "Fullerite-C28": { typeID: 30370, volume: 1 },
  "Fullerite-C32": { typeID: 30371, volume: 1 },
  "Fullerite-C50": { typeID: 30372, volume: 1 },
  "Fullerite-C60": { typeID: 30373, volume: 1 },
  "Fullerite-C70": { typeID: 30374, volume: 1 },
  "Fullerite-C72": { typeID: 30375, volume: 1 },
  "Fullerite-C84": { typeID: 30376, volume: 1 },
};

// --- Category detection ---
function getCategory(name) {
  if (/Bezdnacine|Rakovene|Talassonite/i.test(name)) return "Triglavian";
  if (/Veldspar|Scordite|Plagioclase|Pyroxeres|Omber|Kernite|Jaspet|Hemorphite|Hedbergite|Gneiss|Dark Ochre|Crokite|Bistot|Arkonor/i.test(name)) return "Asteroid";
  if (/Fullerite|Gas|Mykoserocin|Cytoserocin/i.test(name)) return "Gas";
  if (/Ice|Icicle/i.test(name)) return "Ice";
  if (/Moon|R64|R32|R16|R8/i.test(name)) return "Moon";
  return "Other";
}

// --- Live Jita price fetch ---
async function getPrice(typeID) {
  try {
    const resp = await fetch(
      `https://api.evemarketer.com/ec/marketstat/json?typeid=${typeID}&regionlimit=10000002`
    );
    const data = await resp.json();
    return data[0]?.sell?.min || 0;
  } catch (e) {
    console.error(`Price fetch error for typeID ${typeID}`, e);
    return 0;
  }
}

// --- Report generation ---
document.getElementById("generate").addEventListener("click", async () => {
  const input = document.getElementById("miningHold").value.trim();
  if (!input) {
    alert("Paste your mining hold first.");
    return;
  }

  const lines = input.split(/\r?\n/);
  const ores = [];

  lines.forEach(line => {
    const parts = line.split(/\t+|\s{2,}/);
    if (parts.length >= 2) {
      const name = parts.slice(0, -1).join(" ").trim();
      const qty = Number(parts[parts.length - 1].replace(/[,\s]/g, ""));
      if (name && qty) {
        ores.push({ name, qty });
      }
    }
  });

  if (!ores.length) {
    alert("No valid ore lines found.");
    return;
  }

  const categories = {};
  for (const o of ores) {
    const data = oreData[o.name];
    if (!data) continue;
    const price = await getPrice(data.typeID);
    const volume = data.volume * o.qty;
    const total = price * o.qty;
    const cat = getCategory(o.name);
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push({ ...o, volume, price, total });
  }

  const reportDiv = document.getElementById("report");
  reportDiv.innerHTML = "";
  let grandTotal = 0;

  for (const [category, rows] of Object.entries(categories)) {
    let catTotal = rows.reduce((s, r) => s + r.total, 0);
    grandTotal += catTotal;

    const section = document.createElement("div");
    section.className = "report-section category-" + category.replace(/\s+/g, '');
    section.innerHTML = `
      <h2>${category} Ore</h2>
      <table>
        <tr><th>Ore</th><th>Quantity</th><th>m³</th><th>ISK (Jita)</th><th>Total ISK</th></tr>
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
  const filename = `EVE_Mining_Report_${now.getFullYear()}-${String(
    now.getMonth() + 1
  ).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}_${String(
    now.getHours()
  ).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}.xlsx`;

  const reportDiv = document.getElementById("report");
  const sections = reportDiv.querySelectorAll(".report-section");
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

// --- Starfield + Warp Tunnel Animation ---
const starCanvas = document.getElementById("starfield");
const warpCanvas = document.getElementById("warpfield");
const starCtx = starCanvas.getContext("2d");
const warpCtx = warpCanvas.getContext("2d");
let stars = [];

function resizeCanvas() {
  starCanvas.width = warpCanvas.width = window.innerWidth;
  starCanvas.height = warpCanvas.height = window.innerHeight;
  stars = Array.from({ length: 200 }, () => ({
    x: Math.random() * starCanvas.width,
    y: Math.random() * starCanvas.height,
    z: Math.random() * starCanvas.width
  }));
}

function animate() {
  starCtx.fillStyle = "black";
  starCtx.fillRect(0, 0, starCanvas.width, starCanvas.height);
  warpCtx.clearRect(0, 0, warpCanvas.width, warpCanvas.height);

  for (let star of stars) {
    star.z -= 2;
    if (star.z <= 0) {
      star.x = Math.random() * starCanvas.width;
      star.y = Math.random() * starCanvas.height;
      star.z = starCanvas.width;
    }
    const k = 128.0 / star.z;
    const px = star.x * k + starCanvas.width / 2;
    const py = star.y * k + starCanvas.height / 2;
    if (px >= 0 && px <= starCanvas.width && py >= 0 && py <= starCanvas.height) {
      const size = (1 - star.z / starCanvas.width) * 3;
      starCtx.fillStyle = "white";
      starCtx.fillRect(px, py, size, size);
      warpCtx.strokeStyle = "rgba(0,150,255,0.3)";
      warpCtx.beginPath();
      warpCtx.moveTo(starCanvas.width / 2, starCanvas.height / 2);
      warpCtx.lineTo(px, py);
      warpCtx.stroke();
    }
  }
  requestAnimationFrame(animate);
}

window.addEventListener("resize", resizeCanvas);
resizeCanvas();
animate();
