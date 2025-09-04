/* ===============================
   EVE Mining Tracker – main.js
   - Static oreData (wide coverage) + safe fallback lookups
   - Jita prices via EVEMarketer
   - Category grouping, report render, Excel export
   =============================== */

// -------- Helpers: formatters
const fmtISK = (n) => (isFinite(n) ? Math.round(n).toLocaleString() : "0");
const fmtVOL = (n) => (isFinite(n) ? n.toLocaleString() : "0");

// -------- Category labels used in UI
const CATS = {
  AST: "Asteroid",
  MOON: "Moon",
  TRIG: "Triglavian",
  GAS: "Gas",
  ICE: "Ice",
  OTHER: "Other"
};

// =====================================================
//  Big static ore database (IDs/volumes known where possible)
//  NOTE: For any entry with typeID:null or volume:null,
//        the code will fetch it once via Fuzzwork as fallback.
// =====================================================
const oreData = {
  // ========= Asteroid Ores (Hi/Low/Null) =========
  // Veldspar family
  "Veldspar":                       { typeID: 1230,  volume: 0.1,  category: CATS.AST },
  "Concentrated Veldspar":          { typeID: 17471, volume: 0.1,  category: CATS.AST },
  "Dense Veldspar":                 { typeID: 17470, volume: 0.1,  category: CATS.AST },
  "Compressed Veldspar":            { typeID: 28430, volume: 0.1,  category: CATS.AST },
  "Compressed Concentrated Veldspar":{ typeID: 28432, volume: 0.1, category: CATS.AST },
  "Compressed Dense Veldspar":      { typeID: 28431, volume: 0.1,  category: CATS.AST },

  // Scordite family
  "Scordite":                       { typeID: 1228,  volume: 0.15, category: CATS.AST },
  "Condensed Scordite":             { typeID: 17463, volume: 0.15, category: CATS.AST },
  "Massive Scordite":               { typeID: 17464, volume: 0.15, category: CATS.AST },
  "Compressed Scordite":            { typeID: 28428, volume: 0.15, category: CATS.AST },
  "Compressed Condensed Scordite":  { typeID: 28429, volume: 0.15, category: CATS.AST },
  "Compressed Massive Scordite":    { typeID: 28427, volume: 0.15, category: CATS.AST },

  // Plagioclase family
  "Plagioclase":                    { typeID: 18,    volume: 0.35, category: CATS.AST },
  "Azure Plagioclase":              { typeID: 17459, volume: 0.35, category: CATS.AST },
  "Rich Plagioclase":               { typeID: 17460, volume: 0.35, category: CATS.AST },
  "Compressed Plagioclase":         { typeID: 28415, volume: 0.35, category: CATS.AST },
  "Compressed Azure Plagioclase":   { typeID: 28417, volume: 0.35, category: CATS.AST },
  "Compressed Rich Plagioclase":    { typeID: 28416, volume: 0.35, category: CATS.AST },

  // Pyroxeres family
  "Pyroxeres":                      { typeID: 1229,  volume: 0.3,  category: CATS.AST },
  "Solid Pyroxeres":                { typeID: 17461, volume: 0.3,  category: CATS.AST },
  "Viscous Pyroxeres":              { typeID: 17462, volume: 0.3,  category: CATS.AST },
  "Compressed Pyroxeres":           { typeID: 28424, volume: 0.3,  category: CATS.AST },
  "Compressed Solid Pyroxeres":     { typeID: 28426, volume: 0.3,  category: CATS.AST },
  "Compressed Viscous Pyroxeres":   { typeID: 28425, volume: 0.3,  category: CATS.AST },

  // Omber family
  "Omber":                          { typeID: 1231,  volume: 0.6,  category: CATS.AST },
  "Silvery Omber":                  { typeID: 17867, volume: 0.6,  category: CATS.AST },
  "Golden Omber":                   { typeID: 17868, volume: 0.6,  category: CATS.AST },
  "Compressed Omber":               { typeID: 28419, volume: 0.6,  category: CATS.AST },
  "Compressed Silvery Omber":       { typeID: 28433, volume: 0.6,  category: CATS.AST },
  "Compressed Golden Omber":        { typeID: 28434, volume: 0.6,  category: CATS.AST },

  // Kernite family
  "Kernite":                        { typeID: 20,    volume: 1.2,  category: CATS.AST },
  "Luminous Kernite":               { typeID: 17455, volume: 1.2,  category: CATS.AST },
  "Fiery Kernite":                  { typeID: 17456, volume: 1.2,  category: CATS.AST },
  "Compressed Kernite":             { typeID: 28413, volume: 1.2,  category: CATS.AST },
  "Compressed Luminous Kernite":    { typeID: 28436, volume: 1.2,  category: CATS.AST },
  "Compressed Fiery Kernite":       { typeID: 28435, volume: 1.2,  category: CATS.AST },

  // Jaspet family
  "Jaspet":                         { typeID: 1226,  volume: 2.0,  category: CATS.AST },
  "Pure Jaspet":                    { typeID: 17452, volume: 2.0,  category: CATS.AST },
  "Pristine Jaspet":                { typeID: 17453, volume: 2.0,  category: CATS.AST },
  "Compressed Jaspet":              { typeID: 28420, volume: 2.0,  category: CATS.AST },
  "Compressed Pure Jaspet":         { typeID: 28437, volume: 2.0,  category: CATS.AST },
  "Compressed Pristine Jaspet":     { typeID: 28438, volume: 2.0,  category: CATS.AST },

  // Hemorphite family
  "Hemorphite":                     { typeID: 1234,  volume: 3.0,  category: CATS.AST },
  "Vivid Hemorphite":               { typeID: 17444, volume: 3.0,  category: CATS.AST },
  "Radiant Hemorphite":             { typeID: 17445, volume: 3.0,  category: CATS.AST },
  "Compressed Hemorphite":          { typeID: 28421, volume: 3.0,  category: CATS.AST },
  "Compressed Vivid Hemorphite":    { typeID: 28439, volume: 3.0,  category: CATS.AST },
  "Compressed Radiant Hemorphite":  { typeID: 28440, volume: 3.0,  category: CATS.AST },

  // Hedbergite family
  "Hedbergite":                     { typeID: 21,    volume: 3.0,  category: CATS.AST },
  "Vitric Hedbergite":              { typeID: 17440, volume: 3.0,  category: CATS.AST },
  "Glazed Hedbergite":              { typeID: 17441, volume: 3.0,  category: CATS.AST },
  "Compressed Hedbergite":          { typeID: 28422, volume: 3.0,  category: CATS.AST },
  "Compressed Vitric Hedbergite":   { typeID: 28441, volume: 3.0,  category: CATS.AST },
  "Compressed Glazed Hedbergite":   { typeID: 28442, volume: 3.0,  category: CATS.AST },

  // Gneiss family
  "Gneiss":                         { typeID: 1229 + 9997, volume: 5.0, category: CATS.AST }, // typeID placeholder safety
  "Iridescent Gneiss":              { typeID: 17865, volume: 5.0, category: CATS.AST },
  "Prismatic Gneiss":               { typeID: 17866, volume: 5.0, category: CATS.AST },
  "Compressed Gneiss":              { typeID: 28423, volume: 5.0, category: CATS.AST },
  "Compressed Iridescent Gneiss":   { typeID: 28443, volume: 5.0, category: CATS.AST },
  "Compressed Prismatic Gneiss":    { typeID: 28444, volume: 5.0, category: CATS.AST },

  // Dark Ochre family
  "Dark Ochre":                     { typeID: 1227,  volume: 8.0,  category: CATS.AST },
  "Obsidian Ochre":                 { typeID: 17449, volume: 8.0,  category: CATS.AST },
  "Onyx Ochre":                     { typeID: 17450, volume: 8.0,  category: CATS.AST },
  "Compressed Dark Ochre":          { typeID: 28418, volume: 8.0,  category: CATS.AST },
  "Compressed Obsidian Ochre":      { typeID: 28445, volume: 8.0,  category: CATS.AST },
  "Compressed Onyx Ochre":          { typeID: 28446, volume: 8.0,  category: CATS.AST },

  // Crokite family
  "Crokite":                        { typeID: 1236,  volume: 16.0, category: CATS.AST },
  "Sharp Crokite":                  { typeID: 17433, volume: 16.0, category: CATS.AST },
  "Crystalline Crokite":            { typeID: 17432, volume: 16.0, category: CATS.AST },
  "Compressed Crokite":             { typeID: 28412, volume: 16.0, category: CATS.AST },
  "Compressed Sharp Crokite":       { typeID: 28447, volume: 16.0, category: CATS.AST },
  "Compressed Crystalline Crokite": { typeID: 28448, volume: 16.0, category: CATS.AST },

  // Bistot family
  "Bistot":                         { typeID: 1237,  volume: 16.0, category: CATS.AST },
  "Triclinic Bistot":               { typeID: 17428, volume: 16.0, category: CATS.AST },
  "Monoclinic Bistot":              { typeID: 17429, volume: 16.0, category: CATS.AST },
  "Compressed Bistot":              { typeID: 28411, volume: 16.0, category: CATS.AST },
  "Compressed Triclinic Bistot":    { typeID: 28449, volume: 16.0, category: CATS.AST },
  "Compressed Monoclinic Bistot":   { typeID: 28450, volume: 16.0, category: CATS.AST },

  // Arkonor family
  "Arkonor":                        { typeID: 1238,  volume: 16.0, category: CATS.AST },
  "Crimson Arkonor":                { typeID: 17425, volume: 16.0, category: CATS.AST },
  "Prime Arkonor":                  { typeID: 17426, volume: 16.0, category: CATS.AST },
  "Compressed Arkonor":             { typeID: 28410, volume: 16.0, category: CATS.AST },
  "Compressed Crimson Arkonor":     { typeID: 28451, volume: 16.0, category: CATS.AST },
  "Compressed Prime Arkonor":       { typeID: 28452, volume: 16.0, category: CATS.AST },

  // Spodumain family
  "Spodumain":                      { typeID: 19,    volume: 16.0, category: CATS.AST },
  "Bright Spodumain":               { typeID: 17466, volume: 16.0, category: CATS.AST },
  "Gleaming Spodumain":             { typeID: 17467, volume: 16.0, category: CATS.AST },
  "Compressed Spodumain":           { typeID: 28414, volume: 16.0, category: CATS.AST },
  "Compressed Bright Spodumain":    { typeID: 28453, volume: 16.0, category: CATS.AST },
  "Compressed Gleaming Spodumain":  { typeID: 28454, volume: 16.0, category: CATS.AST },

  // Mercoxit family
  "Mercoxit":                       { typeID: 11396, volume: 40.0, category: CATS.AST },
  "Magma Mercoxit":                 { typeID: 17869, volume: 40.0, category: CATS.AST },
  "Vitreous Mercoxit":              { typeID: 17870, volume: 40.0, category: CATS.AST },
  "Compressed Mercoxit":            { typeID: 28409, volume: 40.0, category: CATS.AST },
  "Compressed Magma Mercoxit":      { typeID: 28455, volume: 40.0, category: CATS.AST },
  "Compressed Vitreous Mercoxit":   { typeID: 28456, volume: 40.0, category: CATS.AST },

  // ========= Moon Ores (tiers) =========
  // Ubiquitous
  "Bitumens":                       { typeID: 45490, volume: 1.0, category: CATS.MOON },
  "Coesite":                        { typeID: 45491, volume: 1.0, category: CATS.MOON },
  "Sylvite":                        { typeID: 45492, volume: 1.0, category: CATS.MOON },
  "Zeolites":                       { typeID: 45493, volume: 1.0, category: CATS.MOON },
  "Compressed Bitumens":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Coesite":             { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Sylvite":             { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Zeolites":            { typeID: null,  volume: 1.0, category: CATS.MOON },

  // Common
  "Cobaltite":                      { typeID: 45494, volume: 1.0, category: CATS.MOON },
  "Euxenite":                       { typeID: 45495, volume: 1.0, category: CATS.MOON },
  "Titanite":                       { typeID: 45496, volume: 1.0, category: CATS.MOON },
  "Scheelite":                      { typeID: 45497, volume: 1.0, category: CATS.MOON },
  "Compressed Cobaltite":           { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Euxenite":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Titanite":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Scheelite":           { typeID: null,  volume: 1.0, category: CATS.MOON },

  // Uncommon
  "Otavite":                        { typeID: 45498, volume: 1.0, category: CATS.MOON },
  "Sperrylite":                     { typeID: 45499, volume: 1.0, category: CATS.MOON },
  "Vanadinite":                     { typeID: 45500, volume: 1.0, category: CATS.MOON },
  "Chromite":                       { typeID: 45501, volume: 1.0, category: CATS.MOON },
  "Compressed Otavite":             { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Sperrylite":          { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Vanadinite":          { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Chromite":            { typeID: null,  volume: 1.0, category: CATS.MOON },

  // Rare
  "Carnotite":                      { typeID: 45502, volume: 1.0, category: CATS.MOON },
  "Zircon":                         { typeID: 45503, volume: 1.0, category: CATS.MOON },
  "Pollucite":                      { typeID: 45504, volume: 1.0, category: CATS.MOON },
  "Cinnabar":                       { typeID: 45505, volume: 1.0, category: CATS.MOON },
  "Compressed Carnotite":           { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Zircon":              { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Pollucite":           { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Cinnabar":            { typeID: null,  volume: 1.0, category: CATS.MOON },

  // Exceptional (R64)
  "Xenotime":                       { typeID: 45506, volume: 1.0, category: CATS.MOON },
  "Monazite":                       { typeID: 45507, volume: 1.0, category: CATS.MOON },
  "Loparite":                       { typeID: 45508, volume: 1.0, category: CATS.MOON },
  "Ytterbite":                      { typeID: 45509, volume: 1.0, category: CATS.MOON },
  "Compressed Xenotime":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Monazite":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Loparite":            { typeID: null,  volume: 1.0, category: CATS.MOON },
  "Compressed Ytterbite":           { typeID: null,  volume: 1.0, category: CATS.MOON },

  // ========= Triglavian Ores =========
  "Bezdnacine":                     { typeID: 62578, volume: 16.0, category: CATS.TRIG },
  "Rakovene":                       { typeID: 62580, volume: 16.0, category: CATS.TRIG },
  "Talassonite":                    { typeID: 62582, volume: 8.0,  category: CATS.TRIG },

  "Compressed Bezdnacine":          { typeID: null,  volume: 16.0, category: CATS.TRIG },
  "Compressed Rakovene":            { typeID: null,  volume: 16.0, category: CATS.TRIG },
  "Compressed Talassonite":         { typeID: null,  volume: 8.0,  category: CATS.TRIG },

  // Some UI variants seen in pastes
  "Abyssal Bezdnacine":             { typeID: 62578, volume: 16.0, category: CATS.TRIG },
  "Compressed Abyssal Bezdnacine":  { typeID: null,  volume: 16.0, category: CATS.TRIG },

  // ========= Ice =========
  "Blue Ice":                       { typeID: 16262, volume: 1000.0, category: CATS.ICE },
  "Clear Icicle":                   { typeID: 16263, volume: 1000.0, category: CATS.ICE },
  "Glacial Mass":                   { typeID: 16264, volume: 1000.0, category: CATS.ICE },
  "White Glaze":                    { typeID: 16265, volume: 1000.0, category: CATS.ICE },
  "Dark Glitter":                   { typeID: 17978, volume: 1000.0, category: CATS.ICE },
  "Gelidus":                        { typeID: 17976, volume: 1000.0, category: CATS.ICE },
  "Krystallos":                     { typeID: 17977, volume: 1000.0, category: CATS.ICE },
  "Glare Crust":                    { typeID: 16266, volume: 1000.0, category: CATS.ICE },

  "Compressed Blue Ice":            { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Clear Icicle":        { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Glacial Mass":        { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed White Glaze":         { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Dark Glitter":        { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Gelidus":             { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Krystallos":          { typeID: null,  volume: 1000.0, category: CATS.ICE },
  "Compressed Glare Crust":         { typeID: null,  volume: 1000.0, category: CATS.ICE },

  // ========= Gas (Mykoserocin, Cytoserocin, Fullerites) =========
  // Mykoserocin (HS/LS)
  "Amber Mykoserocin":              { typeID: 25268, volume: 10.0, category: CATS.GAS },
  "Golden Mykoserocin":             { typeID: 25270, volume: 10.0, category: CATS.GAS },
  "Celadon Mykoserocin":            { typeID: 25279, volume: 10.0, category: CATS.GAS },
  "Lime Mykoserocin":               { typeID: 28694, volume: 10.0, category: CATS.GAS },
  "Malachite Mykoserocin":          { typeID: 25274, volume: 10.0, category: CATS.GAS },
  "Vermillion Mykoserocin":         { typeID: 25271, volume: 10.0, category: CATS.GAS },
  "Teal Mykoserocin":               { typeID: 28695, volume: 10.0, category: CATS.GAS },
  "Viridian Mykoserocin":           { typeID: 28696, volume: 10.0, category: CATS.GAS },

  "Compressed Amber Mykoserocin":   { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Golden Mykoserocin":  { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Celadon Mykoserocin": { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Lime Mykoserocin":    { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Malachite Mykoserocin":{ typeID: null, volume: 10.0, category: CATS.GAS },
  "Compressed Vermillion Mykoserocin":{ typeID: null,volume: 10.0, category: CATS.GAS },
  "Compressed Teal Mykoserocin":    { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Viridian Mykoserocin":{ typeID: null,  volume: 10.0, category: CATS.GAS },

  // Cytoserocin (LS/Null)
  "Azure Cytoserocin":              { typeID: 28697, volume: 10.0, category: CATS.GAS },
  "Crimson Cytoserocin":            { typeID: 28698, volume: 10.0, category: CATS.GAS },
  "IVORY Cytoserocin":              { typeID: 28699, volume: 10.0, category: CATS.GAS },
  "LADAR Cytoserocin":              { typeID: 28700, volume: 10.0, category: CATS.GAS }, // placeholder names (rare)
  "Lime Cytoserocin":               { typeID: 28701, volume: 10.0, category: CATS.GAS },
  "Emerald Cytoserocin":            { typeID: 28702, volume: 10.0, category: CATS.GAS },
  "Golden Cytoserocin":             { typeID: 28703, volume: 10.0, category: CATS.GAS },
  "Viridian Cytoserocin":           { typeID: 28704, volume: 10.0, category: CATS.GAS },

  "Compressed Azure Cytoserocin":   { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Crimson Cytoserocin": { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Ivory Cytoserocin":   { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Lime Cytoserocin":    { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Emerald Cytoserocin": { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Golden Cytoserocin":  { typeID: null,  volume: 10.0, category: CATS.GAS },
  "Compressed Viridian Cytoserocin":{ typeID: null,  volume: 10.0, category: CATS.GAS },

  // Fullerites (W-space)
  "Fullerite-C28":                  { typeID: 30375, volume: 1.0, category: CATS.GAS },
  "Fullerite-C32":                  { typeID: 30376, volume: 1.0, category: CATS.GAS },
  "Fullerite-C50":                  { typeID: 30370, volume: 1.0, category: CATS.GAS },
  "Fullerite-C60":                  { typeID: 30371, volume: 1.0, category: CATS.GAS },
  "Fullerite-C70":                  { typeID: 30372, volume: 1.0, category: CATS.GAS },
  "Fullerite-C72":                  { typeID: 30373, volume: 1.0, category: CATS.GAS },
  "Fullerite-C84":                  { typeID: 30374, volume: 1.0, category: CATS.GAS },
  "Fullerite-C320":                 { typeID: 30378, volume: 1.0, category: CATS.GAS },
  "Fullerite-C540":                 { typeID: 30380, volume: 1.0, category: CATS.GAS },

  "Compressed Fullerite-C28":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C32":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C50":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C60":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C70":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C72":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C84":       { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C320":      { typeID: null,  volume: 1.0, category: CATS.GAS },
  "Compressed Fullerite-C540":      { typeID: null,  volume: 1.0, category: CATS.GAS }
};

// =====================================================
// Caches + helpers
// =====================================================
const typeCache = new Map();
const priceCache = new Map();

function normalizeName(name) {
  return String(name || "").trim();
}

function guessCategory(name) {
  if (/Bezdnacine|Rakovene|Talassonite/i.test(name)) return CATS.TRIG;
  if (/Veldspar|Scordite|Plagioclase|Pyroxeres|Omber|Kernite|Jaspet|Hemorphite|Hedbergite|Gneiss|Ochre|Crokite|Bistot|Arkonor|Spodumain|Mercoxit/i.test(name)) return CATS.AST;
  if (/Bitumens|Coesite|Sylvite|Zeolites|Cobaltite|Euxenite|Titanite|Scheelite|Otavite|Sperrylite|Vanadinite|Chromite|Carnotite|Zircon|Pollucite|Cinnabar|Xenotime|Monazite|Loparite|Ytterbite/i.test(name)) return CATS.MOON;
  if (/Fullerite|Mykoserocin|Cytoserocin/i.test(name)) return CATS.GAS;
  if (/Ice|Icicle|Glacial|Glare|Glitter|Gelidus|Krystallos|Glaze/i.test(name)) return CATS.ICE;
  return CATS.OTHER;
}

// --- Resolve type info (static → fallback to Fuzzwork)
async function resolveType(name) {
  const key = normalizeName(name);
  if (typeCache.has(key)) return typeCache.get(key);

  let entry = oreData[key];
  if (entry && entry.typeID && entry.volume != null) {
    const resolved = { typeID: entry.typeID, volume: entry.volume, name: key, category: entry.category || guessCategory(key) };
    typeCache.set(key, resolved);
    return resolved;
  }

  // fallback: fuzzwork lookup
  try {
    const resp = await fetch(`https://www.fuzzwork.co.uk/api/typeid2.php?typename=${encodeURIComponent(key)}`);
    const data = await resp.json();
    if (data && data.typeid) {
      const resolved = {
        typeID: data.typeid,
        volume: Number(data.volume) || (entry?.volume || 0),
        name: data.name || key,
        category: entry?.category || guessCategory(data.name || key)
      };
      typeCache.set(key, resolved);
      return resolved;
    }
  } catch (e) {
    console.error("Resolve error", key, e);
  }

  typeCache.set(key, null);
  return null;
}

// --- Live Price Fetch with fallback ---
async function getPrice(typeID) {
  if (!typeID) return 0;
  if (priceCache.has(typeID)) return priceCache.get(typeID);

  const hubs = {
    "10000002": "Jita",
    "10000043": "Amarr",
    "10000032": "Dodixie",
    "10000030": "Hek",
    "10000042": "Rens"
  };

  const regionID = document.getElementById("hubSelect")?.value || "10000002";
  const hubName = hubs[regionID] || "Jita";

  try {
    // EvEMarketer
    const url = `https://api.evemarketer.com/ec/marketstat/json?typeid=${typeID}&regionlimit=${regionID}`;
    console.log(`Fetching ${hubName} price (EvEMarketer):`, url);
    const resp = await fetch(url);
    const data = await resp.json();
    let price = Number(data?.[0]?.sell?.min) || 0;
    if (price > 0) {
      priceCache.set(typeID, price);
      return price;
    }

    // Fallback: Fuzzwork
    const fwUrl = `https://market.fuzzwork.co.uk/aggregates/?region=${regionID}&types=${typeID}`;
    console.log(`Fallback fetching ${hubName} price (Fuzzwork):`, fwUrl);
    const fwResp = await fetch(fwUrl);
    const fwData = await fwResp.json();
    price = Number(fwData?.[typeID]?.sell?.min) || 0;
    priceCache.set(typeID, price);
    return price;
  } catch (e) {
    console.error("Price fetch failed for typeID", typeID, e);
    priceCache.set(typeID, 0);
    return 0;
  }
}

// =====================================================
// Generate Report
// =====================================================
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

  const buckets = {};
  const unresolved = [];
  let grandTotal = 0;

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
    buckets[resolved.category].push({ name: resolved.name, qty: row.qty, price, total, volume: vol });
    grandTotal += total;
  }

  const reportDiv = document.getElementById("report");
  reportDiv.innerHTML = "";

  const catOrder = [CATS.AST, CATS.MOON, CATS.TRIG, CATS.GAS, CATS.ICE, CATS.OTHER];
  for (const cat of catOrder) {
    const rows = buckets[cat];
    if (!rows || !rows.length) continue;

    rows.sort((a, b) => b.total - a.total);
    const catTotal = rows.reduce((s, r) => s + r.total, 0);

    const section = document.createElement("div");
    section.className = "report-section category-" + cat.replace(/\s+/g, "");
    section.innerHTML = `
      <h2>${cat}</h2>
      <table>
        <tr><th>Ore</th><th>Qty</th><th>m³</th><th>ISK (${document.getElementById("hubSelect")?.selectedOptions[0].text})</th><th>Total ISK</th></tr>
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

  const grand = document.createElement("div");
  grand.className = "report-section grand-total";
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

// =====================================================
// Excel Export
// =====================================================
document.getElementById("downloadExcel").addEventListener("click", () => {
  const wb = XLSX.utils.book_new();
  const now = new Date();
  const filename = `EVE_Mining_Report_${now.toISOString().slice(0,10)}_${now.getHours()}${now.getMinutes()}.xlsx`;

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