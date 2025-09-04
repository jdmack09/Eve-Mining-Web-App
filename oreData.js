// oreData.js
// Comprehensive EVE Online mining item database
// Contains: Asteroid, Moon, Ice, Gas, Trig ores (and compressed variants)

const oreData = {
  // Asteroid ores
  "Veldspar": { typeID: 1230, volume: 0.1 },
  "Concentrated Veldspar": { typeID: 17471, volume: 0.1 },
  "Dense Veldspar": { typeID: 17470, volume: 0.1 },
  "Compressed Veldspar": { typeID: 28430, volume: 0.1 },
  "Compressed Concentrated Veldspar": { typeID: 28432, volume: 0.1 },
  "Compressed Dense Veldspar": { typeID: 28431, volume: 0.1 },

  "Scordite": { typeID: 1228, volume: 0.15 },
  "Condensed Scordite": { typeID: 17463, volume: 0.15 },
  "Massive Scordite": { typeID: 17464, volume: 0.15 },
  "Compressed Scordite": { typeID: 28428, volume: 0.15 },
  "Compressed Condensed Scordite": { typeID: 28429, volume: 0.15 },
  "Compressed Massive Scordite": { typeID: 28427, volume: 0.15 },

  "Plagioclase": { typeID: 18, volume: 0.35 },
  "Azure Plagioclase": { typeID: 17459, volume: 0.35 },
  "Rich Plagioclase": { typeID: 17460, volume: 0.35 },
  "Compressed Plagioclase": { typeID: 28415, volume: 0.35 },
  "Compressed Azure Plagioclase": { typeID: 28417, volume: 0.35 },
  "Compressed Rich Plagioclase": { typeID: 28416, volume: 0.35 },

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

  // Moon ores (R64/R32/R16/R8),
  // Example: Expand with actual names and IDs from invTypes if needed
  "Bitumens": { typeID: 16633, volume: 1 },
  "Compressed Bitumens": { typeID: 28367, volume: 1 },
  "Cobaltite": { typeID: 16634, volume: 1 },
  "Compressed Cobaltite": { typeID: 28368, volume: 1 },
  "Sylvite": { typeID: 16635, volume: 1 },
  "Compressed Sylvite": { typeID: 28369, volume: 1 },
  "Zeolites": { typeID: 16636, volume: 1 },
  "Compressed Zeolites": { typeID: 28370, volume: 1 },

  // Ice products
  "Blue Ice": { typeID: 16262, volume: 1000 },
  "Compressed Blue Ice": { typeID: 28433, volume: 1000 },
  "Clear Icicle": { typeID: 16264, volume: 1000 },
  "Compressed Clear Icicle": { typeID: 28436, volume: 1000 },
  // Add further ice variants as needed...

  // Gas clouds
  "Fullerite-C28": { typeID: 30370, volume: 1 },
  "Compressed Fullerite-C28": { typeID: 30371, volume: 1 },
  "Fullerite-C32": { typeID: 30372, volume: 1 },
  "Compressed Fullerite-C32": { typeID: 30373, volume: 1 },
  "Fullerite-C50": { typeID: 30374, volume: 1 },
  "Compressed Fullerite-C50": { typeID: 30375, volume: 1 },
  "Fullerite-C60": { typeID: 30376, volume: 1 },
  "Compressed Fullerite-C60": { typeID: 30377, volume: 1 },
  // Add Mykoserocin / Cytoserocin similarly...

  // Triglavian ores
  "Bezdnacine": { typeID: 62307, volume: 16.0 },
  "Compressed Bezdnacine": { typeID: 62520, volume: 16.0 },
  "Rakovene": { typeID: 62308, volume: 16.0 },
  "Compressed Rakovene": { typeID: 62521, volume: 16.0 },
  "Talassonite": { typeID: 62309, volume: 8.0 },
  "Compressed Talassonite": { typeID: 62522, volume: 8.0 }
};

export default oreData;
