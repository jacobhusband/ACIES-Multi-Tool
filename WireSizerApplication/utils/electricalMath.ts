import { WIRE_DATA, GROUND_DATA, GEC_DATA, CONDUIT_EMT_40_PERCENT, K_FACTOR_CU, K_FACTOR_AL, MAX_WIRE_SIZE_CU, MAX_WIRE_SIZE_AL, MAX_WIRE_SIZE_FORCED } from '../constants';
import { AppState, CalculationResult, WireSizeData, ConductorMaterial } from '../types';

export const isThreePhaseAllowed = (voltage: number) => voltage !== 120 && voltage !== 277;

export const getHotCount = (phase: number, voltage: number) => {
  if (phase === 3) return 3;
  if (phase === 1) {
    if (voltage === 120 || voltage === 277) return 1;
    if (voltage === 208 || voltage === 240 || voltage === 480) return 2;
  }
  return phase === 3 ? 3 : 2;
};

const indexOf1_0 = WIRE_DATA.findIndex(w => w.size === "1/0");

/** Returns WIRE_DATA filtered to the max allowed size for the given material. */
function getWireTableForMaterial(material: ConductorMaterial, maxSizeOverride?: string): WireSizeData[] {
  const maxSize = maxSizeOverride ?? (material === 'Copper' ? MAX_WIRE_SIZE_CU : MAX_WIRE_SIZE_AL);
  const maxIndex = WIRE_DATA.findIndex(w => w.size === maxSize);
  return maxIndex === -1 ? WIRE_DATA : WIRE_DATA.slice(0, maxIndex + 1);
}

/** Returns the NEC 310.16 ampacity for a wire based on material and temperature rating rules. */
function getAmpacity(wire: WireSizeData, material: ConductorMaterial, wireIndex: number): number {
  const isSmallWire = wireIndex < indexOf1_0;
  if (material === 'Copper') {
    return isSmallWire ? wire.ampacity60Cu : wire.ampacity75Cu;
  } else {
    return isSmallWire ? wire.ampacity60Al : wire.ampacity75Al;
  }
}

interface WireSelection {
  wire: WireSizeData;
  sets: number;
  tempRating: 60 | 75;
}

/**
 * Finds the optimal wire size and minimum number of parallel sets.
 * Iterates sets from 1 upward, for each:
 *   1. Find smallest wire with ampacity >= ampsPerSet
 *   2. Check voltage drop — upsize if needed
 *   3. If upsized wire exceeds material max, try more sets
 * Returns the first valid combination (fewest sets).
 */
function findOptimalWireAndSets(
  amperage: number,
  distance: number,
  voltage: number,
  maxVoltageDrop: number,
  material: ConductorMaterial,
  phaseFactor: number,
  kFactor: number,
  maxSizeOverride?: string
): WireSelection {
  const wireTable = getWireTableForMaterial(material, maxSizeOverride);
  const maxDropVolts = voltage * (maxVoltageDrop / 100);
  const MAX_SETS = 10;

  for (let sets = 1; sets <= MAX_SETS; sets++) {
    const ampsPerSet = amperage / sets;

    // Step 1: Find smallest wire with sufficient ampacity
    let ampacityWire: WireSizeData | undefined;
    for (let i = 0; i < wireTable.length; i++) {
      const globalIndex = WIRE_DATA.findIndex(w => w.size === wireTable[i].size);
      const ampacity = getAmpacity(wireTable[i], material, globalIndex);
      if (ampacity >= ampsPerSet) {
        ampacityWire = wireTable[i];
        break;
      }
    }

    // No wire in allowed table can handle ampsPerSet — need more sets
    if (!ampacityWire) {
      continue;
    }

    // Step 2: Check voltage drop — upsize if needed
    let selectedWire = ampacityWire;

    if (maxDropVolts > 0 && distance > 0) {
      const requiredCM = (kFactor * ampsPerSet * distance * phaseFactor) / maxDropVolts;

      if (selectedWire.circularMils < requiredCM) {
        const upsized = wireTable.find(w => w.circularMils >= requiredCM);
        if (upsized) {
          selectedWire = upsized;
        } else {
          // Largest allowed wire still can't meet VD at this set count — try more sets
          continue;
        }
      }
    }

    // Valid combination found
    const finalIndex = WIRE_DATA.findIndex(w => w.size === selectedWire.size);
    const tempRating: 60 | 75 = finalIndex < indexOf1_0 ? 60 : 75;
    return { wire: selectedWire, sets, tempRating };
  }

  // Fallback (extremely unlikely for real loads)
  const fallbackWire = wireTable[wireTable.length - 1];
  const fallbackIndex = WIRE_DATA.findIndex(w => w.size === fallbackWire.size);
  return {
    wire: fallbackWire,
    sets: MAX_SETS,
    tempRating: fallbackIndex < indexOf1_0 ? 60 : 75
  };
}

function selectWireForSets(
  amperage: number,
  distance: number,
  voltage: number,
  maxVoltageDrop: number,
  material: ConductorMaterial,
  phaseFactor: number,
  kFactor: number,
  sets: number,
  maxSizeOverride?: string
): WireSelection {
  const wireTable = getWireTableForMaterial(material, maxSizeOverride);
  const ampsPerSet = amperage / sets;

  let foundWire: WireSizeData | undefined;
  for (let i = 0; i < wireTable.length; i++) {
    const globalIndex = WIRE_DATA.findIndex(w => w.size === wireTable[i].size);
    const ampacity = getAmpacity(wireTable[i], material, globalIndex);
    if (ampacity >= ampsPerSet) {
      foundWire = wireTable[i];
      break;
    }
  }
  if (!foundWire) {
    foundWire = wireTable[wireTable.length - 1];
  }

  const maxDropVolts = voltage * (maxVoltageDrop / 100);
  if (maxDropVolts > 0 && distance > 0) {
    const requiredCM = (kFactor * ampsPerSet * distance * phaseFactor) / maxDropVolts;
    if (foundWire.circularMils < requiredCM) {
      const upsized = wireTable.find(w => w.circularMils >= requiredCM);
      if (upsized) foundWire = upsized;
    }
  }

  const finalIndex = WIRE_DATA.findIndex(w => w.size === foundWire.size);
  const tempRating: 60 | 75 = finalIndex < indexOf1_0 ? 60 : 75;
  return { wire: foundWire, sets, tempRating };
}

export const calculateEverything = (state: AppState): CalculationResult => {
  const { voltage, phase, material, oversizeConduit } = state;
  const effectivePhase = isThreePhaseAllowed(voltage) ? phase : 1;

  // Sanitize inputs (handle empty strings)
  const amperage = state.amperage === '' ? 0 : state.amperage;
  const distance = state.distance === '' ? 0 : state.distance;
  const maxVoltageDrop = state.maxVoltageDrop === '' ? 3 : state.maxVoltageDrop;

  // User's manual sets value
  const userSets = (state.sets === '' || state.sets === 0) ? 1 : state.sets;
  const forceSets = state.forceSets;

  const kFactor = material === 'Copper' ? K_FACTOR_CU : K_FACTOR_AL;
  const phaseFactor = effectivePhase === 3 ? 1.732 : 2;

  // Auto-calculate optimal wire and sets
  const autoResult = findOptimalWireAndSets(
    amperage, distance, voltage, maxVoltageDrop, material, phaseFactor, kFactor
  );

  // Use the greater of user's manual sets and auto-calculated sets unless forcing
  const effectiveSets = forceSets ? userSets : Math.max(userSets, autoResult.sets);

  // If user override forced more sets, re-select wire for the higher set count
  let selectedWire: WireSizeData;
  let finalTempRating: 60 | 75;

  if (forceSets) {
    const forcedResult = selectWireForSets(
      amperage, distance, voltage, maxVoltageDrop, material, phaseFactor, kFactor, effectiveSets, MAX_WIRE_SIZE_FORCED
    );
    selectedWire = forcedResult.wire;
    finalTempRating = forcedResult.tempRating;
  } else if (effectiveSets > autoResult.sets) {
    const overrideResult = selectWireForSets(
      amperage, distance, voltage, maxVoltageDrop, material, phaseFactor, kFactor, effectiveSets
    );
    selectedWire = overrideResult.wire;
    finalTempRating = overrideResult.tempRating;
  } else {
    selectedWire = autoResult.wire;
    finalTempRating = autoResult.tempRating;
  }

  const sets = effectiveSets;
  const ampsPerSet = amperage / sets;

  // Voltage drop with the selected wire
  const actualCM = selectedWire.circularMils;
  const actualVD = actualCM > 0
    ? (kFactor * ampsPerSet * distance * phaseFactor) / actualCM
    : 0;
  const actualVDPercent = voltage > 0 ? (actualVD / voltage) * 100 : 0;
  const voltsAtLoad = voltage - actualVD;

  // Actual ampacity
  let actualAmpacity = 0;
  if (material === 'Copper') {
    actualAmpacity = finalTempRating === 60 ? selectedWire.ampacity60Cu : selectedWire.ampacity75Cu;
  } else {
    actualAmpacity = finalTempRating === 60 ? selectedWire.ampacity60Al : selectedWire.ampacity75Al;
  }

  // 3. Grounding Conductor Sizing
  let groundWireSize: string;

  if (state.groundingTable === 'GEC') {
    // NEC Table 250.66 — lookup by service-entrance conductor size (circular mils)
    const conductorCM = selectedWire.circularMils * sets;
    const gecRow = GEC_DATA.find(g => {
      const threshold = material === 'Copper' ? g.maxCuCM : g.maxAlCM;
      return conductorCM <= threshold;
    });
    groundWireSize = gecRow
      ? (material === 'Copper' ? gecRow.gecCuSize : gecRow.gecAlSize)
      : "See Eng.";
  } else {
    // NEC Table 250.122 — lookup by overcurrent device rating
    const groundRow = GROUND_DATA.find(g => g.rating >= amperage);
    groundWireSize = groundRow
      ? (material === 'Copper' ? groundRow.cuSize : groundRow.alSize)
      : "See Eng.";
  }

  // 4. Conduit Fill Calculation (per set — each parallel set gets its own conduit)
  const hotCount = getHotCount(effectivePhase, voltage);
  const numCurrentCarrying = hotCount + 1;
  const totalWiresPerSet = numCurrentCarrying;

  const groundWireObj = WIRE_DATA.find(w => w.size === groundWireSize);
  const groundArea = groundWireObj ? groundWireObj.areaSqIn : 0.02;

  const totalFillArea = (selectedWire.areaSqIn * totalWiresPerSet) + groundArea;

  // Constraint: Minimum conduit size is 3/4"
  const minConduitIndex = CONDUIT_EMT_40_PERCENT.findIndex(c => c.size === "3/4");
  const availableConduits = CONDUIT_EMT_40_PERCENT.slice(minConduitIndex);

  let recommendedConduitIndex = availableConduits.findIndex(c => c.area >= totalFillArea);

  if (recommendedConduitIndex === -1) {
    recommendedConduitIndex = availableConduits.length - 1;
  }

  // Handle Oversizing
  if (oversizeConduit && recommendedConduitIndex < availableConduits.length - 1) {
    recommendedConduitIndex++;
  }

  const recommendedConduit = availableConduits[recommendedConduitIndex];

  return {
    recommendedSize: selectedWire.size,
    actualAmpacity: actualAmpacity * sets,
    voltageDrop: actualVD,
    voltageDropPercentage: actualVDPercent,
    voltageAtLoad: voltsAtLoad,
    groundWireSize: groundWireSize,
    conduitSize: recommendedConduit ? recommendedConduit.size : "4\"+",
    conduitType: "EMT",
    maxDistanceFor3Percent: actualCM > 0
      ? (actualCM * (voltage * 0.03)) / (kFactor * ampsPerSet * phaseFactor)
      : 0,
    wireAreaTotal: totalFillArea,
    conduitFillPercentage: recommendedConduit ? (totalFillArea / (recommendedConduit.area / 0.4)) * 100 : 100,
    tempRatingUsed: finalTempRating as 60 | 75,
    sets,
    recommendedSets: autoResult.sets,
  };
};
