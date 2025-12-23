import { WIRE_DATA, GROUND_DATA, CONDUIT_EMT_40_PERCENT, K_FACTOR_CU, K_FACTOR_AL } from '../constants';
import { AppState, CalculationResult, WireSizeData } from '../types';

export const isThreePhaseAllowed = (voltage: number) => voltage !== 120 && voltage !== 277;

export const getHotCount = (phase: number, voltage: number) => {
  if (phase === 3) return 3;
  if (phase === 1) {
    if (voltage === 120 || voltage === 277) return 1;
    if (voltage === 208 || voltage === 240 || voltage === 480) return 2;
  }
  return phase === 3 ? 3 : 2;
};

export const calculateEverything = (state: AppState): CalculationResult => {
  const { voltage, phase, material, oversizeConduit } = state;
  const effectivePhase = isThreePhaseAllowed(voltage) ? phase : 1;
  
  // Sanitize inputs (handle empty strings)
  const amperage = state.amperage === '' ? 0 : state.amperage;
  const distance = state.distance === '' ? 0 : state.distance;
  const maxVoltageDrop = state.maxVoltageDrop === '' ? 3 : state.maxVoltageDrop;
  
  // Avoid division by zero for sets
  const sets = (state.sets === '' || state.sets === 0) ? 1 : state.sets;

  const kFactor = material === 'Copper' ? K_FACTOR_CU : K_FACTOR_AL;
  
  // 1. Determine Minimum Wire Size based on Ampacity (NEC 310.16)
  const ampsPerSet = amperage / sets;
  
  const indexOf1_0 = WIRE_DATA.findIndex(w => w.size === "1/0");

  let selectedWire: WireSizeData | undefined = WIRE_DATA.find((w, index) => {
    const isSmallWire = index < indexOf1_0;
    let ampacity = 0;
    if (material === 'Copper') {
      ampacity = isSmallWire ? w.ampacity60Cu : w.ampacity75Cu;
    } else {
      ampacity = isSmallWire ? w.ampacity60Al : w.ampacity75Al;
    }
    return ampacity >= ampsPerSet;
  });

  if (!selectedWire) {
    selectedWire = WIRE_DATA[WIRE_DATA.length - 1];
  }

  const selectedIndex = WIRE_DATA.findIndex(w => w.size === selectedWire!.size);
  const tempRatingUsed = selectedIndex < indexOf1_0 ? 60 : 75;

  // 2. Voltage Drop Calculation
  const phaseFactor = effectivePhase === 3 ? 1.732 : 2;
  const maxDropVolts = voltage * (maxVoltageDrop / 100);
  
  const requiredCM = maxDropVolts > 0 
    ? (kFactor * ampsPerSet * distance * phaseFactor) / maxDropVolts
    : 0;

  if (selectedWire.circularMils < requiredCM) {
    const upsized = WIRE_DATA.find(w => w.circularMils >= requiredCM);
    if (upsized) {
      selectedWire = upsized;
    } else {
      selectedWire = WIRE_DATA[WIRE_DATA.length - 1];
    }
  }
  
  const finalIndex = WIRE_DATA.findIndex(w => w.size === selectedWire!.size);
  const finalTempRating = finalIndex < indexOf1_0 ? 60 : 75;

  const actualCM = selectedWire.circularMils;
  const actualVD = (kFactor * ampsPerSet * distance * phaseFactor) / actualCM;
  const actualVDPercent = (actualVD / voltage) * 100;
  const voltsAtLoad = voltage - actualVD;
  
  let actualAmpacity = 0;
  if (material === 'Copper') {
    actualAmpacity = finalTempRating === 60 ? selectedWire.ampacity60Cu : selectedWire.ampacity75Cu;
  } else {
    actualAmpacity = finalTempRating === 60 ? selectedWire.ampacity60Al : selectedWire.ampacity75Al;
  }

  // 3. Grounding Conductor Sizing (NEC 250.122)
  const groundRow = GROUND_DATA.find(g => g.rating >= amperage);
  const groundWireSize = groundRow 
    ? (material === 'Copper' ? groundRow.cuSize : groundRow.alSize)
    : "See Eng.";

  // 4. Conduit Fill Calculation
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
    maxDistanceFor3Percent: (actualCM * (voltage * 0.03)) / (kFactor * ampsPerSet * phaseFactor),
    wireAreaTotal: totalFillArea,
    conduitFillPercentage: recommendedConduit ? (totalFillArea / (recommendedConduit.area / 0.4)) * 100 : 100,
    tempRatingUsed: finalTempRating as 60 | 75
  };
};
