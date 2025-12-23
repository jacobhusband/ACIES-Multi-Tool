export type ConductorMaterial = 'Copper' | 'Aluminum';
export type ConduitType = 'EMT' | 'PVC_Sch40' | 'RMC';
export type Phase = 1 | 3;
export type InsulationType = 'THHN/THWN-2' | 'XHHW-2';

export interface WireSizeData {
  size: string;
  circularMils: number;
  ampacity60Cu: number;
  ampacity75Cu: number;
  ampacity60Al: number;
  ampacity75Al: number;
  areaSqIn: number; // For conduit fill
}

export interface GroundingData {
  rating: number; // Rating of overcurrent device
  cuSize: string;
  alSize: string;
}

export interface CalculationResult {
  recommendedSize: string;
  actualAmpacity: number;
  voltageDrop: number;
  voltageDropPercentage: number;
  voltageAtLoad: number;
  groundWireSize: string;
  conduitSize: string;
  conduitType: string;
  maxDistanceFor3Percent: number;
  wireAreaTotal: number;
  conduitFillPercentage: number;
  tempRatingUsed: 60 | 75;
}

export interface AppState {
  voltage: number;
  amperage: number | '';
  distance: number | ''; // in feet
  phase: Phase;
  material: ConductorMaterial;
  maxVoltageDrop: number | ''; // percentage
  sets: number | ''; // Parallel runs
  powerFactor: number;
  oversizeConduit: boolean;
}