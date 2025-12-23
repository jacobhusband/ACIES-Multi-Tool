import { WireSizeData, GroundingData } from './types';

// Simplified Table 310.16 (60°C and 75°C Columns) + Physical Properties
export const WIRE_DATA: WireSizeData[] = [
  { size: "14", circularMils: 4110, ampacity60Cu: 15, ampacity75Cu: 20, ampacity60Al: 0, ampacity75Al: 0, areaSqIn: 0.0097 },
  { size: "12", circularMils: 6530, ampacity60Cu: 20, ampacity75Cu: 25, ampacity60Al: 15, ampacity75Al: 20, areaSqIn: 0.0133 },
  { size: "10", circularMils: 10380, ampacity60Cu: 30, ampacity75Cu: 35, ampacity60Al: 25, ampacity75Al: 30, areaSqIn: 0.0211 },
  { size: "8", circularMils: 16510, ampacity60Cu: 40, ampacity75Cu: 50, ampacity60Al: 30, ampacity75Al: 40, areaSqIn: 0.0366 },
  { size: "6", circularMils: 26240, ampacity60Cu: 55, ampacity75Cu: 65, ampacity60Al: 40, ampacity75Al: 50, areaSqIn: 0.0507 },
  { size: "4", circularMils: 41740, ampacity60Cu: 70, ampacity75Cu: 85, ampacity60Al: 55, ampacity75Al: 65, areaSqIn: 0.0824 },
  { size: "3", circularMils: 52620, ampacity60Cu: 85, ampacity75Cu: 100, ampacity60Al: 65, ampacity75Al: 75, areaSqIn: 0.0973 },
  { size: "2", circularMils: 66360, ampacity60Cu: 95, ampacity75Cu: 115, ampacity60Al: 75, ampacity75Al: 90, areaSqIn: 0.1158 },
  { size: "1", circularMils: 83690, ampacity60Cu: 110, ampacity75Cu: 130, ampacity60Al: 85, ampacity75Al: 100, areaSqIn: 0.1562 },
  { size: "1/0", circularMils: 105600, ampacity60Cu: 125, ampacity75Cu: 150, ampacity60Al: 100, ampacity75Al: 120, areaSqIn: 0.1855 },
  { size: "2/0", circularMils: 133100, ampacity60Cu: 145, ampacity75Cu: 175, ampacity60Al: 115, ampacity75Al: 135, areaSqIn: 0.2223 },
  { size: "3/0", circularMils: 167800, ampacity60Cu: 165, ampacity75Cu: 200, ampacity60Al: 130, ampacity75Al: 155, areaSqIn: 0.2679 },
  { size: "4/0", circularMils: 211600, ampacity60Cu: 195, ampacity75Cu: 230, ampacity60Al: 150, ampacity75Al: 180, areaSqIn: 0.3237 },
  { size: "250", circularMils: 250000, ampacity60Cu: 215, ampacity75Cu: 255, ampacity60Al: 170, ampacity75Al: 205, areaSqIn: 0.3970 },
  { size: "300", circularMils: 300000, ampacity60Cu: 240, ampacity75Cu: 285, ampacity60Al: 190, ampacity75Al: 230, areaSqIn: 0.4608 },
  { size: "350", circularMils: 350000, ampacity60Cu: 260, ampacity75Cu: 310, ampacity60Al: 210, ampacity75Al: 250, areaSqIn: 0.5242 },
  { size: "400", circularMils: 400000, ampacity60Cu: 280, ampacity75Cu: 335, ampacity60Al: 225, ampacity75Al: 270, areaSqIn: 0.5863 },
  { size: "500", circularMils: 500000, ampacity60Cu: 320, ampacity75Cu: 380, ampacity60Al: 260, ampacity75Al: 310, areaSqIn: 0.7073 },
  { size: "600", circularMils: 600000, ampacity60Cu: 350, ampacity75Cu: 420, ampacity60Al: 285, ampacity75Al: 340, areaSqIn: 0.8676 },
];

// Table 250.122 Minimum Size Equipment Grounding Conductors
export const GROUND_DATA: GroundingData[] = [
  { rating: 15, cuSize: "14", alSize: "12" },
  { rating: 20, cuSize: "12", alSize: "10" },
  { rating: 60, cuSize: "10", alSize: "8" },
  { rating: 100, cuSize: "8", alSize: "6" },
  { rating: 200, cuSize: "6", alSize: "4" },
  { rating: 300, cuSize: "4", alSize: "2" },
  { rating: 400, cuSize: "3", alSize: "1" },
  { rating: 500, cuSize: "2", alSize: "1/0" },
  { rating: 600, cuSize: "1", alSize: "2/0" },
  { rating: 800, cuSize: "1/0", alSize: "3/0" },
  { rating: 1000, cuSize: "2/0", alSize: "4/0" },
  { rating: 1200, cuSize: "3/0", alSize: "250" },
  { rating: 1600, cuSize: "4/0", alSize: "350" },
];

// Simplified internal area of EMT Conduit (Trade Size -> Sq Inches @ 40% fill)
// Using Chapter 9 Table 4 (40% fill column for over 2 wires)
export const CONDUIT_EMT_40_PERCENT: { size: string; area: number }[] = [
  { size: "1/2", area: 0.122 },
  { size: "3/4", area: 0.213 },
  { size: "1", area: 0.346 },
  { size: "1-1/4", area: 0.598 },
  { size: "1-1/2", area: 0.814 },
  { size: "2", area: 1.342 },
  { size: "2-1/2", area: 1.924 },
  { size: "3", area: 2.907 },
  { size: "3-1/2", area: 3.848 },
  { size: "4", area: 4.947 },
];

export const K_FACTOR_CU = 12.9;
export const K_FACTOR_AL = 21.2;