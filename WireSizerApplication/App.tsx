import React, { useState, useEffect } from 'react';
import { Settings, Zap, ArrowRight, Cable, AlertTriangle, Info, Ruler, Activity, Copy, CheckCircle2 } from 'lucide-react';
import { AppState, CalculationResult } from './types';
import { calculateEverything, getHotCount, isThreePhaseAllowed } from './utils/electricalMath';
import { InputGroup } from './components/InputGroup';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip } from 'recharts';

// Styles for form inputs
const inputClass = "w-full p-3 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white transition-all text-gray-800 font-medium";
const selectClass = "w-full p-3 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white transition-all text-gray-800 font-medium";

function App() {
  const [state, setState] = useState<AppState>({
    voltage: 208,
    amperage: 20,
    distance: 100,
    phase: 3,
    material: 'Copper',
    maxVoltageDrop: 3,
    sets: 1,
    powerFactor: 0.9,
    oversizeConduit: false,
  });

  const [results, setResults] = useState<CalculationResult | null>(null);
  const [copyFeedback, setCopyFeedback] = useState(false);

  useEffect(() => {
    const res = calculateEverything(state);
    setResults(res);
  }, [state]);

  const handleChange = (field: keyof AppState, value: any) => {
    setState(prev => {
      if (field === 'voltage') {
        const nextVoltage = value;
        const nextPhase = isThreePhaseAllowed(nextVoltage) ? prev.phase : 1;
        return { ...prev, voltage: nextVoltage, phase: nextPhase };
      }
      if (field === 'phase') {
        const nextPhase = value;
        if (nextPhase === 3 && !isThreePhaseAllowed(prev.voltage)) {
          return { ...prev, phase: 1 };
        }
        return { ...prev, phase: nextPhase };
      }
      return { ...prev, [field]: value };
    });
  };

  const handleNumChange = (field: keyof AppState, value: string) => {
    if (value === '') {
      handleChange(field, '');
    } else {
      handleChange(field, Number(value));
    }
  };

  const copySpecToClipboard = () => {
    if (!results) return;
    
    // Format: 3/4"C., 3#8 H, 1#8 N, 1#12 G
    const conduit = `${results.conduitSize}"C.`;
    const effectivePhase = isThreePhaseAllowed(state.voltage) ? state.phase : 1;
    const numHots = getHotCount(effectivePhase, state.voltage);
    const wireSize = results.recommendedSize;
    const gndSize = results.groundWireSize;
    
    const setsPrefix = (state.sets !== '' && state.sets > 1) ? `(${state.sets} sets) ` : '';
    const line1 = `${setsPrefix}${conduit}, ${numHots}#${wireSize} H, 1#${wireSize} N, 1#${gndSize} G`;
    
    const line2 = `LENGTH ~= ${state.distance}'`;
    const line3 = `VOLTAGE DROP ~= ${results.voltageDropPercentage.toFixed(2)}%`;
    
    const spec = `${line1}\n${line2}\n${line3}`;
    
    navigator.clipboard.writeText(spec).then(() => {
      setCopyFeedback(true);
      setTimeout(() => setCopyFeedback(false), 2000);
    });
  };

  const getVoltageStatusColor = (percent: number) => {
    if (percent <= 3) return "text-green-600";
    if (percent <= 5) return "text-yellow-600";
    return "text-red-600";
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12 text-slate-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg">
              <Zap className="text-white w-5 h-5" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-900 leading-tight">ACIES Wire Sizer</h1>
              <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">Engineering Utility</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button 
              onClick={copySpecToClipboard}
              className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 transition-all shadow-sm active:scale-95"
            >
              {copyFeedback ? <CheckCircle2 size={16} className="text-green-600" /> : <Copy size={16} />}
              {copyFeedback ? "Copied!" : "Copy Spec"}
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* INPUTS PANEL */}
          <div className="lg:col-span-4 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200">
              <div className="bg-slate-50 px-5 py-3 border-b border-gray-200 flex items-center gap-2">
                <Settings className="w-4 h-4 text-gray-500" />
                <h2 className="font-bold text-sm text-gray-700">Circuit Configuration</h2>
              </div>
              
              <div className="p-5 space-y-4">
                <div className="grid grid-cols-2 gap-3">
                  <InputGroup label="Phase">
                    <select 
                      value={state.phase} 
                      onChange={(e) => handleChange('phase', Number(e.target.value))}
                      className={selectClass}
                    >
                      <option value={1}>1 Phase</option>
                      <option value={3} disabled={!isThreePhaseAllowed(state.voltage)}>3 Phase</option>
                    </select>
                  </InputGroup>
                  <InputGroup label="Voltage">
                    <select 
                      value={state.voltage} 
                      onChange={(e) => handleChange('voltage', Number(e.target.value))}
                      className={selectClass}
                    >
                      {[120, 208, 240, 277, 480].map(v => <option key={v} value={v}>{v}V</option>)}
                    </select>
                  </InputGroup>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <InputGroup label="Load (Amps)">
                    <input type="number" value={state.amperage} onChange={(e) => handleNumChange('amperage', e.target.value)} className={inputClass} />
                  </InputGroup>
                   <InputGroup label="Dist. (ft)">
                    <input type="number" value={state.distance} onChange={(e) => handleNumChange('distance', e.target.value)} className={inputClass} />
                  </InputGroup>
                </div>

                <InputGroup label="Conductor">
                  <div className="grid grid-cols-2 bg-gray-100 p-1 rounded-lg">
                    {['Copper', 'Aluminum'].map(m => (
                      <button 
                        key={m}
                        onClick={() => handleChange('material', m)}
                        className={`py-2 text-sm font-bold rounded-md transition-all ${state.material === m ? 'bg-white shadow-sm text-blue-600' : 'text-gray-500 hover:text-gray-700'}`}
                      >
                        {m}
                      </button>
                    ))}
                  </div>
                </InputGroup>

                <div className="grid grid-cols-2 gap-3">
                   <InputGroup label="V.D. Max %">
                     <input type="number" step="0.1" value={state.maxVoltageDrop} onChange={(e) => handleNumChange('maxVoltageDrop', e.target.value)} className={inputClass} />
                  </InputGroup>
                  <InputGroup label="Parallel Sets">
                    <input type="number" min="1" value={state.sets} onChange={(e) => handleNumChange('sets', e.target.value)} className={inputClass} />
                  </InputGroup>
                </div>

                <div className="pt-2">
                  <label className="flex items-center gap-3 cursor-pointer group">
                    <input 
                      type="checkbox" 
                      checked={state.oversizeConduit}
                      onChange={(e) => handleChange('oversizeConduit', e.target.checked)}
                      className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                    />
                    <span className="text-sm font-semibold text-gray-700 group-hover:text-blue-600 transition-colors">
                      Oversize conduit for safety
                    </span>
                  </label>
                </div>
              </div>
            </div>
          </div>

          {/* RESULTS PANEL */}
          <div className="lg:col-span-8 space-y-6">
            {results && (
              <>
                <div className="bg-white rounded-xl shadow-md border border-gray-200 overflow-hidden">
                   <div className="p-6">
                     <div className="flex flex-col md:flex-row justify-between gap-6 mb-6">
                       <div>
                         <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1 block">Feeder Specification</span>
                         <div className="flex items-baseline gap-2">
                           <span className="text-4xl font-black text-slate-900">
                             {state.sets !== '' && state.sets > 1 ? `${state.sets}x ` : ''}{results.recommendedSize}
                           </span>
                           <span className="text-lg text-slate-500 font-bold">AWG</span>
                         </div>
                         <p className="text-sm text-slate-500 font-medium mt-1">
                           {state.material} THHN • {state.phase === 3 ? '3PH+N' : '1PH+N'} • {results.tempRatingUsed}°C
                         </p>
                       </div>
                       
                       <div className="flex gap-4">
                         <div className="bg-slate-50 rounded-lg p-3 border border-slate-100 text-center min-w-[100px]">
                           <span className="text-[10px] font-bold text-slate-400 block mb-1 uppercase">Ground</span>
                           <span className="text-lg font-bold text-slate-700">#{results.groundWireSize}</span>
                         </div>
                         <div className="bg-blue-50 rounded-lg p-3 border border-blue-100 text-center min-w-[100px]">
                           <span className="text-[10px] font-bold text-blue-400 block mb-1 uppercase">Conduit</span>
                           <span className="text-lg font-bold text-blue-700">{results.conduitSize}"</span>
                         </div>
                       </div>
                     </div>

                     {/* Visual Stats Row */}
                     <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="border border-slate-100 rounded-xl p-4 bg-slate-50/50">
                          <div className="flex justify-between items-center mb-3">
                            <h4 className="text-xs font-bold text-slate-500 uppercase">Voltage Drop</h4>
                            <span className={`text-xs font-bold ${getVoltageStatusColor(results.voltageDropPercentage)}`}>
                              {results.voltageDropPercentage.toFixed(2)}%
                            </span>
                          </div>
                          <div className="w-full bg-slate-200 rounded-full h-1.5 mb-2">
                            <div 
                              className={`h-1.5 rounded-full transition-all duration-500 ${results.voltageDropPercentage > (state.maxVoltageDrop || 3) ? 'bg-red-500' : 'bg-green-500'}`}
                              style={{ width: `${Math.min(100, (results.voltageDropPercentage / (state.maxVoltageDrop || 3)) * 100)}%` }}
                            ></div>
                          </div>
                          <p className="text-[10px] text-slate-400">Limit: {state.maxVoltageDrop || 3}% • At Load: {results.voltageAtLoad.toFixed(1)}V</p>
                        </div>

                        <div className="border border-slate-100 rounded-xl p-4 bg-slate-50/50">
                          <div className="flex justify-between items-center mb-3">
                            <h4 className="text-xs font-bold text-slate-500 uppercase">Conduit Fill</h4>
                            <span className="text-xs font-bold text-blue-600">{results.conduitFillPercentage.toFixed(1)}% Total</span>
                          </div>
                          <div className="w-full bg-slate-200 rounded-full h-1.5 mb-2">
                            <div 
                              className="h-1.5 bg-blue-500 rounded-full transition-all duration-500"
                              style={{ width: `${Math.min(100, (results.conduitFillPercentage / 40) * 100)}%` }}
                            ></div>
                          </div>
                          <p className="text-[10px] text-slate-400">NEC Limit: 40% • Fill Area: {results.wireAreaTotal.toFixed(3)} sq.in</p>
                        </div>
                     </div>
                   </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-4">
                  <div className="flex gap-3 text-xs text-gray-500">
                    <Info size={14} className="text-blue-500 flex-shrink-0" />
                    <p>
                      Calculations adhere to <strong>NEC 2023</strong> standards. Terminals assumed at 60°C for loads ≤100A or wires &le;1 AWG, and 75°C for loads &gt;100A or wires &ge;1/0 AWG. Bars show 100% at code-defined limits (3% V.D. and 40% Fill).
                    </p>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;
