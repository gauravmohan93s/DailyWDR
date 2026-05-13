"use client";

import React, { useState, useEffect, useRef } from 'react';
import { 
  Play, Settings, Users, History, AlertCircle, CheckCircle2, 
  Terminal, ShieldCheck, Calendar as CalendarIcon, Save,
  Search, Upload, X, Filter, RefreshCcw, LayoutDashboard,
  ChevronRight, ArrowRight, Table, Layers, Plus, Trash2, Mail,
  Eye, EyeOff, MapPin, Briefcase
} from 'lucide-react';
import axios from 'axios';

const API_BASE = "http://localhost:8000";
const WS_BASE = "ws://localhost:8000/ws/logs";

export default function Dashboard() {
  const [activeTab, setActiveTab] = useState('operations');
  const [config, setConfig] = useState({ 
    REPORT_DATE_OVERRIDE: '', 
    DRY_RUN: true,
    NAME_FILTER: '',
    ROLE_FILTER: '',
    TEAM_FILTER: '',
    REGION_FILTER: '',
    SUBREGION_FILTER: '',
    options: { roles: [], regions: [], subregions: [], names: [] }
  });
  const [roster, setRoster] = useState([]);
  const [matrix, setMatrix] = useState([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [matrixSearch, setMatrixSearch] = useState('');
  const [progress, setProgress] = useState({ msg: 'Ready', value: 0, status: 'idle' });
  const [logs, setLogs] = useState([]);
  const logEndRef = useRef(null);
  const fileInputRef = useRef(null);

  // --- Initialization ---
  useEffect(() => {
    fetchConfig();
    fetchRoster();
    fetchMatrix();
    connectLogs();
  }, []);

  const fetchConfig = async () => {
    try {
      const res = await axios.get(`${API_BASE}/config`);
      setConfig(res.data);
    } catch (err) { console.error("Config fetch failed", err); }
  };

  const fetchRoster = async (query = '') => {
    try {
      const res = await axios.get(`${API_BASE}/roster`, { params: { search: query } });
      setRoster(res.data);
    } catch (err) { console.error("Roster fetch failed", err); }
  };

  const fetchMatrix = async (query = '') => {
    try {
      const res = await axios.get(`${API_BASE}/matrix`, { params: { search: query } });
      setMatrix(res.data);
    } catch (err) { console.error("Matrix fetch failed", err); }
  };

  const connectLogs = () => {
    const ws = new WebSocket(WS_BASE);
    ws.onmessage = (event) => {
      const data = JSON.parse(event.data);
      if (data.type === 'status') {
        setProgress({ msg: data.msg, value: data.progress, status: 'running' });
      } else if (data.type === 'success') {
        setProgress({ msg: data.msg, value: 100, status: 'success' });
      } else if (data.type === 'error') {
        setProgress({ msg: data.msg, value: progress.value, status: 'error' });
      }
      setLogs(prev => [...prev, data]);
    };
    ws.onclose = () => setTimeout(connectLogs, 3000);
  };

  const saveConfig = async () => {
    try {
      await axios.post(`${API_BASE}/config`, { ...config, DISK_CLEANUP_DAYS: 14 });
      alert("Reporting scope and filters saved!");
    } catch (err) { alert("Failed to save settings"); }
  };

  const triggerPipeline = async () => {
    if (!confirm("Launch pipeline with current filters?")) return;
    setProgress({ msg: 'Initializing...', value: 5, status: 'running' });
    setLogs([]);
    try { await axios.post(`${API_BASE}/run-pipeline`); } 
    catch (err) { alert("Failed to start pipeline"); setProgress({ ...progress, status: 'error' }); }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    try {
      const res = await axios.post(`${API_BASE}/roster/import`, formData);
      alert(res.data.message);
      fetchRoster();
      fetchMatrix();
    } catch (err) { alert("Import failed: " + (err.response?.data?.detail || err.message)); }
  };

  const saveMatrix = async () => {
    try {
      await axios.post(`${API_BASE}/matrix`, matrix);
      alert("Lookups saved! Roster managers and inclusion status updated.");
      fetchRoster();
    } catch (err) { alert("Failed to save matrix"); }
  };

  // --- Derived Stats ---
  const activeMembers = roster.filter(m => m.Include).length;
  const uniqueRolesCount = new Set(roster.map(m => m.Role)).size;
  const uniqueTeamsCount = new Set(roster.map(m => `${m.Region}|${m.SubRegion}`)).size;

  return (
    <div className="min-h-screen bg-[#F1F5F9] text-slate-900 font-sans flex overflow-hidden h-screen">
      {/* Sidebar */}
      <aside className="w-64 bg-slate-950 text-white h-full flex flex-col p-5 shadow-2xl shrink-0 border-r border-white/5">
        <div className="flex items-center gap-3 mb-10 px-2">
          <div className="p-2 bg-blue-600 rounded-xl shadow-lg shadow-blue-500/30">
            <Layers size={20} className="text-white" />
          </div>
          <div>
            <h1 className="text-base font-black tracking-tight leading-none uppercase italic">DAILY WORK</h1>
            <span className="text-[9px] font-bold text-blue-400 tracking-[0.2em] uppercase">Done Report</span>
          </div>
        </div>

        <nav className="flex-1 space-y-1.5">
          <NavItem icon={<LayoutDashboard size={18}/>} label="Operations" active={activeTab === 'operations'} onClick={() => setActiveTab('operations')} />
          <NavItem icon={<Users size={18}/>} label="Roster Manager" active={activeTab === 'roster'} onClick={() => setActiveTab('roster')} />
          <NavItem icon={<Table size={18}/>} label="Ownership Matrix" active={activeTab === 'matrix'} onClick={() => setActiveTab('matrix')} />
        </nav>
        
        <div className="mt-auto p-4 bg-white/5 rounded-2xl border border-white/5">
            <div className="flex items-center justify-between mb-1.5">
                <span className="text-[8px] font-black text-slate-500 uppercase tracking-widest text-nowrap">Core Engine</span>
                <div className={`w-1.5 h-1.5 rounded-full ${progress.status === 'running' ? 'bg-blue-500 animate-pulse' : 'bg-emerald-500'}`}></div>
            </div>
            <p className="text-[10px] font-bold text-slate-300 truncate uppercase tracking-tighter">{progress.msg}</p>
        </div>
      </aside>

      <main className="flex-1 overflow-y-auto bg-[#F8FAFC]">
        {activeTab === 'operations' && (
          <div className="p-10 space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-500 max-w-[1400px] mx-auto">
            <header className="flex justify-between items-start border-b border-slate-200 pb-8">
              <div>
                <h2 className="text-3xl font-black tracking-tighter text-slate-900 uppercase italic">Command Center</h2>
                <p className="text-slate-400 font-bold text-xs uppercase tracking-widest mt-1">Configure reporting scope and execute delivery pipeline.</p>
              </div>
              <div className="flex gap-3">
                <button onClick={saveConfig} className="px-5 py-2.5 bg-white border border-slate-200 rounded-xl font-bold text-slate-600 hover:bg-slate-50 transition-all flex items-center gap-2 shadow-sm text-xs uppercase tracking-widest">
                  <Save size={16} /> Save Config
                </button>
                <button 
                    onClick={triggerPipeline} 
                    disabled={progress.status === 'running'}
                    className={`px-8 py-2.5 rounded-xl font-black text-white transition-all shadow-lg flex items-center gap-2 text-xs uppercase tracking-widest ${progress.status === 'running' ? 'bg-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 shadow-blue-500/20 active:scale-95'}`}
                >
                  <Play size={16} fill="white" />
                  {progress.status === 'running' ? 'RUNNING...' : 'LAUNCH PIPELINE'}
                </button>
              </div>
            </header>

            {/* Visual Stepper */}
            <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-4">
                    <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">Pipeline Monitor</h3>
                    <span className="text-[10px] font-black text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full">{progress.value}%</span>
                </div>
                <div className="relative h-2.5 bg-slate-100 rounded-full overflow-hidden mb-6">
                    <div className="absolute top-0 left-0 h-full bg-blue-600 transition-all duration-1000 ease-out" style={{ width: `${progress.value}%` }}></div>
                </div>
                <div className="grid grid-cols-4 gap-4">
                    <StepItem label="Validation" active={progress.value >= 10} done={progress.value > 15} />
                    <StepItem label="Data Sync" active={progress.value >= 30} done={progress.value > 35} />
                    <StepItem label="Aggregation" active={progress.value >= 50} done={progress.value > 55} />
                    <StepItem label="Delivery" active={progress.value >= 70} done={progress.value === 100} />
                </div>
            </div>

            {/* High-Density Filter Suite */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <FilterCard label="Report Date" icon={<CalendarIcon size={12}/>}>
                 <input type="date" value={config.REPORT_DATE_OVERRIDE} onChange={(e) => setConfig({...config, REPORT_DATE_OVERRIDE: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1 text-sm"/>
              </FilterCard>
              
              <FilterCard label="Roles (Multi-Select)" icon={<Briefcase size={12}/>}>
                <input type="text" placeholder="Separate with commas..." list="roles-list" value={config.ROLE_FILTER} onChange={(e) => setConfig({...config, ROLE_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1 text-xs placeholder:font-bold placeholder:text-slate-300"/>
                <datalist id="roles-list">{config.options?.roles?.map(r => <option key={r} value={r}/>)}</datalist>
              </FilterCard>

              <FilterCard label="Regions (Multi-Select)" icon={<MapPin size={12}/>}>
                <input type="text" placeholder="Separate with commas..." list="regions-list" value={config.REGION_FILTER || ''} onChange={(e) => setConfig({...config, REGION_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1 text-xs placeholder:font-bold placeholder:text-slate-300"/>
                <datalist id="regions-list">{config.options?.regions?.map(r => <option key={r} value={r}/>)}</datalist>
              </FilterCard>

              <FilterCard label="Sub Regions (Multi-Select)" icon={<MapPin size={12}/>}>
                <input type="text" placeholder="Separate with commas..." list="subregions-list" value={config.SUBREGION_FILTER || ''} onChange={(e) => setConfig({...config, SUBREGION_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1 text-xs placeholder:font-bold placeholder:text-slate-300"/>
                <datalist id="subregions-list">{config.options?.subregions?.map(t => <option key={t} value={t}/>)}</datalist>
              </FilterCard>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <label className="p-5 bg-white rounded-3xl border border-slate-200 shadow-sm flex items-center justify-between cursor-pointer hover:border-blue-300 transition-all">
                    <div>
                        <p className="font-black text-slate-800 text-xs uppercase tracking-widest">DRY RUN MODE</p>
                        <p className="text-[10px] font-bold text-slate-400 uppercase mt-0.5">Generate previews without sending</p>
                    </div>
                    <input type="checkbox" checked={config.DRY_RUN} onChange={(e) => setConfig({...config, DRY_RUN: e.target.checked})} className="w-5 h-5 accent-blue-600 rounded-lg shadow-inner"/>
                </label>
                <div className="p-5 bg-slate-950 rounded-3xl border border-slate-800 shadow-xl flex items-center gap-4 overflow-hidden relative">
                    <Terminal className="text-slate-700 shrink-0" size={24} />
                    <div className="font-mono text-[10px] text-slate-400 truncate tracking-tight">
                        {logs.slice(-1)[0]?.msg || 'Pipeline console ready...'}
                    </div>
                    <div className="absolute right-0 top-0 h-full w-12 bg-gradient-to-l from-slate-950 to-transparent"></div>
                </div>
            </div>
          </div>
        )}

        {activeTab === 'roster' && (
          <div className="p-10 space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-500 max-w-[1400px] mx-auto">
            <header className="flex justify-between items-center border-b border-slate-200 pb-8">
                <div>
                    <h2 className="text-3xl font-black tracking-tighter text-slate-900 uppercase italic">Roster Manager</h2>
                    <p className="text-slate-400 font-bold text-xs uppercase tracking-widest mt-1 italic">Official database of active team members.</p>
                </div>
                <div className="flex gap-3">
                    <div className="relative group">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" size={16} />
                        <input 
                            type="text" placeholder="Search email, name or role..." value={searchQuery}
                            onChange={(e) => {setSearchQuery(e.target.value); fetchRoster(e.target.value);}}
                            className="pl-12 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl outline-none focus:ring-4 focus:ring-blue-500/10 w-80 transition-all font-bold text-sm shadow-sm"
                        />
                    </div>
                    <button onClick={() => fileInputRef.current.click()} className="flex items-center gap-2 px-6 py-2.5 rounded-xl font-black bg-slate-900 text-white hover:bg-slate-800 transition-all shadow-lg text-[10px] uppercase tracking-widest">
                        <Upload size={14} /> Import List
                    </button>
                    <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".xlsx,.csv"/>
                </div>
            </header>

            <div className="grid grid-cols-3 gap-6">
                <StatCard label="Active Staff" value={activeMembers} icon={<CheckCircle2 className="text-emerald-500" size={20}/>} />
                <StatCard label="Unique Roles" value={uniqueRolesCount} icon={<Briefcase className="text-blue-500" size={20}/>} />
                <StatCard label="Total Teams" value={uniqueTeamsCount} icon={<Layers className="text-purple-500" size={20}/>} />
            </div>

            <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-200 overflow-hidden">
              <table className="w-full text-left table-fixed">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-100">
                    <th className="p-6 font-black text-[10px] uppercase tracking-widest text-slate-400 w-1/3">Staff Identity</th>
                    <th className="p-6 font-black text-[10px] uppercase tracking-widest text-slate-400 w-1/3">Role & Team Path</th>
                    <th className="p-6 font-black text-[10px] uppercase tracking-widest text-slate-400 w-1/4">Manager (Lookup)</th>
                    <th className="p-6 font-black text-[10px] uppercase tracking-widest text-slate-400 text-right w-24">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {roster.map((m, i) => (
                    <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                      <td className="p-6 py-3">
                        <div className="font-black text-slate-800 text-sm truncate uppercase tracking-tight">{m.EmployeeName}</div>
                        <div className="text-[10px] text-slate-400 font-bold truncate lowercase mt-0.5">{m.EmployeeEmail}</div>
                      </td>
                      <td className="p-6 py-3">
                        <div className="font-black text-blue-600 text-[10px] uppercase tracking-tighter truncate leading-none mb-1">{m.Role}</div>
                        <div className="text-[10px] font-black text-slate-500 truncate leading-none uppercase tracking-tighter italic">{m.Region} / {m.SubRegion}</div>
                        <div className="text-[8px] font-black text-slate-300 uppercase mt-1 leading-none tracking-widest">Offs: {m.WeekOffs || 'None'}</div>
                      </td>
                      <td className="p-6 py-3">
                          <div className="flex items-center gap-1.5 font-bold text-slate-400 text-[10px] truncate italic lowercase">
                              <Mail size={10} className="text-slate-200 shrink-0" />
                              <span className="truncate">{m.Manager || 'Pending Sync'}</span>
                          </div>
                      </td>
                      <td className="p-6 py-3 text-right">
                        <span className={`px-2.5 py-1 rounded-lg text-[8px] font-black uppercase tracking-widest border ${m.Include ? 'bg-emerald-50 text-emerald-600 border-emerald-100' : 'bg-slate-50 text-slate-400 border-slate-100'}`}>
                            {m.Include ? 'Active' : 'Archived'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {roster.length === 0 && <div className="p-20 text-center text-slate-400 font-bold italic uppercase text-xs tracking-widest">Registry empty matching search.</div>}
            </div>
          </div>
        )}

        {activeTab === 'matrix' && (
             <div className="p-10 space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-500 max-w-[1400px] mx-auto">
                <header className="flex justify-between items-center border-b border-slate-200 pb-8 text-uppercase">
                    <div>
                        <h2 className="text-3xl font-black tracking-tighter text-slate-900 uppercase italic">Ownership Matrix</h2>
                        <p className="text-slate-400 font-bold text-xs uppercase tracking-widest mt-1 italic">Map managers to Role + Team combinations from Roster.</p>
                    </div>
                    <div className="flex gap-3">
                        <div className="relative group">
                            <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" size={16} />
                            <input 
                                type="text" placeholder="Filter combinations..." value={matrixSearch}
                                onChange={(e) => {setMatrixSearch(e.target.value); fetchMatrix(e.target.value);}}
                                className="pl-12 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl outline-none focus:ring-4 focus:ring-blue-500/10 w-80 transition-all font-bold text-sm shadow-sm"
                            />
                        </div>
                        <button onClick={saveMatrix} className="px-8 py-2.5 bg-slate-900 text-white rounded-xl font-black hover:bg-slate-800 transition-all shadow-xl shadow-slate-900/10 text-[10px] uppercase tracking-widest">
                            Save Ownership
                        </button>
                    </div>
                </header>

                <div className="bg-blue-50 border border-blue-100 p-6 rounded-[2rem] flex gap-4 items-start">
                    <AlertCircle className="text-blue-600 shrink-0" size={20} />
                    <p className="text-[11px] font-bold text-blue-900 leading-relaxed uppercase tracking-tighter">
                        This list is <b>Auto-Generated</b> from your roster. Assign a Manager Email to a group, and every person in that group will be updated. Use the eye icon to exclude groups from all reports.
                    </p>
                </div>

                <div className="bg-white rounded-[3rem] shadow-2xl border border-slate-200 overflow-hidden">
                    <table className="w-full text-left border-collapse">
                        <thead>
                            <tr className="bg-slate-50 border-b border-slate-100 uppercase">
                                <th className="p-6 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Role Path</th>
                                <th className="p-6 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Team Path (Region — Sub)</th>
                                <th className="p-6 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Manager Email</th>
                                <th className="p-6 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400 text-right w-32">Visibility</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                            {matrix.map((row, i) => (
                                <tr key={i} className={`group ${!row.IncludeInReporting ? 'bg-slate-50/50' : ''}`}>
                                    <td className="p-6 py-4">
                                        <div className="font-black text-slate-800 text-xs uppercase tracking-tight">{row.Role || 'Generic'}</div>
                                    </td>
                                    <td className="p-6 py-4">
                                        <div className="font-bold text-slate-400 text-[10px] uppercase italic tracking-tighter">{row.Region} — {row.SubRegion}</div>
                                    </td>
                                    <td className="p-6 py-4">
                                        <div className="relative group/input max-w-xs">
                                            <Mail className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within/input:text-blue-500 transition-colors" size={12}/>
                                            <input 
                                                type="text" 
                                                placeholder="Assign manager..."
                                                value={row.ManagerEmail} 
                                                onChange={(e) => {const n=[...matrix]; n[i].ManagerEmail=e.target.value; setMatrix(n);}} 
                                                className={`w-full pl-9 pr-4 py-2 bg-slate-50/50 border border-slate-100 rounded-lg outline-none focus:bg-white focus:ring-2 focus:ring-blue-500 font-bold text-xs transition-all ${row.ManagerEmail ? 'text-blue-600' : 'text-slate-400 italic'}`}
                                            />
                                        </div>
                                    </td>
                                    <td className="p-6 py-4 text-right">
                                        <button 
                                            onClick={() => {const n=[...matrix]; n[i].IncludeInReporting=!n[i].IncludeInReporting; setMatrix(n);}}
                                            className={`p-2.5 rounded-xl transition-all border ${row.IncludeInReporting ? 'bg-white text-blue-600 border-blue-100 hover:bg-blue-50' : 'bg-slate-50 text-slate-300 border-slate-200 hover:bg-slate-100'}`}
                                        >
                                            {row.IncludeInReporting ? <Eye size={16}/> : <EyeOff size={16}/>}
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                    {matrix.length === 0 && (
                        <div className="p-20 text-center text-slate-400 font-black uppercase text-[10px] tracking-widest italic opacity-50">No path combinations found in roster.</div>
                    )}
                </div>
             </div>
        )}
      </main>
    </div>
  );
}

function NavItem({ icon, label, active, onClick }) {
  return (
    <button 
      onClick={onClick}
      className={`w-full flex items-center gap-3 px-5 py-3 rounded-xl font-black text-[10px] uppercase tracking-widest transition-all ${active ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'text-slate-500 hover:text-white hover:bg-white/5'}`}
    >
      {icon} {label}
    </button>
  );
}

function StepItem({ label, active, done }) {
    return (
        <div className={`flex items-center gap-2 p-3 rounded-2xl transition-all border ${done ? 'bg-emerald-50 border-emerald-100' : active ? 'bg-blue-50 border-blue-100' : 'bg-slate-50 border-slate-100 opacity-40'}`}>
            <div className={`w-6 h-6 rounded-full flex items-center justify-center font-black text-[10px] ${done ? 'bg-emerald-600 text-white' : active ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-500'}`}>
                {done ? <CheckCircle2 size={12}/> : <ChevronRight size={12}/>}
            </div>
            <span className={`text-[10px] font-black uppercase tracking-widest ${done ? 'text-emerald-700' : active ? 'text-blue-700' : 'text-slate-500'}`}>{label}</span>
        </div>
    );
}

function FilterCard({ label, children, icon }) {
  return (
    <div className="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm hover:shadow-lg hover:shadow-slate-200/40 transition-all group">
      <div className="flex items-center gap-1.5 text-slate-400 mb-1.5 font-black text-[8px] uppercase tracking-[0.2em] group-focus-within:text-blue-600 transition-colors">
        {icon} <span>{label}</span>
      </div>
      <div className="text-slate-900 leading-none">{children}</div>
    </div>
  );
}

function StatCard({ label, value, icon }) {
    return (
      <div className="bg-white p-5 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4">
        <div className="w-10 h-10 bg-slate-50 rounded-2xl flex items-center justify-center text-slate-400 border border-slate-100 shrink-0">
          {icon}
        </div>
        <div className="min-w-0">
          <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest truncate">{label}</p>
          <p className="text-xl font-black text-slate-900 tracking-tighter leading-none mt-1">{value}</p>
        </div>
      </div>
    );
}
