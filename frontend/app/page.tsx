"use client";

import React, { useState, useEffect, useRef } from 'react';
import { 
  Play, Settings, Users, History, AlertCircle, CheckCircle2, 
  Terminal, ShieldCheck, Calendar as CalendarIcon, Save,
  Search, Upload, X, Filter, RefreshCcw, LayoutDashboard,
  ChevronRight, ArrowRight, Table, Layers
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
    options: { roles: [], teams: [], names: [] }
  });
  const [roster, setRoster] = useState([]);
  const [matrix, setMatrix] = useState([]);
  const [searchQuery, setSearchQuery] = useState('');
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

  const fetchMatrix = async () => {
    try {
      const res = await axios.get(`${API_BASE}/matrix`);
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
      alert("Reporting scope saved!");
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
    } catch (err) { alert("Import failed: " + (err.response?.data?.detail || err.message)); }
  };

  return (
    <div className="min-h-screen bg-[#F1F5F9] text-slate-900 font-sans flex">
      {/* Modern Slim Sidebar */}
      <aside className="w-72 bg-slate-950 text-white h-screen sticky top-0 flex flex-col p-6 shadow-2xl border-r border-white/5">
        <div className="flex items-center gap-4 mb-12 px-2">
          <div className="p-2.5 bg-blue-600 rounded-2xl shadow-lg shadow-blue-500/40">
            <Layers size={24} className="text-white" />
          </div>
          <div>
            <h1 className="text-lg font-black tracking-tight leading-none">DAILY WORK</h1>
            <span className="text-[10px] font-bold text-blue-400 tracking-[0.2em] uppercase">Done Report</span>
          </div>
        </div>

        <nav className="flex-1 space-y-2">
          <NavItem icon={<LayoutDashboard size={20}/>} label="Operations" active={activeTab === 'operations'} onClick={() => setActiveTab('operations')} />
          <NavItem icon={<Users size={20}/>} label="Team Roster" active={activeTab === 'roster'} onClick={() => setActiveTab('roster')} />
          <NavItem icon={<Table size={20}/>} label="Ownership Matrix" active={activeTab === 'matrix'} onClick={() => setActiveTab('matrix')} />
        </nav>
        
        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
            <div className="flex items-center justify-between mb-2">
                <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">System Engine</span>
                <div className={`w-2 h-2 rounded-full ${progress.status === 'running' ? 'bg-blue-500 animate-pulse' : 'bg-emerald-500'}`}></div>
            </div>
            <p className="text-xs font-bold text-slate-300 truncate">{progress.msg}</p>
        </div>
      </aside>

      <main className="flex-1 p-12 overflow-y-auto">
        {activeTab === 'operations' && (
          <div className="max-w-6xl space-y-10">
            <header className="flex justify-between items-start">
              <div>
                <h2 className="text-4xl font-black tracking-tight text-slate-900">Execution Hub</h2>
                <p className="text-slate-500 font-medium mt-1">Configure reporting filters and trigger daily delivery.</p>
              </div>
              <div className="flex gap-3">
                <button onClick={saveConfig} className="px-6 py-3 bg-white border border-slate-200 rounded-2xl font-bold text-slate-600 hover:bg-slate-50 transition-all flex items-center gap-2 shadow-sm">
                  <Save size={18} /> Save Config
                </button>
                <button 
                    onClick={triggerPipeline} 
                    disabled={progress.status === 'running'}
                    className={`px-10 py-3 rounded-2xl font-black text-white transition-all shadow-xl flex items-center gap-3 ${progress.status === 'running' ? 'bg-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 shadow-blue-500/20 active:scale-95 hover:-translate-y-0.5'}`}
                >
                  <Play size={20} fill="white" />
                  {progress.status === 'running' ? 'RUNNING...' : 'LAUNCH PIPELINE'}
                </button>
              </div>
            </header>

            {/* Visual Stepper */}
            <div className="bg-white p-8 rounded-[2.5rem] border border-slate-200 shadow-xl shadow-slate-200/50">
                <div className="flex items-center justify-between mb-8">
                    <h3 className="text-sm font-black text-slate-400 uppercase tracking-[0.2em]">Live Pipeline Progress</h3>
                    <span className="text-xs font-black text-blue-600 bg-blue-50 px-3 py-1 rounded-full uppercase">{progress.value}% Complete</span>
                </div>
                
                <div className="relative h-4 bg-slate-100 rounded-full overflow-hidden mb-12">
                    <div className="absolute top-0 left-0 h-full bg-blue-600 transition-all duration-1000 ease-out" style={{ width: `${progress.value}%` }}></div>
                </div>

                <div className="grid grid-cols-4 gap-4">
                    <StepItem label="Validation" active={progress.value >= 10} done={progress.value > 15} />
                    <StepItem label="Data Sync" active={progress.value >= 30} done={progress.value > 35} />
                    <StepItem label="Aggregation" active={progress.value >= 50} done={progress.value > 55} />
                    <StepItem label="Delivery" active={progress.value >= 70} done={progress.value === 100} />
                </div>
            </div>

            {/* Filter Suite */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
              <FilterCard label="Report Date" icon={<CalendarIcon size={14}/>}>
                 <input type="date" value={config.REPORT_DATE_OVERRIDE} onChange={(e) => setConfig({...config, REPORT_DATE_OVERRIDE: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1"/>
              </FilterCard>
              
              <FilterCard label="Names (CSV)" icon={<Users size={14}/>}>
                <input type="text" placeholder="Search Names..." list="names-list" value={config.NAME_FILTER} onChange={(e) => setConfig({...config, NAME_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1"/>
                <datalist id="names-list">{config.options?.names?.map(n => <option key={n} value={n}/>)}</datalist>
              </FilterCard>

              <FilterCard label="Roles (CSV)" icon={<Settings size={14}/>}>
                <input type="text" placeholder="Search Roles..." list="roles-list" value={config.ROLE_FILTER} onChange={(e) => setConfig({...config, ROLE_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1"/>
                <datalist id="roles-list">{config.options?.roles?.map(r => <option key={r} value={r}/>)}</datalist>
              </FilterCard>

              <FilterCard label="Teams (CSV)" icon={<Filter size={14}/>}>
                <input type="text" placeholder="Search Teams..." list="teams-list" value={config.TEAM_FILTER} onChange={(e) => setConfig({...config, TEAM_FILTER: e.target.value})} className="w-full bg-transparent font-black outline-none mt-1"/>
                <datalist id="teams-list">{config.options?.teams?.map(t => <option key={t} value={t}/>)}</datalist>
              </FilterCard>
            </div>

            {/* Pipeline Controls */}
            <div className="flex gap-6">
                <label className="flex-1 p-6 bg-white rounded-3xl border border-slate-200 shadow-sm flex items-center justify-between cursor-pointer hover:border-blue-300 transition-all">
                    <div>
                        <p className="font-black text-slate-800">DRY RUN MODE</p>
                        <p className="text-xs font-bold text-slate-400">Generate previews without sending emails</p>
                    </div>
                    <input type="checkbox" checked={config.DRY_RUN} onChange={(e) => setConfig({...config, DRY_RUN: e.target.checked})} className="w-6 h-6 accent-blue-600 rounded-lg"/>
                </label>
                <div className="flex-1 p-6 bg-slate-950 rounded-3xl border border-slate-800 shadow-xl flex items-center gap-6 overflow-hidden relative">
                    <Terminal className="text-slate-700 shrink-0" size={40} />
                    <div className="font-mono text-[11px] text-slate-400 truncate">
                        {logs.slice(-1)[0]?.msg || 'Pipeline console ready...'}
                    </div>
                    <div className="absolute right-0 top-0 h-full w-24 bg-gradient-to-l from-slate-950 to-transparent"></div>
                </div>
            </div>
          </div>
        )}

        {activeTab === 'roster' && (
          <div className="space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <header className="flex justify-between items-center">
                <div>
                    <h2 className="text-4xl font-black tracking-tight text-slate-900">Roster Manager</h2>
                    <p className="text-slate-500 font-medium mt-1">Official registry of team members and metadata.</p>
                </div>
                <div className="flex gap-4">
                    <div className="relative group">
                        <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" size={18} />
                        <input 
                            type="text" placeholder="Search staff registry..." value={searchQuery}
                            onChange={(e) => {setSearchQuery(e.target.value); fetchRoster(e.target.value);}}
                            className="pl-14 pr-6 py-4 bg-white border border-slate-200 rounded-2xl outline-none focus:ring-4 focus:ring-blue-500/10 w-96 transition-all font-bold shadow-sm"
                        />
                    </div>
                    <button onClick={() => fileInputRef.current.click()} className="flex items-center gap-3 px-8 py-4 rounded-2xl font-black bg-slate-900 text-white hover:bg-slate-800 transition-all shadow-xl shadow-slate-900/10">
                        <Upload size={20} /> IMPORT EXCEL
                    </button>
                    <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".xlsx,.csv"/>
                </div>
            </header>

            <div className="bg-white rounded-[3rem] shadow-2xl shadow-slate-200/50 border border-slate-200 overflow-hidden">
              <table className="w-full text-left">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-100">
                    <th className="p-8 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Team Member</th>
                    <th className="p-8 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Team Definition</th>
                    <th className="p-8 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400">Ownership</th>
                    <th className="p-8 font-black text-[10px] uppercase tracking-[0.2em] text-slate-400 text-right">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {roster.map((m, i) => (
                    <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                      <td className="p-8">
                        <div className="font-black text-slate-900 text-lg leading-tight">{m.EmployeeName}</div>
                        <div className="text-sm text-slate-400 font-bold">{m.EmployeeEmail}</div>
                      </td>
                      <td className="p-8">
                        <div className="font-black text-slate-600">{m.Region} — {m.SubRegion}</div>
                        <div className="text-[10px] font-black text-blue-500 uppercase mt-1">Offs: {m.WeekOffs || 'None'}</div>
                      </td>
                      <td className="p-8 font-black text-slate-500">{m.Manager}</td>
                      <td className="p-8 text-right">
                        <span className={`px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest ${m.Include ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-400'}`}>
                            {m.Include ? 'Active' : 'Archived'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'matrix' && (
             <div className="space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500">
                <header>
                    <h2 className="text-4xl font-black tracking-tight text-slate-900">Ownership Matrix</h2>
                    <p className="text-slate-500 font-medium mt-1">Define managers for specific Role-Region-SubRegion combinations.</p>
                </header>
                <div className="bg-amber-50 border border-amber-100 p-6 rounded-3xl flex gap-4 items-start">
                    <AlertCircle className="text-amber-500 shrink-0" size={24} />
                    <p className="text-sm font-bold text-amber-800 leading-relaxed">
                        The <b>Manager</b> field in the Roster tab is automatically derived from this matrix. 
                        When you update an email here, the system will re-map all matching employees to the new owner during the next pipeline run.
                    </p>
                </div>
                <div className="bg-white rounded-[3rem] shadow-xl border border-slate-200 p-4">
                    <table className="w-full text-left border-collapse">
                        <thead>
                            <tr className="border-b border-slate-100">
                                <th className="p-6 font-black text-xs uppercase text-slate-400">Role</th>
                                <th className="p-6 font-black text-xs uppercase text-slate-400">Region</th>
                                <th className="p-6 font-black text-xs uppercase text-slate-400">Sub Region</th>
                                <th className="p-6 font-black text-xs uppercase text-slate-400">Owner (Manager Email)</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                            {matrix.map((row, i) => (
                                <tr key={i}>
                                    <td className="p-6 font-black text-slate-800">{row.Role}</td>
                                    <td className="p-6 font-bold text-slate-600">{row.Region}</td>
                                    <td className="p-6 font-bold text-slate-600">{row.SubRegion}</td>
                                    <td className="p-6 font-black text-blue-600 underline underline-offset-4">{row.ManagerEmail}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
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
      className={`w-full flex items-center gap-4 px-6 py-4 rounded-2xl font-black text-xs uppercase tracking-widest transition-all ${active ? 'bg-blue-600 text-white shadow-2xl shadow-blue-600/40' : 'text-slate-500 hover:text-white hover:bg-white/5'}`}
    >
      {icon} {label}
    </button>
  );
}

function StepItem({ label, active, done }) {
    return (
        <div className={`flex items-center gap-3 p-4 rounded-3xl transition-all border ${done ? 'bg-emerald-50 border-emerald-100' : active ? 'bg-blue-50 border-blue-100' : 'bg-slate-50 border-slate-100 opacity-40'}`}>
            <div className={`w-8 h-8 rounded-full flex items-center justify-center font-black text-sm ${done ? 'bg-emerald-600 text-white' : active ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-500'}`}>
                {done ? <CheckCircle2 size={16}/> : <ChevronRight size={16}/>}
            </div>
            <span className={`text-xs font-black uppercase tracking-widest ${done ? 'text-emerald-700' : active ? 'text-blue-700' : 'text-slate-500'}`}>{label}</span>
        </div>
    );
}

function FilterCard({ label, children, icon }) {
  return (
    <div className="bg-white p-6 rounded-[2rem] border border-slate-200 shadow-sm hover:shadow-xl hover:shadow-slate-200/50 transition-all group">
      <div className="flex items-center gap-2 text-slate-400 mb-3 font-black text-[10px] uppercase tracking-[0.2em] group-focus-within:text-blue-600 transition-colors">
        {icon} <span>{label}</span>
      </div>
      <div className="text-slate-900 leading-none">{children}</div>
    </div>
  );
}
