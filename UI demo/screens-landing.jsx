// Shared UI atoms + screen components
const { useState, useMemo, useEffect, useRef } = React;
const I = window.Icon;

// ─────────────────────────────────────────────────────────────────────────
// Atoms

const StatusPill = ({ status }) => {
  const m = {
    'Good':              { cls: 'pill-good',  dot: 'dot-good' },
    'Low Stock':         { cls: 'pill-low',   dot: 'dot-low' },
    'Missing':           { cls: 'pill-miss',  dot: 'dot-miss' },
    'Needs Maintenance': { cls: 'pill-maint', dot: 'dot-maint' },
  }[status] || { cls: 'pill-neutral', dot: 'dot-good' };
  return <span className={`pill ${m.cls}`}><span className={`dot ${m.dot}`}></span>{status}</span>;
};

const HazardPill = ({ tone = 'badge' }) => {
  if (tone === 'stripe') return null;
  return <span className="pill pill-hazard"><I name="hazard" size={12} stroke={2.2}/>Chemical</span>;
};

const Metric = ({ label, value, sub, tone }) => (
  <div className={`card-flat rounded-2xl px-3.5 py-3 ${tone === 'rollout' ? 'border-teal-200 bg-[var(--vsa-teal-soft)]' : ''}`}>
    <div className={`text-[10.5px] font-semibold uppercase tracking-[0.16em] ${tone === 'rollout' ? 'text-teal-700' : 'text-slate-500'}`}>{label}</div>
    <div className={`mt-1 text-[22px] font-semibold leading-none font-display ${tone === 'rollout' ? 'text-teal-900' : 'text-slate-900'}`}>{value}</div>
    {sub && <div className="mt-1 text-[11px] text-slate-500">{sub}</div>}
  </div>
);

const QrGlyph = ({ size = 80 }) => (
  <div className="bg-white border border-slate-900 p-1.5" style={{ width: size, height: size }}>
    <div className="qr-glyph w-full h-full"></div>
  </div>
);

// Header that mimics the renderInventoryHtml_ context strip
const PageHeader = ({ title, sub, badge, right, kicker }) => (
  <header className="px-4 pt-4 pb-3 bg-white border-b border-slate-200">
    {kicker && <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-blue-700 mb-1.5">{kicker}</div>}
    <div className="flex items-start justify-between gap-3">
      <div className="min-w-0">
        <h1 className="font-display text-[20px] font-semibold leading-tight text-slate-900 truncate">{title}</h1>
        {sub && <div className="text-[13px] text-slate-500 mt-0.5">{sub}</div>}
      </div>
      {badge}
    </div>
    {right && <div className="mt-2.5">{right}</div>}
  </header>
);

const ModeBadge = ({ mode }) => {
  if (mode === 'update') return <span className="pill" style={{background:'#fffbeb', color:'#92400e', border:'1px solid #fde68a'}}><I name="edit" size={11} stroke={2.4}/>Update Mode</span>;
  return <span className="pill pill-blue"><I name="search" size={11} stroke={2.4}/>View only</span>;
};

const TopBar = ({ onBack, title, right }) => (
  <div className="px-2 py-2 bg-white border-b border-slate-200 flex items-center gap-1">
    <button onClick={onBack} className="tap rounded-lg hover:bg-slate-100 flex items-center justify-center text-slate-700"><I name="arrow-left" size={20}/></button>
    <div className="flex-1 text-[13.5px] font-semibold text-slate-900 truncate">{title}</div>
    {right}
  </div>
);

// ─────────────────────────────────────────────────────────────────────────
// LANDING

const LandingScreen = ({ data, onOpenLocation, onNav, tweaks }) => {
  const [q, setQ] = useState('');
  const [room, setRoom] = useState('All');

  const filtered = useMemo(() => {
    return data.locations.filter(l => {
      if (room !== 'All' && l.room !== room) return false;
      if (!q) return true;
      const hay = `${l.id} ${l.label} ${l.specific} ${l.code} ${l.room}`.toLowerCase();
      return hay.includes(q.toLowerCase());
    });
  }, [q, room, data.locations]);

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach(l => { (g[l.room] = g[l.room] || []).push(l); });
    return g;
  }, [filtered]);

  const totals = useMemo(() => ({
    rooms: data.rooms.length,
    locations: data.locations.length,
    items: data.rooms.reduce((s,r)=>s+r.itemCount,0),
    attention: data.rooms.reduce((s,r)=>s+r.attention,0),
  }), [data]);

  return (
    <div className="min-h-full">
      {/* App header */}
      <div className="bg-white border-b border-slate-200 px-4 pt-4 pb-3">
        <div className="flex items-center justify-between">
          <div>
            <div className="text-[10.5px] font-semibold uppercase tracking-[0.2em] text-blue-700">Victoria Shanghai Academy</div>
            <h1 className="font-display text-[19px] font-semibold text-slate-900 mt-0.5">D&T QR Inventory</h1>
          </div>
          <div className="flex items-center gap-1.5">
            <button className="tap rounded-lg hover:bg-slate-100 flex items-center justify-center text-slate-600" title="Scan QR"><I name="scan" size={20}/></button>
            <button onClick={()=>onNav && onNav('diag')} className="tap rounded-lg hover:bg-slate-100 flex items-center justify-center text-slate-600" title="Diagnostics"><I name="info" size={20}/></button>
          </div>
        </div>
      </div>

      {/* WEB_APP_BASE_URL warning (toggleable via tweaks) */}
      {tweaks.showBaseUrlWarning && (
        <div className="mx-4 mt-3 rounded-xl bg-amber-50 border border-amber-200 px-3 py-2.5 text-[12.5px] text-amber-900 flex gap-2">
          <I name="warn" size={16} className="shrink-0 mt-0.5"/>
          <div>
            <div className="font-semibold">WEB_APP_BASE_URL not set</div>
            <div className="text-amber-800/90 mt-0.5">In-app links work. QR generation is disabled until you set it.</div>
          </div>
        </div>
      )}

      {/* Metrics — 419A hero row, then 3-col summary so labels never wrap on phone */}
      <div className="px-4 pt-4 space-y-2">
        <Metric label="419A rollout" value={`${data.readiness.rollout419A.pct}%`} sub={`${data.readiness.rollout419A.ready}/${data.readiness.rollout419A.total} storages QR-ready`} tone="rollout"/>
        <div className="grid grid-cols-3 gap-2">
          <Metric label="Rooms" value={totals.rooms}/>
          <Metric label="Storages" value={totals.locations}/>
          <Metric label="Items" value={totals.items}/>
        </div>
      </div>

      {/* Search */}
      <div className="px-4 pt-3">
        <div className="relative">
          <I name="search" size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400"/>
          <input value={q} onChange={e=>setQ(e.target.value)} className="input pl-10" placeholder="Search room, storage ID, label, code..."/>
        </div>
      </div>

      {/* Operations tiles */}
      {onNav && (
        <div className="px-4 pt-3">
          <div className="grid grid-cols-2 gap-2">
            <OpsTile icon="grid"  label="Storage Master" hint={`${data.locations.length} storages · ${data.locations.filter(l=>l.qrReady).length} QR-ready`} onClick={()=>onNav('storage')}/>
            <OpsTile icon="warn"  tone="amber" label="Low stock"     hint={`${countLowStock(data)} items · reorder`} onClick={()=>onNav('lowstock')}/>
            <OpsTile icon="cpu"   tone="orange" label="Maintenance"  hint={`${countMaint(data)} machines · ${countChem(data)} chem`} onClick={()=>onNav('maint')}/>
            <OpsTile icon="list"  label="Audit log"     hint={`${data.audit.length} recent events`} onClick={()=>onNav('audit')}/>
            <OpsTile icon="shield" label="Readiness"    hint={`Score ${data.readiness.score} · ${data.readiness.errors} errors`} onClick={()=>onNav('admin')}/>
            <OpsTile icon="printer" label="QR labels"   hint="Print sheet · A4" onClick={()=>onNav('labels')}/>
          </div>
        </div>
      )}

      {/* Room tabs */}
      <div className="px-4 pt-3 pb-1 flex gap-2 overflow-x-auto no-scrollbar">
        <button onClick={()=>setRoom('All')} className={`room-tab ${room==='All'?'is-active':''}`}>All rooms</button>
        {data.rooms.map(r => (
          <button key={r.code} onClick={()=>setRoom(r.code)} className={`room-tab ${r.rollout?'rollout':''} ${room===r.code?'is-active':''}`}>
            {r.code}{r.rollout && ' · 419A rollout'}
          </button>
        ))}
      </div>

      {/* Groups */}
      <div className="px-4 pt-3 pb-24 space-y-5">
        {Object.keys(grouped).length === 0 && (
          <div className="card rounded-2xl p-6 text-center">
            <div className="text-[14px] font-semibold text-slate-900">No matching storage</div>
            <div className="text-[13px] text-slate-500 mt-1">Try a different room, storage ID or location code.</div>
          </div>
        )}
        {data.rooms.map(meta => ({ rcode: meta.code, locs: grouped[meta.code], meta }))
          .filter(x => x.locs && x.locs.length)
          .map(({ rcode, locs, meta }) => {
          return (
            <section key={rcode}>
              <div className="flex items-baseline justify-between mb-2">
                <div className="flex items-center gap-2">
                  <h2 className="font-display text-[14.5px] font-semibold text-slate-900">{meta?.name || rcode}</h2>
                  {meta?.rollout && <span className="pill pill-teal">419A rollout</span>}
                </div>
                <div className="text-[11.5px] text-slate-500">{locs.length} storage</div>
              </div>
              <div className="space-y-2">
                {locs.map(l => <LocationCard key={l.id} loc={l} onOpen={()=>onOpenLocation(l, 'view')} onUpdate={()=>onOpenLocation(l, 'update')}/>)}
              </div>
            </section>
          );
        })}
      </div>
    </div>
  );
};

window.LandingScreen = LandingScreen;

const countLowStock = (data) => {
  let n = 0;
  Object.values(data.items).forEach(list => list.forEach(it => { if (it.status === 'Low Stock' || it.status === 'Missing') n++; }));
  return n;
};
const countMaint = (data) => {
  let n = 0;
  Object.values(data.items).forEach(list => list.forEach(it => { if (it.status === 'Needs Maintenance' || it.maintenanceDue) n++; }));
  return n;
};
const countChem = (data) => {
  let n = 0;
  Object.values(data.items).forEach(list => list.forEach(it => { if (it.hazard) n++; }));
  return n;
};

const OpsTile = ({ icon, label, hint, tone, onClick }) => {
  const toneCls = tone === 'amber' ? 'bg-amber-50 text-amber-800 border-amber-200'
    : tone === 'orange' ? 'bg-orange-50 text-orange-800 border-orange-200'
    : 'bg-slate-100 text-slate-700 border-slate-200';
  return (
    <button onClick={onClick} className="card rounded-xl p-3 text-left hover:bg-slate-50 active:bg-slate-100 transition-colors">
      <div className={`w-8 h-8 rounded-lg flex items-center justify-center border ${toneCls}`}>
        <I name={icon} size={16}/>
      </div>
      <div className="font-display text-[13px] font-semibold text-slate-900 mt-2">{label}</div>
      {hint && <div className="text-[10.5px] text-slate-500 mt-0.5 truncate">{hint}</div>}
    </button>
  );
};

const LocationCard = ({ loc, onOpen, onUpdate }) => {
  const empty = loc.items === 0;
  const attn = loc.attention > 0;
  return (
    <article className="card rounded-2xl overflow-hidden">
      <button onClick={onOpen} className="w-full text-left px-3.5 py-3 flex items-start gap-3 hover:bg-slate-50">
        {/* category icon block */}
        <div className={`shrink-0 w-10 h-10 rounded-xl flex items-center justify-center ${loc.hazard ? 'bg-red-50 text-red-700' : loc.id.startsWith('TOL')? 'bg-amber-50 text-amber-800' : loc.id.startsWith('MAC')? 'bg-blue-50 text-blue-700' : loc.id.startsWith('MAT')? 'bg-emerald-50 text-emerald-800' : 'bg-slate-100 text-slate-700'}`}>
          <I name={loc.hazard ? 'flask' : loc.id.startsWith('TOL') ? 'wrench' : loc.id.startsWith('MAC') ? 'cpu' : loc.id.startsWith('MAT') ? 'box' : 'box'} size={18}/>
        </div>
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1.5 flex-wrap">
            <h3 className="font-display text-[15px] font-semibold text-slate-900 truncate">{loc.label}</h3>
            {loc.hazard && <HazardPill/>}
          </div>
          <div className="text-[12px] text-slate-500 mt-0.5 truncate">{loc.specific}</div>
          <div className="flex items-center gap-1.5 mt-1.5 flex-wrap">
            <span className="pill pill-neutral font-mono text-[10.5px]">{loc.id}</span>
            {loc.code && <span className="pill pill-blue font-mono text-[10.5px]">{loc.code}</span>}
            {empty && <span className="pill pill-neutral">Empty</span>}
            {!loc.qrReady && <span className="pill" style={{background:'#fff7ed',color:'#9a3412'}}>QR missing</span>}
            {attn && <span className="pill pill-low">{loc.attention} attention</span>}
          </div>
        </div>
        <div className="shrink-0 text-right">
          <div className="font-display text-[20px] font-semibold text-slate-900 leading-none">{loc.items}</div>
          <div className="text-[10.5px] text-slate-500 mt-0.5 uppercase tracking-wide">items</div>
        </div>
      </button>
      <div className="border-t border-slate-100 px-2 py-1.5 flex gap-1">
        <button onClick={onOpen} className="flex-1 py-2 rounded-lg text-[12.5px] font-semibold text-slate-700 hover:bg-slate-50 flex items-center justify-center gap-1.5"><I name="search" size={13}/>View</button>
        <div className="w-px bg-slate-100"></div>
        <button onClick={onUpdate} className="flex-1 py-2 rounded-lg text-[12.5px] font-semibold text-blue-700 hover:bg-blue-50 flex items-center justify-center gap-1.5"><I name="edit" size={13}/>Update</button>
      </div>
    </article>
  );
};

window.LandingScreen = LandingScreen;
window.PageHeader = PageHeader;
window.TopBar = TopBar;
window.ModeBadge = ModeBadge;
window.StatusPill = StatusPill;
window.HazardPill = HazardPill;
window.Metric = Metric;
window.QrGlyph = QrGlyph;
