// View Mode + Tech Mode + Empty + Error screens
const { useState: uS, useMemo: uM, useEffect: uE } = React;
const II = window.Icon;

const ItemCard = ({ item, mode, onChange, hazardTone }) => {
  const stripe = item.hazard && hazardTone === 'stripe';
  const band = item.hazard && hazardTone === 'band';
  const dedicated = item.hazard && hazardTone === 'section';
  return (
    <article className={`card-flat rounded-2xl border bg-white p-3.5 ${stripe ? 'hazard-stripe border-red-200' : 'border-slate-200'} ${band ? 'hazard-band' : ''}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-1.5 flex-wrap">
            <h3 className="font-display text-[15px] font-semibold text-slate-900 leading-tight">{item.name}</h3>
            {item.hazard && (hazardTone === 'badge' || hazardTone === 'diamond') && (
              hazardTone === 'diamond'
                ? <span title="GHS hazard" className="inline-flex items-center justify-center w-5 h-5 bg-red-600 text-white rotate-45 rounded-sm"><span className="-rotate-45"><II name="hazard" size={11} stroke={2.5}/></span></span>
                : <window.HazardPill/>
            )}
          </div>
          <div className="mt-1 text-[11.5px] text-slate-500 flex flex-wrap gap-x-3 gap-y-0.5">
            <span className="font-mono">{item.id}</span>
            <span>· {item.category}</span>
            {item.unit && <span>· {item.unit}</span>}
          </div>
          {item.remarks && <div className="mt-2 text-[12.5px] text-slate-600"><span className="font-medium text-slate-800">Remarks:</span> {item.remarks}</div>}
          {item.hazard && hazardTone !== 'section' && (
            <div className="mt-2 rounded-lg bg-red-50 border border-red-100 px-2.5 py-1.5 text-[11.5px] text-red-800 flex gap-1.5">
              <II name="shield" size={13} className="shrink-0 mt-0.5"/>
              <span>Handle per D&T chemical storage and safety procedures.</span>
            </div>
          )}
        </div>

        {mode === 'view' ? (
          <div className="shrink-0 text-right rounded-xl bg-slate-50 border border-slate-200 px-3 py-2 min-w-[78px]">
            <div className="text-[10px] uppercase tracking-wider text-slate-500 font-semibold">Expected</div>
            <div className="font-display text-[24px] font-semibold leading-none mt-0.5 text-slate-900">{item.qty}</div>
            {item.unit && <div className="text-[10.5px] text-slate-500 mt-0.5">{item.unit}</div>}
            <div className="mt-1.5 flex justify-end"><window.StatusPill status={item.status}/></div>
          </div>
        ) : (
          <TechControls item={item} onChange={onChange}/>
        )}
      </div>
    </article>
  );
};

const TechControls = ({ item, onChange }) => {
  return (
    <div className="shrink-0 w-[160px] flex flex-col gap-2">
      <div>
        <div className="text-[10px] uppercase tracking-wider text-slate-500 font-semibold mb-1">Quantity</div>
        <div className="flex items-center gap-1">
          <button onClick={()=>onChange(item.id, { qty: Math.max(0, item.qty - 1), changed: true })} className="tap rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 flex items-center justify-center"><II name="minus" size={16}/></button>
          <input className="qty-input input flex-1 px-1" type="number" min="0" value={item.qty} onChange={e=>onChange(item.id, { qty: Math.max(0, Number(e.target.value)||0), changed: true })}/>
          <button onClick={()=>onChange(item.id, { qty: item.qty + 1, changed: true })} className="tap rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 flex items-center justify-center"><II name="plus" size={16}/></button>
        </div>
      </div>
      <div>
        <div className="text-[10px] uppercase tracking-wider text-slate-500 font-semibold mb-1">Status</div>
        <select value={item.status} onChange={e=>onChange(item.id, { status: e.target.value, changed: true })} className="select py-2 text-[13px]">
          <option>Good</option>
          <option>Low Stock</option>
          <option>Missing</option>
          <option>Needs Maintenance</option>
        </select>
      </div>
    </div>
  );
};

const ContextStrip = ({ loc, room, mode, onBack, onToggleMode }) => (
  <div className={`px-4 pt-3 pb-3 border-b border-slate-200 ${mode==='tech' ? 'bg-amber-50/40' : 'bg-white'}`}>
    <div className="flex items-center gap-2">
      <button onClick={onBack} className="tap rounded-lg hover:bg-white flex items-center justify-center text-slate-700 -ml-2"><II name="arrow-left" size={20}/></button>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5">
          <span className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-blue-700">{room?.name || `Room ${loc.room}`}</span>
          {room?.rollout && <span className="pill pill-teal" style={{padding:'2px 7px',fontSize:'10px'}}>419A</span>}
        </div>
        <h1 className="font-display text-[18px] font-semibold text-slate-900 leading-tight truncate flex items-center gap-1.5">
          {loc.label}
          {loc.hazard && <II name="hazard" size={15} className="text-red-600"/>}
        </h1>
        <div className="text-[12px] text-slate-500 mt-0.5 truncate">{loc.specific}</div>
      </div>
      <window.ModeBadge mode={mode}/>
    </div>
    <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
      <span className="pill pill-neutral font-mono text-[10.5px]">ID {loc.id}</span>
      {loc.code && <span className="pill pill-blue font-mono text-[10.5px]">{loc.code}</span>}
      <button onClick={onToggleMode} className="ml-auto pill pill-blue hover:bg-blue-100" style={{padding:'5px 10px'}}>
        <II name={mode==='view'?'edit':'search'} size={11} stroke={2.4}/>
        {mode==='view' ? 'Switch to Tech' : 'Switch to View'}
      </button>
    </div>
  </div>
);

const LocationScreen = ({ data, loc, mode, onBack, onToggleMode, tweaks }) => {
  const [items, setItems] = uS(() => (data.items[loc.id] || []).map(i => ({ ...i, _orig: { qty: i.qty, status: i.status }, changed: false })));
  const [q, setQ] = uS('');
  const [status, setStatus] = uS('All');
  const [saveState, setSaveState] = uS('idle'); // idle | saving | saved | error

  const room = data.rooms.find(r => r.code === loc.room);
  const filtered = items.filter(it => {
    if (status !== 'All' && it.status !== status) return false;
    if (!q) return true;
    return `${it.id} ${it.name} ${it.category} ${it.remarks||''}`.toLowerCase().includes(q.toLowerCase());
  });

  const stats = uM(() => ({
    total: items.length,
    hazard: items.filter(i => i.hazard).length,
    attention: items.filter(i => i.status !== 'Good').length,
  }), [items]);

  const dirty = items.some(i => i.changed);

  const onChange = (id, patch) => setItems(prev => prev.map(it => it.id === id ? { ...it, ...patch } : it));

  const doSave = () => {
    setSaveState('saving');
    setTimeout(() => {
      if (tweaks.forceSaveError) {
        setSaveState('error');
      } else {
        setSaveState('saved');
        setItems(prev => prev.map(i => ({ ...i, _orig: { qty: i.qty, status: i.status }, changed: false })));
        setTimeout(()=>setSaveState('idle'), 2200);
      }
    }, 900);
  };

  // Hazard section sort
  let listItems = filtered;
  if (tweaks.hazardTone === 'section') {
    const haz = filtered.filter(i => i.hazard);
    const rest = filtered.filter(i => !i.hazard);
    listItems = [...haz, ...rest];
  }

  if (items.length === 0) {
    return <EmptyLocation data={data} loc={loc} room={room} mode={mode} onBack={onBack} onToggleMode={onToggleMode}/>;
  }

  return (
    <div className={`min-h-full ${mode==='tech' ? 'mode-tech-bg' : ''}`}>
      <ContextStrip loc={loc} room={room} mode={mode} onBack={onBack} onToggleMode={onToggleMode}/>

      {/* stat row */}
      <div className="px-4 pt-3">
        <div className="flex items-center gap-1.5 flex-wrap text-[11.5px]">
          <span className="pill pill-neutral">{stats.total} item{stats.total!==1?'s':''}</span>
          {stats.hazard > 0 && <span className="pill pill-hazard">{stats.hazard} chemical</span>}
          {stats.attention > 0 && <span className="pill pill-low">{stats.attention} need attention</span>}
        </div>
      </div>

      {/* search + status filter */}
      <div className="px-4 pt-2.5 grid grid-cols-[1fr_auto] gap-2">
        <div className="relative">
          <II name="search" size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"/>
          <input value={q} onChange={e=>setQ(e.target.value)} className="input pl-9 py-2.5 text-[14px]" placeholder="Search items..."/>
        </div>
        <select value={status} onChange={e=>setStatus(e.target.value)} className="select py-2.5 text-[13px] w-[125px]">
          <option>All</option><option>Good</option><option>Low Stock</option><option>Missing</option><option>Needs Maintenance</option>
        </select>
      </div>

      {/* save banners */}
      {saveState !== 'idle' && (
        <div className={`mx-4 mt-3 rounded-xl px-3 py-2.5 text-[12.5px] flex items-center gap-2 ${
          saveState==='saving'? 'bg-blue-50 border border-blue-200 text-blue-900' :
          saveState==='saved' ? 'bg-emerald-50 border border-emerald-200 text-emerald-900' :
          'bg-red-50 border border-red-200 text-red-900'
        }`}>
          {saveState==='saving' && <><div className="w-3.5 h-3.5 border-2 border-blue-600 border-t-transparent rounded-full animate-spin"/>Saving updates…</>}
          {saveState==='saved' && <><II name="check" size={15} stroke={2.5}/>Saved. Inventory rows updated in Google Sheets.</>}
          {saveState==='error' && <><II name="warn" size={15}/>Save failed. Check values and retry.</>}
        </div>
      )}

      {/* item list */}
      <div className={`px-4 pt-3 ${mode==='tech' ? 'pb-32' : 'pb-8'} space-y-2.5`}>
        {tweaks.hazardTone === 'section' && listItems.some(i=>i.hazard) && (
          <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-red-700 flex items-center gap-1.5 pt-1"><II name="hazard" size={12}/>Chemical hazards · pinned</div>
        )}
        {listItems.map((it, idx) => (
          <React.Fragment key={it.id}>
            {tweaks.hazardTone === 'section' && idx > 0 && listItems[idx-1].hazard && !it.hazard && (
              <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-slate-500 pt-1">Other items</div>
            )}
            <div className={it.changed ? 'rounded-2xl row-changed' : ''}>
              <ItemCard item={it} mode={mode} onChange={onChange} hazardTone={tweaks.hazardTone}/>
              {it.changed && (
                <div className="px-3 pb-2 -mt-1 text-[11px] text-amber-800 flex items-center gap-1.5">
                  <II name="edit" size={11}/> Changed · {it._orig.qty} → {it.qty}{it.status !== it._orig.status && ` · ${it._orig.status} → ${it.status}`}
                </div>
              )}
            </div>
          </React.Fragment>
        ))}
        {filtered.length === 0 && (
          <div className="text-center py-8 text-[13px] text-slate-500">No matching items.</div>
        )}
      </div>

      {/* sticky save bar */}
      {mode === 'tech' && (
        <div className="save-bar">
          <div className="text-[12.5px]">
            {dirty
              ? <span className="text-amber-800 font-semibold">{items.filter(i=>i.changed).length} unsaved change{items.filter(i=>i.changed).length>1?'s':''}</span>
              : <span className="text-slate-500">No changes yet</span>}
          </div>
          <button disabled={!dirty || saveState==='saving'} onClick={doSave} className={`btn ${dirty ? 'btn-success' : 'btn-secondary opacity-60 cursor-not-allowed'}`}>
            <II name="save" size={15}/> Save updates
          </button>
        </div>
      )}
    </div>
  );
};

const EmptyLocation = ({ data, loc, room, mode, onBack, onToggleMode }) => (
  <div className="min-h-full">
    <ContextStrip loc={loc} room={room} mode={mode} onBack={onBack} onToggleMode={onToggleMode}/>
    <div className="px-4 pt-8">
      <div className="card rounded-2xl p-6 text-center">
        <div className="mx-auto w-12 h-12 rounded-2xl bg-slate-100 text-slate-500 flex items-center justify-center"><II name="box" size={22}/></div>
        <h3 className="font-display text-[15px] font-semibold text-slate-900 mt-3">No items entered yet</h3>
        <p className="text-[13px] text-slate-500 mt-1">This storage exists in the master list but has no inventory rows.</p>
        {mode === 'tech' && <button className="btn btn-primary mt-4 mx-auto"><II name="plus" size={15}/> Add first item</button>}
      </div>
    </div>
  </div>
);

// ─────────────────────────────────────────────────────────────────────────
// ADMIN READINESS

const AdminReadiness = ({ data, onBack }) => {
  const r = data.readiness;
  const circ = 2 * Math.PI * 44;
  const offset = circ * (1 - r.score / 100);
  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Admin · Readiness"/>
      <div className="px-4 pt-4 pb-8 space-y-4">
        <div className="card rounded-2xl p-4 flex items-center gap-4">
          <svg width="120" height="120" className="gauge shrink-0">
            <circle cx="60" cy="60" r="44" stroke="#e2e8f0"/>
            <circle cx="60" cy="60" r="44" stroke={r.score >= 80 ? '#10b981' : r.score >= 60 ? '#f59e0b' : '#ef4444'} strokeDasharray={circ} strokeDashoffset={offset}/>
          </svg>
          <div>
            <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-slate-500">Data Readiness</div>
            <div className="font-display text-[36px] font-semibold leading-none text-slate-900 mt-1">{r.score}<span className="text-[18px] text-slate-400">/100</span></div>
            <div className="text-[12.5px] text-slate-500 mt-1">{r.errors} error · {r.warnings} warnings · run before rollout</div>
          </div>
        </div>

        <div className="grid grid-cols-2 gap-2">
          <Issue n={r.errors} label="Critical errors" tone="red"/>
          <Issue n={r.warnings} label="Warnings" tone="amber"/>
          <Issue n={r.missingQr} label="Missing QR links" tone="amber"/>
          <Issue n={r.missingStorageId} label="Missing Storage IDs" tone="amber"/>
          <Issue n={r.invalidQty} label="Invalid quantities" tone="slate"/>
          <Issue n={r.duplicates} label="Duplicate rows" tone="amber"/>
        </div>

        <div className="card rounded-2xl p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-display text-[14px] font-semibold text-slate-900">419A rollout panel</h3>
            <span className="pill pill-teal">{r.rollout419A.pct}%</span>
          </div>
          <div className="h-2.5 rounded-full bg-slate-100 overflow-hidden">
            <div className="h-full bg-teal-600 rounded-full" style={{width: `${r.rollout419A.pct}%`}}></div>
          </div>
          <div className="mt-2 text-[12.5px] text-slate-600">{r.rollout419A.ready} of {r.rollout419A.total} 419A storages have valid QR + storage ID + status.</div>
          <div className="mt-3 grid grid-cols-3 gap-2">
            <Metric label="QR ready" value={`${r.rollout419A.ready}/${r.rollout419A.total}`}/>
            <Metric label="Item rows" value={87}/>
            <Metric label="Chemicals" value={6}/>
          </div>
        </div>

        <div className="card rounded-2xl divide-y divide-slate-100">
          <div className="px-4 py-3 text-[12.5px] font-semibold text-slate-700">Diagnostics</div>
          <DiagRow ok label="SPREADSHEET_ID" value="configured"/>
          <DiagRow ok={false} label="WEB_APP_BASE_URL" value="not configured" hint="QR generation disabled"/>
          <DiagRow ok label="Inventory sheet" value="Inventory · 142 rows"/>
          <DiagRow ok label="Required columns" value="all 7 present"/>
          <DiagRow ok={false} label="QR Code Image (optional)" value="missing"/>
        </div>
      </div>
    </div>
  );
};

const Issue = ({ n, label, tone }) => {
  const t = { red:'text-red-700 bg-red-50 border-red-200', amber:'text-amber-700 bg-amber-50 border-amber-200', slate:'text-slate-700 bg-slate-50 border-slate-200' }[tone];
  return (
    <div className={`rounded-2xl border ${t} px-3 py-2.5`}>
      <div className="font-display text-[22px] font-semibold leading-none">{n}</div>
      <div className="text-[11.5px] mt-1 font-medium">{label}</div>
    </div>
  );
};

const DiagRow = ({ ok, label, value, hint }) => (
  <div className="px-4 py-2.5 flex items-center gap-3">
    <span className={`w-2 h-2 rounded-full ${ok?'bg-emerald-500':'bg-amber-500'}`}></span>
    <div className="flex-1 min-w-0">
      <div className="text-[13px] font-medium text-slate-800">{label}</div>
      {hint && <div className="text-[11px] text-slate-500">{hint}</div>}
    </div>
    <div className={`text-[12px] font-mono ${ok?'text-emerald-700':'text-amber-700'}`}>{value}</div>
  </div>
);

// ─────────────────────────────────────────────────────────────────────────
// QR LABEL SHEET

const QrLabels = ({ data, onBack }) => {
  const labels = data.locations.slice(0, 8);
  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="QR Labels · print preview" right={<button className="btn btn-primary py-1.5 px-3 text-[12px] mr-2"><II name="printer" size={14}/> Print</button>}/>
      <div className="px-4 pt-3 pb-2 text-[12px] text-slate-500">Build QR Label Sheet · A4, 2 columns, peel-and-stick</div>
      <div className="px-4 pb-8">
        <div className="label-sheet rounded-xl p-4 grid grid-cols-2 gap-3">
          {labels.map(l => (
            <div key={l.id} className="qr-label">
              <QrGlyph size={86}/>
              <div className="min-w-0 flex-1">
                <div className="text-[10px] font-semibold uppercase tracking-[0.18em] text-blue-700">VSA D&T · {l.room}</div>
                <div className="font-display text-[14px] font-semibold text-slate-900 truncate mt-0.5">{l.label}</div>
                <div className="text-[11px] text-slate-500 truncate">{l.specific}</div>
                <div className="mt-1.5 inline-flex items-center gap-1 px-1.5 py-0.5 bg-slate-900 text-white text-[10px] font-mono rounded">{l.id}</div>
                <div className="text-[10px] text-slate-500 mt-1.5">Scan to view inventory</div>
                {l.hazard && <div className="text-[10px] text-red-700 font-semibold mt-0.5 flex items-center gap-1"><II name="hazard" size={9}/>Chemical cabinet</div>}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────
// ERROR / DIAGNOSTICS

const ErrorScreen = ({ kind, onBack }) => {
  const variants = {
    'no-spreadsheet': { title: 'SPREADSHEET_ID is not set', body: 'The web app cannot read inventory until a spreadsheet is configured.', fix: 'Open the Google Sheet → D&T Inventory menu → Set App Config → paste the Sheet ID.', tone:'red' },
    'no-base-url':    { title: 'WEB_APP_BASE_URL is not configured', body: 'Browsing works, but QR codes cannot be generated and external links will not open from outside the spreadsheet.', fix: 'D&T Inventory → Set WEB_APP_BASE_URL → paste the deployed /exec URL.', tone:'amber' },
    'no-match':       { title: 'No matching storage location', body: 'This QR points to a storage row we cannot find: room or location identifier did not match any inventory rows.', fix: 'Verify the QR was printed from this deployment, or run Refresh QR Links to rebuild the URLs.', tone:'amber' },
    'no-bridge':      { title: 'Save bridge unavailable', body: 'google.script.run is not available in this context. Read-only display still works.', fix: 'Open the deployed /exec URL directly to use Tech Mode.', tone:'amber' },
  };
  const v = variants[kind] || variants['no-spreadsheet'];
  const tones = { red:'border-red-200 bg-red-50 text-red-900', amber:'border-amber-200 bg-amber-50 text-amber-900' };
  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Diagnostics"/>
      <div className="px-4 pt-4 pb-8 space-y-3">
        <div className={`rounded-2xl border ${tones[v.tone]} p-4`}>
          <div className="flex items-start gap-3">
            <II name="warn" size={20} className="shrink-0 mt-0.5"/>
            <div>
              <h3 className="font-display text-[15px] font-semibold">{v.title}</h3>
              <p className="text-[13px] mt-1 opacity-90">{v.body}</p>
            </div>
          </div>
        </div>
        <div className="card rounded-2xl p-4">
          <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-slate-500">How to fix</div>
          <p className="text-[13.5px] text-slate-800 mt-1">{v.fix}</p>
          <div className="mt-3 rounded-lg bg-slate-900 text-slate-100 px-3 py-2 font-mono text-[11.5px]">D&T Inventory → Config Status / Diagnostics</div>
        </div>
        <div className="card rounded-2xl divide-y divide-slate-100">
          <DiagRow ok={false} label="WEB_APP_BASE_URL" value="missing" hint="QR generation disabled"/>
          <DiagRow ok label="SPREADSHEET_ID" value="configured"/>
          <DiagRow ok={false} label="QR Code Image column" value="not present" hint="Optional · only needed for in-sheet QR thumbnails"/>
        </div>
      </div>
    </div>
  );
};

window.LocationScreen = LocationScreen;
window.AdminReadiness = AdminReadiness;
window.QrLabels = QrLabels;
window.ErrorScreen = ErrorScreen;
