// Additional screens: Diagnostics, Storage Master, Audit Log, Low Stock, Maintenance/Safety
const { useState: uS2, useMemo: uM2 } = React;
const I3 = window.Icon;

// ─────────────────────────────────────────────────────────────────────────
// DIAGNOSTICS — full admin page (distinct from inline ErrorScreen)

const Diagnostics = ({ data, onBack, tweaks }) => {
  const baseUrlOk = !tweaks.showBaseUrlWarning; // tweak inverts: when warning shown → not configured
  const checks = [
    { ok: true, label: 'App mode', value: 'Production' },
    { ok: true, label: 'SPREADSHEET_ID', value: 'configured', critical: true },
    { ok: baseUrlOk, label: 'WEB_APP_BASE_URL', value: baseUrlOk ? 'configured' : 'not configured',
      hint: baseUrlOk ? '' : 'Browsing works, but QR generation needs WEB_APP_BASE_URL.',
      fix: 'D&T Inventory menu → Set WEB_APP_BASE_URL → paste deployed /exec URL.' },
    { ok: true, label: 'Inventory sheet', value: 'Inventory · 142 rows' },
    { ok: true, label: 'Available sheets', value: 'Inventory · Storage_Master · Inventory_Audit · Inventory_Readiness_Report' },
    { ok: true, label: 'Required columns', value: 'Storage ID · Room · Storage Label · Item ID · Item Name · Quantity · Status' },
    { ok: false, label: 'Optional column · QR Code Image', value: 'missing',
      hint: 'Optional. Only needed for in-sheet QR thumbnails.' },
    { ok: true, label: 'Optional column · Safety Note', value: 'present' },
    { ok: true, label: 'Optional column · SDS Link', value: 'present' },
    { ok: true, label: 'Optional column · Reorder Level', value: 'present' },
    { ok: false, label: 'Optional column · Maintenance Due', value: 'partial · 3 of 14 machine rows', hint: 'Recommended for Machines and chemical cabinets.' },
    { ok: true, label: 'Last QR refresh', value: '2026-04-26 11:30 · admin@vsa.edu' },
  ];

  const next = [
    'Set WEB_APP_BASE_URL before printing labels.',
    'Run Refresh QR Links after any storage rename.',
    'Run Create Readiness Report before each rollout window.',
  ];

  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Admin · Diagnostics"/>
      <div className="px-4 pt-4 pb-8 space-y-3">
        {!baseUrlOk && (
          <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-amber-900 flex items-start gap-3">
            <I3 name="warn" size={20} className="shrink-0 mt-0.5"/>
            <div className="flex-1">
              <h3 className="font-display text-[15px] font-semibold">WEB_APP_BASE_URL is not configured</h3>
              <p className="text-[12.5px] mt-1 text-amber-900/90">Browsing works, but QR generation needs WEB_APP_BASE_URL. QR labels should not be printed until this is set.</p>
            </div>
          </div>
        )}

        <section className="card rounded-2xl divide-y divide-slate-100">
          <div className="px-4 py-3 text-[10.5px] font-semibold uppercase tracking-[0.18em] text-slate-500">Configuration health</div>
          {checks.map((c, i) => (
            <div key={i} className="px-4 py-3">
              <div className="flex items-start gap-3">
                <span className={`w-2 h-2 rounded-full mt-1.5 ${c.ok ? 'bg-emerald-500' : c.critical ? 'bg-red-500' : 'bg-amber-500'}`}></span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-baseline justify-between gap-3">
                    <h3 className="text-[13.5px] font-semibold text-slate-900">{c.label}</h3>
                    <span className={`text-[11.5px] font-mono ${c.ok ? 'text-emerald-700' : c.critical ? 'text-red-700' : 'text-amber-700'}`}>{c.value}</span>
                  </div>
                  {c.hint && <p className="text-[11.5px] text-slate-500 mt-0.5">{c.hint}</p>}
                  {c.fix && !c.ok && (
                    <div className="mt-2 rounded-lg bg-slate-50 border border-slate-200 px-2.5 py-1.5 text-[11.5px] text-slate-700">
                      <strong>How to fix · </strong>{c.fix}
                    </div>
                  )}
                </div>
              </div>
            </div>
          ))}
        </section>

        <section className="card rounded-2xl p-4">
          <div className="text-[10.5px] font-semibold uppercase tracking-[0.18em] text-slate-500 mb-2">Next actions</div>
          <ul className="space-y-1.5 text-[12.5px] text-slate-700">
            {next.map((n, i) => (
              <li key={i} className="flex items-start gap-2">
                <span className="w-1.5 h-1.5 rounded-full bg-slate-400 mt-1.5"></span>{n}
              </li>
            ))}
          </ul>
        </section>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────
// STORAGE MASTER / ROLLOUT DASHBOARD

const StorageMaster = ({ data, onBack }) => {
  const [room, setRoom] = uS2('All');
  const list = data.locations.filter(l => room === 'All' || l.room === room);

  const totals = uM2(() => ({
    total: list.length,
    qrReady: list.filter(l => l.qrReady).length,
    scanTested: list.filter(l => l.scanTested).length,
    chemicals: list.filter(l => l.hazard).length,
  }), [list]);

  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Storage Master · Rollout"/>
      <div className="px-4 pt-4 pb-8 space-y-3">
        <div className="grid grid-cols-2 gap-2">
          <window.Metric label="Storages" value={totals.total}/>
          <window.Metric label="QR ready" value={`${totals.qrReady}/${totals.total}`} tone={totals.qrReady === totals.total ? 'rollout' : undefined}/>
          <window.Metric label="Scan-tested" value={`${totals.scanTested}/${totals.total}`}/>
          <window.Metric label="Chemicals" value={totals.chemicals}/>
        </div>

        <div className="flex gap-2 overflow-x-auto no-scrollbar">
          {['All', ...data.rooms.map(r => r.code)].map(r => (
            <button key={r} onClick={()=>setRoom(r)} className={`room-tab ${room === r ? 'is-active' : ''} ${r === '419A' ? 'rollout' : ''}`}>{r === 'All' ? 'All rooms' : r}</button>
          ))}
        </div>

        {/* Mobile: stacked rows. Desktop: table-like grid */}
        <div className="card rounded-2xl divide-y divide-slate-100 overflow-hidden">
          <div className="hidden md:grid md:grid-cols-[1.5fr_0.7fr_1fr_0.7fr_0.7fr_0.7fr_0.9fr] gap-2 px-4 py-2.5 text-[10.5px] font-semibold uppercase tracking-[0.14em] text-slate-500 bg-slate-50">
            <div>Storage</div><div>Room</div><div>Type</div><div>QR</div><div>Items</div><div>Tested</div><div>Last checked</div>
          </div>
          {list.map(l => (
            <div key={l.id} className="px-4 py-3 md:grid md:grid-cols-[1.5fr_0.7fr_1fr_0.7fr_0.7fr_0.7fr_0.9fr] md:gap-2 md:items-center">
              <div className="md:min-w-0">
                <div className="flex items-center gap-1.5 flex-wrap">
                  <h3 className="font-display text-[14px] font-semibold text-slate-900 truncate">{l.label}</h3>
                  {l.hazard && <span className="pill pill-hazard"><I3 name="hazard" size={10}/>Chem</span>}
                </div>
                <div className="text-[11px] text-slate-500 font-mono truncate mt-0.5">{l.id} · {l.code}</div>
              </div>
              <div className="text-[12px] text-slate-600 mt-1.5 md:mt-0">{l.room}</div>
              <div className="text-[12px] text-slate-600 mt-1 md:mt-0 truncate">{l.storageType}</div>
              <div className="mt-1.5 md:mt-0">
                {l.qrReady ? <span className="pill pill-good">Ready</span> : <span className="pill pill-low">Missing</span>}
              </div>
              <div className="text-[12px] text-slate-700 mt-1 md:mt-0 font-mono">{l.items}</div>
              <div className="mt-1 md:mt-0">
                {l.scanTested ? <span className="pill pill-good">Tested</span> : <span className="pill pill-neutral">—</span>}
              </div>
              <div className="text-[11.5px] text-slate-500 mt-1 md:mt-0 font-mono">{l.lastChecked || '—'}</div>
              {l.notes && <div className="md:col-span-7 text-[11.5px] text-slate-500 mt-1.5 md:mt-1.5 md:pt-1.5 md:border-t md:border-slate-100">{l.notes}</div>}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────
// AUDIT LOG

const AuditLog = ({ data, onBack }) => {
  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Operations · Audit Log"/>
      <div className="px-4 pt-4 pb-8 space-y-3">
        <div className="text-[12px] text-slate-500">Most recent {data.audit.length} events. Every save and admin action is logged with timestamp and user.</div>
        <div className="space-y-2">
          {data.audit.map((e, i) => {
            const isUpdate = e.action === 'Update';
            const qtyChanged = isUpdate && e.oldQty !== e.newQty;
            const stChanged = isUpdate && e.oldStatus !== e.newStatus;
            return (
              <article key={i} className="card rounded-2xl p-3.5">
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <span className={`pill ${isUpdate ? 'pill-blue' : 'pill-teal'}`}>{e.action}</span>
                      <span className="font-mono text-[11px] text-slate-500">{e.ts}</span>
                    </div>
                    <div className="mt-1.5 text-[13px] text-slate-800">
                      <span className="font-semibold">{e.user}</span>
                      <span className="text-slate-500"> · {e.room} / {e.storage}{e.item !== '—' ? ` · ${e.item}` : ''}</span>
                    </div>
                    {(qtyChanged || stChanged) && (
                      <div className="mt-1.5 text-[12px] text-slate-700 flex flex-wrap gap-x-3 gap-y-0.5 font-mono">
                        {qtyChanged && <span><span className="text-slate-400">qty</span> {e.oldQty} → <span className="font-semibold">{e.newQty}</span></span>}
                        {stChanged && <span><span className="text-slate-400">status</span> {e.oldStatus} → <span className="font-semibold">{e.newStatus}</span></span>}
                      </div>
                    )}
                    {e.notes && <p className="text-[11.5px] text-slate-500 mt-1.5 italic">"{e.notes}"</p>}
                  </div>
                </div>
              </article>
            );
          })}
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────
// LOW STOCK / REORDER

const LowStock = ({ data, onBack }) => {
  // Flatten all items with reorder info
  const rows = [];
  Object.entries(data.items).forEach(([sid, list]) => {
    list.forEach(it => {
      if (it.status === 'Low Stock' || it.status === 'Missing') {
        const loc = data.locations.find(l => l.id === sid);
        const priority = it.status === 'Missing' ? 'High' : (it.qty <= 1 ? 'High' : 'Medium');
        rows.push({ ...it, storageId: sid, room: loc?.room || '?', storageLabel: loc?.label || sid, priority });
      }
    });
  });

  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Operations · Low Stock"/>
      <div className="px-4 pt-4 pb-8 space-y-3">
        <div className="grid grid-cols-3 gap-2">
          <window.Metric label="Reorder" value={rows.length}/>
          <window.Metric label="High" value={rows.filter(r => r.priority === 'High').length}/>
          <window.Metric label="Chemicals" value={rows.filter(r => r.hazard).length}/>
        </div>

        {rows.length === 0 ? (
          <div className="card rounded-2xl p-6 text-center">
            <div className="text-[14px] font-semibold text-slate-900">No low stock items</div>
            <div className="text-[12.5px] text-slate-500 mt-1">All inventory is at expected levels.</div>
          </div>
        ) : (
          <div className="space-y-2">
            {rows.map((r, i) => (
              <article key={i} className={`card rounded-2xl p-3.5 ${r.hazard ? 'border-red-200' : ''}`}>
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <h3 className="font-display text-[14.5px] font-semibold text-slate-900">{r.name}</h3>
                      {r.hazard && <window.HazardPill/>}
                      <span className={`pill ${r.priority === 'High' ? 'pill-miss' : 'pill-low'}`}>{r.priority}</span>
                    </div>
                    <div className="text-[11.5px] text-slate-500 mt-1 flex flex-wrap gap-x-3 gap-y-0.5">
                      <span className="font-mono">{r.id}</span>
                      <span>· {r.room} / {r.storageLabel}</span>
                    </div>
                    <div className="mt-2 grid grid-cols-2 gap-2 text-[11.5px]">
                      <div className="rounded-lg bg-slate-50 border border-slate-200 px-2.5 py-1.5">
                        <div className="text-slate-500">On hand</div>
                        <div className="font-mono font-semibold text-slate-900">{r.qty} {r.unit || ''}</div>
                      </div>
                      <div className="rounded-lg bg-slate-50 border border-slate-200 px-2.5 py-1.5">
                        <div className="text-slate-500">Reorder at</div>
                        <div className="font-mono font-semibold text-slate-900">{r.reorderLevel != null ? r.reorderLevel : '—'}</div>
                      </div>
                      {r.supplier && (
                        <div className="rounded-lg bg-slate-50 border border-slate-200 px-2.5 py-1.5 col-span-2">
                          <div className="text-slate-500">Supplier</div>
                          <div className="font-medium text-slate-800">{r.supplier}{r.cost ? ` · ${r.cost}` : ''}</div>
                        </div>
                      )}
                    </div>
                    {r.purchaseLink && <a href={r.purchaseLink} className="mt-2.5 inline-flex items-center gap-1.5 text-[12px] font-semibold text-blue-700 hover:underline"><I3 name="external" size={12}/>Open purchase link</a>}
                  </div>
                  <div className="shrink-0">
                    <window.StatusPill status={r.status}/>
                  </div>
                </div>
              </article>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────
// MAINTENANCE / SAFETY

const Maintenance = ({ data, onBack }) => {
  const maint = [];
  const chems = [];
  Object.entries(data.items).forEach(([sid, list]) => {
    const loc = data.locations.find(l => l.id === sid);
    list.forEach(it => {
      if (it.status === 'Needs Maintenance' || it.maintenanceDue) {
        maint.push({ ...it, storageId: sid, storageLabel: loc?.label, room: loc?.room });
      }
      if (it.hazard) {
        chems.push({ ...it, storageId: sid, storageLabel: loc?.label, room: loc?.room });
      }
    });
  });

  return (
    <div className="min-h-full">
      <window.TopBar onBack={onBack} title="Operations · Maintenance & Safety"/>
      <div className="px-4 pt-4 pb-8 space-y-4">
        <section>
          <div className="flex items-center justify-between mb-2">
            <h2 className="font-display text-[14px] font-semibold text-slate-900 flex items-center gap-2"><I3 name="cpu" size={14}/>Maintenance · {maint.length}</h2>
          </div>
          {maint.length === 0 ? (
            <div className="card rounded-2xl p-4 text-center text-[12.5px] text-slate-500">No machines flagged for maintenance.</div>
          ) : (
            <div className="space-y-2">
              {maint.map((m, i) => (
                <article key={i} className="card rounded-2xl p-3.5">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0 flex-1">
                      <h3 className="font-display text-[14.5px] font-semibold text-slate-900">{m.name}</h3>
                      <div className="text-[11.5px] text-slate-500 mt-1"><span className="font-mono">{m.id}</span> · {m.room} / {m.storageLabel}</div>
                      {m.remarks && <p className="text-[12.5px] text-slate-700 mt-2">{m.remarks}</p>}
                      <div className="mt-2 flex flex-wrap gap-1.5 items-center">
                        <window.StatusPill status={m.status}/>
                        {m.maintenanceDue && <span className="pill pill-maint"><I3 name="cpu" size={10}/>Due {m.maintenanceDue}</span>}
                        {m.assetValue && <span className="pill pill-neutral">{m.assetValue}</span>}
                      </div>
                    </div>
                  </div>
                </article>
              ))}
            </div>
          )}
        </section>

        <section>
          <div className="flex items-center justify-between mb-2">
            <h2 className="font-display text-[14px] font-semibold text-slate-900 flex items-center gap-2"><I3 name="hazard" size={14} className="text-red-600"/>Chemical &amp; safety · {chems.length}</h2>
          </div>
          {chems.length === 0 ? (
            <div className="card rounded-2xl p-4 text-center text-[12.5px] text-slate-500">No chemical entries.</div>
          ) : (
            <div className="space-y-2">
              {chems.map((c, i) => (
                <article key={i} className="card rounded-2xl p-3.5 border-red-200">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <h3 className="font-display text-[14.5px] font-semibold text-slate-900">{c.name}</h3>
                        <window.HazardPill/>
                      </div>
                      <div className="text-[11.5px] text-slate-500 mt-1"><span className="font-mono">{c.id}</span> · {c.room} / {c.storageLabel}</div>
                      {c.safetyNote && (
                        <div className="mt-2 rounded-lg bg-red-50 border border-red-100 px-2.5 py-1.5 text-[11.5px] text-red-800 flex gap-1.5">
                          <I3 name="shield" size={13} className="shrink-0 mt-0.5"/>
                          <span>{c.safetyNote}</span>
                        </div>
                      )}
                      <div className="mt-2 flex flex-wrap gap-1.5 items-center">
                        <window.StatusPill status={c.status}/>
                        {c.sds ? <a className="pill pill-blue hover:underline" href={c.sds}><I3 name="external" size={10}/>SDS</a> : <span className="pill pill-low">SDS missing</span>}
                      </div>
                    </div>
                  </div>
                </article>
              ))}
            </div>
          )}
        </section>
      </div>
    </div>
  );
};

window.Diagnostics = Diagnostics;
window.StorageMaster = StorageMaster;
window.AuditLog = AuditLog;
window.LowStock = LowStock;
window.Maintenance = Maintenance;
