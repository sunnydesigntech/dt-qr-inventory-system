// Main app shell — routing, viewport tweak, navigation between screens
const { useState: uSS, useEffect: uEE } = React;

const App = () => {
  const data = window.DT_DATA;
  const [tweaks, setTweak] = window.useTweaks({
    viewport: 'mobile',         // 'mobile' | 'desktop'
    showBaseUrlWarning: false,
    forceSaveError: false,
    hazardTone: 'badge',        // 'badge' | 'stripe' | 'band' | 'diamond' | 'section'
  });

  const [route, setRoute] = uSS({ name: 'landing' });

  const goLanding = () => setRoute({ name: 'landing' });
  const goLocation = (loc, mode) => setRoute({ name: 'location', loc, mode });
  const toggleMode = () => setRoute(r => ({ ...r, mode: r.mode === 'view' ? 'tech' : 'view' }));

  const screen = (() => {
    if (route.name === 'landing') return <window.LandingScreen data={data} onOpenLocation={goLocation} tweaks={tweaks}/>;
    if (route.name === 'location') return <window.LocationScreen data={data} loc={route.loc} mode={route.mode} onBack={goLanding} onToggleMode={toggleMode} tweaks={tweaks}/>;
    if (route.name === 'admin') return <window.AdminReadiness data={data} onBack={goLanding}/>;
    if (route.name === 'labels') return <window.QrLabels data={data} onBack={goLanding}/>;
    if (route.name === 'error') return <window.ErrorScreen kind={route.kind} onBack={goLanding}/>;
    return null;
  })();

  return (
    <div className="min-h-screen">
      {/* Top toolbar (prototype-only chrome) */}
      <div className="bg-white border-b border-slate-200 px-4 py-2.5 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-slate-900 text-white flex items-center justify-center"><window.Icon name="qr" size={17}/></div>
          <div>
            <div className="font-display text-[13.5px] font-semibold text-slate-900 leading-tight">D&T QR Inventory</div>
            <div className="text-[10.5px] text-slate-500 leading-tight">Victoria Shanghai Academy · Prototype</div>
          </div>
        </div>
        <div className="flex items-center gap-1">
          <NavBtn icon="building" label="Landing" active={route.name==='landing'} onClick={goLanding}/>
          <NavBtn icon="shield" label="Admin" active={route.name==='admin'} onClick={()=>setRoute({name:'admin'})}/>
          <NavBtn icon="printer" label="QR Labels" active={route.name==='labels'} onClick={()=>setRoute({name:'labels'})}/>
          <NavBtn icon="warn" label="Diagnostics" active={route.name==='error'} onClick={()=>setRoute({name:'error', kind:'no-base-url'})}/>
        </div>
      </div>

      {/* Stage */}
      <div className="px-6 py-8 flex items-start justify-center">
        {tweaks.viewport === 'mobile' ? (
          <div className="phone-frame">
            <div className="phone-screen overflow-y-auto">
              <div className="ios-status">
                <span>9:41</span>
                <span className="flex items-center gap-1"><svg width="18" height="11" viewBox="0 0 18 11" fill="currentColor"><rect x="0" y="6" width="3" height="5" rx="0.5"/><rect x="5" y="4" width="3" height="7" rx="0.5"/><rect x="10" y="2" width="3" height="9" rx="0.5"/><rect x="15" y="0" width="3" height="11" rx="0.5"/></svg> <span className="font-semibold text-[13px]">5G</span> <svg width="24" height="11" viewBox="0 0 24 11" fill="none" stroke="currentColor" strokeWidth="1"><rect x="0.5" y="0.5" width="20" height="10" rx="2.5"/><rect x="2" y="2" width="17" height="7" rx="1.2" fill="currentColor"/><path d="M22 4v3"/></svg></span>
              </div>
              {screen}
            </div>
          </div>
        ) : (
          <div className="w-full max-w-[1280px] card rounded-2xl overflow-hidden" style={{minHeight:'780px'}}>
            {screen}
          </div>
        )}
      </div>
    </div>
  );
};

const NavBtn = ({ icon, label, active, onClick }) => (
  <button onClick={onClick} className={`px-2.5 py-1.5 rounded-lg text-[12px] font-medium flex items-center gap-1.5 ${active ? 'bg-slate-900 text-white' : 'text-slate-600 hover:bg-slate-100'}`}>
    <window.Icon name={icon} size={13}/>{label}
  </button>
);

// Tweaks panel
const TweaksUI = () => {
  const [tweaks, setTweak] = window.useTweaks({
    viewport: 'mobile',
    showBaseUrlWarning: false,
    forceSaveError: false,
    hazardTone: 'badge',
  });
  const T = window.TweaksPanel, S = window.TweakSection, Rd = window.TweakRadio, Tg = window.TweakToggle;
  return (
    <T title="Tweaks">
      <S title="Viewport">
        <Rd value={tweaks.viewport} onChange={v=>setTweak('viewport', v)} options={[{value:'mobile',label:'Mobile'},{value:'desktop',label:'Desktop'}]}/>
      </S>
      <S title="Hazard treatment">
        <Rd value={tweaks.hazardTone} onChange={v=>setTweak('hazardTone', v)} options={[
          {value:'badge',label:'Badge'},
          {value:'stripe',label:'Stripe'},
          {value:'band',label:'Band'},
          {value:'diamond',label:'GHS'},
          {value:'section',label:'Section'},
        ]}/>
      </S>
      <S title="States">
        <Tg label="Show WEB_APP_BASE_URL warning" value={tweaks.showBaseUrlWarning} onChange={v=>setTweak('showBaseUrlWarning', v)}/>
        <Tg label="Force save error" value={tweaks.forceSaveError} onChange={v=>setTweak('forceSaveError', v)}/>
      </S>
    </T>
  );
};

// Mount: App owns tweak state via shared hook; TweaksUI renders panel reading the same store.
// useTweaks is a singleton hook so both components share the same persisted state.

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.Fragment>
    <App/>
    <TweaksUI/>
  </React.Fragment>
);
