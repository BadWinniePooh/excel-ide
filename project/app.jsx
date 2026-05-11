/* global React, ReactDOM, Formula */
const { useState, useEffect, useRef, useCallback, useMemo } = React;

// Canonical examples in English; translated to current language when loaded.
const EXAMPLES_EN = [
  { label: "IF · VLOOKUP", src: '=IF(AND(A1>10,B1<20),SUM(C1:C10),IFERROR(VLOOKUP(D1,Sheet1!E:F,2,FALSE),"N/A"))' },
  { label: "LET",          src: '=LET(rate,B1,years,B2,pmt,B3,fv,PV*(1+rate)^years+pmt*((1+rate)^years-1)/rate,ROUND(fv,2))' },
  { label: "SUMPRODUCT",   src: '=SUMPRODUCT((MONTH(Orders[Date])=A1)*(Orders[Region]="EMEA")*Orders[Amount])' },
  { label: "IFS",          src: '=IFS(A1<0,"negative",A1=0,"zero",A1<10,"single digit",A1<100,"two digits",TRUE,"big")' },
];

const DEFAULTS = {
  language: "de",
  indent: 2,
  maxWidth: 60,
  upper: true,
  sep: ";",
};
function decimalFor(lang) { return lang === "de" ? "," : "."; }

function CopyButton({ getText, label = "Copy" }) {
  const [done, setDone] = useState(false);
  const onClick = async () => {
    const t = getText();
    try { await navigator.clipboard.writeText(t); }
    catch (e) {
      const ta = document.createElement("textarea");
      ta.value = t; document.body.appendChild(ta); ta.select();
      document.execCommand("copy"); ta.remove();
    }
    setDone(true); setTimeout(() => setDone(false), 1100);
  };
  return <button className="btn btn-ghost" onClick={onClick}>{done ? "Copied ✓" : label}</button>;
}

function GearIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
         strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1A2 2 0 1 1 4.3 17l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1A2 2 0 1 1 7 4.3l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1A2 2 0 1 1 19.7 7l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z" />
    </svg>
  );
}

function OneLineInput({ value, onChange, decimal, placeholder }) {
  return (
    <div className="oneline-wrap">
      <div className="oneline-scroll">
        <div className="oneline-highlight"
             dangerouslySetInnerHTML={{ __html: Formula.highlight(value || " ", { decimal }) }} />
        <input className="oneline-input" type="text" spellCheck="false"
               autoCorrect="off" autoCapitalize="off"
               placeholder={placeholder} value={value}
               onChange={e => onChange(e.target.value)} />
      </div>
    </div>
  );
}

function PrettyEditor({ value, onChange, decimal, cursor, setCursor }) {
  const taRef = useRef(null);
  const hlRef = useRef(null);
  const lnRef = useRef(null);

  useEffect(() => {
    const ta = taRef.current; if (!ta) return;
    const sync = () => {
      if (hlRef.current) { hlRef.current.scrollTop = ta.scrollTop; hlRef.current.scrollLeft = ta.scrollLeft; }
      if (lnRef.current) { lnRef.current.scrollTop = ta.scrollTop; }
    };
    ta.addEventListener("scroll", sync);
    return () => ta.removeEventListener("scroll", sync);
  }, []);

  const updateCursor = () => {
    const ta = taRef.current; if (!ta) return;
    setCursor(ta.selectionStart);
  };

  const analysis = useMemo(() => {
    try { return Formula.analyze(value || "", cursor, { decimal }); }
    catch (e) { return { bracketAt: null, matchAt: null, diagnostics: [] }; }
  }, [value, cursor, decimal]);

  const errorRanges = useMemo(
    () => (analysis.diagnostics || []).filter(d => d.severity === "error" && d.start != null)
                                       .map(d => ({ start: d.start, end: d.end })),
    [analysis]
  );

  const html = useMemo(
    () => Formula.highlight(value || " ", {
      decimal,
      cursorBracket: analysis.bracketAt,
      matchBracket: analysis.matchAt,
      errorRanges,
    }),
    [value, decimal, analysis, errorRanges]
  );

  const lineCount = useMemo(() => Math.max(1, value.split("\n").length), [value]);
  const lns = []; for (let i = 1; i <= lineCount; i++) lns.push(i);

  return (
    <div className="editor">
      <div className="gutter" ref={lnRef}>
        {lns.map(n => <div key={n} className="gutter-line">{n}</div>)}
      </div>
      <div className="code-area">
        <pre className="code-highlight" ref={hlRef} aria-hidden="true">
          <code dangerouslySetInnerHTML={{ __html: html }} />
        </pre>
        <textarea ref={taRef} className="code-textarea" spellCheck="false"
                  autoCorrect="off" autoCapitalize="off"
                  value={value}
                  onChange={e => { onChange(e.target.value); requestAnimationFrame(updateCursor); }}
                  onKeyUp={updateCursor} onClick={updateCursor} onSelect={updateCursor}
                  onFocus={updateCursor}
                  placeholder="Pretty-printed formula will appear here…" />
      </div>
    </div>
  );
}

function InfoPanel({ analysis, language }) {
  const call = analysis && analysis.activeCall;
  const sig = call ? Formula.getSignature(call.name) : null;
  const argIdx = analysis ? analysis.activeArgIdx : 0;
  const diags = (analysis && analysis.diagnostics) || [];
  const errCount = diags.filter(d => d.severity === "error").length;
  const warnCount = diags.filter(d => d.severity === "warn").length;

  return (
    <aside className="info-panel">
      <div className="ip-section ip-fn">
        {call ? (
          <>
            <div className="ip-fn-head">
              <div className="ip-fn-name">{call.name.toUpperCase()}</div>
              {sig && (sig.en !== call.name.toUpperCase() || sig.de) && (
                <div className="ip-fn-alt">
                  {sig.en !== call.name.toUpperCase() && <span>EN: <code>{sig.en}</code></span>}
                  {sig.de && sig.de !== call.name.toUpperCase() && <span>DE: <code>{sig.de}</code></span>}
                </div>
              )}
            </div>
            {sig ? (
              sig.params.length === 0 ? (
                <div className="ip-empty">No arguments — call as <code>{call.name}()</code></div>
              ) : (
                <ul className="ip-params">
                  {sig.params.map((p, i) => {
                    const isCurrent = i === argIdx
                      || (p.rep && i <= argIdx);
                    return (
                      <li key={i} className={
                        "ip-param" + (isCurrent ? " current" : "")
                                  + (p.r ? " required" : " optional")
                      }>
                        <span className="ip-pointer">{i === argIdx || (p.rep && i === sig.params.length - 1 && argIdx >= i) ? "▸" : ""}</span>
                        <span className="ip-pname">{p.n}</span>
                        <span className="ip-pflag">{p.r ? "required" : "optional"}</span>
                      </li>
                    );
                  })}
                </ul>
              )
            ) : (
              <div className="ip-empty">No signature known for <code>{call.name}</code>. Keep typing — bracket matching and validation still work.</div>
            )}
            {sig && (() => {
              const required = sig.params.filter(p => p.r).length;
              const lastRep = sig.params.length && sig.params[sig.params.length - 1].rep;
              const have = argIdx + 1;
              const stillMissing = Math.max(0, required - have);
              return (
                <div className="ip-progress">
                  <span>Argument <strong>{argIdx + 1}</strong> of {lastRep ? `${required}+` : sig.params.length}</span>
                  {stillMissing > 0 && <span className="ip-missing">· {stillMissing} required still missing</span>}
                </div>
              );
            })()}
          </>
        ) : (
          <div className="ip-fn-idle">
            <div className="ip-fn-name muted">— no function —</div>
            <div className="ip-empty">Place the cursor inside a function call to see its signature.</div>
          </div>
        )}
      </div>

      <div className="ip-section ip-diag">
        <div className="ip-diag-head">
          <span>Diagnostics</span>
          <span className="ip-diag-counts">
            <span className={"badge" + (errCount ? " bad" : " ok")}>{errCount} err</span>
            <span className={"badge" + (warnCount ? " warn" : " ok")}>{warnCount} warn</span>
          </span>
        </div>
        {diags.length === 0 ? (
          <div className="ip-ok">✓ Syntax looks clean</div>
        ) : (
          <ul className="ip-diag-list">
            {diags.slice(0, 20).map((d, i) => (
              <li key={i} className={"ip-diag-item sev-" + d.severity}>
                <span className="ip-diag-dot" aria-hidden="true" />
                <span className="ip-diag-msg">{d.message}</span>
              </li>
            ))}
            {diags.length > 20 && <li className="ip-diag-more">+ {diags.length - 20} more</li>}
          </ul>
        )}
      </div>

      <div className="ip-section ip-tip">
        <div className="ip-tip-title">Tips</div>
        <ul className="ip-tip-list">
          <li>Click any bracket — its partner lights up.</li>
          <li>Type past required args and they're flagged in red.</li>
          <li>{language === "de" ? "Functions are matched against German and English names." : "Both English and German function names are recognised."}</li>
        </ul>
      </div>
    </aside>
  );
}

function SettingsDrawer({ open, onClose, settings, setSetting }) {
  // Trap escape key
  useEffect(() => {
    if (!open) return;
    const h = e => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", h);
    return () => document.removeEventListener("keydown", h);
  }, [open, onClose]);

  return (
    <>
      <div className={"backdrop" + (open ? " open" : "")} onClick={onClose} />
      <aside className={"drawer" + (open ? " open" : "")} aria-hidden={!open}>
        <header className="drawer-head">
          <h2>Settings</h2>
          <button className="icon-btn" onClick={onClose} aria-label="Close settings">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                 strokeWidth="2" strokeLinecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>
          </button>
        </header>

        <div className="drawer-body">
          <section className="set-section">
            <h3>Excel language</h3>
            <p className="set-hint">Formula function names will be translated automatically when you switch.</p>
            <div className="seg">
              <button className={"seg-btn" + (settings.language === "de" ? " on" : "")}
                      onClick={() => setSetting("language", "de")}>
                <span className="flag" aria-hidden="true">🇩🇪</span> German
              </button>
              <button className={"seg-btn" + (settings.language === "en" ? " on" : "")}
                      onClick={() => setSetting("language", "en")}>
                <span className="flag" aria-hidden="true">🇬🇧</span> English
              </button>
            </div>
            <div className="lang-meta">
              <span><strong>Decimal:</strong> <code>{decimalFor(settings.language)}</code></span>
              <span><strong>Dictionary:</strong> {Formula.pairCount}+ functions</span>
            </div>
          </section>

          <section className="set-section">
            <h3>Formatting</h3>

            <div className="set-row">
              <label className="set-label">Indent</label>
              <div className="set-control">
                <input type="range" min="1" max="8" step="1"
                       value={settings.indent}
                       onChange={e => setSetting("indent", parseInt(e.target.value, 10))} />
                <span className="set-value">{settings.indent} {settings.indent === 1 ? "space" : "spaces"}</span>
              </div>
            </div>

            <div className="set-row">
              <label className="set-label">Wrap column</label>
              <div className="set-control">
                <input type="range" min="20" max="140" step="2"
                       value={settings.maxWidth}
                       onChange={e => setSetting("maxWidth", parseInt(e.target.value, 10))} />
                <span className="set-value">{settings.maxWidth} chars</span>
              </div>
            </div>

            <div className="set-row set-row-h">
              <label className="set-label">Uppercase function names</label>
              <button className={"toggle" + (settings.upper ? " on" : "")}
                      role="switch" aria-checked={settings.upper}
                      onClick={() => setSetting("upper", !settings.upper)}>
                <span className="knob" />
              </button>
            </div>

            <div className="set-row">
              <label className="set-label">Argument separator</label>
              <div className="seg seg-sm">
                <button className={"seg-btn" + (settings.sep === "," ? " on" : "")}
                        onClick={() => setSetting("sep", ",")}>
                  Comma <code>,</code>
                </button>
                <button className={"seg-btn" + (settings.sep === ";" ? " on" : "")}
                        onClick={() => setSetting("sep", ";")}>
                  Semicolon <code>;</code>
                </button>
              </div>
            </div>
          </section>

          <section className="set-section">
            <h3>About</h3>
            <p className="set-hint">
              Pure client-side. Nothing is sent anywhere — your formulas never leave the browser.
              Each open tab is its own independent workspace.
            </p>
          </section>
        </div>
      </aside>
    </>
  );
}

function App() {
  const [settings, setSettings] = useState(DEFAULTS);
  const setSetting = useCallback((key, val) => {
    setSettings(s => {
      const next = { ...s, [key]: val };
      // when language changes, also default the separator to the conventional one
      // (but only if it still matches the OLD convention — don't clobber explicit choice).
      if (key === "language") {
        const oldConventional = s.language === "de" ? ";" : ",";
        if (s.sep === oldConventional) next.sep = val === "de" ? ";" : ",";
      }
      return next;
    });
  }, []);

  const initialSrc = useMemo(
    () => Formula.translate(EXAMPLES_EN[0].src, "en", DEFAULTS.language),
    []
  );
  const [oneline, setOneline] = useState(initialSrc);
  const [pretty, setPretty] = useState("");
  const [lastEdited, setLastEdited] = useState("oneline");
  const [showSettings, setShowSettings] = useState(false);
  const [prettyCursor, setPrettyCursor] = useState(0);

  const decimal = decimalFor(settings.language);

  // Re-format when the one-line source changes (only while the user is editing oneline,
  // so typing in the pretty pane doesn't yank the cursor).
  useEffect(() => {
    if (lastEdited !== "oneline") return;
    try {
      setPretty(Formula.format(oneline, {
        indent: settings.indent, maxWidth: settings.maxWidth,
        upper: settings.upper, sep: settings.sep, decimal,
      }));
    } catch (e) {}
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [oneline]);

  // Re-format whenever a formatting *setting* changes — regardless of which pane
  // was edited last. Toggling Uppercase / Indent / Wrap / Separator must always
  // refresh the pretty output.
  useEffect(() => {
    try {
      setPretty(Formula.format(oneline, {
        indent: settings.indent, maxWidth: settings.maxWidth,
        upper: settings.upper, sep: settings.sep, decimal,
      }));
      setLastEdited("oneline");
    } catch (e) {}
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.indent, settings.maxWidth, settings.upper, settings.sep, decimal]);

  // When language changes: translate the current one-line and re-format.
  const prevLang = useRef(settings.language);
  useEffect(() => {
    if (prevLang.current === settings.language) return;
    const translated = Formula.translate(oneline, prevLang.current, settings.language);
    prevLang.current = settings.language;
    setLastEdited("oneline");
    setOneline(translated);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.language]);

  const onOneChange = useCallback(v => { setLastEdited("oneline"); setOneline(v); }, []);
  const onPrettyChange = useCallback(v => {
    setLastEdited("pretty");
    setPretty(v);
    try { setOneline(Formula.minify(v)); } catch (e) {}
  }, []);

  const reformat = () => {
    try {
      setPretty(Formula.format(oneline, {
        indent: settings.indent, maxWidth: settings.maxWidth,
        upper: settings.upper, sep: settings.sep, decimal,
      }));
      setLastEdited("oneline");
    } catch (e) {}
  };

  const loadExample = (i) => {
    const src = Formula.translate(EXAMPLES_EN[i].src, "en", settings.language);
    setLastEdited("oneline");
    setOneline(src);
  };

  const flipLanguage = () => setSetting("language", settings.language === "de" ? "en" : "de");

  const clear = () => { setLastEdited("oneline"); setOneline(""); setPretty(""); };

  const analysis = useMemo(() => {
    try { return Formula.analyze(pretty || "", prettyCursor, { decimal }); }
    catch (e) { return { bracketAt: null, matchAt: null, diagnostics: [] }; }
  }, [pretty, prettyCursor, decimal]);

  const s = useMemo(() => Formula.stats(oneline || "", { decimal }), [oneline, decimal]);

  const placeholder = settings.language === "de"
    ? 'Formel aus Excel einfügen — z.B. =WENN(A1>0;"ja";"nein")'
    : 'Paste a formula from Excel — e.g. =IF(A1>0,"yes","no")';

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark" aria-hidden="true">ƒ<span className="brand-x">x</span></div>
          <div className="brand-text">
            <div className="brand-title">Formula Prettifier</div>
            <div className="brand-sub">Excel · {settings.language === "de" ? "Deutsch" : "English"} · {Formula.pairCount}+ functions translated</div>
          </div>
        </div>
        <div className="top-right">
          <div className="examples">
            <span className="ex-label">Examples</span>
            {EXAMPLES_EN.map((ex, i) => (
              <button key={i} className="chip" onClick={() => loadExample(i)}>{ex.label}</button>
            ))}
          </div>
          <button className="btn btn-translate" onClick={flipLanguage}
                  title={settings.language === "de" ? "Translate to English" : "Translate to German"}>
            {settings.language === "de"
              ? <><span className="flag">🇩🇪</span> → <span className="flag">🇬🇧</span></>
              : <><span className="flag">🇬🇧</span> → <span className="flag">🇩🇪</span></>}
          </button>
          <button className="btn btn-ghost btn-gear" onClick={() => setShowSettings(true)}
                  aria-label="Open settings">
            <GearIcon /> Settings
          </button>
        </div>
      </header>

      <section className="pane pane-one">
        <div className="pane-head">
          <div className="pane-title">
            <span className="pane-num">1</span>
            <span>One-line · paste from Excel</span>
          </div>
          <div className="pane-actions">
            <button className="btn btn-ghost" onClick={clear}>Clear</button>
            <CopyButton getText={() => oneline} label="Copy one-line" />
          </div>
        </div>
        <OneLineInput value={oneline} onChange={onOneChange} decimal={decimal} placeholder={placeholder} />
      </section>

      <div className="pretty-row">
      <section className="pane pane-pretty">
        <div className="pane-head">
          <div className="pane-title">
            <span className="pane-num">2</span>
            <span>Pretty · edit either side, the other syncs live</span>
          </div>
          <div className="pane-actions">
            <button className="btn btn-ghost" onClick={reformat} title="Re-flow indentation">Re-format</button>
            <CopyButton getText={() => pretty} label="Copy pretty" />
          </div>
        </div>
        <PrettyEditor value={pretty} onChange={onPrettyChange} decimal={decimal}
                      cursor={prettyCursor} setCursor={setPrettyCursor} />
      </section>

      <InfoPanel analysis={analysis} language={settings.language} />
      </div>

      <footer className="statusbar">
        <span className={"chip-stat" + (s.balanced ? "" : " bad")}>
          {s.balanced ? "● balanced" : "● unbalanced parens"}
        </span>
        <span className="chip-stat">{s.funcs} func{s.funcs === 1 ? "" : "s"}</span>
        <span className="chip-stat">depth {s.maxDepth}</span>
        <span className="chip-stat">{(oneline || "").length} chars</span>
        <span className="chip-stat">sep <code>{settings.sep}</code> · decimal <code>{decimal}</code></span>
        <span className="spacer" />
        <span className="hint">changes never leave your browser</span>
      </footer>

      <SettingsDrawer open={showSettings} onClose={() => setShowSettings(false)}
                      settings={settings} setSetting={setSetting} />
    </div>
  );
}

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(<App />);
