/* Excel formula tokenizer, parser, formatter, minifier, translator.
 * No dependencies. Exposes window.Formula.
 */
(function () {
  "use strict";

  // ---------- Tokenizer ----------
  // opts:
  //   keepWs : preserve whitespace tokens (default false)
  //   decimal: decimal separator char in NUM literals ('.' or ',', default '.')
  function tokenize(src, opts) {
    opts = opts || {};
    const keepWs = !!opts.keepWs;
    const decimal = opts.decimal || ".";
    const tokens = [];
    const s = src;
    let i = 0;
    function push(type, value, start) {
      tokens.push({ type, value, start, end: i });
    }

    while (i < s.length) {
      const start = i;
      const c = s[i];

      if (c === " " || c === "\t" || c === "\n" || c === "\r") {
        let j = i;
        while (j < s.length && /\s/.test(s[j])) j++;
        if (keepWs) tokens.push({ type: "WS", value: s.slice(i, j), start, end: j });
        i = j; continue;
      }

      if (c === '"') {
        let j = i + 1; let buf = '"';
        while (j < s.length) {
          if (s[j] === '"' && s[j + 1] === '"') { buf += '""'; j += 2; continue; }
          if (s[j] === '"') { buf += '"'; j++; break; }
          buf += s[j]; j++;
        }
        i = j; push("STR", buf, start); continue;
      }

      if (c === "'") {
        let j = i + 1; let buf = "'";
        while (j < s.length) {
          if (s[j] === "'" && s[j + 1] === "'") { buf += "''"; j += 2; continue; }
          if (s[j] === "'") { buf += "'"; j++; break; }
          buf += s[j]; j++;
        }
        i = j; push("SHEETQ", buf, start); continue;
      }

      if (c === "[") {
        let j = i; let depth = 0; let buf = "";
        while (j < s.length) {
          const ch = s[j];
          if (ch === "[") depth++;
          buf += ch; j++;
          if (ch === "]") { depth--; if (depth === 0) break; }
        }
        i = j; push("TABLEREF", buf, start); continue;
      }

      if (c === "(") { i++; push("LPAREN", "(", start); continue; }
      if (c === ")") { i++; push("RPAREN", ")", start); continue; }
      if (c === "{") { i++; push("LBRACE", "{", start); continue; }
      if (c === "}") { i++; push("RBRACE", "}", start); continue; }
      if (c === ";") { i++; push("SEMI", ";", start); continue; }
      if (c === "@") { i++; push("OP", "@", start); continue; }

      // comma: separator unless it's the decimal mark inside a number context (handled in NUM block)
      if (c === ",") {
        // If decimal is ',' AND we're between digits, treat as part of number — but in
        // practice the number block below consumes the entire number including the
        // decimal. So a bare ',' here is always an arg separator.
        i++; push("COMMA", ",", start); continue;
      }

      if (c === "<") {
        if (s[i + 1] === "=") { i += 2; push("OP", "<=", start); continue; }
        if (s[i + 1] === ">") { i += 2; push("OP", "<>", start); continue; }
        i++; push("OP", "<", start); continue;
      }
      if (c === ">") {
        if (s[i + 1] === "=") { i += 2; push("OP", ">=", start); continue; }
        i++; push("OP", ">", start); continue;
      }
      if (c === "=" || c === "+" || c === "-" || c === "*" || c === "/" ||
          c === "^" || c === "&" || c === "%" || c === ":" || c === "!") {
        i++; push("OP", c, start); continue;
      }

      // number — uses configured decimal mark
      if ((c >= "0" && c <= "9") ||
          (c === decimal && s[i + 1] >= "0" && s[i + 1] <= "9")) {
        let j = i; let buf = "";
        while (j < s.length && ((s[j] >= "0" && s[j] <= "9") || s[j] === decimal)) {
          buf += s[j]; j++;
        }
        if (j < s.length && (s[j] === "e" || s[j] === "E")) {
          buf += s[j]; j++;
          if (j < s.length && (s[j] === "+" || s[j] === "-")) { buf += s[j]; j++; }
          while (j < s.length && s[j] >= "0" && s[j] <= "9") { buf += s[j]; j++; }
        }
        i = j; push("NUM", buf, start); continue;
      }

      if (c === "#") {
        let j = i + 1;
        while (j < s.length && /[A-Za-zÄÖÜäöüß0-9_/!?]/.test(s[j])) j++;
        i = j; push("ERR", s.slice(start, i), start); continue;
      }

      if (/[A-Za-z_$\u00A0-\uFFFF]/.test(c)) {
        let j = i; let buf = "";
        while (j < s.length && /[A-Za-z0-9_$.\u00A0-\uFFFF]/.test(s[j])) { buf += s[j]; j++; }
        let k = j;
        while (k < s.length && /\s/.test(s[k])) k++;
        const up = buf.toUpperCase();
        if (s[k] === "(") { i = j; push("FUNC", buf, start); continue; }
        if (up === "TRUE" || up === "FALSE" || up === "WAHR" || up === "FALSCH") {
          i = j; push("BOOL", buf, start); continue;
        }
        i = j; push("REF", buf, start); continue;
      }

      i++;
      push("UNK", c, start);
    }
    return tokens;
  }

  // ---------- Minifier ----------
  function minify(src) {
    let out = ""; const s = src; let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (c === '"') {
        let j = i + 1; out += '"';
        while (j < s.length) {
          if (s[j] === '"' && s[j + 1] === '"') { out += '""'; j += 2; continue; }
          if (s[j] === '"') { out += '"'; j++; break; }
          out += s[j]; j++;
        }
        i = j; continue;
      }
      if (c === "'") {
        let j = i + 1; out += "'";
        while (j < s.length) {
          if (s[j] === "'" && s[j + 1] === "'") { out += "''"; j += 2; continue; }
          if (s[j] === "'") { out += "'"; j++; break; }
          out += s[j]; j++;
        }
        i = j; continue;
      }
      if (c === "[") {
        let j = i; let depth = 0;
        while (j < s.length) {
          const ch = s[j];
          if (ch === "[") depth++;
          out += ch; j++;
          if (ch === "]") { depth--; if (depth === 0) break; }
        }
        i = j; continue;
      }
      if (c === " " || c === "\t" || c === "\n" || c === "\r") { i++; continue; }
      out += c; i++;
    }
    return out;
  }

  // ---------- Group parser ----------
  function parseGroups(tokens) {
    let i = 0;
    function parseArgs(stopType) {
      const args = [{ sep: null, tokens: [] }];
      while (i < tokens.length) {
        const t = tokens[i];
        if (stopType && t.type === stopType) { i++; return args; }
        if (t.type === "COMMA") { i++; args.push({ sep: ",", tokens: [] }); continue; }
        if (t.type === "SEMI")  { i++; args.push({ sep: ";", tokens: [] }); continue; }
        if (t.type === "FUNC") {
          const name = t.value; i++;
          if (tokens[i] && tokens[i].type === "LPAREN") i++;
          const inner = parseArgs("RPAREN");
          args[args.length - 1].tokens.push({ kind: "func", name, items: inner });
          continue;
        }
        if (t.type === "LPAREN") { i++; const inner = parseArgs("RPAREN"); args[args.length - 1].tokens.push({ kind: "paren", items: inner }); continue; }
        if (t.type === "LBRACE") { i++; const inner = parseArgs("RBRACE"); args[args.length - 1].tokens.push({ kind: "brace", items: inner }); continue; }
        if (t.type === "RPAREN" || t.type === "RBRACE") { i++; return args; }
        args[args.length - 1].tokens.push(t); i++;
      }
      return args;
    }
    const root = parseArgs(null);
    return { kind: "root", items: root };
  }

  // ---------- Formatter ----------
  // opts: indent, maxWidth, upper, sep, decimal
  function format(src, options) {
    options = Object.assign(
      { indent: 2, maxWidth: 60, upper: false, sep: ",", decimal: "." },
      options || {});
    const tokens = tokenize(src, { decimal: options.decimal }).filter(t => t.type !== "WS");
    const root = parseGroups(tokens);

    function fnName(n) { return options.upper ? n.toUpperCase() : n; }

    function renderTokenRun(toks) {
      let out = "";
      for (let i = 0; i < toks.length; i++) {
        const t = toks[i];
        const prev = toks[i - 1];
        if (t.type === "OP") {
          const v = t.value;
          if (v === ":" || v === "!" || v === "@" || v === "%") { out += v; continue; }
          const isUnary = (v === "-" || v === "+") &&
            (!prev || prev.type === "OP" || prev.type === "LPAREN" || prev.type === "LBRACE" || prev.type === "COMMA" || prev.type === "SEMI");
          if (isUnary) { out += v; continue; }
          if (v === "=" && !prev) { out += v; continue; }
          out += " " + v + " "; continue;
        }
        out += String(t.value);
      }
      return out.replace(/  +/g, " ").trim();
    }

    function renderArgInline(arg) {
      let out = ""; let run = [];
      const flush = () => { if (run.length) { out += renderTokenRun(run); run = []; } };
      for (const p of arg.tokens) {
        if (p && p.kind) { flush(); out += renderGroupInline(p); }
        else run.push(p);
      }
      flush();
      return out;
    }

    function renderGroupInline(g) {
      if (g.kind === "func") {
        const inner = g.items.map(renderArgInline).join(options.sep);
        return fnName(g.name) + "(" + inner + ")";
      }
      if (g.kind === "paren") {
        const inner = g.items.map(renderArgInline).join(options.sep);
        return "(" + inner + ")";
      }
      if (g.kind === "brace") {
        let o = "{";
        for (let i = 0; i < g.items.length; i++) {
          if (i > 0) o += g.items[i].sep || ",";
          o += renderArgInline(g.items[i]);
        }
        return o + "}";
      }
      let o = "";
      for (let i = 0; i < g.items.length; i++) {
        if (i > 0) o += g.items[i].sep || options.sep;
        o += renderArgInline(g.items[i]);
      }
      return o;
    }

    function lastLineLen(s) {
      const nl = s.lastIndexOf("\n");
      return nl === -1 ? s.length : s.length - nl - 1;
    }

    function renderArg(arg, indent) {
      const inline = renderArgInline(arg);
      if (inline.length + indent <= options.maxWidth) return inline;
      let out = ""; let run = [];
      const flush = () => { if (run.length) { out += renderTokenRun(run); run = []; } };
      for (const p of arg.tokens) {
        if (p && p.kind) {
          flush();
          out += renderGroup(p, indent + lastLineLen(out));
        } else run.push(p);
      }
      flush();
      return out;
    }

    function renderGroup(g, indent) {
      const inline = renderGroupInline(g);
      if (inline.length + indent <= options.maxWidth) return inline;
      if (g.kind === "func" || g.kind === "paren") {
        const head = g.kind === "func" ? fnName(g.name) + "(" : "(";
        const ci = indent + options.indent;
        const pad = " ".repeat(ci);
        const cpad = " ".repeat(indent);

        const isLet = g.kind === "func" && g.name.toUpperCase() === "LET" &&
          g.items.length >= 3 && g.items.length % 2 === 1;
        if (isLet) {
          const lines = [];
          for (let i = 0; i < g.items.length - 1; i += 2) {
            const a = renderArg(g.items[i], ci);
            const b = renderArg(g.items[i + 1], ci + a.length + 2);
            lines.push(a + options.sep + " " + b);
          }
          const last = g.items[g.items.length - 1];
          const lastStr = renderArg(last, ci);
          let body = "";
          for (let i = 0; i < lines.length; i++) body += pad + lines[i] + options.sep + "\n";
          body += pad + lastStr + "\n";
          return head + "\n" + body + cpad + ")";
        }

        let body = "";
        for (let i = 0; i < g.items.length; i++) {
          const argStr = renderArg(g.items[i], ci);
          const sep = i < g.items.length - 1 ? options.sep : "";
          body += pad + argStr + sep + "\n";
        }
        return head + "\n" + body + cpad + ")";
      }
      if (g.kind === "brace") {
        const rows = []; let cur = [];
        for (const it of g.items) {
          if (it.sep === ";") { rows.push(cur); cur = [it]; }
          else cur.push(it);
        }
        rows.push(cur);
        if (rows.length <= 1) return inline;
        const ci = indent + options.indent;
        const pad = " ".repeat(ci); const cpad = " ".repeat(indent);
        let body = "";
        for (let r = 0; r < rows.length; r++) {
          let line = "";
          for (let i = 0; i < rows[r].length; i++) {
            const it = rows[r][i];
            if (i > 0) line += (it.sep || ",");
            line += renderArgInline(it);
          }
          body += pad + line + (r < rows.length - 1 ? ";" : "") + "\n";
        }
        return "{\n" + body + cpad + "}";
      }
      return inline;
    }

    let result = "";
    for (let i = 0; i < root.items.length; i++) {
      if (i > 0) result += options.sep + "\n";
      result += renderArg(root.items[i], 0);
    }
    return result;
  }

  // ---------- Highlighter ----------
  function escapeHtml(s) {
    return s.replace(/[&<>]/g, ch => ch === "&" ? "&amp;" : ch === "<" ? "&lt;" : "&gt;");
  }
  function highlight(src, opts) {
    opts = opts || {};
    const cursorBracket = opts.cursorBracket;   // char position (start) of bracket touched by cursor
    const matchBracket = opts.matchBracket;     // char position of its partner
    const errorRanges = opts.errorRanges || []; // [{start, end}]
    const toks = tokenize(src, { keepWs: true, decimal: opts.decimal || "." });
    const parenStack = []; const braceStack = []; const depthFor = new Map();
    toks.forEach((t, idx) => {
      if (t.type === "LPAREN") { parenStack.push(idx); depthFor.set(idx, parenStack.length - 1); }
      else if (t.type === "RPAREN") {
        const o = parenStack.pop();
        if (o !== undefined) depthFor.set(idx, depthFor.get(o));
        else depthFor.set(idx, -1);
      } else if (t.type === "LBRACE") { braceStack.push(idx); depthFor.set(idx, braceStack.length - 1); }
      else if (t.type === "RBRACE") {
        const o = braceStack.pop();
        if (o !== undefined) depthFor.set(idx, depthFor.get(o));
        else depthFor.set(idx, -1);
      }
    });
    parenStack.forEach(idx => depthFor.set(idx, -1));
    braceStack.forEach(idx => depthFor.set(idx, -1));
    function inError(start) {
      for (const r of errorRanges) if (start >= r.start && start < r.end) return true;
      return false;
    }
    let html = "";
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i]; const v = escapeHtml(t.value);
      const errCls = inError(t.start) ? " tok-error" : "";
      switch (t.type) {
        case "WS": html += v; break;
        case "FUNC": html += `<span class="tok-func${errCls}">${v}</span>`; break;
        case "STR": html += `<span class="tok-str${errCls}">${v}</span>`; break;
        case "SHEETQ": html += `<span class="tok-sheet${errCls}">${v}</span>`; break;
        case "TABLEREF": html += `<span class="tok-tableref${errCls}">${v}</span>`; break;
        case "NUM": html += `<span class="tok-num${errCls}">${v}</span>`; break;
        case "BOOL": html += `<span class="tok-bool${errCls}">${v}</span>`; break;
        case "ERR": html += `<span class="tok-err${errCls}">${v}</span>`; break;
        case "REF": html += `<span class="tok-ref${errCls}">${v}</span>`; break;
        case "OP": html += `<span class="tok-op${errCls}">${v}</span>`; break;
        case "COMMA":
        case "SEMI": html += `<span class="tok-punct${errCls}">${v}</span>`; break;
        case "LPAREN":
        case "RPAREN":
        case "LBRACE":
        case "RBRACE": {
          const d = depthFor.get(i);
          let cls = d === -1 ? "tok-unmatched" : `tok-paren tok-paren-${d % 5}`;
          if (cursorBracket !== undefined && cursorBracket !== null && t.start === cursorBracket) cls += " tok-bracket-cursor";
          if (matchBracket !== undefined && matchBracket !== null && t.start === matchBracket) cls += " tok-bracket-match";
          html += `<span class="${cls}${errCls}">${v}</span>`;
          break;
        }
        default: html += `<span class="tok-unk${errCls}">${v}</span>`;
      }
    }
    return html + "\n";
  }

  function stats(src, opts) {
    opts = opts || {};
    const toks = tokenize(src, { decimal: opts.decimal || "." }).filter(t => t.type !== "WS");
    let depth = 0, maxDepth = 0, funcs = 0, openP = 0;
    for (const t of toks) {
      if (t.type === "FUNC") funcs++;
      if (t.type === "LPAREN") { depth++; openP++; if (depth > maxDepth) maxDepth = depth; }
      else if (t.type === "RPAREN") { depth--; openP--; }
    }
    return { tokens: toks.length, funcs, maxDepth, balanced: openP === 0 };
  }

  // ---------- Translation ----------
  // English (canonical) → German Excel function names.
  // Covers the most common ~230 functions across categories. Names compared
  // case-insensitively; output uses the casing stored here.
  const EN_DE = {
    // Logical
    "IF": "WENN", "IFS": "WENNS", "AND": "UND", "OR": "ODER", "NOT": "NICHT", "XOR": "XODER",
    "IFERROR": "WENNFEHLER", "IFNA": "WENNNV", "SWITCH": "ERSTERWERT", "LET": "LET", "LAMBDA": "LAMBDA",
    // Lookup / reference
    "VLOOKUP": "SVERWEIS", "HLOOKUP": "WVERWEIS", "XLOOKUP": "XVERWEIS", "LOOKUP": "VERWEIS",
    "MATCH": "VERGLEICH", "XMATCH": "XVERGLEICH", "INDEX": "INDEX", "INDIRECT": "INDIREKT",
    "OFFSET": "BEREICH.VERSCHIEBEN", "CHOOSE": "WAHL", "ROW": "ZEILE", "ROWS": "ZEILEN",
    "COLUMN": "SPALTE", "COLUMNS": "SPALTEN", "ADDRESS": "ADRESSE", "AREAS": "BEREICHE",
    "HYPERLINK": "HYPERLINK", "FORMULATEXT": "FORMELTEXT", "GETPIVOTDATA": "PIVOTDATENZUORDNEN",
    // Math
    "SUM": "SUMME", "SUMIF": "SUMMEWENN", "SUMIFS": "SUMMEWENNS", "SUMPRODUCT": "SUMMENPRODUKT",
    "SUMSQ": "QUADRATESUMME", "PRODUCT": "PRODUKT", "QUOTIENT": "QUOTIENT", "MOD": "REST",
    "ABS": "ABS", "SIGN": "VORZEICHEN", "INT": "GANZZAHL", "TRUNC": "KÜRZEN",
    "ROUND": "RUNDEN", "ROUNDUP": "AUFRUNDEN", "ROUNDDOWN": "ABRUNDEN", "MROUND": "VRUNDEN",
    "CEILING": "OBERGRENZE", "FLOOR": "UNTERGRENZE",
    "CEILING.MATH": "OBERGRENZE.MATHEMATIK", "FLOOR.MATH": "UNTERGRENZE.MATHEMATIK",
    "POWER": "POTENZ", "SQRT": "WURZEL", "SQRTPI": "WURZELPI",
    "EXP": "EXP", "LN": "LN", "LOG": "LOG", "LOG10": "LOG10",
    "PI": "PI", "RAND": "ZUFALLSZAHL", "RANDBETWEEN": "ZUFALLSBEREICH", "RANDARRAY": "ZUFALLSMATRIX",
    "GCD": "GGT", "LCM": "KGV", "EVEN": "GERADE", "ODD": "UNGERADE",
    "SIN": "SIN", "COS": "COS", "TAN": "TAN",
    "ASIN": "ARCSIN", "ACOS": "ARCCOS", "ATAN": "ARCTAN", "ATAN2": "ARCTAN2",
    "SINH": "SINHYP", "COSH": "COSHYP", "TANH": "TANHYP",
    "DEGREES": "GRAD", "RADIANS": "BOGENMASS",
    "FACT": "FAKULTÄT", "FACTDOUBLE": "ZWEIFAKULTÄT",
    "COMBIN": "KOMBINATIONEN", "COMBINA": "KOMBINATIONEN2",
    "PERMUT": "VARIATIONEN", "PERMUTATIONA": "VARIATIONEN2",
    "SUBTOTAL": "TEILERGEBNIS", "AGGREGATE": "AGGREGAT",
    // Statistical
    "AVERAGE": "MITTELWERT", "AVERAGEA": "MITTELWERTA",
    "AVERAGEIF": "MITTELWERTWENN", "AVERAGEIFS": "MITTELWERTWENNS",
    "COUNT": "ANZAHL", "COUNTA": "ANZAHL2", "COUNTBLANK": "ANZAHLLEEREZELLEN",
    "COUNTIF": "ZÄHLENWENN", "COUNTIFS": "ZÄHLENWENNS",
    "MAX": "MAX", "MAXA": "MAXA", "MAXIFS": "MAXWENNS",
    "MIN": "MIN", "MINA": "MINA", "MINIFS": "MINWENNS",
    "MEDIAN": "MEDIAN", "MODE": "MODALWERT",
    "MODE.SNGL": "MODUS.EINF", "MODE.MULT": "MODUS.VIELF",
    "LARGE": "KGRÖSSTE", "SMALL": "KKLEINSTE",
    "RANK": "RANG", "RANK.EQ": "RANG.GLEICH", "RANK.AVG": "RANG.MITTELW",
    "STDEV": "STABW", "STDEV.S": "STABW.S", "STDEV.P": "STABW.N",
    "VAR": "VARIANZ", "VAR.S": "VAR.S", "VAR.P": "VAR.P",
    "PERCENTILE": "QUANTIL",
    "PERCENTILE.INC": "QUANTIL.INKL", "PERCENTILE.EXC": "QUANTIL.EXKL",
    "QUARTILE": "QUARTILE", "QUARTILE.INC": "QUARTILE.INKL", "QUARTILE.EXC": "QUARTILE.EXKL",
    "FREQUENCY": "HÄUFIGKEIT", "CORREL": "KORREL", "COVAR": "KOVAR",
    "FORECAST": "SCHÄTZER", "TREND": "TREND",
    "SLOPE": "STEIGUNG", "INTERCEPT": "ACHSENABSCHNITT",
    "GEOMEAN": "GEOMITTEL", "HARMEAN": "HARMITTEL",
    // Text
    "LEFT": "LINKS", "RIGHT": "RECHTS", "MID": "TEIL", "LEN": "LÄNGE",
    "LOWER": "KLEIN", "UPPER": "GROSS", "PROPER": "GROSS2",
    "TRIM": "GLÄTTEN", "CLEAN": "SÄUBERN", "REPT": "WIEDERHOLEN",
    "EXACT": "IDENTISCH", "FIND": "FINDEN", "SEARCH": "SUCHEN",
    "SUBSTITUTE": "WECHSELN", "REPLACE": "ERSETZEN",
    "CONCATENATE": "VERKETTEN", "CONCAT": "TEXTKETTE", "TEXTJOIN": "TEXTVERKETTEN",
    "TEXT": "TEXT", "VALUE": "WERT", "NUMBERVALUE": "ZAHLENWERT",
    "CHAR": "ZEICHEN", "CODE": "CODE", "UNICHAR": "UNIZEICHEN", "UNICODE": "UNICODE",
    "DOLLAR": "DM", "FIXED": "FEST",
    "TEXTBEFORE": "TEXTVOR", "TEXTAFTER": "TEXTNACH", "TEXTSPLIT": "TEXTTEILEN",
    // Date / time
    "TODAY": "HEUTE", "NOW": "JETZT",
    "DATE": "DATUM", "TIME": "ZEIT", "DATEVALUE": "DATWERT", "TIMEVALUE": "ZEITWERT",
    "YEAR": "JAHR", "MONTH": "MONAT", "DAY": "TAG",
    "HOUR": "STUNDE", "MINUTE": "MINUTE", "SECOND": "SEKUNDE",
    "WEEKDAY": "WOCHENTAG", "WEEKNUM": "KALENDERWOCHE", "ISOWEEKNUM": "ISOKALENDERWOCHE",
    "EDATE": "EDATUM", "EOMONTH": "MONATSENDE", "DAYS": "TAGE", "DAYS360": "TAGE360",
    "YEARFRAC": "BRTEILJAHRE", "DATEDIF": "DATEDIF",
    "NETWORKDAYS": "NETTOARBEITSTAGE", "NETWORKDAYS.INTL": "NETTOARBEITSTAGE.INTL",
    "WORKDAY": "ARBEITSTAG", "WORKDAY.INTL": "ARBEITSTAG.INTL",
    // Information
    "ISBLANK": "ISTLEER", "ISNUMBER": "ISTZAHL", "ISTEXT": "ISTTEXT", "ISNONTEXT": "ISTKTEXT",
    "ISERROR": "ISTFEHLER", "ISERR": "ISTFEHL", "ISNA": "ISTNV",
    "ISLOGICAL": "ISTLOG", "ISREF": "ISTBEZUG", "ISFORMULA": "ISTFORMEL",
    "ISEVEN": "ISTGERADE", "ISODD": "ISTUNGERADE",
    "NA": "NV", "N": "N", "TYPE": "TYP", "ERROR.TYPE": "FEHLER.TYP",
    "CELL": "ZELLE", "INFO": "INFO", "SHEET": "BLATT", "SHEETS": "BLÄTTER",
    // Financial
    "PMT": "RMZ", "PPMT": "KAPZ", "IPMT": "ZINSZ",
    "PV": "BW", "FV": "ZW", "NPV": "NBW",
    "IRR": "IKV", "XIRR": "XINTZINSFUSS", "MIRR": "QIKV",
    "NPER": "ZZR", "RATE": "ZINS", "XNPV": "XKAPITALWERT",
    "SLN": "LIA", "DDB": "GDA", "DB": "GDA2", "EFFECT": "EFFEKTIV", "NOMINAL": "NOMINAL",
    // Array / dynamic
    "MAKEARRAY": "MATRIXERSTELLEN",
    "MAP": "MAP", "REDUCE": "REDUCE", "SCAN": "SCAN",
    "BYROW": "NACHZEILE", "BYCOL": "NACHSPALTE",
    "FILTER": "FILTER", "SORT": "SORTIEREN", "SORTBY": "SORTIERENNACH",
    "UNIQUE": "EINDEUTIG", "SEQUENCE": "SEQUENZ", "TRANSPOSE": "MTRANS",
    "TOROW": "INZEILE", "TOCOL": "INSPALTE",
    "WRAPROWS": "ZEILENUMBRUCH", "WRAPCOLS": "SPALTENUMBRUCH",
    "TAKE": "NEHMEN", "DROP": "ENTFERNEN",
    "CHOOSEROWS": "ZEILENWAHL", "CHOOSECOLS": "SPALTENWAHL",
    "HSTACK": "HSTAPELN", "VSTACK": "VSTAPELN", "EXPAND": "ERWEITERN",
  };
  const DE_EN = {};
  Object.keys(EN_DE).forEach(k => { DE_EN[EN_DE[k].toUpperCase()] = k; });

  function translate(src, fromLang, toLang) {
    if (!src || fromLang === toLang) return src;
    const fromDecimal = fromLang === "de" ? "," : ".";
    const toDecimal = toLang === "de" ? "," : ".";
    const fromSep = fromLang === "de" ? ";" : ",";
    const toSep = toLang === "de" ? ";" : ",";
    const dict = fromLang === "en" ? EN_DE : DE_EN;
    const toks = tokenize(src, { keepWs: true, decimal: fromDecimal });
    let out = "";
    for (const t of toks) {
      if (t.type === "FUNC") {
        const up = t.value.toUpperCase();
        out += dict[up] || t.value;
      } else if (t.type === "BOOL") {
        const up = t.value.toUpperCase();
        if (toLang === "de") {
          if (up === "TRUE") out += "WAHR";
          else if (up === "FALSE") out += "FALSCH";
          else out += t.value;
        } else {
          if (up === "WAHR") out += "TRUE";
          else if (up === "FALSCH") out += "FALSE";
          else out += t.value;
        }
      } else if (t.type === "NUM" && fromDecimal !== toDecimal) {
        out += t.value.split(fromDecimal).join(toDecimal);
      } else if (t.type === "COMMA" && fromSep === "," && toSep === ";") {
        out += ";";
      } else if (t.type === "SEMI" && fromSep === ";" && toSep === ",") {
        out += ",";
      } else {
        out += t.value;
      }
    }
    return out;
  }

  // Auto-detect: returns 'de' or 'en' based on which dictionary matches more FUNC tokens.
  function detectLanguage(src) {
    if (!src) return null;
    const toks = tokenize(src).filter(t => t.type === "FUNC");
    let en = 0, de = 0;
    for (const t of toks) {
      const up = t.value.toUpperCase();
      if (EN_DE[up]) en++;
      else if (DE_EN[up]) de++;
    }
    if (en === 0 && de === 0) return null;
    return de > en ? "de" : "en";
  }

  // ---------- Signatures (English canonical) ----------
  const SIGNATURES = {
    IF: [{n:"logical_test",r:1},{n:"value_if_true",r:1},{n:"value_if_false",r:0}],
    IFS: [{n:"logical_test1",r:1},{n:"value_if_true1",r:1},{n:"logical_test…",r:0,rep:1},{n:"value_if_true…",r:0,rep:1}],
    AND: [{n:"logical1",r:1},{n:"logical2…",r:0,rep:1}],
    OR:  [{n:"logical1",r:1},{n:"logical2…",r:0,rep:1}],
    NOT: [{n:"logical",r:1}],
    XOR: [{n:"logical1",r:1},{n:"logical2…",r:0,rep:1}],
    IFERROR: [{n:"value",r:1},{n:"value_if_error",r:1}],
    IFNA: [{n:"value",r:1},{n:"value_if_na",r:1}],
    SWITCH: [{n:"expression",r:1},{n:"value1",r:1},{n:"result1",r:1},{n:"value/result/default…",r:0,rep:1}],
    VLOOKUP: [{n:"lookup_value",r:1},{n:"table_array",r:1},{n:"col_index_num",r:1},{n:"range_lookup",r:0}],
    HLOOKUP: [{n:"lookup_value",r:1},{n:"table_array",r:1},{n:"row_index_num",r:1},{n:"range_lookup",r:0}],
    XLOOKUP: [{n:"lookup_value",r:1},{n:"lookup_array",r:1},{n:"return_array",r:1},{n:"if_not_found",r:0},{n:"match_mode",r:0},{n:"search_mode",r:0}],
    MATCH: [{n:"lookup_value",r:1},{n:"lookup_array",r:1},{n:"match_type",r:0}],
    XMATCH: [{n:"lookup_value",r:1},{n:"lookup_array",r:1},{n:"match_mode",r:0},{n:"search_mode",r:0}],
    INDEX: [{n:"array",r:1},{n:"row_num",r:1},{n:"column_num",r:0},{n:"area_num",r:0}],
    INDIRECT: [{n:"ref_text",r:1},{n:"a1",r:0}],
    OFFSET: [{n:"reference",r:1},{n:"rows",r:1},{n:"cols",r:1},{n:"height",r:0},{n:"width",r:0}],
    CHOOSE: [{n:"index_num",r:1},{n:"value1",r:1},{n:"value…",r:0,rep:1}],
    ROW: [{n:"reference",r:0}],
    ROWS: [{n:"array",r:1}],
    COLUMN: [{n:"reference",r:0}],
    COLUMNS: [{n:"array",r:1}],
    SUM: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    SUMIF: [{n:"range",r:1},{n:"criteria",r:1},{n:"sum_range",r:0}],
    SUMIFS: [{n:"sum_range",r:1},{n:"criteria_range1",r:1},{n:"criteria1",r:1},{n:"range/crit…",r:0,rep:1}],
    SUMPRODUCT: [{n:"array1",r:1},{n:"array…",r:0,rep:1}],
    PRODUCT: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    ABS: [{n:"number",r:1}],
    ROUND: [{n:"number",r:1},{n:"num_digits",r:1}],
    ROUNDUP: [{n:"number",r:1},{n:"num_digits",r:1}],
    ROUNDDOWN: [{n:"number",r:1},{n:"num_digits",r:1}],
    POWER: [{n:"number",r:1},{n:"power",r:1}],
    SQRT: [{n:"number",r:1}],
    MOD: [{n:"number",r:1},{n:"divisor",r:1}],
    INT: [{n:"number",r:1}],
    CEILING: [{n:"number",r:1},{n:"significance",r:1}],
    FLOOR: [{n:"number",r:1},{n:"significance",r:1}],
    AVERAGE: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    AVERAGEIF: [{n:"range",r:1},{n:"criteria",r:1},{n:"average_range",r:0}],
    AVERAGEIFS: [{n:"avg_range",r:1},{n:"crit_range1",r:1},{n:"crit1",r:1},{n:"range/crit…",r:0,rep:1}],
    COUNT: [{n:"value1",r:1},{n:"value…",r:0,rep:1}],
    COUNTA: [{n:"value1",r:1},{n:"value…",r:0,rep:1}],
    COUNTBLANK: [{n:"range",r:1}],
    COUNTIF: [{n:"range",r:1},{n:"criteria",r:1}],
    COUNTIFS: [{n:"crit_range1",r:1},{n:"crit1",r:1},{n:"range/crit…",r:0,rep:1}],
    MAX: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    MIN: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    MAXIFS: [{n:"max_range",r:1},{n:"crit_range1",r:1},{n:"crit1",r:1},{n:"range/crit…",r:0,rep:1}],
    MINIFS: [{n:"min_range",r:1},{n:"crit_range1",r:1},{n:"crit1",r:1},{n:"range/crit…",r:0,rep:1}],
    LARGE: [{n:"array",r:1},{n:"k",r:1}],
    SMALL: [{n:"array",r:1},{n:"k",r:1}],
    MEDIAN: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    STDEV: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    VAR: [{n:"number1",r:1},{n:"number…",r:0,rep:1}],
    LEFT: [{n:"text",r:1},{n:"num_chars",r:0}],
    RIGHT: [{n:"text",r:1},{n:"num_chars",r:0}],
    MID: [{n:"text",r:1},{n:"start_num",r:1},{n:"num_chars",r:1}],
    LEN: [{n:"text",r:1}],
    LOWER: [{n:"text",r:1}],
    UPPER: [{n:"text",r:1}],
    PROPER: [{n:"text",r:1}],
    TRIM: [{n:"text",r:1}],
    CLEAN: [{n:"text",r:1}],
    CONCATENATE: [{n:"text1",r:1},{n:"text…",r:0,rep:1}],
    CONCAT: [{n:"text1",r:1},{n:"text…",r:0,rep:1}],
    TEXTJOIN: [{n:"delimiter",r:1},{n:"ignore_empty",r:1},{n:"text1",r:1},{n:"text…",r:0,rep:1}],
    TEXTSPLIT: [{n:"text",r:1},{n:"col_delim",r:1},{n:"row_delim",r:0},{n:"ignore_empty",r:0},{n:"match_mode",r:0},{n:"pad_with",r:0}],
    SUBSTITUTE: [{n:"text",r:1},{n:"old_text",r:1},{n:"new_text",r:1},{n:"instance_num",r:0}],
    REPLACE: [{n:"old_text",r:1},{n:"start_num",r:1},{n:"num_chars",r:1},{n:"new_text",r:1}],
    FIND: [{n:"find_text",r:1},{n:"within_text",r:1},{n:"start_num",r:0}],
    SEARCH: [{n:"find_text",r:1},{n:"within_text",r:1},{n:"start_num",r:0}],
    TEXT: [{n:"value",r:1},{n:"format_text",r:1}],
    VALUE: [{n:"text",r:1}],
    NUMBERVALUE: [{n:"text",r:1},{n:"decimal_sep",r:0},{n:"group_sep",r:0}],
    REPT: [{n:"text",r:1},{n:"number_times",r:1}],
    EXACT: [{n:"text1",r:1},{n:"text2",r:1}],
    TODAY: [],
    NOW: [],
    DATE: [{n:"year",r:1},{n:"month",r:1},{n:"day",r:1}],
    TIME: [{n:"hour",r:1},{n:"minute",r:1},{n:"second",r:1}],
    YEAR: [{n:"serial_number",r:1}],
    MONTH: [{n:"serial_number",r:1}],
    DAY: [{n:"serial_number",r:1}],
    HOUR: [{n:"serial_number",r:1}],
    MINUTE: [{n:"serial_number",r:1}],
    SECOND: [{n:"serial_number",r:1}],
    WEEKDAY: [{n:"serial_number",r:1},{n:"return_type",r:0}],
    WEEKNUM: [{n:"serial_number",r:1},{n:"return_type",r:0}],
    EDATE: [{n:"start_date",r:1},{n:"months",r:1}],
    EOMONTH: [{n:"start_date",r:1},{n:"months",r:1}],
    DAYS: [{n:"end_date",r:1},{n:"start_date",r:1}],
    DATEDIF: [{n:"start_date",r:1},{n:"end_date",r:1},{n:"unit",r:1}],
    NETWORKDAYS: [{n:"start_date",r:1},{n:"end_date",r:1},{n:"holidays",r:0}],
    WORKDAY: [{n:"start_date",r:1},{n:"days",r:1},{n:"holidays",r:0}],
    DATEVALUE: [{n:"date_text",r:1}],
    ISBLANK: [{n:"value",r:1}],
    ISNUMBER: [{n:"value",r:1}],
    ISTEXT: [{n:"value",r:1}],
    ISNONTEXT: [{n:"value",r:1}],
    ISERROR: [{n:"value",r:1}],
    ISERR: [{n:"value",r:1}],
    ISNA: [{n:"value",r:1}],
    ISLOGICAL: [{n:"value",r:1}],
    ISREF: [{n:"value",r:1}],
    ISEVEN: [{n:"value",r:1}],
    ISODD: [{n:"value",r:1}],
    ISFORMULA: [{n:"reference",r:1}],
    LET: [{n:"name1",r:1},{n:"value1",r:1},{n:"name/value…",r:0,rep:1},{n:"calculation",r:1}],
    LAMBDA: [{n:"parameter…",r:0,rep:1},{n:"calculation",r:1}],
    FILTER: [{n:"array",r:1},{n:"include",r:1},{n:"if_empty",r:0}],
    SORT: [{n:"array",r:1},{n:"sort_index",r:0},{n:"sort_order",r:0},{n:"by_col",r:0}],
    SORTBY: [{n:"array",r:1},{n:"by_array1",r:1},{n:"sort_order1",r:0},{n:"by_array/order…",r:0,rep:1}],
    UNIQUE: [{n:"array",r:1},{n:"by_col",r:0},{n:"exactly_once",r:0}],
    SEQUENCE: [{n:"rows",r:1},{n:"columns",r:0},{n:"start",r:0},{n:"step",r:0}],
    TRANSPOSE: [{n:"array",r:1}],
    RAND: [],
    RANDBETWEEN: [{n:"bottom",r:1},{n:"top",r:1}],
    RANDARRAY: [{n:"rows",r:0},{n:"columns",r:0},{n:"min",r:0},{n:"max",r:0},{n:"integer",r:0}],
    HSTACK: [{n:"array1",r:1},{n:"array…",r:0,rep:1}],
    VSTACK: [{n:"array1",r:1},{n:"array…",r:0,rep:1}],
    TAKE: [{n:"array",r:1},{n:"rows",r:1},{n:"columns",r:0}],
    DROP: [{n:"array",r:1},{n:"rows",r:1},{n:"columns",r:0}],
  };

  function getSignature(name) {
    if (!name) return null;
    const up = name.toUpperCase();
    if (SIGNATURES[up]) return { en: up, de: EN_DE[up] || null, params: SIGNATURES[up] };
    const en = DE_EN[up];
    if (en && SIGNATURES[en]) return { en, de: up, params: SIGNATURES[en] };
    return null;
  }

  // ---------- Analyzer for IDE features ----------
  // Returns: { bracketAt, matchAt, activeCall, activeArgIdx, diagnostics }
  function analyze(src, cursor, opts) {
    opts = opts || {};
    const decimal = opts.decimal || ".";
    const toks = tokenize(src, { keepWs: false, decimal });
    // bracket matches
    const pStack = []; const bStack = []; const matchTok = new Map();
    toks.forEach((t, i) => {
      if (t.type === "LPAREN") pStack.push(i);
      else if (t.type === "RPAREN") {
        const o = pStack.pop();
        if (o !== undefined) { matchTok.set(i, o); matchTok.set(o, i); }
      } else if (t.type === "LBRACE") bStack.push(i);
      else if (t.type === "RBRACE") {
        const o = bStack.pop();
        if (o !== undefined) { matchTok.set(i, o); matchTok.set(o, i); }
      }
    });
    const unmatchedOpens = [...pStack, ...bStack];
    // find bracket touched by cursor
    let bracketAt = null, matchAt = null;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.type === "LPAREN" || t.type === "RPAREN" || t.type === "LBRACE" || t.type === "RBRACE") {
        if (cursor === t.start || cursor === t.end) {
          bracketAt = t.start;
          const m = matchTok.get(i);
          if (m !== undefined) matchAt = toks[m].start;
          break;
        }
      }
    }
    // Build calls list
    const calls = [];
    const stack = [];
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.type === "LPAREN") {
        const prev = i > 0 ? toks[i - 1] : null;
        const isFunc = prev && prev.type === "FUNC";
        if (isFunc) {
          const call = {
            name: prev.value,
            fnTokIdx: i - 1,
            lparenIdx: i,
            rparenIdx: null,
            argSeps: [],
            depth: stack.length,
            hasContent: false,
            innerArgEnds: [],
          };
          stack.push({ kind: "fn", call });
          calls.push(call);
        } else {
          stack.push({ kind: "paren" });
        }
      } else if (t.type === "RPAREN") {
        const f = stack.pop();
        if (f && f.kind === "fn") f.call.rparenIdx = i;
      } else if (t.type === "COMMA" || t.type === "SEMI") {
        const top = stack[stack.length - 1];
        if (top && top.kind === "fn") top.call.argSeps.push(t.start);
      } else {
        const top = stack[stack.length - 1];
        if (top && top.kind === "fn" && t.type !== "WS") top.call.hasContent = true;
      }
    }
    // active call = innermost containing cursor
    let activeCall = null, activeArgIdx = 0;
    for (const c of calls) {
      const lp = toks[c.lparenIdx];
      const lpEnd = lp.end;
      const rpStart = c.rparenIdx !== null ? toks[c.rparenIdx].start : src.length + 1;
      if (cursor >= lpEnd && cursor <= rpStart) {
        if (!activeCall || c.depth > activeCall.depth) {
          activeCall = c;
        }
      }
    }
    if (activeCall) {
      activeArgIdx = activeCall.argSeps.filter(s => s < cursor).length;
    }
    // Diagnostics
    const diagnostics = [];
    if (pStack.length || bStack.length) {
      for (const idx of unmatchedOpens) {
        const t = toks[idx];
        diagnostics.push({
          severity: "error",
          message: `Unclosed ${t.value === "(" ? "parenthesis" : "brace"}`,
          start: t.start, end: t.end,
        });
      }
    }
    // close-paren without open
    let depthRun = 0;
    for (const t of toks) {
      if (t.type === "LPAREN") depthRun++;
      else if (t.type === "RPAREN") {
        if (depthRun === 0) {
          diagnostics.push({ severity: "error", message: "Stray closing )", start: t.start, end: t.end });
        } else depthRun--;
      }
    }
    // Per-call validation
    for (const c of calls) {
      const sig = getSignature(c.name);
      const lp = toks[c.lparenIdx];
      const rp = c.rparenIdx !== null ? toks[c.rparenIdx] : null;
      const closed = !!rp;
      const argCount = c.hasContent ? c.argSeps.length + 1 : 0;
      // empty arg detection
      // sep immediately after lparen, between two seps, or before rparen
      const between = [lp.end, ...c.argSeps.map(s => s + 1)];
      const ends = [...c.argSeps, rp ? rp.start : src.length];
      for (let k = 0; k < ends.length; k++) {
        const segStart = between[k]; const segEnd = ends[k];
        const slice = src.slice(segStart, segEnd).trim();
        if (slice === "" && (c.argSeps.length > 0 || (k > 0 || k < ends.length - 1) && closed)) {
          // only flag if has separators (else empty call is fine)
          if (c.argSeps.length > 0) {
            diagnostics.push({
              severity: "warn",
              message: `Empty argument ${k + 1} in ${c.name}()`,
              start: segStart, end: segEnd + 1,
            });
          }
        }
      }
      if (sig && closed) {
        // count required
        const reqCount = sig.params.filter(p => p.r).length;
        const lastRep = sig.params.length && sig.params[sig.params.length - 1].rep;
        if (argCount < reqCount) {
          const missing = sig.params.slice(argCount).filter(p => p.r).map(p => p.n);
          diagnostics.push({
            severity: "error",
            message: `${c.name}: missing ${missing.join(", ")}`,
            start: lp.start, end: rp.end,
          });
        }
        if (!lastRep && argCount > sig.params.length) {
          diagnostics.push({
            severity: "warn",
            message: `${c.name}: too many arguments (${argCount}, expects max ${sig.params.length})`,
            start: lp.start, end: rp.end,
          });
        }
      }
    }
    return { bracketAt, matchAt, activeCall, activeArgIdx, diagnostics, getSignature };
  }

  window.Formula = {
    tokenize, format, minify, highlight, stats, translate, detectLanguage,
    analyze, getSignature, SIGNATURES,
    EN_DE, DE_EN,
    pairCount: Object.keys(EN_DE).length,
  };
})();
