"use strict";

// ---------------------------------------------------------------------------
// sanitize.js
// Shared, sheet-agnostic data cleaning used by every data-entry script.
// Nothing here is hard-coded to a specific Excel file: callers pass values /
// header names and these helpers normalise them the way the CPCB portal needs.
//
//   - normalizeDateToISO : any date text / Excel serial / Date -> "YYYY-MM-DD"
//                          (portal uses a native <input type="date">)
//   - cleanNumber        : kill IEEE float noise (397670.02999999997 -> 397670.03)
//   - loadStateDistrictMap + correctState / correctDistrict :
//                          fuzzy spell-fix State/District against the master list
//   - isBlank / numericValue : helpers for "skip the row if empty or 0" checks
//   - sanitizeRow        : one entry point each script calls per row
// ---------------------------------------------------------------------------

const fs = require("fs");

// ---------- text ----------
function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
        if (v instanceof Date) return v.toISOString();
        if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("").trim();
        if (v.text !== undefined && v.text !== null) return String(v.text).trim();
        if (v.result !== undefined && v.result !== null) return String(v.result).trim();
    }
    return String(v).trim();
}

function isBlank(v) {
    return cellText(v) === "";
}

// ---------- numbers ----------
// Remove floating-point artifacts; optionally round to a fixed number of decimals.
// Returns a clean string suitable for typing into a form field.
function cleanNumber(v, opts = {}) {
    const { decimals = null } = opts;
    if (v === null || v === undefined || v === "") return "";
    let n;
    if (typeof v === "number") {
        n = v;
    } else if (typeof v === "object" && v.result !== undefined && typeof v.result === "number") {
        n = v.result;
    } else {
        const s = cellText(v).replace(/,/g, "").replace(/[₹\s]/g, "");
        n = Number(s);
    }
    if (!Number.isFinite(n)) return cellText(v); // not numeric -> leave untouched
    // toPrecision(12) removes binary float noise without losing real precision.
    let cleaned = Number(n.toPrecision(12));
    if (decimals !== null && Number.isFinite(decimals)) {
        cleaned = Number(cleaned.toFixed(decimals));
    }
    return String(cleaned);
}

// Numeric value for validation (e.g. "is it 0?"); null when not a number.
function numericValue(v) {
    const s = cleanNumber(v);
    if (s === "") return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
}

// ---------- dates ----------
function pad2(x) {
    return String(x).padStart(2, "0");
}

// Excel serial date (1900 system). 25569 = 1970-01-01.
function excelSerialToISO(serial) {
    const ms = Math.round((serial - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (isNaN(d.getTime())) return null;
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

const MONTHS = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
};

// Convert essentially any date representation to "YYYY-MM-DD". Throws if it
// cannot be understood (caller decides whether to skip the row).
function normalizeDateToISO(v) {
    if (v === null || v === undefined || cellText(v) === "") {
        throw new Error("date is empty");
    }

    // Real Date object (exceljs returns these for true date cells)
    if (v instanceof Date) {
        return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
    }
    if (typeof v === "object" && v.result instanceof Date) {
        const d = v.result;
        return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    }

    const s = cellText(v);

    // Bare number -> treat as Excel serial when in a plausible range (~1954-2089)
    if (typeof v === "number" || /^\d+(\.\d+)?$/.test(s)) {
        const num = typeof v === "number" ? v : Number(s);
        if (num > 20000 && num < 80000) {
            const iso = excelSerialToISO(num);
            if (iso) return iso;
        }
    }

    let m;
    // ISO (optionally with time)
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/);
    if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

    // YYYY/MM/DD or YYYY.MM.DD
    m = s.match(/^(\d{4})[.\/](\d{1,2})[.\/](\d{1,2})$/);
    if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

    // DD.MM.YYYY / DD-MM-YYYY / DD/MM/YYYY  (Indian day-first)
    m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
    if (m) {
        const dd = Number(m[1]);
        const mm = Number(m[2]);
        if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
            return `${m[3]}-${pad2(mm)}-${pad2(dd)}`;
        }
    }

    // 01-Apr-2025 / 1 April 2025
    m = s.match(/^(\d{1,2})[ \-]([A-Za-z]{3,})[ \-](\d{4})$/);
    if (m) {
        const mon = MONTHS[m[2].slice(0, 3).toLowerCase()];
        if (mon) return `${m[3]}-${pad2(mon)}-${pad2(Number(m[1]))}`;
    }

    // Last resort
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
        return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    }

    throw new Error(`Unsupported date format: ${s}`);
}

function safeDateToISO(v) {
    try {
        return normalizeDateToISO(v);
    } catch {
        return cellText(v);
    }
}

// ---------- CSV ----------
function parseCsvLine(line) {
    const out = [];
    let cur = "";
    let q = false;
    for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (q) {
            if (c === '"') {
                if (line[i + 1] === '"') { cur += '"'; i++; }
                else q = false;
            } else cur += c;
        } else if (c === '"') {
            q = true;
        } else if (c === ",") {
            out.push(cur); cur = "";
        } else {
            cur += c;
        }
    }
    out.push(cur);
    return out;
}

// ---------- state / district mapping ----------
function normKey(s) {
    return cellText(s).toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Reads the "State,District" master CSV (ignores the descriptive rows above the
// real header). Returns { states, districts } maps keyed by a normalised key.
function loadStateDistrictMap(csvPath) {
    const raw = fs.readFileSync(csvPath, "utf8");
    const lines = raw.split(/\r?\n/);
    const states = new Map();    // normKey(state)  -> display state
    const districts = new Map(); // normKey(state)  -> Map(normKey(district) -> display district)
    let headerSeen = false;
    for (const line of lines) {
        if (!line.trim()) continue;
        const cols = parseCsvLine(line);
        const stateRaw = (cols[0] || "").trim();
        const distRaw = (cols[1] || "").trim();
        if (!headerSeen) {
            if (normKey(stateRaw) === "STATE" && normKey(distRaw) === "DISTRICT") headerSeen = true;
            continue;
        }
        if (!stateRaw || !distRaw) continue;
        const sk = normKey(stateRaw);
        if (!states.has(sk)) states.set(sk, stateRaw);
        if (!districts.has(sk)) districts.set(sk, new Map());
        districts.get(sk).set(normKey(distRaw), distRaw);
    }
    return { states, districts };
}

// ---------- fuzzy matching ----------
function levenshtein(a, b) {
    a = a || ""; b = b || "";
    const m = a.length, n = b.length;
    if (!m) return n;
    if (!n) return m;
    const dp = Array.from({ length: m + 1 }, (_, i) => i);
    for (let j = 1; j <= n; j++) {
        let prev = dp[0];
        dp[0] = j;
        for (let i = 1; i <= m; i++) {
            const tmp = dp[i];
            dp[i] = Math.min(dp[i] + 1, dp[i - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
            prev = tmp;
        }
    }
    return dp[m];
}

// candidates: [{ key, display }] -> best { display, score, exact } or null
function bestMatch(input, candidates) {
    const ik = normKey(input);
    if (!ik) return null;
    for (const c of candidates) {
        if (c.key === ik) return { display: c.display, score: 1, exact: true };
    }
    let best = null;
    for (const c of candidates) {
        const dist = levenshtein(ik, c.key);
        const maxLen = Math.max(ik.length, c.key.length) || 1;
        let score = 1 - dist / maxLen;
        if (c.key.includes(ik) || ik.includes(c.key)) score = Math.max(score, 0.9);
        if (!best || score > best.score) best = { display: c.display, score, exact: false };
    }
    return best;
}

// Returns { value, changed, score } when a confident match is found, else null.
function correctState(raw, map, threshold = 0.6) {
    if (!map) return null;
    const cands = [...map.states.entries()].map(([key, display]) => ({ key, display }));
    const r = bestMatch(raw, cands);
    if (!r) return null;
    if (r.exact || r.score >= threshold) {
        return { value: r.display, changed: normKey(raw) !== normKey(r.display), score: r.score };
    }
    return null;
}

function correctDistrict(raw, stateDisplay, map, threshold = 0.6) {
    if (!map) return null;
    const dmap = map.districts.get(normKey(stateDisplay));
    if (!dmap) return null;
    const cands = [...dmap.entries()].map(([key, display]) => ({ key, display }));
    const r = bestMatch(raw, cands);
    if (!r) return null;
    if (r.exact || r.score >= threshold) {
        return { value: r.display, changed: normKey(raw) !== normKey(r.display), score: r.score };
    }
    return null;
}

// ---------------------------------------------------------------------------
// sanitizeRow: the one entry point every data-entry script calls per row.
// It is sheet-agnostic — the caller passes the column names that exist in
// THIS sheet, so the same code works for PP, PET, PIBO or any future sheet.
//
//   accessors = { getVal(row, headerMap, name), setVal(row, headerMap, name, v) }
//   spec = {
//     mandatory:    ["Entity Type*", "State*", ...]   // blank -> skip row
//     dateFields:   ["Sales date*"]                   // -> "YYYY-MM-DD"
//     numberFields: [{ name: "Quantity Sold(MT)" },   // clean float noise
//                    { name: "Principal Amount(₹)*", decimals: 2 }]
//     stateField:   "State*"        // optional, fuzzy-corrected vs master list
//     districtField:"District*"     // optional (needs stateField + map)
//     map:          <loadStateDistrictMap result>     // optional
//   }
//
// Returns { skip: boolean, reason: string, notes: string[] }.
// On the way it rewrites the row's cells to the cleaned values, so whatever the
// script reads afterwards is already portal-ready.
// ---------------------------------------------------------------------------
function sanitizeRow(row, headerMap, accessors, spec) {
    const { getVal, setVal } = accessors;
    const s = spec || {};
    const notes = [];
    const present = (name) => headerMap.has(name.trim().replace(/\s+/g, " ").toLowerCase());
    const stop = (reason) => ({ skip: true, reason, notes });

    // 1) State spell-fix
    if (s.stateField && s.map && present(s.stateField)) {
        const raw = cellText(getVal(row, headerMap, s.stateField));
        if (raw) {
            const fix = correctState(raw, s.map);
            if (fix) {
                if (fix.changed) {
                    setVal(row, headerMap, s.stateField, fix.value);
                    notes.push(`State "${raw}" -> "${fix.value}"`);
                }
            } else if ((s.mandatory || []).includes(s.stateField)) {
                return stop(`State not recognized: "${raw}"`);
            }
        }
    }

    // 2) District spell-fix (scoped to the corrected state)
    if (s.districtField && s.map && present(s.districtField)) {
        const raw = cellText(getVal(row, headerMap, s.districtField));
        const stateNow = s.stateField ? cellText(getVal(row, headerMap, s.stateField)) : "";
        if (raw && stateNow) {
            const fix = correctDistrict(raw, stateNow, s.map);
            if (fix) {
                if (fix.changed) {
                    setVal(row, headerMap, s.districtField, fix.value);
                    notes.push(`District "${raw}" -> "${fix.value}"`);
                }
            } else if ((s.mandatory || []).includes(s.districtField)) {
                return stop(`District not recognized: "${raw}" (state ${stateNow})`);
            }
        }
    }

    // 3) Dates -> ISO
    for (const name of s.dateFields || []) {
        if (!present(name)) continue;
        const raw = getVal(row, headerMap, name);
        if (isBlank(raw)) continue; // emptiness handled by mandatory check below
        try {
            const iso = normalizeDateToISO(raw);
            if (iso !== cellText(raw)) {
                setVal(row, headerMap, name, iso);
                notes.push(`${name} "${cellText(raw)}" -> "${iso}"`);
            }
        } catch (e) {
            if ((s.mandatory || []).includes(name)) {
                return stop(`${name} invalid date: "${cellText(raw)}"`);
            }
        }
    }

    // 4) Numbers -> cleaned
    for (const nf of s.numberFields || []) {
        const name = nf.name;
        if (!present(name)) continue;
        const raw = getVal(row, headerMap, name);
        if (isBlank(raw)) continue;
        const cleaned = cleanNumber(raw, { decimals: nf.decimals ?? null });
        if (cleaned !== "" && cleaned !== cellText(raw)) {
            setVal(row, headerMap, name, cleaned);
        }
    }

    // 5) Mandatory: blank -> skip; numeric mandatory that is 0 -> skip
    const numberNames = new Set((s.numberFields || []).map((n) => n.name));
    for (const name of s.mandatory || []) {
        if (!present(name)) continue; // sheet doesn't have this column -> don't enforce
        const val = getVal(row, headerMap, name);
        if (isBlank(val)) return stop(`${name} is empty`);
        if (numberNames.has(name)) {
            const num = numericValue(val);
            if (num === 0) return stop(`${name} is 0`);
        }
    }

    return { skip: false, reason: "", notes };
}

module.exports = {
    cellText,
    isBlank,
    cleanNumber,
    numericValue,
    normalizeDateToISO,
    safeDateToISO,
    excelSerialToISO,
    loadStateDistrictMap,
    correctState,
    correctDistrict,
    sanitizeRow,
};
