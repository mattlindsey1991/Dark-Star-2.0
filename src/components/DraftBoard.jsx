import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { supabase } from "../supabaseClient";
import * as XLSX from "xlsx";

const OFFENSE_POSITIONS = [
  { abbr: "QB", name: "Quarterback" },
  { abbr: "RB", name: "Running back" },
  { abbr: "WR", name: "Wide receiver" },
  { abbr: "TE", name: "Tight end" },
  { abbr: "OT", name: "Offensive tackle" },
  { abbr: "OG", name: "Offensive guard" },
  { abbr: "OC", name: "Center" },
];

const DEFENSE_POSITIONS = [
  { abbr: "DL", name: "Defensive line" },
  { abbr: "EDGE", name: "Edge rusher" },
  { abbr: "LB", name: "Linebacker" },
  { abbr: "DS", name: "Safety" },
  { abbr: "DC", name: "Cornerback" },
  { abbr: "PT", name: "Punter" },
  { abbr: "PK", name: "Kicker" },
  { abbr: "LS", name: "Long snapper" },
];

const YEARS = [2027, 2028, 2029, 2030, 2031];
const ENTRY_YEARS = [2021, 2022, 2023, 2024, 2025, 2026];
const GRADE_SCALE = [1.0, 1.2, 1.5, 1.8, 2.0, 2.2, 2.5, 2.8, 3.3, 3.8, 4.3, 4.8, 5.3, 5.8, 6.3, 6.8, 7.3, 7.8, 8.0, 8.5, 9.0];

const POSITION_BOARD = {};
OFFENSE_POSITIONS.forEach((p) => (POSITION_BOARD[p.abbr] = "OFFENSE"));
DEFENSE_POSITIONS.forEach((p) => (POSITION_BOARD[p.abbr] = "DEFENSE"));
const ALL_POSITIONS = [...OFFENSE_POSITIONS, ...DEFENSE_POSITIONS];
const KNOWN_IMPORT_COLUMNS = new Set(["name", "position", "school", "entry year", "entryyear", "agents"]);

const COLORS = {
  bg: "#15171A",
  surface: "#1D2024",
  surfaceHi: "#242830",
  ink: "#ECE7DC",
  inkDim: "#9C9C93",
  hair: "rgba(236,231,220,0.10)",
  hairStrong: "rgba(236,231,220,0.18)",
  offense: "#C98A3E",
  offenseDim: "rgba(201,138,62,0.14)",
  defense: "#3E7B94",
  defenseDim: "rgba(62,123,148,0.14)",
  ungraded: "#54585F",
  tierGreen: "#4C9A5B",
  tierGreenText: "#EAF6EC",
  tierYellow: "#D9B23C",
  tierYellowText: "#2B2000",
  tierRed: "#C24E4E",
  tierRedText: "#FCEDED",
  tierBlack: "#0D0E10",
  tierBlackText: "#ECE7DC",
};

function computeAvg(grades) {
  if (!grades || grades.length === 0) return null;
  const sum = grades.reduce((a, g) => a + Number(g.grade), 0);
  return sum / grades.length;
}

function gradeTier(avg) {
  // Lower grade is better on this scale: 1.0 is elite, 9.0 is not draftable.
  if (avg === null) return { label: "Ungraded", color: COLORS.ungraded, text: COLORS.ink, filled: false };
  if (avg < 3.5) return { label: "Elite", color: COLORS.tierGreen, text: COLORS.tierGreenText, filled: true };
  if (avg < 5.5) return { label: "Depth / backup", color: COLORS.tierYellow, text: COLORS.tierYellowText, filled: true };
  if (avg < 9.0) return { label: "Priority FA", color: COLORS.tierRed, text: COLORS.tierRedText, filled: true };
  return { label: "Not draftable", color: COLORS.tierBlack, text: COLORS.tierBlackText, filled: true };
}

function fmtGrade(avg) {
  return avg === null ? "—" : avg.toFixed(1);
}

export default function DraftBoard({ session }) {
  const [prospects, setProspects] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [board, setBoard] = useState("OFFENSE");
  const [year, setYear] = useState(2027);
  const [search, setSearch] = useState("");
  const [expandedId, setExpandedId] = useState(null);
  const [addOpenFor, setAddOpenFor] = useState(null);
  const [addDraft, setAddDraft] = useState({ name: "", school: "", entryYear: 2024 });
  const [gradeDraft, setGradeDraft] = useState({ scout: "", grade: GRADE_SCALE[0] });
  const [errorMsg, setErrorMsg] = useState("");
  const [importing, setImporting] = useState(false);
  const [importSummary, setImportSummary] = useState(null);
  const [pwOpen, setPwOpen] = useState(false);
  const [pwDraft, setPwDraft] = useState({ pw1: "", pw2: "" });
  const [pwMsg, setPwMsg] = useState("");
  const fileInputRef = useRef(null);

  const fetchAll = useCallback(async () => {
    const { data, error } = await supabase
      .from("prospects")
      .select("*, grades(*)")
      .order("created_at", { ascending: true });
    if (error) {
      setErrorMsg("Couldn't load the board. Try refreshing.");
    } else {
      setProspects(data);
      setErrorMsg("");
    }
    setLoaded(true);
  }, []);

  useEffect(() => {
    fetchAll();

    const channel = supabase
      .channel("draft-board-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "prospects" }, fetchAll)
      .on("postgres_changes", { event: "*", schema: "public", table: "grades" }, fetchAll)
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, [fetchAll]);

  const positions = board === "OFFENSE" ? OFFENSE_POSITIONS : DEFENSE_POSITIONS;

  const grouped = useMemo(() => {
    const q = search.trim().toLowerCase();
    const map = {};
    positions.forEach((p) => (map[p.abbr] = []));
    prospects
      .filter((pr) => pr.draft_class_year === year && map[pr.position] !== undefined)
      .filter(
        (pr) =>
          !q ||
          pr.name.toLowerCase().includes(q) ||
          (pr.school || "").toLowerCase().includes(q)
      )
      .forEach((pr) => map[pr.position].push(pr));
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => {
        const av = computeAvg(a.grades);
        const bv = computeAvg(b.grades);
        if (av === null && bv === null) return 0;
        if (av === null) return 1;
        if (bv === null) return -1;
        return av - bv;
      });
    });
    return map;
  }, [prospects, positions, year, search]);

  const totalCount = useMemo(
    () => Object.values(grouped).reduce((a, list) => a + list.length, 0),
    [grouped]
  );

  function openAdd(posAbbr) {
    setAddOpenFor(posAbbr);
    setAddDraft({ name: "", school: "", entryYear: 2024 });
  }

  async function submitAdd(posAbbr) {
    if (!addDraft.name.trim()) return;
    const { data, error } = await supabase
      .from("prospects")
      .insert({
        name: addDraft.name.trim(),
        position: posAbbr,
        board,
        school: addDraft.school.trim(),
        draft_class_year: year,
        entry_year: Number(addDraft.entryYear),
        agents: "",
        created_by: session.user.id,
      })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't add that prospect. Try again.");
      return;
    }
    setProspects((prev) => [...prev, { ...data, grades: [] }]);
    setAddOpenFor(null);
    setExpandedId(data.id);
  }

  async function updateProspect(id, patch) {
    setProspects((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
    const { error } = await supabase.from("prospects").update(patch).eq("id", id);
    if (error) setErrorMsg("That change didn't save. Try again.");
  }

  async function deleteProspect(id) {
    setProspects((prev) => prev.filter((p) => p.id !== id));
    if (expandedId === id) setExpandedId(null);
    const { error } = await supabase.from("prospects").delete().eq("id", id);
    if (error) setErrorMsg("Couldn't remove that prospect. Try again.");
  }

  async function addGrade(id) {
    const g = parseFloat(gradeDraft.grade);
    if (!gradeDraft.scout.trim() || isNaN(g) || !GRADE_SCALE.includes(g)) return;
    const { data, error } = await supabase
      .from("grades")
      .insert({ prospect_id: id, scout: gradeDraft.scout.trim(), grade: g, created_by: session.user.id })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't save that grade. Try again.");
      return;
    }
    setProspects((prev) =>
      prev.map((p) => (p.id === id ? { ...p, grades: [...p.grades, data] } : p))
    );
    setGradeDraft({ scout: "", grade: GRADE_SCALE[0] });
  }

  async function deleteGrade(prospectId, gradeId) {
    setProspects((prev) =>
      prev.map((p) =>
        p.id === prospectId ? { ...p, grades: p.grades.filter((g) => g.id !== gradeId) } : p
      )
    );
    const { error } = await supabase.from("grades").delete().eq("id", gradeId);
    if (error) setErrorMsg("Couldn't remove that grade. Try again.");
  }

  function normalizeHeader(h) {
    return String(h || "").trim().toLowerCase();
  }

  async function handleImportFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setImporting(true);
    setImportSummary(null);
    setErrorMsg("");

    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const skipped = [];
      const prospectRows = [];
      const gradeRows = [];

      rows.forEach((row, idx) => {
        const keys = Object.keys(row);
        const keyMap = {};
        keys.forEach((k) => (keyMap[normalizeHeader(k)] = k));

        const nameKey = keyMap["name"];
        const posKey = keyMap["position"];
        const schoolKey = keyMap["school"];
        const entryYearKey = keyMap["entry year"] || keyMap["entryyear"];
        const agentsKey = keyMap["agents"];

        const name = nameKey ? String(row[nameKey]).trim() : "";
        const positionRaw = posKey ? String(row[posKey]).trim().toUpperCase() : "";

        if (!name) {
          skipped.push({ row: idx + 2, reason: "Missing name" });
          return;
        }
        const rowBoard = POSITION_BOARD[positionRaw];
        if (!rowBoard) {
          skipped.push({ row: idx + 2, reason: `Unrecognized position "${positionRaw}" for ${name}` });
          return;
        }

        const entryYearRaw = entryYearKey ? row[entryYearKey] : "";
        const entryYearNum = parseInt(entryYearRaw, 10);

        const tempId = crypto.randomUUID();
        prospectRows.push({
          id: tempId,
          name,
          position: positionRaw,
          board: rowBoard,
          school: schoolKey ? String(row[schoolKey]).trim() : "",
          draft_class_year: year,
          entry_year: Number.isFinite(entryYearNum) ? entryYearNum : null,
          agents: agentsKey ? String(row[agentsKey]).trim() : "",
          created_by: session.user.id,
        });

        keys.forEach((k) => {
          const normalized = normalizeHeader(k);
          if (KNOWN_IMPORT_COLUMNS.has(normalized)) return;
          const val = row[k];
          if (val === "" || val === null || val === undefined) return;
          const num = parseFloat(val);
          if (!Number.isFinite(num)) return;
          gradeRows.push({
            prospect_id: tempId,
            scout: String(k).trim(),
            grade: num,
            created_by: session.user.id,
          });
        });
      });

      if (prospectRows.length > 0) {
        const { error: prospectErr } = await supabase.from("prospects").insert(prospectRows);
        if (prospectErr) {
          setErrorMsg("Import failed while saving prospects: " + prospectErr.message);
          setImporting(false);
          e.target.value = "";
          return;
        }
      }
      if (gradeRows.length > 0) {
        const { error: gradeErr } = await supabase.from("grades").insert(gradeRows);
        if (gradeErr) {
          setErrorMsg("Prospects imported, but grades failed to save: " + gradeErr.message);
        }
      }

      await fetchAll();
      setImportSummary({
        prospectCount: prospectRows.length,
        gradeCount: gradeRows.length,
        skipped,
        year,
      });
    } catch (err) {
      setErrorMsg("Couldn't read that file. Make sure it's a valid .xlsx or .csv.");
    }

    setImporting(false);
    e.target.value = "";
  }

  async function handleSetPassword() {
    setPwMsg("");
    if (pwDraft.pw1.length < 6) {
      setPwMsg("Password must be at least 6 characters.");
      return;
    }
    if (pwDraft.pw1 !== pwDraft.pw2) {
      setPwMsg("Passwords don't match.");
      return;
    }
    const { error } = await supabase.auth.updateUser({ password: pwDraft.pw1 });
    if (error) {
      setPwMsg(error.message);
      return;
    }
    setPwMsg("Password set. You can sign in with it next time.");
    setPwDraft({ pw1: "", pw2: "" });
  }

  const accent = board === "OFFENSE" ? COLORS.offense : COLORS.defense;

  return (
    <div
      style={{
        background: COLORS.bg,
        color: COLORS.ink,
        fontFamily: "'Inter', sans-serif",
        minHeight: "100vh",
        padding: "28px 24px 40px",
        position: "relative",
        overflow: "hidden",
      }}
    >
      <style>{`
        .db-row:hover { background: rgba(236,231,220,0.04) !important; }
        .db-input {
          background: rgba(236,231,220,0.06);
          border: 1px solid ${COLORS.hair};
          color: ${COLORS.ink};
          border-radius: 4px;
          padding: 6px 8px;
          font-family: 'Inter', sans-serif;
          font-size: 12.5px;
          outline: none;
        }
        .db-input:focus { border-color: ${COLORS.hairStrong}; }
        .db-input::placeholder { color: ${COLORS.inkDim}; }
        .db-btn {
          background: transparent;
          border: 1px solid ${COLORS.hair};
          color: ${COLORS.inkDim};
          border-radius: 4px;
          cursor: pointer;
          transition: all 0.12s ease;
        }
        .db-btn:hover { border-color: ${COLORS.hairStrong}; color: ${COLORS.ink}; }
      `}</style>

      <div
        style={{
          position: "absolute",
          top: "-40px",
          right: "20px",
          width: "220px",
          height: "220px",
          borderRadius: "50%",
          border: `6px double ${accent}`,
          opacity: 0.07,
          transform: "rotate(-10deg)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          pointerEvents: "none",
        }}
      >
        <span style={{ fontFamily: "'Anton', sans-serif", fontSize: "56px", color: accent }}>1.0</span>
        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", letterSpacing: "2px", color: accent }}>
          ELITE
        </span>
      </div>

      <div style={{ position: "relative", zIndex: 1 }}>
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: "14px" }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: "14px", flexWrap: "wrap" }}>
            <h1 style={{ fontFamily: "'Anton', sans-serif", fontSize: "40px", letterSpacing: "1px", margin: 0, lineHeight: 1 }}>
              THE BIG BOARD
            </h1>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.inkDim, letterSpacing: "1px" }}>
              {session.user.email}
            </span>
          </div>
          <div style={{ display: "flex", gap: "8px" }}>
            <button className="db-btn" onClick={() => setPwOpen((v) => !v)} style={{ padding: "7px 12px", fontSize: "12px" }}>
              {pwOpen ? "Close" : "Set password"}
            </button>
            <button className="db-btn" onClick={() => supabase.auth.signOut()} style={{ padding: "7px 12px", fontSize: "12px" }}>
              Sign out
            </button>
          </div>
        </div>

        {pwOpen && (
          <div
            style={{
              background: COLORS.surfaceHi,
              border: `1px solid ${COLORS.hair}`,
              borderRadius: "6px",
              padding: "12px 14px",
              marginTop: "10px",
              maxWidth: "320px",
            }}
          >
            <div style={{ fontSize: "11.5px", color: COLORS.inkDim, marginBottom: "8px" }}>
              Set a password to sign in with email + password next time.
            </div>
            <input
              className="db-input"
              type="password"
              placeholder="New password"
              style={{ width: "100%", marginBottom: "6px" }}
              value={pwDraft.pw1}
              onChange={(e) => setPwDraft({ ...pwDraft, pw1: e.target.value })}
            />
            <input
              className="db-input"
              type="password"
              placeholder="Confirm password"
              style={{ width: "100%", marginBottom: "8px" }}
              value={pwDraft.pw2}
              onChange={(e) => setPwDraft({ ...pwDraft, pw2: e.target.value })}
            />
            {pwMsg && (
              <div style={{ fontSize: "11.5px", color: pwMsg.startsWith("Password set") ? "#8FBF8F" : "#D98080", marginBottom: "8px" }}>
                {pwMsg}
              </div>
            )}
            <button className="db-btn" onClick={handleSetPassword} style={{ padding: "6px 12px", fontSize: "12px", color: accent, borderColor: accent }}>
              Save password
            </button>
          </div>
        )}

        <div
          style={{
            height: "10px",
            marginTop: "10px",
            marginBottom: "20px",
            backgroundImage: `repeating-linear-gradient(90deg, ${COLORS.hair} 0px, ${COLORS.hair} 1px, transparent 1px, transparent 24px)`,
          }}
        />

        <div style={{ display: "flex", alignItems: "center", gap: "20px", flexWrap: "wrap", marginBottom: "6px" }}>
          <div style={{ display: "flex", border: `1px solid ${COLORS.hair}`, borderRadius: "6px", overflow: "hidden" }}>
            {["OFFENSE", "DEFENSE"].map((b) => (
              <button
                key={b}
                onClick={() => setBoard(b)}
                style={{
                  fontFamily: "'Bebas Neue', sans-serif",
                  fontSize: "17px",
                  letterSpacing: "1.5px",
                  padding: "8px 22px",
                  border: "none",
                  cursor: "pointer",
                  background: board === b ? (b === "OFFENSE" ? COLORS.offenseDim : COLORS.defenseDim) : "transparent",
                  color: board === b ? (b === "OFFENSE" ? COLORS.offense : COLORS.defense) : COLORS.inkDim,
                }}
              >
                {b}
              </button>
            ))}
          </div>

          <div style={{ display: "flex", gap: "6px" }}>
            {YEARS.map((y) => (
              <button
                key={y}
                onClick={() => setYear(y)}
                className="db-btn"
                style={{
                  fontFamily: "'IBM Plex Mono', monospace",
                  fontSize: "12.5px",
                  padding: "7px 10px",
                  borderColor: year === y ? accent : COLORS.hair,
                  color: year === y ? accent : COLORS.inkDim,
                }}
              >
                {y}
              </button>
            ))}
          </div>

          <input
            className="db-input"
            placeholder="Search name or school"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ width: "180px", marginLeft: "auto" }}
          />

          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            ref={fileInputRef}
            onChange={handleImportFile}
            style={{ display: "none" }}
          />
          <button
            className="db-btn"
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            disabled={importing}
            title={`Imports into the ${year} class. Columns: Name, Position, School, Entry Year, Agents, plus any scout columns.`}
            style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
          >
            {importing ? "Importing…" : "Upload Excel"}
          </button>

          <div style={{ display: "flex", alignItems: "center", gap: "6px", fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.inkDim }}>
            {totalCount} prospects · {year}
          </div>
        </div>

        {importSummary && (
          <div
            style={{
              background: COLORS.surfaceHi,
              border: `1px solid ${COLORS.hair}`,
              borderRadius: "6px",
              padding: "10px 14px",
              marginBottom: "14px",
              fontSize: "12px",
              color: COLORS.inkDim,
            }}
          >
            <div style={{ color: COLORS.ink, marginBottom: importSummary.skipped.length ? "6px" : 0 }}>
              Imported {importSummary.prospectCount} prospect{importSummary.prospectCount === 1 ? "" : "s"} and{" "}
              {importSummary.gradeCount} grade{importSummary.gradeCount === 1 ? "" : "s"} into the {importSummary.year} class.
              {importSummary.skipped.length > 0 && ` ${importSummary.skipped.length} row(s) skipped:`}
            </div>
            {importSummary.skipped.length > 0 && (
              <ul style={{ margin: 0, paddingLeft: "18px" }}>
                {importSummary.skipped.map((s, i) => (
                  <li key={i}>Row {s.row}: {s.reason}</li>
                ))}
              </ul>
            )}
            <button className="db-btn" onClick={() => setImportSummary(null)} style={{ marginTop: "6px", padding: "3px 8px", fontSize: "11px" }}>
              Dismiss
            </button>
          </div>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: "14px", marginBottom: "16px", flexWrap: "wrap" }}>
          <span style={{ fontSize: "11px", color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace" }}>
            LOWER GRADE RANKS HIGHER
          </span>
          {[
            { label: "1.0–3.49", color: COLORS.tierGreen },
            { label: "3.5–5.49", color: COLORS.tierYellow },
            { label: "5.5–8.99", color: COLORS.tierRed },
            { label: "9.0", color: COLORS.tierBlack },
          ].map((s) => (
            <span key={s.label} style={{ display: "flex", alignItems: "center", gap: "5px", fontSize: "11px", color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace" }}>
              <span style={{ width: "9px", height: "9px", borderRadius: "50%", background: s.color, border: "1px solid rgba(255,255,255,0.15)" }} />
              {s.label}
            </span>
          ))}
        </div>

        {errorMsg && (
          <div style={{ fontSize: "12px", color: "#D98080", marginBottom: "12px", fontFamily: "'IBM Plex Mono', monospace" }}>
            {errorMsg}
          </div>
        )}

        {!loaded ? (
          <div style={{ color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace", fontSize: "13px" }}>Loading board…</div>
        ) : (
          <div style={{ display: "flex", gap: "14px", overflowX: "auto", paddingBottom: "12px" }}>
            {positions.map((pos) => {
              const list = grouped[pos.abbr] || [];
              return (
                <div
                  key={pos.abbr}
                  style={{
                    minWidth: "270px",
                    maxWidth: "270px",
                    background: COLORS.surface,
                    borderRadius: "8px",
                    border: `1px solid ${COLORS.hair}`,
                    display: "flex",
                    flexDirection: "column",
                  }}
                >
                  <div style={{ padding: "12px 14px 10px", borderBottom: `1px solid ${COLORS.hair}` }}>
                    <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between" }}>
                      <span style={{ fontFamily: "'Bebas Neue', sans-serif", fontSize: "24px", letterSpacing: "1px", color: accent }}>
                        {pos.abbr}
                      </span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.inkDim }}>
                        {list.length}
                      </span>
                    </div>
                    <div style={{ fontSize: "11.5px", color: COLORS.inkDim }}>{pos.name}</div>
                  </div>

                  <div style={{ flex: 1 }}>
                    {list.length === 0 && addOpenFor !== pos.abbr && (
                      <div style={{ padding: "16px 14px", fontSize: "12px", color: COLORS.inkDim, lineHeight: 1.5 }}>
                        No prospects logged for {year} yet.
                      </div>
                    )}
                    {list.map((p, idx) => {
                      const avg = computeAvg(p.grades);
                      const tier = gradeTier(avg);
                      const isOpen = expandedId === p.id;
                      return (
                        <div key={p.id} style={{ borderBottom: `1px solid ${COLORS.hair}` }}>
                          <div
                            className="db-row"
                            onClick={() => setExpandedId(isOpen ? null : p.id)}
                            style={{ display: "flex", alignItems: "center", gap: "10px", padding: "10px 14px", cursor: "pointer" }}
                          >
                            <span style={{ fontFamily: "'Anton', sans-serif", fontSize: "20px", color: COLORS.hairStrong, width: "22px" }}>
                              {String(idx + 1).padStart(2, "0")}
                            </span>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: "13.5px", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.3px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {p.name}
                              </div>
                              <div style={{ fontSize: "11px", color: COLORS.inkDim, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {p.school || "School unset"}
                              </div>
                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, opacity: 0.75, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {p.agents ? p.agents : "No agent listed"}
                              </div>
                            </div>
                            <div
                              title={tier.label}
                              style={{
                                fontFamily: "'IBM Plex Mono', monospace",
                                fontSize: "12px",
                                fontWeight: 600,
                                color: tier.filled ? tier.text : tier.color,
                                background: tier.filled ? tier.color : "transparent",
                                border: `1.5px solid ${tier.filled ? "rgba(255,255,255,0.15)" : tier.color}`,
                                borderRadius: "50%",
                                width: "34px",
                                height: "34px",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                transform: "rotate(-4deg)",
                                flexShrink: 0,
                              }}
                            >
                              {fmtGrade(avg)}
                            </div>
                          </div>

                          {isOpen && (
                            <div style={{ padding: "4px 14px 14px", background: COLORS.surfaceHi }}>
                              <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "10px" }}>
                                <span style={{ width: "9px", height: "9px", borderRadius: "50%", background: tier.color, border: tier.filled ? "1px solid rgba(255,255,255,0.15)" : "none", flexShrink: 0 }} />
                                <span style={{ fontSize: "10.5px", color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace" }}>
                                  {tier.label.toUpperCase()}
                                </span>
                              </div>

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Position</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.position}
                                    onChange={(e) => {
                                      const newPos = e.target.value;
                                      updateProspect(p.id, { position: newPos, board: POSITION_BOARD[newPos] });
                                    }}
                                  >
                                    <optgroup label="Offense">
                                      {OFFENSE_POSITIONS.map((op) => (
                                        <option key={op.abbr} value={op.abbr}>{op.abbr}</option>
                                      ))}
                                    </optgroup>
                                    <optgroup label="Defense">
                                      {DEFENSE_POSITIONS.map((dp) => (
                                        <option key={dp.abbr} value={dp.abbr}>{dp.abbr}</option>
                                      ))}
                                    </optgroup>
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Draft class</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.draft_class_year}
                                    onChange={(e) => updateProspect(p.id, { draft_class_year: Number(e.target.value) })}
                                  >
                                    {YEARS.map((y) => (
                                      <option key={y} value={y}>{y}</option>
                                    ))}
                                  </select>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>School</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                defaultValue={p.school}
                                onBlur={(e) => updateProspect(p.id, { school: e.target.value })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>College entry year</label>
                              <select
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                value={p.entry_year || ""}
                                onChange={(e) => updateProspect(p.id, { entry_year: Number(e.target.value) })}
                              >
                                {ENTRY_YEARS.map((ey) => (
                                  <option key={ey} value={ey}>Ent: {ey}</option>
                                ))}
                              </select>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Agents assigned</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "10px" }}
                                placeholder="e.g. J. Rosenhaus"
                                defaultValue={p.agents}
                                onBlur={(e) => updateProspect(p.id, { agents: e.target.value })}
                              />

                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, marginBottom: "5px" }}>
                                Scout grades ({p.grades.length})
                              </div>
                              {p.grades.map((g) => (
                                <div key={g.id} style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", marginBottom: "4px" }}>
                                  <span style={{ flex: 1, color: COLORS.ink }}>{g.scout}</span>
                                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkDim }}>{Number(g.grade).toFixed(1)}</span>
                                  <button
                                    className="db-btn"
                                    onClick={(e) => { e.stopPropagation(); deleteGrade(p.id, g.id); }}
                                    style={{ padding: "2px 5px" }}
                                  >
                                    ×
                                  </button>
                                </div>
                              ))}
                              <div style={{ display: "flex", gap: "6px", marginTop: "6px" }} onClick={(e) => e.stopPropagation()}>
                                <input
                                  className="db-input"
                                  placeholder="Scout"
                                  style={{ flex: 1, width: 0 }}
                                  value={gradeDraft.scout}
                                  onChange={(e) => setGradeDraft({ ...gradeDraft, scout: e.target.value })}
                                />
                                <select
                                  className="db-input"
                                  style={{ width: "68px" }}
                                  value={gradeDraft.grade}
                                  onChange={(e) => setGradeDraft({ ...gradeDraft, grade: e.target.value })}
                                >
                                  {GRADE_SCALE.map((g) => (
                                    <option key={g} value={g}>{g.toFixed(1)}</option>
                                  ))}
                                </select>
                                <button className="db-btn" onClick={() => addGrade(p.id)} style={{ padding: "0 8px" }}>
                                  +
                                </button>
                              </div>

                              <button
                                className="db-btn"
                                onClick={(e) => { e.stopPropagation(); deleteProspect(p.id); }}
                                style={{ marginTop: "12px", fontSize: "11px", color: "#C97A7A", borderColor: "rgba(201,122,122,0.3)", padding: "5px 9px" }}
                              >
                                Remove prospect
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  <div style={{ padding: "10px 14px 14px" }}>
                    {addOpenFor === pos.abbr ? (
                      <div>
                        <input
                          className="db-input"
                          placeholder="Player name"
                          style={{ width: "100%", marginBottom: "6px" }}
                          value={addDraft.name}
                          onChange={(e) => setAddDraft({ ...addDraft, name: e.target.value })}
                          autoFocus
                        />
                        <input
                          className="db-input"
                          placeholder="School"
                          style={{ width: "100%", marginBottom: "6px" }}
                          value={addDraft.school}
                          onChange={(e) => setAddDraft({ ...addDraft, school: e.target.value })}
                        />
                        <select
                          className="db-input"
                          style={{ width: "100%", marginBottom: "8px" }}
                          value={addDraft.entryYear}
                          onChange={(e) => setAddDraft({ ...addDraft, entryYear: Number(e.target.value) })}
                        >
                          {ENTRY_YEARS.map((ey) => (
                            <option key={ey} value={ey}>Ent: {ey}</option>
                          ))}
                        </select>
                        <div style={{ display: "flex", gap: "6px" }}>
                          <button
                            className="db-btn"
                            onClick={() => submitAdd(pos.abbr)}
                            style={{ flex: 1, padding: "6px", color: accent, borderColor: accent }}
                          >
                            Add to board
                          </button>
                          <button className="db-btn" onClick={() => setAddOpenFor(null)} style={{ padding: "6px 10px" }}>
                            ×
                          </button>
                        </div>
                      </div>
                    ) : (
                      <button
                        className="db-btn"
                        onClick={() => openAdd(pos.abbr)}
                        style={{ width: "100%", padding: "8px", fontSize: "12px" }}
                      >
                        + Add prospect
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
