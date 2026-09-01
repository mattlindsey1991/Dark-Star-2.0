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

const YEARS = [2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036, 2037];
const ENTRY_YEARS = [2021, 2022, 2023, 2024, 2025, 2026];
const GRADE_SCALE = [1.0, 1.2, 1.5, 1.8, 2.0, 2.2, 2.5, 2.8, 3.3, 3.8, 4.3, 4.8, 5.3, 5.8, 6.3, 6.8, 7.3, 7.8, 8.0, 8.5, 9.0];

const AGENT_INITIALS = ["TA", "KA", "LA", "EB", "BB", "DD", "JF", "TF", "CH", "RH", "CHud", "AK", "SK", "JL", "AL", "KM", "JM", "DM", "BM", "JP", "TP", "BR", "CS", "TS", "AS", "CW", "RW", "TD", "MD", "JS", "DP", "DJ"];

const SCOUT_TEAMS = ["A1", "ARZ", "ATL", "BAL", "BLESTO", "BUF", "CAR", "CHI", "CIN", "CLV", "DAL", "DEN", "DET", "GB", "HOU", "IND", "JAX", "KC", "LAC", "LAR", "LVR", "MIA", "MIN", "NE", "NFS", "NO", "NYG", "NYJ", "PHI", "PIT", "SEA", "SF", "TB", "TEN", "WAS"];
const SCOUT_MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
const SCOUT_YEARS = [2024, 2025, 2026, 2027, 2028, 2029, 2030];

const VET_DRAFT_YEARS = [];
for (let y = 2031; y >= 2000; y--) VET_DRAFT_YEARS.push(y);

const FA_YEARS = [];
for (let y = 2025; y <= 2035; y++) FA_YEARS.push(y);

const PROJECTED_VALUES = ["$5m+", "$10m+", "$15m+", "$20m+", "$25m+", "$30m+", "$35m+", "$40m+", "$45m+", "$50m+"];

const NFL_TEAMS = [
  "ARZ", "ATL", "BAL", "BUF", "CAR", "CHI", "CIN", "CLE", "DAL", "DEN",
  "DET", "FA", "GB", "HST", "IND", "JAX", "KC", "LAC", "LAR", "LVR",
  "MIA", "MIN", "NE", "NO", "NYG", "NYJ", "PHI", "PIT", "SEA", "SF",
  "TB", "TEN", "WAS",
];

const DRAFT_ROUNDS = ["1", "2", "3", "4", "5", "6", "7", "UDFA"];

const BB_POSITIONS = [
  { abbr: "PG", name: "Point guard" },
  { abbr: "SG", name: "Shooting guard" },
  { abbr: "SF", name: "Small forward" },
  { abbr: "PF", name: "Power forward" },
  { abbr: "C", name: "Center" },
];
const BB_YEARS = ["HS-2027", "HS-2028", "HS-2029", "HS-2030", "HS-2031", "College-NIL", "College-Pro", "Pro", "W-NIL"];
const BB_PRIORITIES = ["High", "Medium", "Low"];
const BB_STATUSES = ["Target", "Evaluating", "Contacted", "Warm", "Signed"];

const POSITION_BOARD = {};
OFFENSE_POSITIONS.forEach((p) => (POSITION_BOARD[p.abbr] = "OFFENSE"));
DEFENSE_POSITIONS.forEach((p) => (POSITION_BOARD[p.abbr] = "DEFENSE"));
const ALL_POSITIONS = [...OFFENSE_POSITIONS, ...DEFENSE_POSITIONS];
const KNOWN_IMPORT_COLUMNS = new Set(["name", "position", "school", "entry year", "entryyear", "agents"]);
const KNOWN_VET_IMPORT_COLUMNS = new Set([
  "name", "position", "hometown", "draft year", "draftyear",
  "free agency year", "freeagencyyear", "fa year", "fayear",
  "projected value", "projected $$$", "projectedvalue",
  "current agent", "currentagent", "current agency", "currentagency",
  "assigned agent", "assignedagent",
  "date of birth", "dob", "dateofbirth",
  "meetings", "notes",
]);

function parseExcelDate(val) {
  if (val === null || val === undefined || val === "") return null;
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  if (typeof val === "number") {
    const utcDays = Math.floor(val - 25569);
    const date = new Date(utcDays * 86400 * 1000);
    return isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
  }
  const parsed = new Date(String(val).trim());
  return isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}

function normalizeProjectedValue(raw) {
  if (!raw) return null;
  let v = String(raw).trim();
  if (!v) return null;
  if (PROJECTED_VALUES.includes(v)) return v;
  const match = v.match(/\d+/);
  if (!match) return null;
  const candidate = `$${match[0]}m+`;
  return PROJECTED_VALUES.includes(candidate) ? candidate : null;
}

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
  vetGreen: "#2E7D4F",
  vetGreenText: "#EAF6EC",
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

function projectedValueNum(v) {
  if (!v) return -1;
  const match = String(v).match(/\d+/);
  return match ? parseInt(match[0], 10) : -1;
}

function priorityRank(p) {
  if (p === "High") return 0;
  if (p === "Medium") return 1;
  if (p === "Low") return 2;
  return 3;
}

function priorityColor(p) {
  if (p === "High") return { bg: "#4C9A5B", text: "#EAF6EC" };
  if (p === "Medium") return { bg: "#D9B23C", text: "#2B2000" };
  if (p === "Low") return { bg: "#C24E4E", text: "#FCEDED" };
  return null;
}

function calcAge(dob) {
  if (!dob) return null;
  const birth = new Date(dob + "T00:00:00");
  if (isNaN(birth.getTime())) return null;
  const now = new Date();
  const years = (now - birth) / (365.25 * 24 * 3600 * 1000);
  return years.toFixed(1);
}

export default function DraftBoard({ session }) {
  const [prospects, setProspects] = useState([]);
  const [vets, setVets] = useState([]);
  const [bbProspects, setBbProspects] = useState([]);
  const [sport, setSport] = useState("FOOTBALL");
  const [bbYear, setBbYear] = useState("HS-2027");
  const [loaded, setLoaded] = useState(false);
  const [board, setBoard] = useState("OFFENSE");
  const [year, setYear] = useState(2027);
  const [search, setSearch] = useState("");
  const [expandedId, setExpandedId] = useState(null);
  const [addOpenFor, setAddOpenFor] = useState(null);
  const [addDraft, setAddDraft] = useState({ name: "", school: "", entryYear: 2024 });
  const [bbAddDraft, setBbAddDraft] = useState({ name: "", team: "" });
  const [vetDraft, setVetDraft] = useState({ name: "", hometown: "", draftYear: 2024 });
  const [gradeDraft, setGradeDraft] = useState({ team: "", scout: "", month: "", year: "", grade: GRADE_SCALE[0] });
  const [bbGradeDraft, setBbGradeDraft] = useState({ scout: "", grade: GRADE_SCALE[0] });
  const [errorMsg, setErrorMsg] = useState("");
  const [importing, setImporting] = useState(false);
  const [importSummary, setImportSummary] = useState(null);
  const [pwOpen, setPwOpen] = useState(false);
  const [pwDraft, setPwDraft] = useState({ pw1: "", pw2: "" });
  const [pwMsg, setPwMsg] = useState("");
  const fileInputRef = useRef(null);
  const vetFileInputRef = useRef(null);
  const bbFileInputRef = useRef(null);

  const isVetView = year === "VET" && sport === "FOOTBALL";
  const isBasketball = sport === "BASKETBALL";

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

  const fetchVets = useCallback(async () => {
    const { data, error } = await supabase
      .from("vets")
      .select("*")
      .order("created_at", { ascending: true });
    if (error) {
      setErrorMsg("Couldn't load the vet board. Try refreshing.");
    } else {
      setVets(data);
    }
  }, []);

  const fetchBB = useCallback(async () => {
    const { data, error } = await supabase
      .from("bb_prospects")
      .select("*, bb_grades(*)")
      .order("created_at", { ascending: true });
    if (error) {
      setErrorMsg("Couldn't load the basketball board. Try refreshing.");
    } else {
      setBbProspects(data);
    }
  }, []);

  useEffect(() => {
    fetchAll();
    fetchVets();
    fetchBB();

    const channel = supabase
      .channel("draft-board-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "prospects" }, fetchAll)
      .on("postgres_changes", { event: "*", schema: "public", table: "grades" }, fetchAll)
      .on("postgres_changes", { event: "*", schema: "public", table: "vets" }, fetchVets)
      .on("postgres_changes", { event: "*", schema: "public", table: "bb_prospects" }, fetchBB)
      .on("postgres_changes", { event: "*", schema: "public", table: "bb_grades" }, fetchBB)
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, [fetchAll, fetchVets, fetchBB]);

  const positions = isBasketball ? BB_POSITIONS : (board === "OFFENSE" ? OFFENSE_POSITIONS : DEFENSE_POSITIONS);

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

  const groupedVets = useMemo(() => {
    const q = search.trim().toLowerCase();
    const map = {};
    positions.forEach((p) => (map[p.abbr] = []));
    vets
      .filter((v) => map[v.position] !== undefined)
      .filter(
        (v) =>
          !q ||
          v.name.toLowerCase().includes(q) ||
          (v.hometown || "").toLowerCase().includes(q)
      )
      .forEach((v) => map[v.position].push(v));
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => projectedValueNum(b.projected_value) - projectedValueNum(a.projected_value));
    });
    return map;
  }, [vets, positions, search]);

  const bbGrouped = useMemo(() => {
    const q = search.trim().toLowerCase();
    const map = {};
    BB_POSITIONS.forEach((p) => (map[p.abbr] = []));
    bbProspects
      .filter((bp) => bp.class_year === bbYear && map[bp.position] !== undefined)
      .filter(
        (bp) =>
          !q ||
          bp.name.toLowerCase().includes(q) ||
          (bp.team || "").toLowerCase().includes(q)
      )
      .forEach((bp) => map[bp.position].push(bp));
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => priorityRank(a.priority) - priorityRank(b.priority));
    });
    return map;
  }, [bbProspects, bbYear, search]);

  const totalCount = useMemo(() => {
    const src = isBasketball ? bbGrouped : (isVetView ? groupedVets : grouped);
    return Object.values(src).reduce((a, list) => a + list.length, 0);
  }, [grouped, groupedVets, bbGrouped, isVetView, isBasketball]);

  function openAdd(posAbbr) {
    setAddOpenFor(posAbbr);
    setAddDraft({ name: "", school: "", entryYear: 2024 });
    setVetDraft({ name: "", hometown: "", draftYear: 2024 });
    setBbAddDraft({ name: "", team: "" });
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

  async function submitAddVet(posAbbr) {
    if (!vetDraft.name.trim()) return;
    const { data, error } = await supabase
      .from("vets")
      .insert({
        name: vetDraft.name.trim(),
        position: posAbbr,
        board,
        hometown: vetDraft.hometown.trim(),
        draft_year: Number(vetDraft.draftYear),
        created_by: session.user.id,
      })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't add that player. Try again.");
      return;
    }
    setVets((prev) => [...prev, data]);
    setAddOpenFor(null);
    setExpandedId(data.id);
  }

  async function submitAddBB(posAbbr) {
    if (!bbAddDraft.name.trim()) return;
    const { data, error } = await supabase
      .from("bb_prospects")
      .insert({
        name: bbAddDraft.name.trim(),
        position: posAbbr,
        class_year: bbYear,
        team: bbAddDraft.team.trim(),
        created_by: session.user.id,
      })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't add that prospect. Try again.");
      return;
    }
    setBbProspects((prev) => [...prev, { ...data, bb_grades: [] }]);
    setAddOpenFor(null);
    setExpandedId(data.id);
  }

  async function updateBBProspect(id, patch) {
    setBbProspects((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
    const { error } = await supabase.from("bb_prospects").update(patch).eq("id", id);
    if (error) setErrorMsg("That change didn't save. Try again.");
  }

  async function deleteBBProspect(id) {
    setBbProspects((prev) => prev.filter((p) => p.id !== id));
    if (expandedId === id) setExpandedId(null);
    const { error } = await supabase.from("bb_prospects").delete().eq("id", id);
    if (error) setErrorMsg("Couldn't remove that prospect. Try again.");
  }

  async function addBBGrade(id) {
    const g = parseFloat(bbGradeDraft.grade);
    if (!bbGradeDraft.scout.trim() || isNaN(g) || !GRADE_SCALE.includes(g)) return;
    const { data, error } = await supabase
      .from("bb_grades")
      .insert({ prospect_id: id, scout: bbGradeDraft.scout.trim(), grade: g, created_by: session.user.id })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't save that grade. Try again.");
      return;
    }
    setBbProspects((prev) =>
      prev.map((p) => (p.id === id ? { ...p, bb_grades: [...(p.bb_grades || []), data] } : p))
    );
    setBbGradeDraft({ scout: "", grade: GRADE_SCALE[0] });
  }

  async function deleteBBGrade(prospectId, gradeId) {
    setBbProspects((prev) =>
      prev.map((p) =>
        p.id === prospectId ? { ...p, bb_grades: p.bb_grades.filter((g) => g.id !== gradeId) } : p
      )
    );
    const { error } = await supabase.from("bb_grades").delete().eq("id", gradeId);
    if (error) setErrorMsg("Couldn't remove that grade. Try again.");
  }

  async function updateProspect(id, patch) {
    setProspects((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
    const { error } = await supabase.from("prospects").update(patch).eq("id", id);
    if (error) setErrorMsg("That change didn't save. Try again.");
  }

  async function updateVet(id, patch) {
    setVets((prev) => prev.map((v) => (v.id === id ? { ...v, ...patch } : v)));
    const { error } = await supabase.from("vets").update(patch).eq("id", id);
    if (error) setErrorMsg("That change didn't save. Try again.");
  }

  async function deleteProspect(id) {
    setProspects((prev) => prev.filter((p) => p.id !== id));
    if (expandedId === id) setExpandedId(null);
    const { error } = await supabase.from("prospects").delete().eq("id", id);
    if (error) setErrorMsg("Couldn't remove that prospect. Try again.");
  }

  async function deleteVet(id) {
    setVets((prev) => prev.filter((v) => v.id !== id));
    if (expandedId === id) setExpandedId(null);
    const { error } = await supabase.from("vets").delete().eq("id", id);
    if (error) setErrorMsg("Couldn't remove that player. Try again.");
  }

  async function addGrade(id) {
    const g = parseFloat(gradeDraft.grade);
    if (!gradeDraft.scout.trim() || isNaN(g) || !GRADE_SCALE.includes(g)) return;
    const { data, error } = await supabase
      .from("grades")
      .insert({
        prospect_id: id,
        scout: gradeDraft.scout.trim(),
        team: gradeDraft.team || null,
        scout_name: gradeDraft.scout.trim(),
        month: gradeDraft.month ? Number(gradeDraft.month) : null,
        year: gradeDraft.year ? Number(gradeDraft.year) : null,
        grade: g,
        created_by: session.user.id,
      })
      .select()
      .single();
    if (error) {
      setErrorMsg("Couldn't save that grade. Try again.");
      return;
    }
    setProspects((prev) =>
      prev.map((p) => (p.id === id ? { ...p, grades: [...p.grades, data] } : p))
    );
    setGradeDraft({ team: "", scout: "", month: "", year: "", grade: GRADE_SCALE[0] });
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
        const agent1Key = keyMap["agent 1"] || keyMap["agent1"];
        const agent2Key = keyMap["agent 2"] || keyMap["agent2"];
        const agent3Key = keyMap["agent 3"] || keyMap["agent3"];
        const otherAgencyKey = keyMap["other agency"] || keyMap["otheragency"];

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
          agent_1: agent1Key ? String(row[agent1Key]).trim() : "",
          agent_2: agent2Key ? String(row[agent2Key]).trim() : "",
          agent_3: agent3Key ? String(row[agent3Key]).trim() : "",
          other_agency: otherAgencyKey ? String(row[otherAgencyKey]).trim() : "",
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

  async function handleImportVetFile(e) {
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
      const vetRows = [];

      rows.forEach((row, idx) => {
        const keys = Object.keys(row);
        const keyMap = {};
        keys.forEach((k) => (keyMap[normalizeHeader(k)] = k));

        const nameKey = keyMap["name"];
        const posKey = keyMap["position"];
        const hometownKey = keyMap["hometown"];
        const teamKey = keyMap["current team"] || keyMap["currentteam"] || keyMap["team"] || keyMap["nfl team"];
        const collegeKey = keyMap["college"];
        const draftRoundKey = keyMap["draft round"] || keyMap["draftround"];
        const draftYearKey = keyMap["draft year"] || keyMap["draftyear"];
        const faYearKey = keyMap["free agency year"] || keyMap["freeagencyyear"] || keyMap["fa year"] || keyMap["fayear"];
        const projValKey = keyMap["projected value"] || keyMap["projected $$$"] || keyMap["projectedvalue"];
        const currentAgentKey = keyMap["current agent"] || keyMap["currentagent"];
        const currentAgencyKey = keyMap["current agency"] || keyMap["currentagency"];
        const assignedAgentKey = keyMap["assigned agent"] || keyMap["assignedagent"];
        const dobKey = keyMap["date of birth"] || keyMap["dob"] || keyMap["dateofbirth"];
        const meetingsKey = keyMap["meetings"];
        const notesKey = keyMap["notes"];

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

        const draftYearNum = draftYearKey ? parseInt(row[draftYearKey], 10) : NaN;
        const faYearNum = faYearKey ? parseInt(row[faYearKey], 10) : NaN;

        vetRows.push({
          name,
          position: positionRaw,
          board: rowBoard,
          hometown: hometownKey ? String(row[hometownKey]).trim() : "",
          current_team: teamKey ? String(row[teamKey]).trim().toUpperCase() : null,
          draft_round: draftRoundKey ? String(row[draftRoundKey]).trim() : null,
          college: collegeKey ? String(row[collegeKey]).trim() : "",
          draft_year: Number.isFinite(draftYearNum) ? draftYearNum : null,
          free_agency_year: Number.isFinite(faYearNum) ? faYearNum : null,
          projected_value: projValKey ? normalizeProjectedValue(row[projValKey]) : null,
          current_agent: currentAgentKey ? String(row[currentAgentKey]).trim() : "",
          current_agency: currentAgencyKey ? String(row[currentAgencyKey]).trim() : "",
          assigned_agent: assignedAgentKey ? String(row[assignedAgentKey]).trim() : "",
          date_of_birth: dobKey ? parseExcelDate(row[dobKey]) : null,
          meetings: meetingsKey ? String(row[meetingsKey]).trim() : "",
          notes: notesKey ? String(row[notesKey]).trim() : "",
          created_by: session.user.id,
        });
      });

      if (vetRows.length > 0) {
        const { error: vetErr } = await supabase.from("vets").insert(vetRows);
        if (vetErr) {
          setErrorMsg("Import failed while saving players: " + vetErr.message);
          setImporting(false);
          e.target.value = "";
          return;
        }
      }

      await fetchVets();
      setImportSummary({
        kind: "vet",
        prospectCount: vetRows.length,
        gradeCount: 0,
        skipped,
        year: "VET",
      });
    } catch (err) {
      setErrorMsg("Couldn't read that file. Make sure it's a valid .xlsx or .csv.");
    }

    setImporting(false);
    e.target.value = "";
  }

  const BB_POSITION_SET = new Set(BB_POSITIONS.map((p) => p.abbr));

  async function handleImportBBFile(e) {
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
      const bbRows = [];

      rows.forEach((row, idx) => {
        const keys = Object.keys(row);
        const keyMap = {};
        keys.forEach((k) => (keyMap[normalizeHeader(k)] = k));

        const nameKey = keyMap["name"];
        const posKey = keyMap["position"];
        const teamKey = keyMap["team"];
        const classYearKey = keyMap["class year"] || keyMap["classyear"] || keyMap["category"];
        const aauKey = keyMap["aau program"] || keyMap["aauprogram"] || keyMap["aau"];
        const yearKey = keyMap["year"];
        const primaryRecruiterKey = keyMap["primary recruiter"] || keyMap["primaryrecruiter"];
        const secondaryRecruiterKey = keyMap["secondary recruiter"] || keyMap["secondaryrecruiter"];
        const contactKey = keyMap["contact"];
        const priorityKey = keyMap["priority"];
        const statusKey = keyMap["status"];
        const lastContactKey = keyMap["last contact date"] || keyMap["lastcontactdate"];
        const notesKey = keyMap["notes"];

        const name = nameKey ? String(row[nameKey]).trim() : "";
        const positionRaw = posKey ? String(row[posKey]).trim().toUpperCase() : "";

        if (!name) {
          skipped.push({ row: idx + 2, reason: "Missing name" });
          return;
        }
        if (!BB_POSITION_SET.has(positionRaw)) {
          skipped.push({ row: idx + 2, reason: `Unrecognized position "${positionRaw}" for ${name}` });
          return;
        }

        const classYearRaw = classYearKey ? String(row[classYearKey]).trim() : "";
        const classYearMatch = BB_YEARS.find((y) => y.toLowerCase() === classYearRaw.toLowerCase());
        const priorityRaw = priorityKey ? String(row[priorityKey]).trim() : "";
        const priorityMatch = BB_PRIORITIES.find((pr) => pr.toLowerCase() === priorityRaw.toLowerCase());
        const statusRaw = statusKey ? String(row[statusKey]).trim() : "";
        const statusMatch = BB_STATUSES.find((st) => st.toLowerCase() === statusRaw.toLowerCase());

        bbRows.push({
          name,
          position: positionRaw,
          class_year: classYearMatch || (classYearRaw ? classYearRaw : bbYear),
          team: teamKey ? String(row[teamKey]).trim() : "",
          aau_program: aauKey ? String(row[aauKey]).trim() : "",
          year: yearKey ? String(row[yearKey]).trim() : "",
          primary_recruiter: primaryRecruiterKey ? String(row[primaryRecruiterKey]).trim() : "",
          secondary_recruiter: secondaryRecruiterKey ? String(row[secondaryRecruiterKey]).trim() : "",
          contact: contactKey ? String(row[contactKey]).trim() : "",
          priority: priorityMatch || null,
          status: statusMatch || null,
          last_contact_date: lastContactKey ? parseExcelDate(row[lastContactKey]) : null,
          notes: notesKey ? String(row[notesKey]).trim() : "",
          created_by: session.user.id,
        });
      });

      if (bbRows.length > 0) {
        const { error: bbErr } = await supabase.from("bb_prospects").insert(bbRows);
        if (bbErr) {
          setErrorMsg("Import failed while saving prospects: " + bbErr.message);
          setImporting(false);
          e.target.value = "";
          return;
        }
      }

      await fetchBB();
      setImportSummary({
        kind: "bb",
        prospectCount: bbRows.length,
        gradeCount: 0,
        skipped,
        year: bbYear,
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

  function handleExportExcel() {
    const rows = [];
    if (isBasketball) {
      positions.forEach((pos) => {
        (bbGrouped[pos.abbr] || []).forEach((p) => {
          const avg = computeAvg(p.bb_grades);
          rows.push({
            Name: p.name,
            Position: p.position,
            "Class Year": p.class_year,
            Team: p.team || "",
            "AAU Program": p.aau_program || "",
            Year: p.year || "",
            "Primary Recruiter": p.primary_recruiter || "",
            "Secondary Recruiter": p.secondary_recruiter || "",
            Contact: p.contact || "",
            Priority: p.priority || "",
            Status: p.status || "",
            "Last Contact Date": p.last_contact_date || "",
            "A1 Client": p.is_a1 ? "Yes" : "No",
            "Avg Grade": avg !== null ? avg.toFixed(1) : "",
            Grades: (p.bb_grades || [])
              .map((g) => `${g.scout}: ${Number(g.grade).toFixed(1)}`)
              .join("; "),
            Notes: p.notes || "",
          });
        });
      });
    } else if (isVetView) {
      positions.forEach((pos) => {
        (groupedVets[pos.abbr] || []).forEach((v) => {
          rows.push({
            Name: v.name,
            Position: v.position,
            "Current Team": v.current_team || "",
            "Draft Round": v.draft_round || "",
            College: v.college || "",
            "Draft Year": v.draft_year || "",
            "Free Agency Year": v.free_agency_year || "",
            Hometown: v.hometown || "",
            "Current Agent": v.current_agent || "",
            "Current Agency": v.current_agency || "",
            "Assigned Agent": v.assigned_agent || "",
            "Date of Birth": v.date_of_birth || "",
            Age: calcAge(v.date_of_birth) || "",
            "Projected Value": v.projected_value || "",
            "A1 Client": v.is_a1 ? "Yes" : "No",
            Meetings: v.meetings || "",
            Notes: v.notes || "",
          });
        });
      });
    } else {
      positions.forEach((pos) => {
        (grouped[pos.abbr] || []).forEach((p) => {
          const avg = computeAvg(p.grades);
          rows.push({
            Name: p.name,
            Position: p.position,
            School: p.school || "",
            "Draft Class": p.draft_class_year,
            "Entry Year": p.entry_year || "",
            "Agent 1": p.agent_1 || "",
            "Agent 2": p.agent_2 || "",
            "Agent 3": p.agent_3 || "",
            "Other Agency": p.other_agency || "",
            "A1 Client": p.is_a1 ? "Yes" : "No",
            "Avg Grade": avg !== null ? avg.toFixed(1) : "",
            Grades: (p.grades || [])
              .map((g) => `${g.scout}: ${Number(g.grade).toFixed(1)}`)
              .join("; "),
            Meetings: p.meetings || "",
            Notes: p.notes || "",
          });
        });
      });
    }

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    const sheetName = isBasketball ? `BB ${bbYear}` : (isVetView ? "VET" : `${board} ${year}`);
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    const filename = isBasketball
      ? `BigBoard_BB_${bbYear}.xlsx`
      : isVetView
      ? "BigBoard_VET.xlsx"
      : `BigBoard_${board}_${year}.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  const accent = isBasketball ? "#C9973E" : (board === "OFFENSE" ? COLORS.offense : COLORS.defense);

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
        .db-input option { color: #15171A; background: #ECE7DC; }
        .db-btn {
          background: transparent;
          border: 1px solid ${COLORS.hair};
          color: ${COLORS.inkDim};
          border-radius: 4px;
          cursor: pointer;
          transition: all 0.12s ease;
        }
        .db-btn:hover { border-color: ${COLORS.hairStrong}; color: ${COLORS.ink}; }

        @media print {
          @page { size: landscape; margin: 10mm; }
          .no-print { display: none !important; }
          body, #root, * { background: transparent !important; box-shadow: none !important; }
          body { background: #ffffff !important; }
          * { color: #111111 !important; border-color: #ccc !important; }
          .board-columns { flex-wrap: wrap !important; overflow: visible !important; gap: 8px !important; }
          .board-column { page-break-inside: avoid; min-width: 200px !important; max-width: 220px !important; }
        }
      `}</style>

      <div
        className="no-print"
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
          <div className="no-print" style={{ display: "flex", gap: "8px" }}>
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
            className="no-print"
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
            {["FOOTBALL", "BASKETBALL"].map((s) => (
              <button
                key={s}
                onClick={() => { setSport(s); setExpandedId(null); setAddOpenFor(null); }}
                style={{
                  fontFamily: "'Bebas Neue', sans-serif",
                  fontSize: "15px",
                  letterSpacing: "1px",
                  padding: "8px 16px",
                  border: "none",
                  cursor: "pointer",
                  background: sport === s ? "rgba(201,151,62,0.14)" : "transparent",
                  color: sport === s ? "#C9973E" : COLORS.inkDim,
                }}
              >
                {s}
              </button>
            ))}
          </div>

          {!isBasketball && (
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
          )}

          <div style={{ display: "flex", gap: "6px" }}>
            {isBasketball
              ? BB_YEARS.map((y) => (
                  <button
                    key={y}
                    onClick={() => setBbYear(y)}
                    className="db-btn"
                    style={{
                      fontFamily: "'IBM Plex Mono', monospace",
                      fontSize: "12.5px",
                      padding: "7px 10px",
                      borderColor: bbYear === y ? accent : COLORS.hair,
                      color: bbYear === y ? accent : COLORS.inkDim,
                    }}
                  >
                    {y}
                  </button>
                ))
              : (
                  <select
                    className="db-input"
                    value={typeof year === "number" ? year : ""}
                    onChange={(e) => setYear(Number(e.target.value))}
                    style={{
                      fontFamily: "'IBM Plex Mono', monospace",
                      fontSize: "12.5px",
                      padding: "7px 10px",
                      borderColor: !isVetView ? accent : COLORS.hair,
                      color: !isVetView ? accent : COLORS.inkDim,
                      fontWeight: 700,
                    }}
                  >
                    {YEARS.map((y) => (
                      <option key={y} value={y}>{y}</option>
                    ))}
                  </select>
                )}
            {!isBasketball && (
              <button
                onClick={() => setYear("VET")}
                className="db-btn"
                style={{
                  fontFamily: "'IBM Plex Mono', monospace",
                  fontSize: "12.5px",
                  fontWeight: 700,
                  padding: "7px 10px",
                  borderColor: isVetView ? COLORS.vetGreen : COLORS.hair,
                  color: isVetView ? COLORS.vetGreen : COLORS.inkDim,
                }}
              >
                VET
              </button>
            )}
          </div>

          <input
            className="db-input no-print"
            placeholder={isVetView ? "Search name or hometown" : "Search name or school"}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ width: "180px", marginLeft: "auto" }}
          />

          {isBasketball ? (
            <>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                ref={bbFileInputRef}
                onChange={handleImportBBFile}
                style={{ display: "none" }}
              />
              <button
                className="db-btn no-print"
                onClick={() => bbFileInputRef.current && bbFileInputRef.current.click()}
                disabled={importing}
                title={`Imports into the ${bbYear} class. Columns: Name, Position, Team, Class Year, AAU Program, Year, Primary Recruiter, Secondary Recruiter, Contact, Priority, Status, Last Contact Date, Notes, plus any scout columns.`}
                style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
              >
                {importing ? "Importing…" : "Upload Excel"}
              </button>
            </>
          ) : isVetView ? (
            <>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                ref={vetFileInputRef}
                onChange={handleImportVetFile}
                style={{ display: "none" }}
              />
              <button
                className="db-btn no-print"
                onClick={() => vetFileInputRef.current && vetFileInputRef.current.click()}
                disabled={importing}
                title="Columns: Name, Position, Hometown, Current Team, College, Draft Year, Free Agency Year, Projected Value, Current Agent, Current Agency, Assigned Agent, Date of Birth, Meetings, Notes."
                style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
              >
                {importing ? "Importing…" : "Upload Excel"}
              </button>
            </>
          ) : (
            <>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                ref={fileInputRef}
                onChange={handleImportFile}
                style={{ display: "none" }}
              />
              <button
                className="db-btn no-print"
                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                disabled={importing}
                title={`Imports into the ${year} class. Columns: Name, Position, School, Entry Year, Agents, plus any scout columns.`}
                style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
              >
                {importing ? "Importing…" : "Upload Excel"}
              </button>
            </>
          )}

          <button
            className="db-btn no-print"
            onClick={handleExportExcel}
            title="Exports the players currently shown (this board/year or VET tab) with all their info."
            style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
          >
            Export Excel
          </button>

          <button
            className="db-btn no-print"
            onClick={() => window.print()}
            title="Opens the print dialog. Choose 'Save as PDF' for a printable one-page copy of this view."
            style={{ padding: "7px 12px", fontSize: "12px", whiteSpace: "nowrap" }}
          >
            Print / PDF
          </button>

          <div style={{ display: "flex", alignItems: "center", gap: "6px", fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.inkDim }}>
            {totalCount} {isVetView ? "players" : "prospects"} · {isBasketball ? bbYear : year}
          </div>
        </div>

        {importSummary && (
          <div
            className="no-print"
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
              {importSummary.kind === "vet" ? (
                <>
                  Imported {importSummary.prospectCount} player{importSummary.prospectCount === 1 ? "" : "s"} to the VET board.
                </>
              ) : importSummary.kind === "bb" ? (
                <>
                  Imported {importSummary.prospectCount} prospect{importSummary.prospectCount === 1 ? "" : "s"} into the {importSummary.year} basketball class.
                </>
              ) : (
                <>
                  Imported {importSummary.prospectCount} prospect{importSummary.prospectCount === 1 ? "" : "s"} and{" "}
                  {importSummary.gradeCount} grade{importSummary.gradeCount === 1 ? "" : "s"} into the {importSummary.year} class.
                </>
              )}
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

        {isBasketball ? null : isVetView ? (
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "16px" }}>
            <span style={{ width: "9px", height: "9px", borderRadius: "3px", background: COLORS.vetGreen, border: "1px solid rgba(255,255,255,0.15)" }} />
            <span style={{ fontSize: "11px", color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace" }}>
              SORTED BY PROJECTED VALUE, HIGH TO LOW
            </span>
          </div>
        ) : (
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
        )}

        {errorMsg && (
          <div style={{ fontSize: "12px", color: "#D98080", marginBottom: "12px", fontFamily: "'IBM Plex Mono', monospace" }}>
            {errorMsg}
          </div>
        )}

        {!loaded ? (
          <div style={{ color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace", fontSize: "13px" }}>Loading board…</div>
        ) : (
          <div className="board-columns" style={{ display: "flex", gap: "14px", overflowX: "auto", paddingBottom: "12px" }}>
            {positions.map((pos) => {
              const list = isBasketball ? (bbGrouped[pos.abbr] || []) : isVetView ? (groupedVets[pos.abbr] || []) : (grouped[pos.abbr] || []);
              return (
                <div
                  key={pos.abbr}
                  className="board-column"
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
                        {isBasketball ? `No prospects logged for ${bbYear} yet.` : isVetView ? "No players logged yet." : `No prospects logged for ${year} yet.`}
                      </div>
                    )}

                    {!isBasketball && !isVetView && list.map((p, idx) => {
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
                                {(() => {
                                  const initials = [p.agent_1, p.agent_2, p.agent_3].filter(Boolean).join(", ");
                                  if (!initials && !p.other_agency) return "No agent listed";
                                  return (
                                    <>
                                      {initials}
                                      {p.other_agency && (
                                        <span style={{ color: "#D98A3E" }}>
                                          {initials ? " · " : ""}{p.other_agency}
                                        </span>
                                      )}
                                    </>
                                  );
                                })()}
                              </div>
                            </div>
                            {p.is_a1 && (
                              <span
                                style={{
                                  fontFamily: "'IBM Plex Mono', monospace",
                                  fontSize: "11px",
                                  fontWeight: 800,
                                  color: "#E24C4C",
                                  border: "1.5px solid #E24C4C",
                                  borderRadius: "4px",
                                  padding: "2px 5px",
                                  flexShrink: 0,
                                  letterSpacing: "0.5px",
                                }}
                              >
                                A1
                              </span>
                            )}
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
                            <div className="no-print" style={{ padding: "4px 14px 14px", background: COLORS.surfaceHi }}>
                              <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "10px" }}>
                                <span style={{ width: "9px", height: "9px", borderRadius: "50%", background: tier.color, border: tier.filled ? "1px solid rgba(255,255,255,0.15)" : "none", flexShrink: 0 }} />
                                <span style={{ fontSize: "10.5px", color: COLORS.inkDim, fontFamily: "'IBM Plex Mono', monospace" }}>
                                  {tier.label.toUpperCase()}
                                </span>
                              </div>

                              <label
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: "7px",
                                  marginBottom: "10px",
                                  cursor: "pointer",
                                  fontSize: "12px",
                                  color: p.is_a1 ? "#E24C4C" : COLORS.inkDim,
                                  fontWeight: p.is_a1 ? 700 : 400,
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={!!p.is_a1}
                                  onChange={(e) => {
                                    e.stopPropagation();
                                    updateProspect(p.id, { is_a1: e.target.checked });
                                  }}
                                  onClick={(e) => e.stopPropagation()}
                                  style={{ accentColor: "#E24C4C", width: "14px", height: "14px" }}
                                />
                                Mark as A1 client
                              </label>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Name</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px", fontWeight: 600 }}
                                defaultValue={p.name}
                                onBlur={(e) => {
                                  const v = e.target.value.trim();
                                  if (v) updateProspect(p.id, { name: v });
                                  else e.target.value = p.name;
                                }}
                              />

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

                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, marginBottom: "3px" }}>A1 Agents</div>
                              <div style={{ display: "flex", gap: "4px", marginBottom: "8px" }}>
                                <select
                                  className="db-input"
                                  style={{ flex: 1, width: 0, minWidth: 0, fontSize: "11px", padding: "5px 3px" }}
                                  value={p.agent_1 || ""}
                                  onChange={(e) => updateProspect(p.id, { agent_1: e.target.value || null })}
                                >
                                  <option value="">—</option>
                                  {AGENT_INITIALS.map((a) => (
                                    <option key={a} value={a}>{a}</option>
                                  ))}
                                </select>
                                <select
                                  className="db-input"
                                  style={{ flex: 1, width: 0, minWidth: 0, fontSize: "11px", padding: "5px 3px" }}
                                  value={p.agent_2 || ""}
                                  onChange={(e) => updateProspect(p.id, { agent_2: e.target.value || null })}
                                >
                                  <option value="">—</option>
                                  {AGENT_INITIALS.map((a) => (
                                    <option key={a} value={a}>{a}</option>
                                  ))}
                                </select>
                                <select
                                  className="db-input"
                                  style={{ flex: 1, width: 0, minWidth: 0, fontSize: "11px", padding: "5px 3px" }}
                                  value={p.agent_3 || ""}
                                  onChange={(e) => updateProspect(p.id, { agent_3: e.target.value || null })}
                                >
                                  <option value="">—</option>
                                  {AGENT_INITIALS.map((a) => (
                                    <option key={a} value={a}>{a}</option>
                                  ))}
                                </select>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Other agency</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "10px" }}
                                placeholder="e.g. Rosenhaus, CAA, Excel"
                                defaultValue={p.other_agency}
                                onBlur={(e) => updateProspect(p.id, { other_agency: e.target.value || null })}
                              />

                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, marginBottom: "5px" }}>
                                Scout grades ({p.grades.length})
                              </div>
                              {p.grades.map((g) => (
                                <div key={g.id} style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", marginBottom: "4px" }}>
                                  <span style={{ flex: 1, color: COLORS.ink, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                    {[g.team, g.scout_name || g.scout].filter(Boolean).join(" · ")}
                                    {g.month && g.year && (
                                      <span style={{ color: COLORS.inkDim }}> ({g.month}/{g.year})</span>
                                    )}
                                  </span>
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
                              <div style={{ marginTop: "6px" }} onClick={(e) => e.stopPropagation()}>
                                <div style={{ display: "flex", gap: "4px", marginBottom: "4px" }}>
                                  <select
                                    className="db-input"
                                    style={{ width: "62px", fontSize: "11px", padding: "5px 3px" }}
                                    value={gradeDraft.team}
                                    onChange={(e) => setGradeDraft({ ...gradeDraft, team: e.target.value })}
                                  >
                                    <option value="">Team</option>
                                    {SCOUT_TEAMS.map((t) => (
                                      <option key={t} value={t}>{t}</option>
                                    ))}
                                  </select>
                                  <input
                                    className="db-input"
                                    placeholder="Scout name"
                                    style={{ flex: 1, width: 0 }}
                                    value={gradeDraft.scout}
                                    onChange={(e) => setGradeDraft({ ...gradeDraft, scout: e.target.value })}
                                  />
                                </div>
                                <div style={{ display: "flex", gap: "4px" }}>
                                  <select
                                    className="db-input"
                                    style={{ width: "50px", fontSize: "11px", padding: "5px 3px" }}
                                    value={gradeDraft.month}
                                    onChange={(e) => setGradeDraft({ ...gradeDraft, month: e.target.value })}
                                  >
                                    <option value="">Mo</option>
                                    {SCOUT_MONTHS.map((m) => (
                                      <option key={m} value={m}>{m}</option>
                                    ))}
                                  </select>
                                  <select
                                    className="db-input"
                                    style={{ width: "62px", fontSize: "11px", padding: "5px 3px" }}
                                    value={gradeDraft.year}
                                    onChange={(e) => setGradeDraft({ ...gradeDraft, year: e.target.value })}
                                  >
                                    <option value="">Yr</option>
                                    {SCOUT_YEARS.map((y) => (
                                      <option key={y} value={y}>{y}</option>
                                    ))}
                                  </select>
                                  <select
                                    className="db-input"
                                    style={{ width: "58px", fontSize: "11px", padding: "5px 3px" }}
                                    value={gradeDraft.grade}
                                    onChange={(e) => setGradeDraft({ ...gradeDraft, grade: e.target.value })}
                                  >
                                    {GRADE_SCALE.map((g) => (
                                      <option key={g} value={g}>{g.toFixed(1)}</option>
                                    ))}
                                  </select>
                                  <button className="db-btn" onClick={() => addGrade(p.id)} style={{ padding: "0 8px", flex: 1 }}>
                                    +
                                  </button>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginTop: "10px", marginBottom: "3px" }}>Meetings</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                placeholder="e.g. Combine, 3/12"
                                defaultValue={p.meetings}
                                onBlur={(e) => updateProspect(p.id, { meetings: e.target.value })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Notes</label>
                              <textarea
                                className="db-input"
                                style={{ width: "100%", minHeight: "70px", resize: "vertical", fontFamily: "'Inter', sans-serif" }}
                                defaultValue={p.notes}
                                onBlur={(e) => updateProspect(p.id, { notes: e.target.value })}
                              />

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

                    {isBasketball && list.map((p, idx) => {
                      const pColor = priorityColor(p.priority);
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
                                {p.team || "Team unset"}
                              </div>
                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, opacity: 0.75, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {[p.status, p.priority ? `${p.priority} priority` : null].filter(Boolean).join(" · ") || "No status set"}
                              </div>
                            </div>
                            {p.is_a1 && (
                              <span
                                style={{
                                  fontFamily: "'IBM Plex Mono', monospace",
                                  fontSize: "11px",
                                  fontWeight: 800,
                                  color: "#E24C4C",
                                  border: "1.5px solid #E24C4C",
                                  borderRadius: "4px",
                                  padding: "2px 5px",
                                  flexShrink: 0,
                                  letterSpacing: "0.5px",
                                }}
                              >
                                A1
                              </span>
                            )}
                            <div
                              title={p.priority ? `${p.priority} priority` : "No priority set"}
                              style={{
                                fontFamily: "'IBM Plex Mono', monospace",
                                fontSize: "11px",
                                fontWeight: 700,
                                color: pColor ? pColor.text : COLORS.inkDim,
                                background: pColor ? pColor.bg : "transparent",
                                border: `1.5px solid ${pColor ? "rgba(255,255,255,0.15)" : COLORS.ungraded}`,
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
                              {p.priority ? p.priority[0] : "—"}
                            </div>
                          </div>

                          {isOpen && (
                            <div className="no-print" style={{ padding: "4px 14px 14px", background: COLORS.surfaceHi }}>
                              <label
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: "7px",
                                  marginBottom: "10px",
                                  marginTop: "8px",
                                  cursor: "pointer",
                                  fontSize: "12px",
                                  color: p.is_a1 ? "#E24C4C" : COLORS.inkDim,
                                  fontWeight: p.is_a1 ? 700 : 400,
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={!!p.is_a1}
                                  onChange={(e) => {
                                    e.stopPropagation();
                                    updateBBProspect(p.id, { is_a1: e.target.checked });
                                  }}
                                  onClick={(e) => e.stopPropagation()}
                                  style={{ accentColor: "#E24C4C", width: "14px", height: "14px" }}
                                />
                                Mark as A1 client
                              </label>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Name</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px", fontWeight: 600 }}
                                defaultValue={p.name}
                                onBlur={(e) => {
                                  const v = e.target.value.trim();
                                  if (v) updateBBProspect(p.id, { name: v });
                                  else e.target.value = p.name;
                                }}
                              />

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Position</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.position}
                                    onChange={(e) => updateBBProspect(p.id, { position: e.target.value })}
                                  >
                                    {BB_POSITIONS.map((bp) => (
                                      <option key={bp.abbr} value={bp.abbr}>{bp.abbr}</option>
                                    ))}
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Class / category</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.class_year}
                                    onChange={(e) => updateBBProspect(p.id, { class_year: e.target.value })}
                                  >
                                    {BB_YEARS.map((y) => (
                                      <option key={y} value={y}>{y}</option>
                                    ))}
                                  </select>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Team</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                defaultValue={p.team}
                                onBlur={(e) => updateBBProspect(p.id, { team: e.target.value })}
                              />

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>AAU program</label>
                                  <input
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    defaultValue={p.aau_program}
                                    onBlur={(e) => updateBBProspect(p.id, { aau_program: e.target.value })}
                                  />
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Year</label>
                                  <input
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    placeholder="e.g. Fr, Soph, Jr, Sr"
                                    defaultValue={p.year}
                                    onBlur={(e) => updateBBProspect(p.id, { year: e.target.value })}
                                  />
                                </div>
                              </div>

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Primary recruiter</label>
                                  <input
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    defaultValue={p.primary_recruiter}
                                    onBlur={(e) => updateBBProspect(p.id, { primary_recruiter: e.target.value })}
                                  />
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Secondary recruiter</label>
                                  <input
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    defaultValue={p.secondary_recruiter}
                                    onBlur={(e) => updateBBProspect(p.id, { secondary_recruiter: e.target.value })}
                                  />
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Contact</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                placeholder="Phone, email, etc."
                                defaultValue={p.contact}
                                onBlur={(e) => updateBBProspect(p.id, { contact: e.target.value })}
                              />

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Priority</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.priority || ""}
                                    onChange={(e) => updateBBProspect(p.id, { priority: e.target.value || null })}
                                  >
                                    <option value="">—</option>
                                    {BB_PRIORITIES.map((pr) => (
                                      <option key={pr} value={pr}>{pr}</option>
                                    ))}
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Status</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={p.status || ""}
                                    onChange={(e) => updateBBProspect(p.id, { status: e.target.value || null })}
                                  >
                                    <option value="">—</option>
                                    {BB_STATUSES.map((st) => (
                                      <option key={st} value={st}>{st}</option>
                                    ))}
                                  </select>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Last contact date</label>
                              <input
                                type="date"
                                className="db-input"
                                style={{ width: "100%", marginBottom: "10px" }}
                                defaultValue={p.last_contact_date || ""}
                                onBlur={(e) => updateBBProspect(p.id, { last_contact_date: e.target.value || null })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Notes</label>
                              <textarea
                                className="db-input"
                                style={{ width: "100%", minHeight: "70px", resize: "vertical", fontFamily: "'Inter', sans-serif" }}
                                defaultValue={p.notes}
                                onBlur={(e) => updateBBProspect(p.id, { notes: e.target.value })}
                              />

                              <button
                                className="db-btn"
                                onClick={(e) => { e.stopPropagation(); deleteBBProspect(p.id); }}
                                style={{ marginTop: "12px", fontSize: "11px", color: "#C97A7A", borderColor: "rgba(201,122,122,0.3)", padding: "5px 9px" }}
                              >
                                Remove prospect
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}

                    {!isBasketball && isVetView && list.map((v, idx) => {
                      const isOpen = expandedId === v.id;
                      const age = calcAge(v.date_of_birth);
                      return (
                        <div key={v.id} style={{ borderBottom: `1px solid ${COLORS.hair}` }}>
                          <div
                            className="db-row"
                            onClick={() => setExpandedId(isOpen ? null : v.id)}
                            style={{ display: "flex", alignItems: "center", gap: "10px", padding: "10px 14px", cursor: "pointer" }}
                          >
                            <span style={{ fontFamily: "'Anton', sans-serif", fontSize: "20px", color: COLORS.hairStrong, width: "22px" }}>
                              {String(idx + 1).padStart(2, "0")}
                            </span>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: "13.5px", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.3px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {v.name}
                              </div>
                              <div style={{ fontSize: "11px", color: COLORS.inkDim, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {[v.current_team, v.free_agency_year ? `FA: ${v.free_agency_year}` : "FA year unset"].filter(Boolean).join(" · ")}
                              </div>
                              <div style={{ fontSize: "10.5px", color: COLORS.inkDim, opacity: 0.75, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                {[v.current_agent, v.current_agency].filter(Boolean).join(" · ") || "No agent listed"}
                              </div>
                            </div>
                            {v.is_a1 && (
                              <span
                                style={{
                                  fontFamily: "'IBM Plex Mono', monospace",
                                  fontSize: "11px",
                                  fontWeight: 800,
                                  color: "#E24C4C",
                                  border: "1.5px solid #E24C4C",
                                  borderRadius: "4px",
                                  padding: "2px 5px",
                                  flexShrink: 0,
                                  letterSpacing: "0.5px",
                                }}
                              >
                                A1
                              </span>
                            )}
                            <div
                              title="Projected value"
                              style={{
                                fontFamily: "'IBM Plex Mono', monospace",
                                fontSize: "11.5px",
                                fontWeight: 700,
                                color: v.projected_value ? COLORS.vetGreenText : COLORS.inkDim,
                                background: v.projected_value ? COLORS.vetGreen : "transparent",
                                border: `1.5px solid ${v.projected_value ? "rgba(255,255,255,0.15)" : COLORS.ungraded}`,
                                borderRadius: "8px",
                                minWidth: "48px",
                                padding: "6px 6px",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                transform: "rotate(-2deg)",
                                flexShrink: 0,
                                whiteSpace: "nowrap",
                              }}
                            >
                              {v.projected_value || "—"}
                            </div>
                          </div>

                          {isOpen && (
                            <div className="no-print" style={{ padding: "4px 14px 14px", background: COLORS.surfaceHi }}>
                              <label
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: "7px",
                                  margin: "8px 0 10px",
                                  cursor: "pointer",
                                  fontSize: "12px",
                                  color: v.is_a1 ? "#E24C4C" : COLORS.inkDim,
                                  fontWeight: v.is_a1 ? 700 : 400,
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={!!v.is_a1}
                                  onChange={(e) => {
                                    e.stopPropagation();
                                    updateVet(v.id, { is_a1: e.target.checked });
                                  }}
                                  onClick={(e) => e.stopPropagation()}
                                  style={{ accentColor: "#E24C4C", width: "14px", height: "14px" }}
                                />
                                Mark as A1 client
                              </label>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Name</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px", fontWeight: 600 }}
                                defaultValue={v.name}
                                onBlur={(e) => {
                                  const val = e.target.value.trim();
                                  if (val) updateVet(v.id, { name: val });
                                  else e.target.value = v.name;
                                }}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Position</label>
                              <select
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                value={v.position}
                                onChange={(e) => {
                                  const newPos = e.target.value;
                                  updateVet(v.id, { position: newPos, board: POSITION_BOARD[newPos] });
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

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Draft year</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={v.draft_year || ""}
                                    onChange={(e) => updateVet(v.id, { draft_year: Number(e.target.value) })}
                                  >
                                    {VET_DRAFT_YEARS.map((y) => (
                                      <option key={y} value={y}>{y}</option>
                                    ))}
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Free agency yr</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={v.free_agency_year || ""}
                                    onChange={(e) => updateVet(v.id, { free_agency_year: Number(e.target.value) })}
                                  >
                                    <option value="">—</option>
                                    {FA_YEARS.map((y) => (
                                      <option key={y} value={y}>{y}</option>
                                    ))}
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Draft round</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={v.draft_round || ""}
                                    onChange={(e) => updateVet(v.id, { draft_round: e.target.value || null })}
                                  >
                                    <option value="">—</option>
                                    {DRAFT_ROUNDS.map((r) => (
                                      <option key={r} value={r}>{r}</option>
                                    ))}
                                  </select>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Hometown</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                defaultValue={v.hometown}
                                onBlur={(e) => updateVet(v.id, { hometown: e.target.value })}
                              />

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Current team</label>
                                  <select
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    value={v.current_team || ""}
                                    onChange={(e) => updateVet(v.id, { current_team: e.target.value || null })}
                                  >
                                    <option value="">—</option>
                                    {NFL_TEAMS.map((t) => (
                                      <option key={t} value={t}>{t}</option>
                                    ))}
                                  </select>
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>College</label>
                                  <input
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    defaultValue={v.college}
                                    onBlur={(e) => updateVet(v.id, { college: e.target.value })}
                                  />
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Current agent</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                placeholder="e.g. J. Rosenhaus"
                                defaultValue={v.current_agent}
                                onBlur={(e) => updateVet(v.id, { current_agent: e.target.value })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Current agency</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                defaultValue={v.current_agency}
                                onBlur={(e) => updateVet(v.id, { current_agency: e.target.value })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Assigned agent</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                placeholder="A1 agent assigned to pursue"
                                defaultValue={v.assigned_agent}
                                onBlur={(e) => updateVet(v.id, { assigned_agent: e.target.value })}
                              />

                              <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Date of birth</label>
                                  <input
                                    type="date"
                                    className="db-input"
                                    style={{ width: "100%" }}
                                    defaultValue={v.date_of_birth || ""}
                                    onBlur={(e) => updateVet(v.id, { date_of_birth: e.target.value || null })}
                                  />
                                </div>
                                <div style={{ flex: 1 }}>
                                  <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Current age</label>
                                  <div
                                    className="db-input"
                                    style={{ width: "100%", display: "flex", alignItems: "center", color: COLORS.inkDim, boxSizing: "border-box" }}
                                  >
                                    {age || "—"}
                                  </div>
                                </div>
                              </div>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Projected $$$</label>
                              <select
                                className="db-input"
                                style={{ width: "100%", marginBottom: "10px" }}
                                value={v.projected_value || ""}
                                onChange={(e) => updateVet(v.id, { projected_value: e.target.value || null })}
                              >
                                <option value="">—</option>
                                {PROJECTED_VALUES.map((pv) => (
                                  <option key={pv} value={pv}>{pv}</option>
                                ))}
                              </select>

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Meetings</label>
                              <input
                                className="db-input"
                                style={{ width: "100%", marginBottom: "8px" }}
                                placeholder="e.g. Dinner, 3/12"
                                defaultValue={v.meetings}
                                onBlur={(e) => updateVet(v.id, { meetings: e.target.value })}
                              />

                              <label style={{ fontSize: "10.5px", color: COLORS.inkDim, display: "block", marginBottom: "3px" }}>Notes</label>
                              <textarea
                                className="db-input"
                                style={{ width: "100%", minHeight: "70px", resize: "vertical", fontFamily: "'Inter', sans-serif", marginBottom: "10px" }}
                                defaultValue={v.notes}
                                onBlur={(e) => updateVet(v.id, { notes: e.target.value })}
                              />

                              <button
                                className="db-btn"
                                onClick={(e) => { e.stopPropagation(); deleteVet(v.id); }}
                                style={{ marginTop: "2px", fontSize: "11px", color: "#C97A7A", borderColor: "rgba(201,122,122,0.3)", padding: "5px 9px" }}
                              >
                                Remove player
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  <div className="no-print" style={{ padding: "10px 14px 14px" }}>
                    {addOpenFor === pos.abbr ? (
                      isBasketball ? (
                        <div>
                          <input
                            className="db-input"
                            placeholder="Player name"
                            style={{ width: "100%", marginBottom: "6px" }}
                            value={bbAddDraft.name}
                            onChange={(e) => setBbAddDraft({ ...bbAddDraft, name: e.target.value })}
                            autoFocus
                          />
                          <input
                            className="db-input"
                            placeholder="Team"
                            style={{ width: "100%", marginBottom: "8px" }}
                            value={bbAddDraft.team}
                            onChange={(e) => setBbAddDraft({ ...bbAddDraft, team: e.target.value })}
                          />
                          <div style={{ display: "flex", gap: "6px" }}>
                            <button
                              className="db-btn"
                              onClick={() => submitAddBB(pos.abbr)}
                              style={{ flex: 1, padding: "6px", color: accent, borderColor: accent }}
                            >
                              Add to board
                            </button>
                            <button className="db-btn" onClick={() => setAddOpenFor(null)} style={{ padding: "6px 10px" }}>
                              ×
                            </button>
                          </div>
                        </div>
                      ) : isVetView ? (
                        <div>
                          <input
                            className="db-input"
                            placeholder="Player name"
                            style={{ width: "100%", marginBottom: "6px" }}
                            value={vetDraft.name}
                            onChange={(e) => setVetDraft({ ...vetDraft, name: e.target.value })}
                            autoFocus
                          />
                          <input
                            className="db-input"
                            placeholder="Hometown"
                            style={{ width: "100%", marginBottom: "6px" }}
                            value={vetDraft.hometown}
                            onChange={(e) => setVetDraft({ ...vetDraft, hometown: e.target.value })}
                          />
                          <select
                            className="db-input"
                            style={{ width: "100%", marginBottom: "8px" }}
                            value={vetDraft.draftYear}
                            onChange={(e) => setVetDraft({ ...vetDraft, draftYear: Number(e.target.value) })}
                          >
                            {VET_DRAFT_YEARS.map((y) => (
                              <option key={y} value={y}>Draft: {y}</option>
                            ))}
                          </select>
                          <div style={{ display: "flex", gap: "6px" }}>
                            <button
                              className="db-btn"
                              onClick={() => submitAddVet(pos.abbr)}
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
                      )
                    ) : (
                      <button
                        className="db-btn"
                        onClick={() => openAdd(pos.abbr)}
                        style={{ width: "100%", padding: "8px", fontSize: "12px" }}
                      >
                        {isBasketball ? "+ Add prospect" : isVetView ? "+ Add player" : "+ Add prospect"}
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




