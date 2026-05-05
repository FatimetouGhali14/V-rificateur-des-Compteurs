import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Utilitaires ────────────────────────────────────────────────────────────

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Détecte la ligne d'en-tête réelle (cherche la première ligne contenant plusieurs cellules non-vides)
function detectHeader(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
    let filled = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v != null && String(cell.v).trim() !== "") filled++;
    }
    if (filled >= 2) return r;
  }
  return 0;
}

// Parse un fichier Excel en tableau d'objets, en détectant les en-têtes automatiquement
function parseSheet(wb, sheetName) {
  const sheet = wb.Sheets[sheetName];
  const headerRow = detectHeader(sheet);
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
    range: headerRow,
  });
  return rows;
}

// Normalise une valeur pour comparaison (supprime espaces, met en string)
function norm(v) {
  return String(v ?? "").trim();
}

// Devine la colonne "numéro de compteur" dans les colonnes disponibles
function guessCompteurCol(cols) {
  const candidates = [
    "N COMPTEUR", "N° COMPTEUR", "COMPTEUR", "NUM COMPTEUR",
    "NUMERO COMPTEUR", "N_COMPTEUR", "NUM_COMPTEUR", "Compteur",
  ];
  for (const c of candidates) {
    const found = cols.find(
      (col) => norm(col).toUpperCase() === c.toUpperCase()
    );
    if (found) return found;
  }
  // Cherche par inclusion
  const found = cols.find((col) =>
    norm(col).toUpperCase().includes("COMPTEUR")
  );
  return found || cols[0];
}

// ─── Composants UI ──────────────────────────────────────────────────────────

const Chip = ({ label, color }) => (
  <span
    style={{
      display: "inline-flex",
      alignItems: "center",
      gap: 5,
      padding: "3px 10px",
      borderRadius: 20,
      fontSize: 11,
      fontWeight: 700,
      letterSpacing: 0.4,
      background: color === "green" ? "#d1fae5" : "#fee2e2",
      color: color === "green" ? "#065f46" : "#991b1b",
      border: `1px solid ${color === "green" ? "#6ee7b7" : "#fca5a5"}`,
      whiteSpace: "nowrap",
    }}
  >
    <span
      style={{
        width: 6,
        height: 6,
        borderRadius: "50%",
        background: color === "green" ? "#10b981" : "#ef4444",
        display: "inline-block",
        flexShrink: 0,
      }}
    />
    {label}
  </span>
);

const StatCard = ({ icon, value, label, accent }) => (
  <div
    style={{
      background: "#fff",
      border: "1px solid #e8edf3",
      borderRadius: 14,
      padding: "16px 18px",
      display: "flex",
      alignItems: "center",
      gap: 14,
      flex: 1,
      minWidth: 140,
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}
  >
    <div
      style={{
        width: 42,
        height: 42,
        borderRadius: 10,
        background: accent + "18",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: 18,
        flexShrink: 0,
      }}
    >
      {icon}
    </div>
    <div>
      <div
        style={{
          fontSize: 24,
          fontWeight: 800,
          color: accent,
          lineHeight: 1,
          fontVariantNumeric: "tabular-nums",
        }}
      >
        {value}
      </div>
      <div style={{ fontSize: 11, color: "#64748b", marginTop: 4, fontWeight: 500 }}>
        {label}
      </div>
    </div>
  </div>
);

const DropZone = ({ label, sublabel, onFile, file, accent }) => {
  const [drag, setDrag] = useState(false);

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDrag(false);
      const f = e.dataTransfer.files[0];
      if (f) onFile(f);
    },
    [onFile]
  );

  return (
    <label
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 8,
        border: `2px dashed ${drag ? accent : file ? accent : "#cbd5e1"}`,
        borderRadius: 14,
        padding: "24px 20px",
        cursor: "pointer",
        background: drag ? accent + "08" : file ? accent + "06" : "#f8fafc",
        transition: "all 0.2s",
        flex: 1,
        minHeight: 130,
        textAlign: "center",
      }}
    >
      <input
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])}
      />
      <div style={{ fontSize: 28 }}>{file ? "✅" : "📂"}</div>
      <div style={{ fontSize: 13, fontWeight: 700, color: file ? accent : "#475569" }}>
        {file ? file.name : label}
      </div>
      <div style={{ fontSize: 11, color: "#94a3b8" }}>
        {file ? "Cliquez pour changer" : sublabel}
      </div>
    </label>
  );
};

// Sélecteur de colonne (dropdown)
const ColSelector = ({ label, cols, value, onChange, accent }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1 }}>
    <span style={{ fontSize: 11, fontWeight: 700, color: "#64748b", letterSpacing: 0.5 }}>
      {label}
    </span>
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        padding: "7px 10px",
        border: `1px solid ${accent}60`,
        borderRadius: 8,
        fontSize: 12,
        color: "#1e293b",
        background: "#fff",
        cursor: "pointer",
        outline: "none",
        fontWeight: 600,
      }}
    >
      {cols.map((c) => (
        <option key={c} value={c}>
          {c}
        </option>
      ))}
    </select>
  </div>
);

// ─── App principale ──────────────────────────────────────────────────────────

export default function App() {
  // Fichiers uploadés
  const [fileRef, setFileRef] = useState(null);   // Etat_AbnActif (référence)
  const [fileIn, setFileIn] = useState(null);     // Mellah / fichier de compteurs à vérifier

  // Workbooks parsés
  const [wbRef, setWbRef] = useState(null);
  const [wbIn, setWbIn] = useState(null);

  // Feuilles sélectionnées
  const [sheetRef, setSheetRef] = useState("");
  const [sheetIn, setSheetIn] = useState("");

  // Colonnes sélectionnées
  const [colRef, setColRef] = useState("");   // colonne compteur dans le fichier référence
  const [colIn, setColIn] = useState("");     // colonne compteur dans le fichier d'entrée

  // État
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null); // { actifs, inactifs }
  const [tab, setTab] = useState("actif");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const PAGE_SIZE = 50;

  // ── Chargement fichier référence ──
  const loadRef = useCallback(async (f) => {
    setFileRef(f);
    setError("");
    try {
      const wb = await readExcel(f);
      setWbRef(wb);
      const first = wb.SheetNames[0];
      setSheetRef(first);
      const rows = parseSheet(wb, first);
      const cols = Object.keys(rows[0] || {});
      setColRef(guessCompteurCol(cols));
      setResult(null);
    } catch {
      setError("Erreur lecture fichier référence.");
    }
  }, []);

  // ── Chargement fichier à vérifier ──
  const loadIn = useCallback(async (f) => {
    setFileIn(f);
    setError("");
    try {
      const wb = await readExcel(f);
      setWbIn(wb);
      const first = wb.SheetNames[0];
      setSheetIn(first);
      const rows = parseSheet(wb, first);
      const cols = Object.keys(rows[0] || {});
      setColIn(guessCompteurCol(cols));
      setResult(null);
    } catch {
      setError("Erreur lecture fichier à vérifier.");
    }
  }, []);

  // ── Colonnes disponibles ──
  const colsRef = useMemo(() => {
    if (!wbRef || !sheetRef) return [];
    const rows = parseSheet(wbRef, sheetRef);
    return Object.keys(rows[0] || {});
  }, [wbRef, sheetRef]);

  const colsIn = useMemo(() => {
    if (!wbIn || !sheetIn) return [];
    const rows = parseSheet(wbIn, sheetIn);
    return Object.keys(rows[0] || {});
  }, [wbIn, sheetIn]);

  // ── Croisement ──
  const runCross = useCallback(async () => {
    if (!wbRef || !wbIn || !colRef || !colIn) return;
    setLoading(true);
    setError("");
    try {
      await new Promise((r) => setTimeout(r, 50)); // let UI update

      // Charger tous les abonnés actifs (fichier référence)
      const refRows = parseSheet(wbRef, sheetRef);
      const activeMap = new Map();
      for (const row of refRows) {
        const key = norm(row[colRef]);
        if (key) activeMap.set(key, row);
      }

      // Charger les compteurs à vérifier
      const inRows = parseSheet(wbIn, sheetIn);
      const actifs = [];
      const inactifs = [];

      for (const row of inRows) {
        const key = norm(row[colIn]);
        if (!key) continue;
        if (activeMap.has(key)) {
          actifs.push({ ...row, _refData: activeMap.get(key) });
        } else {
          inactifs.push(row);
        }
      }

      setResult({ actifs, inactifs, totalRef: activeMap.size });
      setTab("actif");
      setSearch("");
      setPage(0);
    } catch (e) {
      setError("Erreur lors du croisement : " + e.message);
    }
    setLoading(false);
  }, [wbRef, wbIn, sheetRef, sheetIn, colRef, colIn]);

  // ── Données affichées (onglet + recherche + pagination) ──
  const currentRows = useMemo(() => {
    if (!result) return [];
    const rows = tab === "actif" ? result.actifs : result.inactifs;
    if (!search.trim()) return rows;
    const q = search.toLowerCase();
    return rows.filter((r) =>
      Object.values(r).some((v) => String(v).toLowerCase().includes(q))
    );
  }, [result, tab, search]);

  const paginated = useMemo(
    () => currentRows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE),
    [currentRows, page]
  );

  const totalPages = Math.ceil(currentRows.length / PAGE_SIZE);

  // ── Colonnes tableau (toutes les colonnes du fichier d'entrée + colonnes ref pour actifs) ──
  const tableCols = useMemo(() => {
    if (!result || paginated.length === 0) return [];
    const sample = paginated[0];
    return Object.keys(sample).filter((k) => k !== "_refData");
  }, [result, paginated]);

  const refCols = useMemo(() => {
    if (!result || result.actifs.length === 0 || tab !== "actif") return [];
    const sample = result.actifs[0]?._refData || {};
    return Object.keys(sample).slice(0, 6); // limiter à 6 colonnes ref
  }, [result, tab]);

  const canRun = wbRef && wbIn && colRef && colIn;

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#f0f4f8",
        fontFamily: "'Segoe UI', system-ui, sans-serif",
      }}
    >
      {/* ── Header ── */}
      <div
        style={{
          background: "linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%)",
          padding: "28px 36px 24px",
          color: "#fff",
        }}
      >
        <div style={{ maxWidth: 1300, margin: "0 auto" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ fontSize: 36 }}>💧</div>
            <div>
              <h1
                style={{
                  margin: 0,
                  fontSize: 22,
                  fontWeight: 800,
                  letterSpacing: -0.5,
                  lineHeight: 1.2,
                }}
              >
                Vérificateur de Compteurs — SNDE
              </h1>
              <p style={{ margin: "4px 0 0", fontSize: 13, color: "#94a3b8" }}>
                Croisement dynamique entre fichier de prélèvement et état des abonnés actifs
              </p>
            </div>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1300, margin: "0 auto", padding: "24px 36px" }}>

        {/* ── Zone upload ── */}
        <div
          style={{
            background: "#fff",
            borderRadius: 16,
            border: "1px solid #e2e8f0",
            padding: "22px 24px",
            marginBottom: 20,
            boxShadow: "0 1px 4px rgba(15,23,42,0.05)",
          }}
        >
          <h2 style={{ fontSize: 14, fontWeight: 800, color: "#1e293b", margin: "0 0 16px", letterSpacing: -0.2 }}>
            1 — Charger les fichiers
          </h2>

          <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
            <DropZone
              label="Fichier référence (Etat_AbnActif)"
              sublabel="Glissez ou cliquez · .xlsx / .xls"
              onFile={loadRef}
              file={fileRef}
              accent="#0ea5e9"
            />
            <DropZone
              label="Fichier à vérifier (fichier de prélèvement)"
              sublabel="Glissez ou cliquez · .xlsx / .xls"
              onFile={loadIn}
              file={fileIn}
              accent="#8b5cf6"
            />
          </div>
        </div>

        {/* ── Configuration colonnes ── */}
        {(wbRef || wbIn) && (
          <div
            style={{
              background: "#fff",
              borderRadius: 16,
              border: "1px solid #e2e8f0",
              padding: "22px 24px",
              marginBottom: 20,
              boxShadow: "0 1px 4px rgba(15,23,42,0.05)",
            }}
          >
            <h2 style={{ fontSize: 14, fontWeight: 800, color: "#1e293b", margin: "0 0 16px" }}>
              2 — Configurer les colonnes
            </h2>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              {wbRef && (
                <div
                  style={{
                    flex: 1,
                    minWidth: 280,
                    padding: "14px 16px",
                    background: "#f0f9ff",
                    borderRadius: 12,
                    border: "1px solid #bae6fd",
                  }}
                >
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#0369a1", marginBottom: 12 }}>
                    📘 Fichier référence
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <ColSelector
                      label="FEUILLE"
                      cols={wbRef.SheetNames}
                      value={sheetRef}
                      onChange={(v) => { setSheetRef(v); setColRef(guessCompteurCol(parseSheet(wbRef, v).length ? Object.keys(parseSheet(wbRef, v)[0]) : [])); }}
                      accent="#0ea5e9"
                    />
                    <ColSelector
                      label="COLONNE COMPTEUR"
                      cols={colsRef}
                      value={colRef}
                      onChange={setColRef}
                      accent="#0ea5e9"
                    />
                  </div>
                </div>
              )}
              {wbIn && (
                <div
                  style={{
                    flex: 1,
                    minWidth: 280,
                    padding: "14px 16px",
                    background: "#faf5ff",
                    borderRadius: 12,
                    border: "1px solid #d8b4fe",
                  }}
                >
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#7c3aed", marginBottom: 12 }}>
                    📗 Fichier à vérifier
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <ColSelector
                      label="FEUILLE"
                      cols={wbIn.SheetNames}
                      value={sheetIn}
                      onChange={(v) => { setSheetIn(v); setColIn(guessCompteurCol(parseSheet(wbIn, v).length ? Object.keys(parseSheet(wbIn, v)[0]) : [])); }}
                      accent="#8b5cf6"
                    />
                    <ColSelector
                      label="COLONNE COMPTEUR"
                      cols={colsIn}
                      value={colIn}
                      onChange={setColIn}
                      accent="#8b5cf6"
                    />
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ── Bouton lancer ── */}
        {canRun && (
          <div style={{ textAlign: "center", marginBottom: 20 }}>
            <button
              onClick={runCross}
              disabled={loading}
              style={{
                padding: "13px 40px",
                background: loading
                  ? "#94a3b8"
                  : "linear-gradient(135deg, #0ea5e9, #8b5cf6)",
                color: "#fff",
                border: "none",
                borderRadius: 12,
                fontSize: 15,
                fontWeight: 800,
                cursor: loading ? "not-allowed" : "pointer",
                letterSpacing: 0.3,
                boxShadow: loading ? "none" : "0 4px 14px rgba(14,165,233,0.35)",
                transition: "all 0.2s",
              }}
            >
              {loading ? "⏳ Analyse en cours..." : "🔍 Lancer le croisement"}
            </button>
          </div>
        )}

        {error && (
          <div
            style={{
              background: "#fef2f2",
              border: "1px solid #fca5a5",
              borderRadius: 10,
              padding: "12px 16px",
              color: "#991b1b",
              fontSize: 13,
              marginBottom: 16,
            }}
          >
            ⚠ {error}
          </div>
        )}

        {/* ── Résultats ── */}
        {result && (
          <>
            {/* Stats */}
            <div
              style={{
                display: "flex",
                gap: 14,
                flexWrap: "wrap",
                marginBottom: 18,
              }}
            >
              <StatCard
                icon="📋"
                value={result.actifs.length + result.inactifs.length}
                label="Total compteurs vérifiés"
                accent="#6366f1"
              />
              <StatCard
                icon="✅"
                value={result.actifs.length}
                label="Actifs (trouvés)"
                accent="#10b981"
              />
              <StatCard
                icon="❌"
                value={result.inactifs.length}
                label="Inactifs / non trouvés"
                accent="#ef4444"
              />
              <StatCard
                icon="📊"
                value={
                  Math.round(
                    (result.actifs.length /
                      (result.actifs.length + result.inactifs.length || 1)) *
                      100
                  ) + "%"
                }
                label="Taux d'activation"
                accent="#f59e0b"
              />
              <StatCard
                icon="🗄"
                value={result.totalRef.toLocaleString()}
                label="Abonnés dans référence"
                accent="#0ea5e9"
              />
            </div>

            {/* Tableau */}
            <div
              style={{
                background: "#fff",
                borderRadius: 16,
                border: "1px solid #e2e8f0",
                overflow: "hidden",
                boxShadow: "0 1px 4px rgba(15,23,42,0.05)",
              }}
            >
              {/* Tabs + Search */}
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  padding: "14px 18px",
                  borderBottom: "1px solid #f1f5f9",
                  flexWrap: "wrap",
                  gap: 10,
                }}
              >
                <div style={{ display: "flex", gap: 8 }}>
                  {[
                    { key: "actif", label: "✅ Actifs", count: result.actifs.length, bg: "#10b981" },
                    { key: "inactif", label: "❌ Inactifs", count: result.inactifs.length, bg: "#ef4444" },
                  ].map((t) => (
                    <button
                      key={t.key}
                      onClick={() => { setTab(t.key); setSearch(""); setPage(0); }}
                      style={{
                        padding: "7px 16px",
                        borderRadius: 9,
                        border: "none",
                        cursor: "pointer",
                        fontWeight: 700,
                        fontSize: 13,
                        background: tab === t.key ? t.bg : "#f1f5f9",
                        color: tab === t.key ? "#fff" : "#475569",
                        boxShadow:
                          tab === t.key ? `0 2px 8px ${t.bg}50` : "none",
                        transition: "all 0.15s",
                      }}
                    >
                      {t.label}{" "}
                      <span
                        style={{
                          background:
                            tab === t.key ? "rgba(255,255,255,0.25)" : "#e2e8f0",
                          padding: "1px 7px",
                          borderRadius: 10,
                          marginLeft: 5,
                          fontSize: 11,
                        }}
                      >
                        {t.count}
                      </span>
                    </button>
                  ))}
                </div>
                <input
                  value={search}
                  onChange={(e) => { setSearch(e.target.value); setPage(0); }}
                  placeholder="🔍 Rechercher dans tous les champs..."
                  style={{
                    padding: "7px 13px",
                    border: "1px solid #e2e8f0",
                    borderRadius: 9,
                    fontSize: 12,
                    width: 280,
                    outline: "none",
                    background: "#f8fafc",
                    color: "#1e293b",
                  }}
                />
              </div>

              {/* Info ligne */}
              <div
                style={{
                  padding: "8px 18px",
                  fontSize: 11,
                  color: "#64748b",
                  borderBottom: "1px solid #f1f5f9",
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <span>
                  {currentRows.length} résultat
                  {currentRows.length > 1 ? "s" : ""}
                  {search && ` pour "${search}"`} · Page {page + 1}/{Math.max(totalPages, 1)}
                </span>
                {totalPages > 1 && (
                  <div style={{ display: "flex", gap: 6 }}>
                    <button
                      onClick={() => setPage((p) => Math.max(0, p - 1))}
                      disabled={page === 0}
                      style={{
                        padding: "3px 10px",
                        border: "1px solid #e2e8f0",
                        borderRadius: 6,
                        background: "#fff",
                        cursor: page === 0 ? "default" : "pointer",
                        fontSize: 12,
                        color: page === 0 ? "#cbd5e1" : "#475569",
                      }}
                    >
                      ‹ Préc
                    </button>
                    <button
                      onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
                      disabled={page >= totalPages - 1}
                      style={{
                        padding: "3px 10px",
                        border: "1px solid #e2e8f0",
                        borderRadius: 6,
                        background: "#fff",
                        cursor: page >= totalPages - 1 ? "default" : "pointer",
                        fontSize: 12,
                        color: page >= totalPages - 1 ? "#cbd5e1" : "#475569",
                      }}
                    >
                      Suiv ›
                    </button>
                  </div>
                )}
              </div>

              {/* Table */}
              <div style={{ overflowX: "auto", maxHeight: 460, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr
                      style={{
                        background: tab === "actif" ? "#f0fdf4" : "#fef2f2",
                        position: "sticky",
                        top: 0,
                        zIndex: 2,
                      }}
                    >
                      <th
                        style={{
                          padding: "9px 14px",
                          textAlign: "left",
                          fontSize: 10,
                          fontWeight: 800,
                          letterSpacing: 0.5,
                          textTransform: "uppercase",
                          color: tab === "actif" ? "#166534" : "#991b1b",
                          borderBottom: `2px solid ${tab === "actif" ? "#bbf7d0" : "#fecaca"}`,
                          whiteSpace: "nowrap",
                        }}
                      >
                        STATUT
                      </th>
                      {tableCols.map((c) => (
                        <th
                          key={c}
                          style={{
                            padding: "9px 14px",
                            textAlign: "left",
                            fontSize: 10,
                            fontWeight: 800,
                            letterSpacing: 0.5,
                            textTransform: "uppercase",
                            color: tab === "actif" ? "#166534" : "#991b1b",
                            borderBottom: `2px solid ${tab === "actif" ? "#bbf7d0" : "#fecaca"}`,
                            whiteSpace: "nowrap",
                          }}
                        >
                          {c}
                        </th>
                      ))}
                      {tab === "actif" &&
                        refCols.map((c) => (
                          <th
                            key={"ref_" + c}
                            style={{
                              padding: "9px 14px",
                              textAlign: "left",
                              fontSize: 10,
                              fontWeight: 800,
                              letterSpacing: 0.5,
                              textTransform: "uppercase",
                              color: "#0369a1",
                              borderBottom: "2px solid #bae6fd",
                              background: "#f0f9ff",
                              whiteSpace: "nowrap",
                            }}
                          >
                            REF: {c}
                          </th>
                        ))}
                    </tr>
                  </thead>
                  <tbody>
                    {paginated.length === 0 ? (
                      <tr>
                        <td
                          colSpan={tableCols.length + 1 + refCols.length}
                          style={{
                            padding: 32,
                            textAlign: "center",
                            color: "#94a3b8",
                            fontSize: 13,
                          }}
                        >
                          Aucun résultat
                        </td>
                      </tr>
                    ) : (
                      paginated.map((row, i) => (
                        <tr
                          key={i}
                          style={{
                            background: i % 2 === 0 ? "#fff" : "#f8fafc",
                          }}
                        >
                          <td style={{ padding: "8px 14px" }}>
                            <Chip
                              label={tab === "actif" ? "ACTIF" : "INACTIF"}
                              color={tab === "actif" ? "green" : "red"}
                            />
                          </td>
                          {tableCols.map((c) => (
                            <td
                              key={c}
                              style={{
                                padding: "8px 14px",
                                color:
                                  c.toUpperCase().includes("COMPTEUR")
                                    ? "#1e293b"
                                    : "#475569",
                                fontWeight: c.toUpperCase().includes("COMPTEUR")
                                  ? 700
                                  : 400,
                                fontFamily: c.toUpperCase().includes("COMPTEUR")
                                  ? "monospace"
                                  : "inherit",
                                whiteSpace: "nowrap",
                                maxWidth: 200,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                              }}
                            >
                              {norm(row[c]) || "—"}
                            </td>
                          ))}
                          {tab === "actif" &&
                            refCols.map((c) => (
                              <td
                                key={"ref_" + c}
                                style={{
                                  padding: "8px 14px",
                                  color: "#0369a1",
                                  background:
                                    i % 2 === 0 ? "#f0f9ff" : "#e0f2fe",
                                  fontSize: 11,
                                  whiteSpace: "nowrap",
                                  maxWidth: 180,
                                  overflow: "hidden",
                                  textOverflow: "ellipsis",
                                }}
                              >
                                {norm(row._refData?.[c]) || "—"}
                              </td>
                            ))}
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
