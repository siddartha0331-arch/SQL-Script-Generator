import React, { useState, useCallback, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Database,
  Download,
  Check,
  ChevronDown,
  FileSpreadsheet,
  AlertCircle,
  Copy,
  CheckCircle2,
  Table2,
  Trash2,
  Pencil,
  Eye,
} from "lucide-react";

import { theme } from "./theme";
import {
  buildTableScript,
  downloadTextFile,
  // NEW — Custom SQL Template feature
  validateTemplate,
  generateCustomSQL,
  getUniqueColumnValues,
  filterRowsByColumnValues,
  buildTemplateFromGuided,
} from "./utils/sqlHelpers";
import { MultiSelectDropdown } from "./components/MultiSelectDropdown";
import { DataPreviewModal } from "./components/DataPreviewModal";
import sagitecLogo from "./sagitec-logo.png";

/* ============================================================
   NEW — ValueSearchPicker
   Small searchable checkbox list used by the row filter to find
   specific values (e.g. type "1002" to find MemberId 1002 instead
   of scrolling through hundreds of rows). Kept local to this file
   since nothing else uses it.
   ============================================================ */
function ValueSearchPicker({ options, selected, onToggle }) {
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    if (!search.trim()) return options;
    const q = search.toLowerCase();
    return options.filter((v) => String(v).toLowerCase().includes(q));
  }, [options, search]);

  return (
    <div>
      <input
        type="text"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        placeholder="Search values..."
        style={{
          width: "100%",
          padding: "7px 10px",
          borderRadius: 6,
          border: `1px solid ${theme.cardBorder}`,
          background: "#fff",
          fontSize: 12,
          boxSizing: "border-box",
          marginBottom: 6,
        }}
      />
      <div
        style={{
          maxHeight: 140,
          overflowY: "auto",
          border: `1px solid ${theme.cardBorder}`,
          borderRadius: 6,
          padding: 4,
          background: "#fff",
        }}
      >
        {filtered.length === 0 ? (
          <p style={{ fontSize: 11, color: theme.textMuted, margin: 4 }}>No matches.</p>
        ) : (
          filtered.map((v) => (
            <label
              key={String(v)}
              style={{ display: "flex", alignItems: "center", gap: 6, padding: "3px 4px", fontSize: 11.5, cursor: "pointer" }}
            >
              <input
                type="checkbox"
                checked={selected.has(v)}
                onChange={() => onToggle(v)}
                style={{ accentColor: theme.brandBlue, flexShrink: 0 }}
              />
              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{String(v)}</span>
            </label>
          ))
        )}
      </div>
    </div>
  );
}

/* ============================================================
   NEW — ColumnCheckboxGrid
   Reusable checkbox grid for the Guided Builder's three pickers
   (INSERT columns, SET columns, WHERE columns). One implementation,
   three uses, so they always look and behave identically.
   ============================================================ */
function ColumnCheckboxGrid({ headers, selected, onToggle }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 6, maxHeight: 150, overflowY: "auto", paddingRight: 2 }}>
      {headers.map((h) => {
        const checked = selected.has(h);
        return (
          <label
            key={h}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              padding: "6px 8px",
              borderRadius: 6,
              border: `1px solid ${checked ? theme.brandBlue : theme.cardBorder}`,
              background: checked ? theme.brandBlueLight : "#fff",
              fontSize: 11.5,
              cursor: "pointer",
            }}
          >
            <input type="checkbox" checked={checked} onChange={() => onToggle(h)} style={{ accentColor: theme.brandBlue, flexShrink: 0 }} />
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{h}</span>
          </label>
        );
      })}
    </div>
  );
}

export default function SqlScriptGenerator() {
  const [sheets, setSheets] = useState([]);
  const [fileNames, setFileNames] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [configs, setConfigs] = useState({});
  const [error, setError] = useState("");
  const [copiedId, setCopiedId] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  const [editingHeader, setEditingHeader] = useState(null);
  const [editValue, setEditValue] = useState("");

  const [previewSheetId, setPreviewSheetId] = useState(null);

  // NEW — collapsible help guide shown inside Raw SQL mode
  const [showCustomSqlHelp, setShowCustomSqlHelp] = useState(false);

  const fileBatchRef = useRef(0);

  const activeSheet = sheets.find((s) => s.id === activeId);
  const activeConfig = activeSheet ? configs[activeSheet.id] : null;
  const previewSheet = sheets.find((s) => s.id === previewSheetId) || null;

  useEffect(() => {
    setEditingHeader(null);
  }, [activeId]);

  useEffect(() => {
    if (error) {
      const timer = setTimeout(() => {
        setError("");
      }, 10000); // 10 seconds

      // This cleanup function runs if 'error' changes before 10s is up
      return () => clearTimeout(timer);
    }
  }, [error]);

  useEffect(() => {
    const prevBodyBg = document.body.style.background;
    const prevHtmlBg = document.documentElement.style.background;
    document.body.style.background = theme.pageBg;
    document.body.style.margin = "0";
    document.documentElement.style.background = theme.pageBg;
    document.documentElement.style.height = "100%";
    document.body.style.minHeight = "100vh";
    return () => {
      document.body.style.background = prevBodyBg;
      document.documentElement.style.background = prevHtmlBg;
    };
  }, []);

  const parseOneFile = (arrayBuffer, name) => {
    const batch = fileBatchRef.current++;
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const parsedSheets = workbook.SheetNames.map((sheetName, i) => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
      const headers =
        rows.length > 0
          ? Object.keys(rows[0])
          : XLSX.utils.sheet_to_json(ws, { header: 1 })[0] || [];
      return {
        id: `f${batch}-s${i}-${sheetName}`,
        originalName: sheetName,
        sourceFile: name,
        headers,
        rows,
      };
    }).filter((s) => s.headers.length > 0);

    const newConfigs = {};
    parsedSheets.forEach((s) => {
      newConfigs[s.id] = {
        tableName: s.originalName,
        keyColumns: new Set(),
        insertColumns: new Set(s.headers),
        columnNameOverrides: {},
        identityInsert: false,
        generated: null,
        // NEW — Custom SQL Template feature, one independent config per sheet
        customSqlEnabled: false,
        customOperation: "INSERT", // "INSERT" | "UPDATE" | "DELETE"
        customTemplate: "",
        guidedInsertColumns: new Set(s.headers),
        guidedSetColumns: new Set(),
        guidedWhereColumns: new Set(),
        // NEW — optional row filter: generate for only selected rows, not all
        rowFilterEnabled: false,
        rowFilterColumn: "",
        rowFilterValues: new Set(),
      };
    });

    return { parsedSheets, newConfigs };
  };

  const readFileAsArrayBuffer = (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
      reader.readAsArrayBuffer(file);
    });

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;

    const valid = files.filter((f) => /\.(xlsx|xls)$/i.test(f.name));
    const invalidCount = files.length - valid.length;

    if (valid.length === 0) {
      setError("Please upload .xlsx or .xls files.");
      return;
    }

    const allNewSheets = [];
    let mergedConfigs = {};
    const failedFiles = [];

    for (const file of valid) {
      try {
        const buffer = await readFileAsArrayBuffer(file);
        const { parsedSheets, newConfigs } = parseOneFile(buffer, file.name);
        if (parsedSheets.length === 0) {
          failedFiles.push(`${file.name} (no usable sheets — check row 1 has headers)`);
          continue;
        }
        allNewSheets.push(...parsedSheets);
        mergedConfigs = { ...mergedConfigs, ...newConfigs };
      } catch (e) {
        failedFiles.push(file.name);
      }
    }

    if (allNewSheets.length > 0) {
      setSheets((prev) => [...prev, ...allNewSheets]);
      setConfigs((prev) => ({ ...prev, ...mergedConfigs }));
      setFileNames((prev) => [...prev, ...valid.map((f) => f.name)]);
      setActiveId((prev) => prev ?? allNewSheets[0].id);
    }

    if (invalidCount > 0 || failedFiles.length > 0) {
      const parts = [];
      if (invalidCount > 0) parts.push(`${invalidCount} file(s) skipped — not .xlsx/.xls`);
      if (failedFiles.length > 0) parts.push(`Couldn't read: ${failedFiles.join(", ")}`);
      setError(parts.join(". "));
    } else {
      setError("");
    }
  }, []);

  const updateConfig = (id, patch) => {
    setConfigs((prev) => ({ ...prev, [id]: { ...prev[id], ...patch, generated: null } }));
  };

  const toggleKeyColumn = (id, col) => {
    setConfigs((prev) => {
      const set = new Set(prev[id].keyColumns);
      set.has(col) ? set.delete(col) : set.add(col);
      return { ...prev, [id]: { ...prev[id], keyColumns: set, generated: null } };
    });
  };

  const toggleInsertColumn = (id, col) => {
    setConfigs((prev) => {
      const set = new Set(prev[id].insertColumns);
      set.has(col) ? set.delete(col) : set.add(col);
      return { ...prev, [id]: { ...prev[id], insertColumns: set, generated: null } };
    });
  };

  // Plain Set-toggle — used only by the row filter (rowFilterValues),
  // which has nothing to do with the SQL template text.
  const toggleGuidedField = (id, field, val) => {
    setConfigs((prev) => {
      const set = new Set(prev[id][field]);
      set.has(val) ? set.delete(val) : set.add(val);
      return { ...prev, [id]: { ...prev[id], [field]: set, generated: null } };
    });
  };

  // NEW — toggles a template-builder checkbox (INSERT/SET/WHERE column) AND
  // immediately rebuilds customTemplate from the updated selections. This is
  // what makes the textarea "start correct" every time you check a box —
  // but once you type into the textarea by hand, nothing overwrites it again
  // until you toggle another checkbox (an explicit, expected regenerate).
  const toggleGuidedColumn = (id, field, col) => {
    setConfigs((prev) => {
      const cfg = prev[id];
      const set = new Set(cfg[field]);
      set.has(col) ? set.delete(col) : set.add(col);

      let newTemplate = cfg.customTemplate;
      if (cfg.customOperation === "INSERT" && field === "guidedInsertColumns") {
        newTemplate = buildTemplateFromGuided("INSERT", cfg.tableName, set, null);
      } else if (cfg.customOperation === "UPDATE") {
        const setCols = field === "guidedSetColumns" ? set : cfg.guidedSetColumns;
        const whereCols = field === "guidedWhereColumns" ? set : cfg.guidedWhereColumns;
        newTemplate = buildTemplateFromGuided("UPDATE", cfg.tableName, setCols, whereCols);
      } else if (cfg.customOperation === "DELETE" && field === "guidedWhereColumns") {
        newTemplate = buildTemplateFromGuided("DELETE", cfg.tableName, null, set);
      }

      return {
        ...prev,
        [id]: { ...cfg, [field]: set, customTemplate: newTemplate, generated: null },
      };
    });
  };

  const startEditingColumn = (header) => {
    setEditingHeader(header);
    setEditValue(activeConfig?.columnNameOverrides?.[header] || header);
  };

  const commitColumnRename = (header) => {
    const trimmed = editValue.trim();
    updateConfig(activeSheet.id, {
      columnNameOverrides: {
        ...(activeConfig.columnNameOverrides || {}),
        [header]: trimmed || header,
      },
    });
    setEditingHeader(null);
  };

  const cancelColumnRename = () => setEditingHeader(null);

  const deleteSheet = (id) => {
    setSheets((prev) => {
      const next = prev.filter((s) => s.id !== id);
      setActiveId((prevActive) =>
        prevActive === id ? (next.length > 0 ? next[0].id : null) : prevActive
      );
      return next;
    });
    setConfigs((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setPreviewSheetId((prev) => (prev === id ? null : prev));
  };

  const generateForActiveSheet = () => {
    if (!activeSheet) return;
    const cfg = configs[activeSheet.id];

    // ============================================================
    // NEW — Custom SQL Template path. getEffectiveTemplate() resolves
    // to either the Guided-Builder output or the Raw textarea content,
    // and validateTemplate() ALWAYS runs first. If it fails, `generated`
    // is never set, so nothing dangerous reaches Export.
    // ============================================================
    if (cfg.customSqlEnabled) {
      const validation = validateTemplate(cfg.customTemplate, cfg.customOperation);
      if (!validation.valid) {
        setError(validation.message);
        return;
      }

      const targetRows = cfg.rowFilterEnabled
        ? filterRowsByColumnValues(activeSheet.rows, cfg.rowFilterColumn, cfg.rowFilterValues)
        : activeSheet.rows;

      if (cfg.rowFilterEnabled && targetRows.length === 0) {
        setError("Row filter is on but matches 0 rows — nothing to generate. Pick at least one value.");
        return;
      }

      setError("");
      const script = generateCustomSQL(cfg.customTemplate, targetRows);
      setConfigs((prev) => ({
        ...prev,
        [activeSheet.id]: { ...prev[activeSheet.id], generated: script },
      }));
      return;
    }

    // ============================================================
    // EXISTING — default generator path (unchanged from your original)
    // ============================================================
    // 1. Check for empty values in selected key columns
    for (const col of cfg.keyColumns) {
      const hasEmpty = activeSheet.rows.some(
        (row) => row[col] === null || row[col] === undefined || String(row[col]).trim() === ""
      );

      if (hasEmpty) {
        setError(`Error: The column "${col}" selected in your WHERE clause contains empty values. Please clean your data.`);
        return;
      } else {
        // Optional: Log or notify that the column is clean
        console.log(`Success: Column "${col}" has no empty values.`);
      }
    }

    // 2. Validate other inputs
    if (!cfg.tableName.trim()) {
      setError("Table name can't be empty.");
      return;
    }
    if (cfg.keyColumns.size === 0) {
      setError(`Pick at least one key column for "${cfg.tableName}".`);
      return;
    }
    if (cfg.insertColumns.size === 0) {
      setError(`Select at least one column to include.`);
      return;
    }

    setError("");

    // Proceed with generation...
    const script = buildTableScript(
      activeSheet,
      cfg.tableName.trim(),
      Array.from(cfg.keyColumns),
      cfg.insertColumns,
      cfg.columnNameOverrides || {},
      cfg.identityInsert
    );

    setConfigs((prev) => ({
      ...prev,
      [activeSheet.id]: { ...prev[activeSheet.id], generated: script },
    }));
  };

  const generatedSheets = useMemo(
    () => sheets.filter((s) => configs[s.id]?.generated),
    [sheets, configs]
  );

  const combinedScript = useMemo(() => {
    return generatedSheets
      .map(
        (s) =>
          `-- =====================================================\n-- TABLE: ${configs[s.id].tableName}\n-- =====================================================\n\n${configs[s.id].generated}`
      )
      .join("\n\n");
  }, [generatedSheets, configs]);

  // NEW — live preview for the Custom SQL panel: row/statement count,
  // which {{tags}} are used, unknown-column warnings, and current
  // validation state. Computed from the SAME getEffectiveTemplate()
  // that generation uses, so preview and reality can never disagree.
  const customSqlPreview = useMemo(() => {
    if (!activeSheet || !activeConfig?.customSqlEnabled) return null;

    const template = activeConfig.customTemplate || "";
    const operation = activeConfig.customOperation;

    const targetRows = activeConfig.rowFilterEnabled
      ? filterRowsByColumnValues(activeSheet.rows, activeConfig.rowFilterColumn, activeConfig.rowFilterValues)
      : activeSheet.rows;
    const rowCount = targetRows.length;

    const tagsUsed = [...template.matchAll(/\{\{\s*([^}]+?)\s*\}\}/g)].map((m) => m[1]);
    const uniqueTags = [...new Set(tagsUsed)];
    const unknownTags = uniqueTags.filter((t) => !activeSheet.headers.includes(t));

    const validation = validateTemplate(template, operation);

    return { rowCount, uniqueTags, unknownTags, validation, operation };
  }, [activeSheet, activeConfig]);

  const copyToClipboard = async (text, id) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(id);
      setTimeout(() => setCopiedId(""), 1500);
    } catch (e) {
      /* clipboard can fail in sandboxed contexts; export still works */
    }
  };

  const [collapsedFiles, setCollapsedFiles] = useState(new Set());

  const toggleCollapse = (fileName) => {
    setCollapsedFiles((prev) => {
      const next = new Set(prev);
      next.has(fileName) ? next.delete(fileName) : next.add(fileName);
      return next;
    });
  };

  const reset = () => {
    setSheets([]);
    setConfigs({});
    setFileNames([]);
    setError("");
    setActiveId(null);
    setPreviewSheetId(null);
  };

  const orangePanelStyle = {
    background: theme.panelBg,
    border: `1px solid ${theme.panelBorder}`,
    borderRadius: 12,
    padding: 16,
    minWidth: 0,
  };

  // This will be used for the Middle (Configure + Preview) panels
  const bluePanelStyle = {
    background: theme.cardBg,
    border: `1px solid ${theme.cardBorder}`,
    borderRadius: 12,
    boxShadow: theme.cardShadow,
    padding: 16,
    minWidth: 0,
  };

  const outlineButtonStyle = {
    fontSize: 12,
    display: "flex",
    alignItems: "center",
    gap: 5,
    padding: "6px 11px",
    borderRadius: 7,
    border: `1px solid ${theme.brandBlueBorder}`,
    background: theme.brandBlueLight,
    color: theme.brandBlueDark,
    cursor: "pointer",
    fontWeight: 600,
  };

  const iconRowButtonStyle = {
    background: "none",
    border: "none",
    cursor: "pointer",
    padding: 3,
    flexShrink: 0,
    display: "flex",
  };

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif",
        background: theme.pageBg,
        minHeight: "100vh",
        width: "100%",
        boxSizing: "border-box",
        color: theme.textPrimary,
      }}
    >
      {/* HEADER — full-bleed blue band */}
      <div
        style={{
          width: "100%",
          boxSizing: "border-box",
          background: `linear-gradient(135deg, ${theme.headerBgFrom}, ${theme.headerBgTo})`,
          padding: "0 28px",     // removed vertical padding — height now controls that instead
          height: 90,             // ← NEW: fixed height on the outer band
          marginBottom: 20,
          display: "flex",        // ← NEW: needed so alignItems below can vertically center content
          alignItems: "center",   // ← NEW: centers everything inside this fixed-height band
        }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div>
              <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-0.01em", color: theme.headerText }}>
                SQL Script Generator
              </h1>
              <p style={{ fontSize: 13, color: theme.headerSubtext, margin: 0 }}>
                Convert spreadsheet data to database-ready SQL in seconds.
              </p>
            </div>
          </div>

          {/* Logo wrapper — this is the key fix */}
          <div style={{ height: 65, display: "flex", alignItems: "center", flexShrink: 0 }}>
            <img
              src={sagitecLogo}
              alt="Sagitec"
              style={{ height: "100%", width: "auto" }}
            />
          </div>
        </div>
      </div>

      <div style={{ padding: "0 28px 28px", boxSizing: "border-box" }}>
        {sheets.length === 0 && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragging(false);
              handleFiles(e.dataTransfer.files);
            }}
            style={{
              maxWidth: 620,
              margin: "40px auto 0",
              border: `2px dashed ${isDragging ? theme.brandBlue : theme.cardBorder}`,
              borderRadius: 14,
              padding: "60px 24px",
              textAlign: "center",
              background: isDragging ? theme.brandBlueLight : theme.cardBg,
              boxShadow: theme.cardShadow,
              transition: "border-color 120ms ease, background 120ms ease",
            }}
          >
            <Upload size={30} color={theme.brandBlue} style={{ marginBottom: 12 }} />
            <p style={{ margin: "0 0 4px", fontSize: 15, fontWeight: 600, color: theme.textPrimary }}>
              Drop one or more .xlsx files here, or click to browse
            </p>
            <p style={{ margin: "0 0 18px", fontSize: 13, color: theme.textSecondary }}>
              Every sheet, across every file, becomes a table. First row must be column headers.
            </p>
            <label
              style={{
                display: "inline-block",
                padding: "10px 22px",
                borderRadius: 8,
                background: theme.brandBlue,
                color: "#fff",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              Choose file(s)
              <input
                type="file"
                accept=".xlsx,.xls"
                multiple
                onChange={(e) => handleFiles(e.target.files)}
                style={{ display: "none" }}
              />
            </label>
          </div>
        )}

        {error && (
          <div
            style={{
              marginTop: 16,
              padding: "10px 14px",
              borderRadius: 8,
              background: theme.dangerBg,
              border: `1px solid ${theme.dangerBorder}`,
              display: "flex",
              gap: 8,
              alignItems: "flex-start",
              fontSize: 13,
              color: theme.danger,
              // NEW — SECURITY WARNING messages render bold for visibility
              fontWeight: error.startsWith("SECURITY WARNING") ? 700 : 400,
            }}
          >
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        {sheets.length > 0 && (
          <div style={{ marginTop: 4 }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 14,
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: theme.textSecondary, overflow: "hidden" }}>
                <FileSpreadsheet size={15} style={{ flexShrink: 0 }} />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {fileNames.join(", ")}
                </span>
                <span style={{ flexShrink: 0 }}>·</span>
                <span style={{ flexShrink: 0 }}>{sheets.length} sheet{sheets.length > 1 ? "s" : ""}</span>
              </div>
              <div style={{ display: "flex", gap: 14, flexShrink: 0 }}>
                <label style={{ fontSize: 12, color: theme.brandBlue, cursor: "pointer", fontWeight: 600 }}>
                  Add more files
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    multiple
                    onChange={(e) => handleFiles(e.target.files)}
                    style={{ display: "none" }}
                  />
                </label>
                <button
                  onClick={reset}
                  style={{
                    fontSize: 12,
                    color: theme.textSecondary,
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    textDecoration: "underline",
                  }}
                >
                  Start over
                </button>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "260px minmax(420px, 1fr) 320px",
                gap: 18,
                alignItems: "start",
              }}
            >
              {/* LEFT: sheet list */}
              <div style={orangePanelStyle}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.brandBlue, margin: "0 0 10px", fontWeight: 700 }}>
                  Sheets
                </p>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {fileNames.map((fn) => {
                    const sheetsForFile = sheets.filter((s) => s.sourceFile === fn);
                    const isCollapsed = collapsedFiles.has(fn);

                    return (
                      <div key={fn} style={{ marginBottom: "8px" }}>
                        <div
                          onClick={() => toggleCollapse(fn)}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: "6px",
                            cursor: "pointer",
                            padding: "4px 0",
                            color: theme.textPrimary,
                            fontWeight: 700,
                            fontSize: "9px",
                            textTransform: "uppercase",
                          }}
                        >
                          <ChevronDown
                            size={12}
                            style={{
                              transform: isCollapsed ? "rotate(-90deg)" : "rotate(0deg)",
                              transition: "transform 0.2s",
                            }}
                          />
                          {fn}
                        </div>

                        {!isCollapsed && (
                          <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 4 }}>
                            {sheetsForFile.map((s) => {
                              const cfg = configs[s.id];
                              const isActive = s.id === activeId;
                              return (
                                <div
                                  key={s.id}
                                  onClick={() => setActiveId(s.id)}
                                  style={{
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "space-between",
                                    padding: "8px 10px",
                                    borderRadius: 8,
                                    borderLeft: isActive ? `3px solid ${theme.brandBlue}` : "3px solid transparent",
                                    border: isActive ? `1px solid ${theme.brandBlueBorder}` : "1px solid transparent",
                                    background: isActive ? theme.brandBlueLight : "transparent",
                                    cursor: "pointer",
                                  }}
                                >
                                  <div style={{ display: "flex", alignItems: "center", gap: 6, overflow: "hidden" }}>
                                    <Table2 size={13} color={isActive ? theme.brandBlue : theme.textMuted} style={{ flexShrink: 0 }} />
                                    <span
                                      style={{
                                        fontSize: 12.5,
                                        fontWeight: isActive ? 700 : 500,
                                        color: isActive ? theme.brandBlueDark : theme.textPrimary,
                                        overflow: "hidden",
                                        textOverflow: "ellipsis",
                                        whiteSpace: "nowrap",
                                      }}
                                    >
                                      {cfg?.tableName || s.originalName}
                                    </span>
                                    {/* NEW — badge shown when a sheet uses the Custom SQL feature */}
                                    {cfg?.customSqlEnabled && (
                                      <span
                                        title="Custom SQL Template"
                                        style={{
                                          fontSize: 8.5,
                                          fontWeight: 700,
                                          color: theme.brandBlueDark,
                                          background: theme.brandBlueLight,
                                          border: `1px solid ${theme.brandBlueBorder}`,
                                          borderRadius: 4,
                                          padding: "1px 4px",
                                          flexShrink: 0,
                                        }}
                                      >
                                        CUSTOM
                                      </span>
                                    )}
                                    {cfg?.generated && <CheckCircle2 size={12} color={theme.success} style={{ flexShrink: 0 }} />}
                                  </div>
                                  <div style={{ display: "flex", alignItems: "center", gap: 2, flexShrink: 0 }}>
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setPreviewSheetId(s.id);
                                      }}
                                      title="View uploaded data"
                                      style={{ ...iconRowButtonStyle, color: theme.brandBlue }}
                                    >
                                      <Eye size={13} />
                                    </button>
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        deleteSheet(s.id);
                                      }}
                                      title="Remove this table"
                                      style={{ ...iconRowButtonStyle, color: theme.textMuted }}
                                    >
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* MIDDLE: configure (top) + preview (below) */}
              <div style={{ display: "flex", flexDirection: "column", gap: 18, minWidth: 0 }}>
                <div style={orangePanelStyle}>
                  {!activeSheet ? (
                    <p style={{ fontSize: 13, color: theme.textSecondary }}>No sheet selected.</p>
                  ) : (
                    <>
                      <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.brandBlue, margin: "0 0 12px", fontWeight: 700 }}>
                        Configure
                      </p>

                      {/* ============================================================
                          NEW — Use Custom SQL Template toggle switch
                          ============================================================ */}
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          marginBottom: 16,
                          padding: "8px 10px",
                          borderRadius: 8,
                          background: theme.brandBlueLight,
                          border: `1px solid ${theme.brandBlueBorder}`,
                        }}
                      >
                        <span style={{ fontSize: 12.5, fontWeight: 600, color: theme.brandBlueDark }}>
                          Use Custom SQL Template
                        </span>
                        <button
                          onClick={() =>
                            updateConfig(activeSheet.id, { customSqlEnabled: !activeConfig.customSqlEnabled })
                          }
                          style={{
                            width: 38,
                            height: 20,
                            borderRadius: 999,
                            border: "none",
                            cursor: "pointer",
                            position: "relative",
                            background: activeConfig.customSqlEnabled ? theme.brandBlue : "#CBD5E1",
                            transition: "background 150ms ease",
                            flexShrink: 0,
                          }}
                          aria-label="Toggle custom SQL template"
                        >
                          <span
                            style={{
                              position: "absolute",
                              top: 2,
                              left: activeConfig.customSqlEnabled ? 20 : 2,
                              width: 16,
                              height: 16,
                              borderRadius: "50%",
                              background: "#fff",
                              transition: "left 150ms ease",
                              boxShadow: "0 1px 2px rgba(0,0,0,0.25)",
                            }}
                          />
                        </button>
                      </div>

                      {activeConfig.customSqlEnabled ? (
                        <>
                          {/* Operation type selector — picking a new operation also
                              rebuilds the template from whatever columns are already
                              checked for that operation. */}
                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            Operation type
                          </label>
                          <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
                            {["INSERT", "UPDATE", "DELETE"].map((op) => (
                              <button
                                key={op}
                                onClick={() => {
                                  const cfg = configs[activeSheet.id];
                                  let newTemplate = "";
                                  if (op === "INSERT") {
                                    newTemplate = buildTemplateFromGuided("INSERT", cfg.tableName, cfg.guidedInsertColumns, null);
                                  } else if (op === "UPDATE") {
                                    newTemplate = buildTemplateFromGuided("UPDATE", cfg.tableName, cfg.guidedSetColumns, cfg.guidedWhereColumns);
                                  } else if (op === "DELETE") {
                                    newTemplate = buildTemplateFromGuided("DELETE", cfg.tableName, null, cfg.guidedWhereColumns);
                                  }
                                  updateConfig(activeSheet.id, { customOperation: op, customTemplate: newTemplate });
                                }}
                                style={{
                                  flex: 1,
                                  padding: "6px 0",
                                  borderRadius: 6,
                                  fontSize: 11.5,
                                  fontWeight: 700,
                                  cursor: "pointer",
                                  border: `1px solid ${activeConfig.customOperation === op ? theme.brandBlue : theme.cardBorder}`,
                                  background: activeConfig.customOperation === op ? theme.brandBlue : "#FBFCFE",
                                  color: activeConfig.customOperation === op ? "#fff" : theme.textPrimary,
                                }}
                              >
                                {op}
                              </button>
                            ))}
                          </div>

                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            Table name
                          </label>
                          <input
                            type="text"
                            value={activeConfig.tableName}
                            onChange={(e) => updateConfig(activeSheet.id, { tableName: e.target.value })}
                            style={{
                              width: "100%",
                              padding: "9px 12px",
                              borderRadius: 8,
                              border: `1px solid ${theme.cardBorder}`,
                              background: "#FBFCFE",
                              color: theme.textPrimary,
                              fontSize: 13,
                              fontFamily: "monospace",
                              boxSizing: "border-box",
                              marginBottom: 14,
                            }}
                          />

                          {/* Column pickers — only the ones relevant to the chosen
                              operation appear. Checking/unchecking a box rebuilds
                              the template below; typing in the template afterward
                              is never overwritten until a box changes again. */}
                          {activeConfig.customOperation === "INSERT" && (
                            <>
                              <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                Columns to INSERT
                              </label>
                              <div style={{ marginBottom: 14 }}>
                                <ColumnCheckboxGrid
                                  headers={activeSheet.headers}
                                  selected={activeConfig.guidedInsertColumns}
                                  onToggle={(h) => toggleGuidedColumn(activeSheet.id, "guidedInsertColumns", h)}
                                />
                              </div>
                            </>
                          )}

                          {activeConfig.customOperation === "UPDATE" && (
                            <>
                              <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                Columns to SET (update)
                              </label>
                              <div style={{ marginBottom: 14 }}>
                                <ColumnCheckboxGrid
                                  headers={activeSheet.headers}
                                  selected={activeConfig.guidedSetColumns}
                                  onToggle={(h) => toggleGuidedColumn(activeSheet.id, "guidedSetColumns", h)}
                                />
                              </div>
                              <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                Match rows WHERE (usually your key column, e.g. MemberId)
                              </label>
                              <div style={{ marginBottom: 14 }}>
                                <ColumnCheckboxGrid
                                  headers={activeSheet.headers}
                                  selected={activeConfig.guidedWhereColumns}
                                  onToggle={(h) => toggleGuidedColumn(activeSheet.id, "guidedWhereColumns", h)}
                                />
                              </div>
                            </>
                          )}

                          {activeConfig.customOperation === "DELETE" && (
                            <>
                              <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                Match rows WHERE (usually your key column, e.g. MemberId)
                              </label>
                              <div style={{ marginBottom: 14 }}>
                                <ColumnCheckboxGrid
                                  headers={activeSheet.headers}
                                  selected={activeConfig.guidedWhereColumns}
                                  onToggle={(h) => toggleGuidedColumn(activeSheet.id, "guidedWhereColumns", h)}
                                />
                              </div>
                            </>
                          )}

                          <button
                            onClick={() => setShowCustomSqlHelp((v) => !v)}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 4,
                              background: "none",
                              border: "none",
                              cursor: "pointer",
                              padding: 0,
                              marginBottom: showCustomSqlHelp ? 8 : 14,
                              fontSize: 11.5,
                              fontWeight: 600,
                              color: theme.brandBlue,
                            }}
                          >
                            <ChevronDown
                              size={12}
                              style={{
                                transform: showCustomSqlHelp ? "rotate(0deg)" : "rotate(-90deg)",
                                transition: "transform 0.2s",
                              }}
                            />
                            How does this work?
                          </button>

                          {showCustomSqlHelp && (
                            <div
                              style={{
                                marginBottom: 14,
                                padding: "10px 12px",
                                borderRadius: 8,
                                background: "#FBFCFE",
                                border: `1px solid ${theme.cardBorder}`,
                                fontSize: 11.5,
                                color: theme.textSecondary,
                                lineHeight: 1.6,
                              }}
                            >
                              <b style={{ color: theme.textPrimary }}>What this does:</b> checking columns above
                              builds a starting SQL template automatically — then you can edit the text below by
                              hand if you need something the checkboxes can't express (a join, a custom condition, etc).
                              <br /><br />
                              <b style={{ color: theme.textPrimary }}>1. Check the columns you need</b> — the
                              template below rebuilds itself each time.
                              <br />
                              <b style={{ color: theme.textPrimary }}>2. Edit freely</b> — your typing is never
                              overwritten unless you check/uncheck a box again.
                              <br />
                              <b style={{ color: theme.textPrimary }}>3. UPDATE/DELETE require a WHERE clause</b> —
                              otherwise every row in your real table is affected.
                              <br />
                              <b style={{ color: theme.textPrimary }}>4. Never wrap a{" "}
                              <code style={{ background: theme.brandBlueLight, padding: "1px 4px", borderRadius: 3 }}>
                                {"{{tag}}"}
                              </code>{" "}
                              in quotes</b> — quotes are added automatically for text values; adding your own
                              doubles them up and breaks the SQL.
                            </div>
                          )}

                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            Available tags (click to insert at the end)
                          </label>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 10 }}>
                            {activeSheet.headers.map((h) => (
                              <button
                                key={h}
                                onClick={() =>
                                  updateConfig(activeSheet.id, {
                                    customTemplate: (activeConfig.customTemplate || "") + `{{${h}}}`,
                                  })
                                }
                                style={{
                                  fontSize: 10.5,
                                  fontFamily: "monospace",
                                  padding: "3px 7px",
                                  borderRadius: 5,
                                  border: `1px solid ${theme.brandBlueBorder}`,
                                  background: theme.brandBlueLight,
                                  color: theme.brandBlueDark,
                                  cursor: "pointer",
                                }}
                              >
                                {`{{${h}}}`}
                              </button>
                            ))}
                          </div>

                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            SQL template (auto-filled from your column picks above — edit freely)
                          </label>
                          <textarea
                            value={activeConfig.customTemplate}
                            onChange={(e) => updateConfig(activeSheet.id, { customTemplate: e.target.value })}
                            placeholder="Check columns above to auto-fill this, or type your own SQL — don't wrap {{tags}} in quotes."
                            rows={7}
                            style={{
                              width: "100%",
                              padding: 10,
                              borderRadius: 8,
                              border: `1px solid ${theme.cardBorder}`,
                              background: "#FBFCFE",
                              color: theme.textPrimary,
                              fontSize: 12,
                              fontFamily: "'JetBrains Mono', monospace",
                              boxSizing: "border-box",
                              resize: "vertical",
                              marginBottom: 10,
                            }}
                          />

                          {/* Optional row filter — target specific rows instead of all */}
                          <div style={{ marginBottom: 12 }}>
                            <label
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "space-between",
                                fontSize: 12,
                                color: theme.textSecondary,
                                marginBottom: 6,
                                cursor: "pointer",
                                fontWeight: 600,
                              }}
                            >
                              <span>Filter which rows to generate for (optional)</span>
                              <input
                                type="checkbox"
                                checked={activeConfig.rowFilterEnabled}
                                onChange={() =>
                                  updateConfig(activeSheet.id, { rowFilterEnabled: !activeConfig.rowFilterEnabled })
                                }
                                style={{ accentColor: theme.brandBlue }}
                              />
                            </label>

                            {activeConfig.rowFilterEnabled && (
                              <div
                                style={{
                                  padding: 10,
                                  borderRadius: 8,
                                  background: "#FBFCFE",
                                  border: `1px solid ${theme.cardBorder}`,
                                }}
                              >
                                <label style={{ fontSize: 11.5, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                  Column to filter by
                                </label>
                                <select
                                  value={activeConfig.rowFilterColumn}
                                  onChange={(e) =>
                                    updateConfig(activeSheet.id, {
                                      rowFilterColumn: e.target.value,
                                      rowFilterValues: new Set(),
                                    })
                                  }
                                  style={{
                                    width: "100%",
                                    padding: "7px 8px",
                                    borderRadius: 6,
                                    border: `1px solid ${theme.cardBorder}`,
                                    background: "#fff",
                                    fontSize: 12,
                                    marginBottom: 10,
                                    boxSizing: "border-box",
                                  }}
                                >
                                  <option value="">-- choose a column --</option>
                                  {activeSheet.headers.map((h) => (
                                    <option key={h} value={h}>
                                      {h}
                                    </option>
                                  ))}
                                </select>

                                {activeConfig.rowFilterColumn && (
                                  <>
                                    <label style={{ fontSize: 11.5, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                                      Search and pick value(s) — only matching rows will be included
                                    </label>
                                    <ValueSearchPicker
                                      options={getUniqueColumnValues(activeSheet.rows, activeConfig.rowFilterColumn)}
                                      selected={activeConfig.rowFilterValues}
                                      onToggle={(val) => toggleGuidedField(activeSheet.id, "rowFilterValues", val)}
                                    />
                                    <p style={{ fontSize: 10.5, color: theme.textMuted, margin: "6px 0 0" }}>
                                      {activeConfig.rowFilterValues.size === 0
                                        ? "No values selected — filter is on but matches nothing yet."
                                        : `${activeConfig.rowFilterValues.size} value(s) selected.`}
                                    </p>
                                  </>
                                )}
                              </div>
                            )}
                          </div>

                          {/* Live preview: rows/statements affected, tags, warnings */}
                          {customSqlPreview && (
                            <div
                              style={{
                                marginBottom: 12,
                                padding: "10px 12px",
                                borderRadius: 8,
                                background: customSqlPreview.validation.valid ? theme.brandBlueLight : theme.dangerBg,
                                border: `1px solid ${
                                  customSqlPreview.validation.valid ? theme.brandBlueBorder : theme.dangerBorder
                                }`,
                                fontSize: 12,
                              }}
                            >
                              <div
                                style={{
                                  fontWeight: 700,
                                  marginBottom: 4,
                                  color: customSqlPreview.validation.valid ? theme.brandBlueDark : theme.danger,
                                  display: "flex",
                                  alignItems: "center",
                                  gap: 5,
                                }}
                              >
                                {!customSqlPreview.validation.valid && <AlertCircle size={13} />}
                                {customSqlPreview.validation.valid
                                  ? customSqlPreview.operation === "INSERT"
                                    ? `Statements to be generated: ${customSqlPreview.rowCount}`
                                    : `Rows to be affected: ${customSqlPreview.rowCount}`
                                  : customSqlPreview.validation.message}
                              </div>

                              <div style={{ color: theme.textSecondary }}>
                                {customSqlPreview.uniqueTags.length > 0
                                  ? `Tags used: ${customSqlPreview.uniqueTags.join(", ")}`
                                  : "No {{tags}} used yet — every row will produce an identical statement."}
                              </div>

                              {customSqlPreview.unknownTags.length > 0 && (
                                <div style={{ color: theme.danger, marginTop: 4, fontWeight: 600 }}>
                                  Unknown column{customSqlPreview.unknownTags.length > 1 ? "s" : ""}:{" "}
                                  {customSqlPreview.unknownTags.join(", ")} — not found in this sheet's headers.
                                </div>
                              )}
                            </div>
                          )}

                          {/* Always-visible caution note */}
                          <div
                            style={{
                              marginBottom: 16,
                              padding: "8px 10px",
                              borderRadius: 8,
                              background: "#FFF7ED",
                              border: "1px solid #FED7AA",
                              fontSize: 11.5,
                              color: "#9A3412",
                              lineHeight: 1.5,
                            }}
                          >
                            <b>Note:</b> Custom SQL is powerful but dangerous. Always verify your WHERE clause
                            to avoid accidental data loss.
                          </div>
                        </>
                      ) : (
                        <>
                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            Table name
                          </label>
                          <div style={{ position: "relative", marginBottom: 16 }}>
                            <Pencil size={12} style={{ position: "absolute", right: 10, top: 11, color: theme.textMuted }} />
                            <input
                              type="text"
                              value={activeConfig.tableName}
                              onChange={(e) => updateConfig(activeSheet.id, { tableName: e.target.value })}
                              style={{
                                width: "100%",
                                padding: "9px 30px 9px 12px",
                                borderRadius: 8,
                                border: `1px solid ${theme.cardBorder}`,
                                background: "#FBFCFE",
                                color: theme.textPrimary,
                                fontSize: 13,
                                fontFamily: "monospace",
                                boxSizing: "border-box",
                              }}
                            />
                          </div>

                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            IF NOT EXISTS - Unique Key
                          </label>
                          <div style={{ marginBottom: 16 }}>
                            <MultiSelectDropdown
                              options={activeSheet.headers}
                              selected={activeConfig.keyColumns}
                              onToggle={(col) => toggleKeyColumn(activeSheet.id, col)}
                            />
                          </div>

                          <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                            Columns to include in the INSERT
                            <span style={{ color: theme.textMuted, fontWeight: 400 }}> (pencil renames the DB column)</span>
                          </label>
                          <div
                            style={{
                              display: "grid",
                              gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
                              gap: 6,
                              marginBottom: 16,
                              maxHeight: 220,
                              overflowY: "auto",
                              paddingRight: 2,
                            }}
                          >
                            {activeSheet.headers.map((h) => {
                              const checked = activeConfig.insertColumns.has(h);
                              const overrideName = activeConfig.columnNameOverrides?.[h];
                              const isRenamed = overrideName && overrideName !== h;
                              const isEditing = editingHeader === h;
                              return (
                                <div
                                  key={h}
                                  style={{
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "space-between",
                                    gap: 4,
                                    padding: "6px 6px 6px 8px",
                                    borderRadius: 6,
                                    border: `1px solid ${checked ? theme.brandBlue : theme.cardBorder}`,
                                    background: checked ? theme.brandBlueLight : "#FBFCFE",
                                    fontSize: 11.5,
                                  }}
                                >
                                  <label
                                    style={{
                                      display: "flex",
                                      alignItems: "center",
                                      gap: 6,
                                      overflow: "hidden",
                                      cursor: "pointer",
                                      flex: 1,
                                      minWidth: 0,
                                    }}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={checked}
                                      onChange={() => toggleInsertColumn(activeSheet.id, h)}
                                      style={{ accentColor: theme.brandBlue, flexShrink: 0 }}
                                    />
                                    {isEditing ? (
                                      <input
                                        autoFocus
                                        value={editValue}
                                        onClick={(e) => e.preventDefault()}
                                        onChange={(e) => setEditValue(e.target.value)}
                                        onBlur={() => commitColumnRename(h)}
                                        onKeyDown={(e) => {
                                          if (e.key === "Enter") commitColumnRename(h);
                                          if (e.key === "Escape") cancelColumnRename();
                                        }}
                                        style={{
                                          width: "100%",
                                          minWidth: 0,
                                          fontSize: 11.5,
                                          padding: "2px 4px",
                                          borderRadius: 4,
                                          border: `1px solid ${theme.brandBlue}`,
                                          background: "#fff",
                                          color: theme.textPrimary,
                                        }}
                                      />
                                    ) : (
                                      <span
                                        title={h}
                                        style={{
                                          overflow: "hidden",
                                          textOverflow: "ellipsis",
                                          whiteSpace: "nowrap",
                                          color: isRenamed ? theme.brandBlueDark : theme.textPrimary,
                                          fontWeight: isRenamed ? 700 : 500,
                                        }}
                                      >
                                        {overrideName || h}
                                      </span>
                                    )}
                                  </label>
                                  {!isEditing && (
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        startEditingColumn(h);
                                      }}
                                      title={`Rename column (Excel: ${h})`}
                                      style={{
                                        background: "none",
                                        border: "none",
                                        color: theme.textMuted,
                                        cursor: "pointer",
                                        padding: 2,
                                        flexShrink: 0,
                                        display: "flex",
                                      }}
                                    >
                                      <Pencil size={11} />
                                    </button>
                                  )}
                                </div>
                              );
                            })}
                          </div>

                          <label
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                              fontSize: 13,
                              marginBottom: 18,
                              cursor: "pointer",
                              color: theme.textPrimary,
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={activeConfig.identityInsert}
                              onChange={() =>
                                updateConfig(activeSheet.id, { identityInsert: !activeConfig.identityInsert })
                              }
                              style={{ accentColor: theme.brandBlue }}
                            />
                            Wrap with SET IDENTITY_INSERT ON / OFF
                          </label>
                        </>
                      )}

                      <button
                        onClick={generateForActiveSheet}
                        style={{
                          width: "100%",
                          padding: "10px 18px",
                          borderRadius: 8,
                          background: theme.accentOrange,
                          color: "#fff",
                          fontSize: 13,
                          fontWeight: 700,
                          border: "none",
                          cursor: "pointer",
                        }}
                      >
                        Generate SQL for this table
                      </button>
                    </>
                  )}
                </div>

                <div style={orangePanelStyle}>
                  <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.brandBlue, margin: "0 0 12px", fontWeight: 700 }}>
                    SQL Preview
                  </p>
                  {activeConfig?.generated ? (
                    <>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <span style={{ fontSize: 12, color: theme.textSecondary, fontFamily: "monospace" }}>
                          {activeConfig.tableName}.sql
                        </span>
                        <button
                          onClick={() => copyToClipboard(activeConfig.generated, activeSheet.id)}
                          style={outlineButtonStyle}
                        >
                          {copiedId === activeSheet.id ? <Check size={12} /> : <Copy size={12} />}
                          {copiedId === activeSheet.id ? "Copied" : "Copy"}
                        </button>
                      </div>
                      <pre
                        style={{
                          background: theme.panelBg,
                          border: `1px solid ${theme.panelBorder}`,
                          borderRadius: 10,
                          padding: "14px 14px 20px",
                          fontSize: 11.5,
                          lineHeight: 1.6,
                          color: theme.codeText,
                          overflow: "auto",
                          height: 340,
                          minHeight: 160,
                          maxHeight: "65vh",
                          resize: "vertical",
                          fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
                          margin: 0,
                          whiteSpace: "pre",
                        }}
                      >
                        {activeConfig.generated}
                      </pre>
                      <p style={{ fontSize: 10.5, color: theme.textMuted, margin: "6px 0 0", textAlign: "right" }}>
                        Drag the bottom-right corner to resize ↘
                      </p>
                    </>
                  ) : (
                    <p style={{ fontSize: 12.5, color: theme.textMuted }}>
                      Configure the table above, then generate to see the script here.
                    </p>
                  )}
                </div>
              </div>

              {/* RIGHT: export */}
              <div style={orangePanelStyle}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.brandBlue, margin: "0 0 12px", fontWeight: 700 }}>
                  Export
                </p>

                {generatedSheets.length === 0 ? (
                  <p style={{ fontSize: 12.5, color: theme.textMuted }}>
                    Generate at least one table to enable export.
                  </p>
                ) : (
                  <>
                    <button
                      onClick={() => downloadTextFile("all_tables.sql", combinedScript)}
                      style={{
                        width: "100%",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: 6,
                        padding: "10px 14px",
                        borderRadius: 8,
                        background: theme.brandBlue,
                        color: "#ffffff",
                        fontWeight: 700,
                        fontSize: 13,
                        border: "none",
                        cursor: "pointer",
                        marginBottom: 14,
                      }}
                    >
                      <Download size={14} /> Export all ({generatedSheets.length})
                    </button>

                    <p style={{ fontSize: 11, color: theme.textMuted, margin: "0 0 8px", fontWeight: 600 }}>
                      Individually
                    </p>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {generatedSheets.map((s) => (
                        <div
                          key={s.id}
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            padding: "8px 10px",
                            borderRadius: 7,
                            background: "#FBFCFE",
                            border: `1px solid ${theme.cardBorder}`,
                          }}
                        >
                          <span
                            style={{
                              fontSize: 12,
                              fontFamily: "monospace",
                              color: theme.textPrimary,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {configs[s.id].tableName}.sql
                          </span>
                          <button
                            onClick={() =>
                              downloadTextFile(`${configs[s.id].tableName}.sql`, configs[s.id].generated)
                            }
                            title="Export this table"
                            style={{
                              background: "none",
                              border: "none",
                              color: theme.brandBlue,
                              cursor: "pointer",
                              padding: 2,
                              flexShrink: 0,
                              display: "flex",
                            }}
                          >
                            <Download size={14} />
                          </button>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </div>

      {previewSheet && (
        <DataPreviewModal sheet={previewSheet} onClose={() => setPreviewSheetId(null)} />
      )}
    </div>
  );
}
