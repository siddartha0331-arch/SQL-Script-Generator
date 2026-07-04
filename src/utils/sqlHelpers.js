
export const isAutoTimestampColumn = (header) => {
  const normalized = header.toUpperCase().replace(/[^A-Z]/g, "");
  return (
    (normalized.includes("CREATED") && normalized.includes("DATE")) ||
    (normalized.includes("MODIFIED") && normalized.includes("DATE"))
  );
};

function formatSqlValue(value) {
  if (value === null || value === undefined || value === "") return "NULL";
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "1" : "0";
  const escaped = String(value).replace(/'/g, "''");
  return `'${escaped}'`;
}

export const escapeSqlString = (value) => String(value).replace(/'/g, "''");
export function hasWhereClause(template) {
  return /\bWHERE\b/i.test(template);
}

export const formatValue = (header, rawValue) => {
  if (isAutoTimestampColumn(header)) return "CURRENT_TIMESTAMP";
  if (rawValue === undefined || rawValue === null || rawValue === "")
    return "NULL";
  if (typeof rawValue === "number") return String(rawValue);
  if (typeof rawValue === "boolean") return rawValue ? "1" : "0";
  return `'${escapeSqlString(rawValue)}'`;
};

function extractWhereClause(template) {
  const match = template.match(/\bWHERE\b([\s\S]*)/i);
  return match ? match[1] : "";
}

export const buildWhereClause = (row, keyColumns, columnOverrides = {}) =>
  keyColumns
    .map((col) => `${columnOverrides[col] || col}=${formatValue(col, row[col])}`)
    .join(" AND ");

export const buildRowScript = (
  tableName,
  insertColumnPairs,
  row,
  keyColumns,
  columnOverrides
) => {
  const whereClause = buildWhereClause(row, keyColumns, columnOverrides);
  const columnList = insertColumnPairs.map((p) => `[${p.display}]`).join(",");
  const valueList = insertColumnPairs
    .map((p) => formatValue(p.original, row[p.original]))
    .join(",");

  return (
    `IF NOT EXISTS (SELECT 1 FROM dbo.${tableName} WHERE ${whereClause})\n` +
    ` BEGIN \n` +
    ` INSERT dbo.${tableName} (${columnList})\n` +
    ` VALUES (${valueList})\n` +
    ` END \n`
  );
};


export const buildTableScript = (
  sheet,
  tableName,
  keyColumns,
  insertColumns,
  columnOverrides,
  identityInsert
) => {
  const insertColumnPairs = sheet.headers
    .filter((h) => insertColumns.has(h))
    .map((h) => ({ original: h, display: (columnOverrides[h] || "").trim() || h }));

  const parts = [];
  if (identityInsert) parts.push(`SET IDENTITY_INSERT ${tableName} ON\n`);
  sheet.rows.forEach((row) => {
    parts.push(buildRowScript(tableName, insertColumnPairs, row, keyColumns, columnOverrides));
  });
  if (identityInsert) parts.push(`SET IDENTITY_INSERT ${tableName} OFF `);

  return parts.join("\n");
};

// Triggers a browser file download for a plain-text file.
export const downloadTextFile = (filename, text) => {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// Catches obvious "always true" patterns written literally (1=1, 'a'='a', TRUE).
// Heuristic, not a SQL parser — can't catch every possible tautology shape.
function isObviousTautology(whereClause) {
  const patterns = [/\b(\d+)\s*=\s*\1\b/, /'([^']*)'\s*=\s*'\1'/, /\bTRUE\b/i];
  return patterns.some((p) => p.test(whereClause));
}
 
// before generateCustomSQL() is ever called — if it fails, nothing generates.
export function validateTemplate(template, operation) {
  if (!template || !template.trim()) {
    return { valid: false, message: "Your custom SQL template is empty." };
  }
  if (operation === "INSERT") {
    return { valid: true, message: "" };
  }
  if (!hasWhereClause(template)) {
    return { valid: false, message: "SECURITY WARNING: Missing WHERE clause. This will affect all rows." };
  }
  const whereClause = extractWhereClause(template);
  if (isObviousTautology(whereClause)) {
    return {
      valid: false,
      message: "SECURITY WARNING: WHERE clause looks like it's always true (e.g. 1=1). This will affect all rows.",
    };
  }
  if (!/\{\{\s*[^}]+?\s*\}\}/.test(whereClause)) {
    return {
      valid: false,
      message: "SECURITY WARNING: WHERE clause doesn't reference any {{column}} — every row will use the exact same condition.",
    };
  }
  return { valid: true, message: "" };
}
 
// already adds quotes for text values, so manual quotes would double them up.
export function generateCustomSQL(template, rows) {
  return rows
    .map((row) => template.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (m, colName) => formatSqlValue(row[colName])))
    .join("\n");
}
 
// Sorted unique values in a column — powers the "search and pick rows" filter.
export function getUniqueColumnValues(rows, column) {
  const values = new Set();
  rows.forEach((row) => {
    const v = row[column];
    if (v !== null && v !== undefined && v !== "") values.add(v);
  });
  return Array.from(values).sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
}
 
// Narrows rows to only those matching selectedValues on `column`.
export function filterRowsByColumnValues(rows, column, selectedValues) {
  if (!column || !selectedValues || selectedValues.size === 0) return rows;
  return rows.filter((row) => selectedValues.has(row[column]));
}
 
// Turns Guided-Builder checkbox selections into an actual template string.
export function buildTemplateFromGuided(operation, tableName, primaryColumns, whereColumns) {
  const table = (tableName || "MyTable").trim() || "MyTable";
 
  if (operation === "INSERT") {
    const cols = Array.from(primaryColumns);
    if (cols.length === 0) return "";
    return `INSERT INTO ${table} (${cols.join(", ")})\nVALUES (${cols.map((c) => `{{${c}}}`).join(", ")});`;
  }
  if (operation === "UPDATE") {
    const setCols = Array.from(primaryColumns);
    const whereCols = Array.from(whereColumns);
    if (setCols.length === 0 || whereCols.length === 0) return "";
    const setList = setCols.map((c) => `${c} = {{${c}}}`).join(",\n    ");
    const whereList = whereCols.map((c) => `${c} = {{${c}}}`).join(" AND ");
    return `UPDATE ${table}\nSET ${setList}\nWHERE ${whereList};`;
  }
  if (operation === "DELETE") {
    const whereCols = Array.from(whereColumns);
    if (whereCols.length === 0) return "";
    return `DELETE FROM ${table}\nWHERE ${whereCols.map((c) => `${c} = {{${c}}}`).join(" AND ")};`;
  }
  return "";
}
 
// THE single place that decides "what template actually runs" —

export function getEffectiveTemplate(cfg) {
  if (cfg.builderMode === "raw") return cfg.customTemplate || "";
  if (cfg.customOperation === "INSERT") return buildTemplateFromGuided("INSERT", cfg.tableName, cfg.guidedInsertColumns, null);
  if (cfg.customOperation === "UPDATE") return buildTemplateFromGuided("UPDATE", cfg.tableName, cfg.guidedSetColumns, cfg.guidedWhereColumns);
  if (cfg.customOperation === "DELETE") return buildTemplateFromGuided("DELETE", cfg.tableName, null, cfg.guidedWhereColumns);
  return "";
}
