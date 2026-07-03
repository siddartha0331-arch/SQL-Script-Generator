// Detects columns like "Created Date" / "ModifiedDate" so we can force
// them to CURRENT_TIMESTAMP instead of copying whatever was in Excel.
export const isAutoTimestampColumn = (header) => {
  const normalized = header.toUpperCase().replace(/[^A-Z]/g, "");
  return (
    (normalized.includes("CREATED") && normalized.includes("DATE")) ||
    (normalized.includes("MODIFIED") && normalized.includes("DATE"))
  );
};

// SQL strings escape a single quote by doubling it: O'Brien -> O''Brien.
export const escapeSqlString = (value) => String(value).replace(/'/g, "''");

// Turns one Excel cell into the correct SQL literal for its type.
export const formatValue = (header, rawValue) => {
  if (isAutoTimestampColumn(header)) return "CURRENT_TIMESTAMP";
  if (rawValue === undefined || rawValue === null || rawValue === "")
    return "NULL";
  if (typeof rawValue === "number") return String(rawValue);
  if (typeof rawValue === "boolean") return rawValue ? "1" : "0";
  return `'${escapeSqlString(rawValue)}'`;
};

// Builds the WHERE clause used inside "IF NOT EXISTS (...)" to check for
// duplicates before inserting.
export const buildWhereClause = (row, keyColumns, columnOverrides = {}) =>
  keyColumns
    .map((col) => `${columnOverrides[col] || col}=${formatValue(col, row[col])}`)
    .join(" AND ");

// insertColumnPairs = [{ original, display }] — original is the Excel
// header, display is the renamed column name (falls back to original).
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

// Builds the full script for one sheet: one buildRowScript() per row,
// optionally wrapped in SET IDENTITY_INSERT ON/OFF.
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
