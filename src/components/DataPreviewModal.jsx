import React from "react";
import { Table2, X } from "lucide-react";
import { theme } from "../theme";

export function DataPreviewModal({ sheet, onClose }) {
  if (!sheet) return null;
  const previewRows = sheet.rows.slice(0, 100);

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(16,24,40,0.45)",
        zIndex: 100,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 24,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: theme.cardBg,
          borderRadius: 14,
          width: "min(1100px, 100%)",
          maxHeight: "85vh",
          display: "flex",
          flexDirection: "column",
          boxShadow: "0 24px 60px rgba(16,24,40,0.28)",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            padding: "16px 20px",
            borderBottom: `1px solid ${theme.cardBorder}`,
            flexShrink: 0,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10, overflow: "hidden" }}>
            <Table2 size={16} color={theme.brandBlue} style={{ flexShrink: 0 }} />
            <div style={{ overflow: "hidden" }}>
              <p style={{ margin: 0, fontSize: 14, fontWeight: 700, color: theme.textPrimary }}>
                {sheet.originalName}
              </p>
              <p style={{ margin: 0, fontSize: 11.5, color: theme.textMuted }}>
                {sheet.sourceFile} · {sheet.rows.length} row{sheet.rows.length !== 1 ? "s" : ""} ·{" "}
                {sheet.headers.length} column{sheet.headers.length !== 1 ? "s" : ""}
                {sheet.rows.length > 100 ? " · showing first 100" : ""}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              color: theme.textMuted,
              padding: 4,
              display: "flex",
              flexShrink: 0,
            }}
          >
            <X size={18} />
          </button>
        </div>

        <div style={{ overflow: "auto", flex: 1 }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
            <thead>
              <tr>
                {sheet.headers.map((h) => (
                  <th
                    key={h}
                    style={{
                      position: "sticky",
                      top: 0,
                      background: theme.brandBlueLight,
                      color: theme.brandBlueDark,
                      textAlign: "left",
                      padding: "9px 12px",
                      fontWeight: 700,
                      whiteSpace: "nowrap",
                      borderBottom: `1px solid ${theme.brandBlueBorder}`,
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewRows.map((row, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? theme.cardBg : "#FBFCFE" }}>
                  {sheet.headers.map((h) => (
                    <td
                      key={h}
                      style={{
                        padding: "8px 12px",
                        color: theme.textPrimary,
                        whiteSpace: "nowrap",
                        borderBottom: `1px solid ${theme.cardBorder}`,
                      }}
                    >
                      {row[h] === null || row[h] === undefined || row[h] === "" ? (
                        <span style={{ color: theme.textMuted }}>NULL</span>
                      ) : (
                        String(row[h])
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
