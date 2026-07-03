import React, { useState, useRef, useEffect } from "react";
import { ChevronDown } from "lucide-react";
import { theme } from "../theme";

export function MultiSelectDropdown({ options, selected, onToggle }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const summary =
    selected.size === 0 ? "Select column(s)..." : Array.from(selected).join(" : ");

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((o) => !o)}
        style={{
          width: "100%",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "9px 12px",
          borderRadius: 8,
          border: `1px solid ${selected.size > 0 ? theme.brandBlue : theme.cardBorder}`,
          background: selected.size > 0 ? theme.brandBlueLight : "#FBFCFE",
          color: selected.size > 0 ? theme.brandBlueDark : theme.textSecondary,
          fontSize: 12.5,
          cursor: "pointer",
          textAlign: "left",
        }}
      >
        <span
          style={{
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
            fontFamily: selected.size > 0 ? "monospace" : "inherit",
          }}
        >
          {summary}
        </span>
        <ChevronDown size={14} style={{ flexShrink: 0, marginLeft: 8 }} />
      </button>

      {open && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 6px)",
            left: 0,
            right: 0,
            zIndex: 20,
            background: theme.cardBg,
            border: `1px solid ${theme.cardBorder}`,
            borderRadius: 10,
            padding: 8,
            maxHeight: 220,
            overflowY: "auto",
            boxShadow: "0 12px 28px rgba(16,24,40,0.14)",
          }}
        >
          {options.map((opt) => {
            const checked = selected.has(opt);
            return (
              <label
                key={opt}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "7px 8px",
                  borderRadius: 6,
                  fontSize: 12.5,
                  cursor: "pointer",
                  color: theme.textPrimary,
                  background: checked ? theme.brandBlueLight : "transparent",
                }}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => onToggle(opt)}
                  style={{ accentColor: theme.brandBlue }}
                />
                {opt}
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
}
