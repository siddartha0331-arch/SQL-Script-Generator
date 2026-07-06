export const theme = {

  pageBg: "#0B1220",          // unchanged — this one was already right
  cardBg: "#121C33",          // was white → now a dark navy surface, one shade up from pageBg
  cardBorder: "#26324A",      // was a light grey border → needs to be visible on dark, not darker
  cardShadow: "0 1px 2px rgba(0,0,0,0.35), 0 1px 3px rgba(0,0,0,0.45)", // shadows need to be darker/stronger on dark bg, a light shadow just disappears
  textPrimary: "#E8ECF6",     // was near-black → flipped to near-white
  textSecondary: "#96A2BC",   // was medium-grey → lightened so it's still readable, but dimmer than textPrimary
  textMuted: "#6C7690",       // kept the "dimmest" role, just lightened enough to stay legible on dark

  // Header stays blue-on-blue — this pairing already worked in your screenshot, don't touch it
  brandBlue: "#0B5FA5",
  brandBlueDark: "#084A82",
  brandBlueLight: "#152A42",  // was a pale blue tint (for white bg) → now a dark blue tint (for dark bg), same ROLE, inverted lightness
  brandBlueBorder: "#3A6FA5", // needs to be lighter than brandBlueLight to still read as a border
  headerBgFrom: "#0B5FA5",
  headerBgTo: "#073E68",
  headerText: "#FFFFFF",
  headerSubtext: "#CFE3F7",

  // Orange accent panels (Sheets / Export) — same idea, dark + warm instead of light + warm
  accentOrange: "#F2994A",
  accentOrangeDark: "#D9822B",
  panelBg: "#1E160E",         // was cream (FFF7F0) → dark, but keeping a warm brown undertone so it still reads as "the orange-accented panel," not just another navy card
  panelBorder: "#3E2E1C",
  panelAccent: "#F2994A",
  panelAccentDark: "#D9822B",
  panelAccentText: "#F2C08A", // was a dark brown (8A4A12) for text on cream → needs to be a light warm tone for text on dark brown

  success: "#3BC48B",         // brightened slightly so it still pops against dark
  successBg: "#0F2A1E",       // was pale green → dark green tint
  danger: "#FF6B6B",
  dangerBg: "#2A1414",        // was pale red → dark red tint
  dangerBorder: "#5C2A2A",

  codeBg: "#080D1A",          // darkest surface, terminal-like — unchanged in spirit from before
  codeBorder: "#1B2740",
  codeText: "#CBD5E6",        // was dark navy text on light code bg → light grey-blue text on near-black
};
