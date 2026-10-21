import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ─── LOCAL STORAGE SHIM ────────────────────────────────────────────────────
// Replaces Claude artifact's window.storage API with localStorage for standalone use
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key) {
      try {
        const val = localStorage.getItem("bs_" + key);
        return val !== null ? { key, value: val } : null;
      } catch { return null; }
    },
    async set(key, value) {
      try {
        localStorage.setItem("bs_" + key, value);
        return { key, value };
      } catch { return null; }
    },
    async delete(key) {
      try {
        localStorage.removeItem("bs_" + key);
        return { key, deleted: true };
      } catch { return null; }
    },
    async list(prefix) {
      try {
        const keys = Object.keys(localStorage)
          .filter(k => k.startsWith("bs_"))
          .map(k => k.slice(3))
          .filter(k => !prefix || k.startsWith(prefix));
        return { keys };
      } catch { return { keys: [] }; }
    },
  };
}


// ─── QUESTION DEFINITIONS ───────────────────────────────────────────────────
// Column index in Excel → question metadata
const QUESTIONS = [
  { idx: 0, num: "1", text: "Are cattle being purchased?", topic: "A" },
  { idx: 1, num: "2.1", text: "How often a year are pregnant cows bought?", topic: "A" },
  { idx: 2, num: "2.2", text: "How often a year are lactating cows bought?", topic: "A" },
  { idx: 3, num: "2.3", text: "How often a year are calves or nonpregnant heifers bought?", topic: "A" },
  { idx: 4, num: "2.4", text: "How often a year are bulls for reproduction bought?", topic: "A" },
  { idx: 5, num: "3", text: "Are your cattle during the past 2 years always bought from the same original source?", topic: "A" },
  { idx: 6, num: "4", text: "Before the cattle arrives on your farm, is contact between your cattle and animals from different farms possible?", topic: "A" },
  { idx: 7, num: "5", text: "Whenever cattle are bought from another farm, is proof requested to ensure sanitary statute and health management of the farm of origin is equal or higher?", topic: "A" },
  { idx: 8, num: "6", text: "Is the level of maternal immunity checked when buying calves?", topic: "A" },
  { idx: 9, num: "7", text: "Are the cattle tested for specific diseases when entering the farm?", topic: "A" },
  { idx: 10, num: "8", text: "Are all new cattle put into quarantine?", topic: "A" },
  { idx: 12, num: "9", text: "What is the minimum duration (in days) of the quarantine period?", topic: "A" },
  { idx: 13, num: "10", text: "Are ... before entering the quarantine?", topic: "A" },
  { idx: 14, num: "11", text: "Is the quarantine empty after an animal has been in quarantine?", topic: "A" },
  { idx: 15, num: "12", text: "Is the quarantine area ... before the introduction of new cattle?", topic: "A" },
  { idx: 16, num: "13", text: "Are the newly introduced cows milked separately during their quarantine period?", topic: "A" },
  { idx: 17, num: "14", text: "Is a milk sample from the newly introduced cows taken and tested?", topic: "A" },
  { idx: 18, num: "15", text: "Are there any cattle that leave the farm and return afterwards?", topic: "A" },
  { idx: 19, num: "16", text: "Are these returning cattle put into quarantine like described before?", topic: "A" },
  { idx: 20, num: "17", text: "Are the cattle on your farm bred? If yes, how?", topic: "A" },
  { idx: 21, num: "18", text: "Has the bull's semen been tested for sexually transmitted diseases?", topic: "A" },
  { idx: 22, num: "19", text: "Does the semen come from a farm/institution with health status known to be higher or equal?", topic: "A" },
  { idx: 23, num: "20", text: "Do all vehicles have to pass through clean transport baths before entering the farm?", topic: "B" },
  { idx: 24, num: "21", text: "Do external transport vehicles and transporters have access to areas where cattle are kept?", topic: "B" },
  { idx: 25, num: "22", text: "When cattle are delivered, are only the animals supposed to be delivered in the transport vehicle?", topic: "B" },
  { idx: 26, num: "23", text: "Is the transport vehicle for the cattle empty on arrival at the farm?", topic: "B" },
  { idx: 27, num: "24", text: "Is the transport vehicle always cleaned and disinfected before entering the farm?", topic: "B" },
  { idx: 28, num: "25", text: "Is there a separate carcass storage space with a hard surface floor present?", topic: "B" },
  { idx: 29, num: "26", text: "Is this carcass storage space cleaned and disinfected after each use?", topic: "B" },
  { idx: 30, num: "27", text: "Is the carcass storage space protected from vermin, cats and dogs?", topic: "B" },
  { idx: 31, num: "28", text: "Can the carcasses be removed by the rendering company without them entering the premises?", topic: "B" },
  { idx: 32, num: "29", text: "Are carcasses manipulated with gloves, or are hands cleaned and disinfected after?", topic: "B" },
  { idx: 33, num: "30", text: "Is all the material used for the manipulation of carcasses cleaned and disinfected?", topic: "B" },
  { idx: 34, num: "31", text: "Are the feed storage facilities protected from pets and vermin?", topic: "C" },
  { idx: 35, num: "32", text: "Are feeding utensils used only for feed?", topic: "C" },
  { idx: 37, num: "33", text: "Is the quality of the drinking water checked every year at the source by bacteriological analysis?", topic: "C" },
  { idx: 38, num: "34", text: "Is the quality of the drinking water checked every year at the main outlets?", topic: "C" },
  { idx: 39, num: "35", text: "Are visitors obliged to notify you of their presence before entering the stables?", topic: "D" },
  { idx: 40, num: "36", text: "Is there a separate space available for changing boots and clothes and washing hands?", topic: "D" },
  { idx: 41, num: "37", text: "Are there any farmworkers who also work at other farms?", topic: "D" },
  { idx: 42, num: "38", text: "Upon entering the farm, does the farm personnel ...", topic: "D" },
  { idx: 43, num: "39", text: "Upon entering the farm, does the veterinarian...", topic: "D" },
  { idx: 44, num: "40", text: "Does the artificial insemination technician come to the farm?", topic: "D" },
  { idx: 45, num: "41", text: "Upon entering the farm, does the artificial insemination technician...", topic: "D" },
  { idx: 46, num: "42", text: "Does the cattle salesman come to the farm?", topic: "D" },
  { idx: 47, num: "43", text: "Upon entering the farm, does the cattle salesman...", topic: "D" },
  { idx: 48, num: "44", text: "Does the hoof trimmer come to the farm?", topic: "D" },
  { idx: 49, num: "45", text: "Upon entering the farm, does the hoof trimmer...", topic: "D" },
  { idx: 50, num: "46", text: "Are there any other visitors that enter the farm and come into contact with the cattle?", topic: "D" },
  { idx: 51, num: "47", text: "Upon entering the farm, do these other visitors...", topic: "D" },
  { idx: 52, num: "48", text: "Is an insect control programme present on the farm?", topic: "E" },
  { idx: 53, num: "49", text: "Is a rodent control programme present on the farm?", topic: "E" },
  { idx: 54, num: "50", text: "Is a bird control programme present on the farm?", topic: "E" },
  { idx: 55, num: "51", text: "Do your cattle, including the youngstock, have access to the outside?", topic: "E" },
  { idx: 56, num: "52", text: "When your cattle go outside, do they have access to natural water bodies?", topic: "E" },
  { idx: 57, num: "53", text: "Is it possible for your cattle to come into contact with animals from other farms?", topic: "E" },
  { idx: 59, num: "54", text: "Are there other commercially exploited cattle present on the farm?", topic: "E" },
  { idx: 60, num: "55", text: "Can these other commercially exploited cattle come into contact with the dairy cows?", topic: "E" },
  { idx: 61, num: "56", text: "Are any other farm animals being kept?", topic: "E" },
  { idx: 62, num: "57", text: "Do pets have access to the stables?", topic: "E" },
  { idx: 63, num: "58", text: "Is manure from other farms being spread on farmlands within 500m of your farm?", topic: "E" },
  { idx: 64, num: "59", text: "Are the sick cattle physically isolated from the healthy cattle?", topic: "F" },
  { idx: 65, num: "60", text: "Are there equipment and materials specific for the sick cattle in the hospital pen?", topic: "F" },
  { idx: 66, num: "61", text: "Is this specific equipment cleaned and disinfected before a new animal enters?", topic: "F" },
  { idx: 67, num: "62", text: "Are the cattle in the hospital pen fully separated from the other cattle?", topic: "F" },
  { idx: 68, num: "63", text: "Are ... before entering the hospital pen?", topic: "F" },
  { idx: 69, num: "64", text: "Is the hospital pen empty after each use?", topic: "F" },
  { idx: 70, num: "65", text: "Is the hospital pen ... before each new introduction of sick cattle?", topic: "F" },
  { idx: 71, num: "66", text: "Are the sick cattle taken care of before or after the healthy cattle?", topic: "F" },
  { idx: 72, num: "67", text: "Can a unit of sick cattle be completely separated in case of a disease outbreak?", topic: "F" },
  { idx: 74, num: "68", text: "Is a register with the animal health data being kept?", topic: "F" },
  { idx: 75, num: "69", text: "Are there written protocols for vaccination, disease treatment and hygiene?", topic: "F" },
  { idx: 76, num: "70", text: "What happens to the disease carriers that are detected?", topic: "F" },
  { idx: 77, num: "71", text: "Are there dedicated injection needles specific to each age group?", topic: "F" },
  { idx: 78, num: "72", text: "Are there maternity pens and/or a box for C-sections available?", topic: "F" },
  { idx: 79, num: "73", text: "Is the maternity pen ever used to house sick cattle or adjacent to sick cattle?", topic: "F" },
  { idx: 80, num: "74", text: "Are cattle in the maternity pen fully separated from the other animals?", topic: "F" },
  { idx: 81, num: "75", text: "Are ... in the maternity pen?", topic: "F" },
  { idx: 82, num: "76", text: "Is the maternity pen ... before each new introduction of cattle?", topic: "F" },
  { idx: 83, num: "77", text: "Are hands and obstetric materials always cleaned and disinfected before and after each calving/abortion?", topic: "G" },
  { idx: 85, num: "78", text: "Are the cow's hindquarters always cleaned and disinfected before each calving?", topic: "G" },
  { idx: 86, num: "79", text: "When does the separation of the calf from the mother take place?", topic: "G" },
  { idx: 87, num: "80", text: "If an abortion takes place, is the cow tested afterwards?", topic: "G" },
  { idx: 88, num: "81", text: "Where are the foetal membranes and tissues disposed of after a calving/abortion?", topic: "G" },
  { idx: 89, num: "82", text: "How many litres of colostrum are administered to the calf within the first six hours?", topic: "G" },
  { idx: 90, num: "83", text: "Is the colostrum given from the mother or frozen colostrum from healthy cows?", topic: "G" },
  { idx: 91, num: "84", text: "Is it checked if the colostrum quality is sufficient?", topic: "G" },
  { idx: 92, num: "85", text: "Is there a frozen or artificial reserve of colostrum present?", topic: "G" },
  { idx: 93, num: "86", text: "If colostrum is not administered directly after milking, is it stored in the refrigerator?", topic: "G" },
  { idx: 94, num: "87", text: "Are the materials used for colostrum administration cleaned and disinfected after each use?", topic: "G" },
  { idx: 95, num: "88", text: "Are the calves housed in individual calf boxes/hutches or separate areas?", topic: "H" },
  { idx: 96, num: "89", text: "Are the individual calf boxes/hutches/separate areas empty after each use?", topic: "H" },
  { idx: 97, num: "90", text: "Are the individual calf boxes/hutches/separate areas ... before each new introduction?", topic: "H" },
  { idx: 98, num: "91", text: "Is contact possible with calves in different hutches/boxes/separate areas?", topic: "H" },
  { idx: 99, num: "92", text: "Are milk feeding buckets/teats reused between calves during the same feeding session?", topic: "H" },
  { idx: 100, num: "93", text: "Are the calves ever fed with waste milk?", topic: "H" },
  { idx: 102, num: "94", text: "Are the feeding buckets cleaned after each feeding?", topic: "H" },
  { idx: 103, num: "95", text: "How large are the groups of calves regrouped from individual hutches to group pens?", topic: "H" },
  { idx: 104, num: "96", text: "Is the group housing empty after each use?", topic: "H" },
  { idx: 105, num: "97", text: "Is the group housing ... before each new introduction of calves?", topic: "H" },
  { idx: 106, num: "98", text: "Are the cows milked with a milking robot or manually?", topic: "I" },
  { idx: 107, num: "99", text: "How many times per year is a static measurement of milking equipment performed?", topic: "I" },
  { idx: 108, num: "100", text: "How many cows are milked on average?", topic: "I" },
  { idx: 109, num: "101.1", text: "How often a day do you milk?", topic: "I" },
  { idx: 110, num: "101.2", text: "What is the average number of milkings per cow per day?", topic: "I" },
  { idx: 111, num: "102", text: "How many milking clusters are there in the milking parlour?", topic: "I" },
  { idx: 112, num: "103", text: "Do you use rubber or silicone teat cup liners?", topic: "I" },
  { idx: 113, num: "104", text: "After how many months are the teat cup liners replaced?", topic: "I" },
  { idx: 114, num: "105", text: "Are milking clusters being disinfected between cows?", topic: "I" },
  { idx: 115, num: "106", text: "How are the milking clusters between cows disinfected?", topic: "I" },
  { idx: 116, num: "107", text: "Are the teats cleaned before milking? If yes, how?", topic: "I" },
  { idx: 117, num: "108", text: "Is the foremilk examined during fore-stripping?", topic: "I" },
  { idx: 118, num: "109", text: "Are the teats disinfected after the teat cups are removed?", topic: "I" },
  { idx: 120, num: "110", text: "Are cows kept upright for a period after milking?", topic: "I" },
  { idx: 121, num: "111", text: "Are the cows milked in a specific order?", topic: "I" },
  { idx: 122, num: "112", text: "Are the udders of the lactating cows clipped?", topic: "J" },
  { idx: 123, num: "113", text: "Are the tails of the lactating cows clipped?", topic: "J" },
  { idx: 124, num: "114", text: "Is there a regular bacterial examination of the udder of all cows?", topic: "J" },
  { idx: 125, num: "115.1", text: "How often a year is the adult stable cleaned?", topic: "J" },
  { idx: 127, num: "115.2", text: "How often a year is the adult stable disinfected?", topic: "J" },
  { idx: 128, num: "116", text: "In which groups are the cows on your farm divided?", topic: "J" },
  { idx: 129, num: "117", text: "Do the cows have to regularly pass through a hoof disinfection footbath?", topic: "J" },
  { idx: 130, num: "118", text: "Are the cattle grouped per age category in the stable?", topic: "K" },
  { idx: 131, num: "119", text: "Has a full separation between age groups been established?", topic: "K" },
  { idx: 132, num: "120", text: "Are ... between age groups?", topic: "K" },
  { idx: 133, num: "121", text: "Is farm work performed in a specific order?", topic: "K" },
  { idx: 134, num: "122", text: "Has clearly recognisable, separate material been foreseen for each age group?", topic: "K" },
  { idx: 135, num: "123", text: "Is there any material being shared with other farms that enters the stables?", topic: "K" },
  { idx: 136, num: "124", text: "What measures do you take before this shared material enters your stable?", topic: "K" },
];

// Score column indices in Excel
const SCORE_COLS = {
  11: "8. Question score",
  36: "32. Question score",
  58: "53. Question score",
  73: "67. Question score",
  84: "77. Question score",
  101: "93. Question score",
  119: "109. Question score",
  126: "115.1 Question score",
  137: "124. Question score",
};

// Topic score column indices
const TOPIC_INDICES = {
  138: "A. Purchase and reproduction",
  139: "B. Transport and carcass removal",
  140: "C. Feed and water",
  141: "D. Visitors and farmworkers",
  142: "E. Vermin control and other animals",
  143: "F. Health management",
  144: "G. Calving management",
  145: "H. Calf management",
  146: "I. Dairy management",
  147: "J. Adult cattle management",
  148: "K. Working organisation and equipment",
};

const AGGREGATE_INDICES = {
  149: "External score",
  150: "Internal score",
  151: "Total score",
};

const TOPICS = {
  A: "Purchase and reproduction",
  B: "Transport and carcass removal",
  C: "Feed and water",
  D: "Visitors and farmworkers",
  E: "Vermin control and other animals",
  F: "Health management",
  G: "Calving management",
  H: "Calf management",
  I: "Dairy management",
  J: "Adult cattle management",
  K: "Working organisation and equipment",
};

// IBT indices (cols 139-149 in Excel = topics B through External score)
const IBT_INDICES = [138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148];

// ─── STYLES ─────────────────────────────────────────────────────────────────
const theme = {
  bg: "#0a0e17",
  surface: "#111827",
  surfaceAlt: "#1a2235",
  border: "#1e293b",
  borderLight: "#334155",
  accent: "#10b981",
  accentDark: "#059669",
  accentGlow: "rgba(16, 185, 129, 0.15)",
  danger: "#ef4444",
  warning: "#f59e0b",
  info: "#3b82f6",
  text: "#e2e8f0",
  textMuted: "#94a3b8",
  textDim: "#64748b",
  fc1: "#ef4444",
  fc2: "#f59e0b",
  fc3: "#10b981",
};

// ─── HELPER FUNCTIONS ───────────────────────────────────────────────────────
function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function percentile(arr, val) {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  let count = 0;
  for (const v of sorted) {
    if (v < val) count++;
    else if (v === val) count += 0.5;
  }
  return (count / sorted.length) * 100;
}

function classifyByPercentile(pct) {
  if (pct <= 25) return "FC1";
  if (pct <= 75) return "FC2";
  return "FC3";
}

function runClassification(farms, sbmCriteria, selectedIBT) {
  if (farms.length === 0) return [];

  // Only use selected IBT topics for Scale 2
  const activeIBT = selectedIBT && selectedIBT.length > 0 ? selectedIBT : IBT_INDICES;

  // Extract score arrays for percentile computation
  const totalScores = farms.map((f) => f.scores?.[151] ?? 0);
  const externalScores = farms.map((f) => f.scores?.[149] ?? 0);
  const internalScores = farms.map((f) => f.scores?.[150] ?? 0);

  // IBT scores arrays (only selected topics)
  const ibtArrays = {};
  for (const idx of activeIBT) {
    ibtArrays[idx] = farms.map((f) => f.scores?.[idx] ?? 0);
  }

  return farms.map((farm) => {
    const totalPct = percentile(totalScores, farm.scores?.[151] ?? 0);
    const extPct = percentile(externalScores, farm.scores?.[149] ?? 0);
    const intPct = percentile(internalScores, farm.scores?.[150] ?? 0);

    // Scale 1: classification based on total, external, internal percentiles
    const scale1Class = classifyByPercentile(Math.min(totalPct, extPct, intPct));

    // Scale 2: classification based on selected IBT percentiles only
    const ibtPcts = activeIBT.map((idx) =>
      percentile(ibtArrays[idx], farm.scores?.[idx] ?? 0)
    );
    const minIbtPct = ibtPcts.length > 0 ? Math.min(...ibtPcts) : 100;
    const scale2Class = classifyByPercentile(minIbtPct);

    // Scale 3: SBM check using regex
    let scale3Pass = true;
    const sbmDetails = [];
    for (const crit of sbmCriteria) {
      const farmAnswer = String(farm.answers?.[crit.questionIdx] ?? "").trim();
      let pass = false;
      let regexError = null;
      try {
        const re = new RegExp(crit.regexPattern, "i");
        pass = re.test(farmAnswer);
      } catch (e) {
        regexError = e.message;
        pass = false;
      }
      if (!pass) scale3Pass = false;
      sbmDetails.push({
        questionIdx: crit.questionIdx,
        questionNum: crit.questionNum,
        questionText: crit.questionText,
        regexPattern: crit.regexPattern,
        actual: farm.answers?.[crit.questionIdx] ?? "(no answer)",
        pass,
        regexError,
      });
    }

    // Final classification
    // Worst of scale1 and scale2, then check scale3
    const scaleClasses = [scale1Class, scale2Class];
    let combined = "FC3";
    if (scaleClasses.includes("FC1")) combined = "FC1";
    else if (scaleClasses.includes("FC2")) combined = "FC2";

    // If scale3 fails, downgrade to FC1
    if (!scale3Pass && sbmCriteria.length > 0) combined = "FC1";

    return {
      farmId: farm.id,
      farmName: farm.name,
      scale1: { classification: scale1Class, totalPct, extPct, intPct },
      scale2: { classification: scale2Class, ibtPcts, minIbtPct, activeIBT },
      scale3: { pass: scale3Pass, details: sbmDetails },
      finalClassification: combined,
      scores: farm.scores,
    };
  });
}

// ─── COMPONENTS ─────────────────────────────────────────────────────────────
function Badge({ type, children }) {
  const colors = {
    FC1: { bg: "rgba(239,68,68,0.15)", color: theme.fc1, border: "rgba(239,68,68,0.3)" },
    FC2: { bg: "rgba(245,158,11,0.15)", color: theme.fc2, border: "rgba(245,158,11,0.3)" },
    FC3: { bg: "rgba(16,185,129,0.15)", color: theme.fc3, border: "rgba(16,185,129,0.3)" },
    info: { bg: "rgba(59,130,246,0.12)", color: theme.info, border: "rgba(59,130,246,0.3)" },
    dim: { bg: "rgba(100,116,139,0.12)", color: theme.textDim, border: "rgba(100,116,139,0.3)" },
  };
  const c = colors[type] || colors.dim;
  return (
    <span
      style={{
        display: "inline-block",
        padding: "2px 10px",
        borderRadius: 4,
        fontSize: 12,
        fontWeight: 700,
        letterSpacing: "0.05em",
        background: c.bg,
        color: c.color,
        border: `1px solid ${c.border}`,
      }}
    >
      {children}
    </span>
  );
}

function Button({ onClick, variant = "primary", disabled, children, style: sx }) {
  const base = {
    padding: "8px 18px",
    borderRadius: 6,
    border: "none",
    cursor: disabled ? "not-allowed" : "pointer",
    fontWeight: 600,
    fontSize: 13,
    transition: "all 0.2s",
    opacity: disabled ? 0.5 : 1,
    fontFamily: "inherit",
    ...sx,
  };
  const variants = {
    primary: { background: theme.accent, color: "#000" },
    secondary: { background: theme.surfaceAlt, color: theme.text, border: `1px solid ${theme.border}` },
    danger: { background: "rgba(239,68,68,0.15)", color: theme.danger, border: `1px solid rgba(239,68,68,0.3)` },
    ghost: { background: "transparent", color: theme.textMuted },
  };
  return (
    <button onClick={disabled ? undefined : onClick} style={{ ...base, ...variants[variant] }}>
      {children}
    </button>
  );
}

function Card({ children, style: sx, onClick }) {
  return (
    <div
      onClick={onClick}
      style={{
        background: theme.surface,
        border: `1px solid ${theme.border}`,
        borderRadius: 10,
        padding: 20,
        cursor: onClick ? "pointer" : "default",
        transition: "border-color 0.2s, box-shadow 0.2s",
        ...sx,
      }}
      onMouseEnter={(e) => {
        if (onClick) {
          e.currentTarget.style.borderColor = theme.accent;
          e.currentTarget.style.boxShadow = `0 0 20px ${theme.accentGlow}`;
        }
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.borderColor = theme.border;
        e.currentTarget.style.boxShadow = "none";
      }}
    >
      {children}
    </div>
  );
}

function Input({ label, value, onChange, type = "text", placeholder, style: sx }) {
  return (
    <div style={{ marginBottom: 14 }}>
      {label && (
        <label style={{ display: "block", fontSize: 12, color: theme.textMuted, marginBottom: 5, fontWeight: 600 }}>
          {label}
        </label>
      )}
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        style={{
          width: "100%",
          padding: "9px 12px",
          background: theme.bg,
          border: `1px solid ${theme.border}`,
          borderRadius: 6,
          color: theme.text,
          fontSize: 13,
          fontFamily: "inherit",
          outline: "none",
          boxSizing: "border-box",
          ...sx,
        }}
      />
    </div>
  );
}

// ─── DASHBOARD VIEW ─────────────────────────────────────────────────────────
function Dashboard({ processes, onSelect, onNew }) {
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28 }}>
        <div>
          <h2 style={{ margin: 0, color: theme.text, fontSize: 22 }}>Classification Processes</h2>
          <p style={{ margin: "4px 0 0", color: theme.textDim, fontSize: 13 }}>
            Manage farm biosecurity classification workflows
          </p>
        </div>
        <Button onClick={onNew}>+ New Process</Button>
      </div>

      {processes.length === 0 ? (
        <Card style={{ textAlign: "center", padding: 50 }}>
          <div style={{ fontSize: 40, marginBottom: 12, opacity: 0.3 }}>🏗️</div>
          <p style={{ color: theme.textMuted, margin: 0, fontSize: 14 }}>
            No classification processes yet. Create one to get started.
          </p>
        </Card>
      ) : (
        <div style={{ display: "grid", gap: 12 }}>
          {processes.map((p) => (
            <Card key={p.id} onClick={() => onSelect(p.id)}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: theme.text }}>{p.name}</div>
                  <div style={{ fontSize: 12, color: theme.textDim, marginTop: 3 }}>
                    ID: {p.id} · Created: {new Date(p.created).toLocaleDateString()}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <Badge type="info">{p.farms?.length || 0} farms</Badge>
                  {p.classified && <Badge type="FC3">Classified</Badge>}
                </div>
              </div>
              {p.description && (
                <div style={{ fontSize: 13, color: theme.textMuted, marginTop: 8 }}>{p.description}</div>
              )}
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── NEW PROCESS FORM ───────────────────────────────────────────────────────
function NewProcessForm({ onSave, onCancel }) {
  const [name, setName] = useState("");
  const [desc, setDesc] = useState("");
  return (
    <div>
      <h2 style={{ color: theme.text, marginBottom: 20 }}>Create New Classification Process</h2>
      <Card>
        <Input label="Process Name" value={name} onChange={setName} placeholder="e.g., Spring 2025 Assessment" />
        <Input label="Description (optional)" value={desc} onChange={setDesc} placeholder="Brief description of this classification round" />
        <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
          <Button onClick={() => onSave({ name, description: desc })} disabled={!name.trim()}>
            Create Process
          </Button>
          <Button variant="secondary" onClick={onCancel}>
            Cancel
          </Button>
        </div>
      </Card>
    </div>
  );
}

// ─── SBM CONFIGURATION ─────────────────────────────────────────────────────
// ─── IBT TOPIC SELECTION (Scale 2) ──────────────────────────────────────────
function IBTConfig({ selectedIBT, onChange }) {
  const toggle = (idx) => {
    if (selectedIBT.includes(idx)) {
      onChange(selectedIBT.filter((i) => i !== idx));
    } else {
      onChange([...selectedIBT, idx].sort((a, b) => a - b));
    }
  };

  const allSelected = IBT_INDICES.every((idx) => selectedIBT.includes(idx));

  const toggleAll = () => {
    onChange(allSelected ? [] : [...IBT_INDICES]);
  };

  const checkboxStyle = (checked) => ({
    width: 16, height: 16, borderRadius: 4, cursor: "pointer",
    border: `1.5px solid ${checked ? theme.accent : theme.borderLight}`,
    background: checked ? theme.accent : "transparent",
    display: "flex", alignItems: "center", justifyContent: "center",
    flexShrink: 0, transition: "all 0.15s",
  });

  return (
    <Card style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: theme.text }}>
            Scale 2 — Important Biosecurity Topics (IBT)
          </div>
          <div style={{ fontSize: 12, color: theme.textDim, marginTop: 2 }}>
            Select which topics are examined in Scale 2 percentile check
          </div>
        </div>
        <button
          onClick={toggleAll}
          style={{
            background: "none", border: `1px solid ${theme.border}`, borderRadius: 5,
            color: theme.textMuted, cursor: "pointer", fontSize: 11,
            fontWeight: 600, fontFamily: "inherit", padding: "4px 10px",
          }}
        >
          {allSelected ? "Deselect all" : "Select all"}
        </button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
        {IBT_INDICES.map((idx) => {
          const label = TOPIC_INDICES[idx] || `Col ${idx}`;
          const checked = selectedIBT.includes(idx);
          return (
            <div
              key={idx}
              onClick={() => toggle(idx)}
              style={{
                display: "flex", alignItems: "center", gap: 8, padding: "7px 10px",
                background: checked ? theme.accentGlow : theme.bg,
                border: `1px solid ${checked ? "rgba(16,185,129,0.25)" : theme.border}`,
                borderRadius: 6, cursor: "pointer", transition: "all 0.15s",
              }}
            >
              <div style={checkboxStyle(checked)}>
                {checked && <span style={{ color: "#000", fontSize: 11, fontWeight: 800, lineHeight: 1 }}>✓</span>}
              </div>
              <span style={{ fontSize: 12, color: checked ? theme.text : theme.textMuted }}>{label}</span>
            </div>
          );
        })}
      </div>

      {selectedIBT.length === 0 && (
        <div style={{
          marginTop: 10, fontSize: 12, color: theme.warning, fontStyle: "italic",
          padding: "6px 10px", background: "rgba(245,158,11,0.08)", borderRadius: 5,
        }}>
          ⚠ No topics selected — Scale 2 will be skipped (all farms pass).
        </div>
      )}
      {selectedIBT.length > 0 && selectedIBT.length < IBT_INDICES.length && (
        <div style={{ marginTop: 10, fontSize: 11, color: theme.textDim }}>
          {selectedIBT.length} of {IBT_INDICES.length} topics selected
        </div>
      )}
    </Card>
  );
}

// ─── SBM CONFIGURATION (Scale 3) ───────────────────────────────────────────
function SBMConfig({ sbmCriteria, onChange }) {
  const [showAdd, setShowAdd] = useState(false);
  const [editIdx, setEditIdx] = useState(null);
  const [selQ, setSelQ] = useState("");
  const [regexPat, setRegexPat] = useState("");
  const [testVal, setTestVal] = useState("");

  const regexTestResult = useMemo(() => {
    if (!regexPat || !testVal) return null;
    try {
      const re = new RegExp(regexPat, "i");
      return re.test(testVal);
    } catch {
      return "error";
    }
  }, [regexPat, testVal]);

  const resetForm = () => {
    setSelQ("");
    setRegexPat("");
    setTestVal("");
    setShowAdd(false);
    setEditIdx(null);
  };

  const startEdit = (i) => {
    const c = sbmCriteria[i];
    setEditIdx(i);
    setSelQ(String(c.questionIdx));
    setRegexPat(c.regexPattern);
    setTestVal("");
    setShowAdd(false);
  };

  const startAdd = () => {
    resetForm();
    setShowAdd(true);
  };

  const saveCriterion = () => {
    const q = QUESTIONS.find((q) => String(q.idx) === selQ);
    if (!q || !regexPat.trim()) return;
    try { new RegExp(regexPat, "i"); } catch { return; }

    const entry = {
      questionIdx: q.idx,
      questionNum: q.num,
      questionText: q.text,
      regexPattern: regexPat.trim(),
    };

    if (editIdx !== null) {
      const updated = [...sbmCriteria];
      updated[editIdx] = entry;
      onChange(updated);
    } else {
      onChange([...sbmCriteria, entry]);
    }
    resetForm();
  };

  const remove = (i) => {
    if (editIdx === i) resetForm();
    onChange(sbmCriteria.filter((_, j) => j !== i));
  };

  const showForm = showAdd || editIdx !== null;

  const formUI = (
    <div style={{
      background: theme.bg, borderRadius: 8, padding: 14, marginBottom: 12,
      border: `1px solid ${editIdx !== null ? "rgba(59,130,246,0.3)" : theme.border}`,
    }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: editIdx !== null ? theme.info : theme.accent, marginBottom: 10 }}>
        {editIdx !== null ? `Editing criterion #${editIdx + 1}` : "New criterion"}
      </div>

      <div style={{ marginBottom: 10 }}>
        <label style={{ display: "block", fontSize: 12, color: theme.textMuted, marginBottom: 4, fontWeight: 600 }}>
          Select Question
        </label>
        <select
          value={selQ}
          onChange={(e) => setSelQ(e.target.value)}
          style={{
            width: "100%", padding: "8px 10px", background: theme.surface,
            border: `1px solid ${theme.border}`, borderRadius: 6,
            color: theme.text, fontSize: 12, fontFamily: "inherit",
          }}
        >
          <option value="">— Choose question —</option>
          {QUESTIONS.map((q) => (
            <option key={q.idx} value={q.idx}>
              Q{q.num}. {q.text.slice(0, 80)}
            </option>
          ))}
        </select>
      </div>

      <div style={{ marginBottom: 10 }}>
        <label style={{ display: "block", fontSize: 12, color: theme.textMuted, marginBottom: 4, fontWeight: 600 }}>
          Answer validation pattern (RegEx, case-insensitive)
        </label>
        <input
          value={regexPat}
          onChange={(e) => setRegexPat(e.target.value)}
          placeholder="e.g., ^Always$  or  ^(Yes|Always)  or  (?=.*cleaned)(?=.*disinfected)"
          style={{
            width: "100%", padding: "7px 10px", background: theme.surface,
            border: `1px solid ${theme.border}`, borderRadius: 5, color: theme.text,
            fontSize: 12, fontFamily: "'DM Mono', 'Fira Code', monospace",
            outline: "none", boxSizing: "border-box",
          }}
        />
        <div style={{ fontSize: 10, color: theme.textDim, marginTop: 3 }}>
          Examples: <code style={{ color: theme.accent }}>^Always$</code> exact match ·{" "}
          <code style={{ color: theme.accent }}>^(Yes|Always)</code> starts with Yes or Always ·{" "}
          <code style={{ color: theme.accent }}>cleaned.*disinfected</code> contains both words in order
        </div>
      </div>

      <div style={{ marginBottom: 10 }}>
        <label style={{ display: "block", fontSize: 12, color: theme.textMuted, marginBottom: 4, fontWeight: 600 }}>
          Test value (optional — test your regex against a sample answer)
        </label>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <input
            value={testVal}
            onChange={(e) => setTestVal(e.target.value)}
            placeholder="Paste a sample answer to test..."
            style={{
              flex: 1, padding: "7px 10px", background: theme.surface,
              border: `1px solid ${theme.border}`, borderRadius: 5, color: theme.text,
              fontSize: 12, fontFamily: "inherit", outline: "none", boxSizing: "border-box",
            }}
          />
          {regexTestResult !== null && (
            <span style={{
              fontSize: 12, fontWeight: 700, padding: "4px 10px", borderRadius: 4,
              background: regexTestResult === "error" ? "rgba(239,68,68,0.12)"
                : regexTestResult ? "rgba(16,185,129,0.12)" : "rgba(239,68,68,0.12)",
              color: regexTestResult === "error" ? theme.danger
                : regexTestResult ? theme.fc3 : theme.fc1,
            }}>
              {regexTestResult === "error" ? "Invalid regex" : regexTestResult ? "MATCH" : "NO MATCH"}
            </span>
          )}
        </div>
      </div>

      <div style={{ display: "flex", gap: 8 }}>
        <Button
          onClick={saveCriterion}
          disabled={!selQ || !regexPat.trim()}
          style={{ fontSize: 12, padding: "5px 12px" }}
        >
          {editIdx !== null ? "Save Changes" : "Add Criterion"}
        </Button>
        <Button variant="secondary" onClick={resetForm} style={{ fontSize: 12, padding: "5px 12px" }}>
          Cancel
        </Button>
      </div>
    </div>
  );

  return (
    <Card style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: theme.text }}>
            Scale 3 — Specific Biosecurity Measures (SBM)
          </div>
          <div style={{ fontSize: 12, color: theme.textDim, marginTop: 2 }}>
            Define minimum criteria: questions and RegEx patterns for valid answers
          </div>
        </div>
        {!showForm && (
          <Button variant="secondary" onClick={startAdd} style={{ fontSize: 12, padding: "5px 12px" }}>
            + Add Criterion
          </Button>
        )}
      </div>

      {showForm && formUI}

      {sbmCriteria.length === 0 && !showForm ? (
        <div style={{ fontSize: 12, color: theme.textDim, fontStyle: "italic" }}>
          No SBM criteria defined yet. Farms won't be checked against Scale 3.
        </div>
      ) : (
        <div style={{ display: "grid", gap: 6 }}>
          {sbmCriteria.map((c, i) => {
            const isEditing = editIdx === i;
            return (
              <div
                key={i}
                style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  padding: "8px 12px", background: isEditing ? "rgba(59,130,246,0.08)" : theme.bg,
                  borderRadius: 6, fontSize: 12,
                  border: `1px solid ${isEditing ? "rgba(59,130,246,0.3)" : theme.border}`,
                  cursor: "pointer", transition: "all 0.15s",
                }}
                onClick={() => { if (!showForm) startEdit(i); }}
                onMouseEnter={(e) => { if (!showForm) e.currentTarget.style.borderColor = theme.info; }}
                onMouseLeave={(e) => { if (!isEditing) e.currentTarget.style.borderColor = theme.border; }}
              >
                <div style={{ minWidth: 0, overflow: "hidden", display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ color: theme.textDim, fontSize: 10, flexShrink: 0 }}>#{i + 1}</span>
                  <span style={{ color: theme.accent, fontWeight: 700 }}>Q{c.questionNum}</span>
                  <span style={{ color: theme.textMuted }}>→</span>
                  <span style={{ color: theme.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.questionText.slice(0, 45)}...</span>
                  <span style={{ color: theme.textMuted }}>match</span>
                  <code style={{
                    color: theme.warning, fontWeight: 700, fontSize: 11, flexShrink: 0,
                    background: "rgba(245,158,11,0.08)", padding: "1px 5px", borderRadius: 3,
                  }}>
                    /{c.regexPattern}/i
                  </code>
                </div>
                <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                  <button
                    onClick={(e) => { e.stopPropagation(); startEdit(i); }}
                    style={{
                      background: "none", border: "none", color: theme.info,
                      cursor: "pointer", fontSize: 12, padding: "2px 5px",
                    }}
                    title="Edit"
                  >
                    ✎
                  </button>
                  <button
                    onClick={(e) => { e.stopPropagation(); remove(i); }}
                    style={{
                      background: "none", border: "none", color: theme.danger,
                      cursor: "pointer", fontSize: 14, padding: "2px 5px",
                    }}
                    title="Delete"
                  >
                    ✕
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </Card>
  );
}

// ─── EXCEL IMPORT ───────────────────────────────────────────────────────────
function ExcelImport({ onImport }) {
  const fileRef = useRef();
  const [status, setStatus] = useState(null);

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setStatus("Reading file...");
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      if (rows.length < 2) {
        setStatus("Error: File must have at least 2 rows (questions + 1 farm)");
        return;
      }

      // Row 0 = questions, rows 1+ = farm data
      const farms = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const answers = {};
        const scores = {};

        // Map all columns
        for (let j = 0; j < row.length; j++) {
          if (row[j] !== null && row[j] !== undefined && row[j] !== "") {
            if (j >= 138 && j <= 151) {
              scores[j] = parseFloat(row[j]) || 0;
            } else if (SCORE_COLS[j] !== undefined) {
              scores[j] = parseFloat(row[j]) || 0;
            } else {
              answers[j] = row[j];
            }
          }
        }

        farms.push({
          id: genId(),
          name: `Farm ${i}`,
          answers,
          scores,
          importedAt: new Date().toISOString(),
        });
      }

      setStatus(`Successfully parsed ${farms.length} farm(s)`);
      onImport(farms);
    } catch (err) {
      setStatus(`Error: ${err.message}`);
    }
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <Card style={{ marginBottom: 16 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: theme.text, marginBottom: 4 }}>
        Import Farm Data from Excel
      </div>
      <div style={{ fontSize: 12, color: theme.textDim, marginBottom: 12 }}>
        Upload an Excel file matching the Gent questionnaire format (first row = questions, subsequent rows = farm answers)
      </div>
      <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFile}
          style={{ fontSize: 12, color: theme.textMuted }}
        />
      </div>
      {status && (
        <div
          style={{
            marginTop: 10,
            fontSize: 12,
            color: status.startsWith("Error") ? theme.danger : theme.accent,
            padding: "6px 10px",
            background: status.startsWith("Error") ? "rgba(239,68,68,0.1)" : "rgba(16,185,129,0.1)",
            borderRadius: 6,
          }}
        >
          {status}
        </div>
      )}
    </Card>
  );
}

// ─── MANUAL ENTRY FORM ──────────────────────────────────────────────────────
// Score fields needed for classification, grouped for the UI
const SCORE_FIELDS = [
  { group: "Topic scores (Scale 2 — IBT)", fields: [
    { idx: 138, label: "A. Purchase and reproduction" },
    { idx: 139, label: "B. Transport and carcass removal" },
    { idx: 140, label: "C. Feed and water" },
    { idx: 141, label: "D. Visitors and farmworkers" },
    { idx: 142, label: "E. Vermin control and other animals" },
    { idx: 143, label: "F. Health management" },
    { idx: 144, label: "G. Calving management" },
    { idx: 145, label: "H. Calf management" },
    { idx: 146, label: "I. Dairy management" },
    { idx: 147, label: "J. Adult cattle management" },
    { idx: 148, label: "K. Working organisation and equipment" },
  ]},
  { group: "Aggregate scores (Scale 1)", fields: [
    { idx: 149, label: "External score" },
    { idx: 150, label: "Internal score" },
    { idx: 151, label: "Total score" },
  ]},
  { group: "Question scores", fields: [
    { idx: 11,  label: "Q8 score" },
    { idx: 36,  label: "Q32 score" },
    { idx: 58,  label: "Q53 score" },
    { idx: 73,  label: "Q67 score" },
    { idx: 84,  label: "Q77 score" },
    { idx: 101, label: "Q93 score" },
    { idx: 119, label: "Q109 score" },
    { idx: 126, label: "Q115.1 score" },
    { idx: 137, label: "Q124 score" },
  ]},
];

function ManualEntry({ onSave, onCancel, existingFarm }) {
  const isEdit = !!existingFarm;
  const [farmName, setFarmName] = useState(existingFarm?.name || "");
  const [answers, setAnswers] = useState(existingFarm?.answers || {});
  const [scores, setScores] = useState(existingFarm?.scores || {});
  const [activeTab, setActiveTab] = useState("A"); // A-K for questions, "scores" for scores

  const topicQuestions = activeTab !== "scores" ? QUESTIONS.filter((q) => q.topic === activeTab) : [];

  const setAnswer = (idx, val) => {
    setAnswers((prev) => ({ ...prev, [idx]: val }));
  };

  const setScore = (idx, val) => {
    const num = val === "" ? undefined : parseFloat(val);
    setScores((prev) => {
      const next = { ...prev };
      if (num === undefined || isNaN(num)) delete next[idx];
      else next[idx] = num;
      return next;
    });
  };

  const handleSave = () => {
    if (!farmName.trim()) return;
    onSave({
      id: existingFarm?.id || genId(),
      name: farmName.trim(),
      answers,
      scores,
      importedAt: existingFarm?.importedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  };

  const answeredCount = Object.keys(answers).length;
  const scoredCount = Object.keys(scores).length;
  const requiredScoreCount = SCORE_FIELDS.reduce((n, g) => n + g.fields.length, 0);

  // Check if all classification-critical scores are filled (Scale 1 + Scale 2)
  const criticalIndices = [138,139,140,141,142,143,144,145,146,147,148,149,150,151];
  const criticalFilled = criticalIndices.filter((i) => scores[i] !== undefined).length;

  const inputStyle = {
    width: "100%",
    padding: "7px 10px",
    background: theme.bg,
    border: `1px solid ${theme.border}`,
    borderRadius: 5,
    color: theme.text,
    fontSize: 12,
    fontFamily: "inherit",
    outline: "none",
    boxSizing: "border-box",
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h3 style={{ color: theme.text, margin: 0 }}>{isEdit ? `Edit Farm — ${existingFarm.name}` : "Manual Farm Entry"}</h3>
        <div style={{ fontSize: 12, color: theme.textDim }}>
          {answeredCount} questions · {scoredCount} scores
        </div>
      </div>

      <Card style={{ marginBottom: 16 }}>
        <Input label="Farm Name" value={farmName} onChange={setFarmName} placeholder="Enter farm name or identifier" />
      </Card>

      {/* Validation hint */}
      {criticalFilled < 14 && (
        <div style={{
          padding: "8px 12px", marginBottom: 12, borderRadius: 6, fontSize: 12,
          background: "rgba(245,158,11,0.1)", border: "1px solid rgba(245,158,11,0.2)", color: theme.warning,
        }}>
          ⚠ Classification requires all 14 key scores (11 topic + External + Internal + Total).
          Currently {criticalFilled}/14 filled. Go to the <b>Scores</b> tab to enter them.
        </div>
      )}

      {/* Tabs: A-K topics + Scores */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 14 }}>
        {Object.entries(TOPICS).map(([key, label]) => {
          const count = QUESTIONS.filter((q) => q.topic === key && answers[q.idx] !== undefined).length;
          const total = QUESTIONS.filter((q) => q.topic === key).length;
          return (
            <button
              key={key}
              onClick={() => setActiveTab(key)}
              style={{
                padding: "6px 12px", borderRadius: 5,
                border: `1px solid ${activeTab === key ? theme.accent : theme.border}`,
                background: activeTab === key ? theme.accentGlow : "transparent",
                color: activeTab === key ? theme.accent : theme.textMuted,
                cursor: "pointer", fontSize: 11, fontWeight: 600, fontFamily: "inherit",
              }}
            >
              {key}. {label.slice(0, 14)}
              {count > 0 && <span style={{ marginLeft: 4, opacity: 0.7 }}>{count}/{total}</span>}
            </button>
          );
        })}

        {/* Scores tab */}
        <button
          onClick={() => setActiveTab("scores")}
          style={{
            padding: "6px 14px", borderRadius: 5,
            border: `1px solid ${activeTab === "scores" ? theme.info : theme.border}`,
            background: activeTab === "scores" ? "rgba(59,130,246,0.12)" : "transparent",
            color: activeTab === "scores" ? theme.info : theme.textMuted,
            cursor: "pointer", fontSize: 11, fontWeight: 700, fontFamily: "inherit",
          }}
        >
          ★ Scores
          {scoredCount > 0 && <span style={{ marginLeft: 4, opacity: 0.7 }}>{scoredCount}/{requiredScoreCount}</span>}
        </button>
      </div>

      {/* ─── SCORES TAB ─── */}
      {activeTab === "scores" && (
        <div style={{ display: "grid", gap: 14 }}>
          {SCORE_FIELDS.map((group) => (
            <Card key={group.group}>
              <div style={{ fontSize: 13, fontWeight: 700, color: theme.info, marginBottom: 12 }}>
                {group.group}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                {group.fields.map((f) => (
                  <div key={f.idx}>
                    <label style={{ display: "block", fontSize: 11, color: theme.textMuted, marginBottom: 3, fontWeight: 600 }}>
                      {f.label}
                      {criticalIndices.includes(f.idx) && (
                        <span style={{ color: theme.danger, marginLeft: 3 }}>*</span>
                      )}
                    </label>
                    <input
                      type="number"
                      min="0" max="100"
                      value={scores[f.idx] !== undefined ? scores[f.idx] : ""}
                      onChange={(e) => setScore(f.idx, e.target.value)}
                      placeholder="0–100"
                      style={inputStyle}
                    />
                  </div>
                ))}
              </div>
            </Card>
          ))}
          <div style={{ fontSize: 11, color: theme.textDim, fontStyle: "italic" }}>
            Fields marked <span style={{ color: theme.danger }}>*</span> are required for classification (Scale 1 and Scale 2).
          </div>
        </div>
      )}

      {/* ─── QUESTIONS TAB ─── */}
      {activeTab !== "scores" && (
        <Card>
          <div style={{ fontSize: 13, fontWeight: 700, color: theme.accent, marginBottom: 14 }}>
            {activeTab}. {TOPICS[activeTab]}
          </div>
          <div style={{ display: "grid", gap: 14 }}>
            {topicQuestions.map((q) => (
              <div key={q.idx} style={{ borderBottom: `1px solid ${theme.border}`, paddingBottom: 12 }}>
                <div style={{ fontSize: 12, color: theme.text, marginBottom: 6 }}>
                  <span style={{ color: theme.accent, fontWeight: 700 }}>Q{q.num}.</span> {q.text}
                </div>
                <input
                  value={answers[q.idx] || ""}
                  onChange={(e) => setAnswer(q.idx, e.target.value)}
                  placeholder="Enter answer..."
                  style={inputStyle}
                />
              </div>
            ))}
          </div>
        </Card>
      )}

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <Button onClick={handleSave} disabled={!farmName.trim()}>
          {isEdit ? "Update Farm" : "Save Farm"}
        </Button>
        <Button variant="secondary" onClick={onCancel}>
          Cancel
        </Button>
      </div>
    </div>
  );
}

// ─── CLASSIFICATION RESULTS ─────────────────────────────────────────────────
function ClassificationResults({ results, processName }) {
  const [expandedFarm, setExpandedFarm] = useState(null);

  const summary = { FC1: 0, FC2: 0, FC3: 0 };
  results.forEach((r) => summary[r.finalClassification]++);

  return (
    <div>
      <h3 style={{ color: theme.text, marginBottom: 16 }}>Classification Results — {processName}</h3>

      {/* Summary bar */}
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 20, justifyContent: "center" }}>
          {[
            { label: "FC1 — Poor", count: summary.FC1, color: theme.fc1, type: "FC1" },
            { label: "FC2 — Medium", count: summary.FC2, color: theme.fc2, type: "FC2" },
            { label: "FC3 — Good", count: summary.FC3, color: theme.fc3, type: "FC3" },
          ].map((item) => (
            <div key={item.type} style={{ textAlign: "center" }}>
              <div style={{ fontSize: 28, fontWeight: 800, color: item.color }}>{item.count}</div>
              <div style={{ fontSize: 11, color: theme.textDim, marginTop: 2 }}>{item.label}</div>
            </div>
          ))}
        </div>

        {/* Distribution bar */}
        <div
          style={{
            display: "flex",
            height: 8,
            borderRadius: 4,
            overflow: "hidden",
            marginTop: 14,
            background: theme.bg,
          }}
        >
          {results.length > 0 && (
            <>
              <div style={{ width: `${(summary.FC1 / results.length) * 100}%`, background: theme.fc1 }} />
              <div style={{ width: `${(summary.FC2 / results.length) * 100}%`, background: theme.fc2 }} />
              <div style={{ width: `${(summary.FC3 / results.length) * 100}%`, background: theme.fc3 }} />
            </>
          )}
        </div>
      </Card>

      {/* Farm results */}
      <div style={{ display: "grid", gap: 8 }}>
        {results.map((r) => (
          <Card key={r.farmId}>
            <div
              onClick={() => setExpandedFarm(expandedFarm === r.farmId ? null : r.farmId)}
              style={{ cursor: "pointer" }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <Badge type={r.finalClassification}>{r.finalClassification}</Badge>
                  <span style={{ fontWeight: 600, color: theme.text, fontSize: 14 }}>{r.farmName}</span>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <span style={{ fontSize: 11, color: theme.textDim }}>
                    Total: {r.scores?.[151] ?? "—"} · Ext: {r.scores?.[149] ?? "—"} · Int: {r.scores?.[150] ?? "—"}
                  </span>
                  <span style={{ color: theme.textDim, fontSize: 14 }}>
                    {expandedFarm === r.farmId ? "▼" : "▶"}
                  </span>
                </div>
              </div>
            </div>

            {expandedFarm === r.farmId && (
              <div style={{ marginTop: 14, borderTop: `1px solid ${theme.border}`, paddingTop: 14 }}>
                {/* Scale 1 */}
                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 6 }}>
                    Scale 1 — Score Distribution
                    <Badge type={r.scale1.classification} style={{ marginLeft: 8 }}>
                      {r.scale1.classification}
                    </Badge>
                  </div>
                  <div style={{ display: "flex", gap: 16, fontSize: 11, color: theme.textMuted }}>
                    <span>
                      Total percentile: <b style={{ color: theme.text }}>{r.scale1.totalPct.toFixed(1)}%</b>
                    </span>
                    <span>
                      External percentile: <b style={{ color: theme.text }}>{r.scale1.extPct.toFixed(1)}%</b>
                    </span>
                    <span>
                      Internal percentile: <b style={{ color: theme.text }}>{r.scale1.intPct.toFixed(1)}%</b>
                    </span>
                  </div>
                </div>

                {/* Scale 2 */}
                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 6 }}>
                    Scale 2 — Important Biosecurity Topics
                    <Badge type={r.scale2.classification} style={{ marginLeft: 8 }}>
                      {r.scale2.classification}
                    </Badge>
                  </div>
                  <div style={{ fontSize: 11, color: theme.textMuted, marginBottom: 4 }}>
                    Min IBT percentile: <b style={{ color: theme.text }}>{r.scale2.minIbtPct.toFixed(1)}%</b>
                    {r.scale2.activeIBT && r.scale2.activeIBT.length < IBT_INDICES.length && (
                      <span style={{ marginLeft: 8 }}>({r.scale2.activeIBT.length} of {IBT_INDICES.length} topics selected)</span>
                    )}
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                    {(r.scale2.activeIBT || IBT_INDICES).map((idx, i) => {
                      const label = TOPIC_INDICES[idx]?.split(". ")[0] || idx;
                      const pct = r.scale2.ibtPcts[i];
                      return (
                        <span
                          key={idx}
                          style={{
                            padding: "2px 8px",
                            borderRadius: 4,
                            fontSize: 10,
                            background: pct <= 25 ? "rgba(239,68,68,0.1)" : pct <= 75 ? "rgba(245,158,11,0.1)" : "rgba(16,185,129,0.1)",
                            color: pct <= 25 ? theme.fc1 : pct <= 75 ? theme.fc2 : theme.fc3,
                            border: `1px solid ${pct <= 25 ? "rgba(239,68,68,0.2)" : pct <= 75 ? "rgba(245,158,11,0.2)" : "rgba(16,185,129,0.2)"}`,
                          }}
                        >
                          {label}: {pct.toFixed(0)}%
                        </span>
                      );
                    })}
                  </div>
                </div>

                {/* Scale 3 */}
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 6 }}>
                    Scale 3 — Specific Biosecurity Measures
                    {r.scale3.details.length > 0 ? (
                      <Badge type={r.scale3.pass ? "FC3" : "FC1"} style={{ marginLeft: 8 }}>
                        {r.scale3.pass ? "PASS" : "FAIL"}
                      </Badge>
                    ) : (
                      <Badge type="dim" style={{ marginLeft: 8 }}>
                        No criteria defined
                      </Badge>
                    )}
                  </div>
                  {r.scale3.details.length > 0 && (
                    <div style={{ display: "grid", gap: 4 }}>
                      {r.scale3.details.map((d, i) => (
                        <div
                          key={i}
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            padding: "4px 8px",
                            background: d.pass ? "rgba(16,185,129,0.05)" : "rgba(239,68,68,0.05)",
                            borderRadius: 4,
                            fontSize: 11,
                            border: `1px solid ${d.pass ? "rgba(16,185,129,0.15)" : "rgba(239,68,68,0.15)"}`,
                          }}
                        >
                          <span style={{ color: theme.textMuted, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" }}>
                            Q{d.questionNum}: pattern <code style={{ color: theme.warning, fontSize: 10 }}>/{d.regexPattern}/i</code> → got "{d.actual}"
                            {d.regexError && <span style={{ color: theme.danger, marginLeft: 4 }}>(regex error)</span>}
                          </span>
                          <span style={{ color: d.pass ? theme.fc3 : theme.fc1, fontWeight: 700, flexShrink: 0 }}>
                            {d.pass ? "✓" : "✕"}
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Topic scores */}
                <div style={{ marginTop: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 8 }}>
                    Topic Scores
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    {Object.entries(TOPIC_INDICES).map(([idx, label]) => {
                      const val = r.scores?.[parseInt(idx)] ?? "—";
                      return (
                        <div
                          key={idx}
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            padding: "4px 8px",
                            background: theme.bg,
                            borderRadius: 4,
                            fontSize: 11,
                          }}
                        >
                          <span style={{ color: theme.textMuted }}>{label}</span>
                          <span style={{ color: theme.text, fontWeight: 600 }}>{val}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}
          </Card>
        ))}
      </div>
    </div>
  );
}

// ─── FARM LIST WITH DETAIL VIEW ─────────────────────────────────────────────
function FarmList({ farms, onEdit, onRemove }) {
  const [expandedId, setExpandedId] = useState(null);

  const toggle = (id) => setExpandedId(expandedId === id ? null : id);

  return (
    <div>
      <div style={{ fontSize: 14, fontWeight: 700, color: theme.text, marginBottom: 10 }}>
        Farms ({farms.length})
      </div>
      <div style={{ display: "grid", gap: 6 }}>
        {farms.map((f) => {
          const expanded = expandedId === f.id;
          const ansCount = Object.keys(f.answers || {}).length;
          const scoreCount = Object.keys(f.scores || {}).length;
          return (
            <div
              key={f.id}
              style={{
                background: theme.surface,
                border: `1px solid ${expanded ? theme.accent : theme.border}`,
                borderRadius: 8,
                transition: "border-color 0.15s",
                overflow: "hidden",
              }}
            >
              {/* Header row — always visible */}
              <div
                onClick={() => toggle(f.id)}
                style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  padding: "10px 14px", cursor: "pointer",
                }}
                onMouseEnter={(e) => { if (!expanded) e.currentTarget.parentElement.style.borderColor = theme.accent; }}
                onMouseLeave={(e) => { if (!expanded) e.currentTarget.parentElement.style.borderColor = theme.border; }}
              >
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: theme.text }}>{f.name}</div>
                  <div style={{ fontSize: 11, color: theme.textDim }}>
                    {ansCount} answers · {scoreCount} scores ·{" "}
                    {f.scores?.[151] !== undefined ? `Total: ${f.scores[151]}` : "Scores pending"}
                    {f.scores?.[149] !== undefined && ` · Ext: ${f.scores[149]}`}
                    {f.scores?.[150] !== undefined && ` · Int: ${f.scores[150]}`}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                  <button
                    onClick={(e) => { e.stopPropagation(); onEdit(f.id); }}
                    title="Edit farm"
                    style={{
                      background: "rgba(59,130,246,0.1)", border: "1px solid rgba(59,130,246,0.2)",
                      color: theme.info, cursor: "pointer", fontSize: 11, fontWeight: 600,
                      padding: "3px 10px", borderRadius: 4, fontFamily: "inherit",
                    }}
                  >
                    Edit
                  </button>
                  <button
                    onClick={(e) => { e.stopPropagation(); onRemove(f.id); }}
                    title="Remove farm"
                    style={{
                      background: "none", border: "none", color: theme.danger,
                      cursor: "pointer", fontSize: 16, padding: "2px 6px",
                    }}
                  >
                    ✕
                  </button>
                  <span style={{ color: theme.textDim, fontSize: 13, marginLeft: 2 }}>
                    {expanded ? "▼" : "▶"}
                  </span>
                </div>
              </div>

              {/* Expanded detail */}
              {expanded && (
                <div style={{ padding: "0 14px 14px", borderTop: `1px solid ${theme.border}` }}>
                  {/* Topic scores */}
                  {scoreCount > 0 && (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 8 }}>
                        Scores
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
                        {Object.entries(TOPIC_INDICES).map(([idx, label]) => {
                          const val = f.scores?.[parseInt(idx)];
                          return val !== undefined ? (
                            <div key={idx} style={{
                              display: "flex", justifyContent: "space-between",
                              padding: "3px 8px", background: theme.bg, borderRadius: 4, fontSize: 11,
                            }}>
                              <span style={{ color: theme.textMuted }}>{label}</span>
                              <span style={{ color: theme.text, fontWeight: 600 }}>{val}</span>
                            </div>
                          ) : null;
                        })}
                        {Object.entries(AGGREGATE_INDICES).map(([idx, label]) => {
                          const val = f.scores?.[parseInt(idx)];
                          return val !== undefined ? (
                            <div key={idx} style={{
                              display: "flex", justifyContent: "space-between",
                              padding: "3px 8px", background: "rgba(16,185,129,0.05)",
                              border: "1px solid rgba(16,185,129,0.1)", borderRadius: 4, fontSize: 11,
                            }}>
                              <span style={{ color: theme.accent }}>{label}</span>
                              <span style={{ color: theme.text, fontWeight: 700 }}>{val}</span>
                            </div>
                          ) : null;
                        })}
                      </div>
                    </div>
                  )}

                  {/* Answers (grouped by topic) */}
                  {ansCount > 0 && (
                    <div style={{ marginTop: 14 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: theme.text, marginBottom: 8 }}>
                        Answers ({ansCount})
                      </div>
                      {Object.entries(TOPICS).map(([topicKey, topicLabel]) => {
                        const topicQs = QUESTIONS.filter(
                          (q) => q.topic === topicKey && f.answers?.[q.idx] !== undefined
                        );
                        if (topicQs.length === 0) return null;
                        return (
                          <div key={topicKey} style={{ marginBottom: 8 }}>
                            <div style={{ fontSize: 11, fontWeight: 700, color: theme.accent, marginBottom: 4 }}>
                              {topicKey}. {topicLabel}
                            </div>
                            {topicQs.map((q) => (
                              <div key={q.idx} style={{
                                display: "flex", gap: 8, padding: "3px 8px",
                                background: theme.bg, borderRadius: 4, marginBottom: 2, fontSize: 11,
                              }}>
                                <span style={{ color: theme.textDim, fontWeight: 600, flexShrink: 0 }}>Q{q.num}</span>
                                <span style={{ color: theme.textMuted, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                  {q.text.slice(0, 50)}...
                                </span>
                                <span style={{ color: theme.text, fontWeight: 500, flexShrink: 0, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                  {String(f.answers[q.idx])}
                                </span>
                              </div>
                            ))}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {ansCount === 0 && scoreCount === 0 && (
                    <div style={{ marginTop: 12, fontSize: 12, color: theme.textDim, fontStyle: "italic" }}>
                      No data entered yet. Click Edit to add answers and scores.
                    </div>
                  )}

                  <div style={{ marginTop: 12 }}>
                    <Button variant="secondary" onClick={() => onEdit(f.id)} style={{ fontSize: 11, padding: "4px 12px" }}>
                      Open in Editor
                    </Button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── PROCESS DETAIL VIEW ────────────────────────────────────────────────────
function ProcessDetail({ process, onUpdate, onBack }) {
  const [view, setView] = useState("overview"); // overview, manual, editFarm, results
  const [results, setResults] = useState(process.results || null);
  const [editFarmId, setEditFarmId] = useState(null);

  const updateProcess = (updates) => {
    onUpdate({ ...process, ...updates });
  };

  const handleImport = (farms) => {
    const existing = process.farms || [];
    updateProcess({ farms: [...existing, ...farms] });
  };

  const handleManualSave = (farm) => {
    const existing = process.farms || [];
    updateProcess({ farms: [...existing, farm] });
    setView("overview");
  };

  const handleFarmUpdate = (updatedFarm) => {
    const updated = (process.farms || []).map((f) => f.id === updatedFarm.id ? updatedFarm : f);
    updateProcess({ farms: updated });
    setView("overview");
    setEditFarmId(null);
  };

  const openEditFarm = (farmId) => {
    setEditFarmId(farmId);
    setView("editFarm");
  };

  const removeFarm = (farmId) => {
    updateProcess({ farms: (process.farms || []).filter((f) => f.id !== farmId) });
  };

  const handleClassify = () => {
    const classResults = runClassification(process.farms || [], process.sbmCriteria || [], process.selectedIBT || []);
    setResults(classResults);
    updateProcess({ results: classResults, classified: true });
    setView("results");
  };

  const handleSBMChange = (newCriteria) => {
    updateProcess({ sbmCriteria: newCriteria });
  };

  const handleIBTChange = (newSelection) => {
    updateProcess({ selectedIBT: newSelection });
  };

  const farms = process.farms || [];

  if (view === "manual") {
    return <ManualEntry onSave={handleManualSave} onCancel={() => setView("overview")} />;
  }

  if (view === "editFarm" && editFarmId) {
    const farmToEdit = (process.farms || []).find((f) => f.id === editFarmId);
    if (farmToEdit) {
      return (
        <ManualEntry
          existingFarm={farmToEdit}
          onSave={handleFarmUpdate}
          onCancel={() => { setView("overview"); setEditFarmId(null); }}
        />
      );
    }
  }

  if (view === "results" && results) {
    return (
      <div>
        <Button variant="ghost" onClick={() => setView("overview")} style={{ marginBottom: 14, padding: "4px 0" }}>
          ← Back to process
        </Button>
        <ClassificationResults results={results} processName={process.name} />
      </div>
    );
  }

  return (
    <div>
      <Button variant="ghost" onClick={onBack} style={{ marginBottom: 14, padding: "4px 0" }}>
        ← All Processes
      </Button>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
        <div>
          <h2 style={{ margin: 0, color: theme.text, fontSize: 20 }}>{process.name}</h2>
          <div style={{ fontSize: 12, color: theme.textDim, marginTop: 3 }}>
            ID: {process.id} · {farms.length} farm(s) loaded
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {results && (
            <Button variant="secondary" onClick={() => setView("results")} style={{ fontSize: 12, padding: "6px 14px" }}>
              View Results
            </Button>
          )}
          <Button onClick={handleClassify} disabled={farms.length < 2} style={{ fontSize: 12, padding: "6px 14px" }}>
            Run Classification
          </Button>
        </div>
      </div>

      {farms.length < 2 && (
        <div
          style={{
            padding: "10px 14px",
            background: "rgba(245,158,11,0.1)",
            border: "1px solid rgba(245,158,11,0.2)",
            borderRadius: 8,
            fontSize: 12,
            color: theme.warning,
            marginBottom: 16,
          }}
        >
          ⚠ At least 2 farms are needed to run percentile-based classification. Add more farm data.
        </div>
      )}

      {/* IBT Topic Selection (Scale 2) */}
      <IBTConfig selectedIBT={process.selectedIBT || []} onChange={handleIBTChange} />

      {/* SBM Configuration (Scale 3) */}
      <SBMConfig sbmCriteria={process.sbmCriteria || []} onChange={handleSBMChange} />

      {/* Import section */}
      <ExcelImport onImport={handleImport} />

      {/* Manual entry button */}
      <Card
        onClick={() => setView("manual")}
        style={{ marginBottom: 16, textAlign: "center", padding: 16, cursor: "pointer" }}
      >
        <div style={{ fontSize: 14, fontWeight: 600, color: theme.accent }}>+ Add Farm Manually</div>
        <div style={{ fontSize: 12, color: theme.textDim, marginTop: 2 }}>
          Answer all biosecurity questions for a single farm
        </div>
      </Card>

      {/* Farm list */}
      {farms.length > 0 && (
        <FarmList farms={farms} onEdit={openEditFarm} onRemove={removeFarm} />
      )}
    </div>
  );
}

// ─── MAIN APP ───────────────────────────────────────────────────────────────
export default function App() {
  const [processes, setProcesses] = useState([]);
  const [activeProcess, setActiveProcess] = useState(null);
  const [view, setView] = useState("dashboard");
  const [loaded, setLoaded] = useState(false);

  // Load from persistent storage
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("biosecurity-processes");
        if (res?.value) {
          setProcesses(JSON.parse(res.value));
        }
      } catch (e) {
        // no saved data
      }
      setLoaded(true);
    })();
  }, []);

  // Save to persistent storage on change
  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await window.storage.set("biosecurity-processes", JSON.stringify(processes));
      } catch (e) {
        console.error("Storage save failed:", e);
      }
    })();
  }, [processes, loaded]);

  const createProcess = ({ name, description }) => {
    const proc = {
      id: genId(),
      name,
      description,
      created: new Date().toISOString(),
      farms: [],
      sbmCriteria: [],
      selectedIBT: [...IBT_INDICES],
      results: null,
      classified: false,
    };
    setProcesses((prev) => [...prev, proc]);
    setActiveProcess(proc.id);
    setView("detail");
  };

  const updateProcess = (updated) => {
    setProcesses((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
  };

  const currentProcess = processes.find((p) => p.id === activeProcess);

  return (
    <div
      style={{
        minHeight: "100vh",
        background: theme.bg,
        color: theme.text,
        fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif",
      }}
    >
      {/* Header */}
      <div
        style={{
          borderBottom: `1px solid ${theme.border}`,
          padding: "14px 24px",
          display: "flex",
          alignItems: "center",
          gap: 14,
          background: theme.surface,
        }}
      >
        <div
          style={{
            width: 32,
            height: 32,
            borderRadius: 8,
            background: `linear-gradient(135deg, ${theme.accent}, #059669)`,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 16,
          }}
        >
          🛡
        </div>
        <div>
          <div style={{ fontSize: 15, fontWeight: 800, color: theme.text, letterSpacing: "-0.02em" }}>
            BioSecure
          </div>
          <div style={{ fontSize: 10, color: theme.textDim, letterSpacing: "0.08em", textTransform: "uppercase" }}>
            Farm Classification System
          </div>
        </div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: 11, color: theme.textDim }}>3-Scale Selection Model · PoC v1.0</div>
      </div>

      {/* Content */}
      <div style={{ maxWidth: 860, margin: "0 auto", padding: "24px 20px" }}>
        {view === "dashboard" && (
          <Dashboard
            processes={processes}
            onSelect={(id) => {
              setActiveProcess(id);
              setView("detail");
            }}
            onNew={() => setView("new")}
          />
        )}

        {view === "new" && <NewProcessForm onSave={createProcess} onCancel={() => setView("dashboard")} />}

        {view === "detail" && currentProcess && (
          <ProcessDetail
            process={currentProcess}
            onUpdate={updateProcess}
            onBack={() => {
              setActiveProcess(null);
              setView("dashboard");
            }}
          />
        )}
      </div>
    </div>
  );
}
