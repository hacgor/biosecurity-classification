// ─── 3-SCALE SELECTION MODEL (3-SSM) — CLASSIFICATION LOGIC ──────────────────
// Pure, UI-free classification logic. Extracted from App.jsx so it can be unit
// tested without a JSX/React runtime (see classification.test.js).
//
// Runs per classification process, over ALL farms in that process. Percentiles
// are computed for each farm relative to the other farms in the SAME process.
//
// Algorithm (per farm, in this exact order):
//   1. FC1 "Poor" if ANY of:
//        a. any Scale 1 percentile (Total / External / Internal) <= 25
//        b. any selected Scale 2 IBT percentile <= 25
//        c. Scale 3 fails (at least one SBM question lacks its required answer)
//   2. Otherwise, look ONLY at the Scale 1 Total percentile:
//        - Total > 25 and <= 75  -> FC2 "Medium"
//        - Total > 75            -> FC3 "Good"
//      External / Internal / IBT percentiles do NOT affect the FC2 vs FC3 split.

// ─── NAMED THRESHOLDS ───────────────────────────────────────────────────────
// A percentile at or below FC1_MAX_PCT forces FC1. Above FC3_MIN_PCT (and not
// an FC1 trigger) is FC3; everything in between is FC2.
export const FC1_MAX_PCT = 25;
export const FC3_MIN_PCT = 75;

// ─── EXCEL COLUMN INDICES (0-based) ─────────────────────────────────────────
// Scale 1 aggregate scores.
export const EXTERNAL_COL = 149;
export const INTERNAL_COL = 150;
export const TOTAL_COL = 151;

// Scale 2 — Important Biosecurity Topics: topic score columns A..K (138..148).
export const IBT_INDICES = [138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148];

// Topic letter (A..K) for an IBT column index — used in human-readable reasons.
function ibtLetter(idx) {
  return String.fromCharCode(65 + (idx - IBT_INDICES[0]));
}

// Format a percentile for a reason string without a trailing ".0".
function fmtPct(p) {
  return Number.isInteger(p) ? String(p) : p.toFixed(1);
}

// ─── PERCENTILE ─────────────────────────────────────────────────────────────
// Percentile of `val` within `arr` using the midpoint method: values strictly
// below get full credit, ties get half credit.
export function percentile(arr, val) {
  if (!arr || arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  let count = 0;
  for (const v of sorted) {
    if (v < val) count++;
    else if (v === val) count += 0.5;
  }
  return (count / sorted.length) * 100;
}

// ─── DECISION RULE ──────────────────────────────────────────────────────────
// Decide the final class from pre-computed percentiles + the Scale 3 result.
// Kept separate from percentile computation so the rule can be unit tested with
// exact percentile inputs (boundaries at 25 and 75).
//
// Returns { classification: "FC1"|"FC2"|"FC3", reasons: string[] } where each
// reason is prefixed with the class it supports, e.g.
//   "FC1: External percentile 18 <= 25"
//   "FC1: SBM Q42 wrong answer"
//   "FC3: Total percentile 82 > 75"
export function classify({
  totalPct,
  extPct,
  intPct,
  ibtPcts = [],
  ibtLabels = [],
  scale3Pass = true,
  sbmFailLabels = [],
}) {
  const fc1 = [];

  // 1a — Scale 1 percentiles.
  if (totalPct <= FC1_MAX_PCT) fc1.push(`Total percentile ${fmtPct(totalPct)} <= ${FC1_MAX_PCT}`);
  if (extPct <= FC1_MAX_PCT) fc1.push(`External percentile ${fmtPct(extPct)} <= ${FC1_MAX_PCT}`);
  if (intPct <= FC1_MAX_PCT) fc1.push(`Internal percentile ${fmtPct(intPct)} <= ${FC1_MAX_PCT}`);

  // 1b — Scale 2 IBT percentiles.
  ibtPcts.forEach((p, i) => {
    if (p <= FC1_MAX_PCT) {
      const label = ibtLabels[i] ?? `#${i + 1}`;
      fc1.push(`IBT ${label} percentile ${fmtPct(p)} <= ${FC1_MAX_PCT}`);
    }
  });

  // 1c — Scale 3 SBM.
  if (!scale3Pass) {
    if (sbmFailLabels.length > 0) {
      for (const lbl of sbmFailLabels) fc1.push(`SBM ${lbl} wrong answer`);
    } else {
      fc1.push("SBM criteria not met");
    }
  }

  if (fc1.length > 0) {
    return { classification: "FC1", reasons: fc1.map((r) => `FC1: ${r}`) };
  }

  // 2 — passed the FC1 check: Total percentile alone decides FC2 vs FC3.
  if (totalPct <= FC3_MIN_PCT) {
    return {
      classification: "FC2",
      reasons: [`FC2: Total percentile ${fmtPct(totalPct)} > ${FC1_MAX_PCT} and <= ${FC3_MIN_PCT}`],
    };
  }
  return {
    classification: "FC3",
    reasons: [`FC3: Total percentile ${fmtPct(totalPct)} > ${FC3_MIN_PCT}`],
  };
}

// ─── FULL CLASSIFICATION ────────────────────────────────────────────────────
// Classify every farm in a process. Computes percentiles relative to the farm
// set, evaluates the Scale 3 SBM criteria (regex, case-insensitive), then applies
// the decision rule above.
export function runClassification(farms, sbmCriteria = [], selectedIBT = []) {
  if (!farms || farms.length === 0) return [];

  // Only the topics the expert selected are examined in Scale 2. If none are
  // selected, fall back to all topics.
  const activeIBT = selectedIBT && selectedIBT.length > 0 ? selectedIBT : IBT_INDICES;

  // Score arrays for percentile computation (relative to this farm set only).
  const totalScores = farms.map((f) => f.scores?.[TOTAL_COL] ?? 0);
  const externalScores = farms.map((f) => f.scores?.[EXTERNAL_COL] ?? 0);
  const internalScores = farms.map((f) => f.scores?.[INTERNAL_COL] ?? 0);
  const ibtArrays = {};
  for (const idx of activeIBT) {
    ibtArrays[idx] = farms.map((f) => f.scores?.[idx] ?? 0);
  }

  return farms.map((farm) => {
    const totalPct = percentile(totalScores, farm.scores?.[TOTAL_COL] ?? 0);
    const extPct = percentile(externalScores, farm.scores?.[EXTERNAL_COL] ?? 0);
    const intPct = percentile(internalScores, farm.scores?.[INTERNAL_COL] ?? 0);

    const ibtPcts = activeIBT.map((idx) => percentile(ibtArrays[idx], farm.scores?.[idx] ?? 0));
    const ibtLabels = activeIBT.map((idx) => `topic ${ibtLetter(idx)}`);

    // Scale 3: every SBM criterion must match its required answer.
    let scale3Pass = true;
    const sbmDetails = [];
    const sbmFailLabels = [];
    for (const crit of sbmCriteria) {
      const farmAnswer = String(farm.answers?.[crit.questionIdx] ?? "").trim();
      let pass = false;
      let regexError = null;
      try {
        pass = new RegExp(crit.regexPattern, "i").test(farmAnswer);
      } catch (e) {
        regexError = e.message;
        pass = false;
      }
      if (!pass) {
        scale3Pass = false;
        sbmFailLabels.push(`Q${crit.questionNum}`);
      }
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

    const { classification, reasons } = classify({
      totalPct,
      extPct,
      intPct,
      ibtPcts,
      ibtLabels,
      scale3Pass,
      sbmFailLabels,
    });

    // Per-scale FC1 triggers (for display/debugging).
    const scale1FC1 = [];
    if (totalPct <= FC1_MAX_PCT) scale1FC1.push("Total");
    if (extPct <= FC1_MAX_PCT) scale1FC1.push("External");
    if (intPct <= FC1_MAX_PCT) scale1FC1.push("Internal");
    const scale2FC1 = activeIBT.filter((idx, i) => ibtPcts[i] <= FC1_MAX_PCT);

    const minIbtPct = ibtPcts.length > 0 ? Math.min(...ibtPcts) : 100;

    return {
      farmId: farm.id,
      farmName: farm.name,
      finalClassification: classification,
      reason: reasons[0],
      reasons,
      scale1: { totalPct, extPct, intPct, triggersFC1: scale1FC1.length > 0, fc1Scores: scale1FC1 },
      scale2: { activeIBT, ibtPcts, minIbtPct, triggersFC1: scale2FC1.length > 0, fc1Topics: scale2FC1 },
      scale3: { pass: scale3Pass, details: sbmDetails, fc1Questions: sbmFailLabels },
      scores: farm.scores,
    };
  });
}
