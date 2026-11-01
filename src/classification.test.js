// Unit tests for the 3-Scale Selection Model classification logic.
// Run with:  npm test   (which runs `node --test`)
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  percentile,
  classify,
  runClassification,
  FC1_MAX_PCT,
  FC3_MIN_PCT,
  TOTAL_COL,
  EXTERNAL_COL,
  INTERNAL_COL,
  IBT_INDICES,
} from "./classification.js";

// Baseline percentiles that are clearly in the "pass the FC1 check" zone.
const SAFE = { totalPct: 80, extPct: 80, intPct: 80, ibtPcts: [80], scale3Pass: true };

// ─── named thresholds ───────────────────────────────────────────────────────
test("thresholds are the named 25 / 75 constants", () => {
  assert.equal(FC1_MAX_PCT, 25);
  assert.equal(FC3_MIN_PCT, 75);
});

// ─── percentile() ───────────────────────────────────────────────────────────
test("percentile: midpoint method (strictly-below full credit, ties half)", () => {
  // 2 values: the higher is at (1 + 0.5)/2 = 75, the lower at (0 + 0.5)/2 = 25.
  assert.equal(percentile([10, 20], 20), 75);
  assert.equal(percentile([10, 20], 10), 25);
  // empty set is defined as 0.
  assert.equal(percentile([], 5), 0);
});

// ─── FC1 via Scale 1 ─────────────────────────────────────────────────────────
test("FC1 via Scale 1: an External percentile of 18 (<= 25) forces FC1", () => {
  const { classification, reasons } = classify({ ...SAFE, extPct: 18 });
  assert.equal(classification, "FC1");
  assert.ok(reasons.includes("FC1: External percentile 18 <= 25"), reasons.join(" | "));
});

// ─── FC1 via Scale 2 ─────────────────────────────────────────────────────────
test("FC1 via Scale 2: a selected IBT percentile of 20 (<= 25) forces FC1", () => {
  const { classification, reasons } = classify({
    ...SAFE,
    ibtPcts: [60, 20, 90],
    ibtLabels: ["topic A", "topic B", "topic C"],
  });
  assert.equal(classification, "FC1");
  assert.ok(reasons.some((r) => r.includes("IBT topic B percentile 20 <= 25")), reasons.join(" | "));
});

// ─── FC1 via Scale 3 ─────────────────────────────────────────────────────────
test("FC1 via Scale 3: a failed SBM measure forces FC1", () => {
  const { classification, reasons } = classify({
    ...SAFE,
    scale3Pass: false,
    sbmFailLabels: ["Q42"],
  });
  assert.equal(classification, "FC1");
  assert.ok(reasons.includes("FC1: SBM Q42 wrong answer"), reasons.join(" | "));
});

// ─── FC2 at Total exactly 75 ─────────────────────────────────────────────────
test("FC2 at the Total == 75 boundary (<= 75 is still FC2)", () => {
  const { classification, reasons } = classify({ ...SAFE, totalPct: 75 });
  assert.equal(classification, "FC2");
  assert.ok(reasons[0].startsWith("FC2: Total percentile 75"), reasons[0]);
});

// ─── FC3 at Total 76 with External 40 ────────────────────────────────────────
test("FC3 at Total 76 even when External percentile is only 40", () => {
  // External 40 is > 25 so it is NOT an FC1 trigger, and External is ignored
  // for the FC2/FC3 split — Total 76 (> 75) alone makes this FC3.
  const { classification, reasons } = classify({
    totalPct: 76,
    extPct: 40,
    intPct: 80,
    ibtPcts: [80],
    scale3Pass: true,
  });
  assert.equal(classification, "FC3");
  assert.equal(reasons[0], "FC3: Total percentile 76 > 75");
});

// ─── boundary at exactly 25 ──────────────────────────────────────────────────
test("boundary: a Total percentile of exactly 25 is FC1 (<= 25)", () => {
  assert.equal(classify({ ...SAFE, totalPct: 25 }).classification, "FC1");
});

test("boundary: Total percentile of 25.0001 (just above 25) is not FC1", () => {
  assert.equal(classify({ ...SAFE, totalPct: 25.0001 }).classification, "FC2");
});

// ─── full pipeline through runClassification() ───────────────────────────────
function farm(id, total, ext, int, ibt) {
  const scores = { [TOTAL_COL]: total, [EXTERNAL_COL]: ext, [INTERNAL_COL]: int };
  // Give every IBT topic the same value so they don't accidentally trip FC1.
  for (const col of IBT_INDICES) scores[col] = ibt;
  return { id, name: id, answers: {}, scores };
}

test("runClassification: Total percentile is computed relative to the farm set", () => {
  // Two farms; the higher one sits at the 75th percentile, the lower at 25th.
  const farms = [farm("hi", 90, 90, 90, 90), farm("lo", 10, 10, 10, 10)];
  const results = runClassification(farms, [], [...IBT_INDICES]);
  const hi = results.find((r) => r.farmId === "hi");
  const lo = results.find((r) => r.farmId === "lo");

  assert.equal(hi.scale1.totalPct, 75);
  assert.equal(hi.finalClassification, "FC2"); // 75 -> FC2, nothing <= 25
  assert.equal(lo.scale1.totalPct, 25);
  assert.equal(lo.finalClassification, "FC1"); // every percentile 25 -> FC1
});

test("runClassification: FC2 vs FC3 ignores External/Internal/IBT", () => {
  // Build a set where the target farm has the TOP total (90th percentile) but a
  // mid-ranked External/Internal/IBT (50th percentile). It must still land FC3
  // on Total alone — the mid-range External/Internal/IBT neither trigger FC1
  // (all > 25) nor pull it down from FC3.
  const farms = [
    farm("top", 100, 50, 50, 50), // highest total; mid on every other dimension
    farm("a", 10, 10, 10, 10),
    farm("b", 20, 20, 20, 20),
    farm("c", 30, 90, 90, 90),
    farm("d", 40, 80, 80, 80),
  ];
  const results = runClassification(farms, [], [...IBT_INDICES]);
  const top = results.find((r) => r.farmId === "top");
  assert.equal(top.scale1.totalPct, 90); // (4 + 0.5)/5 * 100
  assert.equal(top.scale1.extPct, 50); // (2 + 0.5)/5 * 100
  assert.ok(top.scale1.extPct > FC1_MAX_PCT && top.scale1.extPct <= FC3_MIN_PCT);
  assert.equal(top.finalClassification, "FC3");
  assert.equal(top.reason, "FC3: Total percentile 90 > 75");
});

test("runClassification: a failed SBM regex overrides to FC1 with a reason", () => {
  const farms = [
    { id: "x", name: "x", scores: farm("x", 90, 90, 90, 90).scores, answers: { 46: "No" } },
    { id: "y", name: "y", scores: farm("y", 10, 10, 10, 10).scores, answers: { 46: "Yes" } },
  ];
  const sbm = [{ questionIdx: 46, questionNum: "42", questionText: "Does the cattle salesman come to the farm?", regexPattern: "^Yes" }];
  const results = runClassification(farms, sbm, [...IBT_INDICES]);
  const x = results.find((r) => r.farmId === "x");
  assert.equal(x.scale3.pass, false);
  assert.equal(x.finalClassification, "FC1"); // would be FC2 on scores alone
  assert.ok(x.reasons.includes("FC1: SBM Q42 wrong answer"), x.reasons.join(" | "));
});

test("runClassification: every farm carries a human-readable reason", () => {
  const farms = [farm("hi", 90, 90, 90, 90), farm("lo", 10, 10, 10, 10)];
  for (const r of runClassification(farms, [], [...IBT_INDICES])) {
    assert.equal(typeof r.reason, "string");
    assert.ok(r.reason.startsWith(r.finalClassification + ":"), r.reason);
  }
});
