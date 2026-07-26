// Small geometric helpers shared by src/tables.js and src/layout.js.

/** Cluster sorted-or-not numbers into contiguous bands separated by gaps > `gap`. */
export function clusterByGap(values, gap) {
  const xs = [...values].sort((a, b) => a - b);
  const bands = [];
  for (const x of xs) {
    const last = bands[bands.length - 1];
    if (last && x - last.x1 <= gap) last.x1 = x;
    else bands.push({ x0: x, x1: x });
  }
  return bands;
}

export function nearestBandIndex(x, bands) {
  let best = 0, bestDist = Infinity;
  bands.forEach((b, i) => {
    const d = Math.abs(x - b.x0);
    if (d < bestDist) { bestDist = d; best = i; }
  });
  return best;
}
