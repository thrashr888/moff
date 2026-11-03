const timestamp = /\b\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?\b/g;

export function hasRepetitionLoop(text: string, sourceText = ''): boolean {
  const tail = text.slice(-8192);
  const lines = tail.split('\n').map(line =>
    line.trim().replace(/^\d+[.)]\s+/, '').replace(/\s+/g, ' '),
  ).filter(Boolean);
  const last = lines.at(-1);
  if (last && last.length >= 40 && lines.length >= 6 &&
      lines.slice(-6).every(line => line === last)) return true;

  // A stuck model may advance the list number and invent a new timestamp on
  // each copy. Only flag timestamp variants when several lack source evidence.
  const completed = text.endsWith('\n') ? lines : lines.slice(0, -1);
  const recent = completed.slice(-6);
  if (recent.length === 6) {
    const patterns = recent.map(line => line.replace(timestamp, '<timestamp>'));
    if (patterns[0].length >= 40 && patterns.every(pattern => pattern === patterns[0])) {
      const sourceTimes = new Set(sourceText.match(timestamp) || []);
      const unsupported = recent.filter(line => {
        const times = line.match(timestamp) || [];
        return times.length > 0 && times.some(time => !sourceTimes.has(time));
      });
      if (unsupported.length >= 3) return true;
    }
  }

  // Catch repeated phrases/paragraphs even when they arrive without newlines.
  // Require six exact copies of a substantial block, not repeated short words.
  for (let size = 40; size <= Math.min(1024, Math.floor(tail.length / 6)); size++) {
    const block = tail.slice(-size);
    if (block.trim().length < 32 || !/\s/.test(block)) continue;
    let repeated = true;
    for (let copy = 2; copy <= 6; copy++) {
      if (tail.slice(-copy * size, -(copy - 1) * size) !== block) {
        repeated = false;
        break;
      }
    }
    if (repeated) return true;
  }
  return false;
}
