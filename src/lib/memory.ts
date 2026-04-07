// Lightweight memory system for Cel - no dependencies, just localStorage

export interface CelMemory {
  preferences: {
    provider?: string;
    model?: string;
    language?: string;
  };
  mistakes: { what: string; fix: string; date: string }[];
  commonOps: { op: string; count: number }[];
  rules: string[];
  conventions: {
    currency?: string;
    dateFormat?: string;
    decimalPlaces?: number;
    negativeStyle?: string;
  };
  checkpoints: {
    id: string;
    sheet: string;
    address: string;
    values: any[][];
    timestamp: number;
    description: string;
  }[];
}

const MEMORY_KEY = 'cel-memory';

const defaultMemory: CelMemory = {
  preferences: {},
  mistakes: [],
  commonOps: [],
  rules: [],
  conventions: {},
  checkpoints: [],
};

export function loadMemory(): CelMemory {
  try {
    const saved = localStorage.getItem(MEMORY_KEY);
    if (saved) return { ...defaultMemory, ...JSON.parse(saved) };
  } catch {}
  return { ...defaultMemory };
}

export function saveMemory(data: Partial<CelMemory>) {
  const current = loadMemory();
  const merged = { ...current, ...data };
  localStorage.setItem(MEMORY_KEY, JSON.stringify(merged));
}

export function addMistake(what: string, fix: string) {
  const memory = loadMemory();
  memory.mistakes.unshift({ what, fix, date: new Date().toISOString() });
  if (memory.mistakes.length > 20) memory.mistakes = memory.mistakes.slice(0, 20);
  saveMemory({ mistakes: memory.mistakes });
}

export function trackOperation(op: string) {
  const memory = loadMemory();
  const existing = memory.commonOps.find((o) => o.op === op);
  if (existing) {
    existing.count++;
  } else {
    memory.commonOps.push({ op, count: 1 });
  }
  memory.commonOps.sort((a, b) => b.count - a.count);
  if (memory.commonOps.length > 20) memory.commonOps = memory.commonOps.slice(0, 20);
  saveMemory({ commonOps: memory.commonOps });
}

export function addCheckpoint(checkpoint: Omit<CelMemory['checkpoints'][0], 'id' | 'timestamp'>) {
  const memory = loadMemory();
  const newCheckpoint = {
    ...checkpoint,
    id: `chk_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    timestamp: Date.now(),
  };
  memory.checkpoints.unshift(newCheckpoint);
  if (memory.checkpoints.length > 50) memory.checkpoints = memory.checkpoints.slice(0, 50);
  saveMemory({ checkpoints: memory.checkpoints });
  return newCheckpoint.id;
}

export function getCheckpoints(): CelMemory['checkpoints'] {
  return loadMemory().checkpoints;
}

export function removeCheckpoint(id: string) {
  const memory = loadMemory();
  memory.checkpoints = memory.checkpoints.filter((c) => c.id !== id);
  saveMemory({ checkpoints: memory.checkpoints });
}

export function clearCheckpoints() {
  saveMemory({ checkpoints: [] });
}

export function getMemoryContext(): string {
  const memory = loadMemory();
  const parts: string[] = [];

  if (memory.rules.length > 0) {
    parts.push(`USER RULES (follow these):\n${memory.rules.map((r, i) => `${i + 1}. ${r}`).join('\n')}`);
  }

  if (memory.mistakes.length > 0) {
    const recent = memory.mistakes.slice(0, 5);
    parts.push(`PAST MISTAKES TO AVOID:\n${recent.map((m) => `- ${m.what} → Fix: ${m.fix}`).join('\n')}`);
  }

  if (memory.conventions.currency) {
    parts.push(`CURRENCY: Use ${memory.conventions.currency} for all currency formatting`);
  }
  if (memory.conventions.dateFormat) {
    parts.push(`DATE FORMAT: Use ${memory.conventions.dateFormat}`);
  }
  if (memory.conventions.decimalPlaces !== undefined) {
    parts.push(`DECIMAL PLACES: Use ${memory.conventions.decimalPlaces} decimal places`);
  }
  if (memory.conventions.negativeStyle) {
    parts.push(`NEGATIVE STYLE: ${memory.conventions.negativeStyle}`);
  }

  return parts.join('\n\n');
}
