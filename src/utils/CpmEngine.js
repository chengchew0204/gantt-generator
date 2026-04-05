/**
 * Critical Path Method (CPM) engine.
 *
 * Computes earlyStart, earlyFinish, lateStart, lateFinish, and totalFloat
 * for each task. Tasks with totalFloat === 0 lie on the critical path.
 *
 * The engine respects actual calendar dates: tasks without predecessors
 * use their startDate as the anchor. All scheduling is in day-offset
 * units from the earliest date in the project.
 *
 * Each task is expected to have at minimum:
 *   - id         (string|number) unique task identifier
 *   - duration   (number) duration in days
 *   - dependency (string) comma-separated predecessor task IDs, or empty
 *   - startDate  (string|null) ISO date YYYY-MM-DD
 *   - endDate    (string|null) ISO date YYYY-MM-DD
 */

export function computeCpm(tasks) {
  if (!tasks || tasks.length === 0) return [];

  const projectStart = findProjectStart(tasks);

  const taskMap = new Map();
  const results = [];

  for (const task of tasks) {
    const id = String(task.id);
    const duration = parseDuration(task.duration);
    const predecessors = parseDependencies(task.dependency);
    const dateOffset = task.startDate ? daysBetween(projectStart, task.startDate) : null;

    const entry = {
      id,
      duration,
      predecessors,
      successors: [],
      dateOffset,
      earlyStart: 0,
      earlyFinish: 0,
      lateStart: Infinity,
      lateFinish: Infinity,
      totalFloat: 0,
      isCritical: false,
    };
    taskMap.set(id, entry);
    results.push(entry);
  }

  for (const entry of results) {
    for (const predId of entry.predecessors) {
      const pred = taskMap.get(predId);
      if (pred) {
        pred.successors.push(entry.id);
      }
    }
  }

  const sorted = topologicalSort(results, taskMap);

  // Forward pass
  for (const id of sorted) {
    const entry = taskMap.get(id);

    let maxPredFinish = 0;
    for (const predId of entry.predecessors) {
      const pred = taskMap.get(predId);
      if (pred && pred.earlyFinish > maxPredFinish) {
        maxPredFinish = pred.earlyFinish;
      }
    }

    if (entry.predecessors.length === 0 && entry.dateOffset != null) {
      entry.earlyStart = Math.max(entry.dateOffset, 0);
    } else {
      entry.earlyStart = maxPredFinish;
    }

    entry.earlyFinish = entry.earlyStart + entry.duration;
  }

  let projectFinish = 0;
  for (const entry of results) {
    if (entry.earlyFinish > projectFinish) {
      projectFinish = entry.earlyFinish;
    }
  }

  // Backward pass
  for (const entry of results) {
    if (entry.successors.length === 0) {
      entry.lateFinish = projectFinish;
    }
  }

  for (let i = sorted.length - 1; i >= 0; i--) {
    const entry = taskMap.get(sorted[i]);

    if (entry.lateFinish === Infinity) {
      entry.lateFinish = projectFinish;
    }
    entry.lateStart = entry.lateFinish - entry.duration;

    for (const predId of entry.predecessors) {
      const pred = taskMap.get(predId);
      if (pred && entry.lateStart < pred.lateFinish) {
        pred.lateFinish = entry.lateStart;
      }
    }
  }

  for (const entry of results) {
    entry.totalFloat = entry.lateStart - entry.earlyStart;
    entry.isCritical = entry.totalFloat === 0 && entry.duration > 0;
  }

  return results;
}

function parseDuration(value) {
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : 0;
}

function parseDependencies(value) {
  if (!value || typeof value !== 'string') return [];
  return value
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

function findProjectStart(tasks) {
  let min = null;
  for (const t of tasks) {
    if (t.startDate && (!min || t.startDate < min)) {
      min = t.startDate;
    }
  }
  if (min) return min;
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

function parseLocal(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d, 12);
}

function daysBetween(a, b) {
  return Math.round((parseLocal(b) - parseLocal(a)) / (1000 * 60 * 60 * 24));
}

/**
 * Kahn's algorithm for topological sort.
 * Falls back to insertion order for tasks involved in cycles.
 */
function topologicalSort(entries, taskMap) {
  const inDegree = new Map();
  for (const entry of entries) {
    inDegree.set(entry.id, 0);
  }

  for (const entry of entries) {
    for (const predId of entry.predecessors) {
      if (taskMap.has(predId)) {
        inDegree.set(entry.id, (inDegree.get(entry.id) || 0) + 1);
      }
    }
  }

  const queue = [];
  for (const entry of entries) {
    if (inDegree.get(entry.id) === 0) {
      queue.push(entry.id);
    }
  }

  const sorted = [];
  while (queue.length > 0) {
    const id = queue.shift();
    sorted.push(id);

    const entry = taskMap.get(id);
    for (const succId of entry.successors) {
      const newDeg = inDegree.get(succId) - 1;
      inDegree.set(succId, newDeg);
      if (newDeg === 0) {
        queue.push(succId);
      }
    }
  }

  if (sorted.length < entries.length) {
    const sortedSet = new Set(sorted);
    for (const entry of entries) {
      if (!sortedSet.has(entry.id)) {
        sorted.push(entry.id);
      }
    }
  }

  return sorted;
}
