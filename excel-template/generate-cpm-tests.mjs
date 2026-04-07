/**
 * Generate Excel test files specifically designed to exercise the
 * CPM engine's critical path detection and slack (total float) calculation.
 *
 * Run: node excel-template/generate-cpm-tests.mjs
 */

import XLSX from 'xlsx';
import { writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

const TASK_COLUMNS = [
  'Task ID', 'Task Name', 'Dependency', 'Task Category',
  'Start Date', 'End Date', 'Duration', 'Progress (%)',
  'Status', 'Owner', 'Remarks', 'Baseline Start', 'Baseline End', 'Parent ID',
];

const SETTINGS_FIELDS = [
  ['Project Name', ''],
  ['Theme Name', 'Linear Dark'],
  ['Theme BG Primary', '#0f0f12'],
  ['Theme BG Secondary', '#1a1a23'],
  ['Theme BG Tertiary', '#23232f'],
  ['Theme BG Hover', '#2a2a38'],
  ['Theme Border', '#2e2e3a'],
  ['Theme Border Subtle', '#232330'],
  ['Theme Text Primary', '#e8e8ed'],
  ['Theme Text Secondary', '#9898a8'],
  ['Theme Text Muted', '#66667a'],
  ['Theme Accent', '#6366f1'],
  ['Theme Accent Hover', '#818cf8'],
  ['Theme Success', '#22c55e'],
  ['Theme Warning', '#f59e0b'],
  ['Theme Danger', '#ef4444'],
  ['Theme Info', '#3b82f6'],
  ['Theme Critical Path', '#f43f5e'],
  ['Show Critical Path', 'true'],
  ['Show Slack', 'true'],
  ['Show Dependencies', 'true'],
  ['Show Today Line', 'true'],
  ['Show Baseline', 'true'],
  ['Skip Weekends', 'false'],
  ['Show Scale Buttons', 'true'],
  ['Show Week Labels', 'false'],
  ['Show Month Labels', 'true'],
  ['Show Day Labels', 'true'],
  ['Visible Columns', 'id,name,dependency,duration,startDate,endDate,progress,status,owner'],
  ['Category Colors', '{}'],
];

function addDays(iso, days) {
  const [y, m, d] = iso.split('-').map(Number);
  const date = new Date(y, m - 1, d + days);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function buildWorkbook(projectName, tasks) {
  const wb = XLSX.utils.book_new();

  const taskRows = [TASK_COLUMNS];
  for (const t of tasks) {
    taskRows.push([
      t.id, t.name, t.dep || '', t.cat || '',
      t.start, t.end, t.dur,
      t.progress ?? 0, t.status || 'Not Started',
      t.owner || '', t.remarks || '',
      t.blStart || '', t.blEnd || '',
      t.parentId || '',
    ]);
  }
  const taskSheet = XLSX.utils.aoa_to_sheet(taskRows);
  taskSheet['!cols'] = TASK_COLUMNS.map((c) => ({ wch: Math.max(c.length + 2, 14) }));
  XLSX.utils.book_append_sheet(wb, taskSheet, 'Tasks');

  const settings = SETTINGS_FIELDS.map(([k, v]) =>
    k === 'Project Name' ? [k, projectName] : [k, v],
  );
  settings.unshift(['Setting', 'Value']);
  const settingsSheet = XLSX.utils.aoa_to_sheet(settings);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  return wb;
}

function save(filename, wb) {
  const path = join(__dirname, filename);
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  writeFileSync(path, buf);
  console.log(`  Created: ${filename}`);
}

// ---------------------------------------------------------------------------
// Sample 10: Simple Linear Chain (all tasks are critical, zero slack)
// A -> B -> C -> D -> E  (pure sequential)
// ---------------------------------------------------------------------------
function genLinearChain() {
  const base = '2026-05-04';
  const tasks = [
    { id: 1, name: 'Requirements Gathering', dep: '', start: base, end: addDays(base, 4), dur: 5, status: 'Completed', progress: 100, owner: 'Alice' },
    { id: 2, name: 'System Design', dep: '1', start: addDays(base, 5), end: addDays(base, 9), dur: 5, owner: 'Bob' },
    { id: 3, name: 'Implementation', dep: '2', start: addDays(base, 10), end: addDays(base, 19), dur: 10, owner: 'Carol' },
    { id: 4, name: 'Testing', dep: '3', start: addDays(base, 20), end: addDays(base, 24), dur: 5, owner: 'Dave' },
    { id: 5, name: 'Deployment', dep: '4', start: addDays(base, 25), end: addDays(base, 27), dur: 3, owner: 'Eve' },
  ];
  save('10-cpm-linear-chain.xlsx', buildWorkbook('CPM Linear Chain', tasks));
}

// ---------------------------------------------------------------------------
// Sample 11: Two Parallel Paths -- one longer (critical), one shorter (has slack)
//
//   1 (3d) ---> 2 (7d) ---> 4 (3d)   <-- critical path (total=13d)
//   1 (3d) ---> 3 (4d) ---> 4 (3d)   <-- non-critical (total=10d, slack=3d on task 3)
// ---------------------------------------------------------------------------
function genParallelPaths() {
  const base = '2026-05-04';
  const tasks = [
    { id: 1, name: 'Project Kickoff', dep: '', start: base, end: addDays(base, 2), dur: 3, owner: 'Alice', cat: 'Planning' },
    { id: 2, name: 'Backend Development', dep: '1', start: addDays(base, 3), end: addDays(base, 9), dur: 7, owner: 'Bob', cat: 'Backend', remarks: 'CRITICAL PATH -- longest route' },
    { id: 3, name: 'Frontend Development', dep: '1', start: addDays(base, 3), end: addDays(base, 6), dur: 4, owner: 'Carol', cat: 'Frontend', remarks: 'Has 3 days of slack (total float)' },
    { id: 4, name: 'Integration Testing', dep: '2,3', start: addDays(base, 10), end: addDays(base, 12), dur: 3, owner: 'Dave', cat: 'QA', remarks: 'Waits for both paths' },
  ];
  save('11-cpm-parallel-paths.xlsx', buildWorkbook('CPM Parallel Paths', tasks));
}

// ---------------------------------------------------------------------------
// Sample 12: Diamond Dependency with Multiple Slack Values
//
//          +-- B (5d) --+
//   A (2d)-+-- C (3d) --+-- F (2d)
//          +-- D (8d) --+
//          +-- E (1d) --+
//
// Critical: A -> D -> F (total = 12d)
// B has 3d slack, C has 5d slack, E has 7d slack
// ---------------------------------------------------------------------------
function genDiamondDependency() {
  const base = '2026-06-01';
  const tasks = [
    { id: 1, name: 'Requirements Analysis', dep: '', start: base, end: addDays(base, 1), dur: 2, owner: 'PM', cat: 'Planning' },
    { id: 2, name: 'Database Schema Design', dep: '1', start: addDays(base, 2), end: addDays(base, 6), dur: 5, owner: 'DBA', cat: 'Backend', remarks: 'Slack = 3 days' },
    { id: 3, name: 'UI Wireframes', dep: '1', start: addDays(base, 2), end: addDays(base, 4), dur: 3, owner: 'Designer', cat: 'Design', remarks: 'Slack = 5 days' },
    { id: 4, name: 'Core API Development', dep: '1', start: addDays(base, 2), end: addDays(base, 9), dur: 8, owner: 'Lead Dev', cat: 'Backend', remarks: 'CRITICAL PATH -- zero slack' },
    { id: 5, name: 'Documentation', dep: '1', start: addDays(base, 2), end: addDays(base, 2), dur: 1, owner: 'Tech Writer', cat: 'Docs', remarks: 'Slack = 7 days' },
    { id: 6, name: 'System Integration', dep: '2,3,4,5', start: addDays(base, 10), end: addDays(base, 11), dur: 2, owner: 'Lead Dev', cat: 'Integration', remarks: 'CRITICAL -- waits for all four paths' },
  ];
  save('12-cpm-diamond-dependency.xlsx', buildWorkbook('CPM Diamond Dependency', tasks));
}

// ---------------------------------------------------------------------------
// Sample 13: Multiple Independent Chains
// Chain A: 1->2->3  (10d total)
// Chain B: 4->5     (8d total)
// Chain C: 6->7->8  (12d total -- longest, drives project finish)
// All start at the same date. Chains A and B have slack relative to project finish.
// ---------------------------------------------------------------------------
function genIndependentChains() {
  const base = '2026-06-15';
  const tasks = [
    // Chain A (10 days total)
    { id: 1, name: 'Chain A: Research', dep: '', start: base, end: addDays(base, 3), dur: 4, owner: 'Team A', cat: 'Research', remarks: 'Slack = 2 days (project finishes at day 12)' },
    { id: 2, name: 'Chain A: Prototype', dep: '1', start: addDays(base, 4), end: addDays(base, 6), dur: 3, owner: 'Team A', cat: 'Dev' },
    { id: 3, name: 'Chain A: Review', dep: '2', start: addDays(base, 7), end: addDays(base, 9), dur: 3, owner: 'Team A', cat: 'Review' },
    // Chain B (8 days total)
    { id: 4, name: 'Chain B: Vendor Setup', dep: '', start: base, end: addDays(base, 4), dur: 5, owner: 'Team B', cat: 'Procurement', remarks: 'Slack = 4 days' },
    { id: 5, name: 'Chain B: Installation', dep: '4', start: addDays(base, 5), end: addDays(base, 7), dur: 3, owner: 'Team B', cat: 'Infra' },
    // Chain C (12 days total -- critical)
    { id: 6, name: 'Chain C: Foundation', dep: '', start: base, end: addDays(base, 4), dur: 5, owner: 'Team C', cat: 'Build', remarks: 'CRITICAL' },
    { id: 7, name: 'Chain C: Structure', dep: '6', start: addDays(base, 5), end: addDays(base, 8), dur: 4, owner: 'Team C', cat: 'Build', remarks: 'CRITICAL' },
    { id: 8, name: 'Chain C: Finishing', dep: '7', start: addDays(base, 9), end: addDays(base, 11), dur: 3, owner: 'Team C', cat: 'Build', remarks: 'CRITICAL' },
  ];
  save('13-cpm-independent-chains.xlsx', buildWorkbook('CPM Independent Chains', tasks));
}

// ---------------------------------------------------------------------------
// Sample 14: Complex Project (WBS + milestones + baselines + mixed statuses)
// Tests CPM on a realistic project with parent tasks, milestones (dur=0),
// baseline comparison, and dependency chains of varying lengths.
//
// Structure:
//   Phase 1 (parent)
//     1.1 -> 1.2 -> 1.3      (critical: 5+4+3 = 12d)
//     1.1 -> 1.4              (non-critical: 5+2 = 7d, slack = 5d on 1.4)
//   M1 milestone (dep: 1.3, 1.4)
//   Phase 2 (parent)
//     2.1 -> 2.2 -> 2.3      (critical: 6+5+3 = 14d)
//     2.1 -> 2.4              (non-critical: 6+4 = 10d, slack = 4d)
//   M2 milestone (dep: 2.3, 2.4)
//   Phase 3 (parent)
//     3.1 -> 3.2              (3+2 = 5d)
//   M3 final milestone
// ---------------------------------------------------------------------------
function genComplexProject() {
  const base = '2026-05-11';
  const tasks = [
    // Phase 1
    { id: 100, name: 'Phase 1: Planning & Design', dep: '', start: base, end: addDays(base, 11), dur: 12, parentId: '', cat: 'Phase 1' },
    { id: 101, name: 'Stakeholder Interviews', dep: '', start: base, end: addDays(base, 4), dur: 5, owner: 'Alice', parentId: '100', cat: 'Planning',
      blStart: base, blEnd: addDays(base, 4), status: 'Completed', progress: 100 },
    { id: 102, name: 'Architecture Design', dep: '101', start: addDays(base, 5), end: addDays(base, 8), dur: 4, owner: 'Bob', parentId: '100', cat: 'Design',
      blStart: addDays(base, 5), blEnd: addDays(base, 7), remarks: 'CRITICAL -- baseline was 3d, actual is 4d (slipped 1d)' },
    { id: 103, name: 'Technical Spec Review', dep: '102', start: addDays(base, 9), end: addDays(base, 11), dur: 3, owner: 'Carol', parentId: '100', cat: 'Review',
      blStart: addDays(base, 8), blEnd: addDays(base, 10), remarks: 'CRITICAL -- started 1d late due to 102 slip' },
    { id: 104, name: 'Risk Assessment', dep: '101', start: addDays(base, 5), end: addDays(base, 6), dur: 2, owner: 'Dave', parentId: '100', cat: 'Planning',
      blStart: addDays(base, 5), blEnd: addDays(base, 6), remarks: 'Non-critical -- has 5 days of slack' },

    // M1 milestone
    { id: 110, name: 'M1: Planning Complete', dep: '103,104', start: addDays(base, 12), end: addDays(base, 12), dur: 0, cat: 'Milestone', remarks: 'Gate milestone' },

    // Phase 2
    { id: 200, name: 'Phase 2: Development', dep: '', start: addDays(base, 12), end: addDays(base, 25), dur: 14, parentId: '', cat: 'Phase 2' },
    { id: 201, name: 'Backend API', dep: '110', start: addDays(base, 12), end: addDays(base, 17), dur: 6, owner: 'Eve', parentId: '200', cat: 'Backend',
      blStart: addDays(base, 11), blEnd: addDays(base, 16), status: 'In Progress', progress: 40, remarks: 'CRITICAL -- baseline started 1d earlier' },
    { id: 202, name: 'Frontend UI', dep: '201', start: addDays(base, 18), end: addDays(base, 22), dur: 5, owner: 'Frank', parentId: '200', cat: 'Frontend',
      blStart: addDays(base, 17), blEnd: addDays(base, 21), remarks: 'CRITICAL' },
    { id: 203, name: 'End-to-End Testing', dep: '202', start: addDays(base, 23), end: addDays(base, 25), dur: 3, owner: 'Grace', parentId: '200', cat: 'QA',
      blStart: addDays(base, 22), blEnd: addDays(base, 24), remarks: 'CRITICAL' },
    { id: 204, name: 'API Documentation', dep: '201', start: addDays(base, 18), end: addDays(base, 21), dur: 4, owner: 'Hank', parentId: '200', cat: 'Docs',
      blStart: addDays(base, 17), blEnd: addDays(base, 20), remarks: 'Non-critical -- has 4 days of slack' },

    // M2 milestone
    { id: 210, name: 'M2: Development Complete', dep: '203,204', start: addDays(base, 26), end: addDays(base, 26), dur: 0, cat: 'Milestone' },

    // Phase 3
    { id: 300, name: 'Phase 3: Deployment', dep: '', start: addDays(base, 26), end: addDays(base, 30), dur: 5, parentId: '', cat: 'Phase 3' },
    { id: 301, name: 'Production Setup', dep: '210', start: addDays(base, 26), end: addDays(base, 28), dur: 3, owner: 'Ivy', parentId: '300', cat: 'DevOps',
      blStart: addDays(base, 25), blEnd: addDays(base, 27), remarks: 'CRITICAL' },
    { id: 302, name: 'Go-Live & Monitoring', dep: '301', start: addDays(base, 29), end: addDays(base, 30), dur: 2, owner: 'Jack', parentId: '300', cat: 'DevOps',
      blStart: addDays(base, 28), blEnd: addDays(base, 29), remarks: 'CRITICAL' },

    // Final milestone
    { id: 310, name: 'M3: Project Complete', dep: '302', start: addDays(base, 31), end: addDays(base, 31), dur: 0, cat: 'Milestone' },
  ];
  save('14-cpm-complex-project.xlsx', buildWorkbook('CPM Complex Project', tasks));
}

// ---------------------------------------------------------------------------
// Sample 15: Ladder / Staggered Dependencies
// Creates a "ladder" pattern where tasks on two parallel tracks alternate
// their dependencies, forcing the CPM to pick the longest combination.
//
//   A1 (3d) -----> A2 (4d) -----> A3 (2d) -----> Finish (1d)
//     \              \              \
//      +-> B1 (2d) --+-> B2 (6d) --+-> B3 (1d) ---> Finish
//
// Critical path: A1 -> B1 -> A2? No. Let's compute carefully.
// A1(3) -> A2(4)  earliest A2 finishes at 7
// A1(3) -> B1(2)  earliest B1 finishes at 5
// B1(5) -> A2(4)  but A2 also depends on A1 (finishes at 3), and B1 finishes at 5 -- so A2 starts at 5, finishes at 9
// Wait, let me design it more carefully:
//
//   A1(3) --> A2(4) --> A3(2) --> Merge(1)
//   A1(3) --> B1(5) --> B2(3) --> Merge(1)
//   B1 also depends on A1 (already included)
//   A2 depends on A1 AND B1
//   B2 depends on B1 AND A2
//   A3 depends on A2
//   Merge depends on A3 AND B2
//
// Forward:
//   A1: ES=0, EF=3
//   B1: ES=3, EF=8  (depends on A1)
//   A2: ES=max(3,8)=8, EF=12  (depends on A1, B1)
//   B2: ES=max(8,12)=12, EF=15  (depends on B1, A2)
//   A3: ES=12, EF=14  (depends on A2)
//   Merge: ES=max(14,15)=15, EF=16  (depends on A3, B2)
//
// Critical path: A1 -> B1 -> A2 -> B2 -> Merge (total 16d)
// A3 has 1d slack (finishes at 14, but Merge doesn't start until 15)
// ---------------------------------------------------------------------------
function genLadderDependency() {
  const base = '2026-07-01';
  const tasks = [
    { id: 1, name: 'A1: Initial Setup', dep: '', start: base, end: addDays(base, 2), dur: 3, owner: 'Team A', cat: 'Setup' },
    { id: 2, name: 'B1: Infrastructure', dep: '1', start: addDays(base, 3), end: addDays(base, 7), dur: 5, owner: 'Team B', cat: 'Infra', remarks: 'CRITICAL -- longest parallel task after A1' },
    { id: 3, name: 'A2: Core Module', dep: '1,2', start: addDays(base, 8), end: addDays(base, 11), dur: 4, owner: 'Team A', cat: 'Dev', remarks: 'CRITICAL -- waits for both A1 and B1' },
    { id: 4, name: 'B2: Data Pipeline', dep: '2,3', start: addDays(base, 12), end: addDays(base, 14), dur: 3, owner: 'Team B', cat: 'Data', remarks: 'CRITICAL' },
    { id: 5, name: 'A3: UI Polish', dep: '3', start: addDays(base, 12), end: addDays(base, 13), dur: 2, owner: 'Team A', cat: 'UI', remarks: 'Has 1 day of slack' },
    { id: 6, name: 'Merge & Deploy', dep: '4,5', start: addDays(base, 15), end: addDays(base, 15), dur: 1, owner: 'DevOps', cat: 'Deploy', remarks: 'CRITICAL' },
  ];
  save('15-cpm-ladder-dependency.xlsx', buildWorkbook('CPM Ladder Dependency', tasks));
}

// ---------------------------------------------------------------------------
// Sample 16: Near-Critical Paths (multiple paths with very small slack differences)
// This tests that the CPM highlights the EXACT critical path and shows
// correct slack values for near-critical tasks.
//
//   Start(1d) --> A(10d) --> End(1d)            total=12  CRITICAL
//   Start(1d) --> B(9d)  --> End(1d)            total=11  slack=1
//   Start(1d) --> C(8d)  --> End(1d)            total=10  slack=2
//   Start(1d) --> D(7d)  --> End(1d)            total=9   slack=3
// ---------------------------------------------------------------------------
function genNearCritical() {
  const base = '2026-07-13';
  const tasks = [
    { id: 1, name: 'Project Start', dep: '', start: base, end: base, dur: 1, owner: 'PM', cat: 'Kickoff' },
    { id: 2, name: 'Track A: Server Migration', dep: '1', start: addDays(base, 1), end: addDays(base, 10), dur: 10, owner: 'Alice', cat: 'Infra',
      remarks: 'CRITICAL -- zero slack' },
    { id: 3, name: 'Track B: Database Upgrade', dep: '1', start: addDays(base, 1), end: addDays(base, 9), dur: 9, owner: 'Bob', cat: 'Database',
      remarks: 'Near-critical -- 1 day slack' },
    { id: 4, name: 'Track C: Network Config', dep: '1', start: addDays(base, 1), end: addDays(base, 8), dur: 8, owner: 'Carol', cat: 'Network',
      remarks: '2 days slack' },
    { id: 5, name: 'Track D: Security Audit', dep: '1', start: addDays(base, 1), end: addDays(base, 7), dur: 7, owner: 'Dave', cat: 'Security',
      remarks: '3 days slack' },
    { id: 6, name: 'Final Verification', dep: '2,3,4,5', start: addDays(base, 11), end: addDays(base, 11), dur: 1, owner: 'PM', cat: 'Verification',
      remarks: 'CRITICAL -- waits for all tracks' },
  ];
  save('16-cpm-near-critical.xlsx', buildWorkbook('CPM Near-Critical Paths', tasks));
}

// ---------------------------------------------------------------------------
// Sample 17: Long Project with Mixed Patterns
// A realistic software project with 25 tasks, WBS, milestones, baselines,
// multiple critical and non-critical paths, various statuses and owners.
// ---------------------------------------------------------------------------
function genLongProject() {
  const base = '2026-04-06';
  const tasks = [
    // Phase 1: Discovery
    { id: 10, name: 'Phase 1: Discovery', dep: '', start: base, end: addDays(base, 14), dur: 15, parentId: '', cat: 'Discovery' },
    { id: 11, name: 'Market Research', dep: '', start: base, end: addDays(base, 6), dur: 7, owner: 'Alice', parentId: '10', cat: 'Research',
      blStart: base, blEnd: addDays(base, 6), status: 'Completed', progress: 100, remarks: 'CRITICAL' },
    { id: 12, name: 'User Interviews', dep: '11', start: addDays(base, 7), end: addDays(base, 11), dur: 5, owner: 'Bob', parentId: '10', cat: 'Research',
      blStart: addDays(base, 7), blEnd: addDays(base, 11), status: 'Completed', progress: 100, remarks: 'CRITICAL' },
    { id: 13, name: 'Competitive Analysis', dep: '11', start: addDays(base, 7), end: addDays(base, 9), dur: 3, owner: 'Carol', parentId: '10', cat: 'Research',
      blStart: addDays(base, 7), blEnd: addDays(base, 8), remarks: 'Slack = 2 days' },
    { id: 14, name: 'Requirements Doc', dep: '12,13', start: addDays(base, 12), end: addDays(base, 14), dur: 3, owner: 'Alice', parentId: '10', cat: 'Planning',
      blStart: addDays(base, 12), blEnd: addDays(base, 14), status: 'Completed', progress: 100, remarks: 'CRITICAL' },

    // M1
    { id: 19, name: 'M1: Discovery Complete', dep: '14', start: addDays(base, 15), end: addDays(base, 15), dur: 0, cat: 'Milestone', status: 'Completed', progress: 100 },

    // Phase 2: Design
    { id: 20, name: 'Phase 2: Design', dep: '', start: addDays(base, 15), end: addDays(base, 29), dur: 15, parentId: '', cat: 'Design' },
    { id: 21, name: 'UX Wireframes', dep: '19', start: addDays(base, 15), end: addDays(base, 21), dur: 7, owner: 'Diana', parentId: '20', cat: 'UX',
      blStart: addDays(base, 15), blEnd: addDays(base, 20), status: 'In Progress', progress: 60, remarks: 'CRITICAL -- baseline was 6d, actual 7d' },
    { id: 22, name: 'Visual Design', dep: '21', start: addDays(base, 22), end: addDays(base, 26), dur: 5, owner: 'Eve', parentId: '20', cat: 'UI',
      blStart: addDays(base, 21), blEnd: addDays(base, 25), remarks: 'CRITICAL' },
    { id: 23, name: 'Design System', dep: '21', start: addDays(base, 22), end: addDays(base, 24), dur: 3, owner: 'Frank', parentId: '20', cat: 'UI',
      blStart: addDays(base, 21), blEnd: addDays(base, 23), remarks: 'Slack = 2 days' },
    { id: 24, name: 'Prototype', dep: '22,23', start: addDays(base, 27), end: addDays(base, 29), dur: 3, owner: 'Diana', parentId: '20', cat: 'UX',
      blStart: addDays(base, 26), blEnd: addDays(base, 28), remarks: 'CRITICAL' },

    // M2
    { id: 29, name: 'M2: Design Approved', dep: '24', start: addDays(base, 30), end: addDays(base, 30), dur: 0, cat: 'Milestone' },

    // Phase 3: Development
    { id: 30, name: 'Phase 3: Development', dep: '', start: addDays(base, 30), end: addDays(base, 54), dur: 25, parentId: '', cat: 'Development' },
    { id: 31, name: 'Backend API v1', dep: '29', start: addDays(base, 30), end: addDays(base, 39), dur: 10, owner: 'Grace', parentId: '30', cat: 'Backend',
      blStart: addDays(base, 29), blEnd: addDays(base, 38), remarks: 'CRITICAL' },
    { id: 32, name: 'Frontend Shell', dep: '29', start: addDays(base, 30), end: addDays(base, 36), dur: 7, owner: 'Hank', parentId: '30', cat: 'Frontend',
      blStart: addDays(base, 29), blEnd: addDays(base, 35), remarks: 'Slack = 3 days' },
    { id: 33, name: 'Auth Module', dep: '31', start: addDays(base, 40), end: addDays(base, 44), dur: 5, owner: 'Grace', parentId: '30', cat: 'Backend',
      blStart: addDays(base, 39), blEnd: addDays(base, 43), remarks: 'CRITICAL' },
    { id: 34, name: 'Frontend Pages', dep: '32', start: addDays(base, 37), end: addDays(base, 44), dur: 8, owner: 'Hank', parentId: '30', cat: 'Frontend',
      blStart: addDays(base, 36), blEnd: addDays(base, 43), remarks: 'Non-critical -- joins at integration' },
    { id: 35, name: 'Integration', dep: '33,34', start: addDays(base, 45), end: addDays(base, 49), dur: 5, owner: 'Ivan', parentId: '30', cat: 'Integration',
      blStart: addDays(base, 44), blEnd: addDays(base, 48), remarks: 'CRITICAL' },
    { id: 36, name: 'Performance Tuning', dep: '35', start: addDays(base, 50), end: addDays(base, 54), dur: 5, owner: 'Judy', parentId: '30', cat: 'DevOps',
      blStart: addDays(base, 49), blEnd: addDays(base, 53), remarks: 'CRITICAL' },

    // M3
    { id: 39, name: 'M3: Code Freeze', dep: '36', start: addDays(base, 55), end: addDays(base, 55), dur: 0, cat: 'Milestone' },

    // Phase 4: QA & Launch
    { id: 40, name: 'Phase 4: QA & Launch', dep: '', start: addDays(base, 55), end: addDays(base, 64), dur: 10, parentId: '', cat: 'QA' },
    { id: 41, name: 'Regression Testing', dep: '39', start: addDays(base, 55), end: addDays(base, 59), dur: 5, owner: 'Karen', parentId: '40', cat: 'QA',
      blStart: addDays(base, 54), blEnd: addDays(base, 58), remarks: 'CRITICAL' },
    { id: 42, name: 'Security Pen Test', dep: '39', start: addDays(base, 55), end: addDays(base, 57), dur: 3, owner: 'Leo', parentId: '40', cat: 'Security',
      blStart: addDays(base, 54), blEnd: addDays(base, 56), remarks: 'Slack = 2 days' },
    { id: 43, name: 'Bug Fixes', dep: '41,42', start: addDays(base, 60), end: addDays(base, 62), dur: 3, owner: 'Grace', parentId: '40', cat: 'Dev',
      blStart: addDays(base, 59), blEnd: addDays(base, 61), remarks: 'CRITICAL' },
    { id: 44, name: 'Deployment', dep: '43', start: addDays(base, 63), end: addDays(base, 64), dur: 2, owner: 'Ivan', parentId: '40', cat: 'DevOps',
      blStart: addDays(base, 62), blEnd: addDays(base, 63), remarks: 'CRITICAL' },

    // Final
    { id: 49, name: 'M4: Launch!', dep: '44', start: addDays(base, 65), end: addDays(base, 65), dur: 0, cat: 'Milestone' },
  ];
  save('17-cpm-long-project.xlsx', buildWorkbook('CPM Long Software Project', tasks));
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
console.log('Generating CPM/Slack test Excel files...\n');

genLinearChain();
genParallelPaths();
genDiamondDependency();
genIndependentChains();
genComplexProject();
genLadderDependency();
genNearCritical();
genLongProject();

console.log('\nDone! Files written to excel-template/');
