/**
 * Generates example Excel templates for GanttGen.
 * Each file demonstrates a different project management scenario.
 *
 * Run:  node scripts/generate-templates.mjs
 */

import * as XLSX from 'xlsx';
import { mkdirSync, writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = resolve(__dirname, '..', 'excel-template');

mkdirSync(OUT_DIR, { recursive: true });

const TASK_COLUMNS = [
  'Task ID',
  'Task Name',
  'Dependency',
  'Task Category',
  'Start Date',
  'End Date',
  'Duration',
  'Progress (%)',
  'Status',
  'Owner',
  'Remarks',
  'Baseline Start',
  'Baseline End',
  'Parent ID',
];

const SETTINGS_FIELDS = [
  { label: 'Theme Name', value: 'Linear Dark' },
  { label: 'Theme BG Primary', value: '#0f0f12' },
  { label: 'Theme BG Secondary', value: '#1a1a23' },
  { label: 'Theme BG Tertiary', value: '#23232f' },
  { label: 'Theme BG Hover', value: '#2a2a38' },
  { label: 'Theme Border', value: '#2e2e3a' },
  { label: 'Theme Border Subtle', value: '#232330' },
  { label: 'Theme Text Primary', value: '#e8e8ed' },
  { label: 'Theme Text Secondary', value: '#9898a8' },
  { label: 'Theme Text Muted', value: '#66667a' },
  { label: 'Theme Accent', value: '#6366f1' },
  { label: 'Theme Accent Hover', value: '#818cf8' },
  { label: 'Theme Success', value: '#22c55e' },
  { label: 'Theme Warning', value: '#f59e0b' },
  { label: 'Theme Danger', value: '#ef4444' },
  { label: 'Theme Info', value: '#3b82f6' },
  { label: 'Theme Critical Path', value: '#f43f5e' },
  { label: 'Show Critical Path', value: 'true' },
  { label: 'Show Slack', value: 'true' },
  { label: 'Show Dependencies', value: 'true' },
  { label: 'Show Today Line', value: 'true' },
  { label: 'Show Baseline', value: 'true' },
  { label: 'Skip Weekends', value: 'true' },
  { label: 'Visible Columns', value: 'id,name,duration,startDate,endDate,progress,status' },
];

function buildWorkbook(taskRows, settingsOverrides = {}) {
  const wb = XLSX.utils.book_new();

  const data = [TASK_COLUMNS, ...taskRows];
  const taskSheet = XLSX.utils.aoa_to_sheet(data);
  taskSheet['!cols'] = TASK_COLUMNS.map((col) => ({
    wch: Math.max(col.length + 2, 14),
  }));
  XLSX.utils.book_append_sheet(wb, taskSheet, 'Tasks');

  const merged = SETTINGS_FIELDS.map((f) => {
    const override = settingsOverrides[f.label];
    return [f.label, override !== undefined ? override : f.value];
  });
  merged.unshift(['Setting', 'Value']);
  const settingsSheet = XLSX.utils.aoa_to_sheet(merged);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  return wb;
}

function writeTemplate(filename, taskRows, settingsOverrides = {}) {
  const wb = buildWorkbook(taskRows, settingsOverrides);
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  const path = resolve(OUT_DIR, filename);
  writeFileSync(path, buf);
  console.log(`  Created: ${filename}`);
}

// ---------------------------------------------------------------------------
// 1. Simple Linear Project -- sequential tasks with finish-to-start deps
// ---------------------------------------------------------------------------
writeTemplate('01-simple-linear-project.xlsx', [
  // ID, Name, Dep, Category, Start, End, Dur, Prog, Status, Owner, Remarks, BL Start, BL End, Parent
  ['1', 'Project Kickoff', '', 'Planning', '2026-05-04', '2026-05-08', 5, 100, 'Completed', 'Alice', 'Initial planning and alignment', '', '', ''],
  ['2', 'Requirements Gathering', '1', 'Planning', '2026-05-11', '2026-05-22', 10, 80, 'In Progress', 'Bob', 'Stakeholder interviews', '', '', ''],
  ['3', 'System Design', '2', 'Design', '2026-05-25', '2026-06-05', 10, 0, 'Not Started', 'Carol', 'Architecture documentation', '', '', ''],
  ['4', 'Backend Development', '3', 'Development', '2026-06-08', '2026-07-03', 20, 0, 'Not Started', 'Dave', 'API and database layer', '', '', ''],
  ['5', 'Frontend Development', '4', 'Development', '2026-07-06', '2026-07-24', 15, 0, 'Not Started', 'Eve', 'UI implementation', '', '', ''],
  ['6', 'Testing', '5', 'QA', '2026-07-27', '2026-08-07', 10, 0, 'Not Started', 'Frank', 'Integration and UAT', '', '', ''],
  ['7', 'Deployment', '6', 'Delivery', '2026-08-10', '2026-08-14', 5, 0, 'Not Started', 'Alice', 'Production rollout', '', '', ''],
]);

// ---------------------------------------------------------------------------
// 2. Parallel Workstreams -- independent tasks running concurrently
// ---------------------------------------------------------------------------
writeTemplate('02-parallel-workstreams.xlsx', [
  ['1', 'Sprint Planning', '', 'Planning', '2026-05-04', '2026-05-08', 5, 100, 'Completed', 'PM', '', '', '', ''],
  ['2', 'API Module A', '1', 'Backend', '2026-05-11', '2026-05-29', 15, 60, 'In Progress', 'Alice', 'User service', '', '', ''],
  ['3', 'API Module B', '1', 'Backend', '2026-05-11', '2026-06-05', 20, 40, 'In Progress', 'Bob', 'Order service', '', '', ''],
  ['4', 'API Module C', '1', 'Backend', '2026-05-11', '2026-05-22', 10, 90, 'In Progress', 'Carol', 'Notification service', '', '', ''],
  ['5', 'UI Component Library', '1', 'Frontend', '2026-05-11', '2026-06-05', 20, 50, 'In Progress', 'Dave', 'Shared components', '', '', ''],
  ['6', 'Integration Testing', '2,3,4,5', 'QA', '2026-06-08', '2026-06-19', 10, 0, 'Not Started', 'Eve', 'All modules must complete', '', '', ''],
  ['7', 'Performance Tuning', '6', 'DevOps', '2026-06-22', '2026-06-26', 5, 0, 'Not Started', 'Frank', '', '', '', ''],
  ['8', 'Release', '7', 'Delivery', '2026-06-29', '2026-07-03', 5, 0, 'Not Started', 'PM', '', '', '', ''],
]);

// ---------------------------------------------------------------------------
// 3. WBS Hierarchy -- parent/child task structure
// ---------------------------------------------------------------------------
writeTemplate('03-wbs-hierarchy.xlsx', [
  ['1', 'Website Redesign', '', 'Summary', '2026-05-04', '2026-08-28', 85, 25, 'In Progress', 'PM', 'Top-level project', '', '', ''],
  ['1.1', 'Research Phase', '', 'Summary', '2026-05-04', '2026-05-22', 15, 100, 'Completed', 'PM', '', '', '', '1'],
  ['1.1.1', 'Competitor Analysis', '', 'Research', '2026-05-04', '2026-05-15', 10, 100, 'Completed', 'Alice', '', '', '', '1.1'],
  ['1.1.2', 'User Surveys', '', 'Research', '2026-05-11', '2026-05-22', 10, 100, 'Completed', 'Bob', '', '', '', '1.1'],
  ['1.2', 'Design Phase', '1.1', 'Summary', '2026-05-25', '2026-06-26', 25, 40, 'In Progress', 'PM', '', '', '', '1'],
  ['1.2.1', 'Wireframes', '1.1', 'Design', '2026-05-25', '2026-06-05', 10, 100, 'Completed', 'Carol', '', '', '', '1.2'],
  ['1.2.2', 'Visual Design', '1.2.1', 'Design', '2026-06-08', '2026-06-19', 10, 30, 'In Progress', 'Carol', '', '', '', '1.2'],
  ['1.2.3', 'Prototype Review', '1.2.2', 'Design', '2026-06-22', '2026-06-26', 5, 0, 'Not Started', 'PM', '', '', '', '1.2'],
  ['1.3', 'Development Phase', '1.2', 'Summary', '2026-06-29', '2026-08-14', 35, 0, 'Not Started', 'PM', '', '', '', '1'],
  ['1.3.1', 'CMS Setup', '1.2', 'Development', '2026-06-29', '2026-07-10', 10, 0, 'Not Started', 'Dave', '', '', '', '1.3'],
  ['1.3.2', 'Page Templates', '1.3.1', 'Development', '2026-07-13', '2026-07-31', 15, 0, 'Not Started', 'Eve', '', '', '', '1.3'],
  ['1.3.3', 'Content Migration', '1.3.2', 'Development', '2026-08-03', '2026-08-14', 10, 0, 'Not Started', 'Frank', '', '', '', '1.3'],
  ['1.4', 'Launch', '1.3', 'Summary', '2026-08-17', '2026-08-28', 10, 0, 'Not Started', 'PM', '', '', '', '1'],
  ['1.4.1', 'QA and Bug Fixes', '1.3', 'QA', '2026-08-17', '2026-08-21', 5, 0, 'Not Started', 'Eve', '', '', '', '1.4'],
  ['1.4.2', 'Go Live', '1.4.1', 'Delivery', '2026-08-24', '2026-08-28', 5, 0, 'Not Started', 'PM', '', '', '', '1.4'],
]);

// ---------------------------------------------------------------------------
// 4. Milestones and Baselines -- zero-duration milestones with baseline tracking
// ---------------------------------------------------------------------------
writeTemplate('04-milestones-and-baselines.xlsx', [
  ['1', 'Project Charter Approved', '', 'Milestone', '2026-05-04', '2026-05-04', 0, 100, 'Completed', 'PM', 'Milestone', '2026-05-01', '2026-05-01', ''],
  ['2', 'Requirements Analysis', '1', 'Analysis', '2026-05-05', '2026-05-16', 10, 100, 'Completed', 'Alice', '', '2026-05-05', '2026-05-16', ''],
  ['3', 'Requirements Signed Off', '2', 'Milestone', '2026-05-19', '2026-05-19', 0, 100, 'Completed', 'PM', 'Milestone', '2026-05-16', '2026-05-16', ''],
  ['4', 'Architecture Design', '3', 'Design', '2026-05-19', '2026-06-06', 15, 70, 'In Progress', 'Bob', 'Slipped 3 days vs baseline', '2026-05-19', '2026-05-30', ''],
  ['5', 'Design Review Gate', '4', 'Milestone', '2026-06-09', '2026-06-09', 0, 0, 'Not Started', 'PM', 'Milestone', '2026-06-02', '2026-06-02', ''],
  ['6', 'Core Development', '5', 'Development', '2026-06-09', '2026-07-18', 30, 0, 'Not Started', 'Carol', '', '2026-06-02', '2026-07-11', ''],
  ['7', 'Code Freeze', '6', 'Milestone', '2026-07-20', '2026-07-20', 0, 0, 'Not Started', 'PM', 'Milestone', '2026-07-13', '2026-07-13', ''],
  ['8', 'System Testing', '7', 'QA', '2026-07-20', '2026-08-07', 15, 0, 'Not Started', 'Dave', '', '2026-07-13', '2026-07-25', ''],
  ['9', 'UAT Complete', '8', 'Milestone', '2026-08-10', '2026-08-10', 0, 0, 'Not Started', 'PM', 'Milestone', '2026-07-28', '2026-07-28', ''],
  ['10', 'Production Deployment', '9', 'Delivery', '2026-08-10', '2026-08-14', 5, 0, 'Not Started', 'Eve', '', '2026-07-28', '2026-08-01', ''],
  ['11', 'Project Closure', '10', 'Milestone', '2026-08-17', '2026-08-17', 0, 0, 'Not Started', 'PM', 'Milestone', '2026-08-04', '2026-08-04', ''],
], {
  'Show Baseline': 'true',
  'Visible Columns': 'id,name,duration,startDate,endDate,progress,status,owner',
});

// ---------------------------------------------------------------------------
// 5. Critical Path with Slack -- demonstrates CPM analysis
// ---------------------------------------------------------------------------
writeTemplate('05-critical-path-demo.xlsx', [
  ['1', 'Foundation Work', '', 'Construction', '2026-05-04', '2026-05-22', 15, 100, 'Completed', 'Team A', 'Critical: no slack', '', '', ''],
  ['2', 'Structural Framing', '1', 'Construction', '2026-05-25', '2026-06-19', 20, 50, 'In Progress', 'Team A', 'Critical: no slack', '', '', ''],
  ['3', 'Electrical Rough-In', '2', 'MEP', '2026-06-22', '2026-07-03', 10, 0, 'Not Started', 'Team B', 'Critical path continues', '', '', ''],
  ['4', 'Plumbing Rough-In', '2', 'MEP', '2026-06-22', '2026-06-26', 5, 0, 'Not Started', 'Team C', 'Has slack: shorter than electrical', '', '', ''],
  ['5', 'HVAC Installation', '2', 'MEP', '2026-06-22', '2026-07-03', 10, 0, 'Not Started', 'Team D', 'Parallel with electrical', '', '', ''],
  ['6', 'Insulation', '3,4,5', 'Construction', '2026-07-06', '2026-07-10', 5, 0, 'Not Started', 'Team A', 'Waits for all MEP', '', '', ''],
  ['7', 'Drywall', '6', 'Construction', '2026-07-13', '2026-07-24', 10, 0, 'Not Started', 'Team E', 'Critical', '', '', ''],
  ['8', 'Painting', '7', 'Finish', '2026-07-27', '2026-08-07', 10, 0, 'Not Started', 'Team F', 'Critical', '', '', ''],
  ['9', 'Flooring', '7', 'Finish', '2026-07-27', '2026-07-31', 5, 0, 'Not Started', 'Team G', 'Has slack: shorter than painting', '', '', ''],
  ['10', 'Landscaping', '1', 'Exterior', '2026-05-25', '2026-06-05', 10, 30, 'In Progress', 'Team H', 'Independent: large slack', '', '', ''],
  ['11', 'Final Inspection', '8,9,10', 'QA', '2026-08-10', '2026-08-14', 5, 0, 'Not Started', 'Inspector', 'Converging dependencies', '', '', ''],
], {
  'Show Critical Path': 'true',
  'Show Slack': 'true',
  'Show Dependencies': 'true',
});

// ---------------------------------------------------------------------------
// 6. Software Sprint -- agile-style iteration with categories and owners
// ---------------------------------------------------------------------------
writeTemplate('06-software-sprint.xlsx', [
  ['1', 'Sprint 12 Planning', '', 'Ceremony', '2026-05-04', '2026-05-04', 1, 100, 'Completed', 'Scrum Master', '', '', '', ''],
  ['2', 'User Auth - OAuth2 Flow', '1', 'Feature', '2026-05-05', '2026-05-09', 5, 100, 'Completed', 'Alice', '', '', '', ''],
  ['3', 'User Auth - Session Mgmt', '2', 'Feature', '2026-05-12', '2026-05-14', 3, 80, 'In Progress', 'Alice', '', '', '', ''],
  ['4', 'Dashboard Charts', '1', 'Feature', '2026-05-05', '2026-05-14', 8, 60, 'In Progress', 'Bob', '', '', '', ''],
  ['5', 'API Rate Limiting', '1', 'Tech Debt', '2026-05-05', '2026-05-07', 3, 100, 'Completed', 'Carol', '', '', '', ''],
  ['6', 'Fix: Memory Leak in WS', '1', 'Bug', '2026-05-05', '2026-05-06', 2, 100, 'Completed', 'Dave', 'Hotfix priority', '', '', ''],
  ['7', 'E2E Test Suite Expansion', '2,5', 'QA', '2026-05-12', '2026-05-16', 5, 20, 'In Progress', 'Eve', '', '', '', ''],
  ['8', 'Database Index Optimization', '1', 'Tech Debt', '2026-05-05', '2026-05-09', 5, 50, 'In Progress', 'Frank', '', '', '', ''],
  ['9', 'Code Review and Merge', '3,4,7,8', 'QA', '2026-05-19', '2026-05-20', 2, 0, 'Not Started', 'Lead', '', '', '', ''],
  ['10', 'Sprint 12 Retrospective', '9', 'Ceremony', '2026-05-21', '2026-05-21', 1, 0, 'Not Started', 'Scrum Master', '', '', '', ''],
], {
  'Visible Columns': 'id,name,category,duration,startDate,endDate,progress,status,owner',
});

// ---------------------------------------------------------------------------
// 7. Event Planning -- non-software scenario with varied statuses
// ---------------------------------------------------------------------------
writeTemplate('07-event-planning.xlsx', [
  ['1', 'Define Event Concept', '', 'Strategy', '2026-05-04', '2026-05-08', 5, 100, 'Completed', 'Director', '', '', '', ''],
  ['2', 'Secure Venue', '1', 'Logistics', '2026-05-11', '2026-05-22', 10, 100, 'Completed', 'Logistics Lead', 'Contract signed', '', '', ''],
  ['3', 'Book Speakers', '1', 'Content', '2026-05-11', '2026-06-05', 20, 60, 'In Progress', 'Content Lead', '4 of 6 confirmed', '', '', ''],
  ['4', 'Marketing Campaign', '2', 'Marketing', '2026-05-25', '2026-07-03', 30, 30, 'In Progress', 'Marketing Lead', 'Social media + email', '', '', ''],
  ['5', 'Sponsorship Outreach', '1', 'Finance', '2026-05-11', '2026-06-19', 30, 70, 'In Progress', 'Biz Dev', '3 sponsors locked', '', '', ''],
  ['6', 'Catering Selection', '2', 'Logistics', '2026-05-25', '2026-06-05', 10, 0, 'Not Started', 'Logistics Lead', '', '', '', ''],
  ['7', 'AV Equipment Setup', '2', 'Logistics', '2026-07-06', '2026-07-10', 5, 0, 'Not Started', 'AV Vendor', '', '', '', ''],
  ['8', 'Rehearsal Day', '3,6,7', 'Logistics', '2026-07-13', '2026-07-13', 1, 0, 'Not Started', 'Director', '', '', '', ''],
  ['9', 'Event Day', '8', 'Execution', '2026-07-14', '2026-07-14', 1, 0, 'Not Started', 'All Teams', '', '', '', ''],
  ['10', 'Post-Event Report', '9', 'Wrap-up', '2026-07-15', '2026-07-24', 8, 0, 'Not Started', 'Director', '', '', '', ''],
], {
  'Theme Name': 'Notion Light',
  'Theme BG Primary': '#ffffff',
  'Theme BG Secondary': '#f7f7f5',
  'Theme BG Tertiary': '#f0f0ee',
  'Theme BG Hover': '#e8e8e5',
  'Theme Border': '#e0e0dc',
  'Theme Border Subtle': '#ebebea',
  'Theme Text Primary': '#1a1a1a',
  'Theme Text Secondary': '#6b6b6b',
  'Theme Text Muted': '#9b9b9b',
  'Theme Accent': '#2f7df6',
  'Theme Accent Hover': '#528bff',
  'Theme Success': '#0f7b3f',
  'Theme Warning': '#c27800',
  'Theme Danger': '#e03e3e',
  'Theme Info': '#2f7df6',
  'Theme Critical Path': '#e03e3e',
});

// ---------------------------------------------------------------------------
// 8. Delayed Project with Baseline Variance -- schedule slippage scenario
// ---------------------------------------------------------------------------
writeTemplate('08-baseline-variance.xlsx', [
  ['1', 'Requirements', '', 'Analysis', '2026-05-04', '2026-05-15', 10, 100, 'Completed', 'Alice', 'On schedule', '2026-05-04', '2026-05-15', ''],
  ['2', 'Design', '1', 'Design', '2026-05-18', '2026-06-05', 15, 100, 'Completed', 'Bob', 'Slipped 5 days', '2026-05-18', '2026-05-29', ''],
  ['3', 'Dev Sprint 1', '2', 'Development', '2026-06-08', '2026-06-26', 15, 80, 'In Progress', 'Carol', 'Slipped from baseline', '2026-06-01', '2026-06-12', ''],
  ['4', 'Dev Sprint 2', '3', 'Development', '2026-06-29', '2026-07-17', 15, 0, 'Not Started', 'Dave', 'Cascading delay', '2026-06-15', '2026-06-26', ''],
  ['5', 'QA Phase', '4', 'QA', '2026-07-20', '2026-08-07', 15, 0, 'Not Started', 'Eve', 'Compressed from 20 to 15 days', '2026-06-29', '2026-07-24', ''],
  ['6', 'Staging Deploy', '5', 'DevOps', '2026-08-10', '2026-08-14', 5, 0, 'Not Started', 'Frank', '', '2026-07-27', '2026-07-31', ''],
  ['7', 'Go/No-Go Decision', '6', 'Milestone', '2026-08-17', '2026-08-17', 0, 0, 'Not Started', 'PM', 'Milestone', '2026-08-03', '2026-08-03', ''],
  ['8', 'Production Release', '7', 'Delivery', '2026-08-17', '2026-08-21', 5, 0, 'At Risk', 'Ops', 'Deadline is Aug 22', '2026-08-03', '2026-08-07', ''],
], {
  'Show Baseline': 'true',
  'Visible Columns': 'id,name,duration,startDate,endDate,progress,status,owner,remarks',
});

console.log('\nAll templates generated in excel-template/');
