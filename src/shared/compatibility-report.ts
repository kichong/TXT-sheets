import type { CompatibilityIssue, CompatibilityReportRequest } from './types';

interface CompatibilityReportContext extends CompatibilityReportRequest {
  appName: string;
  appVersion: string;
  operatingSystem: string;
  repositoryUrl: string;
}

function singleLine(value: string): string {
  return value.replace(/\s+/gu, ' ').trim();
}

const SAFE_ISSUES: Record<string, string> = {
  Images: 'Images are not displayed or preserved when the workbook is saved.',
  'Excel tables': 'Table values and styling are shown, but structured-table metadata may be simplified.',
  'Conditional formatting': 'Conditional-formatting rules are not currently evaluated.',
  'Data validation': 'Data-validation rules are not currently editable.',
  'Hidden sheets': 'Hidden worksheets are visible while editing.',
  'Advanced formulas': 'Some formulas use functions that TXT Sheets cannot recalculate yet.',
};

function safeIssues(issues: CompatibilityIssue[]) {
  return [...new Set(issues.map((issue) => issue.feature))]
    .filter((feature) => Object.hasOwn(SAFE_ISSUES, feature))
    .map((feature) => ({ feature, detail: SAFE_ISSUES[feature] }));
}

export function buildCompatibilityReportUrl(context: CompatibilityReportContext): string {
  const issues = safeIssues(context.issues);
  const issueNames = issues.map((issue) => issue.feature);
  const title = `[Compatibility] ${issueNames.slice(0, 2).join(', ')}${issueNames.length > 2 ? ` +${issueNames.length - 2} more` : ''}`;
  const body = [
    '## Compatibility report', '',
    '> Privacy note: this report contains only the app version, general platform, file type, and compatibility categories. It does not include the filename, file path, worksheet names, cell values, formulas, comments, authors, or other workbook metadata.', '',
    `- App: ${singleLine(context.appName)} ${singleLine(context.appVersion)}`,
    `- Platform: ${['win32', 'darwin', 'linux'].includes(context.operatingSystem) ? context.operatingSystem : 'other'}`,
    `- Source format: ${singleLine(context.sourceFormat)}`, '',
    '### Detected issues', '', ...issues.map((issue) => `- **${issue.feature}**: ${issue.detail}`), '',
    '### What happened?', '',
    '<!-- Describe the behavior without pasting private workbook data. A small newly-created sample file is safest if a reproduction is needed. -->', '',
    'Expected:', '', 'Actual:',
  ].join('\n');
  const params = new URLSearchParams({ title, body });
  return `${context.repositoryUrl.replace(/\/$/u, '')}/issues/new?${params.toString()}`;
}
