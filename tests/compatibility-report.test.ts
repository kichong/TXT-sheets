import { describe, expect, it } from 'vitest';
import { buildCompatibilityReportUrl } from '../src/shared/compatibility-report';

describe('compatibility report', () => {
  it('creates a sanitized, prefilled GitHub issue', () => {
    const report = new URL(buildCompatibilityReportUrl({
      appName: 'TXT Sheets', appVersion: '0.2.0', operatingSystem: 'win32 10.0.26100',
      sourceFormat: 'xlsx', repositoryUrl: 'https://github.com/kichong/TXT-sheets',
      issues: [{ feature: 'Conditional formatting', detail: 'PRIVATE worksheet name and formula' }],
    }));
    expect(report.origin + report.pathname).toBe('https://github.com/kichong/TXT-sheets/issues/new');
    expect(report.searchParams.get('title')).toContain('Conditional formatting');
    expect(report.searchParams.get('body')).not.toContain('PRIVATE');
    expect(report.searchParams.get('body')).toContain('does not include the filename');
    expect(report.searchParams.get('body')).toContain('Platform: other');
  });
});
