const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const resultsPath = path.join(repoRoot, 'execution_results.csv');
const evidenceDir = path.join(repoRoot, 'test-results');
const evidencePath = path.join(evidenceDir, 'resize-preview.png');
const fixturePath = path.join(repoRoot, 'fixtures', 'sample-preview.png');

function csvEscape(value) {
  const text = String(value ?? '');
  return `"${text.replace(/"/g, '""')}"`;
}

function writeExecutionResult(row) {
  const headers = [
    'TC ID',
    'Application Feature Tested',
    'Scenario',
    'Input',
    'Expected Output',
    'Actual Output',
    'Status',
    'Execution Date Time',
    'Evidence'
  ];
  const lines = [
    headers.map(csvEscape).join(','),
    headers.map((header) => csvEscape(row[header])).join(',')
  ];
  fs.writeFileSync(resultsPath, `${lines.join('\n')}\n`, 'utf8');
}

test('Pos_0001 - Resize Image preview displays uploaded PNG', async ({ page }) => {
  const startedAt = new Date();
  let actualOutput = 'The test did not complete.';
  let status = 'Fail';

  try {
    await page.goto('/resize-image', { waitUntil: 'networkidle' });
    await expect(page.getByText('Resize Image').first()).toBeVisible();

    await page.locator('input[type="file"]').setInputFiles(fixturePath);

    await expect(page.getByText(/Original:\s*320.180/)).toBeVisible();
    const canvas = page.locator('canvas');
    await expect(canvas).toBeVisible();

    await page.waitForFunction(() => {
      const element = document.querySelector('canvas');
      if (!element) return false;
      const context = element.getContext('2d');
      const pixel = context.getImageData(10, 10, 1, 1).data;
      return element.width === 320 && element.height === 180 && pixel[3] > 0;
    });

    fs.mkdirSync(evidenceDir, { recursive: true });
    await page.screenshot({ path: evidencePath, fullPage: true });

    actualOutput = 'The uploaded PNG was rendered in the Preview canvas with 320x180 dimensions.';
    status = 'Pass';
  } catch (error) {
    actualOutput = error.message.replace(/\s+/g, ' ').slice(0, 500);
    throw error;
  } finally {
    writeExecutionResult({
      'TC ID': 'Pos_0001',
      'Application Feature Tested': 'Image resizing',
      Scenario: 'Verify that a valid PNG upload is displayed in the Resize Image preview area.',
      Input: 'fixtures/sample-preview.png',
      'Expected Output': 'The uploaded PNG should appear in the Preview section with the original image dimensions.',
      'Actual Output': actualOutput,
      Status: status,
      'Execution Date Time': startedAt.toISOString(),
      Evidence: 'test-results/resize-preview.png'
    });
  }
});
