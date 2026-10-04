import { chromium, expect } from '@playwright/test';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
page.on('pageerror', error => errors.push(error.message));
await page.addInitScript(() => {
  if (sessionStorage.getItem('trash-undo-fixture')) return;
  sessionStorage.setItem('trash-undo-fixture', 'true');
  localStorage.setItem('folio-files-v1', JSON.stringify([{
    id: 'preserved-draft', name: 'Preserved draft', kind: 'document', content: '<p>Every paragraph survives.</p>',
    createdAt: 1, updatedAt: 1, starred: false, trashed: false,
    versions: [{ id: 'saved-original', name: 'Saved original', createdAt: 1, content: '<p>Original version remains.</p>' }],
  }]));
});

async function showFiles() {
  await page.locator('.sidebar').getByRole('button', { name: 'My files', exact: false }).click();
  await expect(page.locator('.file-row')).toHaveCount(1);
  await expect(page.locator('.file-name')).toContainText('Preserved draft');
}

try {
  await page.goto(process.env.BASE_URL || 'http://localhost:5173');
  await page.getByRole('checkbox', { name: 'Select Preserved draft', exact: true }).check();
  await page.getByRole('toolbar', { name: 'Selected file actions' }).getByRole('button', { name: 'Move to trash' }).click();
  await page.locator('.sidebar').getByRole('button', { name: 'Trash', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Select all visible files' }).check();
  await page.getByRole('button', { name: 'Delete forever', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Delete selected files', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Delete selected files', exact: true })).toHaveCount(0);
  await showFiles();

  await page.getByRole('button', { name: 'More options for Preserved draft' }).click();
  await page.getByRole('button', { name: 'Move to trash', exact: true }).click();
  await page.locator('.sidebar').getByRole('button', { name: 'Trash', exact: true }).click();
  await page.getByRole('button', { name: 'More options for Preserved draft' }).click();
  await page.getByRole('button', { name: 'Delete permanently', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Delete this file for good?' })).toBeVisible();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await showFiles();
  await page.reload();
  await expect(page.locator('.file-row')).toHaveCount(1);
  await page.getByRole('button', { name: 'More options for Preserved draft' }).click();
  await page.getByRole('button', { name: 'Version history', exact: true }).click();
  await expect(page.locator('.version-item')).toHaveCount(1);
  await expect(page.locator('.version-item')).toContainText('Original version remains.');
  await page.getByRole('button', { name: 'Close version history' }).click();
  await page.locator('.file-name').filter({ hasText: 'Preserved draft' }).click();
  await expect(page.locator('.tiptap')).toHaveText('Every paragraph survives.');
  expect(errors).toEqual([]);
  console.log('PASS: Undo invalidates bulk and single permanent-delete confirmations; restored files, current content, and saved versions survive reload.');
} finally {
  await browser.close();
}
