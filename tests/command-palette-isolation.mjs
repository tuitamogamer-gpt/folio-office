import { chromium, expect } from '@playwright/test';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
page.on('pageerror', error => errors.push(error.message));
await page.addInitScript(() => {
  localStorage.setItem('folio-files-v1', JSON.stringify([{
    id: 'keyboard-isolation-deck', name: 'Keyboard isolation deck', kind: 'presentation',
    content: [{ id: 'slide-1', title: 'Keyboard isolation', body: '', layout: 'blank', background: '#ffffff',
      elements: [{ id: 'shape-1', type: 'shape', shape: 'rectangle', x: 10, y: 10, width: 30, height: 30, fill: '#bde1d0' }],
    }], createdAt: 1, updatedAt: 1, starred: false, trashed: false,
  }]));
});
const shape = page.locator('.slides-canvas [data-element-id="shape-1"]');
const query = page.getByRole('combobox', { name: 'Search files and actions' });
const dialog = page.getByRole('dialog', { name: 'Search files and commands' });
async function openPalette() {
  await page.keyboard.press('Control+k');
  await expect(query).toBeFocused();
  await expect(page.locator('.editor-shell')).toHaveJSProperty('inert', true);
}

try {
  await page.goto(process.env.BASE_URL || 'http://localhost:5173');
  await page.locator('.file-name').filter({ hasText: 'Keyboard isolation deck' }).click();
  await page.locator('.slides-layer-row button').click();
  await page.getByLabel('Object x', { exact: true }).fill('12');
  await expect(shape).toHaveCSS('left', /.+/);
  expect(await shape.evaluate(element => element.style.left)).toBe('12%');

  await openPalette();
  await query.fill('workspace');
  await query.press('Control+z');
  expect(await shape.evaluate(element => element.style.left)).toBe('12%');
  await query.press('Control+y');
  expect(await shape.evaluate(element => element.style.left)).toBe('12%');
  await query.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.editor-shell')).toHaveJSProperty('inert', false);

  // Establish a real redo entry, then ensure editing a command search does not consume it.
  await page.keyboard.press('Control+z');
  expect(await shape.evaluate(element => element.style.left)).toBe('10%');
  await openPalette();
  await query.fill('find');
  await query.press('Control+y');
  expect(await shape.evaluate(element => element.style.left)).toBe('10%');
  await query.press('Escape');
  await page.keyboard.press('Control+y');
  expect(await shape.evaluate(element => element.style.left)).toBe('12%');

  await openPalette();
  await query.fill('');
  const close = page.getByRole('button', { name: 'Close command menu' });
  await close.focus();
  await expect(close).toBeFocused();
  const geometry = await shape.evaluate(element => ({ left: element.style.left, top: element.style.top }));
  for (const key of ['ArrowLeft', 'ArrowUp', 'ArrowRight', 'ArrowDown', 'Delete', 'Backspace', 'Control+z', 'Control+y']) {
    await page.keyboard.press(key);
    await expect(shape).toHaveCount(1);
    expect(await shape.evaluate(element => ({ left: element.style.left, top: element.style.top }))).toEqual(geometry);
    await expect(dialog).toBeVisible();
  }
  await page.keyboard.press('Control+k');
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.editor-shell')).toHaveJSProperty('inert', false);

  // Executing a command that opens another dialog must transfer focus into it.
  for (const [command, title, closeName] of [
    ['Workspace settings', 'Make this space yours.', 'Close dialog'],
    ['Version history', 'Version history', 'Close version history'],
  ]) {
    await page.getByLabel('Object x', { exact: true }).focus();
    await openPalette();
    await query.fill(command);
    await query.press('Enter');
    const nextDialog = page.getByRole('dialog', { name: title, exact: true });
    await expect(dialog).toHaveCount(0);
    await expect(nextDialog).toBeVisible();
    await expect.poll(() => nextDialog.evaluate(element => element.contains(document.activeElement))).toBe(true);
    await expect(page.locator('.editor-shell')).toHaveJSProperty('inert', true);
    const modalClose = page.getByRole('button', { name: closeName, exact: true });
    await modalClose.focus();
    await page.keyboard.press('Shift+Tab');
    expect(await nextDialog.evaluate(element => element.contains(document.activeElement))).toBe(true);
    await page.keyboard.press('Tab');
    await expect(modalClose).toBeFocused();
    await page.keyboard.press('Control+z');
    await page.keyboard.press('Delete');
    await expect(shape).toHaveCount(1);
    expect(await shape.evaluate(element => ({ left: element.style.left, top: element.style.top }))).toEqual(geometry);
    await modalClose.click();
    await expect(nextDialog).toHaveCount(0);
    await expect(page.locator('.editor-shell')).toHaveJSProperty('inert', false);
  }
  expect(errors).toEqual([]);
  console.log('PASS: command search Ctrl+Z/Y leaves slide history intact; palette Close Delete/Backspace/arrows/undo/redo cannot mutate selected objects; background is inert; Ctrl+K closes the palette; settings/history receive and trap focus.');
} finally {
  await browser.close();
}
