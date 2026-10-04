import { chromium, expect } from '@playwright/test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
page.on('pageerror', error => errors.push(error.message));
const selectText = async (text, collapse = false) => page.locator('.tiptap').evaluate((node, { text, collapse }) => {
  const editor = node.editor;
  let range;
  editor.state.doc.descendants((child, pos) => {
    if (!range && child.isTextblock && child.textContent.includes(text)) {
      const start = pos + 1 + child.textContent.indexOf(text);
      range = { from: collapse ? start + text.length : start, to: start + text.length };
    }
  });
  if (!range) throw new Error(`Missing text ${text}`);
  editor.chain().focus().setTextSelection(range).run();
}, { text, collapse });
const stat = (name, column = 0) => page.locator(`.doc-statistics [data-stat="${name}"] td`).nth(column);
const selectionHeading = () => page.locator('.tiptap').evaluate(node => node.editor.state.selection.$from.parent.textContent);

try {
  await page.goto(process.env.BASE_URL || 'http://localhost:5173');
  await page.getByRole('button', { name: 'New document', exact: false }).click();
  await expect(page.locator('.tiptap')).toBeVisible();
  await page.getByLabel('Document name').fill('A focused writing space');
  await page.getByLabel('Document name').press('Enter');
  await page.locator('.tiptap').evaluate(node => node.editor.commands.setContent('<h1>Overview</h1><p>Hello <strong>world</strong>! 👩‍💻</p><p></p><blockquote><h2>Nested section</h2><p>Alpha beta.</p></blockquote><h2>Closing</h2><p>Čaj i čokolada.</p><table><tbody><tr><td><p>Cell text</p></td></tr></tbody></table>'));

  await page.getByRole('button', { name: 'Document statistics', exact: true }).click();
  await expect(stat('words')).toHaveText('13');
  await expect(stat('characters')).toHaveText('78');
  await expect(stat('charactersWithoutSpaces')).toHaveText('71');
  await expect(stat('paragraphs')).toHaveText('7');
  await expect(stat('readingMinutes')).toHaveText('< 1 min');
  await expect(stat('words', 1)).toHaveText('—');
  await selectText('Hello world! 👩‍💻');
  await expect(stat('words', 1)).toHaveText('2');
  await expect(stat('characters', 1)).toHaveText('14');
  await expect(stat('charactersWithoutSpaces', 1)).toHaveText('12');
  await expect(stat('paragraphs', 1)).toHaveText('1');
  await expect(page.getByRole('button', { name: 'Document statistics', exact: true })).toHaveText('2 of 13 words');

  await page.locator('.tiptap').evaluate(node => {
    const editor = node.editor;
    let end;
    editor.state.doc.descendants((child, pos) => { if (child.type.name === 'heading' && child.textContent === 'Nested section') end = pos + 1 + child.content.size; });
    editor.chain().focus().setTextSelection({ from: 1, to: end }).run();
  });
  await expect(stat('words', 1)).toHaveText('5');
  await expect(stat('paragraphs', 1)).toHaveText('3');
  await selectText('Čaj i čokolada.', true);
  await page.keyboard.type(' More words.');
  await expect(stat('words')).toHaveText('15');
  await expect(stat('characters')).toHaveText('90');
  await expect(stat('words', 1)).toHaveText('—');
  await page.getByRole('button', { name: 'Close statistics' }).click();

  await page.getByRole('button', { name: 'View', exact: true }).click();
  await page.getByRole('button', { name: 'Outline', exact: true }).click();
  const outline = page.getByRole('navigation', { name: 'Document headings', exact: true });
  await expect(outline.getByRole('button')).toHaveCount(3);
  await outline.getByRole('button', { name: 'Heading 2: Nested section', exact: true }).click();
  assert.equal(await selectionHeading(), 'Nested section');
  await expect(outline.getByRole('button', { name: 'Heading 2: Nested section', exact: true })).toHaveAttribute('aria-current', 'location');
  await selectText('Alpha beta.', true);
  await expect(outline.getByRole('button', { name: 'Heading 2: Nested section', exact: true })).toHaveAttribute('aria-current', 'location');
  await outline.getByRole('button', { name: 'Heading 1: Overview', exact: true }).focus();
  await page.keyboard.press('ArrowDown');
  await expect(outline.getByRole('button', { name: 'Heading 2: Nested section', exact: true })).toBeFocused();
  await page.keyboard.press('End');
  await expect(outline.getByRole('button', { name: 'Heading 2: Closing', exact: true })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(outline.getByRole('button', { name: 'Heading 1: Overview', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  assert.equal(await selectionHeading(), 'Nested section');
  await expect(page.locator('.tiptap')).toBeFocused();
  await page.getByLabel('Filter headings').fill('CLOSING');
  await expect(outline.getByRole('button')).toHaveCount(1);
  await outline.getByRole('button').click();
  assert.equal(await selectionHeading(), 'Closing');
  await page.getByLabel('Filter headings').fill('missing');
  await expect(page.getByLabel('Document outline', { exact: true })).toContainText('No headings match');
  await page.getByLabel('Filter headings').press('Escape');
  await expect(page.getByLabel('Filter headings')).toHaveValue('');
  await expect(outline.getByRole('button')).toHaveCount(3);

  const beforeFocus = await page.locator('.tiptap').evaluate(node => node.editor.getHTML());
  await page.getByRole('button', { name: 'Focus mode', exact: true }).click();
  await expect(page.locator('.doc-ribbon')).toBeHidden();
  await expect(page.getByLabel('Document outline', { exact: true })).toBeHidden();
  await expect(page.getByRole('button', { name: 'Exit focus mode', exact: true })).toBeVisible();
  assert.equal(await page.locator('.tiptap').evaluate(node => node.editor.getHTML()), beforeFocus);
  await selectText('More words.', true);
  await page.keyboard.type(' Writing freely.');
  await expect(page.locator('.tiptap')).toContainText('Writing freely.');
  await page.keyboard.press('Control+f');
  await expect(page.getByLabel('Find text', { exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.locator('.doc-find-panel')).toHaveCount(0);
  await expect(page.locator('.doc-focus-bar')).toBeVisible();
  await page.keyboard.press('Control+Shift+g');
  await expect(page.getByLabel('Writing statistics', { exact: true })).toBeVisible();
  await expect(stat('words')).toHaveText('17');
  await page.keyboard.press('Escape');
  await expect(page.getByLabel('Writing statistics', { exact: true })).toHaveCount(0);
  await expect(page.locator('.doc-focus-bar')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('.doc-focus-bar')).toHaveCount(0);
  await expect(page.getByLabel('Document outline', { exact: true })).toBeVisible();

  await page.getByTitle('Switch to reading').click();
  await page.keyboard.press('Control+Shift+f');
  await expect(page.locator('.doc-focus-bar')).toContainText('Focus · Reading');
  await expect(page.locator('.tiptap')).toHaveAttribute('contenteditable', 'false');
  await page.getByRole('button', { name: 'Document statistics', exact: true }).click();
  await expect(stat('words')).toHaveText('17');
  await page.pdf({ path: '/tmp/folio-writing-focus.pdf', preferCSSPageSize: true });
  const printed = execFileSync('pdftotext', ['-layout', '/tmp/folio-writing-focus.pdf', '-'], { encoding: 'utf8' });
  assert.match(printed, /Overview/); assert.match(printed, /Nested section/); assert.match(printed, /Writing freely/);
  assert.doesNotMatch(printed, /Word count|Focus · Reading|A little perspective|Exit focus/);
  await page.getByRole('button', { name: 'Close statistics' }).click();
  await page.screenshot({ path: '/tmp/folio-writing-focus.png' });
  await page.getByRole('button', { name: 'Exit focus mode', exact: true }).click();
  await page.getByRole('button', { name: 'Document statistics', exact: true }).click();
  await page.screenshot({ path: '/tmp/folio-writing-statistics.png' });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: '/tmp/folio-writing-mobile.png' });
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await expect(page.getByRole('button', { name: 'Close statistics' })).toBeVisible();
  expect(errors).toEqual([]);
  console.log('PASS: Unicode document/selection statistics, block boundaries and live updates; nested outline positions, current section, filter and keyboard navigation; focus editing/reading/shortcuts/layered Escape; clean PDF, responsive panels, no browser errors.');
} finally { await browser.close(); }
