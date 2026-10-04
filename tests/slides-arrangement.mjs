import { chromium, expect } from '@playwright/test';

const browser = await chromium.launch({ headless: true, args: ['--no-sandbox'] });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, hasTouch: true });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.addInitScript(() => {
    if (sessionStorage.getItem('arrange-seeded')) return;
    sessionStorage.setItem('arrange-seeded', '1');
    localStorage.setItem('folio-files-v1', JSON.stringify([{ id: 'arrange-test', name: 'Arrange a story', kind: 'presentation', updatedAt: Date.now(), createdAt: Date.now(), starred: false, trashed: false, content: [
      { id: 'slide-one', title: 'Arrange objects', body: '', background: '#ffffff', layout: 'blank', notes: 'These notes stay private until requested.', elements: [
        { id: 'a', type: 'shape', shape: 'rectangle', x: 10, y: 10, width: 10, height: 10, fill: '#2d6554' },
        { id: 'b', type: 'shape', shape: 'ellipse', x: 35, y: 20, width: 15, height: 15, fill: '#8d5e3d' },
        { id: 'c', type: 'shape', shape: 'rectangle', x: 70, y: 40, width: 10, height: 20, fill: '#6f5689' },
        { id: 'caption', type: 'text', x: 10, y: 75, width: 50, height: 10, text: 'An editable caption', fontSize: 22 },
      ] },
      { id: 'slide-two', title: 'A second slide', body: 'Ready for questions.', background: '#244e40', layout: 'title' },
    ] }]));
  });
  await page.goto('http://localhost:5173');
  await page.getByText('Arrange a story', { exact: true }).click();
  const objects = () => page.locator('.slides-canvas [data-element-id]');
  const object = id => page.locator(`.slides-canvas [data-element-id="${id}"]`);
  const selected = () => page.locator('.slides-canvas .slides-element-selected');
  const positions = () => objects().evaluateAll(elements => Object.fromEntries(elements.map(element => [element.dataset.elementId, { x: parseFloat(element.style.left), y: parseFloat(element.style.top), width: parseFloat(element.style.width), height: parseFloat(element.style.height) }])));
  const button = name => page.getByRole('button', { name, exact: true });
  const selectThree = async () => { await object('a').click(); await object('b').click({ modifiers: ['Shift'] }); await object('c').click({ modifiers: ['Shift'] }); await expect(selected()).toHaveCount(3); };

  await selectThree();
  await button('Align selected objects top').click();
  let geometry = await positions();
  expect([geometry.a.y, geometry.b.y, geometry.c.y]).toEqual([10, 10, 10]);
  await button('Undo').click();
  geometry = await positions();
  expect([geometry.a.y, geometry.b.y, geometry.c.y]).toEqual([10, 20, 40]);
  await button('Redo').click();
  await button('Distribute objects horizontally').click();
  geometry = await positions();
  expect([geometry.a.x, geometry.b.x, geometry.c.x]).toEqual([10, 37.5, 70]);
  await button('Undo').click(); expect((await positions()).b.x).toBe(35);
  await button('Redo').click(); expect((await positions()).b.x).toBe(37.5);

  const beforeMove = await positions();
  const move = await button('Move selected objects').boundingBox();
  await page.mouse.move(move.x + 10, move.y + 10); await page.mouse.down();
  await page.mouse.move(move.x + 45, move.y + 35, { steps: 8 }); await page.mouse.up();
  geometry = await positions();
  expect(geometry.a.x).toBeGreaterThan(beforeMove.a.x);
  expect(geometry.b.x - beforeMove.b.x).toBeCloseTo(geometry.a.x - beforeMove.a.x, 5);
  expect(geometry.c.y - beforeMove.c.y).toBeCloseTo(geometry.a.y - beforeMove.a.y, 5);
  await button('Undo').click(); expect(await positions()).toEqual(beforeMove);

  await button('Duplicate selected objects').click(); await expect(objects()).toHaveCount(7); await expect(selected()).toHaveCount(3);
  await button('Delete selected objects').click(); await expect(objects()).toHaveCount(4);
  await button('Undo').click(); await expect(objects()).toHaveCount(7);
  await button('Redo').click(); await expect(objects()).toHaveCount(4);
  await page.getByRole('region', { name: 'Slide canvas', exact: true }).focus();
  await page.keyboard.press('Control+a'); await expect(selected()).toHaveCount(4);
  const crowded = await positions();
  await button('Distribute objects horizontally').click();
  await expect(page.getByText('Make the objects smaller or give them more room to space them evenly.', { exact: true })).toBeVisible();
  expect(await positions()).toEqual(crowded);
  await page.keyboard.press('Control+d'); await expect(objects()).toHaveCount(8);
  await page.keyboard.press('Control+z'); await expect(objects()).toHaveCount(4);
  await page.keyboard.press('Control+a'); await page.keyboard.press('Delete'); await expect(objects()).toHaveCount(0);
  await page.keyboard.press('Control+z'); await expect(objects()).toHaveCount(4);

  const caption = page.getByLabel('Text box content', { exact: true });
  await caption.click(); await page.keyboard.press('Control+a');
  await expect(selected()).toHaveCount(1);
  expect(await caption.evaluate(element => element.selectionEnd - element.selectionStart)).toBe('An editable caption'.length);
  await object('a').click();
  await page.keyboard.press('ArrowRight');
  expect((await positions()).a.x).toBe(10.5);
  await button('Undo').click(); expect((await positions()).a.x).toBe(10);

  await button('Grid').click(); await expect(page.locator('.slides-grid-overlay')).toBeVisible();
  let anchor = await button('Move selected object').boundingBox();
  await page.mouse.move(anchor.x + 10, anchor.y + 10); await page.mouse.down();
  await page.mouse.move(anchor.x + 33, anchor.y + 29, { steps: 5 }); await page.mouse.up();
  geometry = await positions(); expect(geometry.a.x % 2).toBeCloseTo(0); expect(geometry.a.y % 2).toBeCloseTo(0);
  anchor = await button('Move selected object').boundingBox();
  await page.keyboard.down('Alt'); await page.mouse.move(anchor.x + 10, anchor.y + 10); await page.mouse.down();
  await page.mouse.move(anchor.x + 21, anchor.y + 17, { steps: 5 }); await page.mouse.up(); await page.keyboard.up('Alt');
  geometry = await positions(); expect(Math.abs(geometry.a.x % 2)).toBeGreaterThan(.1);
  await button('Undo').click();
  await page.getByLabel('Alignment target', { exact: true }).selectOption('slide');
  await button('Align selected objects right').click(); expect((await positions()).a.x).toBe(90);
  await button('Undo').click();

  const present = button('Present'); await present.click();
  const presenter = page.getByRole('dialog', { name: 'Presentation', exact: true });
  await expect(presenter).toBeFocused();
  await page.keyboard.press('Tab'); await expect(button('Show speaker notes')).toBeFocused();
  await page.keyboard.press('Space'); await expect(page.getByLabel('Presenter notes', { exact: true })).toBeVisible();
  await expect(presenter).toContainText('1 / 2');
  await page.keyboard.press('Shift+Tab'); await expect(button('Next slide')).toBeFocused();
  await page.keyboard.press('Tab'); await expect(button('Hide speaker notes')).toBeFocused();
  await page.keyboard.press('ArrowRight'); await expect(presenter).toContainText('2 / 2');
  await page.keyboard.press('Home'); await expect(presenter).toContainText('1 / 2');
  await page.keyboard.press('Escape'); await expect(present).toBeFocused();

  await page.setViewportSize({ width: 390, height: 844 });
  await button('Clear').click();
  await page.locator('[data-layer-id="a"] input').check();
  await page.locator('[data-layer-id="b"] input').check();
  await expect(selected()).toHaveCount(2);
  await button('Close design panel').click();
  const beforeTouch = await positions();
  const touchAnchor = await button('Move selected objects').boundingBox();
  const cdp = await page.context().newCDPSession(page);
  const start = { x: touchAnchor.x + touchAnchor.width / 2, y: touchAnchor.y + touchAnchor.height / 2 };
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [start] });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: start.x + 12, y: start.y + 9 }] });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  geometry = await positions();
  expect(geometry.a.x).toBeGreaterThan(beforeTouch.a.x);
  expect(geometry.a.x - beforeTouch.a.x).toBeCloseTo(geometry.b.x - beforeTouch.b.x, 4);
  expect(geometry.a.y - beforeTouch.a.y).toBeCloseTo(geometry.b.y - beforeTouch.b.y, 4);
  await expect(selected()).toHaveCount(2);

  await page.waitForTimeout(500);
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('folio-files-v1'))[0].content);
  expect(saved[0].elements.find(element => element.id === 'a').x).toBeCloseTo(geometry.a.x, 5);
  const roundtrip = await page.evaluate(async () => {
    const source = JSON.parse(localStorage.getItem('folio-files-v1'))[0].content;
    const { buildPresentationBlob } = await import('/src/editors/slidesPptx.ts');
    const { importOfficeFile } = await import('/src/lib/fileIO.ts');
    const blob = await buildPresentationBlob(source, 'Arranged objects');
    return (await importOfficeFile(new File([blob], 'Arranged objects.pptx'))).content;
  });
  for (const source of saved[0].elements) {
    const imported = roundtrip[0].elements.find(element => element.id === source.id);
    for (const key of ['x', 'y', 'width', 'height']) expect(imported[key]).toBeCloseTo(source[key], 3);
  }
  await page.screenshot({ path: '/tmp/folio-slides-arrangement-mobile.png', fullPage: true });
  expect(errors).toEqual([]);
  console.log('PASS: slide multiselection, alignment, spacing, group drag/duplicate/delete/undo, keyboard editing, grid snapping, presenter focus, mobile touch grouping, and PPTX geometry roundtrip.');
} finally { await browser.close(); }
