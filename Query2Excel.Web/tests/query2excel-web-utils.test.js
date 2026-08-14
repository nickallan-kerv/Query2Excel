const test = require('node:test');
const assert = require('node:assert/strict');
const utils = require('../wwwroot/scripts/query2excel-web-utils.js');

test('normalizeRowStyleNames trims and de-duplicates case-insensitively', () => {
  const result = utils.normalizeRowStyleNames([
    ' Accent1 ',
    'accent1',
    'Accent2',
    '',
    null,
    '  ',
    'ACCENT2',
    'Good'
  ]);

  assert.deepEqual(result, ['Accent1', 'Accent2', 'Good']);
});

test('computeMetadataTitleSync auto-syncs while title is empty or last auto value', () => {
  const first = utils.computeMetadataTitleSync('Orders', '', '');
  assert.equal(first.title, 'Orders');
  assert.equal(first.lastAutoSyncedTitle, 'Orders');
  assert.equal(first.wasAutoSynced, true);

  const follow = utils.computeMetadataTitleSync('Orders Q4', first.title, first.lastAutoSyncedTitle);
  assert.equal(follow.title, 'Orders Q4');
  assert.equal(follow.lastAutoSyncedTitle, 'Orders Q4');
  assert.equal(follow.wasAutoSynced, true);

  const manual = utils.computeMetadataTitleSync('Should Not Override', 'Custom Title', follow.lastAutoSyncedTitle);
  assert.equal(manual.title, 'Custom Title');
  assert.equal(manual.lastAutoSyncedTitle, follow.lastAutoSyncedTitle);
  assert.equal(manual.wasAutoSynced, false);
});

test('computeInsertRange prefers selected range and clamps to script bounds', () => {
  const replace = utils.computeInsertRange(25, 12, 5, 30);
  assert.deepEqual(replace, { start: 5, end: 25, isReplace: true });

  const insert = utils.computeInsertRange(25, 33, null, null);
  assert.deepEqual(insert, { start: 25, end: 25, isReplace: false });
});

test('computeSubmenuLayout flips left and stays within viewport bottom', () => {
  const layout = utils.computeSubmenuLayout({
    menuRight: 990,
    wrapperTop: 680,
    submenuWidth: 220,
    submenuHeight: 500,
    viewportWidth: 1000,
    viewportHeight: 736,
    viewportPadding: 8,
    minimumHeight: 120
  });

  assert.equal(layout.openLeft, true);
  assert.ok(layout.submenuTop < 0);
  assert.ok(layout.maxHeight >= 120);
  assert.ok(680 + layout.submenuTop + layout.maxHeight <= 736);
});
