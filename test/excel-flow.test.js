import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';
import { DIMENSIONS } from '../dist/batch-excel.js';
import { generateDimensionGroups } from '../dist/batch-request.js';
import { snapshotDraft } from '../dist/batch-draft.js';

async function setup(request) {
  class Element {
    constructor() { this.handlers = {}; this.children = []; this.value = ''; this.style = {}; this.files = []; this.hidden = false; }
    addEventListener(type, fn) { this.handlers[type] = fn; }
    append(...children) { this.children.push(...children); }
    replaceChildren(...children) { this.children = children; }
    querySelectorAll(selector) { return this.children.flatMap(label => label.children).filter(item => item.type === 'checkbox' && (selector !== 'input:checked' || item.checked)); }
    focus() {}
  }
  const elements = new Map();
  const document = { getElementById(id) { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); }, createElement() { return new Element(); }, querySelectorAll() { return []; } };
  document.getElementById('note-length').value = '140';
  document.getElementById('access-code').value = 'fake';
  const writes = [];
  const context = vm.createContext({ document, DIMENSIONS, generateDimensionGroups, snapshotDraft, createDraftStore: () => ({ read: async () => null, write: async () => {}, clear: async () => {} }), AbortController, setTimeout, clearTimeout, console,
    hasOriginalResult: () => false, writeIntelligenceResult: (_, result) => writes.push(...result),
    fetch: async url => url === '/api/health' ? Response.json({ model_configured: true }) : request(),
  });
  const source = (await readFile(new URL('../dist/excel-page.js', import.meta.url), 'utf8')).replace(/import[\s\S]*?from\s+'[^']+';/g, '');
  vm.runInContext(source, context);
  vm.runInContext("state.workbook = {}; state.matches = [{ student: '测试学生', notes: '愿意参与课堂', targets: [{}] }]; state.readyCount = 1;", context);
  return { elements, writes, context, click: () => elements.get('generate-batch').handlers.click() };
}

test('pause finishes the current group; resume skips successes and completes remaining dimensions', async () => {
  let calls = 0;
  let app;
  app = await setup(() => {
    calls++;
    if (calls === 1) app.elements.get('pause-batch').handlers.click();
    const group = DIMENSIONS.slice((calls - 1) * 2, calls * 2);
    return Response.json({ intelligences: group.map(d => ({ id: d.id, level: 4, note: '积极参与。' })) });
  });
  await app.click();
  assert.equal(calls, 1);
  assert.equal(app.writes.length, 2);
  assert.equal(app.elements.get('generate-batch').disabled, false);
  await app.click();
  assert.equal(calls, 4);
  assert.equal(app.writes.length, 8);
  assert.equal(new Set(app.writes.map(item => item.id)).size, 8);
  assert.equal(app.elements.get('generate-batch').disabled, true);
  assert.equal(app.elements.get('download-excel').disabled, false);
  assert.equal(app.elements.get('overall-progress').value, 8);
});

test('empty input bypasses model and produces zeros; changing length makes items pending again', async () => {
  const app = await setup(() => { throw new Error('must not call model'); });
  vm.runInContext("state.matches[0].notes = '/';", app.context);
  await app.click();
  assert.equal(app.writes.length, 8);
  assert.ok(app.writes.every(item => item.level === 0 && item.note === ''));
  app.elements.get('note-length').value = '100';
  app.elements.get('note-length').handlers.input();
  assert.equal(app.elements.get('generate-batch').disabled, false);
});

test('preview writes one selected student and batch continues without repeating that student', async () => {
  const app = await setup(() => { throw new Error('no model needed'); });
  vm.runInContext("state.matches = [{student:'甲', notes:'/', targets:[{}]}, {student:'乙', notes:'/', targets:[{}]}, {student:'丙', notes:'/', targets:[{}], selected:false}]; state.readyCount = 3;", app.context);
  await app.elements.get('preview-one').handlers.click();
  assert.equal(app.writes.length, 8);
  assert.equal(vm.runInContext('state.matches[0].expanded', app.context), true);
  await app.click();
  assert.equal(app.writes.length, 16);
  assert.equal(vm.runInContext('state.matches[2].results', app.context), undefined);
});

test('selection respects search and pending filter without selecting hidden students', async () => {
  const app = await setup(() => {});
  vm.runInContext("state.matches = [{student:'甲', notes:'/', targets:[{}]}, {student:'乙', notes:'/', targets:[{}]}]; state.readyCount = 2;", app.context);
  app.elements.get('student-search').value = '乙';
  app.elements.get('students-none').handlers.click();
  assert.equal(vm.runInContext('selectedStudents().length', app.context), 1);
  assert.equal(vm.runInContext('selectedStudents()[0].student', app.context), '甲');
  await app.click();
  app.elements.get('student-search').value = '';
  app.elements.get('student-filter').value = 'pending';
  assert.equal(vm.runInContext('visibleStudents().length', app.context), 1);
  assert.equal(vm.runInContext('visibleStudents()[0].student', app.context), '乙');
});

test('draft snapshot includes template, results, selection and progress but excludes credentials', () => {
  const item = { student: '甲', key: '甲', notes: '观察内容', selected: false, results: new Map([['music', { id:'music', level:4, note:'参与积极' }]]) };
  const draft = snapshotDraft({ matches:[item], templateBytes:new Uint8Array([1,2]), templateName:'模板.xlsx', password:'must-not-store' }, new Map([[item, new Map([['music', '130:150']])]]), { areas:['music'], length:'140', metadata:{} });
  assert.equal(draft.records[0].selected, false);
  assert.deepEqual(draft.records[0].completed, [['music', '130:150']]);
  assert.equal(draft.records[0].results[0].level, 4);
  assert.ok(!JSON.stringify(draft).includes('must-not-store'));
});

test('regenerating one completed dimension updates only that dimension of the named student', async () => {
  const app = await setup(() => { throw new Error('no model needed'); });
  vm.runInContext("state.matches[0].notes = '/';", app.context);
  await app.click();
  const run = vm.runInContext("runBatch(false, state.matches[0], 'music')", app.context);
  await run;
  assert.equal(app.writes.length, 9);
  assert.equal(app.writes[8].id, 'music');
  assert.equal(app.writes[8].overwrite, true);
});
