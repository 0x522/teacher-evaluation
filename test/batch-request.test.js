import test from 'node:test';
import assert from 'node:assert/strict';
import { generateDimensionGroups } from '../dist/batch-request.js';
import { evaluateIntelligencesWithDeepSeek } from '../src/worker.js';
const areas = ['language', 'spatial', 'music', 'logic', 'self', 'body', 'social', 'nature'];
const results = group => group.map(id => ({ id, level: 3, note: '课堂参与积极。' }));

test('all eight dimensions are saved in small groups without repeating successful groups', async () => {
  const calls = [];
  const saved = [];
  const failures = await generateDimensionGroups(areas, async group => {
    calls.push([...group]);
    if (calls.length === 2) throw Object.assign(new Error('暂时不可用'), { status: 502 });
    return results(group);
  }, items => saved.push(...items), () => {}, async () => {});
  assert.deepEqual(failures, []);
  assert.equal(calls.length, 5);
  assert.deepEqual(calls[1], calls[2]);
  assert.deepEqual(saved.map(item => item.id), areas);
});

test('a failed group preserves successful groups and continues with remaining dimensions', async () => {
  const saved = [];
  const failures = await generateDimensionGroups(areas, async group => group.includes('music') ? [] : results(group), items => saved.push(...items), () => {}, async () => {});
  assert.equal(failures.length, 1);
  assert.deepEqual(failures[0].areas, ['music', 'logic']);
  assert.equal(saved.length, 6);
});

test('eight long notes receive enough output budget and report truncation explicitly', async () => {
  let budget;
  await assert.rejects(evaluateIntelligencesWithDeepSeek({ student: '小禾', notes: '积极参与', areas, noteMin: 100, noteMax: 120 }, 'fake', async (_, options) => {
    budget = JSON.parse(options.body).max_tokens;
    return Response.json({ choices: [{ finish_reason: 'length', message: { content: '{' } }] });
  }), error => error.status === 502 && /截断/.test(error.message));
  assert.ok(budget >= 8 * 120 * 3);
});

test('slash observations become zero and save to Excel without retries', async () => {
  const { default: ExcelJS } = await import('exceljs');
  const { writeIntelligenceResult } = await import('../dist/batch-excel.js');
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('记录');
  let calls = 0;
  const failures = await generateDimensionGroups(['music'], async () => {
    calls++;
    return evaluateIntelligencesWithDeepSeek({ student: '小禾', notes: '参与课堂', areas: ['music'] }, 'fake',
      async () => Response.json({ choices: [{ message: { content: JSON.stringify({ intelligences: [{ id: 'music', level: 4, note: ' / ' }] }) } }] }));
  }, result => writeIntelligenceResult({ worksheet: sheet, row: 1, columns: { music: { level: 1, note: 2 } } }, result), () => {});
  assert.deepEqual(failures, []);
  assert.equal(calls, 1);
  assert.equal(sheet.getCell('A1').value, 0);
  assert.equal(sheet.getCell('B1').value, null);
});

test('existing grade or note protects the pair until explicit overwrite; edits persist to Excel', async () => {
  const { default: ExcelJS } = await import('exceljs');
  const { writeIntelligenceResult } = await import('../dist/batch-excel.js');
  const sheet = new ExcelJS.Workbook().addWorksheet('记录');
  const target = { worksheet: sheet, row: 1, columns: { music: { level: 1, note: 2 } } };
  sheet.getCell('B1').value = '已有记录';
  writeIntelligenceResult(target, [{ id: 'music', level: 0, note: '' }]);
  assert.equal(sheet.getCell('A1').value, null);
  assert.equal(sheet.getCell('B1').value, '已有记录');
  writeIntelligenceResult(target, [{ id: 'music', level: 4, note: '教师修订', overwrite: true }]);
  assert.equal(sheet.getCell('A1').value, 4);
  assert.equal(sheet.getCell('B1').value, '教师修订');
});
