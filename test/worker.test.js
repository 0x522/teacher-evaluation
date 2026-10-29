import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import worker, { evaluateWithDeepSeek, evaluateIntelligencesWithDeepSeek } from '../src/worker.js';
import { DIMENSIONS, clearIntelligenceResult, extractTemplateGuidance, findInputRecords, findTemplateMetadataTargets, findTemplateTargets, matchRecords, prepareTemplateRecords, writeIntelligenceResult, writeTemplateMetadata } from '../dist/batch-excel.js';

const origin = 'https://teacher.example';
const env = {
  DEEPSEEK_API_KEY: 'fake-key-for-tests',
  APP_PASSWORD: 'long-fake-password-for-tests',
  ASSETS: { fetch: async () => new Response('page', { status: 200 }) },
};
const input = {
  student: '小禾',
  course: '英语戏剧课',
  notes: '愿意跟着节奏模仿动物叫声。',
  length: 120,
  areas: ['music'],
};

function post(body = input, headers = {}) {
  return new Request(`${origin}/api/evaluate`, {
    method: 'POST',
    headers: {
      Origin: origin,
      'Content-Type': 'application/json',
      'X-App-Password': env.APP_PASSWORD,
      ...headers,
    },
    body: JSON.stringify(body),
  });
}

test('health reports whether secrets are configured, without disclosing them', async () => {
  const response = await worker.fetch(new Request(`${origin}/api/health`), env);
  assert.deepEqual(await response.json(), { model_configured: true });
  assert.equal(response.headers.get('Cache-Control'), 'no-store');
  const noSecrets = await worker.fetch(new Request(`${origin}/api/health`), { ASSETS: env.ASSETS });
  assert.deepEqual(await noSecrets.json(), { model_configured: false });
});

test('static pages use Cloudflare asset binding', async () => {
  const response = await worker.fetch(new Request(`${origin}/`), env);
  assert.equal(await response.text(), 'page');
});

test('single evaluation and Excel batch tools live on separate pages', async () => {
  const [home, excel] = await Promise.all([
    readFile(new URL('../dist/index.html', import.meta.url), 'utf8'),
    readFile(new URL('../dist/excel.html', import.meta.url), 'utf8'),
  ]);
  assert.match(home, /先写下你看见的/);
  assert.match(home, /href="\/excel\.html"/);
  assert.doesNotMatch(home, /id="input-excel"/);
  assert.match(excel, /id="input-excel"/);
  assert.match(excel, /id="batch-tags"/);
  assert.match(excel, /id="note-length"/);
  assert.match(excel, /返回单人课程评价/);
  assert.doesNotMatch(excel, /id="notes"/);
});

test('cross-origin submissions and wrong password cannot call the model', async () => {
  const crossOrigin = await worker.fetch(post(input, { Origin: 'https://other.example' }), env);
  assert.equal(crossOrigin.status, 403);
  const wrongPassword = await worker.fetch(post(input, { 'X-App-Password': 'wrong' }), env);
  assert.equal(wrongPassword.status, 401);
});

test('missing secrets and invalid input return safe errors', async () => {
  const missing = await worker.fetch(post(), { ASSETS: env.ASSETS });
  assert.equal(missing.status, 503);
  const bad = await worker.fetch(post({ ...input, length: 10 }), env);
  assert.equal(bad.status, 400);
});

test('server calls official DeepSeek endpoint and returns generated text', async () => {
  let requestBody;
  const mockFetch = async (url, options) => {
    assert.equal(url, 'https://api.deepseek.com/chat/completions');
    assert.equal(options.headers.Authorization, 'Bearer fake-key-for-tests');
    requestBody = JSON.parse(options.body);
    return new Response(JSON.stringify({ choices: [{ message: { content: '  小禾愿意参与表演。  ' } }] }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  };
  const text = await evaluateWithDeepSeek(input, env.DEEPSEEK_API_KEY, mockFetch);
  assert.equal(text, '小禾愿意参与表演。');
  assert.equal(requestBody.model, 'deepseek-flash');
  assert.equal(requestBody.thinking.type, 'disabled');
  assert.match(requestBody.messages[1].content, /愿意跟着节奏/);
});

test('DeepSeek authentication failure gets a specific message', async () => {
  await assert.rejects(
    evaluateWithDeepSeek(input, env.DEEPSEEK_API_KEY, async () => new Response('', { status: 401 })),
    error => error.status === 502 && /API Key 无效/.test(error.message),
  );
});

test('valid API request works end to end with a mocked model', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(JSON.stringify({ choices: [{ message: { content: '测试评价。' } }] }), { status: 200 });
  try {
    const response = await worker.fetch(post(), env);
    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), { text: '测试评价。', source: 'deepseek' });
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('intelligence evaluation requests JSON for only the selected dimensions', async () => {
  let requestBody;
  const responseItems = [{ id: 'body', level: 4, note: '愿意跟随示范完成动作。' }];
  const result = await evaluateIntelligencesWithDeepSeek(
    { student: '小禾', notes: '愿意跟随老师完成动作。', areas: ['body'], guidance: '4分表示表现优秀。', noteMin: 30, noteMax: 50 },
    env.DEEPSEEK_API_KEY,
    async (url, options) => {
      requestBody = JSON.parse(options.body);
      return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify({ intelligences: responseItems }) } }] }), { status: 200 });
    },
  );
  assert.equal(requestBody.response_format.type, 'json_object');
  assert.match(requestBody.messages[0].content, /4分表示表现优秀/);
  assert.match(requestBody.messages[0].content, /30–50 个汉字/);
  assert.match(requestBody.messages[0].content, /不要写“课堂评价未提及”/);
  assert.equal(result.length, 1);
  assert.deepEqual(result.find(item => item.id === 'body'), { id: 'body', level: 4, note: '愿意跟随示范完成动作。' });
});

test('weak validation fills missing dimensions and normalizes grades without losing valid notes', async () => {
  const result = await evaluateIntelligencesWithDeepSeek(
    { student: '小禾', notes: '本周课堂表现积极。', areas: ['music', 'logic', 'body', 'self'] },
    env.DEEPSEEK_API_KEY,
    async () => Response.json({ choices: [{ message: { content: JSON.stringify({ intelligences: [
      { id: 'music', level: '4', note: '愿意跟唱歌曲。' },
      { id: 'logic', level: null, note: '' },
      { id: 'body', level: 9, note: '建议继续观察动作协调性。' },
    ] }) } }] }),
  );
  assert.deepEqual(result, [
    { id: 'music', level: 4, note: '愿意跟唱歌曲。' },
    { id: 'logic', level: 0, note: '' },
    { id: 'body', level: 3, note: '建议继续观察动作协调性。' },
    { id: 'self', level: 0, note: '' },
  ]);
});

test('empty structured results mark missing dimensions zero without inventing observations', async () => {
  const result = await evaluateIntelligencesWithDeepSeek(
    { student: '小禾', notes: '愿意参与活动。', areas: ['music', 'body'] }, 'fake',
    async () => Response.json({ choices: [{ message: { content: '```json\n{}\n```' } }] }),
  );
  assert.deepEqual(result, ['music', 'body'].map(id => ({ id, level: 0, note: '' })));
});

test('Excel parsing supports variable row counts and Chinese-English template names', () => {
  const inputBook = new ExcelJS.Workbook();
  const inputSheet = inputBook.addWorksheet('本周评价');
  inputSheet.addRow(['序号', '孩子姓名', '本周课堂评价']);
  inputSheet.addRow([1, '高誉畅', '主动参与动作游戏。']);
  inputSheet.addRow([]);
  inputSheet.addRow([2, '新同学', '愿意跟唱歌曲。']);

  const templateBook = new ExcelJS.Workbook();
  const templateSheet = templateBook.addWorksheet('成长记录');
  const guideSheet = templateBook.addWorksheet('填写说明');
  guideSheet.addRow(['智能名称', '核心内涵', '幼儿课堂表现参考']);
  guideSheet.addRow(['音乐智能', '音准与节奏感', '喜欢跟唱、能模仿旋律']);
  guideSheet.addRow(['4分', '表现优秀', '能主动展现相关能力']);
  templateSheet.mergeCells('A3:B3');templateSheet.getCell('A3').value = '周次';
  templateSheet.mergeCells('C3:D3');templateSheet.getCell('C3').value = '第 ___ 周';
  templateSheet.mergeCells('E3:F3');templateSheet.getCell('E3').value = '记录日期';
  templateSheet.mergeCells('G3:H3');templateSheet.getCell('G3').value = '____年__月__日';
  templateSheet.mergeCells('I3:J3');templateSheet.getCell('I3').value = '带班老师';
  templateSheet.mergeCells('K3:L3');
  templateSheet.mergeCells('M3:N3');templateSheet.getCell('M3').value = '班级';
  templateSheet.mergeCells('O3:V3');
  templateSheet.getCell('A5').value = '序号';
  templateSheet.getCell('B5').value = '孩子姓名';
  DIMENSIONS.forEach((item, index) => {
    const column = 3 + index * 2;
    templateSheet.getCell(5, column).value = `${item.name}\nEnglish`;
    templateSheet.getCell(6, column).value = '表现等级';
    templateSheet.getCell(6, column + 1).value = '亮点/观察记录';
  });
  templateSheet.getCell('B7').value = '高誉畅Kevin';
  templateSheet.getCell('A7').value = 1;
  templateSheet.getCell('A8').value = 2;
  templateSheet.getCell('B9').value = '另一位学生';
  templateSheet.getCell('A9').value = 3;
  templateSheet.getCell('C7').value = 5;
  templateSheet.getCell('S7').value = '保持不变';

  const records = findInputRecords(inputBook);
  assert.equal(findTemplateTargets(templateBook).length, 2);
  const prepared = prepareTemplateRecords(templateBook, records);
  const matches = matchRecords(records, prepared.targets);
  assert.equal(records.length, 2);
  assert.equal(prepared.targets.length, 3);
  assert.deepEqual(prepared.appended.map(item => item.student), ['新同学']);
  assert.equal(templateSheet.getCell('B8').value, '新同学');
  assert.equal(matches[0].targets.length, 1);
  assert.equal(matches[1].targets.length, 1);

  clearIntelligenceResult(matches[0].targets[0], ['body']);
  writeIntelligenceResult(matches[0].targets[0], [{ id: 'body', level: 4, note: '动作模仿积极。' }]);
  assert.equal(templateSheet.getCell('M7').value, 4);
  assert.equal(templateSheet.getCell('N7').value, '动作模仿积极。');
  assert.equal(templateSheet.getCell('C7').value, 5);
  assert.equal(templateSheet.getCell('S7').value, '保持不变');

  const metadataTargets = findTemplateMetadataTargets(templateBook);
  writeTemplateMetadata(metadataTargets, { week: '6', date: '2026年9月23日', teacher: '王老师', className: 'Yoyo班' });
  assert.equal(templateSheet.getCell('C3').value, '第 6 周');
  assert.equal(templateSheet.getCell('G3').value, '2026年9月23日');
  assert.equal(templateSheet.getCell('K3').value, '王老师');
  assert.equal(templateSheet.getCell('O3').value, 'Yoyo班');
  const guidance = extractTemplateGuidance(templateBook);
  assert.match(guidance, /音乐智能 \| 音准与节奏感 \| 喜欢跟唱/);
  assert.match(guidance, /4分 \| 表现优秀/);
});
