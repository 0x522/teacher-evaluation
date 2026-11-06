import { createDraftStore, snapshotDraft } from './batch-draft.js';
import { generateDimensionGroups } from './batch-request.js';
import {
  DIMENSIONS,
  hasOriginalResult,
  extractTemplateGuidance,
  findInputRecords,
  findTemplateMetadataTargets,
  matchRecords,
  prepareTemplateRecords,
  writeIntelligenceResult,
  writeTemplateMetadata,
} from './batch-excel.js';

const $ = id => document.getElementById(id);
const state = {
  workbook: null,
  matches: [],
  targets: [],
  metadataTargets: [],
  guidance: '',
  readyCount: 0,
  templateName: '多元智能成长记录模板.xlsx',
};
let modelReady = false;
let generating = false;
let parsing = false;
let connecting = false;
let pauseRequested = false;
let completedGroups = new Map();
const draftStore = createDraftStore();
let availableDraft = null;
let draftEnabled = true;
let draftRevision = 0;
const selectedStudents = () => state.matches.filter(item => item.targets.length && item.selected !== false);
const currentSignature = () => { const range = selectedNoteRange(); return `${range.noteMin}:${range.noteMax}`; };
const protectedFor = (item, id) => !$('overwrite-existing').checked &&
  item.targets.length > 0 && item.targets.every(target => hasOriginalResult(target, id));
const completedFor = (item, id, signature) =>
  completedGroups.get(item)?.get(id) === signature &&
  (!$('overwrite-existing').checked || item.results?.get(id)?.overwrite ||
    !item.targets.some(target => hasOriginalResult(target, id)));
const pendingFor = item => selectedAreas().some(id => !protectedFor(item, id) && !completedFor(item, id, currentSignature()));
$('overwrite-existing').addEventListener('change', () => { updateSelection(); renderResults(); saveDraft(); });
function visibleStudents() {
  const query = $('student-search').value.trim();
  const filter = $('student-filter').value;
  return state.matches.filter(item => item.student.includes(query) && (filter === 'failed' ? item.error : filter === 'pending' ? pendingFor(item) : true));
}
function draftSettings() {
  return { areas: selectedAreas(), length: $('note-length').value, overwrite: $('overwrite-existing').checked,
    metadata: Object.fromEntries(['meta-week', 'meta-date', 'meta-teacher', 'meta-class'].map(id => [id, $(id).value])) };
}
async function saveDraft() {
  if (!state.workbook || !state.templateBytes || !draftEnabled) return;
  const revision = ++draftRevision;
  try {
    $('draft-status').textContent = '正在保存本机草稿…';
    await draftStore.write(snapshotDraft(state, completedGroups, draftSettings()));
    if (revision === draftRevision) $('draft-status').textContent = '草稿已保存，可刷新后恢复任务。';
  } catch {
    if (revision === draftRevision) $('draft-status').textContent = '本机草稿保存失败，请下载结果后再离开。';
  }
}
$('clear-draft').addEventListener('click', async () => {
  draftEnabled = false;
  draftRevision++;
  try {
    await draftStore.clear();
    availableDraft = null;
    $('restore-draft').hidden = true;
    $('draft-status').textContent = '本机草稿已清除；当前任务仍可继续。本次不再自动保存，新上传文件后恢复保存。';
  } catch { $('draft-status').textContent = '草稿清除失败，请检查浏览器存储权限。'; }
});
$('restore-draft').addEventListener('click', async () => {
  if (!availableDraft || generating || parsing || connecting) return;
  parsing = true;
  updateSelection();
  try {
    const draft = availableDraft;
    const book = new globalThis.ExcelJS.Workbook();
    await book.xlsx.load(draft.templateBytes);
    const prepared = prepareTemplateRecords(book, draft.records);
    const matches = matchRecords(draft.records, prepared.targets);
    const restored = new Map();
    matches.forEach((item, index) => {
      const saved = draft.records[index];
      item.selected = saved.selected;
      item.error = saved.error;
      item.results = new Map(saved.results.map(result => [result.id, result]));
      item.targets.forEach(target => writeIntelligenceResult(target, saved.results));
      restored.set(item, new Map(saved.completed));
    });
    Object.assign(state, { workbook: book, templateBytes: draft.templateBytes, templateName: draft.templateName,
      matches, targets: prepared.targets, readyCount: matches.filter(item => item.targets.length).length,
      metadataTargets: findTemplateMetadataTargets(book), guidance: extractTemplateGuidance(book) });
    completedGroups = restored;
    tags.querySelectorAll('input').forEach(input => { input.checked = draft.settings.areas.includes(input.value); });
    $('note-length').value = draft.settings.length;
    $('overwrite-existing').checked = Boolean(draft.settings.overwrite);
    Object.entries(draft.settings.metadata).forEach(([id, value]) => { $(id).value = value; });
    draftEnabled = true;
    $('download-excel').disabled = false;
    $('restore-draft').hidden = true;
    $('draft-status').textContent = '上次任务已恢复，请重新输入口令后继续。';
    $('next-step').textContent = '已恢复学生、设置和结果。继续生成时会跳过成功项。';
    updateNoteRange();
    renderResults();
  } catch { showSummary('草稿恢复失败，请重新上传文件；也可清除旧草稿。', 'error'); }
  finally { parsing = false; updateSelection(); renderResults(); }
});
for (const id of ['student-search', 'student-filter']) $(id).addEventListener(id === 'student-search' ? 'input' : 'change', renderResults);
for (const [id, selected] of [['students-all', true], ['students-none', false]]) {
  $(id).addEventListener('click', () => {
    if (generating || parsing || connecting) return;
    visibleStudents().filter(item => item.targets.length).forEach(item => { item.selected = selected; });
    renderResults(); updateSelection(); saveDraft();
  });
}
for (const id of ['meta-week', 'meta-date', 'meta-teacher', 'meta-class']) $(id).addEventListener('change', saveDraft);


const tags = $('batch-tags');
for (const item of DIMENSIONS) {
  const label = document.createElement('label');
  label.className = 'tag';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.value = item.id;
  input.name = 'batch-intelligence';
  input.checked = true;
  const span = document.createElement('span');
  span.textContent = item.name;
  label.append(input, span);
  tags.append(label);
}

const selectedAreas = () => [...tags.querySelectorAll('input:checked')].map(input => input.value);
const selectedNoteRange = () => {
  const center = Number($('note-length').value);
  return { noteMin: Math.max(20, center - 10), noteMax: Math.min(150, center + 10) };
};

function updateNoteRange() {
  const { noteMin, noteMax } = selectedNoteRange();
  $('note-length-value').textContent = `${noteMin}–${noteMax} 字`;
}

function showSummary(message, kind = 'info') {
  const element = $('batch-summary');
  element.textContent = message;
  element.hidden = !message;
  element.className = `summary ${kind}`;
}

function updateSelection() {
  const selected = selectedAreas();
  $('batch-area-count').textContent = selected.length
    ? `已选择 ${selected.length} 个智能维度，可再次点击取消`
    : '请选择至少一个智能维度';
  const { noteMin, noteMax } = selectedNoteRange();
  const signature = `${noteMin}:${noteMax}`;
  const pending = selectedStudents().reduce((sum, item) => sum + selected.filter(id => !protectedFor(item, id) && !completedFor(item, id, signature)).length, 0);
  $('generate-batch').disabled = parsing || connecting || generating || !(state.workbook && state.readyCount && selected.length && pending);
  $('preview-one').disabled = $('generate-batch').disabled;
  if (!generating && state.workbook) {
    const total = selectedStudents().length * selected.length;
    $('overall-progress').max = total || 1;
    $('overall-progress').value = total - pending;
  }
  $('student-count').textContent = `已选 ${selectedStudents().length} / ${state.readyCount} 人`;
  if (!generating) $('generate-batch').textContent = state.workbook && selected.length && !pending ? '所选项目已完成' : completedGroups.size ? '继续生成未完成项' : '开始批量生成';
}

tags.addEventListener('change', () => { updateSelection(); renderResults(); saveDraft(); });
$('note-length').addEventListener('input', () => { updateNoteRange(); updateSelection(); renderResults(); saveDraft(); });
for (const [id, checked] of [['select-all', true], ['select-none', false]]) {
  $(id).addEventListener('click', () => {
    tags.querySelectorAll('input').forEach(input => { input.checked = checked; });
    updateSelection(); renderResults(); saveDraft();
  });
}
$('pause-batch').addEventListener('click', () => {
  pauseRequested = true;
  $('pause-batch').disabled = true;
  $('pause-batch').textContent = '完成当前组后暂停…';
});
function renderResults() {
  const list = $('result-list');
  list.replaceChildren();
  for (const item of visibleStudents()) {
    const selection = document.createElement('label');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = item.selected !== false && Boolean(item.targets.length);
    checkbox.disabled = generating || parsing || !item.targets.length;
    checkbox.addEventListener('change', () => { item.selected = checkbox.checked; updateSelection(); saveDraft(); });
    const label = document.createElement('span');
    label.textContent = ` 选择 ${item.student}`;
    selection.append(checkbox, label);
    list.append(selection);
    const row = document.createElement('details');
    row.open = Boolean(item.expanded);
    row.addEventListener('toggle', () => { item.expanded = row.open; });
    row.style.cssText = 'padding:10px 0;border-bottom:1px solid #edf0ed;font-size:.82rem';
    const title = document.createElement('summary');
    const count = item.results?.size || 0;
    title.textContent = `${item.student} · ${!item.targets.length ? '未匹配，请检查模板空位' : item.error || (count ? `已完成 ${count} 项，点击核对` : !item.notes?.trim() || item.notes.trim() === '/' ? '评价为空，将记 0 分' : '已匹配，等待生成')}`;
    row.append(title);
    const action = (label, fn) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'secondary';
      button.textContent = label;
      button.disabled = generating || parsing || connecting;
      button.addEventListener('click', fn);
      return button;
    };
    if (item.targets.length) {
      const preserved = DIMENSIONS.filter(d => protectedFor(item, d.id)).map(d => d.name);
      if (preserved.length) {
        const hint = document.createElement('p');
        hint.textContent = '保留模板已有记录：' + preserved.join('、');
        row.append(hint);
      }
      row.append(action('重新生成该学生所选智能', () => runBatch(false, item)));
    }
    for (const result of item.results?.values() || []) {
      const text = document.createElement('p');
      text.textContent = `${DIMENSIONS.find(d => d.id === result.id)?.name} · ${result.level} 分：${result.note || '跳过（无观察记录）'}`;
      text.style.cssText = 'line-height:1.7;color:#627066';
      const edit = action('编辑分数和记录', () => {
        if (generating || parsing || connecting) return;
        const level = document.createElement('input');
        level.type = 'number'; level.min = '0'; level.max = '5'; level.step = '1';
        level.value = result.level; level.setAttribute('aria-label', '表现等级');
        const note = document.createElement('textarea');
        note.value = result.note; note.setAttribute('aria-label', '观察记录');
        const save = action('保存修改', () => {
          const score = Number(level.value);
          if (generating || parsing || connecting) return;
          if (level.value === '' || !Number.isInteger(score) || score < 0 || score > 5) {
            return showSummary('请输入 0 到 5 的整数分数。', 'error');
          }
          const updated = { ...result, level: score, note: note.value.trim() };
          item.targets.forEach(target => writeIntelligenceResult(target, [updated]));
          item.results.set(result.id, updated);
          item.expanded = true;
          renderResults(); saveDraft();
        });
        text.replaceChildren(level, note, save, action('取消', renderResults));
      });
      row.append(text, edit, action('重新生成该维度', () => runBatch(false, item, result.id)));
    }
    list.append(row);
  }
}
for (const id of ['input-excel', 'output-template']) {
  $(id).addEventListener('change', () => {
    state.workbook = null;
    state.matches = [];
    state.templateBytes = null;
    $('overwrite-existing').checked = false;
    state.readyCount = 0;
    completedGroups.clear();
    $('download-excel').disabled = true;
    updateSelection();
    $('result-list').replaceChildren();
    if ($('input-excel').files[0] && $('output-template').files[0]) $('parse-excel').click();
  });
}

updateNoteRange();

async function checkConnection() {
  const element = $('connection');
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 10_000);
  try {
    const response = await fetch('/api/health', { cache: 'no-store', signal: controller.signal });
    if (!response.ok) throw new Error('unavailable');
    const data = await response.json();
    modelReady = Boolean(data.model_configured);
    element.textContent = modelReady ? 'DeepSeek 已配置' : 'DeepSeek 尚未配置';
    element.className = `connection ${modelReady ? 'ready' : 'offline'}`;
  } catch {
    modelReady = false;
    element.textContent = '未连接到评价服务';
    element.className = 'connection offline';
  } finally {
    clearTimeout(timer);
  }
}

async function loadExcel(file) {
  if (!globalThis.ExcelJS) throw new Error('Excel 解析组件加载失败，请刷新页面后重试。');
  if (!file || !file.name.toLowerCase().endsWith('.xlsx')) throw new Error('请选择 .xlsx 格式的 Excel 文件。');
  if (file.size > 15 * 1024 * 1024) throw new Error('Excel 文件不能超过 15MB。');
  const workbook = new globalThis.ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return workbook;
}

$('parse-excel').addEventListener('click', async () => {
  const inputFile = $('input-excel').files[0];
  const templateFile = $('output-template').files[0];
  if (!inputFile || !templateFile) {
    showSummary('请选择本周课堂评价文件和多元智能成长记录模板。', 'error');
    return;
  }
  if (generating || parsing || connecting) return;
  parsing = true;
  for (const id of ['input-excel', 'output-template', 'restore-draft']) $(id).disabled = true;
  const button = $('parse-excel');
  button.disabled = true;
  $('generate-batch').disabled = true;
  $('download-excel').disabled = true;
  $('batch-progress').textContent = '正在解析…';
  try {
    const [inputWorkbook, templateWorkbook] = await Promise.all([loadExcel(inputFile), loadExcel(templateFile)]);
    const templateBytes = await templateFile.arrayBuffer();
    const records = findInputRecords(inputWorkbook);
    const prepared = prepareTemplateRecords(templateWorkbook, records);
    const targets = prepared.targets;
    const metadataTargets = findTemplateMetadataTargets(templateWorkbook);
    const guidance = extractTemplateGuidance(templateWorkbook);
    const matches = matchRecords(records, targets);
    const ready = matches.filter(item => item.targets.length);
    const noNotes = matches.filter(item => !item.notes?.trim() || item.notes.trim() === '/');
    const unmatched = matches.filter(item => item.notes && !item.targets.length);
    state.workbook = templateWorkbook;
    state.templateBytes = templateBytes;
    draftEnabled = true;
    availableDraft = null;
    $('restore-draft').hidden = true;
    state.matches = matches;
    completedGroups.clear();
    $('overall-progress').value = 0;
    renderResults();
    $('next-step').textContent = `已匹配 ${ready.length} 人。核对下方名单后开始；有课堂评价时需填写口令。`;
    state.targets = targets;
    state.metadataTargets = metadataTargets;
    state.guidance = guidance;
    state.readyCount = ready.length;
    state.templateName = templateFile.name;
    $('download-excel').disabled = false;
    updateSelection();
    const parts = [`读取 ${records.length} 名学生，输出模板中有 ${targets.length} 个姓名行，可生成 ${ready.length} 人。`];
    if (noNotes.length) parts.push(`${noNotes.length} 人的“本周课堂评价”为空或为 /，所选智能将记 0 分并跳过生成。`);
    if (prepared.appended.length) parts.push(`已自动新增 ${prepared.appended.length} 名学生：${prepared.appended.map(item => item.student).join('、')}。`);
    if (unmatched.length) parts.push(`模板空白姓名行不足，未能加入：${unmatched.map(item => item.student).join('、')}。`);
    parts.push(`模板还可新增 ${prepared.remainingSlots} 名学生。`);
    parts.push(guidance ? `已读取“填写说明”（${guidance.length} 字）。` : '模板中未找到“填写说明”，将使用系统默认标准。');
    showSummary(parts.join(' '));
    $('batch-progress').textContent = '解析完成';
    await saveDraft();
  } catch (error) {
    state.workbook = null;
    state.matches = [];
    state.templateBytes = null;
    $('overwrite-existing').checked = false;
    state.readyCount = 0;
    showSummary(error?.message || 'Excel 解析失败，请检查文件格式。', 'error');
    $('batch-progress').textContent = '';
  } finally {
    parsing = false;
    for (const id of ['input-excel', 'output-template', 'restore-draft']) $(id).disabled = false;
    renderResults();
    button.disabled = false;
    updateSelection();
  }
});

async function requestIntelligences(student, notes, areas, guidance, noteMin, noteMax, accessCode) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 75_000);
  try {
  const response = await fetch('/api/evaluate-intelligences', {
    signal: controller.signal,
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-App-Password': accessCode },
    credentials: 'omit',
    body: JSON.stringify({ student, notes, areas, guidance, noteMin, noteMax }),
  });
  let payload = {};
  try { payload = await response.json(); } catch {}
  if (!response.ok) {
    const error = new Error(payload?.message || 'DeepSeek 解析失败。');
    error.status = response.status;
    throw error;
  }
  if (!Array.isArray(payload.intelligences)) throw new Error('DeepSeek 返回的数据格式不正确。');
  return payload.intelligences;
  } catch (error) {
    if (controller.signal.aborted) {
      const timeout = new Error('请求超时，请重试。');
      timeout.status = 504;
      throw timeout;
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function runBatch(preview = false, onlyStudent = null, onlyArea = null) {
  if (generating || parsing || connecting) return;
  if (!state.workbook) return showSummary('请先选择并匹配两个 Excel 文件。', 'error');
  const areas = onlyArea ? [onlyArea] : selectedAreas();
  if (!areas.length) return showSummary('请至少选择一个智能维度。', 'error');
  let ready = onlyStudent ? [onlyStudent] : selectedStudents();
  if (preview) ready = ready.filter(pendingFor).slice(0, 1);
  if (!ready.length) return showSummary('请勾选需要处理的学生。');
  const { noteMin, noteMax } = selectedNoteRange();
  const signature = `${noteMin}:${noteMax}`;
  const overwrite = Boolean(onlyStudent || $('overwrite-existing').checked);
  const regenerated = new Set();
  const isDone = (item, id) => onlyStudent ? regenerated.has(id) :
    protectedFor(item, id) || completedFor(item, id, signature);
  const needsModel = ready.some(item => areas.some(id => !isDone(item, id)) && item.notes?.trim() && item.notes.trim() !== '/');
  const accessCode = $('access-code').value.trim();
  if (needsModel) {
    connecting = true;
    updateSelection();
    try {
      if (!accessCode) {
        $('access-code').focus();
        return showSummary('请填写个人访问口令。', 'error');
      }
      await checkConnection();
      if (!modelReady) return showSummary('暂时无法连接生成服务，请稍后再试。', 'error');
    } finally { connecting = false; updateSelection(); }
  }
  const updateProgress = () => {
    const count = ready.reduce((sum, item) => sum + areas.filter(id => isDone(item, id)).length, 0);
    $('overall-progress').max = ready.length * areas.length || 1;
    $('overall-progress').value = count;
    $('next-step').textContent = `已完成 ${count} / ${ready.length * areas.length} 项。结果保存在当前页面，可随时下载。`;
  };
  generating = true;
  pauseRequested = false;
  const controls = [...document.querySelectorAll('input'), $('parse-excel'), $('select-all'), $('select-none'), $('students-all'), $('students-none'), $('student-filter'), $('restore-draft'), $('preview-one')];
  controls.forEach(control => { control.disabled = true; });
  $('generate-batch').disabled = true;
  $('pause-batch').hidden = false;
  $('pause-batch').disabled = false;
  $('pause-batch').textContent = '暂停';
  showSummary('正在生成。可以展开学生姓名查看结果；暂停会在当前组完成后生效。');
  updateProgress();
  renderResults();
  try {
    for (let index = 0; index < ready.length && !pauseRequested; index++) {
      const item = ready[index];
      item.error = '';
      const pending = areas.filter(id => !isDone(item, id));
      for (let offset = 0; offset < pending.length && !pauseRequested; offset += 2) {
        const group = pending.slice(offset, offset + 2);
        const failures = await generateDimensionGroups(
          group,
          ids => !item.notes?.trim() || item.notes.trim() === '/'
            ? Promise.resolve(ids.map(id => ({ id, level: 0, note: '' })))
            : requestIntelligences(item.student, item.notes, ids, state.guidance, noteMin, noteMax, accessCode),
          results => {
            const result = results.map(value => overwrite || item.results?.get(value.id)?.overwrite ? { ...value, overwrite: true } : value);
            result.forEach(value => regenerated.add(value.id));
            item.targets.forEach(target => writeIntelligenceResult(target, result));
            if (!completedGroups.has(item)) completedGroups.set(item, new Map());
            if (!item.results) item.results = new Map();
            result.forEach(value => {
              completedGroups.get(item).set(value.id, signature);
              item.results.set(value.id, value);
            });
            updateProgress();
            renderResults();
            saveDraft();
          },
          (ids, _, attempt) => {
            $('batch-progress').textContent = `${index + 1} / ${ready.length}：${item.student} · ${ids.map(id => DIMENSIONS.find(d => d.id === id)?.name).join('、')}${attempt > 1 ? '（重试中）' : ''}`;
          },
        );
        if (failures.length) {
          item.error = `部分未完成：${failures[0].error?.message || '请重试'}`;
          renderResults();
          if (/口令|API Key|余额|尚未配置/.test(item.error)) { pauseRequested = true; break; }
        }
      }
    }
    const remaining = ready.reduce((sum, item) => sum + areas.filter(id => !isDone(item, id)).length, 0);
    $('generate-batch').textContent = remaining ? '继续生成未完成项' : '补充生成所选智能';
    $('batch-progress').textContent = remaining ? `剩余 ${remaining} 项` : '全部完成';
    if (preview) {
      ready[0].expanded = true;
      renderResults();
      showSummary(remaining ? '预览尚未全部完成，可重试未完成项。' : '一人预览已完成。展开姓名核对，满意后点击批量生成继续其余学生；调整字数后可重新预览。');
    } else showSummary(remaining ? `还有 ${remaining} 项未完成。点击继续只补未完成项，已成功内容不会重复生成。可以先下载当前结果。` : '全部完成！展开学生姓名核对记录，然后下载填写结果。');
  } catch (error) {
    showSummary(`处理已停止，已完成结果保留：${error?.message || '请重试'}。可继续或下载。`, 'error');
  } finally {
    generating = false;
    controls.forEach(control => { control.disabled = false; });
    $('pause-batch').hidden = true;
    $('download-excel').disabled = false;
    updateSelection();
    renderResults();
    await saveDraft();
  }
}
$('generate-batch').addEventListener('click', () => runBatch(false));
$('preview-one').addEventListener('click', () => runBatch(true));

$('download-excel').addEventListener('click', async () => {
  if (!state.workbook) return;
  const button = $('download-excel');
  button.disabled = true;
  $('batch-progress').textContent = '正在生成文件…';
  try {
    const dateInput = $('meta-date').value;
    let formattedDate = '';
    if (dateInput) {
      const [year, month, day] = dateInput.split('-');
      formattedDate = `${year}年${Number(month)}月${Number(day)}日`;
    }
    writeTemplateMetadata(state.metadataTargets, {
      week: $('meta-week').value,
      date: formattedDate,
      teacher: $('meta-teacher').value.trim(),
      className: $('meta-class').value.trim(),
    });
    const buffer = await state.workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${state.templateName.replace(/\.xlsx$/i, '')}-已填写.xlsx`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    if (!generating) $('batch-progress').textContent = '下载完成';
  } catch (error) {
    showSummary(error?.message || '生成 Excel 失败。', 'error');
    $('batch-progress').textContent = '';
  } finally {
    button.disabled = false;
  }
});

checkConnection();

updateSelection();

(async () => {
  try {
    const draft = await draftStore.read();
    if (state.workbook || !draftEnabled) return;
    if (draft?.version === 1) {
      availableDraft = draft;
      $('restore-draft').hidden = false;
      $('draft-status').textContent = `找到上次草稿（${draft.records.length} 名学生），点击恢复后继续。`;
    } else $('draft-status').textContent = '匹配文件后自动保存本机草稿，支持刷新后恢复。';
  } catch { $('draft-status').textContent = '本机存储不可用，请及时下载结果。'; }
})();
