export const DIMENSIONS = [
  { id: 'language', name: '语言智能' },
  { id: 'spatial', name: '空间智能' },
  { id: 'music', name: '音乐智能' },
  { id: 'logic', name: '逻辑智能' },
  { id: 'self', name: '内省智能' },
  { id: 'body', name: '运动智能' },
  { id: 'social', name: '人际智能' },
  { id: 'nature', name: '自然智能' },
];

export function normalizedText(value) {
  return String(value ?? '').replace(/[\s\u00a0]+/g, '').trim();
}

export function canonicalStudentName(value) {
  const text = String(value ?? '').trim();
  const chinese = text.match(/[\u3400-\u9fff·]+/g)?.join('') || '';
  return normalizedText(chinese || text).toLowerCase();
}

export function findInputRecords(workbook) {
  const records = [];
  for (const worksheet of workbook.worksheets) {
    let headerRow = 0;
    let nameColumn = 0;
    let notesColumn = 0;
    for (let row = 1; row <= Math.min(worksheet.rowCount, 20); row += 1) {
      worksheet.getRow(row).eachCell({ includeEmpty: false }, (cell, column) => {
        const header = normalizedText(cell.value);
        if (header === '孩子姓名' || header === '学生姓名' || header === '姓名') nameColumn = column;
        if (header === '本周课堂评价' || header === '课堂评价' || header === '课堂观察') notesColumn = column;
      });
      if (nameColumn && notesColumn) { headerRow = row; break; }
      nameColumn = 0;
      notesColumn = 0;
    }
    if (!headerRow) continue;
    for (let row = headerRow + 1; row <= worksheet.rowCount; row += 1) {
      const student = String(worksheet.getCell(row, nameColumn).text || '').trim();
      const notes = String(worksheet.getCell(row, notesColumn).text || '').trim();
      if (!student && !notes) continue;
      if (student) records.push({ student, key: canonicalStudentName(student), notes, sheetName: worksheet.name, row });
    }
  }
  if (!records.length) throw new Error('输入文件中没有找到学生数据。请使用“孩子姓名 / 本周课堂评价”表头。');
  return records;
}

function findTemplateStructure(workbook) {
  const sheets = [];
  for (const worksheet of workbook.worksheets) {
    let nameHeaderRow = 0;
    let nameColumn = 0;
    let serialColumn = 0;
    for (let row = 1; row <= Math.min(worksheet.rowCount, 20); row += 1) {
      worksheet.getRow(row).eachCell({ includeEmpty: false }, (cell, column) => {
        const header = normalizedText(cell.value);
        if (header === '序号') serialColumn = column;
        if (header === '孩子姓名') {
          nameHeaderRow = row;
          nameColumn = column;
        }
      });
      if (nameHeaderRow) break;
    }
    if (!nameHeaderRow) continue;
    const subheaderRow = nameHeaderRow + 1;
    const columns = {};
    for (let column = 1; column <= worksheet.columnCount; column += 1) {
      const dimensionText = normalizedText(worksheet.getCell(nameHeaderRow, column).value);
      const dimension = DIMENSIONS.find(item => dimensionText.startsWith(item.name));
      if (!dimension) continue;
      const levelText = normalizedText(worksheet.getCell(subheaderRow, column).value);
      const noteText = normalizedText(worksheet.getCell(subheaderRow, column + 1).value);
      if (levelText === '表现等级' && noteText === '亮点/观察记录') {
        columns[dimension.id] = { level: column, note: column + 1 };
      }
    }
    if (Object.keys(columns).length !== DIMENSIONS.length) continue;
    const targets = [];
    const slots = [];
    for (let row = subheaderRow + 1; row <= worksheet.rowCount; row += 1) {
      const student = String(worksheet.getCell(row, nameColumn).text || '').trim();
      const serial = serialColumn ? String(worksheet.getCell(row, serialColumn).text || '').trim() : '';
      if (serialColumn && !/^\d+$/.test(serial)) continue;
      const base = { worksheet, sheetName: worksheet.name, row, nameColumn, columns };
      if (student) targets.push({ ...base, student, key: canonicalStudentName(student) });
      else slots.push(base);
    }
    sheets.push({ worksheet, sheetName: worksheet.name, targets, slots });
  }
  if (!sheets.length) throw new Error('输出模板中没有找到八大智能记录表。');
  return sheets;
}

export function findTemplateTargets(workbook) {
  return findTemplateStructure(workbook).flatMap(sheet => sheet.targets);
}

export function prepareTemplateRecords(workbook, records) {
  const structure = findTemplateStructure(workbook);
  const targets = structure.flatMap(sheet => sheet.targets);
  const slots = structure.flatMap(sheet => sheet.slots);
  const existingKeys = new Set(targets.map(target => target.key));
  const newByKey = new Map();
  for (const record of records) {
    if (!existingKeys.has(record.key) && !newByKey.has(record.key)) newByKey.set(record.key, record);
  }
  const appended = [];
  const unassigned = [];
  for (const record of newByKey.values()) {
    const slot = slots.shift();
    if (!slot) {
      unassigned.push(record);
      continue;
    }
    slot.worksheet.getCell(slot.row, slot.nameColumn).value = record.student;
    const target = { ...slot, student: record.student, key: record.key };
    targets.push(target);
    appended.push({ student: record.student, sheetName: slot.sheetName, row: slot.row });
  }
  return { targets, appended, unassigned, remainingSlots: slots.length };
}

export function findTemplateMetadataTargets(workbook) {
  const fields = [
    { id: 'week', label: '周次' },
    { id: 'date', label: '记录日期' },
    { id: 'teacher', label: '带班老师' },
    { id: 'className', label: '班级' },
  ];
  const output = [];
  for (const worksheet of workbook.worksheets) {
    const hasGrowthTable = normalizedText(worksheet.getCell('B5').value) === '孩子姓名';
    if (!hasGrowthTable) continue;
    const targets = {};
    for (let row = 1; row <= Math.min(worksheet.rowCount, 10); row += 1) {
      for (let column = 1; column <= worksheet.columnCount; column += 1) {
        const cell = worksheet.getCell(row, column);
        const text = normalizedText(cell.value);
        const field = fields.find(item => text.includes(item.label));
        if (!field || targets[field.id]) continue;
        const headerMaster = cell.master.address;
        for (let next = column + 1; next <= worksheet.columnCount; next += 1) {
          const candidate = worksheet.getCell(row, next);
          if (candidate.master.address !== headerMaster) {
            targets[field.id] = candidate.master;
            break;
          }
        }
      }
    }
    if (Object.keys(targets).length) output.push({ worksheet, sheetName: worksheet.name, targets });
  }
  return output;
}

export function extractTemplateGuidance(workbook) {
  let sheets = workbook.worksheets.filter(worksheet => normalizedText(worksheet.name).includes('填写说明'));
  if (!sheets.length) {
    sheets = workbook.worksheets.filter(worksheet => {
      for (let row = 1; row <= Math.min(worksheet.rowCount, 30); row += 1) {
        for (let column = 1; column <= Math.min(worksheet.columnCount, 12); column += 1) {
          const text = normalizedText(worksheet.getCell(row, column).value);
          if (text.includes('八大智能核心解读') || text.includes('评分等级标准')) return true;
        }
      }
      return false;
    });
  }
  const lines = [];
  for (const worksheet of sheets) {
    lines.push(`工作表：${worksheet.name}`);
    for (let row = 1; row <= Math.min(worksheet.rowCount, 100); row += 1) {
      const values = [];
      for (let column = 1; column <= Math.min(worksheet.columnCount, 20); column += 1) {
        const text = String(worksheet.getCell(row, column).text || '').replace(/\s+/g, ' ').trim();
        if (text && !values.includes(text)) values.push(text);
      }
      if (values.length) lines.push(values.join(' | '));
    }
  }
  return lines.join('\n').slice(0, 6000);
}

export function writeTemplateMetadata(metadataTargets, values) {
  for (const item of metadataTargets) {
    if (values.week) item.targets.week && (item.targets.week.value = `第 ${values.week} 周`);
    if (values.date) item.targets.date && (item.targets.date.value = values.date);
    if (values.teacher) item.targets.teacher && (item.targets.teacher.value = values.teacher);
    if (values.className) item.targets.className && (item.targets.className.value = values.className);
  }
}

export function matchRecords(records, targets) {
  const targetByName = new Map();
  for (const target of targets) {
    if (!targetByName.has(target.key)) targetByName.set(target.key, []);
    targetByName.get(target.key).push(target);
  }
  return records.map(record => ({ ...record, targets: targetByName.get(record.key) || [] }));
}

export function clearIntelligenceResult(target, dimensionIds) {
  for (const id of dimensionIds) {
    const columns = target.columns[id];
    if (!columns) continue;
    target.worksheet.getCell(target.row, columns.level).value = null;
    target.worksheet.getCell(target.row, columns.note).value = null;
  }
}

export function writeIntelligenceResult(target, intelligences) {
  const byId = new Map(intelligences.map(item => [item.id, item]));
  for (const dimension of DIMENSIONS.filter(item => byId.has(item.id))) {
    const columns = target.columns[dimension.id];
    const item = byId.get(dimension.id);
    const hasEvidence = Number.isInteger(item?.level) && item.level >= 1 && item.level <= 5 && item.note;
    target.worksheet.getCell(target.row, columns.level).value = hasEvidence ? item.level : 0;
    target.worksheet.getCell(target.row, columns.note).value = hasEvidence ? String(item.note).trim() : null;
  }
}
