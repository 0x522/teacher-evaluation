// Keep responses small, retry only the failed group, and save each completed group.
export async function generateDimensionGroups(areas, request, save, progress, pause = ms => new Promise(resolve => setTimeout(resolve, ms))) {
  const failures = [];
  for (let offset = 0; offset < areas.length; offset += 2) {
    const group = areas.slice(offset, offset + 2);
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      progress(group, offset, attempt);
      try {
        const result = await request(group);
        if (!Array.isArray(result) || group.some(id => result.filter(item => item?.id === id).length !== 1 ||
          !result.some(item => item?.id === id && Number.isInteger(item.level) && item.level >= 0 && item.level <= 5 && typeof item.note === 'string' && (item.level === 0 || item.note.trim())))) {
          const error = new Error('返回的智能维度不完整或格式不正确。');
          error.status = 502;
          throw error;
        }
        save(result.filter(item => group.includes(item.id)));
        break;
      } catch (error) {
        const retryable = [429, 502, 503, 504].includes(error?.status) || error instanceof TypeError || /超时|连接|频繁|暂时不可用/.test(error?.message || '');
        if (!retryable || attempt === 2) {
          failures.push({ areas: group, error });
          // Authentication and configuration errors affect every remaining group.
          if ([401, 403, 503].includes(error?.status)) return failures;
          break;
        }
        await pause(1500);
      }
    }
  }
  return failures;
}
