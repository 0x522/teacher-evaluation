// One local draft per browser origin. Never include credentials.
export function createDraftStore(factory = globalThis.indexedDB) {
  let queue = Promise.resolve();
  async function transact(mode, action) {
    if (!factory) throw new Error('浏览器不支持本机草稿');
    const db = await new Promise((resolve, reject) => {
      const open = factory.open('teacher-evaluation-drafts', 1);
      open.onupgradeneeded = () => open.result.createObjectStore('drafts');
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction('drafts', mode);
        const request = action(tx.objectStore('drafts'));
        tx.oncomplete = () => resolve(request.result);
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error || new Error('草稿保存中断'));
      });
    } finally { db.close(); }
  }
  function run(mode, action) {
    const next = queue.then(() => transact(mode, action));
    queue = next.catch(() => {});
    return next;
  }
  return {
    read: () => run('readonly', store => store.get('current')),
    write: draft => run('readwrite', store => store.put(draft, 'current')),
    clear: () => run('readwrite', store => store.delete('current')),
  };
}

export function snapshotDraft(state, completed, settings) {
  return {
    version: 1, savedAt: Date.now(), templateBytes: state.templateBytes,
    templateName: state.templateName, settings,
    records: state.matches.map(item => ({
      student: item.student, key: item.key, notes: item.notes,
      selected: item.selected !== false, error: item.error || '',
      results: [...(item.results?.values() || [])],
      completed: [...(completed.get(item)?.entries() || [])],
    })),
  };
}
