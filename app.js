const CHINA_HOLIDAYS_2026 = [
  '2026-01-01','2026-01-02','2026-01-03','2026-02-14','2026-02-15','2026-02-16','2026-02-17','2026-02-18','2026-02-19','2026-02-20',
  '2026-04-04','2026-04-05','2026-04-06','2026-05-01','2026-05-02','2026-05-03','2026-05-04','2026-05-05','2026-06-19','2026-06-20','2026-06-21',
  '2026-09-25','2026-09-26','2026-09-27','2026-10-01','2026-10-02','2026-10-03','2026-10-04','2026-10-05','2026-10-06','2026-10-07'
];
const CHINA_WORKDAYS_2026 = ['2026-01-04','2026-02-11','2026-02-22','2026-04-08','2026-04-26','2026-05-09','2026-06-28','2026-09-19','2026-10-10'];
const holidaySet = new Set(CHINA_HOLIDAYS_2026);
const workdayOverrideSet = new Set(CHINA_WORKDAYS_2026);

const DEFAULT_STAGES = [
  { id: 'stg_1', name: '提需+案子宣讲', days: 0, enabled: true, isDefault: false },
  { id: 'stg_2', name: '原画', days: 0, enabled: true, isDefault: false },
  { id: 'stg_3', name: '模型', days: 0, enabled: true, isDefault: false },
  { id: 'stg_4', name: '绑定', days: 0, enabled: true, isDefault: false },
  { id: 'stg_5', name: '动画', days: 0, enabled: true, isDefault: false },
  { id: 'stg_6', name: '特效', days: 0, enabled: true, isDefault: false }
];
const DEFAULT_DEPENDENCIES = [
  { id: 'dep_1', fromId: 'stg_1', toId: 'stg_2', linkType: 'FS', gapDays: 0 },
  { id: 'dep_2', fromId: 'stg_2', toId: 'stg_3', linkType: 'FS', gapDays: 0 },
  { id: 'dep_3', fromId: 'stg_3', toId: 'stg_4', linkType: 'FS', gapDays: 0 },
  { id: 'dep_4', fromId: 'stg_4', toId: 'stg_5', linkType: 'FS', gapDays: 0 },
  { id: 'dep_5', fromId: 'stg_5', toId: 'stg_6', linkType: 'FS', gapDays: 0 }
];

// 环节选项库 - 按大类分组
const STAGE_NAME_CATEGORIES = [
  { category: '运营/策划', items: [
    { name: '提需+案子宣讲', children: [] },
    { name: 'CE', children: [] }
  ]},
  { category: '美术', items: [
    { name: '原画', children: ['概念设计', '三视图', '染色设计'] },
    { name: '模型', children: ['中模', '高模', '低模', '烘焙', '贴图', 'LOD', '终版'] },
    { name: '绑定', children: [] },
    { name: '动画', children: ['动画初版', '动画终版'] },
    { name: '特效', children: ['特效初版', '特效终版'] },
    { name: '音频', children: [] }
  ]},
  { category: '程序', items: [
    { name: '服务端开发', children: [] },
    { name: '客户端开发', children: [] },
    { name: '程序联调', children: [] }
  ]},
  { category: 'UX', items: [
    { name: 'iCON', children: [] },
    { name: 'Layout', children: [] },
    { name: 'KV', children: [] }
  ]},
  { category: 'QA/测试', items: [
    { name: 'QA/测试', children: [] }
  ]}
];

// 兼容旧逻辑：扁平化为 STAGE_NAME_TREE
const STAGE_NAME_TREE = STAGE_NAME_CATEGORIES.flatMap((cat) => cat.items);

// 扁平化选项列表（兼容旧逻辑）
function flattenStageNameTree() {
  const list = [];
  STAGE_NAME_TREE.forEach((item) => {
    list.push(item.name);
    if (item.children && item.children.length) {
      item.children.forEach((child) => list.push(`${item.name}-${child}`));
    }
  });
  return list;
}
const DEFAULT_STAGE_NAME_OPTIONS = flattenStageNameTree();
const LINK_META = {
  FS: { label: '尾→首', arrow: '→', fromAnchor: '尾', toAnchor: '首', desc: '前置结束后，当前环节才能开始' },
  SS: { label: '首→首', arrow: '⇉', fromAnchor: '首', toAnchor: '首', desc: '前置开始后，当前环节才能开始' },
  FF: { label: '尾→尾', arrow: '⇇', fromAnchor: '尾', toAnchor: '尾', desc: '前置结束后，当前环节才能结束' },
  SF: { label: '首→尾', arrow: '←', fromAnchor: '首', toAnchor: '尾', desc: '前置开始后，当前环节才能结束' }
};
const STORAGE_KEYS = { templates: 'schedule_tool_templates_v2', stageNames: 'schedule_tool_stage_names_v2', stageNameAliases: 'schedule_tool_stage_name_aliases_v1', savedRecords: 'schedule_tool_records_v1', workspace: 'schedule_tool_workspace_v1' };

let stages = clone(DEFAULT_STAGES);
let dependencies = clone(DEFAULT_DEPENDENCIES);
let stageNameOptions = [];
let stageNameAliases = {};
let stageIdCounter = 100;
let dependencyIdCounter = 100;
let ganttZoom = 1;
let ganttZoomBwd = 1;
let autoCalcTimer = null;
let toastTimer = null;
let ocrToken = 0;
let historyStack = [];
let isRestoringHistory = false;
let importBuffer = { ocr: [], excel: [] };
let lastPrimaryData = null;
let lastCompareData = null;
let currentViewingRecord = null; // 当前正在查看的记录名

function clone(v) { return JSON.parse(JSON.stringify(v)); }
function safeInt(v, fallback = 0) { const n = parseInt(v, 10); return Number.isFinite(n) ? n : fallback; }
function esc(v) { return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
function fmt(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }
function parseDate(str) { const [y, m, d] = str.split('-').map(Number); return new Date(y, m - 1, d); }
function addDays(date, n) { const next = new Date(date); next.setDate(next.getDate() + n); return next; }
function mergeUnique(list) { const res = []; const seen = new Set(); list.forEach((item) => { const v = String(item || '').trim(); if (!v || seen.has(v)) return; seen.add(v); res.push(v); }); return res; }

function showToast(msg, type = 'success') {
  document.querySelectorAll('.toast').forEach((node) => node.remove());
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = msg;
  document.body.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('show'));
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 280); }, 2200);
}
function highlightEmptyField(id) {
  const el = document.getElementById(id);
  if (el) { el.classList.add('field-error-shake'); el.style.borderColor = '#ef4444'; setTimeout(() => { el.classList.remove('field-error-shake'); el.style.borderColor = ''; }, 1500); }
}
function modal(id, show) { const el = document.getElementById(id); if (el) el.style.display = show ? 'flex' : 'none'; }
function getJSON(key, fallback) { try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : fallback; } catch { return fallback; } }
function setJSON(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch { showToast('本地存储失败', 'error'); return false; } }

function createSnapshot() {
  return { stages: clone(stages), dependencies: clone(dependencies), stageNameOptions: clone(stageNameOptions), stageNameAliases: clone(stageNameAliases), stageIdCounter, dependencyIdCounter };
}
function pushHistory() {
  if (isRestoringHistory) return;
  historyStack.push(createSnapshot());
  if (historyStack.length > 80) historyStack.shift();
}
function undoLastAction() {
  if (!historyStack.length) { showToast('没有可撤销的操作', 'warning'); return; }
  const snapshot = historyStack.pop();
  isRestoringHistory = true;
  stages = clone(snapshot.stages); dependencies = clone(snapshot.dependencies);
  stageNameOptions = clone(snapshot.stageNameOptions);
  stageNameAliases = clone(snapshot.stageNameAliases || {});
  stageIdCounter = snapshot.stageIdCounter; dependencyIdCounter = snapshot.dependencyIdCounter;
  isRestoringHistory = false;
  renderAll(); autoRunCalculation(); showToast('已撤销上一步');
}

function isHoliday(date) { return holidaySet.has(fmt(date)); }
function isWeekend(date) { const day = date.getDay(); return day === 0 || day === 6; }
function isWorkday(date, rule) {
  const ds = fmt(date); const day = date.getDay();
  if (holidaySet.has(ds)) return false;
  if (workdayOverrideSet.has(ds)) return true;
  if (rule === 'single-rest') return day !== 0;
  return day !== 0 && day !== 6;
}
function ensureWorkday(date, rule, dir = 1) {
  let current = new Date(date);
  while (!isWorkday(current, rule)) current = addDays(current, dir >= 0 ? 1 : -1);
  return current;
}
function shiftWorkdays(date, steps, rule) {
  let current = ensureWorkday(date, rule, steps >= 0 ? 1 : -1);
  if (steps === 0) return current;
  const dir = steps > 0 ? 1 : -1; let remain = Math.abs(steps);
  while (remain > 0) { current = addDays(current, dir); current = ensureWorkday(current, rule, dir); remain -= 1; }
  return current;
}
function countWorkdaysInclusive(start, end, rule) {
  let total = 0; let current = new Date(start);
  while (current <= end) { if (isWorkday(current, rule)) total += 1; current = addDays(current, 1); }
  return total;
}

function getStageById(id) { return stages.find((s) => s.id === id) || null; }
function getCurrentDirection() { return document.querySelector('input[name="direction"]:checked')?.value || 'forward'; }
function getOppositeDirection(dir) { return dir === 'forward' ? 'backward' : 'forward'; }
function getDirectionLabel(dir) { return dir === 'forward' ? '正推' : '倒推'; }

function normalizeDependencies() {
  const stageIds = new Set(stages.map((s) => s.id)); const unique = new Set();
  dependencies = dependencies.filter((dep) => {
    // 无依赖：只需 toId 有效
    if (dep.fromId === 'none') {
      if (!stageIds.has(dep.toId)) return false;
      const key = `none__${dep.toId}`;
      if (unique.has(key)) return false; unique.add(key); return true;
    }
    if (!stageIds.has(dep.fromId) || !stageIds.has(dep.toId) || dep.fromId === dep.toId) return false;
    dep.linkType = LINK_META[dep.linkType] ? dep.linkType : 'FS';
    dep.gapDays = Math.max(0, safeInt(dep.gapDays, 0));
    const key = `${dep.fromId}__${dep.toId}`;
    if (unique.has(key)) return false; unique.add(key); return true;
  });
}
function getEnabledStages() { return stages.filter((s) => s.enabled); }
function getEnabledDependencies() {
  const ids = new Set(getEnabledStages().map((s) => s.id));
  return dependencies.filter((d) => d.fromId !== 'none' && ids.has(d.fromId) && ids.has(d.toId));
}
function hasCycle(customDeps = dependencies) {
  const enabled = getEnabledStages(); const ids = new Set(enabled.map((s) => s.id));
  const graph = {}; const indeg = {};
  enabled.forEach((s) => { graph[s.id] = []; indeg[s.id] = 0; });
  customDeps.forEach((d) => { if (d.fromId === 'none' || !ids.has(d.fromId) || !ids.has(d.toId)) return; graph[d.fromId].push(d.toId); indeg[d.toId] += 1; });
  const queue = Object.keys(indeg).filter((id) => indeg[id] === 0); let count = 0;
  while (queue.length) { const cur = queue.shift(); count += 1; graph[cur].forEach((n) => { indeg[n] -= 1; if (indeg[n] === 0) queue.push(n); }); }
  return count !== enabled.length;
}
function buildTopologicalOrder(enabledStages, enabledDeps) {
  const graph = {}; const indeg = {};
  enabledStages.forEach((s) => { graph[s.id] = []; indeg[s.id] = 0; });
  enabledDeps.forEach((d) => { if (d.fromId === 'none') return; graph[d.fromId].push(d.toId); indeg[d.toId] += 1; });
  const order = []; const queue = enabledStages.filter((s) => indeg[s.id] === 0).map((s) => s.id);
  while (queue.length) { const cur = queue.shift(); order.push(cur); graph[cur].forEach((n) => { indeg[n] -= 1; if (indeg[n] === 0) queue.push(n); }); }
  return order.length === enabledStages.length ? order : null;
}

function loadStageNameOptions() {
  stageNameAliases = getJSON(STORAGE_KEYS.stageNameAliases, {});
  const savedNames = getJSON(STORAGE_KEYS.stageNames, []).map(resolveStageNameAlias);
  const defaultNames = DEFAULT_STAGE_NAME_OPTIONS.map(resolveStageNameAlias);
  const defaultStageNames = DEFAULT_STAGES.map((s) => resolveStageNameAlias(s.name));
  stageNameOptions = mergeUnique([...savedNames, ...defaultNames, ...defaultStageNames]);
  setJSON(STORAGE_KEYS.stageNames, stageNameOptions);
}
function saveStageNameOptions() { setJSON(STORAGE_KEYS.stageNames, stageNameOptions); }
function saveStageNameAliases() { setJSON(STORAGE_KEYS.stageNameAliases, stageNameAliases); }
function resolveStageNameAlias(name) {
  const key = String(name || '').trim();
  return stageNameAliases[key] || key;
}
function getDefaultStagesWithAliases() {
  return DEFAULT_STAGES.map((stage) => ({ ...clone(stage), name: resolveStageNameAlias(stage.name) }));
}
function getStageNameAliasSource(name) {
  const value = String(name || '').trim();
  const pair = Object.entries(stageNameAliases).find(([, alias]) => alias === value);
  return pair ? pair[0] : value;
}
function resolveStageTreeOptionName(parentName, childName = '') {
  const parent = String(parentName || '').trim();
  const child = String(childName || '').trim();
  if (!child) return resolveStageNameAlias(parent);
  const fullName = `${parent}-${child}`;
  return stageNameAliases[fullName] || `${resolveStageNameAlias(parent)}-${child}`;
}
function getManagedDefaultNameSet() {
  const names = new Set();
  STAGE_NAME_TREE.forEach((item) => {
    names.add(resolveStageTreeOptionName(item.name));
    (item.children || []).forEach((child) => names.add(resolveStageTreeOptionName(item.name, child)));
  });
  return names;
}
function renameManagedStageName(oldName, newName) {
  const from = String(oldName || '').trim();
  const to = String(newName || '').trim();
  if (!from || !to || from === to) return false;

  const source = getStageNameAliasSource(from);
  const isDefaultOption = DEFAULT_STAGE_NAME_OPTIONS.includes(source) || DEFAULT_STAGES.some((s) => s.name === source);
  if (isDefaultOption) {
    if (to === source) delete stageNameAliases[source];
    else stageNameAliases[source] = to;
    saveStageNameAliases();
  }

  const namesToReplace = new Set([from, source, stageNameAliases[source]].filter(Boolean));
  stageNameOptions = mergeUnique(stageNameOptions.map((name) => (namesToReplace.has(name) ? to : resolveStageNameAlias(name))));
  const changedStages = stages.reduce((count, stage) => {
    if (!namesToReplace.has(stage.name)) return count;
    stage.name = to;
    return count + 1;
  }, 0);
  saveStageNameOptions();
  if (changedStages) {
    renderStages();
    renderDependencies();
    autoRunCalculation();
    saveWorkspace();
  }
  return changedStages > 0;
}

/* ============ 工作区状态自动持久化 ============ */
let _saveWorkspaceTimer = null;
function saveWorkspace() {
  clearTimeout(_saveWorkspaceTimer);
  _saveWorkspaceTimer = setTimeout(() => {
    setJSON(STORAGE_KEYS.workspace, {
      stages: clone(stages),
      dependencies: clone(dependencies),
      stageNameOptions: clone(stageNameOptions),
      stageNameAliases: clone(stageNameAliases),
      stageIdCounter,
      dependencyIdCounter,
      settings: typeof collectSettings === 'function' ? collectSettings() : {}
    });
  }, 300);
}
function clearWorkspace() {
  clearTimeout(_saveWorkspaceTimer);
  try { localStorage.removeItem(STORAGE_KEYS.workspace); } catch {}
}
function loadWorkspace() {
  const ws = getJSON(STORAGE_KEYS.workspace, null);
  if (!ws || !Array.isArray(ws.stages) || ws.stages.length === 0) return false;
  stages = clone(ws.stages);
  dependencies = clone(ws.dependencies || []);
  if (ws.stageNameAliases) {
    stageNameAliases = clone(ws.stageNameAliases);
    saveStageNameAliases();
  }
  if (Array.isArray(ws.stageNameOptions)) {
    stageNameOptions = mergeUnique([...ws.stageNameOptions.map(resolveStageNameAlias), ...stageNameOptions.map(resolveStageNameAlias), ...stages.map((s) => s.name)]);
    saveStageNameOptions();
  }
  stageIdCounter = ws.stageIdCounter || 100;
  dependencyIdCounter = ws.dependencyIdCounter || 100;
  normalizeDependencies();
  return ws.settings || null;
}

function upsertStageNameOption(name) {
  const v = String(name || '').trim(); if (!v) return;
  if (!stageNameOptions.includes(v)) { stageNameOptions.push(v); saveStageNameOptions(); }
}

function renderStageNameDropdown(keyword = '') {
  const list = document.getElementById('stage-name-dropdown'); if (!list) return;
  const q = keyword.trim().toLowerCase();

  // 构建分类菜单 HTML
  let html = '';
  STAGE_NAME_CATEGORIES.forEach((cat, catIdx) => {
    let catHtml = '';
    cat.items.forEach((item) => {
      const optionName = resolveStageTreeOptionName(item.name);
      const nameMatch = !q || optionName.toLowerCase().includes(q) || item.name.toLowerCase().includes(q);
      const childMatches = (item.children || []).filter((c) => {
        const childOptionName = resolveStageTreeOptionName(item.name, c);
        return !q || childOptionName.toLowerCase().includes(q) || c.toLowerCase().includes(q) || item.name.toLowerCase().includes(q);
      });
      if (!nameMatch && !childMatches.length) return;

      catHtml += `<div class="dropdown-item dropdown-parent" data-option-name="${esc(optionName)}"><span>${esc(optionName)}</span>${item.children && item.children.length ? '<span class="dropdown-arrow">▸</span>' : ''}</div>`;
      if (item.children && item.children.length) {
        const showChildren = nameMatch || childMatches.length;
        if (showChildren) {
          item.children.forEach((child) => {
            const childOptionName = resolveStageTreeOptionName(item.name, child);
            if (q && !childOptionName.toLowerCase().includes(q) && !child.toLowerCase().includes(q) && !item.name.toLowerCase().includes(q)) return;
            catHtml += `<div class="dropdown-item dropdown-child" data-option-name="${esc(childOptionName)}"><span class="dropdown-child-prefix">└</span> ${esc(childOptionName.replace(`${optionName}-`, ''))}</div>`;
          });
        }
      }
    });
    if (catHtml) {
      if (html) html += '<div class="dropdown-divider"></div>';
      html += `<div class="dropdown-category-label">${esc(cat.category)}</div>`;
      html += catHtml;
    }
  });

  // 追加自定义选项（不在树中的）
  const treeNames = getManagedDefaultNameSet();
  const customOptions = stageNameOptions.filter((n) => !treeNames.has(n) && (!q || n.toLowerCase().includes(q)));
  if (customOptions.length) {
    if (html) html += '<div class="dropdown-divider"></div>';
    html += '<div class="dropdown-category-label">自定义</div>';
    customOptions.forEach((n) => {
      html += `<div class="dropdown-item" data-option-name="${esc(n)}"><span>${esc(n)}</span></div>`;
    });
  }

  list.innerHTML = html || '<div class="dropdown-empty">没有匹配项，可直接输入新名称</div>';
  list.querySelectorAll('[data-option-name]').forEach((item) => {
    item.addEventListener('click', () => { document.getElementById('new-stage-name').value = item.dataset.optionName; toggleStageNameDropdown(false); });
  });
}
function toggleStageNameDropdown(show) { const l = document.getElementById('stage-name-dropdown'); if (l) l.style.display = show ? 'block' : 'none'; }

function renderManageNameOptions() {
  const wrap = document.getElementById('name-options-list'); if (!wrap) return;
  wrap.innerHTML = stageNameOptions.length
    ? stageNameOptions.map((name, i) => `<div class="name-option-item" data-name-index="${i}" draggable="true"><span class="name-option-handle" title="拖拽排序">≡</span><input type="text" value="${esc(name)}" data-name-input="${i}"><button class="btn-danger-sm" data-name-delete="${i}">删除</button></div>`).join('')
    : '<p class="placeholder-text">暂无选项</p>';
  wrap.querySelectorAll('[data-name-input]').forEach((input) => {
    input.addEventListener('change', (e) => {
      const idx = safeInt(e.target.dataset.nameInput, -1); const val = e.target.value.trim();
      if (idx < 0) return; if (!val) { showToast('选项名称不能为空', 'warning'); e.target.value = stageNameOptions[idx]; return; }
      const oldName = stageNameOptions[idx];
      const changedStages = renameManagedStageName(oldName, val);
      renderManageNameOptions(); renderStageNameDropdown(document.getElementById('new-stage-name').value || '');
      showToast(changedStages ? '环节名称已同步更新' : '选项名称已更新');
    });
  });
  wrap.querySelectorAll('[data-name-delete]').forEach((btn) => {
    btn.addEventListener('click', () => { const idx = safeInt(btn.dataset.nameDelete, -1); if (idx < 0) return; stageNameOptions.splice(idx, 1); saveStageNameOptions(); renderManageNameOptions(); renderStageNameDropdown(document.getElementById('new-stage-name').value || ''); });
  });
  let dragIdx = -1;
  wrap.querySelectorAll('.name-option-item').forEach((item) => {
    item.addEventListener('dragstart', () => { dragIdx = safeInt(item.dataset.nameIndex, -1); item.classList.add('dragging-option'); });
    item.addEventListener('dragend', () => { dragIdx = -1; item.classList.remove('dragging-option'); wrap.querySelectorAll('.name-option-item').forEach((n) => n.classList.remove('drag-over')); });
    item.addEventListener('dragover', (e) => { e.preventDefault(); wrap.querySelectorAll('.name-option-item').forEach((n) => n.classList.remove('drag-over')); item.classList.add('drag-over'); });
    item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
    item.addEventListener('drop', (e) => {
      e.preventDefault(); item.classList.remove('drag-over');
      const dropIdx = safeInt(item.dataset.nameIndex, -1);
      if (dragIdx < 0 || dropIdx < 0 || dragIdx === dropIdx) return;
      const [moved] = stageNameOptions.splice(dragIdx, 1); stageNameOptions.splice(dropIdx, 0, moved);
      saveStageNameOptions(); renderManageNameOptions(); renderStageNameDropdown(document.getElementById('new-stage-name').value || '');
    });
  });
}

function getTemplates() { return getJSON(STORAGE_KEYS.templates, []); }
function saveTemplates(list) { return setJSON(STORAGE_KEYS.templates, list); }
function collectSettings() {
  return { direction: getCurrentDirection(), baseDate: document.getElementById('base-date').value, extraDays: document.getElementById('extra-days').value, workday: document.querySelector('input[name="workday"]:checked')?.value || 'double-rest', expectedEndDate: document.getElementById('expected-end-date')?.value || '' };
}
function applySettings(settings = {}) {
  const dir = settings.direction || 'forward';
  const dirInput = document.querySelector(`input[name="direction"][value="${dir}"]`); if (dirInput) dirInput.checked = true;
  const wd = settings.workday || 'double-rest';
  const wdInput = document.querySelector(`input[name="workday"][value="${wd}"]`); if (wdInput) wdInput.checked = true;
  document.getElementById('base-date').value = settings.baseDate || fmt(new Date());
  document.getElementById('extra-days').value = settings.extraDays ?? 0;
  const expectedEndInput = document.getElementById('expected-end-date');
  if (expectedEndInput) expectedEndInput.value = settings.expectedEndDate || '';
  syncDirectionTexts(); bindCompareToggle();
}
function refreshTemplateSelect(selected = '') {
  const sel = document.getElementById('template-select'); if (!sel) return;
  const tpls = getTemplates().sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  sel.innerHTML = '<option value="">-- 选择标效模板 --</option>' + tpls.map((t) => `<option value="${esc(t.name)}">${esc(t.name)}</option>`).join('');
  if (selected) sel.value = selected;
}
function saveCurrentTemplate(name) {
  const n = String(name || '').trim(); if (!n) { showToast('请输入模板名称', 'warning'); return; }
  const payload = { name: n, updatedAt: Date.now(), data: { stages: clone(stages), dependencies: clone(dependencies), stageNameOptions: clone(stageNameOptions), stageNameAliases: clone(stageNameAliases), stageIdCounter, dependencyIdCounter, settings: collectSettings() } };
  const tpls = getTemplates(); const idx = tpls.findIndex((t) => t.name === n);
  if (idx >= 0) tpls[idx] = payload; else tpls.push(payload);
  if (!saveTemplates(tpls)) return;
  refreshTemplateSelect(n); modal('modal-save-tpl', false); showToast(`标效模板「${n}」已保存`);
}
function loadSelectedTemplate() {
  const name = document.getElementById('template-select').value;
  if (!name) { showToast('请先选择标效模板', 'warning'); return; }
  const tpl = getTemplates().find((t) => t.name === name);
  if (!tpl) { showToast('模板不存在', 'error'); refreshTemplateSelect(); return; }
  pushHistory();
  stages = clone(tpl.data.stages || DEFAULT_STAGES); dependencies = clone(tpl.data.dependencies || DEFAULT_DEPENDENCIES);
  stageNameAliases = clone(tpl.data.stageNameAliases || stageNameAliases || {});
  saveStageNameAliases();
  stageNameOptions = mergeUnique([...(tpl.data.stageNameOptions || []).map(resolveStageNameAlias), ...DEFAULT_STAGE_NAME_OPTIONS.map(resolveStageNameAlias), ...stages.map((s) => s.name)]);
  stageIdCounter = tpl.data.stageIdCounter || 100; dependencyIdCounter = tpl.data.dependencyIdCounter || 100;
  normalizeDependencies(); saveStageNameOptions(); applySettings(tpl.data.settings || {});
  renderAll(); autoRunCalculation(); showToast(`已加载标效模板「${name}」`);
}
function deleteSelectedTemplate() {
  const name = document.getElementById('template-select').value;
  if (!name) { showToast('请先选择标效模板', 'warning'); return; }
  if (!saveTemplates(getTemplates().filter((t) => t.name !== name))) return;
  refreshTemplateSelect(); showToast(`已删除标效模板「${name}」`);
}

/* ============ 流程环节列表（纯环节管理） ============ */
function renderStages() {
  normalizeDependencies();
  const list = document.getElementById('stage-list'); if (!list) return;
  list.innerHTML = stages.map((stage, idx) => {
    return `
      <div class="stage-item${stage.enabled ? '' : ' disabled'}" data-stage-id="${stage.id}" data-stage-index="${idx}" draggable="true">
        <div class="stage-item-header">
          <div class="drag-handle" title="拖拽排序">≡</div>
          <span class="stage-order">${idx + 1}</span>
          <div class="stage-name-dropdown-wrap" data-stage-dropdown="${stage.id}">
            <input class="stage-name-input" type="text" value="${esc(stage.name)}" data-stage-name="${stage.id}" autocomplete="off" placeholder="选择或输入环节名称">
            <button class="stage-name-dropdown-btn" type="button" data-stage-dropdown-toggle="${stage.id}">▼</button>
            <div class="stage-name-dropdown-list" data-stage-dropdown-list="${stage.id}" style="display:none;"></div>
          </div>
          <div class="stage-days-box${stage.days ? '' : ' stage-days-empty'}">
            <span>工期</span>
            <div class="stage-days-dropdown-wrap">
              <input type="text" inputmode="numeric" value="${stage.days || ''}" data-stage-days="${stage.id}" placeholder="必填" readonly>
              <div class="stage-days-dropdown-list" data-days-dropdown-list="${stage.id}" style="display:none;"></div>
            </div>
            <em>天</em>
          </div>
          <label class="toggle"><input type="checkbox" ${stage.enabled ? 'checked' : ''} data-stage-toggle="${stage.id}"><span class="toggle-slider"></span></label>
          ${stage.isDefault ? '' : `<button class="btn-danger-sm" data-stage-delete="${stage.id}" title="删除">×</button>`}
        </div>
      </div>`;
  }).join('');
  bindStageEvents();
  bindStageDragSort();
  bindStageNameDropdowns();
}

/* 为流程环节卡片中的环节名称渲染下拉选项 */
function renderStageCardDropdown(listEl, keyword, stageId) {
  if (!listEl) return;
  const q = (keyword || '').trim().toLowerCase();

  let html = '';
  STAGE_NAME_CATEGORIES.forEach((cat) => {
    let catHtml = '';
    cat.items.forEach((item) => {
      const optionName = resolveStageTreeOptionName(item.name);
      const nameMatch = !q || optionName.toLowerCase().includes(q) || item.name.toLowerCase().includes(q);
      const childMatches = (item.children || []).filter((c) => {
        const childOptionName = resolveStageTreeOptionName(item.name, c);
        return !q || childOptionName.toLowerCase().includes(q) || c.toLowerCase().includes(q) || item.name.toLowerCase().includes(q);
      });
      if (!nameMatch && !childMatches.length) return;

      catHtml += `<div class="dropdown-item dropdown-parent" data-card-option="${esc(optionName)}"><span>${esc(optionName)}</span>${item.children && item.children.length ? '<span class="dropdown-arrow">▸</span>' : ''}</div>`;
      if (item.children && item.children.length) {
        const showChildren = nameMatch || childMatches.length;
        if (showChildren) {
          item.children.forEach((child) => {
            const childOptionName = resolveStageTreeOptionName(item.name, child);
            if (q && !childOptionName.toLowerCase().includes(q) && !child.toLowerCase().includes(q) && !item.name.toLowerCase().includes(q)) return;
            catHtml += `<div class="dropdown-item dropdown-child" data-card-option="${esc(childOptionName)}"><span class="dropdown-child-prefix">└</span> ${esc(childOptionName.replace(`${optionName}-`, ''))}</div>`;
          });
        }
      }
    });
    if (catHtml) {
      if (html) html += '<div class="dropdown-divider"></div>';
      html += `<div class="dropdown-category-label">${esc(cat.category)}</div>`;
      html += catHtml;
    }
  });

  // 追加自定义选项
  const treeNames = getManagedDefaultNameSet();
  const customOptions = stageNameOptions.filter((n) => !treeNames.has(n) && (!q || n.toLowerCase().includes(q)));
  if (customOptions.length) {
    if (html) html += '<div class="dropdown-divider"></div>';
    html += '<div class="dropdown-category-label">自定义</div>';
    customOptions.forEach((n) => {
      html += `<div class="dropdown-item" data-card-option="${esc(n)}"><span>${esc(n)}</span></div>`;
    });
  }

  listEl.innerHTML = html || '<div class="dropdown-empty">没有匹配项，可直接输入新名称</div>';
  listEl.querySelectorAll('[data-card-option]').forEach((item) => {
    item.addEventListener('mousedown', (e) => {
      e.preventDefault(); // 阻止input失焦
      const name = item.dataset.cardOption;
      const input = listEl.closest('.stage-name-dropdown-wrap').querySelector('.stage-name-input');
      if (input) {
        input.value = name;
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }
      listEl.style.display = 'none';
    });
  });
}

function toggleStageCardDropdown(stageId, show) {
  const listEl = document.querySelector(`[data-stage-dropdown-list="${stageId}"]`);
  if (listEl) listEl.style.display = show ? 'block' : 'none';
}

function bindStageNameDropdowns() {
  // 为每个卡片的环节名称绑定下拉交互
  document.querySelectorAll('.stage-name-dropdown-wrap').forEach((wrap) => {
    const stageId = wrap.dataset.stageDropdown;
    const input = wrap.querySelector('.stage-name-input');
    const listEl = wrap.querySelector('.stage-name-dropdown-list');
    const toggleBtn = wrap.querySelector('.stage-name-dropdown-btn');

    if (!input || !listEl) return;

    function showDropdown() {
      // 先关闭其他所有打开的下拉框
      document.querySelectorAll('.stage-name-dropdown-list').forEach((el) => { if (el !== listEl) el.style.display = 'none'; });
      renderStageCardDropdown(listEl, input.value, stageId);
      listEl.style.display = 'block';
    }

    input.addEventListener('click', showDropdown);
    input.addEventListener('focus', showDropdown);
    input.addEventListener('input', () => {
      renderStageCardDropdown(listEl, input.value, stageId);
      listEl.style.display = 'block';
    });
    input.addEventListener('blur', () => {
      // 延迟关闭，让mousedown事件有机会触发
      setTimeout(() => { listEl.style.display = 'none'; }, 150);
    });

    if (toggleBtn) {
      toggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const shown = listEl.style.display === 'block';
        if (!shown) showDropdown();
        else listEl.style.display = 'none';
      });
    }
  });

  // 全局点击关闭
  // 注意：避免重复绑定，使用标记
  if (!bindStageNameDropdowns._globalBound) {
    bindStageNameDropdowns._globalBound = true;
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.stage-name-dropdown-wrap')) {
        document.querySelectorAll('.stage-name-dropdown-list').forEach((el) => { el.style.display = 'none'; });
      }
      if (!e.target.closest('.stage-days-dropdown-wrap')) {
        document.querySelectorAll('.stage-days-dropdown-list').forEach((el) => { el.style.display = 'none'; });
      }
    });
  }
}

function bindStageEvents() {
  document.querySelectorAll('[data-stage-name]').forEach((input) => {
    input.addEventListener('change', (e) => {
      const stage = getStageById(e.target.dataset.stageName); if (!stage) return;
      const v = e.target.value.trim();
      if (!v) { e.target.value = stage.name; showToast('环节名称不能为空', 'warning'); return; }
      pushHistory(); stage.name = v; upsertStageNameOption(v); renderStages(); renderDependencies(); autoRunCalculation();
    });
  });
  document.querySelectorAll('[data-stage-days]').forEach((input) => {
    input.addEventListener('input', (e) => {
      const stage = getStageById(e.target.dataset.stageDays); if (!stage) return;
      const val = e.target.value.trim();
      const box = e.target.closest('.stage-days-box');
      if (!val || safeInt(val, 0) <= 0) {
        stage.days = 0;
        if (box) box.classList.add('stage-days-empty');
      } else {
        stage.days = Math.max(1, safeInt(val, 1));
        if (box) box.classList.remove('stage-days-empty');
      }
      autoRunCalculation();
    });
    // 聚焦时隐藏placeholder
    input.addEventListener('focus', () => { input.placeholder = ''; });
    // 失焦时如果没填写，恢复placeholder
    input.addEventListener('blur', () => { if (!input.value.trim()) input.placeholder = '必填'; });
  });
  document.querySelectorAll('[data-stage-toggle]').forEach((toggle) => {
    toggle.addEventListener('change', (e) => {
      const stage = getStageById(e.target.dataset.stageToggle); if (!stage) return;
      pushHistory(); stage.enabled = e.target.checked; renderStages(); renderDependencies(); autoRunCalculation();
    });
  });
  document.querySelectorAll('[data-stage-delete]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const sid = btn.dataset.stageDelete; pushHistory();
      stages = stages.filter((s) => s.id !== sid);
      dependencies = dependencies.filter((d) => d.fromId !== sid && d.toId !== sid);
      renderStages(); renderDependencies(); autoRunCalculation(); showToast('环节已删除');
    });
  });
  // 工期下拉选择数字：点击input展开
  document.querySelectorAll('[data-stage-days]').forEach((input) => {
    input.addEventListener('click', (e) => {
      e.stopPropagation();
      const stageId = input.dataset.stageDays;
      const listEl = document.querySelector(`[data-days-dropdown-list="${stageId}"]`);
      if (!listEl) return;
      const isOpen = listEl.style.display !== 'none';
      // 关闭其他所有下拉
      document.querySelectorAll('.stage-days-dropdown-list').forEach((el) => { el.style.display = 'none'; });
      if (isOpen) return;
      // 生成选项：1~60天
      const currentVal = safeInt(input.value, 0);
      listEl.innerHTML = '';
      for (let i = 1; i <= 60; i++) {
        const opt = document.createElement('div');
        opt.className = 'days-dropdown-item' + (i === currentVal ? ' days-dropdown-active' : '');
        opt.textContent = i;
        opt.addEventListener('click', () => {
          input.value = i; input.dispatchEvent(new Event('input', { bubbles: true }));
          listEl.style.display = 'none';
        });
        listEl.appendChild(opt);
      }
      listEl.style.display = 'block';
      // 滚动到当前选中值
      if (currentVal > 0) {
        const activeEl = listEl.querySelector('.days-dropdown-active');
        if (activeEl) activeEl.scrollIntoView({ block: 'center' });
      }
    });
  });
}

function bindStageDragSort() {
  const list = document.getElementById('stage-list'); if (!list) return;
  let dragIdx = -1;
  let dragFromHandle = false;
  list.querySelectorAll('.stage-item').forEach((item) => {
    // mousedown 时记录是否点击在 handle 上
    item.addEventListener('mousedown', (e) => {
      dragFromHandle = !!e.target.closest('.drag-handle');
    });
    item.addEventListener('dragstart', (e) => {
      if (!dragFromHandle) { e.preventDefault(); return; }
      dragIdx = safeInt(item.dataset.stageIndex, -1); item.classList.add('dragging-stage');
      e.dataTransfer.effectAllowed = 'move';
    });
    item.addEventListener('dragend', () => { dragIdx = -1; dragFromHandle = false; item.classList.remove('dragging-stage'); list.querySelectorAll('.stage-item').forEach((n) => n.classList.remove('drag-over')); });
    item.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; list.querySelectorAll('.stage-item').forEach((n) => n.classList.remove('drag-over')); item.classList.add('drag-over'); });
    item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
    item.addEventListener('drop', (e) => {
      e.preventDefault(); item.classList.remove('drag-over');
      const dropIdx = safeInt(item.dataset.stageIndex, -1);
      if (dragIdx < 0 || dropIdx < 0 || dragIdx === dropIdx) return;
      pushHistory(); const [moved] = stages.splice(dragIdx, 1); stages.splice(dropIdx, 0, moved);
      renderStages(); renderDependencies(); autoRunCalculation();
    });
  });
}

/* ============ 流程依赖关系（独立 section） ============ */
function isDepAdvancedMode() {
  const chk = document.getElementById('chk-dep-advanced');
  return chk ? chk.checked : false;
}

const LINK_ORDER = ['FS', 'SS', 'FF', 'SF'];

function renderDependencies() {
  normalizeDependencies();
  const list = document.getElementById('dependency-list'); if (!list) return;
  const advanced = isDepAdvancedMode();
  if (!dependencies.length) {
    list.innerHTML = '<div class="dep-empty-tip">暂无依赖关系。点击「+ 添加依赖」来设置环节之间的先后关系。</div>';
    bindDepEvents();
    return;
  }
  list.innerHTML = dependencies.map((dep, idx) => {
    const allStages = stages;
    const meta = LINK_META[dep.linkType] || LINK_META.FS;
    const isNone = dep.fromId === 'none';
    return `
      <div class="dep-row-compact${isNone ? ' dep-row-none' : ''}" data-dependency-id="${dep.id}" data-dep-index="${idx}" draggable="true">
        <div class="dep-drag-handle" title="拖拽排序">≡</div>
        <select class="dep-sel${isNone ? ' dep-sel-none' : ''}" data-dep-from="${dep.id}">
          <option value="none" ${isNone ? 'selected' : ''} class="dep-option-none">— 无依赖 —</option>
          ${allStages.map((s) => `<option value="${s.id}" ${s.id === dep.fromId ? 'selected' : ''}>${esc(s.name)}${s.enabled ? '' : '（已停用）'}</option>`).join('')}
        </select>
        <span class="dep-arrow-btn${isNone ? ' dep-arrow-disabled' : ''}" data-dep-arrow="${dep.id}" title="点击切换连接方式">
          <span class="dep-arrow-icon">${isNone ? '—' : meta.arrow}</span>
          <span class="dep-arrow-tip">${isNone ? '无依赖' : esc(meta.label)}</span>
        </span>
        <select class="dep-sel" data-dep-to="${dep.id}">
          ${allStages.map((s) => `<option value="${s.id}" ${s.id === dep.toId ? 'selected' : ''}>${esc(s.name)}${s.enabled ? '' : '（已停用）'}</option>`).join('')}
        </select>
        ${advanced && !isNone ? `<select class="dep-link-sel" data-dep-link="${dep.id}">
          ${Object.entries(LINK_META).map(([k, m]) => `<option value="${k}" ${k === dep.linkType ? 'selected' : ''} title="${esc(m.desc)}">${esc(m.label)}</option>`).join('')}
        </select>` : ''}
        ${!isNone ? `<div class="dep-gap-compact">
          <span class="dep-gap-label">间隔</span>
          <input type="number" min="0" value="${dep.gapDays}" class="dep-gap-num" data-dep-gap="${dep.id}">
          <span class="dep-gap-unit">天</span>
        </div>` : ''}
        <button class="dep-del-btn" data-dep-delete="${dep.id}" title="删除此依赖">×</button>
      </div>`;
  }).join('');
  bindDepEvents();
  bindDepDragSort();
}

function findNextDependencyPair() {
  const availableStages = stages.filter((s) => s.enabled);
  if (availableStages.length < 2) return null;
  const existingPairs = new Set(dependencies.map((d) => `${d.fromId}__${d.toId}`));
  const candidates = [];

  for (let i = 0; i < availableStages.length - 1; i++) {
    candidates.push([availableStages[i], availableStages[i + 1]]);
  }
  availableStages.forEach((from) => {
    availableStages.forEach((to) => {
      if (from.id !== to.id) candidates.push([from, to]);
    });
  });

  for (const [from, to] of candidates) {
    if (existingPairs.has(`${from.id}__${to.id}`)) continue;
    const testDeps = [...dependencies, { id: '__test__', fromId: from.id, toId: to.id, linkType: 'FS', gapDays: 0 }];
    if (!hasCycle(testDeps)) return { from, to };
  }
  return null;
}

function bindDepEvents() {
  // 添加依赖按钮
  const addBtn = document.getElementById('btn-add-dep');
  if (addBtn) {
    const newBtn = addBtn.cloneNode(true);
    addBtn.parentNode.replaceChild(newBtn, addBtn);
    newBtn.addEventListener('click', () => {
      if (stages.length < 2) { showToast('至少需要 2 个环节才能添加依赖', 'warning'); return; }
      const pair = findNextDependencyPair();
      if (!pair) {
        showToast('没有可自动添加的依赖关系，请编辑已有依赖或删除后调整', 'warning');
        return;
      }
      pushHistory();
      dependencies.push({ id: `dep_custom_${++dependencyIdCounter}`, fromId: pair.from.id, toId: pair.to.id, linkType: 'FS', gapDays: 0 });
      renderDependencies(); autoRunCalculation();
      showToast(`已添加依赖：${pair.from.name} → ${pair.to.name}`);
    });
  }
  // 修改前置环节
  document.querySelectorAll('#dependency-list [data-dep-from]').forEach((sel) => {
    sel.addEventListener('change', (e) => {
      const dep = dependencies.find((d) => d.id === e.target.dataset.depFrom); if (!dep) return;
      const prev = dep.fromId; dep.fromId = e.target.value;
      if (dep.fromId === 'none') {
        // 选择无依赖，清空连接方式和间隔
        pushHistory(); renderDependencies(); autoRunCalculation(); return;
      }
      if (dep.fromId === dep.toId || hasCycle()) { dep.fromId = prev; renderDependencies(); showToast('该依赖会形成循环或自依赖', 'error'); return; }
      pushHistory(); renderDependencies(); autoRunCalculation();
    });
  });
  // 修改后续环节
  document.querySelectorAll('#dependency-list [data-dep-to]').forEach((sel) => {
    sel.addEventListener('change', (e) => {
      const dep = dependencies.find((d) => d.id === e.target.dataset.depTo); if (!dep) return;
      const prev = dep.toId; dep.toId = e.target.value;
      if (dep.fromId === dep.toId || hasCycle()) { dep.toId = prev; renderDependencies(); showToast('该依赖会形成循环或自依赖', 'error'); return; }
      pushHistory(); renderDependencies(); autoRunCalculation();
    });
  });
  // 连接方式（高级模式才有）
  document.querySelectorAll('#dependency-list [data-dep-link]').forEach((sel) => {
    sel.addEventListener('change', (e) => {
      const dep = dependencies.find((d) => d.id === e.target.dataset.depLink); if (!dep) return;
      pushHistory(); dep.linkType = e.target.value; renderDependencies(); autoRunCalculation();
    });
  });
  // 间隔天数
  document.querySelectorAll('#dependency-list [data-dep-gap]').forEach((input) => {
    input.addEventListener('input', (e) => {
      const dep = dependencies.find((d) => d.id === e.target.dataset.depGap); if (!dep) return;
      dep.gapDays = Math.max(0, safeInt(e.target.value, 0)); autoRunCalculation();
    });
  });
  // 删除
  document.querySelectorAll('#dependency-list [data-dep-delete]').forEach((btn) => {
    btn.addEventListener('click', () => {
      pushHistory(); dependencies = dependencies.filter((d) => d.id !== btn.dataset.depDelete);
      renderDependencies(); autoRunCalculation(); showToast('依赖已删除');
    });
  });
  // 箭头点击切换连接方式
  document.querySelectorAll('#dependency-list [data-dep-arrow]').forEach((arrow) => {
    arrow.addEventListener('click', () => {
      const dep = dependencies.find((d) => d.id === arrow.dataset.depArrow); if (!dep) return;
      if (dep.fromId === 'none') return; // 无依赖不切换
      pushHistory();
      const curIdx = LINK_ORDER.indexOf(dep.linkType);
      dep.linkType = LINK_ORDER[(curIdx + 1) % LINK_ORDER.length];
      renderDependencies(); autoRunCalculation();
    });
  });
}

function bindDepDragSort() {
  const list = document.getElementById('dependency-list'); if (!list) return;
  let dragIdx = -1;
  let dragFromHandle = false;
  list.querySelectorAll('.dep-row-compact').forEach((item) => {
    item.addEventListener('mousedown', (e) => {
      dragFromHandle = !!e.target.closest('.dep-drag-handle');
    });
    item.addEventListener('dragstart', (e) => {
      if (!dragFromHandle) { e.preventDefault(); return; }
      dragIdx = safeInt(item.dataset.depIndex, -1);
      item.classList.add('dep-dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    item.addEventListener('dragend', () => {
      dragIdx = -1; dragFromHandle = false; item.classList.remove('dep-dragging');
      list.querySelectorAll('.dep-row-compact').forEach((n) => n.classList.remove('dep-drag-over'));
    });
    item.addEventListener('dragover', (e) => {
      e.preventDefault(); e.dataTransfer.dropEffect = 'move';
      list.querySelectorAll('.dep-row-compact').forEach((n) => n.classList.remove('dep-drag-over'));
      item.classList.add('dep-drag-over');
    });
    item.addEventListener('dragleave', () => item.classList.remove('dep-drag-over'));
    item.addEventListener('drop', (e) => {
      e.preventDefault(); item.classList.remove('dep-drag-over');
      const dropIdx = safeInt(item.dataset.depIndex, -1);
      if (dragIdx < 0 || dropIdx < 0 || dragIdx === dropIdx) return;
      pushHistory();
      const [moved] = dependencies.splice(dragIdx, 1);
      dependencies.splice(dropIdx, 0, moved);
      renderDependencies();
    });
  });
}

/* ============ 排期计算 ============ */
function getSuccessorStart(predResult, dep, currentDays, rule) {
  if (dep.linkType === 'SS') return shiftWorkdays(predResult.startDate, dep.gapDays, rule);
  if (dep.linkType === 'FF') return shiftWorkdays(shiftWorkdays(predResult.endDate, dep.gapDays, rule), -(currentDays - 1), rule);
  if (dep.linkType === 'SF') return shiftWorkdays(shiftWorkdays(predResult.startDate, dep.gapDays, rule), -(currentDays - 1), rule);
  return shiftWorkdays(predResult.endDate, dep.gapDays + 1, rule);
}
function getPredecessorEnd(succResult, dep, predDays, rule) {
  if (dep.linkType === 'SS') return shiftWorkdays(shiftWorkdays(succResult.startDate, -dep.gapDays, rule), predDays - 1, rule);
  if (dep.linkType === 'FF') return shiftWorkdays(succResult.endDate, -dep.gapDays, rule);
  if (dep.linkType === 'SF') return shiftWorkdays(shiftWorkdays(succResult.endDate, -dep.gapDays, rule), predDays - 1, rule);
  return shiftWorkdays(succResult.startDate, -(dep.gapDays + 1), rule);
}

function calculateDirection(direction, options = {}) {
  normalizeDependencies();
  const enabledStages = getEnabledStages(); if (!enabledStages.length) return null;
  const enabledDeps = getEnabledDependencies();
  const order = buildTopologicalOrder(enabledStages, enabledDeps);
  if (!order) { showToast('检测到循环依赖，请先调整依赖关系', 'error'); return null; }
  const rule = document.querySelector('input[name="workday"]:checked')?.value || 'double-rest';
  const extraDays = Math.max(0, safeInt(document.getElementById('extra-days').value, 0));
  const baseDateText = options.anchorDate || document.getElementById('base-date').value;
  if (!baseDateText) return null;
  const baseDate = parseDate(baseDateText);
  const resultMap = {};

  if (direction === 'forward') {
    const actualBase = ensureWorkday(baseDate, rule, 1);
    order.forEach((sid) => {
      const stage = getStageById(sid);
      const incoming = enabledDeps.filter((d) => d.toId === sid);
      let startDate = new Date(actualBase);
      incoming.forEach((d) => { const pr = resultMap[d.fromId]; if (!pr) return; const c = getSuccessorStart(pr, d, stage.days, rule); if (c > startDate) startDate = c; });
      startDate = ensureWorkday(startDate, rule, 1);
      const endDate = shiftWorkdays(startDate, stage.days - 1, rule);
      resultMap[sid] = { id: stage.id, name: stage.name, days: stage.days, startDate, endDate, incoming };
    });
    const results = enabledStages.map((s) => resultMap[s.id]);
    const projectStart = results.reduce((min, r) => r.startDate < min ? r.startDate : min, results[0].startDate);
    const stageEnd = results.reduce((max, r) => r.endDate > max ? r.endDate : max, results[0].endDate);
    const projectEnd = extraDays > 0 ? shiftWorkdays(stageEnd, extraDays, rule) : stageEnd;
    return { direction, rule, baseDate: actualBase, projectStart, stageEnd, projectEnd, extraDays, results, dependencies: enabledDeps };
  }

  const deadline = ensureWorkday(baseDate, rule, -1);
  const stageDeadline = extraDays > 0 ? shiftWorkdays(deadline, -extraDays, rule) : deadline;
  [...order].reverse().forEach((sid) => {
    const stage = getStageById(sid);
    const outgoing = enabledDeps.filter((d) => d.fromId === sid);
    let endDate = new Date(stageDeadline);
    outgoing.forEach((d) => { const sr = resultMap[d.toId]; if (!sr) return; const c = getPredecessorEnd(sr, d, stage.days, rule); if (c < endDate) endDate = c; });
    endDate = ensureWorkday(endDate, rule, -1);
    const startDate = shiftWorkdays(endDate, -(stage.days - 1), rule);
    resultMap[sid] = { id: stage.id, name: stage.name, days: stage.days, startDate, endDate, incoming: enabledDeps.filter((d) => d.toId === sid) };
  });
  const results = enabledStages.map((s) => resultMap[s.id]);
  const projectStart = results.reduce((min, r) => r.startDate < min ? r.startDate : min, results[0].startDate);
  const stageEnd = results.reduce((max, r) => r.endDate > max ? r.endDate : max, results[0].endDate);
  return { direction, rule, baseDate: deadline, projectStart, stageEnd, projectEnd: deadline, extraDays, results, dependencies: enabledDeps };
}

function formatDependencyText(dep, data) {
  if (dep.fromId === 'none') return '<span class="badge badge-serial">无依赖</span>';
  const fromName = data.results.find((r) => r.id === dep.fromId)?.name || getStageById(dep.fromId)?.name || '未知';
  const meta = LINK_META[dep.linkType] || LINK_META.FS;
  return `${esc(fromName)} <span class="badge badge-serial">${meta.label}</span>${dep.gapDays ? ` +${dep.gapDays}天` : ''}`;
}

function renderSummary(data, target) {
  const el = target || document.getElementById('summary-content'); if (!el || !data) return;
  const week = ['日','一','二','三','四','五','六'];
  const ruleText = data.rule === 'single-rest' ? '单休 + 法定节假日' : '双休 + 法定节假日';

  // 期望结束日期对比分析
  const expectedEndStr = document.getElementById('expected-end-date')?.value;
  let deadlineHtml = '';
  if (expectedEndStr && !target) {
    const expectedEnd = parseDate(expectedEndStr);
    const actualEnd = data.projectEnd;
    const diffMs = actualEnd.getTime() - expectedEnd.getTime();
    const diffDays = Math.round(diffMs / 86400000);
    if (diffDays > 0) {
      deadlineHtml = `<div class="summary-item summary-item-danger"><div class="val">超期 ${diffDays} 天</div><div class="lbl">实际 ${fmt(actualEnd)} 超出期望 ${fmt(expectedEnd)}</div></div>`;
    } else if (diffDays === 0) {
      deadlineHtml = `<div class="summary-item summary-item-success"><div class="val">刚好达标</div><div class="lbl">实际完成日 = 期望结束日 ${fmt(expectedEnd)}</div></div>`;
    } else {
      deadlineHtml = `<div class="summary-item summary-item-success"><div class="val">提前 ${Math.abs(diffDays)} 天</div><div class="lbl">实际 ${fmt(actualEnd)} 早于期望 ${fmt(expectedEnd)}</div></div>`;
    }
  }

  el.innerHTML = `<div class="summary-grid">
    <div class="summary-item"><div class="val">${fmt(data.projectStart)}</div><div class="lbl">${data.direction === 'forward' ? '最早启动' : '倒推出的启动'}（周${week[data.projectStart.getDay()]}）</div></div>
    <div class="summary-item"><div class="val">${fmt(data.projectEnd)}</div><div class="lbl">${data.direction === 'forward' ? '预计完成' : '上线日期'}（周${week[data.projectEnd.getDay()]}）</div></div>
    <div class="summary-item"><div class="val">${data.results.length} 个</div><div class="lbl">启用环节</div></div>
    <div class="summary-item"><div class="val">${data.extraDays} 天</div><div class="lbl">额外预留</div></div>
    <div class="summary-item"><div class="val">${getDirectionLabel(data.direction)}</div><div class="lbl">计算方向</div></div>
    <div class="summary-item summary-item-rule"><div class="val">${ruleText}</div><div class="lbl">工作日规则</div></div>
    ${deadlineHtml}
  </div>`;
}

function renderTable(data, target, compareData) {
  const el = target || document.getElementById('table-content'); if (!el || !data) return;
  const week = ['日','一','二','三','四','五','六'];

  // 构建对比映射：如果有compareData，用来检测被压缩的环节
  const compareMap = {};
  if (compareData && compareData.results) {
    compareData.results.forEach((r) => { compareMap[r.id] = r; });
  }

  // 汇总统计
  const totalDays = data.results.reduce((sum, r) => sum + r.days, 0);
  const earliestStart = data.results.reduce((min, r) => r.startDate < min ? r.startDate : min, data.results[0].startDate);
  const latestEnd = data.results.reduce((max, r) => r.endDate > max ? r.endDate : max, data.results[0].endDate);
  const totalWorkdays = countWorkdaysInclusive(earliestStart, latestEnd, data.rule);

  // 期望结束日期对比
  const expectedEndStr = document.getElementById('expected-end-date')?.value;
  let deadlineNote = '';
  if (expectedEndStr && !target) {
    const expectedEnd = parseDate(expectedEndStr);
    const diffMs = latestEnd.getTime() - expectedEnd.getTime();
    const diffDays = Math.round(diffMs / 86400000);
    if (diffDays > 0) {
      deadlineNote = `<span class="table-deadline-badge table-deadline-over">⚠ 超期 ${diffDays} 天</span>`;
    } else if (diffDays === 0) {
      deadlineNote = `<span class="table-deadline-badge table-deadline-ok">✓ 刚好达标</span>`;
    } else {
      deadlineNote = `<span class="table-deadline-badge table-deadline-ok">✓ 提前 ${Math.abs(diffDays)} 天</span>`;
    }
  }

  el.innerHTML = `<table class="result-table"><thead><tr><th>环节</th><th>工期</th><th>开始</th><th>结束</th><th>前置依赖</th></tr></thead><tbody>
    <tr class="table-summary-row">
      <td><strong>📊 汇总</strong></td>
      <td><strong>${totalDays} 天</strong><span class="summary-sub">（跨度 ${totalWorkdays} 工作日）</span></td>
      <td><strong>${fmt(earliestStart)}</strong>（周${week[earliestStart.getDay()]}）</td>
      <td><strong>${fmt(latestEnd)}</strong>（周${week[latestEnd.getDay()]}）${deadlineNote}</td>
      <td>${data.results.length} 个环节</td>
    </tr>
    ${data.results.map((item) => {
      const deps = data.dependencies.filter((d) => d.toId === item.id);
      // 检测该环节是否被压缩
      let compressHtml = '';
      const origRef = compareMap[item.id];
      if (origRef) {
        // 对比两种方向的实际工作日跨度
        const thisSpan = countWorkdaysInclusive(item.startDate, item.endDate, data.rule);
        const refSpan = countWorkdaysInclusive(origRef.startDate, origRef.endDate, data.rule);
        if (refSpan > thisSpan) {
          const compressed = refSpan - thisSpan;
          compressHtml = `<span class="compress-badge">压缩 ${compressed} 天</span>`;
        }
      }
      // 也对比原始配置工期：如果当前结果的工作日跨度比配置的工期少
      if (!compressHtml) {
        const actualSpan = countWorkdaysInclusive(item.startDate, item.endDate, data.rule);
        if (item.days > actualSpan) {
          const compressed = item.days - actualSpan;
          compressHtml = `<span class="compress-badge">压缩 ${compressed} 天</span>`;
        }
      }
      return `<tr><td><strong>${esc(item.name)}</strong></td><td>${item.days} 天${compressHtml}</td><td>${fmt(item.startDate)}（周${week[item.startDate.getDay()]}）</td><td>${fmt(item.endDate)}（周${week[item.endDate.getDay()]}）</td><td>${deps.length ? deps.map((d) => formatDependencyText(d, data)).join('<br>') : '—'}</td></tr>`;
    }).join('')}
    ${data.extraDays > 0 ? (() => {
      // 额外预留行：从最后环节结束日的下一个工作日到projectEnd
      const bufferStart = shiftWorkdays(data.stageEnd, 1, data.rule);
      const bufferEnd = data.projectEnd;
      return `<tr class="table-buffer-row"><td><strong>⏳ 额外预留</strong></td><td>${data.extraDays} 天</td><td>${fmt(bufferStart)}（周${week[bufferStart.getDay()]}）</td><td>${fmt(bufferEnd)}（周${week[bufferEnd.getDay()]}）</td><td>—</td></tr>`;
    })() : ''}
  </tbody></table>`;
}

function applyGanttZoom() {
  // 顺排甘特图缩放
  const fwdWrapper = document.getElementById('gantt-content');
  if (fwdWrapper) {
    fwdWrapper.querySelectorAll('[data-gantt-outer]').forEach((outer) => {
      outer.style.transform = `scale(${ganttZoom})`; outer.style.transformOrigin = 'top left';
    });
  }
  const fwdLabel = document.getElementById('fwd-zoom-level'); if (fwdLabel) fwdLabel.textContent = `${Math.round(ganttZoom * 100)}%`;
  // 倒排甘特图缩放
  const bwdWrapper = document.getElementById('backward-gantt-content');
  if (bwdWrapper) {
    bwdWrapper.querySelectorAll('[data-gantt-outer]').forEach((outer) => {
      outer.style.transform = `scale(${ganttZoomBwd})`; outer.style.transformOrigin = 'top left';
    });
  }
  const bwdLabel = document.getElementById('bwd-zoom-level'); if (bwdLabel) bwdLabel.textContent = `${Math.round(ganttZoomBwd * 100)}%`;
}

function getWeekNumber(date, projectStart) {
  // 开始周永远是第1周，不管从周几开始
  const startMonday = new Date(projectStart);
  // 找到开始日期所在周的周一
  const dayOfWeek = startMonday.getDay();
  const diff = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  startMonday.setDate(startMonday.getDate() - diff);
  startMonday.setHours(0, 0, 0, 0);
  const target = new Date(date);
  target.setHours(0, 0, 0, 0);
  const diffDays = Math.floor((target - startMonday) / 86400000);
  return Math.floor(diffDays / 7) + 1;
}

function renderGantt(data, targetWrapper) {
  const wrapper = targetWrapper || document.getElementById('gantt-content');
  const tip = document.getElementById('gantt-drag-tip');
  if (!wrapper || !data) return; if (tip) tip.style.display = 'block';

  // 甘特图日期范围：开始前3天 ~ 结束后3个月
  const start = addDays(data.projectStart, -3);
  const endPlus3m = new Date(data.projectEnd);
  endPlus3m.setMonth(endPlus3m.getMonth() + 3);
  const end = endPlus3m;

  const days = []; for (let c = new Date(start); c <= end; c = addDays(c, 1)) days.push(new Date(c));
  const cellW = 32; const labelW = 132; const today = fmt(new Date());

  // 构建月份行
  let monthHtml = '<tr class="month-row"><th class="gantt-frozen-col"></th>';
  let curMonth = ''; let monthSpan = 0; let monthCells = [];
  days.forEach((d, i) => {
    const mKey = `${d.getFullYear()}年${d.getMonth() + 1}月`;
    if (mKey !== curMonth) {
      if (curMonth) monthCells.push({ label: curMonth, span: monthSpan });
      curMonth = mKey; monthSpan = 1;
    } else { monthSpan++; }
    if (i === days.length - 1) monthCells.push({ label: curMonth, span: monthSpan });
  });
  monthHtml += monthCells.map((m) => `<th colspan="${m.span}">${m.label}</th>`).join('');
  monthHtml += '</tr>';

  // 构建周号行
  let weekHtml = '<tr class="week-row"><th class="gantt-frozen-col"></th>';
  let curWeek = -999; let weekSpan = 0; let weekCells = [];
  days.forEach((d, i) => {
    const wn = getWeekNumber(d, data.projectStart);
    if (wn !== curWeek) {
      if (curWeek !== -999) weekCells.push({ label: `W${curWeek}`, span: weekSpan });
      curWeek = wn; weekSpan = 1;
    } else { weekSpan++; }
    if (i === days.length - 1) weekCells.push({ label: `W${curWeek}`, span: weekSpan });
  });
  weekHtml += weekCells.map((w) => `<th colspan="${w.span}">${w.label}</th>`).join('');
  weekHtml += '</tr>';

  let html = `<table class="gantt-table"><colgroup><col style="width:130px">`;
  html += days.map(() => `<col style="width:32px">`).join('');
  html += `</colgroup><thead>${monthHtml}${weekHtml}<tr class="date-row"><th class="gantt-frozen-col">环节</th>`;
  html += days.map((d) => { let cls = ''; if (isWeekend(d)) cls += ' col-weekend'; if (isHoliday(d)) cls += ' col-holiday'; if (fmt(d) === today) cls += ' col-today'; return `<th class="${cls.trim()}">${d.getDate()}</th>`; }).join('');
  html += '</tr></thead><tbody>';
  html += data.results.map((item) => `<tr data-stage-row="${item.id}"><td class="gantt-frozen-col">${esc(item.name)}</td>${days.map((d) => { let cls = ''; if (isWeekend(d)) cls += ' col-weekend'; if (isHoliday(d)) cls += ' col-holiday'; if (fmt(d) === today) cls += ' col-today'; return `<td class="${cls.trim()}" data-date="${fmt(d)}"></td>`; }).join('')}</tr>`).join('');
  // 额外预留行
  if (data.extraDays > 0) {
    html += `<tr data-stage-row="__buffer__"><td class="gantt-frozen-col gantt-buffer-label">⏳ 额外预留</td>${days.map((d) => { let cls = ''; if (isWeekend(d)) cls += ' col-weekend'; if (isHoliday(d)) cls += ' col-holiday'; if (fmt(d) === today) cls += ' col-today'; return `<td class="${cls.trim()}" data-date="${fmt(d)}"></td>`; }).join('')}</tr>`;
  }
  html += '</tbody></table>';
  const outer = document.createElement('div'); outer.className = 'gantt-outer'; outer.dataset.ganttOuter = 'true'; outer.innerHTML = html;
  wrapper.innerHTML = ''; wrapper.appendChild(outer);

  const rule = data.rule;
  const rows = outer.querySelectorAll('tbody tr');

  // 关键：从实际 DOM 获取单元格位置，避免硬编码导致错位
  function getCellPositions(rowEl) {
    const cells = rowEl.querySelectorAll('td');
    const positions = [];
    for (let c = 1; c < cells.length; c++) { // 跳过frozen col (index 0)
      positions.push({
        left: cells[c].offsetLeft,
        width: cells[c].offsetWidth,
        right: cells[c].offsetLeft + cells[c].offsetWidth
      });
    }
    return positions;
  }

  // 获取表头日期行的单元格位置（作为所有行的基准）
  const dateHeaderRow = outer.querySelector('.date-row');
  const headerCells = dateHeaderRow ? dateHeaderRow.querySelectorAll('th') : [];
  const headerPositions = [];
  for (let c = 1; c < headerCells.length; c++) {
    headerPositions.push({
      left: headerCells[c].offsetLeft,
      width: headerCells[c].offsetWidth
    });
  }

  // 实际的单元格宽度和标签列宽度
  const realCellW = headerPositions.length > 1 ? headerPositions[1].left - headerPositions[0].left : cellW;
  const realLabelW = headerPositions.length > 0 ? headerPositions[0].left : labelW;

  data.results.forEach((item, i) => {
    const row = rows[i]; if (!row) return;
    const si = Math.round((item.startDate - start) / 86400000);
    const ei = Math.round((item.endDate - start) / 86400000);

    // 使用实际DOM位置定位色条
    const startPos = headerPositions[si] || headerPositions[0];
    const endPos = headerPositions[ei] || headerPositions[headerPositions.length - 1];
    const barLeft = startPos.left;
    const barWidth = endPos.left + endPos.width - startPos.left;

    const bar = document.createElement('div');
    bar.className = `gantt-continuous-bar gantt-color-${i % 8}`;
    bar.dataset.stageId = item.id;
    bar.style.top = `${row.offsetTop + 5}px`;
    bar.style.left = `${barLeft}px`;
    bar.style.width = `${barWidth}px`;
    bar.style.height = `${row.offsetHeight - 10}px`;
    bar.innerHTML = `<div class="gantt-bar-handle gantt-bar-handle-left" data-handle="left"></div><span class="gantt-bar-text">${esc(item.name)} (${item.days}天)</span><div class="gantt-bar-handle gantt-bar-handle-right" data-handle="right"></div>`;
    outer.appendChild(bar);

    // 拖拽调整开始/结束时间
    bindBarDrag(bar, item, i, start, days, realCellW, realLabelW, row, outer, data, rule, headerPositions);
  });

  // 拖选功能
  bindGanttDragSelect(outer, days, data, realCellW, realLabelW, rows);

  // 额外预留色条
  if (data.extraDays > 0) {
    const bufferStart = shiftWorkdays(data.stageEnd, 1, data.rule);
    const bufferEnd = data.projectEnd;
    const bsi = Math.round((bufferStart - start) / 86400000);
    const bei = Math.round((bufferEnd - start) / 86400000);
    const bufferRow = outer.querySelector('tr[data-stage-row="__buffer__"]');
    if (bufferRow && bsi >= 0 && bei >= 0 && bsi < headerPositions.length && bei < headerPositions.length) {
      const bStartPos = headerPositions[bsi] || headerPositions[0];
      const bEndPos = headerPositions[bei] || headerPositions[headerPositions.length - 1];
      const bBar = document.createElement('div');
      bBar.className = 'gantt-continuous-bar gantt-buffer-bar';
      bBar.style.top = `${bufferRow.offsetTop + 5}px`;
      bBar.style.left = `${bStartPos.left}px`;
      bBar.style.width = `${bEndPos.left + bEndPos.width - bStartPos.left}px`;
      bBar.style.height = `${bufferRow.offsetHeight - 10}px`;
      bBar.innerHTML = `<span class="gantt-bar-text">额外预留 (${data.extraDays}天)</span>`;
      outer.appendChild(bBar);
    }
  }

  // 期望结束日期标线
  const expectedEndStr = document.getElementById('expected-end-date')?.value;
  if (expectedEndStr) {
    const expectedEnd = parseDate(expectedEndStr);
    const expectedIdx = Math.round((expectedEnd - start) / 86400000);
    if (expectedIdx >= 0 && expectedIdx < headerPositions.length) {
      const pos = headerPositions[expectedIdx];
      const lineX = pos.left + pos.width / 2;
      const line = document.createElement('div');
      line.className = 'gantt-expected-end-line';
      line.style.left = `${lineX}px`;
      line.style.top = '0px';
      line.style.height = `${outer.scrollHeight}px`;
      outer.appendChild(line);
      const label = document.createElement('div');
      label.className = 'gantt-expected-end-label';
      label.textContent = '🎯 期望结束';
      label.style.left = `${lineX + 2}px`;
      label.style.top = '0px';
      outer.appendChild(label);

      // 在标线底部也添加一个icon标记
      const bottomIcon = document.createElement('div');
      bottomIcon.className = 'gantt-expected-end-icon';
      bottomIcon.textContent = '🎯';
      bottomIcon.style.left = `${lineX - 8}px`;
      bottomIcon.style.top = `${outer.scrollHeight - 22}px`;
      outer.appendChild(bottomIcon);
    }
  }

  applyGanttZoom();
}

function bindBarDrag(bar, item, rowIdx, ganttStart, days, cellW, labelW, row, outer, data, rule, headerPositions) {
  let dragType = null; // 'left' | 'right' | 'move'
  let startX = 0;
  let origStartIdx = 0; // 色块起始日期索引
  let origEndIdx = 0;   // 色块结束日期索引
  let tooltip = null;
  let hasMoved = false; // 是否实际发生了拖动
  let lastSnapStartIdx = -1; // 上一次吸附的起始索引，避免重复渲染
  let lastSnapEndIdx = -1;

  // 通过bar当前位置计算初始日期索引
  function getBarDateIndices() {
    const barLeft = parseFloat(bar.style.left);
    const barWidth = parseFloat(bar.style.width);
    const si = dayIdxFromPixel(barLeft + 2);
    const ei = dayIdxFromPixel(barLeft + barWidth - 2);
    return { startIdx: si, endIdx: ei };
  }

  // 将bar精确吸附到指定的日期索引位置
  function snapBarToIndices(si, ei) {
    si = Math.max(0, Math.min(si, headerPositions.length - 1));
    ei = Math.max(si, Math.min(ei, headerPositions.length - 1));
    const newLeft = headerPositions[si].left;
    const newRight = headerPositions[ei].left + headerPositions[ei].width;
    bar.style.left = `${newLeft}px`;
    bar.style.width = `${newRight - newLeft}px`;
    return { startIdx: si, endIdx: ei };
  }

  function createTooltip() {
    tooltip = document.createElement('div');
    tooltip.className = 'gantt-drag-tooltip';
    outer.appendChild(tooltip);
  }
  function updateTooltip(startIdx, endIdx) {
    if (!tooltip) return;
    const startDate = days[startIdx] || days[0];
    const endDate = days[endIdx] || days[days.length - 1];
    const workDays = countWorkdaysInclusive(startDate, endDate, rule);
    tooltip.textContent = `${fmt(startDate)} → ${fmt(endDate)}（${workDays}天）`;
    const barLeft = parseFloat(bar.style.left);
    const barWidth = parseFloat(bar.style.width);
    tooltip.style.left = `${barLeft + barWidth / 2}px`;
    tooltip.style.top = `${parseFloat(bar.style.top) - 30}px`;
  }
  function removeTooltip() {
    if (tooltip) { tooltip.remove(); tooltip = null; }
  }

  function dayIdxFromPixel(px) {
    let bestIdx = 0;
    let bestDist = Infinity;
    for (let i = 0; i < headerPositions.length; i++) {
      const center = headerPositions[i].left + headerPositions[i].width / 2;
      const dist = Math.abs(px - center);
      if (dist < bestDist) { bestDist = dist; bestIdx = i; }
    }
    return bestIdx;
  }

  // 通过像素偏移量计算日期列数偏移
  function pixelToColumnOffset(dx) {
    const avgCellW = headerPositions.length > 1
      ? (headerPositions[headerPositions.length - 1].left + headerPositions[headerPositions.length - 1].width - headerPositions[0].left) / headerPositions.length
      : cellW;
    return Math.round(dx / avgCellW);
  }

  function onMouseDown(e) {
    const handle = e.target.closest('[data-handle]');
    if (handle) {
      dragType = handle.dataset.handle;
    } else {
      dragType = 'move';
    }
    startX = e.clientX / ganttZoom;
    const indices = getBarDateIndices();
    origStartIdx = indices.startIdx;
    origEndIdx = indices.endIdx;
    lastSnapStartIdx = origStartIdx;
    lastSnapEndIdx = origEndIdx;
    hasMoved = false;
    bar.classList.add('dragging');
    createTooltip();
    updateTooltip(origStartIdx, origEndIdx);
    e.preventDefault();
    e.stopPropagation();
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  function onMouseMove(e) {
    const dx = e.clientX / ganttZoom - startX;
    // 至少移动3像素才算真正拖动
    if (!hasMoved && Math.abs(dx) < 3) return;
    hasMoved = true;

    const colOffset = pixelToColumnOffset(dx);
    let newStartIdx, newEndIdx;

    if (dragType === 'left') {
      newStartIdx = Math.max(0, Math.min(origStartIdx + colOffset, origEndIdx));
      newEndIdx = origEndIdx;
    } else if (dragType === 'right') {
      newStartIdx = origStartIdx;
      newEndIdx = Math.max(origStartIdx, Math.min(origEndIdx + colOffset, headerPositions.length - 1));
    } else { // move
      const span = origEndIdx - origStartIdx;
      newStartIdx = Math.max(0, origStartIdx + colOffset);
      newEndIdx = newStartIdx + span;
      if (newEndIdx >= headerPositions.length) {
        newEndIdx = headerPositions.length - 1;
        newStartIdx = newEndIdx - span;
      }
    }

    // 只在吸附位置变化时更新DOM，避免抖动
    if (newStartIdx !== lastSnapStartIdx || newEndIdx !== lastSnapEndIdx) {
      lastSnapStartIdx = newStartIdx;
      lastSnapEndIdx = newEndIdx;
      snapBarToIndices(newStartIdx, newEndIdx);
      updateTooltip(newStartIdx, newEndIdx);
    }
  }

  function onMouseUp(e) {
    document.removeEventListener('mousemove', onMouseMove);
    document.removeEventListener('mouseup', onMouseUp);
    bar.classList.remove('dragging');
    removeTooltip();

    // 如果没有实际拖动（只是单击），恢复原位并跳过更新
    if (!hasMoved) {
      snapBarToIndices(origStartIdx, origEndIdx);
      return;
    }

    const newStartDate = days[lastSnapStartIdx] || days[0];
    const newEndDate = days[lastSnapEndIdx] || days[days.length - 1];

    // 计算新的工期（工作日数）
    const newDays = countWorkdaysInclusive(newStartDate, newEndDate, rule);
    if (newDays < 1) {
      snapBarToIndices(origStartIdx, origEndIdx);
      return;
    }

    // 回写到 stages
    const stage = getStageById(item.id);
    if (stage) {
      pushHistory();
      stage.days = newDays;
      // 重新渲染并计算
      renderStages();
      runCalculation();
      showToast(`${stage.name} 工期调整为 ${newDays} 天`);
    }
  }

  bar.addEventListener('mousedown', onMouseDown);
}

function bindGanttDragSelect(outer, days, data, cellW, labelW, rows) {
  let isSelecting = false;
  let selectStartCell = null;
  let selectOverlay = null;

  outer.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td[data-date]');
    if (!td) return;
    // 不在色条上时才开始拖选
    if (e.target.closest('.gantt-continuous-bar')) return;
    isSelecting = true;
    selectStartCell = td;
    // 创建选区覆盖层
    if (selectOverlay) selectOverlay.remove();
    selectOverlay = document.createElement('div');
    selectOverlay.className = 'gantt-select-overlay';
    outer.appendChild(selectOverlay);
    updateSelectOverlay(td, td);
    e.preventDefault();
  });

  outer.addEventListener('mousemove', (e) => {
    if (!isSelecting) return;
    const td = e.target.closest('td[data-date]');
    if (!td) return;
    updateSelectOverlay(selectStartCell, td);
  });

  outer.addEventListener('mouseup', (e) => {
    if (!isSelecting) return;
    isSelecting = false;
    const td = e.target.closest('td[data-date]');
    if (!td || !selectStartCell) { if (selectOverlay) { selectOverlay.remove(); selectOverlay = null; } return; }

    const startDate = parseDate(selectStartCell.dataset.date);
    const endDate = parseDate(td.dataset.date);
    const realStart = startDate < endDate ? startDate : endDate;
    const realEnd = startDate < endDate ? endDate : startDate;

    // 找到拖选所在行对应的环节
    const startRow = selectStartCell.closest('tr');
    const stageId = startRow?.dataset.stageRow;

    if (selectOverlay) { selectOverlay.remove(); selectOverlay = null; }

    if (stageId) {
      const rule = data.rule;
      const newDays = countWorkdaysInclusive(realStart, realEnd, rule);
      if (newDays < 1) return;
      const stage = getStageById(stageId);
      if (stage) {
        pushHistory();
        stage.days = newDays;
        renderStages();
        runCalculation();
        showToast(`${stage.name} 工期调整为 ${newDays} 天（拖选 ${fmt(realStart)} ~ ${fmt(realEnd)}）`);
      }
    }
  });

  function updateSelectOverlay(cell1, cell2) {
    if (!selectOverlay || !cell1 || !cell2) return;
    const r1 = cell1.getBoundingClientRect();
    const r2 = cell2.getBoundingClientRect();
    const outerRect = outer.getBoundingClientRect();
    const left = Math.min(r1.left, r2.left) - outerRect.left + outer.scrollLeft;
    const top = Math.min(r1.top, r2.top) - outerRect.top + outer.scrollTop;
    const right = Math.max(r1.right, r2.right) - outerRect.left + outer.scrollLeft;
    const bottom = Math.max(r1.bottom, r2.bottom) - outerRect.top + outer.scrollTop;
    selectOverlay.style.left = `${left}px`;
    selectOverlay.style.top = `${top}px`;
    selectOverlay.style.width = `${right - left}px`;
    selectOverlay.style.height = `${bottom - top}px`;
  }
}

function renderCompareResult(primaryDir) {
  const checked = document.getElementById('chk-backward')?.checked;
  // 是否显示对比结果区块
  const showCompare = !!checked;

  const backwardSummaryBlock = document.getElementById('backward-summary-block');
  const backwardTableBlock = document.getElementById('backward-table-block');
  const backwardGanttSection = document.getElementById('sec-gantt-backward');
  const forwardSummaryBlock = document.getElementById('forward-summary-block');
  const forwardTableBlock = document.getElementById('forward-table-block');

  if (backwardSummaryBlock) backwardSummaryBlock.style.display = showCompare ? 'block' : 'none';
  if (backwardTableBlock) backwardTableBlock.style.display = showCompare ? 'block' : 'none';
  if (backwardGanttSection) backwardGanttSection.style.display = showCompare ? 'block' : 'none';

  // 单方向时去除主方向区块的边框等装饰
  [forwardSummaryBlock, forwardTableBlock].forEach((el) => {
    if (el) el.classList.toggle('dual-result-solo', !showCompare);
  });

  if (!showCompare) return;

  // 对比方向 = 主方向的反向
  const oppositeDir = getOppositeDirection(primaryDir);

  // 用主方向的计算结果推导对比方向的锚定日期
  // 正推为主 → 倒排对比：用正推的"预计完成日"作为倒排截止日
  // 倒推为主 → 正推对比：用倒推的"倒推启动日"作为正推起始日
  let anchorDate = null;
  if (lastPrimaryData) {
    if (primaryDir === 'forward') {
      anchorDate = fmt(lastPrimaryData.projectEnd);
    } else {
      anchorDate = fmt(lastPrimaryData.projectStart);
    }
  }
  const compareData = calculateDirection(oppositeDir, anchorDate ? { anchorDate } : {});
  lastCompareData = compareData;
  if (!compareData) {
    document.getElementById('backward-summary-content').innerHTML = '<p class="placeholder-text">暂无可展示结果</p>';
    document.getElementById('backward-table-content').innerHTML = '';
    document.getElementById('backward-gantt-content').innerHTML = '';
    return;
  }
  renderSummary(compareData, document.getElementById('backward-summary-content'));
  renderTable(compareData, document.getElementById('backward-table-content'), lastPrimaryData);
  renderGantt(compareData, document.getElementById('backward-gantt-content'));
}

function runCalculation() {
  const baseDate = document.getElementById('base-date').value;
  if (!baseDate) { showToast('请填写需求开始日期', 'warning'); highlightEmptyField('base-date'); return null; }
  if (!getEnabledStages().length) { showToast('请至少启用一个环节', 'warning'); return null; }
  // 检查所有启用环节的工期是否已填
  const emptyDaysStages = getEnabledStages().filter((s) => !s.days || s.days <= 0);
  if (emptyDaysStages.length) {
    showToast(`请填写以下环节的工期：${emptyDaysStages.map((s) => s.name).join('、')}`, 'error');
    // 高亮所有未填工期的输入框
    emptyDaysStages.forEach((s) => {
      const box = document.querySelector(`[data-stage-days="${s.id}"]`)?.closest('.stage-days-box');
      if (box) { box.classList.add('stage-days-empty', 'stage-days-shake'); setTimeout(() => box.classList.remove('stage-days-shake'), 600); }
    });
    return null;
  }
  const dir = getCurrentDirection(); const result = calculateDirection(dir); if (!result) return null;
  lastPrimaryData = result; renderSummary(result); renderTable(result); renderGantt(result); renderCompareResult(dir); return result;
}
function autoRunCalculation() {
  clearTimeout(autoCalcTimer);
  autoCalcTimer = setTimeout(() => { if (!document.getElementById('base-date').value || !getEnabledStages().length) return; runCalculation(); }, 160);
  saveWorkspace();
}

function clearResultViews() {
  lastPrimaryData = null;
  lastCompareData = null;
  currentViewingRecord = null;
  ganttZoom = 1;
  ganttZoomBwd = 1;
  const summary = document.getElementById('summary-content');
  const table = document.getElementById('table-content');
  const gantt = document.getElementById('gantt-content');
  if (summary) summary.innerHTML = '<p class="placeholder-text">👈 请在左侧配置后点击「计算排期」</p>';
  if (table) table.innerHTML = '<p class="placeholder-text">等待计算...</p>';
  if (gantt) gantt.innerHTML = '<p class="placeholder-text">等待计算...</p>';
  const backwardSummary = document.getElementById('backward-summary-content');
  const backwardTable = document.getElementById('backward-table-content');
  const backwardGantt = document.getElementById('backward-gantt-content');
  if (backwardSummary) backwardSummary.innerHTML = '';
  if (backwardTable) backwardTable.innerHTML = '';
  if (backwardGantt) backwardGantt.innerHTML = '';
  ['backward-summary-block', 'backward-table-block', 'sec-gantt-backward'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  ['forward-summary-block', 'forward-table-block'].forEach((id) => {
    document.getElementById(id)?.classList.add('dual-result-solo');
  });
  const chk = document.getElementById('chk-backward');
  if (chk) chk.checked = false;
  const fwdZoom = document.getElementById('fwd-zoom-level');
  const bwdZoom = document.getElementById('bwd-zoom-level');
  if (fwdZoom) fwdZoom.textContent = '100%';
  if (bwdZoom) bwdZoom.textContent = '100%';
  renderRecordViewingBanner();
  renderSavedRecordsList();
}

function clearCurrentConfiguration() {
  const ok = window.confirm('确定要清空当前输入配置并新建排期吗？已保存记录、标效模板和环节名称选项不会删除。');
  if (!ok) return;
  pushHistory();
  stages = getDefaultStagesWithAliases();
  dependencies = clone(DEFAULT_DEPENDENCIES);
  stageIdCounter = 100;
  dependencyIdCounter = 100;
  document.querySelector('input[name="direction"][value="forward"]').checked = true;
  document.querySelector('input[name="workday"][value="double-rest"]').checked = true;
  document.getElementById('base-date').value = '';
  document.getElementById('extra-days').value = '0';
  const expectedEndInput = document.getElementById('expected-end-date');
  if (expectedEndInput) expectedEndInput.value = '';
  syncDirectionTexts();
  bindCompareToggle();
  renderAll();
  clearResultViews();
  hideTplHint();
  clearWorkspace();
  showToast('已清空当前配置');
}

/* ============ 排期结果保存/加载 ============ */
const MAX_SAVED_RECORDS = 5;
function getSavedRecords() { return getJSON(STORAGE_KEYS.savedRecords, []); }
function setSavedRecords(records) { return setJSON(STORAGE_KEYS.savedRecords, records); }

function fmtDateTime(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
}

function saveCurrentRecord(name) {
  const n = String(name || '').trim();
  if (!n) { showToast('请输入记录备注名称', 'warning'); return; }
  const records = getSavedRecords();
  if (records.length >= MAX_SAVED_RECORDS) {
    showToast(`最多保存 ${MAX_SAVED_RECORDS} 条记录，请先清理后再保存`, 'warning');
    return;
  }
  // 检查是否有同名记录
  if (records.some((r) => r.name === n)) {
    showToast(`已存在名为「${n}」的记录，请换个名称`, 'warning');
    return;
  }
  // 收集当前数据
  const payload = {
    name: n,
    savedAt: Date.now(),
    data: {
      stages: clone(stages),
      dependencies: clone(dependencies),
      stageNameOptions: clone(stageNameOptions),
      stageNameAliases: clone(stageNameAliases),
      stageIdCounter,
      dependencyIdCounter,
      settings: collectSettings()
    }
  };
  records.unshift(payload);
  if (!setSavedRecords(records)) return;
  modal('modal-save-record', false);
  renderSavedRecordsList();
  showToast(`排期记录「${n}」已保存`);
}

function loadSavedRecord(idx) {
  const records = getSavedRecords();
  const record = records[idx];
  if (!record) { showToast('记录不存在', 'error'); return; }
  pushHistory();
  stages = clone(record.data.stages || DEFAULT_STAGES);
  dependencies = clone(record.data.dependencies || DEFAULT_DEPENDENCIES);
  stageNameAliases = clone(record.data.stageNameAliases || stageNameAliases || {});
  saveStageNameAliases();
  stageNameOptions = mergeUnique([...(record.data.stageNameOptions || []).map(resolveStageNameAlias), ...DEFAULT_STAGE_NAME_OPTIONS.map(resolveStageNameAlias), ...stages.map((s) => s.name)]);
  stageIdCounter = record.data.stageIdCounter || 100;
  dependencyIdCounter = record.data.dependencyIdCounter || 100;
  normalizeDependencies(); saveStageNameOptions();
  applySettings(record.data.settings || {});
  renderAll();
  // 执行计算
  const result = runCalculation();
  saveWorkspace();
  currentViewingRecord = record.name;
  renderRecordViewingBanner();
  showToast(`已加载记录「${record.name}」`);
}

function deleteSavedRecord(idx) {
  const records = getSavedRecords();
  const name = records[idx]?.name || '';
  records.splice(idx, 1);
  setSavedRecords(records);
  renderSavedRecordsList();
  if (currentViewingRecord === name) { currentViewingRecord = null; renderRecordViewingBanner(); }
  showToast(`已删除记录「${name}」`);
}

function renderSavedRecordsList() {
  const wrap = document.getElementById('saved-records-list');
  if (!wrap) return;
  const records = getSavedRecords();
  if (!records.length) {
    wrap.innerHTML = '<div class="saved-records-empty">暂无保存的记录</div>';
    return;
  }
  wrap.innerHTML = records.map((r, i) => `
    <div class="saved-record-item${currentViewingRecord === r.name ? ' saved-record-active' : ''}" data-record-idx="${i}">
      <div class="saved-record-info" data-record-load="${i}">
        <span class="saved-record-name">${esc(r.name)}</span>
        <span class="saved-record-time">${fmtDateTime(r.savedAt)}</span>
      </div>
      <button class="btn-danger-sm saved-record-del" data-record-del="${i}" title="删除此记录">×</button>
    </div>
  `).join('');
  // 绑定事件
  wrap.querySelectorAll('[data-record-load]').forEach((el) => {
    el.addEventListener('click', () => loadSavedRecord(safeInt(el.dataset.recordLoad, -1)));
  });
  wrap.querySelectorAll('[data-record-del]').forEach((btn) => {
    btn.addEventListener('click', (e) => { e.stopPropagation(); deleteSavedRecord(safeInt(btn.dataset.recordDel, -1)); });
  });
}

function renderRecordViewingBanner() {
  const banner = document.getElementById('record-viewing-banner');
  if (!banner) return;
  if (currentViewingRecord) {
    banner.style.display = 'flex';
    banner.innerHTML = `<span class="record-banner-icon">📋</span><span>当前显示的是保存记录：<strong>${esc(currentViewingRecord)}</strong></span><button class="btn-sm record-banner-close" id="btn-close-banner">✕ 关闭提示</button>`;
    document.getElementById('btn-close-banner').addEventListener('click', () => {
      currentViewingRecord = null;
      renderRecordViewingBanner();
      renderSavedRecordsList();
    });
  } else {
    banner.style.display = 'none';
    banner.innerHTML = '';
  }
}

/* ============ 导入相关 ============ */
function normalizeText(text) {
  return String(text || '').replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 65248))
    .replace(/[：]/g, ':').replace(/[（]/g, '(').replace(/[）]/g, ')').replace(/[，]/g, ',').replace(/[；]/g, ';').replace(/\r/g, '');
}
function dedupeItems(items) {
  const map = new Map();
  items.forEach((item) => { const n = String(item.name || '').trim(); if (!n) return; map.set(n, { name: n, days: Math.max(1, safeInt(item.days, 5)) }); });
  return Array.from(map.values());
}
function parseStageItemsFromText(text) {
  const lines = normalizeText(text).split('\n').map((l) => l.trim()).filter(Boolean); const items = [];
  lines.forEach((line) => {
    let row = line.replace(/^[\d一二三四五六七八九十]+[、.）)]\s*/, '').replace(/^[\-•·]\s*/, '').trim(); if (!row) return;
    const parts = row.split(/[;,，；]+/).map((p) => p.trim()).filter(Boolean);
    (parts.length ? parts : [row]).forEach((part) => {
      let match = part.match(/^(.+?)[\s:：-]*([0-9]{1,3})(?:\s*(?:个?工作?日|天|d|day|days))?$/i);
      if (!match) { const cols = part.split(/\t+|\s{2,}/).map((v) => v.trim()).filter(Boolean); if (cols.length >= 2) { const d = cols.find((v) => /\d+/.test(v)) || '5'; const n = cols.filter((v) => v !== d).join(' ').trim(); if (n) match = [part, n, d.match(/\d+/)?.[0] || '5']; } }
      if (match) { items.push({ name: match[1].trim(), days: safeInt(match[2], 5) }); }
      else if (/^[\u4e00-\u9fa5A-Za-z0-9()（）_-]{2,40}$/.test(part) && !/^\d+$/.test(part)) { items.push({ name: part, days: 5 }); }
    });
  });
  return dedupeItems(items);
}
function buildExcelPreview(rows) {
  const show = rows.slice(0, 6); if (!show.length) return '<p class="placeholder-text">表格为空</p>';
  return `<div class="table-wrapper"><table class="result-table"><tbody>${show.map((r) => `<tr>${r.slice(0, 6).map((c) => `<td>${esc(c ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
function parseStageItemsFromSheetRows(rows) {
  if (!rows?.length) return [];
  const nameKeys = ['环节','环节名称','名称','节点','流程','stage','name'];
  const dayKeys = ['工期','天数','工作日','周期','days','duration'];
  let nameCol = 0, dayCol = 1, dataRows = rows;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const lower = rows[i].map((v) => String(v ?? '').trim().toLowerCase());
    const n = lower.findIndex((c) => nameKeys.some((k) => c.includes(k.toLowerCase())));
    const d = lower.findIndex((c) => dayKeys.some((k) => c.includes(k.toLowerCase())));
    if (n !== -1 || d !== -1) { if (n !== -1) nameCol = n; if (d !== -1) dayCol = d; dataRows = rows.slice(i + 1); break; }
  }
  const items = [];
  dataRows.forEach((row) => {
    const cells = row.map((v) => String(v ?? '').trim()); if (!cells.some(Boolean)) return;
    const name = (cells[nameCol] || cells.find((v) => /[\u4e00-\u9fa5A-Za-z]/.test(v)) || '').replace(/^[\d一二三四五六七八九十]+[、.）)]\s*/, '').trim();
    const dayText = cells[dayCol] || cells.find((v) => /\d+/.test(v)) || '5';
    if (name) items.push({ name, days: safeInt(dayText.match(/\d+/)?.[0], 5) });
  });
  return dedupeItems(items);
}
function renderImportResultList(type, items) {
  const list = document.getElementById(type === 'ocr' ? 'ocr-result-list' : 'excel-result-list');
  const area = document.getElementById(type === 'ocr' ? 'ocr-result-area' : 'excel-result-area');
  if (!list || !area) return; importBuffer[type] = items;
  if (!items.length) { list.innerHTML = '<p class="placeholder-text">未识别到可导入环节。</p>'; area.style.display = 'block'; return; }
  list.innerHTML = items.map((item, i) => `<div class="ocr-result-item" data-import-index="${i}"><input type="checkbox" checked><input class="ocr-name-input" type="text" value="${esc(item.name)}"><input class="ocr-days-input" type="number" min="1" value="${Math.max(1, safeInt(item.days, 5))}"><span>天</span></div>`).join('');
  area.style.display = 'block';
}
function collectImportItems(type) {
  const list = document.getElementById(type === 'ocr' ? 'ocr-result-list' : 'excel-result-list'); if (!list) return [];
  const items = [];
  list.querySelectorAll('.ocr-result-item').forEach((row) => {
    if (!row.querySelector('input[type="checkbox"]')?.checked) return;
    const name = row.querySelector('.ocr-name-input')?.value.trim();
    const days = Math.max(1, safeInt(row.querySelector('.ocr-days-input')?.value, 5));
    if (name) items.push({ name, days });
  });
  return dedupeItems(items);
}
function importStageItems(items, label) {
  const finalItems = dedupeItems(items); if (!finalItems.length) { showToast('没有可导入的环节', 'warning'); return; }
  pushHistory();
  let prevId = stages.length ? stages[stages.length - 1].id : null;
  finalItems.forEach((item) => {
    const newId = `stg_custom_${++stageIdCounter}`;
    stages.push({ id: newId, name: item.name, days: item.days, enabled: true, isDefault: false });
    upsertStageNameOption(item.name);
    if (prevId) {
      dependencies.push({ id: `dep_custom_${++dependencyIdCounter}`, fromId: prevId, toId: newId, linkType: 'FS', gapDays: 0 });
    }
    prevId = newId;
  });
  renderAll(); autoRunCalculation(); showToast(`${label}已导入 ${finalItems.length} 个环节`);
}

function bindFileImports() {
  const imgInput = document.getElementById('file-img');
  const excelInput = document.getElementById('file-excel');
  document.getElementById('btn-upload-img').addEventListener('click', () => imgInput.click());
  document.getElementById('btn-upload-excel').addEventListener('click', () => excelInput.click());
  imgInput.addEventListener('change', () => {
    const file = imgInput.files?.[0]; if (!file) return;
    modal('modal-ocr', true);
    document.getElementById('ocr-preview-area').innerHTML = `<img src="${URL.createObjectURL(file)}" alt="OCR预览"><div class="ocr-progress"><div>正在识别...</div><div class="progress-bar"><div class="progress-fill" style="width:0%"></div></div></div>`;
    document.getElementById('ocr-result-area').style.display = 'none';
    if (!window.Tesseract) { document.getElementById('ocr-preview-area').innerHTML = '<p class="placeholder-text">OCR 组件未加载成功。</p>'; imgInput.value = ''; return; }
    const token = ++ocrToken;
    Tesseract.recognize(file, 'chi_sim+eng', {
      logger: (msg) => { if (token !== ocrToken) return; if (msg.status === 'recognizing text') { const p = Math.max(0, Math.min(100, Math.round((msg.progress || 0) * 100))); const fill = document.querySelector('#ocr-preview-area .progress-fill'); if (fill) fill.style.width = `${p}%`; } }
    }).then(({ data }) => { if (token !== ocrToken) return; const items = parseStageItemsFromText(data?.text || ''); renderImportResultList('ocr', items); showToast(`识别完成，${items.length} 个候选环节`); })
    .catch(() => { document.getElementById('ocr-preview-area').innerHTML = '<p class="placeholder-text">图片识别失败。</p>'; showToast('图片识别失败', 'error'); })
    .finally(() => { imgInput.value = ''; });
  });
  excelInput.addEventListener('change', () => {
    const file = excelInput.files?.[0]; if (!file) return;
    modal('modal-excel', true);
    document.getElementById('excel-preview-area').innerHTML = '<p class="placeholder-text">解析中...</p>';
    document.getElementById('excel-result-area').style.display = 'none';
    if (!window.XLSX) { document.getElementById('excel-preview-area').innerHTML = '<p class="placeholder-text">Excel 组件未加载。</p>'; excelInput.value = ''; return; }
    const reader = new FileReader();
    reader.onload = (e) => {
      try { const wb = XLSX.read(e.target.result, { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }); const items = parseStageItemsFromSheetRows(rows); document.getElementById('excel-preview-area').innerHTML = buildExcelPreview(rows); renderImportResultList('excel', items); showToast(`解析完成，${items.length} 个候选环节`); }
      catch { document.getElementById('excel-preview-area').innerHTML = '<p class="placeholder-text">Excel 解析失败。</p>'; showToast('Excel 解析失败', 'error'); }
      excelInput.value = '';
    };
    reader.readAsArrayBuffer(file);
  });
  document.getElementById('btn-ocr-cancel').addEventListener('click', () => modal('modal-ocr', false));
  document.getElementById('btn-ocr-import').addEventListener('click', () => { importStageItems(collectImportItems('ocr'), '图片识别'); modal('modal-ocr', false); });
  document.getElementById('btn-excel-cancel').addEventListener('click', () => modal('modal-excel', false));
  document.getElementById('btn-excel-import').addEventListener('click', () => { importStageItems(collectImportItems('excel'), 'Excel'); modal('modal-excel', false); });
}

function syncDirectionTexts() {
  const dir = getCurrentDirection();
  document.getElementById('label-date').innerHTML = dir === 'forward' ? '需求开始日期 <span class="required-mark">*</span>' : '上线日期（倒推目标）<span class="required-mark">*</span>';
  // 更新对比toggle的文案：始终显示反方向名称
  const toggleWrap = document.getElementById('backward-toggle-wrap');
  if (toggleWrap) {
    const chk = document.getElementById('chk-backward');
    const wasChecked = chk?.checked || false;
    const oppositeLabel = dir === 'forward' ? '倒排' : '正推';
    toggleWrap.innerHTML = `<input type="checkbox" id="chk-backward" ${wasChecked ? 'checked' : ''}> 显示${oppositeLabel}对比结果`;
    // 重新绑定 change 事件
    document.getElementById('chk-backward').addEventListener('change', () => renderCompareResult(getCurrentDirection()));
  }
  // 更新顺排/倒排的label标签
  const forwardSummaryLabel = document.querySelector('#forward-summary-block .dual-result-label');
  const backwardSummaryLabel = document.querySelector('#backward-summary-block .dual-result-label');
  const forwardTableLabel = document.querySelector('#forward-table-block .dual-result-label');
  const backwardTableLabel = document.querySelector('#backward-table-block .dual-result-label');
  const forwardGanttLabel = document.getElementById('forward-gantt-label');
  const backwardGanttLabel = document.getElementById('backward-gantt-label');
  if (dir === 'forward') {
    if (forwardSummaryLabel) forwardSummaryLabel.textContent = '📈 顺排结果';
    if (backwardSummaryLabel) backwardSummaryLabel.textContent = '📉 倒排结果';
    if (forwardTableLabel) forwardTableLabel.textContent = '📈 顺排明细';
    if (backwardTableLabel) backwardTableLabel.textContent = '📉 倒排明细';
    if (forwardGanttLabel) forwardGanttLabel.textContent = '📈 顺排甘特图';
    if (backwardGanttLabel) backwardGanttLabel.textContent = '📉 倒排甘特图';
  } else {
    if (forwardSummaryLabel) forwardSummaryLabel.textContent = '📉 倒排结果';
    if (backwardSummaryLabel) backwardSummaryLabel.textContent = '📈 顺排结果';
    if (forwardTableLabel) forwardTableLabel.textContent = '📉 倒排明细';
    if (backwardTableLabel) backwardTableLabel.textContent = '📈 顺排明细';
    if (forwardGanttLabel) forwardGanttLabel.textContent = '📉 倒排甘特图';
    if (backwardGanttLabel) backwardGanttLabel.textContent = '📈 顺排甘特图';
  }
}

function bindStageModal() {
  let suppressNextStageNameFocus = false;

  function openAddStageModal() {
    document.getElementById('new-stage-name').value = ''; document.getElementById('new-stage-days').value = '';
    renderStageNameDropdown(''); modal('modal-stage', true);
    setTimeout(() => {
      const nameInput = document.getElementById('new-stage-name');
      if (!nameInput) return;
      suppressNextStageNameFocus = true;
      nameInput.focus({ preventScroll: true });
      toggleStageNameDropdown(false);
    }, 50);
  }
  document.getElementById('btn-add-stage').addEventListener('click', openAddStageModal);
  // 底部新增环节入口
  const bottomBtn = document.getElementById('btn-add-stage-bottom');
  if (bottomBtn) bottomBtn.addEventListener('click', openAddStageModal);
  document.getElementById('btn-stage-cancel').addEventListener('click', () => modal('modal-stage', false));
  document.getElementById('btn-stage-confirm').addEventListener('click', () => {
    const name = document.getElementById('new-stage-name').value.trim();
    const daysVal = document.getElementById('new-stage-days').value.trim();
    if (!name) { showToast('请输入环节名称', 'warning'); return; }
    if (!daysVal || safeInt(daysVal, 0) <= 0) { showToast('请输入工期天数', 'warning'); highlightEmptyField('new-stage-days'); return; }
    const days = Math.max(1, safeInt(daysVal, 1));
    pushHistory();
    const newId = `stg_custom_${++stageIdCounter}`;
    // 自动创建依赖：从最后一个环节到新环节
    const lastStage = stages.length ? stages[stages.length - 1] : null;
    stages.push({ id: newId, name, days, enabled: true, isDefault: false });
    if (lastStage) {
      dependencies.push({ id: `dep_custom_${++dependencyIdCounter}`, fromId: lastStage.id, toId: newId, linkType: 'FS', gapDays: 0 });
      if (hasCycle()) { dependencies.pop(); dependencyIdCounter--; }
    }
    upsertStageNameOption(name); modal('modal-stage', false); renderAll(); autoRunCalculation(); showToast('环节已添加');
  });
  document.getElementById('btn-toggle-dropdown').addEventListener('click', (e) => {
    e.stopPropagation(); const shown = document.getElementById('stage-name-dropdown').style.display === 'block';
    if (!shown) renderStageNameDropdown(document.getElementById('new-stage-name').value || '');
    toggleStageNameDropdown(!shown);
  });
  const newStageNameInput = document.getElementById('new-stage-name');
  newStageNameInput.addEventListener('input', (e) => { renderStageNameDropdown(e.target.value); toggleStageNameDropdown(true); });
  newStageNameInput.addEventListener('focus', (e) => {
    if (suppressNextStageNameFocus) {
      suppressNextStageNameFocus = false;
      return;
    }
    renderStageNameDropdown(e.target.value);
    toggleStageNameDropdown(true);
  });
  newStageNameInput.addEventListener('click', (e) => { renderStageNameDropdown(e.target.value); toggleStageNameDropdown(true); });
  document.getElementById('btn-manage-stage-names').addEventListener('click', () => { renderManageNameOptions(); modal('modal-manage-names', true); });
  document.getElementById('btn-close-manage-names').addEventListener('click', () => modal('modal-manage-names', false));
  document.getElementById('btn-add-option').addEventListener('click', () => {
    const input = document.getElementById('new-option-name'); const v = input.value.trim();
    if (!v) { showToast('请输入选项名称', 'warning'); return; }
    upsertStageNameOption(v); input.value = ''; renderManageNameOptions(); renderStageNameDropdown(document.getElementById('new-stage-name').value || ''); showToast('已添加');
  });
  document.addEventListener('click', (e) => { const box = document.getElementById('stage-name-container'); if (box && !box.contains(e.target)) toggleStageNameDropdown(false); });
}

function bindTemplateActions() {
  document.getElementById('btn-save-tpl').addEventListener('click', () => { document.getElementById('tpl-name-input').value = document.getElementById('template-select').value || ''; modal('modal-save-tpl', true); document.getElementById('tpl-name-input').focus(); });
  document.getElementById('btn-tpl-save-cancel').addEventListener('click', () => modal('modal-save-tpl', false));
  document.getElementById('btn-tpl-save-confirm').addEventListener('click', () => saveCurrentTemplate(document.getElementById('tpl-name-input').value));
  document.getElementById('btn-load-tpl').addEventListener('click', () => { loadSelectedTemplate(); hideTplHint(); });
  document.getElementById('btn-del-tpl').addEventListener('click', () => { deleteSelectedTemplate(); hideTplHint(); });
  // 选择模板时提示用户点击"加载"
  document.getElementById('template-select').addEventListener('change', (e) => {
    const hint = document.getElementById('tpl-select-hint');
    if (hint) {
      if (e.target.value) {
        hint.style.display = 'block';
        hint.innerHTML = `💡 已选择模板「<strong>${esc(e.target.value)}</strong>」，请点击 <strong>📥 加载</strong> 按钮使其生效`;
      } else {
        hint.style.display = 'none';
      }
    }
  });
}
function hideTplHint() { const h = document.getElementById('tpl-select-hint'); if (h) h.style.display = 'none'; }

function bindRecordActions() {
  document.getElementById('btn-open-save-record').addEventListener('click', () => {
    if (!lastPrimaryData) { showToast('请先计算排期后再保存', 'warning'); return; }
    const records = getSavedRecords();
    if (records.length >= MAX_SAVED_RECORDS) { showToast(`已达上限 ${MAX_SAVED_RECORDS} 条，请先删除旧记录`, 'warning'); return; }
    document.getElementById('record-name-input').value = '';
    modal('modal-save-record', true);
    document.getElementById('record-name-input').focus();
  });
  document.getElementById('btn-record-cancel').addEventListener('click', () => modal('modal-save-record', false));
  document.getElementById('btn-record-confirm').addEventListener('click', () => saveCurrentRecord(document.getElementById('record-name-input').value));
}

function bindBasicEvents() {
  const dateInput = document.getElementById('base-date');
  // 需求开始日期默认为空，必填
  dateInput.value = '';
  ['click', 'focus'].forEach((t) => dateInput.addEventListener(t, function () { if (this.showPicker) try { this.showPicker(); } catch {} }));
  document.getElementById('btn-calc').addEventListener('click', runCalculation);
  const clearConfigBtn = document.getElementById('btn-clear-config');
  if (clearConfigBtn) clearConfigBtn.addEventListener('click', clearCurrentConfiguration);
  // 分界引导icon也支持点击触发计算
  const dividerCalcBtn = document.getElementById('divider-calc-btn');
  if (dividerCalcBtn) dividerCalcBtn.addEventListener('click', runCalculation);
  document.querySelectorAll('input[name="direction"]').forEach((input) => input.addEventListener('change', () => { syncDirectionTexts(); bindCompareToggle(); autoRunCalculation(); }));
  document.querySelectorAll('input[name="workday"]').forEach((input) => input.addEventListener('change', autoRunCalculation));
  document.getElementById('base-date').addEventListener('change', autoRunCalculation);
  document.getElementById('extra-days').addEventListener('input', autoRunCalculation);
  const expectedEndInput = document.getElementById('expected-end-date');
  if (expectedEndInput) {
    ['click', 'focus'].forEach((t) => expectedEndInput.addEventListener(t, function () { if (this.showPicker) try { this.showPicker(); } catch {} }));
    expectedEndInput.addEventListener('change', autoRunCalculation);
  }
  document.getElementById('btn-fwd-zoom-in').addEventListener('click', () => { ganttZoom = Math.min(2, ganttZoom + 0.1); applyGanttZoom(); });
  document.getElementById('btn-fwd-zoom-out').addEventListener('click', () => { ganttZoom = Math.max(0.5, ganttZoom - 0.1); applyGanttZoom(); });
  document.getElementById('btn-fwd-zoom-reset').addEventListener('click', () => { ganttZoom = 1; applyGanttZoom(); });
  document.getElementById('btn-bwd-zoom-in').addEventListener('click', () => { ganttZoomBwd = Math.min(2, ganttZoomBwd + 0.1); applyGanttZoom(); });
  document.getElementById('btn-bwd-zoom-out').addEventListener('click', () => { ganttZoomBwd = Math.max(0.5, ganttZoomBwd - 0.1); applyGanttZoom(); });
  document.getElementById('btn-bwd-zoom-reset').addEventListener('click', () => { ganttZoomBwd = 1; applyGanttZoom(); });
  document.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); undoLastAction(); } });
  document.querySelectorAll('.modal-overlay').forEach((ov) => ov.addEventListener('click', (e) => { if (e.target === ov) ov.style.display = 'none'; }));
  // 导出按钮
  document.getElementById('btn-export-table-excel').addEventListener('click', () => exportTableToExcel());
  document.getElementById('btn-export-table-img').addEventListener('click', () => exportElementToImage('sec-table', '时间明细表'));
  // 顺排甘特图导出
  document.getElementById('btn-export-fwd-gantt-excel').addEventListener('click', () => exportGanttToExcel(lastPrimaryData, '顺排甘特图'));
  document.getElementById('btn-export-fwd-gantt-img').addEventListener('click', () => exportGanttContentToImage('gantt-content', '顺排甘特图', ganttZoom));
  // 倒排甘特图导出
  document.getElementById('btn-export-bwd-gantt-excel').addEventListener('click', () => exportGanttToExcel(lastCompareData, '倒排甘特图'));
  document.getElementById('btn-export-bwd-gantt-img').addEventListener('click', () => exportGanttContentToImage('backward-gantt-content', '倒排甘特图', ganttZoomBwd));
  // 全屏展开按钮
  document.getElementById('btn-fwd-gantt-fullscreen').addEventListener('click', () => openGanttFullscreen('gantt-content', '顺排甘特图'));
  document.getElementById('btn-bwd-gantt-fullscreen').addEventListener('click', () => openGanttFullscreen('backward-gantt-content', '倒排甘特图'));
}

/* ============ 导出功能 ============ */
function exportTableToExcel() {
  if (!lastPrimaryData) { showToast('请先计算排期', 'warning'); return; }
  const week = ['日','一','二','三','四','五','六'];
  const data = lastPrimaryData;
  const rows = [['环节', '工期（天）', '开始日期', '结束日期', '前置依赖']];
  data.results.forEach((item) => {
    const deps = data.dependencies.filter((d) => d.toId === item.id);
    const depText = deps.length ? deps.map((d) => {
      if (d.fromId === 'none') return '无依赖';
      const fromName = data.results.find((r) => r.id === d.fromId)?.name || getStageById(d.fromId)?.name || '未知';
      const meta = LINK_META[d.linkType] || LINK_META.FS;
      return `${fromName}(${meta.label})${d.gapDays ? '+' + d.gapDays + '天' : ''}`;
    }).join('; ') : '—';
    rows.push([item.name, item.days, `${fmt(item.startDate)}（周${week[item.startDate.getDay()]}）`, `${fmt(item.endDate)}（周${week[item.endDate.getDay()]}）`, depText]);
  });
  if (!window.XLSX) { showToast('Excel组件未加载', 'error'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  // 设置列宽
  ws['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 25 }, { wch: 25 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws, '排期明细');
  XLSX.writeFile(wb, `排期明细_${fmt(new Date())}.xlsx`);
  showToast('Excel导出成功');
}

function exportGanttToExcel(inputData, title) {
  const data = inputData || lastPrimaryData;
  if (!data) { showToast('请先计算排期', 'warning'); return; }
  const sheetTitle = title || '甘特图';
  const start = addDays(data.projectStart, -1);
  const end = addDays(data.projectEnd, 3);
  const days = []; for (let c = new Date(start); c <= end; c = addDays(c, 1)) days.push(new Date(c));
  // 表头：环节 + 各日期
  const header = ['环节', ...days.map((d) => fmt(d))];
  const rows = [header];
  data.results.forEach((item) => {
    const row = [item.name];
    days.forEach((d) => {
      if (d >= item.startDate && d <= item.endDate && isWorkday(d, data.rule)) {
        row.push('■');
      } else {
        row.push('');
      }
    });
    rows.push(row);
  });
  if (!window.XLSX) { showToast('Excel组件未加载', 'error'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
  XLSX.writeFile(wb, `${sheetTitle}_${fmt(new Date())}.xlsx`);
  showToast(`${sheetTitle}Excel导出成功`);
}

function exportElementToImage(elementId, title) {
  const el = document.getElementById(elementId);
  if (!el) { showToast('内容不存在', 'error'); return; }
  if (!window.html2canvas) { showToast('截图组件未加载', 'error'); return; }
  showToast('正在生成图片...', 'success');
  html2canvas(el, {
    backgroundColor: '#ffffff',
    scale: 2,
    useCORS: true,
    scrollX: 0,
    scrollY: 0,
    windowWidth: el.scrollWidth + 40,
    windowHeight: el.scrollHeight + 40,
  }).then((canvas) => {
    const link = document.createElement('a');
    link.download = `${title}_${fmt(new Date())}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('图片已保存');
  }).catch(() => {
    showToast('图片生成失败', 'error');
  });
}

function exportGanttContentToImage(wrapperId, title, zoom) {
  const wrapper = document.getElementById(wrapperId);
  if (!wrapper) { showToast('内容不存在', 'error'); return; }
  const outer = wrapper.querySelector('[data-gantt-outer]');
  if (!outer) { showToast('甘特图内容不存在', 'error'); return; }
  if (!window.html2canvas) { showToast('截图组件未加载', 'error'); return; }
  showToast('正在生成图片...', 'success');
  // 临时恢复缩放为1，保存完整甘特图
  const origTransform = outer.style.transform;
  outer.style.transform = 'scale(1)';
  // 临时让wrapper不限制高度和滚动
  const origMaxH = wrapper.style.maxHeight;
  const origOverflow = wrapper.style.overflow;
  wrapper.style.maxHeight = 'none';
  wrapper.style.overflow = 'visible';
  // 临时降低"今天"色块透明度到40%
  const todayCells = outer.querySelectorAll('.col-today');
  const todayOrigStyles = [];
  todayCells.forEach((cell) => {
    todayOrigStyles.push({ bg: cell.style.backgroundColor, boxShadow: cell.style.boxShadow });
    cell.style.backgroundColor = 'rgba(255, 251, 235, 0.4)';
    cell.style.boxShadow = 'inset 2px 0 0 rgba(245, 158, 11, 0.4)';
  });
  html2canvas(outer, {
    backgroundColor: '#ffffff',
    scale: 2,
    useCORS: true,
    scrollX: 0,
    scrollY: 0,
    width: outer.scrollWidth,
    height: outer.scrollHeight,
    windowWidth: outer.scrollWidth + 40,
    windowHeight: outer.scrollHeight + 40,
  }).then((canvas) => {
    const link = document.createElement('a');
    link.download = `${title}_${fmt(new Date())}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('图片已保存');
  }).catch(() => {
    showToast('图片生成失败', 'error');
  }).finally(() => {
    outer.style.transform = origTransform;
    wrapper.style.maxHeight = origMaxH;
    wrapper.style.overflow = origOverflow;
    // 恢复"今天"色块样式
    todayCells.forEach((cell, i) => {
      cell.style.backgroundColor = todayOrigStyles[i].bg;
      cell.style.boxShadow = todayOrigStyles[i].boxShadow;
    });
  });
}

/* ============ 甘特图全屏展开 ============ */
function openGanttFullscreen(wrapperId, title) {
  const wrapper = document.getElementById(wrapperId);
  if (!wrapper) { showToast('甘特图内容不存在', 'error'); return; }
  const outer = wrapper.querySelector('[data-gantt-outer]');
  if (!outer) { showToast('甘特图未生成', 'error'); return; }

  // 创建覆盖层（半透明背景）
  const overlay = document.createElement('div');
  overlay.className = 'gantt-fullscreen-overlay';
  overlay.innerHTML = `
    <div class="gantt-fullscreen-inner">
      <div class="gantt-fullscreen-header">
        <div class="gantt-fullscreen-title">${title}</div>
        <div class="gantt-fullscreen-actions">
          <button class="btn-sm" id="fs-zoom-out">−</button>
          <span id="fs-zoom-level">100%</span>
          <button class="btn-sm" id="fs-zoom-in">+</button>
          <button class="btn-sm" id="fs-zoom-reset">重置</button>
          <button class="btn-close-fullscreen" id="fs-close">✕ 关闭</button>
        </div>
      </div>
      <div class="gantt-fullscreen-body" id="fs-body"></div>
    </div>
  `;
  document.body.appendChild(overlay);
  document.body.style.overflow = 'hidden';

  // 点击半透明遮罩区域关闭（不点inner内部）
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) { overlay.remove(); document.body.style.overflow = ''; }
  });

  // 克隆gantt-outer到全屏body
  const cloned = outer.cloneNode(true);
  cloned.style.transform = 'scale(1)';
  cloned.style.transformOrigin = 'top left';
  document.getElementById('fs-body').appendChild(cloned);

  let fsZoom = 1;
  function applyFsZoom() {
    cloned.style.transform = `scale(${fsZoom})`;
    document.getElementById('fs-zoom-level').textContent = `${Math.round(fsZoom * 100)}%`;
  }

  document.getElementById('fs-zoom-in').addEventListener('click', () => { fsZoom = Math.min(3, fsZoom + 0.1); applyFsZoom(); });
  document.getElementById('fs-zoom-out').addEventListener('click', () => { fsZoom = Math.max(0.3, fsZoom - 0.1); applyFsZoom(); });
  document.getElementById('fs-zoom-reset').addEventListener('click', () => { fsZoom = 1; applyFsZoom(); });
  document.getElementById('fs-close').addEventListener('click', () => {
    overlay.remove();
    document.body.style.overflow = '';
  });
  // ESC 关闭
  function onEsc(e) { if (e.key === 'Escape') { overlay.remove(); document.body.style.overflow = ''; document.removeEventListener('keydown', onEsc); } }
  document.addEventListener('keydown', onEsc);
}

function bindCompareToggle() {
  const toggleWrap = document.getElementById('backward-toggle-wrap'); if (!toggleWrap) return;
  const checked = document.getElementById('chk-backward')?.checked;
  const dir = getCurrentDirection();
  const oppositeLabel = dir === 'forward' ? '倒排' : '正推';
  toggleWrap.innerHTML = `<input type="checkbox" id="chk-backward" ${checked ? 'checked' : ''}> 显示${oppositeLabel}对比结果`;
  document.getElementById('chk-backward').addEventListener('change', () => renderCompareResult(getCurrentDirection()));
}

/* ============ 新手引导 ============ */
const GUIDE_KEY = 'schedule_tool_guide_shown_v2';
const GUIDE_STEPS = [
  { target: '#sec-basic', title: '① 基础设置', desc: '在这里设置计算方向、需求开始日期等基础参数。需求开始日期为必填项。', position: 'right' },
  { target: '#sec-stages', title: '② 流程环节', desc: '配置每个制作环节的名称和工期。工期为必填项，可拖拽 ≡ 调整顺序。底部也有快捷新增入口。', position: 'right' },
  { target: '#sec-dependencies', title: '③ 流程依赖', desc: '设置环节之间的先后依赖关系。开启"高级"可切换连接方式（如首→首、尾→尾）。', position: 'right' },
  { target: '#btn-calc', title: '④ 开始计算', desc: '点击此按钮计算排期。修改左侧配置后也会自动触发重算。', position: 'right' },
  { target: '#sec-summary', title: '⑤ 项目汇总', desc: '查看计算后的排期汇总信息，包括起止日期、启用环节数等。可勾选显示倒排结果做对比。', position: 'left' },
  { target: '#sec-table', title: '⑥ 时间明细表', desc: '每个环节的详细开始/结束时间一目了然。支持导出Excel和保存为图片。', position: 'left' },
  { target: '#sec-gantt-forward', title: '⑦ 甘特图', desc: '直观的可视化甘特图，可拖拽色条调整工期，支持缩放、导出和全屏查看。顺排和倒排各有独立的甘特图。', position: 'left' }
];

let guideOverlay = null, guideHighlight = null, guideTooltip = null, guideActiveTarget = null;

function initGuide() {
  if (localStorage.getItem(GUIDE_KEY)) return;
  showGuide();
}

function cleanupGuide() {
  [guideOverlay, guideHighlight, guideTooltip].forEach((el) => { if (el) el.remove(); });
  guideOverlay = guideHighlight = guideTooltip = null;
  if (guideActiveTarget) guideActiveTarget.classList.remove('guide-target-active');
  guideActiveTarget = null;
}

function showGuide(stepIdx = 0) {
  if (stepIdx >= GUIDE_STEPS.length) {
    if (guideTooltip) { guideTooltip.style.opacity = '0'; guideTooltip.style.transform = 'translateY(10px)'; }
    setTimeout(() => { cleanupGuide(); localStorage.setItem(GUIDE_KEY, '1'); showToast('🎉 引导完成！开始使用吧'); }, 400);
    return;
  }

  const step = GUIDE_STEPS[stepIdx];
  const targetEl = document.querySelector(step.target);
  if (!targetEl) { showGuide(stepIdx + 1); return; }

  // 滚动到目标元素
  targetEl.scrollIntoView({ behavior: 'smooth', block: 'center' });

  setTimeout(() => {
    if (guideActiveTarget && guideActiveTarget !== targetEl) guideActiveTarget.classList.remove('guide-target-active');
    guideActiveTarget = targetEl;
    guideActiveTarget.classList.add('guide-target-active');

    if (!guideTooltip) {
      guideTooltip = document.createElement('div');
      guideTooltip.className = 'guide-tooltip guide-tooltip-entering';
      document.body.appendChild(guideTooltip);
    }

    guideTooltip.classList.remove('guide-tooltip-visible');
    guideTooltip.classList.add('guide-tooltip-entering');

    setTimeout(() => {
      guideTooltip.innerHTML = `
        <div class="guide-tooltip-title">${step.title}</div>
        <div class="guide-tooltip-desc">${step.desc}</div>
        <div class="guide-tooltip-footer">
          <span class="guide-tooltip-progress">${stepIdx + 1} / ${GUIDE_STEPS.length}</span>
          <div class="guide-tooltip-btns">
            <button class="guide-btn-skip" id="guide-btn-skip">跳过引导</button>
            <button class="guide-btn-next" id="guide-btn-next">${stepIdx < GUIDE_STEPS.length - 1 ? '下一步 →' : '完成 ✓'}</button>
          </div>
        </div>
      `;

      requestAnimationFrame(() => {
        guideTooltip.classList.remove('guide-tooltip-entering');
        guideTooltip.classList.add('guide-tooltip-visible');
      });

      document.getElementById('guide-btn-next').addEventListener('click', () => showGuide(stepIdx + 1));
      document.getElementById('guide-btn-skip').addEventListener('click', () => {
        cleanupGuide();
        localStorage.setItem(GUIDE_KEY, '1');
        showToast('已跳过引导。如需重新查看，可清除浏览器缓存。');
      });
    }, 80);
  }, 260);
}

function renderAll() { renderStages(); renderDependencies(); refreshTemplateSelect(document.getElementById('template-select')?.value || ''); }

/* ============ URL 参数预填功能 ============ */
// 支持 ?preset=base64JSON 方式传入排期数据，自动填充环节、依赖和设置并计算
function loadPresetFromURL() {
  try {
    const params = new URLSearchParams(window.location.search);
    const preset = params.get('preset');
    if (!preset) return null;
    const json = JSON.parse(decodeURIComponent(atob(preset)));
    if (!json || !Array.isArray(json.stages) || json.stages.length === 0) return null;
    if (json.stageNameAliases) {
      stageNameAliases = clone(json.stageNameAliases);
      saveStageNameAliases();
    }
    stages = clone(json.stages);
    dependencies = clone(json.dependencies || []);
    stageIdCounter = json.stageIdCounter || Math.max(100, ...stages.map((s) => parseInt(String(s.id).replace(/\D/g, ''), 10) || 0)) + 1;
    dependencyIdCounter = json.dependencyIdCounter || Math.max(100, ...dependencies.map((d) => parseInt(String(d.id).replace(/\D/g, ''), 10) || 0)) + 1;
    normalizeDependencies();
    const settings = json.settings || {};
    setJSON(STORAGE_KEYS.workspace, { stages: clone(stages), dependencies: clone(dependencies), stageNameOptions: clone(stageNameOptions), stageNameAliases: clone(stageNameAliases), stageIdCounter, dependencyIdCounter, settings });
    return settings;
  } catch (e) {
    console.warn('[preset] URL参数解析失败:', e);
    return null;
  }
}

function init() {
  loadStageNameOptions();
  // 优先检查 URL 参数预填数据
  const presetSettings = loadPresetFromURL();
  // 恢复上次保存的工作区状态
  const savedSettings = presetSettings || loadWorkspace();
  if (!savedSettings) {
    stages = getDefaultStagesWithAliases();
    dependencies = clone(DEFAULT_DEPENDENCIES);
  }
  renderAll(); renderStageNameDropdown('');
  bindBasicEvents(); bindStageModal(); bindTemplateActions(); bindFileImports(); bindRecordActions();
  if (savedSettings) {
    applySettings(savedSettings);
  } else {
    syncDirectionTexts(); bindCompareToggle();
  }
  renderSavedRecordsList();
  // 恢复后自动计算（URL预填或工作区恢复都触发）
  if (savedSettings && document.getElementById('base-date').value && getEnabledStages().length) {
    setTimeout(() => runCalculation(), 100);
  }
  // 依赖关系高级开关
  const advChk = document.getElementById('chk-dep-advanced');
  if (advChk) advChk.addEventListener('change', () => renderDependencies());
  // 首次打开新手引导（URL预填模式跳过引导）
  if (!presetSettings) setTimeout(() => initGuide(), 500);
}
document.addEventListener('DOMContentLoaded', init);
