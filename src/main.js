import * as XLSX from 'xlsx';
import './styles.css';

const FIELD_DEFS = [
  { id: 'branch', label: 'الفرع', aliases: ['الفرع'] },
  { id: 'registration', label: 'رقم القيد', aliases: ['رقم القيد'] },
  { id: 'type', label: 'نوع المعاملة', aliases: ['نوع المعاملة'] },
  { id: 'importance', label: 'الأهمية', aliases: ['الاهمية', 'الأهمية'] },
  { id: 'movement', label: 'رقم الحركة', aliases: ['رقم الحركة'] },
  { id: 'department', label: 'اسم الجهة', aliases: ['اسم الجهة'] },
  { id: 'duration', label: 'المدة المحددة', aliases: ['المدة المحددة'] },
  { id: 'durationType', label: 'نوع المدة', aliases: ['نوع المدة'] },
  { id: 'requiredAction', label: 'الإجراء المطلوب', aliases: ['الاجراء المطلوب', 'الإجراء المطلوب'] },
  { id: 'takenAction', label: 'الإجراء المتخذ', aliases: ['الاجراء المتخذ', 'الإجراء المتخذ'] },
  { id: 'sentDate', label: 'تاريخ الإرسال/الاستلام', aliases: ['تاريخ الارسال/الاستلام', 'تاريخ الإرسال/الاستلام'] },
  { id: 'endDate', label: 'تاريخ الانتهاء', aliases: ['تاريخ الانتهاء'] },
  { id: 'overdue', label: 'متأخرة', aliases: ['متأخرة'] },
  { id: 'status', label: 'حالة التنفيذ', aliases: ['حالة التنفيذ'] },
  { id: 'subject', label: 'الموضوع', aliases: ['الموضوع'] },
  { id: 'beneficiary', label: 'المستفيد', aliases: ['المستفيد'] },
  { id: 'phone', label: 'رقم الهاتف', aliases: ['رقم الهاتف'] }
];

const SEARCH_FIELDS = [
  { id: 'all', label: 'كل الحقول' },
  { id: 'registration', label: 'رقم القيد' },
  { id: 'subject', label: 'الموضوع' },
  { id: 'beneficiary', label: 'المستفيد' },
  { id: 'phone', label: 'رقم الهاتف' },
  { id: 'movement', label: 'رقم الحركة' },
  { id: 'department', label: 'اسم الجهة' }
];

const state = {
  records: [],
  headers: [],
  source: '',
  sheet: '',
  view: 'dashboard',
  query: '',
  searchField: 'all',
  branch: 'all',
  status: 'all',
  importance: 'all',
  page: 1,
  pageSize: 20,
  selectedIndex: null,
  editMode: false,
  toast: null
};

const app = document.querySelector('#app');
const STORAGE_KEY = 'mualajat-data-v1';
const SETTINGS_KEY = 'mualajat-settings-v1';

const clean = (value) => String(value ?? '').trim();
const normalizeText = (value) => clean(value).toLowerCase().replace(/[\u064B-\u065F\u0670]/g, '').replace(/[إأآا]/g, 'ا').replace(/ة/g, 'ه');
const numberValue = (value) => {
  const number = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(number) ? number : 0;
};
const formatNumber = (number) => new Intl.NumberFormat('ar-EG').format(number || 0);
const escapeHtml = (value) => clean(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
const getHeaderKeys = () => state.headers.map((header, index) => clean(header) || `عمود ${index + 1}`);
const fieldDef = (id) => FIELD_DEFS.find((field) => field.id === id);

function normalizeKey(value) {
  return normalizeText(value).replace(/[\s_\-]/g, '');
}

function findKey(record, aliases = []) {
  const keys = Object.keys(record);
  const wanted = aliases.map(normalizeKey);
  return keys.find((key) => wanted.includes(normalizeKey(key))) || keys.find((key) => wanted.some((alias) => normalizeKey(key).includes(alias)));
}

function valueOf(record, id) {
  const field = fieldDef(id);
  if (!field) return '';
  const key = findKey(record, field.aliases);
  return key ? clean(record[key]) : '';
}

function isDone(record) {
  const value = normalizeText(valueOf(record, 'status'));
  return ['1', 'تم', 'منجز', 'مكتمل', 'مكتملة', 'مغلق', 'مغلقة', 'منتهى'].some((token) => value === token || value.includes(token));
}

function isLate(record) {
  const value = normalizeText(valueOf(record, 'overdue'));
  return ['1', 'نعم', 'متأخر', 'متاخر', 'متأخرة', 'متاخره'].some((token) => value === token || value.includes(token));
}

function displayStatus(record) {
  return isDone(record) ? 'مكتملة' : 'قيد التنفيذ';
}

function formatDate(value) {
  const text = clean(value);
  if (!text) return '—';
  const date = new Date(text);
  if (!Number.isNaN(date.getTime()) && /[-/T]/.test(text)) {
    return new Intl.DateTimeFormat('ar-EG', { year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
  }
  return text.replace(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/, '$1/$2/$3');
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ source: state.source, sheet: state.sheet, headers: state.headers, records: state.records }));
}

function showToast(message, kind = 'success') {
  state.toast = { message, kind };
  render();
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => { state.toast = null; render(); }, 2800);
}

async function loadData() {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (stored?.records?.length) {
      state.source = stored.source || 'نسخة محلية';
      state.sheet = stored.sheet || '';
      state.headers = stored.headers || Object.keys(stored.records[0]);
      state.records = stored.records;
      return;
    }
  } catch { /* ignore corrupt local data and use the bundled source */ }

  const response = await fetch('/data.json');
  const payload = await response.json();
  state.source = payload.source || 'data.json';
  state.sheet = payload.sheet || '';
  state.headers = payload.headers || Object.keys(payload.records?.[0] || {});
  state.records = payload.records || [];
  persist();
}

function registerOffline() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {});
}

function uniqueValues(id) {
  return [...new Set(state.records.map((record) => valueOf(record, id)).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ar'));
}

function matchesQuery(record) {
  const query = normalizeText(state.query);
  if (!query) return true;
  const ids = state.searchField === 'all' ? SEARCH_FIELDS.filter((field) => field.id !== 'all').map((field) => field.id) : [state.searchField];
  return ids.some((id) => normalizeText(valueOf(record, id)).includes(query));
}

function getFilteredEntries() {
  return state.records.map((record, index) => ({ record, index })).filter(({ record }) => {
    const branchMatch = state.branch === 'all' || valueOf(record, 'branch') === state.branch;
    const statusMatch = state.status === 'all' || (state.status === 'done' ? isDone(record) : !isDone(record));
    const importanceMatch = state.importance === 'all' || valueOf(record, 'importance') === state.importance;
    return branchMatch && statusMatch && importanceMatch && matchesQuery(record);
  }).sort((a, b) => numberValue(valueOf(a.record, 'movement')) - numberValue(valueOf(b.record, 'movement')));
}

function statsFor(entries) {
  return {
    total: entries.length,
    done: entries.filter(({ record }) => isDone(record)).length,
    pending: entries.filter(({ record }) => !isDone(record)).length,
    late: entries.filter(({ record }) => isLate(record)).length
  };
}

function statCard(label, value, tone, icon, foot) {
  return `<article class="stat-card ${tone}"><div class="stat-top"><span>${label}</span><b>${icon}</b></div><strong class="stat-value">${formatNumber(value)}</strong><div class="stat-foot">${foot}</div></article>`;
}

function renderShell() {
  const pageTitle = { dashboard: 'لوحة المتابعة', search: 'استعلام المعاملات', records: 'السجل الكامل', settings: 'إعدادات البيانات' }[state.view];
  app.innerHTML = `<div class="app-shell">
    <aside class="sidebar">
      <div class="brand"><div class="brand-mark">م</div><div class="brand-text"><strong>مُتابع المعاملات</strong><span>مركز العمليات المحلي</span></div></div>
      <nav class="nav"><div class="nav-label">التنقل الرئيسي</div>
        ${navButton('dashboard', '▦', 'لوحة المتابعة')}
        ${navButton('search', '⌕', 'استعلام وتتبع')}
        ${navButton('records', '▤', 'السجل الكامل')}
        ${navButton('settings', '⚙', 'إعدادات البيانات')}
      </nav>
      <div class="sidebar-footer"><span class="offline-dot"></span><strong>يعمل محليًا</strong><br />لا حاجة لاتصال الإنترنت بعد تثبيت التطبيق.</div>
    </aside>
    <main class="main">
      <header class="topbar"><div class="topbar-title"><div><h1>${pageTitle}</h1><p>آخر تحديث محلي: ${new Intl.DateTimeFormat('ar-EG', { dateStyle: 'medium' }).format(new Date())}</p></div></div>
        <div class="topbar-actions"><button class="secondary-button" data-action="export" title="تصدير البيانات">⇩ <span>تصدير</span></button><button class="primary-button" data-action="import" title="استيراد Excel">＋ <span>استيراد Excel</span></button><input id="file-input" type="file" accept=".xlsx,.xls,.csv" hidden /></div>
      </header>
      <div class="content" id="content"></div>
    </main>
  </div>
  <div id="modal-root"></div>
  ${state.toast ? `<div class="toast ${state.toast.kind}">${escapeHtml(state.toast.message)}</div>` : ''}`;
}

function navButton(view, icon, label) {
  return `<button class="nav-button ${state.view === view ? 'active' : ''}" data-view="${view}"><span class="nav-icon">${icon}</span>${label}</button>`;
}

function renderFilters(includeHeading = true) {
  const branches = uniqueValues('branch');
  const importances = uniqueValues('importance');
  return `${includeHeading ? `<div class="quick-search"><div class="search-wrap"><span class="search-glyph">⌕</span><input id="query" value="${escapeHtml(state.query)}" placeholder="ابحث برقم القيد، الموضوع، المستفيد أو الهاتف" /></div><select class="filter-select" id="search-field">${SEARCH_FIELDS.map((field) => `<option value="${field.id}" ${state.searchField === field.id ? 'selected' : ''}>${field.label}</option>`).join('')}</select></div>` : ''}
  <div class="filter-row"><label>تضييق النتائج الحالية:</label><select class="filter-select" id="branch-filter"><option value="all">كل الفروع</option>${branches.map((value) => `<option value="${escapeHtml(value)}" ${state.branch === value ? 'selected' : ''}>الفرع ${escapeHtml(value)}</option>`).join('')}</select><select class="filter-select" id="status-filter"><option value="all" ${state.status === 'all' ? 'selected' : ''}>كل الحالات</option><option value="done" ${state.status === 'done' ? 'selected' : ''}>مكتملة</option><option value="pending" ${state.status === 'pending' ? 'selected' : ''}>قيد التنفيذ</option></select><select class="filter-select" id="importance-filter"><option value="all">كل درجات الأهمية</option>${importances.map((value) => `<option value="${escapeHtml(value)}" ${state.importance === value ? 'selected' : ''}>${escapeHtml(value)}</option>`).join('')} </select><button class="secondary-button" data-action="clear-filters">مسح التصفية</button></div>`;
}

function renderDashboard() {
  const entries = getFilteredEntries();
  const stats = statsFor(entries);
  const branchCounts = countBy(entries, 'branch').slice(0, 6);
  const departmentCounts = countBy(entries, 'department').slice(0, 6);
  const recent = entries.slice(0, 7);
  return `<section class="hero"><div><p class="eyebrow">نظرة تشغيلية سريعة</p><h2>اعرف أين وصلت المعاملة</h2><p>ملخص حي للسجلات المحلية، مع ترتيب رقم الحركة تصاعديًا وفلترة فورية حسب الفرع والحالة والأهمية.</p></div><div class="hero-note"><span>مصدر البيانات الحالي</span><strong>${escapeHtml(state.source || 'نسخة محلية')}</strong><span style="margin:7px 0 0">${formatNumber(state.records.length)} سجل محفوظ على هذا الجهاز</span></div></section>
  ${renderFilters(false)}
  <section class="stats">${statCard('إجمالي المعاملات', stats.total, '', 'Σ', 'ضمن التصفية الحالية')}${statCard('مكتملة', stats.done, 'green', '✓', stats.total ? `${Math.round((stats.done / stats.total) * 100)}٪ من النتائج` : 'لا توجد نتائج')}${statCard('قيد التنفيذ', stats.pending, 'gold', '◷', 'تحتاج إلى متابعة')}${statCard('متأخرة', stats.late, 'red', '!', 'بحسب حقل متأخرة في الملف')}</section>
  <section class="section-grid"><article class="panel"><div class="panel-header"><div><h3 class="panel-title">توزيع المعاملات حسب الفرع</h3><p class="panel-subtitle">أكثر الفروع حضورًا في النتائج الحالية</p></div><button class="panel-link" data-view="records">عرض السجل ←</button></div>${renderBars(branchCounts, stats.total, 'branch', 'لا تتوفر بيانات فروع')}</article><article class="panel"><div class="panel-header"><div><h3 class="panel-title">الجهات المحيلة/المستقبلة</h3><p class="panel-subtitle">بحسب اسم الجهة في السجلات</p></div><button class="panel-link" data-view="search">استعلام ←</button></div>${renderBars(departmentCounts, stats.total, 'department', 'لا تتوفر بيانات جهات')}</article></section>
  <section class="panel table-panel" style="margin-top:18px"><div class="table-toolbar"><div><h3>آخر السجلات مرتبة برقم الحركة</h3><span>اضغط على أي صف لعرض تفاصيل المسار وتعديل السجل محليًا</span></div><button class="secondary-button" data-view="records">فتح السجل الكامل</button></div>${renderTable(recent, false)}</section>`;
}

function countBy(entries, id) {
  const map = new Map();
  entries.forEach(({ record }) => { const value = valueOf(record, id) || 'غير محدد'; map.set(value, (map.get(value) || 0) + 1); });
  return [...map.entries()].sort((a, b) => b[1] - a[1]).map(([label, count]) => ({ label, count }));
}

function renderBars(items, total, id, emptyMessage) {
  if (!items.length) return `<div class="empty">${emptyMessage}</div>`;
  return `<div class="bar-list">${items.map(({ label, count }) => `<div class="bar-row"><span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.max(3, Math.round((count / Math.max(total, 1)) * 100))}%"></div></div><span class="bar-value">${formatNumber(count)}</span></div>`).join('')}</div>`;
}

function renderSearch() {
  const entries = getFilteredEntries();
  const stats = statsFor(entries);
  return `<section class="hero"><div><p class="eyebrow">بحث متعدد المعايير</p><h2>ابحث وتابع مسار المعاملة</h2><p>اكتب رقم القيد أو الموضوع أو اسم المستفيد أو رقم الهاتف للوصول إلى السجل والجهة الحالية ورقم الحركة.</p></div><div class="hero-note"><span>عدد النتائج</span><strong>${formatNumber(stats.total)} معاملة</strong><span style="margin:7px 0 0">مرتبة تصاعديًا حسب رقم الحركة</span></div></section>${renderFilters(true)}<section class="panel table-panel"><div class="table-toolbar"><div><h3>نتائج الاستعلام</h3><span>${state.query ? `نتائج البحث عن «${escapeHtml(state.query)}»` : 'كل السجلات المطابقة للفلاتر'}</span></div><button class="secondary-button" data-action="export">⇩ تصدير النتائج</button></div>${renderTable(entries.slice((state.page - 1) * state.pageSize, state.page * state.pageSize), true)}${renderPagination(entries.length)}</section>`;
}

function renderRecords() {
  const entries = getFilteredEntries();
  const shown = entries.slice((state.page - 1) * state.pageSize, state.page * state.pageSize);
  return `<section class="hero"><div><p class="eyebrow">الأرشيف التشغيلي</p><h2>السجل الكامل</h2><p>جميع السجلات محفوظة محليًا ويمكن فتح أي معاملة لمراجعة جميع الحقول أو تعديلها.</p></div><div class="hero-note"><span>الصفحة الحالية</span><strong>${formatNumber(shown.length)} من ${formatNumber(entries.length)}</strong><span style="margin:7px 0 0">${formatNumber(state.records.length)} إجمالي محلي</span></div></section>${renderFilters(true)}<section class="panel table-panel"><div class="table-toolbar"><div><h3>سجل المعاملات</h3><span>مرتّب تصاعديًا حسب رقم الحركة</span></div><button class="secondary-button" data-action="export">⇩ تصدير النتائج</button></div>${renderTable(shown, true)}${renderPagination(entries.length)}</section>`;
}

function renderTable(entries, withIndex) {
  if (!entries.length) return '<div class="empty">لا توجد سجلات مطابقة. جرّب تغيير كلمة البحث أو مسح التصفية.</div>';
  return `<div class="table-wrap"><table><thead><tr><th>رقم الحركة</th><th>رقم القيد</th><th>الفرع</th><th>نوع المعاملة</th><th>اسم الجهة</th><th>الموضوع</th><th>الحالة</th><th>التأخير</th></tr></thead><tbody>${entries.map(({ record, index }) => `<tr class="data-row" data-index="${index}"><td class="td-strong">${escapeHtml(valueOf(record, 'movement') || '—')}</td><td>${escapeHtml(valueOf(record, 'registration') || '—')}</td><td>${escapeHtml(valueOf(record, 'branch') || '—')}</td><td>${escapeHtml(valueOf(record, 'type') || '—')}</td><td>${escapeHtml(valueOf(record, 'department') || '—')}</td><td class="subject-cell" title="${escapeHtml(valueOf(record, 'subject'))}">${escapeHtml(valueOf(record, 'subject') || '—')}</td><td><span class="badge ${isDone(record) ? 'done' : 'pending'}">${displayStatus(record)}</span></td><td><span class="badge ${isLate(record) ? 'late' : 'neutral'}">${isLate(record) ? 'متأخرة' : 'ضمن المدة'}</span></td></tr>`).join('')}</tbody></table></div>`;
}

function renderPagination(total) {
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  if (state.page > pages) state.page = pages;
  return `<div class="pagination"><button data-page="${state.page - 1}" ${state.page <= 1 ? 'disabled' : ''}>السابق</button><span>صفحة ${formatNumber(state.page)} من ${formatNumber(pages)}</span><button data-page="${state.page + 1}" ${state.page >= pages ? 'disabled' : ''}>التالي</button></div>`;
}

function renderSettings() {
  const headerKeys = getHeaderKeys();
  return `<section class="hero"><div><p class="eyebrow">حفظ محلي ومرونة في البيانات</p><h2>إعدادات البيانات</h2><p>إدارة النسخة المحلية، استيراد ملف جديد، أو إعادة تحميل بيانات البداية المرفقة مع التطبيق.</p></div><div class="hero-note"><span>حالة التخزين</span><strong>محفوظ على الجهاز</strong><span style="margin:7px 0 0">${formatNumber(state.records.length)} سجل · ${formatNumber(headerKeys.length)} عمود</span></div></section><section class="settings-grid"><article class="panel settings-card"><h3>المصدر الحالي</h3><p>يستخدم التطبيق نسخة البيانات المحلية على هذا الجهاز. عند استيراد Excel جديد يتم استبدال النسخة الحالية محليًا فقط، ولا يتم رفعها إلى الإنترنت.</p><span class="source-chip">${escapeHtml(state.source || 'نسخة محلية')}</span><div class="settings-actions"><button class="primary-button" data-action="import">＋ استيراد Excel</button><button class="secondary-button" data-action="export">⇩ تصدير كل البيانات</button></div></article><article class="panel settings-card"><h3>الحقول المكتشفة</h3><p>تم الاحتفاظ بكل عناوين الأعمدة كما وصلت من ملف Excel، بما في ذلك العمود غير المسمى في المصدر.</p><div class="record-grid" style="margin-top:14px">${headerKeys.map((key) => `<div class="detail-card"><p>${escapeHtml(key)}</p></div>`).join('')}</div></article><article class="panel settings-card"><h3>إعادة بيانات البداية</h3><p>إذا رغبت في التراجع عن الاستيراد الحالي والعودة إلى ملف أكتوبر المرفق، استخدم الزر التالي. سيتم حذف النسخة المحلية الحالية من هذا الجهاز.</p><div class="settings-actions"><button class="secondary-button danger" data-action="reset">إعادة بيانات البداية</button></div></article><article class="panel settings-card"><h3>العمل دون اتصال</h3><p>التطبيق مجهز بملف Manifest وService Worker. بعد فتحه مرة واحدة، يمكن تثبيته من قائمة المتصفح واستخدام البحث والبيانات دون اتصال.</p><div class="source-chip"><span class="offline-dot"></span> جاهز للاستخدام المحلي</div></article></section>`;
}

function renderModal() {
  const root = document.querySelector('#modal-root');
  if (state.selectedIndex === null || !state.records[state.selectedIndex]) { root.innerHTML = ''; return; }
  const record = state.records[state.selectedIndex];
  root.innerHTML = state.editMode ? renderEditModal(record, state.selectedIndex) : renderDetailModal(record, state.selectedIndex);
}

function renderDetailModal(record, index) {
  const detailFields = FIELD_DEFS;
  return `<div class="modal-backdrop" data-action="close-modal"><div class="modal" role="dialog" aria-modal="true" data-modal-content><div class="modal-head"><div><h2>معاملة رقم ${escapeHtml(valueOf(record, 'registration') || '—')}</h2><p>رقم الحركة ${escapeHtml(valueOf(record, 'movement') || '—')} · ${escapeHtml(valueOf(record, 'department') || 'الجهة غير محددة')}</p></div><button class="modal-close" data-action="close-modal">×</button></div><div class="modal-body"><div class="record-grid">${detailFields.map((field) => `<div class="detail-card ${['subject', 'requiredAction', 'takenAction'].includes(field.id) ? 'full-span' : ''}"><label>${field.label}</label><p>${escapeHtml(field.id.includes('Date') ? formatDate(valueOf(record, field.id)) : valueOf(record, field.id) || '—')}</p></div>`).join('')}</div><div class="timeline"><div class="timeline-item"><div class="timeline-dot">1</div><div><strong>الإرسال / الاستلام</strong><span>${formatDate(valueOf(record, 'sentDate'))} · المدة المحددة ${escapeHtml(valueOf(record, 'duration') || '—')} ${escapeHtml(valueOf(record, 'durationType') || '')}</span></div></div><div class="timeline-item"><div class="timeline-dot">2</div><div><strong>${escapeHtml(valueOf(record, 'department') || 'الجهة الحالية')}</strong><span>رقم الحركة ${escapeHtml(valueOf(record, 'movement') || '—')} · الإجراء المطلوب: ${escapeHtml(valueOf(record, 'requiredAction') || '—')}</span></div></div><div class="timeline-item"><div class="timeline-dot">3</div><div><strong>${displayStatus(record)}</strong><span>الإجراء المتخذ: ${escapeHtml(valueOf(record, 'takenAction') || 'لم يسجل بعد')} · تاريخ الانتهاء: ${formatDate(valueOf(record, 'endDate'))}</span></div></div></div></div><div class="modal-foot"><button class="primary-button" data-action="edit-record" data-index="${index}">✎ تعديل محلي</button><button class="secondary-button" data-action="close-modal">إغلاق</button></div></div></div>`;
}

function renderEditModal(record, index) {
  return `<div class="modal-backdrop" data-action="close-modal"><div class="modal" role="dialog" aria-modal="true" data-modal-content><div class="modal-head"><div><h2>تعديل السجل محليًا</h2><p>لن يتم إرسال أي تعديل خارج هذا الجهاز.</p></div><button class="modal-close" data-action="close-modal">×</button></div><form id="edit-form"><div class="modal-body"><div class="form-grid">${FIELD_DEFS.map((field) => { const value = valueOf(record, field.id); const key = findKey(record, field.aliases) || field.label; const area = ['subject', 'requiredAction', 'takenAction'].includes(field.id); return `<div class="form-field ${area ? 'full' : ''}"><label for="field-${field.id}">${field.label}</label>${area ? `<textarea id="field-${field.id}" data-field-key="${escapeHtml(key)}">${escapeHtml(value)}</textarea>` : `<input id="field-${field.id}" data-field-key="${escapeHtml(key)}" value="${escapeHtml(value)}" />`}</div>`; }).join('')}</div></div><div class="modal-foot"><button class="primary-button" type="submit">حفظ التعديل</button><button class="secondary-button" type="button" data-action="cancel-edit">إلغاء</button></div></form></div></div>`;
}

function renderMain() {
  const content = document.querySelector('#content');
  if (!content) return;
  content.innerHTML = state.view === 'dashboard' ? renderDashboard() : state.view === 'search' ? renderSearch() : state.view === 'records' ? renderRecords() : renderSettings();
  renderModal();
}

function render({ preserveQueryFocus = false } = {}) {
  const previousQuery = state.query;
  renderShell();
  renderMain();
  if (preserveQueryFocus) {
    const input = document.querySelector('#query');
    if (input) { input.focus(); input.setSelectionRange(previousQuery.length, previousQuery.length); }
  }
}

function exportRecords(entries = state.records.map((record, index) => ({ record, index })), label = 'المعاملات') {
  const headers = getHeaderKeys();
  const rows = entries.map(({ record }) => headers.reduce((row, header) => { row[header] = record[header] ?? ''; return row; }, {}));
  const sheet = XLSX.utils.json_to_sheet(rows, { header: headers });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'المعاملات');
  XLSX.writeFile(workbook, `${label}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`تم تصدير ${formatNumber(rows.length)} سجل إلى Excel`);
}

function importWorkbook(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
      if (!matrix.length) throw new Error('empty');
      const headers = matrix[0].map((header, index) => clean(header) || `عمود ${index + 1}`);
      const records = matrix.slice(1).filter((row) => row.some((value) => clean(value))).map((row) => headers.reduce((record, header, index) => { record[header] = clean(row[index]); return record; }, {}));
      if (!records.length) throw new Error('empty');
      state.records = records;
      state.headers = headers;
      state.source = file.name;
      state.sheet = sheetName;
      state.page = 1;
      state.query = '';
      state.branch = 'all';
      state.status = 'all';
      state.importance = 'all';
      persist();
      showToast(`تم استيراد ${formatNumber(records.length)} سجل من ${file.name}`);
    } catch (error) {
      showToast('تعذر قراءة الملف. تأكد من أنه Excel أو CSV صالح.', 'warn');
    }
  };
  reader.readAsArrayBuffer(file);
}

function resetToBundled() {
  if (!window.confirm('سيتم حذف النسخة المحلية الحالية وإعادة بيانات البداية. هل تريد المتابعة؟')) return;
  localStorage.removeItem(STORAGE_KEY);
  window.location.reload();
}

function updateFilter(id, value) {
  if (id === 'query') state.query = value;
  if (id === 'search-field') state.searchField = value;
  if (id === 'branch-filter') state.branch = value;
  if (id === 'status-filter') state.status = value;
  if (id === 'importance-filter') state.importance = value;
  state.page = 1;
}

app.addEventListener('click', (event) => {
  const viewButton = event.target.closest('[data-view]');
  if (viewButton) { state.view = viewButton.dataset.view; state.page = 1; render(); return; }
  const actionButton = event.target.closest('[data-action]');
  if (!actionButton) return;
  const action = actionButton.dataset.action;
  if (action === 'import') { document.querySelector('#file-input')?.click(); return; }
  if (action === 'export') { exportRecords(getFilteredEntries(), state.view === 'dashboard' ? 'ملخص-المعاملات' : 'نتائج-المعاملات'); return; }
  if (action === 'clear-filters') { state.query = ''; state.searchField = 'all'; state.branch = 'all'; state.status = 'all'; state.importance = 'all'; state.page = 1; render(); return; }
  if (action === 'reset') { resetToBundled(); return; }
  if (action === 'close-modal') { if (event.target.closest('[data-modal-content]') && !event.target.closest('.modal-close') && !event.target.closest('[data-action="close-modal"]')) return; state.selectedIndex = null; state.editMode = false; renderModal(); return; }
  if (action === 'edit-record') { state.selectedIndex = Number(actionButton.dataset.index); state.editMode = true; renderModal(); return; }
  if (action === 'cancel-edit') { state.editMode = false; renderModal(); return; }
  const row = event.target.closest('[data-index]');
  if (row && row.classList.contains('data-row')) { state.selectedIndex = Number(row.dataset.index); state.editMode = false; renderModal(); }
});

document.addEventListener('change', (event) => {
  if (event.target.id === 'file-input' && event.target.files?.[0]) { importWorkbook(event.target.files[0]); event.target.value = ''; return; }
  if (['search-field', 'branch-filter', 'status-filter', 'importance-filter'].includes(event.target.id)) { updateFilter(event.target.id, event.target.value); render(); }
});

document.addEventListener('input', (event) => {
  if (event.target.id === 'query') { updateFilter('query', event.target.value); render({ preserveQueryFocus: true }); }
});

document.addEventListener('click', (event) => {
  const pageButton = event.target.closest('[data-page]');
  if (pageButton && !pageButton.disabled) { state.page = Number(pageButton.dataset.page); render(); }
});

document.addEventListener('submit', (event) => {
  if (event.target.id !== 'edit-form') return;
  event.preventDefault();
  const record = state.records[state.selectedIndex];
  event.target.querySelectorAll('[data-field-key]').forEach((input) => { record[input.dataset.fieldKey] = input.value; });
  persist();
  state.editMode = false;
  showToast('تم حفظ التعديل محليًا');
});

document.addEventListener('keydown', (event) => { if (event.key === 'Escape' && state.selectedIndex !== null) { state.selectedIndex = null; state.editMode = false; renderModal(); } });

(async function boot() {
  await loadData();
  registerOffline();
  render();
})();
