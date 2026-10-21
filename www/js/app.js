/**
 * Outlook / Mail 快速取件 - 前端逻辑
 * 
 * 功能：
 * 1. 批量导入邮箱（localStorage 存储）
 * 2. 双协议并行取件（IMAP + Graph）
 * 3. 邮件展示与搜索
 */

// ==================== 全局状态 ====================
const STORAGE_KEY = 'outlook_accounts';
const MAIL_STORAGE_KEY = 'mail_accounts';
const APP_VERSION = '1.5.2-import';
const EXPORT_SEPARATOR = '----';
const SECURITY_RETRY_LIMIT = 1;
const TOKEN_REFRESH_COOLDOWN_MS = 10 * 60 * 1000;
const TOKEN_REFRESH_CONCURRENCY = 2;

// 追踪新导入的账号ID，用于高亮动画
let _newlyAddedIds = new Set();
let _apiSecuritySession = null;
let _mailOperation = null;
const _refreshingAccountIds = new Set();
let _tokenExportDialog = null;
let _activeProvider = 'outlook';
let _fetchSequence = 0;
const _providerStates = {
  outlook: createProviderState(),
  mail: createProviderState(),
};

function createProviderState() {
  return { selectedIds: new Set(), fields: {}, emails: [], resultOptions: {}, hasFetched: false, errors: [], status: { state: 'ready', text: '就绪' }, refreshReport: '' };
}

function accountStorageKey(provider = _activeProvider) {
  return provider === 'mail' ? MAIL_STORAGE_KEY : STORAGE_KEY;
}

function saveProviderView() {
  const state = _providerStates[_activeProvider];
  state.selectedIds = new Set(Array.from(document.querySelectorAll('.account-check:checked'), cb => cb.dataset.id));
  for (const id of ['accountSearch', 'searchKeyword', 'searchSender', 'fetchLimit', 'importTextarea']) {
    state.fields[id] = document.getElementById(id)?.value || '';
  }
  state.refreshReport = document.getElementById('tokenRefreshReport')?.innerHTML || '';
}

function applyProviderUI() {
  const isMail = _activeProvider === 'mail';
  const state = _providerStates[_activeProvider];
  document.body.dataset.provider = _activeProvider;
  document.title = `${isMail ? 'Mail' : 'Outlook'} 快速取件 | 邮箱快速取件`;
  document.getElementById('tokenRefreshControls').hidden = isMail;
  document.getElementById('mailSessionHint').hidden = !isMail;
  document.getElementById('protocolToggles').hidden = isMail;
  document.getElementById('mailFolderHint').hidden = !isMail;
  const report = document.getElementById('tokenRefreshReport');
  report.innerHTML = state.refreshReport;
  report.hidden = isMail || !state.refreshReport;
  document.getElementById('importFormat').textContent = isMail ? '邮箱:密码 / 邮箱---密码 / 邮箱----密码' : '账号----密码----clientid----刷新令牌';
  document.getElementById('importHint').textContent = isMail
    ? '支持中英文冒号、2–4 个短横线、竖线、Tab 或空格，可混合粘贴。密码以短横线开头时建议用冒号分隔。'
    : '分隔符支持 1~4 个短横线（- 到 ----）';
  document.getElementById('importTextarea').placeholder = isMail
    ? 'user@mail.com:password123\nuser2@mail.com---password456\nuser3@mail.com----password789'
    : 'user@outlook.com----password123----client-id-here----refresh-token-here\nuser2@outlook.com--pass--clientid--token';
  const limit = document.getElementById('fetchLimit');
  limit.innerHTML = (isMail ? [1, 3, 5, 10, 20] : [1, 3, 5, 10, 20, 30])
    .map(value => `<option value="${value}">${value} 封</option>`).join('');
  for (const id of ['accountSearch', 'searchKeyword', 'searchSender', 'fetchLimit', 'importTextarea']) {
    document.getElementById(id).value = state.fields[id] ?? (id === 'fetchLimit' ? '5' : '');
  }
  document.querySelectorAll('.provider-tab').forEach(tab => {
    const selected = tab.dataset.provider === _activeProvider;
    tab.classList.toggle('active', selected);
    tab.setAttribute('aria-selected', String(selected));
    tab.tabIndex = selected ? 0 : -1;
  });
  document.getElementById('mailWorkspace').setAttribute('aria-labelledby', isMail ? 'providerMail' : 'providerOutlook');
  renderAccountList(null, state.selectedIds);
  if (state.hasFetched) renderEmailList(state.emails, state.resultOptions);
  else {
    document.getElementById('resultsHeader').style.display = 'none';
    document.getElementById('emailList').innerHTML = `<div class="empty-state" id="emailEmpty"><svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="0.8" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 7L12 13L2 7"/></svg><p>在左侧导入邮箱，然后点击取件</p><p class="text-muted">${isMail ? '收件箱与垃圾箱一并查看，登录状态自动复用' : '支持 IMAP OAuth2 + Graph API 双协议并行取件'}</p></div>`;
    window._currentEmails = [];
  }
  renderFetchIssues(state.errors);
  showProgress(false);
  setStatus(state.status.state, state.status.text);
}

function switchProvider(provider) {
  if (!Object.hasOwn(_providerStates, provider) || provider === _activeProvider) return false;
  if (_mailOperation || document.getElementById('btnConfirmImport')?.disabled) {
    showToast('正在处理邮箱，请完成后再切换', 'info');
    return false;
  }
  saveProviderView();
  if (_tokenExportDialog) closeTokenRefreshExportDialog();
  document.getElementById('importModal')?.classList.remove('active');
  document.getElementById('emailDetailModal')?.classList.remove('active');
  _activeProvider = provider;
  try { localStorage.setItem('mail_active_provider', provider); } catch { /* 切换不依赖浏览器持久化。 */ }
  applyProviderUI();
  return true;
}

function initializeProviderSwitcher() {
  try {
    if (localStorage.getItem('mail_active_provider') === 'mail') _activeProvider = 'mail';
  } catch { /* 默认打开 Outlook。 */ }
  document.querySelectorAll('.provider-tab').forEach(tab => {
    tab.addEventListener('click', () => switchProvider(tab.dataset.provider));
    tab.addEventListener('keydown', event => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const next = event.key === 'Home' ? 'outlook' : event.key === 'End' ? 'mail' : _activeProvider === 'mail' ? 'outlook' : 'mail';
      if (switchProvider(next)) document.getElementById(next === 'mail' ? 'providerMail' : 'providerOutlook').focus();
    });
  });
  applyProviderUI();
}

/**
 * 获取所有已存储的账号
 */
function getAccounts(provider = _activeProvider) {
  try {
    const accounts = JSON.parse(localStorage.getItem(accountStorageKey(provider)) || '[]');
    return normalizeAccountOrder(accounts);
  } catch {
    return [];
  }
}

/**
 * 保存账号列表
 */
function saveAccounts(accounts, provider = _activeProvider) {
  localStorage.setItem(accountStorageKey(provider), JSON.stringify(normalizeAccountOrder(accounts)));
}

function normalizeAccountOrder(accounts) {
  if (!Array.isArray(accounts)) return [];
  return accounts.map((account, index) => ({
    ...account,
    importIndex: Number.isFinite(account.importIndex) ? account.importIndex : index,
  }));
}

function getAccountSearchKeyword() {
  const input = document.getElementById('accountSearch');
  return (input?.value || '').trim().toLowerCase();
}

// ==================== 解析导入 ====================

/**
 * 解析导入文本
 * 格式：账号----密码----clientid----刷新令牌
 * 分隔符支持 1-4 个短横线
 * @param {string} text - 导入文本
 * @returns {Array} - 解析后的账号列表
 */
function parseImportText(text, provider = _activeProvider) {
  if (provider === 'mail') return parseMailImportText(text);
  const lines = text.trim().split('\n');
  const accounts = [];
  const errors = [];

  lines.forEach((line, index) => {
    const trimmed = line.trim();
    if (!trimmed) return; // 跳过空行

    // 使用正则匹配 1-4 个短横线作为分隔符
    // 需要小心不要匹配 UUID 中的单个短横线
    // 策略：先尝试4个分隔，再3个，再2个，再1个
    let parts = null;

    // 尝试从最多的短横线开始分割
    for (let dashCount = 4; dashCount >= 1; dashCount--) {
      const sep = '-'.repeat(dashCount);
      // 构建正则：匹配恰好 dashCount 个短横线（前后不能是短横线）
      // 但为了简化，我们使用 split 然后验证
      const testParts = trimmed.split(sep);
      
      // 有效分割应该得到恰好 4 个部分
      if (testParts.length === 4) {
        // 验证每个部分都不为空（去除首尾空格后）
        const cleaned = testParts.map(p => p.trim());
        if (cleaned.every(p => p.length > 0)) {
          parts = cleaned;
          break;
        }
      }

      // 如果分割出超过4个部分，尝试只分割前3个分隔符
      if (testParts.length > 4) {
        const cleaned = [];
        let remaining = trimmed;
        for (let i = 0; i < 3; i++) {
          const idx = remaining.indexOf(sep);
          if (idx === -1) break;
          cleaned.push(remaining.substring(0, idx).trim());
          remaining = remaining.substring(idx + sep.length);
        }
        if (cleaned.length === 3 && remaining.trim().length > 0) {
          cleaned.push(remaining.trim());
          parts = cleaned;
          break;
        }
      }
    }

    if (!parts || parts.length !== 4) {
      errors.push(`第 ${index + 1} 行格式错误: ${trimmed.substring(0, 40)}...`);
      return;
    }

    const [email, password, clientId, refreshToken] = parts;

    // 简单验证邮箱格式
    if (!email.includes('@')) {
      errors.push(`第 ${index + 1} 行邮箱格式无效: ${email}`);
      return;
    }

    accounts.push({
      id: generateId(),
      email,
      password,
      clientId,
      refreshToken,
      importIndex: index,
      addedAt: new Date().toISOString(),
    });
  });

  return { accounts, errors };
}

function parseMailImportText(text) {
  const accounts = [];
  const errors = [];
  const emailPattern = /^[a-z0-9.!#$%&'*+\/=?^_`{}~-]+@(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i;
  // 仅在合法邮箱末尾拆分一次，保留密码内的分隔符及首尾空格。
  String(text).split(/\r\n|\n|\r/).forEach((line, index) => {
    if (!line.trim()) return;
    let email = '';
    let password = '';
    for (const separator of line.matchAll(/----|---|--|[:：|]|[ \t]+/g)) {
      const candidate = line.slice(0, separator.index).trim();
      if (!emailPattern.test(candidate)) continue;
      const remainder = line.slice(separator.index + separator[0].length);
      // “邮箱 :密码”优先识别冒号；纯空白分隔时消耗整段空白。
      if (/^[ \t]/.test(separator[0]) && /^(?:----|---|--|[:：|])/.test(remainder)) continue;
      email = candidate;
      password = remainder;
      break;
    }
    if (!email || !password.length) {
      errors.push(`第 ${index + 1} 行格式错误，请使用 邮箱:密码、邮箱---密码 等支持格式`);
      return;
    }
    accounts.push({ id: generateId(), email, password, importIndex: index, addedAt: new Date().toISOString() });
  });
  return { accounts, errors };
}

/**
 * 生成唯一 ID
 */
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

// ==================== UI 渲染 ====================

/**
 * 渲染邮箱列表
 * @param {Set} highlightIds - 需要高亮的新导入账号ID
 */
function renderAccountList(highlightIds = null, restoredSelection = null) {
  // 保留当前勾选项，刷新进度、导入或其它标签页更新均可能重绘列表。
  const selectedIds = restoredSelection || new Set(Array.from(document.querySelectorAll('.account-check:checked'), cb => cb.dataset.id));
  const accounts = getAccounts();
  const searchKeyword = getAccountSearchKeyword();
  const visibleAccounts = searchKeyword
    ? accounts.filter(acc => (acc.email || '').toLowerCase().includes(searchKeyword))
    : accounts;
  const listEl = document.getElementById('accountList');
  const countEl = document.getElementById('accountCount');

  // 更新计数，带动画
  const oldCount = parseInt(countEl.textContent) || 0;
  const newCount = accounts.length;
  countEl.textContent = searchKeyword ? `${visibleAccounts.length}/${newCount}` : newCount;
  if (newCount !== oldCount) {
    countEl.classList.add('badge-pulse');
    setTimeout(() => countEl.classList.remove('badge-pulse'), 600);
  }

  if (accounts.length === 0) {
    listEl.innerHTML = `
      <div class="empty-state">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
          <rect x="2" y="4" width="20" height="16" rx="2"/>
          <path d="M22 7L12 13L2 7"/>
        </svg>
        <p>暂无邮箱</p>
        <p class="text-muted">点击上方按钮导入</p>
      </div>
    `;
    // 隐藏批量删除按钮
    updateBulkDeleteButton(0);
    return;
  }

  if (visibleAccounts.length === 0) {
    listEl.innerHTML = `
      <div class="empty-state">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
          <circle cx="11" cy="11" r="8"/>
          <path d="M21 21l-4.35-4.35"/>
        </svg>
        <p>未找到邮箱</p>
        <p class="text-muted">换个关键词试试</p>
      </div>
    `;
    updateBulkDeleteButton(0);
    return;
  }

  // 全选容器 + 批量删除按钮
  let html = `
    <div class="select-all-wrapper">
      <input type="checkbox" class="account-checkbox" id="selectAll" />
      <label for="selectAll" style="cursor:pointer;">全选</label>
      <button class="btn btn-ghost btn-small btn-danger-ghost" id="btnDeleteSelected" style="display:none;margin-left:auto;">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M3 6h18M8 6V4h8v2M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/>
        </svg>
        删除选中
      </button>
    </div>
  `;

  visibleAccounts.forEach((acc, index) => {
    const isNew = highlightIds && highlightIds.has(acc.id);
    html += `
      <div class="account-item ${isNew ? 'account-item-new' : ''} ${selectedIds.has(acc.id) ? 'selected' : ''}" data-id="${acc.id}" style="animation-delay: ${isNew ? index * 0.05 : 0}s">
        <input type="checkbox" class="account-checkbox account-check" data-id="${acc.id}" ${selectedIds.has(acc.id) ? 'checked' : ''} />
        <div class="account-info">
          <span class="account-email" title="${escapeAttr(acc.email)}">${escapeHtml(acc.email)}</span>
          ${_activeProvider === 'mail' ? renderMailSessionStatus(acc) : renderTokenRefreshStatus(acc)}
        </div>
        <button class="account-copy" data-email="${escapeAttr(acc.email)}" title="复制邮箱">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="9" y="9" width="13" height="13" rx="2"/>
            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
          </svg>
        </button>
        <button class="account-delete" onclick="event.stopPropagation(); deleteAccount('${acc.id}')" title="删除">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
          </svg>
        </button>
      </div>
    `;
  });

  listEl.innerHTML = html;

  // 清除新导入高亮（3秒后移除动画类）
  if (highlightIds && highlightIds.size > 0) {
    setTimeout(() => {
      document.querySelectorAll('.account-item-new').forEach(el => {
        el.classList.remove('account-item-new');
      });
    }, 3000);
  }

  // 全选事件
  const selectAllEl = document.getElementById('selectAll');
  if (selectAllEl) {
    selectAllEl.addEventListener('change', (e) => {
      document.querySelectorAll('.account-check').forEach(cb => {
        cb.checked = e.target.checked;
        cb.closest('.account-item')?.classList.toggle('selected', e.target.checked);
      });
      updateBulkDeleteButton();
    });
  }

  // 单个选中事件（同时支持点击行来切换选中）
  document.querySelectorAll('.account-item').forEach(item => {
    // 点击行切换选中
    item.addEventListener('click', (e) => {
      // 如果点击的是 checkbox 或删除按钮，不处理
      if (e.target.closest('.account-checkbox') || e.target.closest('.account-delete') || e.target.closest('.account-copy') || e.target.closest('.token-status-error')) return;
      const cb = item.querySelector('.account-check');
      if (cb) {
        cb.checked = !cb.checked;
        item.classList.toggle('selected', cb.checked);
        updateSelectAllState();
        updateBulkDeleteButton();
      }
    });
  });

  document.querySelectorAll('.account-check').forEach(cb => {
    cb.addEventListener('change', (e) => {
      e.target.closest('.account-item')?.classList.toggle('selected', e.target.checked);
      updateSelectAllState();
      updateBulkDeleteButton();
    });
  });

  // 批量删除按钮事件
  const btnDeleteSelected = document.getElementById('btnDeleteSelected');
  if (btnDeleteSelected) {
    btnDeleteSelected.addEventListener('click', deleteSelectedAccounts);
  }

  document.querySelectorAll('.account-copy').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      copyText(btn.dataset.email || '', '邮箱已复制');
    });
  });
  document.querySelectorAll('.token-status-error').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const account = getAccounts().find(acc => acc.id === btn.dataset.id);
      if (account?.tokenRefreshError) showToast(account.tokenRefreshError, 'error', 8000);
    });
  });
  updateSelectAllState();
  updateBulkDeleteButton();
}

/**
 * 更新全选复选框状态
 */
function updateSelectAllState() {
  const selectAllEl = document.getElementById('selectAll');
  if (!selectAllEl) return;
  const all = document.querySelectorAll('.account-check');
  const checked = document.querySelectorAll('.account-check:checked');
  selectAllEl.checked = all.length > 0 && all.length === checked.length;
  selectAllEl.indeterminate = checked.length > 0 && checked.length < all.length;
}

/**
 * 更新批量删除按钮显隐
 */
function updateBulkDeleteButton(count) {
  const btn = document.getElementById('btnDeleteSelected');
  if (!btn) return;
  if (count === undefined) {
    count = document.querySelectorAll('.account-check:checked').length;
  }
  if (count > 0) {
    btn.style.display = 'inline-flex';
    btn.textContent = '';
    btn.innerHTML = `
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M3 6h18M8 6V4h8v2M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/>
      </svg>
      删除选中 (${count})
    `;
  } else {
    btn.style.display = 'none';
  }
}

/**
 * 删除选中的账号
 */
function deleteSelectedAccounts() {
  const provider = _activeProvider;
  const checkedIds = [];
  document.querySelectorAll('.account-check:checked').forEach(cb => {
    checkedIds.push(cb.dataset.id);
  });

  if (checkedIds.length === 0) {
    showToast('请先选择要删除的邮箱', 'warning');
    return;
  }

  if (!confirm(`确定要删除选中的 ${checkedIds.length} 个邮箱吗？`)) return;

  // 播放移除动画
  checkedIds.forEach(id => {
    const item = document.querySelector(`.account-item[data-id="${id}"]`);
    if (item) {
      item.classList.add('account-item-removing');
    }
  });

  // 等动画结束后再真正删除
  setTimeout(() => {
    const accounts = getAccounts(provider).filter(a => !checkedIds.includes(a.id));
    saveAccounts(accounts, provider);
    if (_activeProvider === provider) {
      renderAccountList();
      showToast(`已删除 ${checkedIds.length} 个邮箱`, 'success');
    }
  }, 300);
}

/**
 * 删除单个账号（带动画）
 */
function deleteAccount(id) {
  const provider = _activeProvider;
  const item = document.querySelector(`.account-item[data-id="${id}"]`);
  if (item) {
    item.classList.add('account-item-removing');
  }

  // 等动画结束后再真正删除
  setTimeout(() => {
    const accounts = getAccounts(provider).filter(a => a.id !== id);
    saveAccounts(accounts, provider);
    if (_activeProvider === provider) {
      renderAccountList();
      showToast('已删除邮箱', 'info');
    }
  }, 300);
}

/**
 * 获取选中的账号
 */
function getSelectedAccounts() {
  const checkedIds = [];
  document.querySelectorAll('.account-check:checked').forEach(cb => {
    checkedIds.push(cb.dataset.id);
  });
  const accounts = getAccounts();
  return accounts.filter(a => checkedIds.includes(a.id));
}

/**
 * 导出账号，优先导出已选邮箱；未选择时导出全部
 */
function exportAccounts() {
  const selected = getSelectedAccounts();
  const accounts = selected.length > 0 ? selected : getAccounts();

  if (accounts.length === 0) {
    showToast('没有可导出的邮箱', 'warning');
    return;
  }

  const content = downloadAccounts(accounts);
  copyText(content, selected.length > 0 ? `已导出并复制选中的 ${accounts.length} 个邮箱` : `已导出并复制全部 ${accounts.length} 个邮箱`, false);
}

function downloadAccounts(accounts, prefix = `${_activeProvider}-accounts`) {
  const content = accounts.map(account => formatAccountForExport(account)).join('\n');
  downloadTextFile(`${prefix}-${new Date().toISOString().slice(0, 10)}.txt`, content);
  return content;
}

function formatAccountForExport(account, provider = _activeProvider) {
  if (provider === 'mail') return [account.email || '', account.password || ''].join(EXPORT_SEPARATOR);
  return [
    account.email || '',
    account.password || '',
    account.clientId || '',
    account.refreshToken || '',
  ].join(EXPORT_SEPARATOR);
}

function downloadTextFile(filename, content) {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

// ==================== 令牌刷新 ====================

function renderMailSessionStatus(account) {
  const expiresAt = Date.parse(account.sessionExpiresAt || '');
  if (!Number.isFinite(expiresAt)) return '';
  if (expiresAt <= Date.now()) return '<span class="token-status">下次取件自动登录</span>';
  return '<span class="token-status token-status-success" title="取件时自动复用有效登录，失效时自动登录">✓ 登录状态可复用</span>';
}

function tokenRefreshCooldownRemaining(account, now = Date.now()) {
  const refreshedAt = Date.parse(account.tokenRefreshedAt || '');
  return Number.isFinite(refreshedAt) ? Math.max(0, refreshedAt + TOKEN_REFRESH_COOLDOWN_MS - now) : 0;
}

function renderTokenRefreshStatus(account) {
  if (_refreshingAccountIds.has(account.id)) {
    return '<span class="token-status token-status-loading">刷新中…</span>';
  }
  if (account.tokenRefreshError) {
    return `<button type="button" class="token-status token-status-error" data-id="${escapeAttr(account.id)}" title="${escapeAttr(account.tokenRefreshError)}">刷新失败 · 查看原因</button>`;
  }
  const refreshedAt = Date.parse(account.tokenRefreshedAt || '');
  if (!Number.isFinite(refreshedAt)) return '';
  const remaining = tokenRefreshCooldownRemaining(account);
  const time = new Date(refreshedAt).toLocaleString('zh-CN', {
    month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false,
  });
  const cooldown = remaining > 0 ? ` · 冷却 ${Math.ceil(remaining / 60000)} 分钟` : '';
  return `<span class="token-status token-status-success" title="${escapeAttr(new Date(refreshedAt).toLocaleString('zh-CN'))}">✓ 已刷新 ${escapeHtml(time)}${cooldown}</span>`;
}

function updateMailOperationButtons() {
  for (const id of ['btnFetchSelected', 'btnFetchAll', 'btnRefreshSelected', 'btnRefreshAll', 'providerOutlook', 'providerMail']) {
    const button = document.getElementById(id);
    if (button) button.disabled = Boolean(_mailOperation);
  }
}

async function runMailOperation(kind, operation) {
  if (_mailOperation) {
    showToast(_mailOperation === 'refresh' ? '令牌刷新中，请完成后再操作' : '取件中，请完成后再操作', 'warning');
    return;
  }
  _mailOperation = kind;
  updateMailOperationButtons();
  try {
    // 同源多标签页共用锁；不支持 Web Locks 的浏览器仍有本页互斥及持久化冷却。
    if (navigator.locks?.request) {
      return await navigator.locks.request(`${_activeProvider}-mail-operation`, { ifAvailable: true }, async lock => {
        if (!lock) {
          showToast('其它标签页正在取件或刷新令牌，请稍后再试', 'warning');
          return;
        }
        return operation();
      });
    }
    return await operation();
  } catch (error) {
    setStatus('error', kind === 'refresh' ? '令牌刷新失败' : '取件失败');
    showToast(simplifyErrorMessage(error.message || '操作失败，请重试'), 'error', 6000);
  } finally {
    _mailOperation = null;
    updateMailOperationButtons();
  }
}

function accountCredentialsMatch(current, original) {
  return current && current.id === original.id && current.email === original.email
    && current.clientId === original.clientId && current.refreshToken === original.refreshToken;
}

// 每次回包都重新读取列表，只更新仍使用此次请求凭据的账号，避免复活已删除账号或覆盖新导入数据。
function saveTokenRefreshResult(original, changes) {
  const latest = getAccounts();
  const index = latest.findIndex(account => account.id === original.id);
  if (index === -1 || !accountCredentialsMatch(latest[index], original)) return null;
  latest[index] = { ...latest[index], ...changes };
  saveAccounts(latest);
  return latest[index];
}

function renderTokenRefreshReport(summary, running = false) {
  const report = document.getElementById('tokenRefreshReport');
  if (!report) return;
  report.hidden = false;
  const changed = summary.changed ? ` · 列表变更跳过 ${summary.changed}` : '';
  report.innerHTML = `
    <p class="token-refresh-report-title">${running ? '令牌刷新中' : '令牌刷新完成'}${running ? `（${summary.completed}/${summary.total}）` : ''}</p>
    <p>成功 ${summary.success} · 失败 ${summary.failed} · 冷却跳过 ${summary.cooldown}${changed}</p>
    ${summary.errors.length ? `<details class="token-refresh-errors" open><summary>查看失败原因（${summary.errors.length}）</summary><ul>${summary.errors.map(item => `<li><strong>${escapeHtml(item.email)}</strong>：${escapeHtml(item.message)}</li>`).join('')}</ul></details>` : ''}
  `;
}

function getTokenRefreshExportAccounts(successfulTokens) {
  return getAccounts().filter(account => successfulTokens.has(account.id) && successfulTokens.get(account.id) === account.refreshToken);
}

function closeTokenRefreshExportDialog() {
  const previousFocus = _tokenExportDialog?.previousFocus;
  _tokenExportDialog = null;
  const modal = document.getElementById('tokenExportModal');
  modal.classList.remove('active');
  modal.setAttribute('aria-hidden', 'true');
  if (previousFocus?.isConnected && !previousFocus.disabled) previousFocus.focus();
}

function showTokenRefreshExportDialog(summary, successfulTokens, previousFocus = document.activeElement) {
  const exportable = getTokenRefreshExportAccounts(successfulTokens);
  if (!exportable.length) return;
  if (_tokenExportDialog) closeTokenRefreshExportDialog();
  _tokenExportDialog = { successfulTokens: new Map(successfulTokens), previousFocus };
  document.getElementById('tokenExportSuccessCount').textContent = summary.success;
  document.getElementById('tokenExportFailedCount').textContent = summary.failed;
  document.getElementById('tokenExportCooldownCount').textContent = summary.cooldown;
  const changed = document.getElementById('tokenExportChanged');
  changed.hidden = !summary.changed;
  changed.textContent = summary.changed ? `另有 ${summary.changed} 个账号因列表变更已跳过` : '';
  document.getElementById('tokenExportNote').textContent = `可导出本次成功的 ${exportable.length} 个账号，方便备份或迁移。`;
  document.getElementById('tokenExportButtonText').textContent = `导出成功账号（${exportable.length}）`;
  const modal = document.getElementById('tokenExportModal');
  modal.classList.add('active');
  modal.setAttribute('aria-hidden', 'false');
  document.getElementById('btnConfirmTokenExport').focus();
}

function exportRefreshedAccountsFromDialog() {
  if (!_tokenExportDialog) return;
  // 弹窗打开后列表仍可能在其它标签页变化；点击导出时再读取和核对一次。
  const exportable = getTokenRefreshExportAccounts(_tokenExportDialog.successfulTokens);
  if (!exportable.length) {
    closeTokenRefreshExportDialog();
    showToast('本次成功账号已被删除或更新，当前没有可导出的账号', 'info');
    return;
  }
  downloadAccounts(exportable, 'outlook-refreshed-accounts');
  closeTokenRefreshExportDialog();
  showToast(`已导出本次刷新成功的 ${exportable.length} 个邮箱`, 'success');
}

function initializeTokenRefreshExportDialog() {
  const modal = document.getElementById('tokenExportModal');
  document.getElementById('btnCloseTokenExport').addEventListener('click', closeTokenRefreshExportDialog);
  document.getElementById('btnCancelTokenExport').addEventListener('click', closeTokenRefreshExportDialog);
  document.getElementById('btnConfirmTokenExport').addEventListener('click', exportRefreshedAccountsFromDialog);
  modal.addEventListener('click', event => {
    if (event.target === modal) closeTokenRefreshExportDialog();
  });
  document.addEventListener('keydown', event => {
    if (!_tokenExportDialog) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      closeTokenRefreshExportDialog();
      return;
    }
    if (event.key !== 'Tab') return;
    const focusable = ['btnCloseTokenExport', 'btnCancelTokenExport', 'btnConfirmTokenExport'].map(id => document.getElementById(id));
    const index = focusable.indexOf(document.activeElement);
    if (index === -1 || (event.shiftKey && index === 0) || (!event.shiftKey && index === focusable.length - 1)) {
      event.preventDefault();
      focusable[event.shiftKey ? focusable.length - 1 : 0].focus();
    }
  });
}

async function refreshAccountTokens(accounts) {
  if (_activeProvider !== 'outlook') return;
  if (accounts.length === 0) {
    showToast('请先选择或导入要刷新令牌的邮箱', 'warning');
    return;
  }
  const focusBeforeRefresh = document.activeElement;
  let completedBatch = null;
  const result = await runMailOperation('refresh', async () => {
    const queue = [...new Map(accounts.map(account => [account.id, account])).values()];
    const summary = { total: queue.length, completed: 0, success: 0, failed: 0, cooldown: 0, changed: 0, errors: [] };
    const successfulTokens = new Map();
    let cursor = 0;
    setStatus('loading', '刷新令牌中…');
    renderTokenRefreshReport(summary, true);

    async function refreshNextAccount() {
      while (cursor < queue.length) {
        const queued = queue[cursor++];
        const account = getAccounts().find(item => item.id === queued.id);
        if (!accountCredentialsMatch(account, queued)) {
          summary.changed++;
          summary.completed++;
          renderTokenRefreshReport(summary, true);
          continue;
        }
        if (tokenRefreshCooldownRemaining(account) > 0) {
          summary.cooldown++;
          summary.completed++;
          renderTokenRefreshReport(summary, true);
          continue;
        }
        _refreshingAccountIds.add(account.id);
        renderAccountList();
        try {
          const response = await secureApiFetch('/api/refresh-token', {
            email: account.email, clientId: account.clientId, refreshToken: account.refreshToken,
          });
          const data = await parseApiResponse(response);
          if (!data.success) {
            throw new Error([data.reason || data.error || data.message || '令牌刷新失败', data.action].filter(Boolean).join('；'));
          }
          const refreshToken = data.refreshToken || data.newRefreshToken || account.refreshToken;
          if (typeof refreshToken !== 'string' || !refreshToken.trim()) throw new Error('刷新接口未返回可用令牌');
          // 不保存响应中的 accessToken；成功标志与最新 RT 一起写入浏览器账号列表。
          const saved = saveTokenRefreshResult(account, {
            refreshToken,
            tokenRefreshedAt: new Date().toISOString(),
            tokenRefreshError: null,
          });
          if (saved) {
            summary.success++;
            successfulTokens.set(saved.id, saved.refreshToken);
          } else {
            summary.changed++;
          }
        } catch (error) {
          let message = simplifyErrorMessage(error.message || '令牌刷新失败，请重试');
          // 若本地存储写入失败，仍完成其他账号并在汇总里显示原因。
          try {
            const saved = saveTokenRefreshResult(account, { tokenRefreshError: message });
            if (!saved) {
              summary.changed++;
              continue;
            }
          } catch {
            message = `无法保存刷新结果，请检查浏览器存储空间后重试。${message}`;
          }
          summary.failed++;
          summary.errors.push({ email: account.email, message });
        } finally {
          _refreshingAccountIds.delete(account.id);
          summary.completed++;
          renderAccountList();
          renderTokenRefreshReport(summary, true);
        }
      }
    }

    await Promise.all(Array.from({ length: Math.min(TOKEN_REFRESH_CONCURRENCY, queue.length) }, refreshNextAccount));
    renderTokenRefreshReport(summary);
    setStatus(summary.failed && !summary.success ? 'error' : 'ready', '令牌刷新完成');
    showToast(`令牌刷新完成：成功 ${summary.success}，失败 ${summary.failed}，冷却跳过 ${summary.cooldown}${summary.changed ? `，列表变更跳过 ${summary.changed}` : ''}`, summary.failed ? 'warning' : 'success', 6000);

    completedBatch = { summary, successfulTokens };
    return summary;
  });
  // 先释放取件/刷新锁，再展示非阻塞的导出选择。
  if (completedBatch) showTokenRefreshExportDialog(completedBatch.summary, completedBatch.successfulTokens, focusBeforeRefresh);
  return result;
}

// ==================== 邮件取件 ====================

/**
 * 双协议并行取件
 * @param {Array} accounts - 要取件的账号列表
 * @param {object} options - 搜索选项
 */
async function fetchEmails(accounts, options = {}) {
  if (accounts.length === 0) {
    showToast('请先选择邮箱', 'warning');
    return;
  }

  const provider = _activeProvider;
  const sequence = ++_fetchSequence;
  const useMail = provider === 'mail';
  const useImap = !useMail && document.getElementById('toggleImap').checked;
  const useGraph = !useMail && document.getElementById('toggleGraph').checked;

  if (!useMail && !useImap && !useGraph) {
    showToast('请至少选择一种协议', 'warning');
    return;
  }

  // 计算总步骤数（每个邮箱 = 协议数个步骤）
  const protocolCount = useMail ? 1 : (useImap ? 1 : 0) + (useGraph ? 1 : 0);
  const totalSteps = accounts.length * protocolCount;
  let completedSteps = 0;

  // 显示进度
  setStatus('loading', '取件中...');
  showProgress(true);
  updateProgress(0, useMail ? `准备取件 ${accounts.length} 个 Mail 邮箱...` : `准备取件 ${accounts.length} 个邮箱 (${protocolCount} 种协议)...`);

  // 先显示骨架屏占位
  showSkeletonCards(3);

  let visibleEmails = [];
  const failedAccountErrors = [];
  let successfulAccounts = 0;
  let failedAccounts = 0;
  renderFetchIssues([]);
  renderEmailList([], {
    accounts,
    completedAccounts: 0,
    isFetching: true,
  });

  for (let i = 0; i < accounts.length; i++) {
    const account = accounts[i];
    const emailShort = account.email.split('@')[0];
    const accountLabel = `[${i + 1}/${accounts.length}] ${emailShort}`;

    const promises = [];

    if (useMail) {
      promises.push((async () => {
        updateProgress(Math.round((completedSteps / totalSteps) * 95), `${accountLabel} — Mail 取件中...`);
        try {
          const result = await fetchViaMail(account, options);
          completedSteps++;
          updateProgress(Math.round((completedSteps / totalSteps) * 95), `${accountLabel} — Mail ✅ ${result.count || 0} 封`);
          return result;
        } catch (err) {
          completedSteps++;
          updateProgress(Math.round((completedSteps / totalSteps) * 95), `${accountLabel} — Mail ❌`);
          return { success: false, emails: [], protocol: 'mail', _error: normalizeFetchError(account.email, 'mail', err) };
        }
      })());
    }

    // 双协议并行请求，每个完成时更新进度
    if (useGraph) {
      promises.push(
        (async () => {
          updateProgress(
            Math.round((completedSteps / totalSteps) * 95),
            `${accountLabel} — Graph API 取件中...`
          );
          try {
            const result = await fetchViaGraph(account, options);
            completedSteps++;
            updateProgress(
              Math.round((completedSteps / totalSteps) * 95),
              `${accountLabel} — Graph ✅ ${result.count || 0} 封`
            );
            return result;
          } catch (err) {
            completedSteps++;
            updateProgress(
              Math.round((completedSteps / totalSteps) * 95),
              `${accountLabel} — Graph ❌`
            );
            return {
              success: false,
              emails: [],
              protocol: 'graph',
              _error: normalizeFetchError(account.email, 'graph', err),
            };
          }
        })()
      );
    }

    if (useImap) {
      promises.push(
        (async () => {
          updateProgress(
            Math.round((completedSteps / totalSteps) * 95),
            `${accountLabel} — IMAP 取件中...`
          );
          try {
            const result = await fetchViaIMAP(account, options);
            completedSteps++;
            updateProgress(
              Math.round((completedSteps / totalSteps) * 95),
              `${accountLabel} — IMAP ✅ ${result.count || 0} 封`
            );
            return result;
          } catch (err) {
            completedSteps++;
            updateProgress(
              Math.round((completedSteps / totalSteps) * 95),
              `${accountLabel} — IMAP ❌`
            );
            return {
              success: false,
              emails: [],
              protocol: 'imap',
              _error: normalizeFetchError(account.email, 'imap', err),
            };
          }
        })()
      );
    }

    const results = await Promise.all(promises);
    const accountErrors = results.map(result => result._error).filter(Boolean);
    const accountSucceeded = results.some(result => result.success === true);

    if (accountSucceeded) {
      successfulAccounts++;
    } else {
      failedAccounts++;
      failedAccountErrors.push(...accountErrors);
    }

    // 收集邮件并标记来源账号
    results.forEach(result => {
      if (result.emails && result.emails.length > 0) {
        result.emails.forEach(email => {
          email._account = account.email;
          email._accountIndex = i;
        });
        visibleEmails.push(...result.emails);
      }
    });

    visibleEmails = deduplicateEmails(visibleEmails);
    visibleEmails.sort((a, b) => {
      const accountDelta = (a._accountIndex ?? 0) - (b._accountIndex ?? 0);
      if (accountDelta !== 0) return accountDelta;
      return new Date(b.date) - new Date(a.date);
    });
    renderEmailList(visibleEmails, {
      accounts,
      completedAccounts: i + 1,
      isFetching: i + 1 < accounts.length,
    });
    renderFetchIssues(failedAccountErrors);
  }

  // 合并去重
  updateProgress(97, '去重合并中...');
  const merged = deduplicateEmails(visibleEmails);
  merged.sort((a, b) => {
    const accountDelta = (a._accountIndex ?? 0) - (b._accountIndex ?? 0);
    if (accountDelta !== 0) return accountDelta;
    return new Date(b.date) - new Date(a.date);
  });

  // 渲染结果
  updateProgress(100, '取件完成 ✅');
  renderEmailList(merged, {
    accounts,
    completedAccounts: accounts.length,
    isFetching: false,
  });
  renderFetchIssues(failedAccountErrors);

  const imapCount = merged.filter(e => e.protocol === 'imap').length;
  const graphCount = merged.filter(e => e.protocol === 'graph').length;

  // 延迟隐藏进度条让用户看到完成状态
  setTimeout(() => {
    if (_activeProvider === provider && _fetchSequence === sequence) showProgress(false);
  }, 2000);

  if (failedAccountErrors.length > 0) {
    setStatus(successfulAccounts > 0 ? 'ready' : 'error', `完成：${successfulAccounts} 成功 / ${failedAccounts} 失败`);
    showToast(`取件完成，有 ${failedAccounts} 个邮箱${useMail ? '取件失败' : '所有协议都失败'}，详情已整理在结果区`, successfulAccounts > 0 ? 'warning' : 'error', 5000);
  } else {
    setStatus('ready', '就绪');
  }

  if (successfulAccounts > 0) showToast(useMail ? `取件完成：共 ${merged.length} 封 Mail 邮件` : `取件完成：共 ${merged.length} 封 (IMAP: ${imapCount} / Graph: ${graphCount})`, 'success', 4000);
}

/**
 * 显示骨架屏卡片
 */
function showSkeletonCards(count) {
  const listEl = document.getElementById('emailList');
  let html = '';
  for (let i = 0; i < count; i++) {
    html += `<div class="skeleton skeleton-card"></div>`;
  }
  listEl.innerHTML = html;
}

/**
 * 通过 Graph API 取件
 */
async function fetchViaGraph(account, options) {
  const response = await secureApiFetch('/api/fetch-graph', {
    email: account.email,
    clientId: account.clientId,
    refreshToken: account.refreshToken,
    keyword: options.keyword || '',
    limit: Math.max(1, Number.parseInt(options.limit, 10) || 1),
    sender: options.sender || '',
    filters: options.filters || [],
  });

  const data = await parseApiResponse(response);
  if (!data.success) {
    const err = new Error(data.error || data.message || 'Graph API 取件失败');
    err.code = data.code;
    err.reason = data.reason;
    err.action = data.action;
    err.detail = data.detail;
    throw err;
  }
  return data;
}

/**
 * 通过 IMAP 取件
 */
async function fetchViaIMAP(account, options) {
  const response = await secureApiFetch('/api/fetch-imap', {
    email: account.email,
    clientId: account.clientId,
    refreshToken: account.refreshToken,
    keyword: options.keyword || '',
    limit: Math.max(1, Math.min(Number.parseInt(options.limit, 10) || 1, 10)),
    sender: options.sender || '',
    filters: options.filters || [],
  });

  const data = await parseApiResponse(response);
  if (!data.success) {
    const err = new Error(data.error || data.message || 'IMAP 取件失败');
    err.code = data.code;
    err.reason = data.reason;
    err.action = data.action;
    err.detail = data.detail;
    throw err;
  }
  return data;
}

async function fetchViaMail(account, options = {}) {
  const response = await secureApiFetch('/api/fetch-mail', {
    email: account.email,
    password: account.password,
    keyword: options.keyword || '',
    sender: options.sender || '',
    limit: Math.min(20, Math.max(1, Number.parseInt(options.limit, 10) || 1)),
    filters: options.filters || [],
  });
  const data = await parseApiResponse(response);
  if (!data.success) {
    const err = new Error(data.error || data.message || 'Mail 取件失败');
    for (const key of ['code', 'reason', 'action', 'detail']) err[key] = data[key];
    throw err;
  }
  // 登录会话由服务端持久化；前端只记录有效时间，不保存响应中的会话令牌。
  const latest = getAccounts('mail');
  const saved = latest.find(item => item.id === account.id && item.email === account.email && item.password === account.password);
  if (saved && Number.isFinite(Date.parse(data.sessionExpiresAt || ''))) {
    saved.sessionExpiresAt = data.sessionExpiresAt;
    try { saveAccounts(latest, 'mail'); } catch { /* 邮件已成功取回，不因本地状态写入失败丢弃结果。 */ }
    if (_activeProvider === 'mail') renderAccountList();
  }
  return { ...data, protocol: 'mail', emails: (data.emails || []).map(email => ({ ...email, protocol: 'mail' })) };
}

async function secureApiFetch(url, payload, retryCount = 0) {
  // 非安全上下文（例如用 http://内网IP 直接访问）拿不到 crypto.subtle，
  // 此时退化为明文 JSON；通过 https / localhost 访问时仍使用 AES-GCM + HMAC 加密信封。
  const envelope = window.crypto?.subtle ? await createSecureEnvelope(payload) : { ...payload };
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(envelope),
  });

  if ((response.status === 401 || response.status === 403) && retryCount < SECURITY_RETRY_LIMIT) {
    // 业务认证失败不能靠重建安全封包修复，重试反而会重复登录邮箱。
    // 使用 clone 检查错误，保留原始响应供取件流程读取。
    const failure = await response.clone().json().catch(() => null);
    const securityFailure = failure?.code === 'SECURITY_ENVELOPE_INVALID'
      || (!failure?.code && /安全会话|请求签名|\bnonce\b|加密请求体|请求(?:已)?过期/i.test([failure?.error, failure?.message].filter(value => typeof value === 'string').join(' ')));
    if (securityFailure) {
      _apiSecuritySession = null;
      return secureApiFetch(url, payload, retryCount + 1);
    }
  }

  return response;
}

async function getApiSecuritySession() {
  const now = Date.now();
  if (_apiSecuritySession && _apiSecuritySession.expiresAtMs - now > 60 * 1000) {
    return _apiSecuritySession;
  }

  const response = await fetch('/api/security-session', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{}',
  });
  const data = await response.json();

  if (!response.ok || !data.success) {
    throw new Error(data.error || '安全会话初始化失败');
  }

  _apiSecuritySession = {
    sessionId: data.sessionId,
    sessionToken: data.sessionToken,
    sessionKey: data.sessionKey,
    expiresAtMs: new Date(data.expiresAt).getTime(),
  };

  return _apiSecuritySession;
}

async function createSecureEnvelope(payload) {
  const session = await getApiSecuritySession();
  const keyBytes = base64UrlToBytes(session.sessionKey);
  const cryptoKey = await crypto.subtle.importKey('raw', keyBytes, { name: 'AES-GCM' }, false, ['encrypt']);
  const hmacKey = await crypto.subtle.importKey('raw', keyBytes, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const ivBytes = crypto.getRandomValues(new Uint8Array(12));
  const nonce = bytesToBase64Url(crypto.getRandomValues(new Uint8Array(16)));
  const timestamp = Date.now();
  const plaintext = new TextEncoder().encode(JSON.stringify(payload));
  const encrypted = await crypto.subtle.encrypt({ name: 'AES-GCM', iv: ivBytes }, cryptoKey, plaintext);
  const iv = bytesToBase64Url(ivBytes);
  const ciphertext = bytesToBase64Url(new Uint8Array(encrypted));
  const signedText = `${session.sessionId}.${nonce}.${timestamp}.${iv}.${ciphertext}`;
  const signatureBytes = await crypto.subtle.sign('HMAC', hmacKey, new TextEncoder().encode(signedText));

  return {
    secure: true,
    sessionId: session.sessionId,
    sessionToken: session.sessionToken,
    nonce,
    timestamp,
    iv,
    ciphertext,
    signature: bytesToBase64Url(new Uint8Array(signatureBytes)),
  };
}

function bytesToBase64Url(bytes) {
  let binary = '';
  bytes.forEach(byte => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlToBytes(value) {
  const base64 = value.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(value.length / 4) * 4, '=');
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function parseApiResponse(response) {
  const text = await response.text();
  let data = {};

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { success: false, error: text || `HTTP ${response.status}` };
  }

  if (!response.ok && data.success !== false) {
    data.success = false;
    data.error = data.error || `HTTP ${response.status}`;
  }

  return data;
}

function normalizeFetchError(email, protocol, err) {
  const raw = err?.message || String(err || '未知错误');
  return {
    email,
    protocol,
    code: err?.code || '',
    error: simplifyErrorMessage(raw),
    reason: err?.reason || '',
    action: err?.action || '',
    raw: err?.detail || raw,
  };
}

function simplifyErrorMessage(message) {
  const text = String(message || '').replace(/\s+/g, ' ').trim();

  if (/AADSTS90023\b/i.test(text)) return '当前应用未提供可用的邮箱授权权限，请检查 Mail.Read 授权';
  if (/invalid_grant/i.test(text)) return '刷新令牌无效或已过期，请重新获取 refresh_token';
  if (/invalid_client/i.test(text)) return 'client_id 无效，请检查导入的 clientid';
  if (/interaction_required|consent_required/i.test(text)) return '账号需要重新授权或补充权限';
  if (/AUTHENTICATE|Authentication failed|Invalid credentials|LOGIN failed/i.test(text)) {
    return _activeProvider === 'mail'
      ? 'IMAP 登录失败，请检查账号密码，并确认该邮箱已开启 IMAP'
      : 'IMAP 认证失败，请检查账号权限和令牌 scope';
  }
  if (/timeout|ETIMEDOUT|ESOCKETTIMEDOUT/i.test(text)) return '连接超时，请稍后重试或降低取件数量';
  if (/Failed to fetch|NetworkError/i.test(text)) return '网络请求失败，请检查服务器网络';
  if (/Graph API 错误: ErrorInvalidUser|MailboxNotEnabledForRESTAPI/i.test(text)) return 'Graph 无法访问该邮箱，请检查账号类型或权限';

  return text.length > 180 ? `${text.slice(0, 180)}...` : text;
}

/**
 * 邮件去重
 */
function deduplicateEmails(emails) {
  // 同一封邮件可能被 IMAP 与 Graph 各取一次，而两者的 messageId 形式不同
  // （Graph 给 RFC822 Message-ID，IMAP 工具给 文件夹+UID）。
  // 因此统一用「账号 + 发件地址 + 主题 + 正文指纹」做内容归并，
  // 时间只取到分钟级容差，避免两个协议收到的时刻差几十秒而漏合并。
  const fingerprint = email => {
    const subject = String(email.subject || '').replace(/\s+/g, ' ').trim().toLowerCase();
    const fromAddr = String(email.from || '').toLowerCase().replace(/^.*<|>$/g, '').trim();
    const body = String(email.bodyPreview || email.bodyText || '').replace(/\s+/g, ' ').trim().slice(0, 120).toLowerCase();
    return `${email._account || ''}|${fromAddr}|${subject}|${body}`;
  };
  const seen = new Map();
  emails.forEach(email => {
    const key = fingerprint(email);
    const rawId = String(email.messageId || '').trim().replace(/^<|>$/g, '');
    // 指纹为空（理论上不会）时退回 messageId
    const finalKey = key.replace(/\|/g, '').length > 4 ? key : rawId;
    if (!seen.has(finalKey)) {
      seen.set(finalKey, email);
    } else if (seen.get(finalKey).protocol !== 'graph' && email.protocol === 'graph') {
      // 重复时保留 graph 的（正文通常更完整）
      seen.set(finalKey, email);
    }
  });
  return Array.from(seen.values());
}

// ==================== 邮件渲染 ====================

/**
 * 渲染邮件列表
 */
function renderEmailList(emails, options = {}) {
  Object.assign(_providerStates[_activeProvider], { emails, resultOptions: options, hasFetched: true });
  const listEl = document.getElementById('emailList');
  const headerEl = document.getElementById('resultsHeader');
  const imapCountEl = document.getElementById('imapCount');
  const graphCountEl = document.getElementById('graphCount');
  const totalCountEl = document.getElementById('totalCount');
  const {
    accounts = [],
    completedAccounts = accounts.length,
    isFetching = false,
  } = options;

  headerEl.style.display = 'flex';

  const imapEmails = emails.filter(e => e.protocol === 'imap');
  const graphEmails = emails.filter(e => e.protocol === 'graph');

  imapCountEl.style.display = imapEmails.length > 0 ? 'inline' : 'none';
  graphCountEl.style.display = graphEmails.length > 0 ? 'inline' : 'none';
  imapCountEl.textContent = `IMAP: ${imapEmails.length}`;
  graphCountEl.textContent = `Graph: ${graphEmails.length}`;
  totalCountEl.textContent = isFetching
    ? `已完成 ${completedAccounts}/${accounts.length} · ${emails.length} 封`
    : `共 ${emails.length} 封`;

  if (emails.length === 0) {
    listEl.innerHTML = `
      <div class="empty-state">
        <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="0.8">
          <rect x="2" y="4" width="20" height="16" rx="2"/>
          <path d="M22 7L12 13L2 7"/>
        </svg>
        <p>${isFetching ? '正在等待第一批邮件' : '未获取到邮件'}</p>
        <p class="text-muted">${isFetching ? '邮箱完成后会立即显示在这里' : _activeProvider === 'mail' ? '请检查关键词、发件人或取件数量' : '请检查关键词或协议设置'}</p>
      </div>
    `;
    window._currentEmails = [];
    return;
  }

  window._currentEmails = emails;
  const shouldGroupByAccount = accounts.length > 1 || new Set(emails.map(email => email._account).filter(Boolean)).size > 1;

  let html = '';
  if (shouldGroupByAccount) {
    html = renderGroupedEmailList(emails, accounts, completedAccounts, isFetching);
  } else {
    html = emails.map((email, index) => renderEmailCard(email, index)).join('');
  }

  listEl.innerHTML = html;
}

function renderGroupedEmailList(emails, accounts, completedAccounts, isFetching) {
  const accountMap = new Map();
  accounts.forEach((account, index) => {
    accountMap.set(account.email, {
      account,
      index,
      emails: [],
    });
  });

  emails.forEach((email, index) => {
    const key = email._account || '未知账号';
    if (!accountMap.has(key)) {
      accountMap.set(key, {
        account: { email: key },
        index: email._accountIndex ?? accountMap.size,
        emails: [],
      });
    }
    accountMap.get(key).emails.push({ email, index });
  });

  const groups = Array.from(accountMap.values())
    .filter(group => group.emails.length > 0 || (isFetching && group.index < completedAccounts))
    .sort((a, b) => a.index - b.index);

  return groups.map(group => {
    const imapCount = group.emails.filter(item => item.email.protocol === 'imap').length;
    const graphCount = group.emails.filter(item => item.email.protocol === 'graph').length;
    const statusText = group.index < completedAccounts ? '已完成' : '等待中';
    const countText = `${group.emails.length} 封`;

    return `
      <section class="email-account-group">
        <div class="email-account-group-header">
          <div class="email-account-group-main">
            <span class="account-index-badge">#${group.index + 1}</span>
            <span class="email-account-group-address" title="${escapeAttr(group.account.email)}">${escapeHtml(group.account.email)}</span>
            <span class="email-account-group-status ${group.index < completedAccounts ? 'done' : 'pending'}">${statusText}</span>
          </div>
          <div class="email-account-group-meta">
            <span>${countText}</span>
            ${imapCount ? `<span class="email-protocol imap">IMAP ${imapCount}</span>` : ''}
            ${graphCount ? `<span class="email-protocol graph">Graph ${graphCount}</span>` : ''}
          </div>
        </div>
        <div class="email-account-group-list">
          ${group.emails.map(item => renderEmailCard(item.email, item.index)).join('')}
        </div>
      </section>
    `;
  }).join('');
}

function renderEmailCard(email, index) {
    const date = formatDate(email.date);
    const fromDisplay = email.fromName || email.from;
    const preview = (email.bodyPreview || email.bodyText || '').substring(0, 120);

    return `
      <div class="email-card" style="animation-delay: ${index * 0.04}s" onclick="showEmailDetail(${index})">
        <div class="email-card-header">
          <span class="email-from">
            ${escapeHtml(fromDisplay)}
            <span class="email-protocol ${email.protocol}">${email.protocol}</span>
          </span>
          <span class="email-date">${date}</span>
        </div>
        <div class="email-subject">${escapeHtml(email.subject)}</div>
        <div class="email-preview">${escapeHtml(preview)}</div>
        <div class="email-account-tag">
          <span>${escapeHtml(email._account || '')}</span>
          ${email.folder ? `<span>${escapeHtml(formatFolderName(email.folder))}</span>` : ''}
        </div>
      </div>
    `;
}

function renderFetchIssues(errors) {
  _providerStates[_activeProvider].errors = errors || [];
  const issuesEl = document.getElementById('fetchIssues');
  if (!issuesEl) return;

  if (!errors || errors.length === 0) {
    issuesEl.style.display = 'none';
    issuesEl.innerHTML = '';
    return;
  }

  const grouped = errors.reduce((acc, err) => {
    const key = err.protocol;
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const summary = Object.entries(grouped)
    .map(([protocol, count]) => `${protocol.toUpperCase()} ${count} 个`)
    .join(' / ');

  issuesEl.style.display = 'block';
  issuesEl.innerHTML = `
    <div class="fetch-issues-title">
      <span>${_activeProvider === 'mail' ? '邮箱取件失败' : '部分协议取件失败'}</span>
      <span class="fetch-issues-count">${summary}</span>
    </div>
    <div class="fetch-issues-list">
      ${errors.map(err => `
        <details class="fetch-issue-item">
          <summary>
            <span class="email-protocol ${err.protocol}">${escapeHtml(err.protocol)}</span>
            <span class="fetch-issue-email">${escapeHtml(err.email)}</span>
            <span class="fetch-issue-message">${escapeHtml(err.error)}</span>
          </summary>
          <div class="fetch-issue-detail">
            ${err.code ? `<div><strong>错误码：</strong>${escapeHtml(err.code)}</div>` : ''}
            ${err.reason ? `<div><strong>原因：</strong>${escapeHtml(err.reason)}</div>` : ''}
            ${err.action ? `<div><strong>建议：</strong>${escapeHtml(err.action)}</div>` : ''}
            <pre>${escapeHtml(err.raw || err.error)}</pre>
          </div>
        </details>
      `).join('')}
    </div>
  `;
}

/**
 * 显示邮件详情
 */
function showEmailDetail(index) {
  const email = window._currentEmails?.[index];
  if (!email) return;

  document.getElementById('emailDetailSubject').textContent = email.subject;

  const metaEl = document.getElementById('emailDetailMeta');
  metaEl.innerHTML = `
    <span><span class="meta-label">发件人</span> ${escapeHtml(email.fromName ? `${email.fromName} <${email.from}>` : email.from)}</span>
    <span><span class="meta-label">时间</span> ${formatDate(email.date, true)}</span>
    <span><span class="meta-label">${email.protocol === 'mail' ? '邮箱类型' : '协议'}</span> <span class="email-protocol ${email.protocol}" style="display:inline;font-size:0.75rem;">${email.protocol.toUpperCase()}</span></span>
    ${email.folder ? `<span><span class="meta-label">文件夹</span> ${escapeHtml(formatFolderName(email.folder))}</span>` : ''}
    <span><span class="meta-label">账号</span> ${escapeHtml(email._account || '')}</span>
  `;

  const bodyEl = document.getElementById('emailDetailBody');
  if (email.bodyHtml) {
    // 使用 iframe 隔离 HTML 内容安全风险
    bodyEl.innerHTML = `<iframe srcdoc="${escapeAttr(email.bodyHtml)}" style="width:100%;min-height:300px;border:none;border-radius:8px;background:white;" sandbox="allow-same-origin"></iframe>`;
  } else {
    bodyEl.innerHTML = `<pre style="white-space:pre-wrap;font-family:inherit;">${escapeHtml(email.bodyText || email.bodyPreview || '(无内容)')}</pre>`;
  }

  document.getElementById('emailDetailModal').classList.add('active');
}

// ==================== 工具函数 ====================

/**
 * 格式化日期
 */
function formatDate(dateStr, full = false) {
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;

    const now = new Date();
    const today = now.toDateString();
    const dateDay = d.toDateString();

    if (full) {
      return d.toLocaleString('zh-CN', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit'
      });
    }

    if (today === dateDay) {
      return d.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
    }

    return d.toLocaleDateString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
  } catch {
    return dateStr;
  }
}

function formatFolderName(folder) {
  const value = String(folder || '');
  const normalized = value.toLowerCase();
  if (normalized === 'inbox') return '收件箱';
  if (normalized === 'junkemail' || normalized.includes('junk') || normalized.includes('spam') || value.includes('垃圾')) {
    return '垃圾邮件';
  }
  return value;
}

/**
 * HTML 转义
 */
function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

/**
 * 属性值转义
 */
function escapeAttr(str) {
  return String(str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function copyText(text, successMessage = '已复制', showOnFailure = true) {
  if (!text) {
    if (showOnFailure) showToast('没有可复制的内容', 'warning');
    return false;
  }

  try {
    await navigator.clipboard.writeText(text);
    showToast(successMessage, 'success', 1800);
    return true;
  } catch {
    const input = document.createElement('textarea');
    input.value = text;
    input.style.position = 'fixed';
    input.style.left = '-9999px';
    document.body.appendChild(input);
    input.focus();
    input.select();
    const ok = document.execCommand('copy');
    input.remove();
    if (ok) {
      showToast(successMessage, 'success', 1800);
      return true;
    }
    if (showOnFailure) showToast('复制失败，请手动复制', 'error');
    return false;
  }
}

/**
 * Toast 通知（支持成功大横幅模式）
 */
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  const icons = {
    success: '✅',
    error: '❌',
    warning: '⚠️',
    info: 'ℹ️',
  };

  toast.innerHTML = `<span>${icons[type] || ''}</span><span>${escapeHtml(message)}</span>`;
  container.appendChild(toast);

  // 点击 toast 提前关闭
  toast.addEventListener('click', () => {
    toast.classList.add('toast-exit');
    setTimeout(() => toast.remove(), 300);
  });
  toast.style.cursor = 'pointer';

  setTimeout(() => {
    toast.classList.add('toast-exit');
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

/**
 * 显示导入成功大横幅
 */
function showImportSuccessBanner(uniqueCount, duplicateCount) {
  const banner = document.createElement('div');
  banner.className = 'import-success-banner';
  
  let msg = `🎉 成功导入 <strong>${uniqueCount}</strong> 个邮箱`;
  if (duplicateCount > 0) {
    msg += `，跳过 <strong>${duplicateCount}</strong> 个重复`;
  }

  banner.innerHTML = `
    <div class="import-success-icon">
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
    </div>
    <div class="import-success-text">${msg}</div>
  `;

  document.body.appendChild(banner);

  // 触发动画
  requestAnimationFrame(() => {
    banner.classList.add('active');
  });

  // 3秒后移除
  setTimeout(() => {
    banner.classList.add('exit');
    setTimeout(() => banner.remove(), 400);
  }, 2500);
}

/**
 * 保存当前邮箱提供商的操作状态
 */
function setStatus(state, text) {
  _providerStates[_activeProvider].status = { state, text };
}

/**
 * 进度条控制
 */
function showProgress(show) {
  document.getElementById('progressContainer').style.display = show ? 'block' : 'none';
}

function updateProgress(percent, text) {
  document.getElementById('progressFill').style.width = `${percent}%`;
  document.getElementById('progressText').textContent = text;
  document.getElementById('progressPercent').textContent = `${percent}%`;
}

/**
 * Textarea 输入验证震动效果
 */
function shakeElement(el) {
  el.classList.add('shake');
  setTimeout(() => el.classList.remove('shake'), 500);
}

/**
 * 按钮加载状态
 */
function setButtonLoading(btn, loading, originalText = '') {
  if (loading) {
    btn.dataset.originalHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `
      <span class="btn-spinner"></span>
      处理中...
    `;
  } else {
    btn.disabled = false;
    btn.innerHTML = btn.dataset.originalHtml || originalText;
  }
}

// ==================== 事件绑定 ====================

document.addEventListener('DOMContentLoaded', () => {
  initializeTokenRefreshExportDialog();
  initializeProviderSwitcher();
  initializeFilters();
  // 另一标签页刷新成功后同步标记；定时更新剩余冷却时间。
  window.addEventListener('storage', event => {
    if (event.key === accountStorageKey()) renderAccountList();
  });
  setInterval(() => {
    if (getAccounts().some(account => account.tokenRefreshedAt || account.sessionExpiresAt)) renderAccountList();
  }, 30000);

  // 导入弹窗
  const importModal = document.getElementById('importModal');
  const emailDetailModal = document.getElementById('emailDetailModal');
  const importTextarea = document.getElementById('importTextarea');
  const btnConfirmImport = document.getElementById('btnConfirmImport');
  const accountSearch = document.getElementById('accountSearch');

  accountSearch.addEventListener('input', () => {
    renderAccountList();
  });

  document.getElementById('btnOpenImport').addEventListener('click', () => {
    importModal.classList.add('active');
    importTextarea.focus();
    // 重置 textarea 状态
    importTextarea.classList.remove('textarea-error');
    updateImportPreview();
  });

  document.getElementById('btnCloseModal').addEventListener('click', () => {
    importModal.classList.remove('active');
  });

  document.getElementById('btnCancelImport').addEventListener('click', () => {
    importModal.classList.remove('active');
  });

  // 点击背景关闭弹窗
  importModal.addEventListener('click', (e) => {
    if (e.target === importModal) importModal.classList.remove('active');
  });

  emailDetailModal.addEventListener('click', (e) => {
    if (e.target === emailDetailModal) emailDetailModal.classList.remove('active');
  });

  document.getElementById('btnCloseDetail').addEventListener('click', () => {
    emailDetailModal.classList.remove('active');
  });

  // 实时预览导入数据
  importTextarea.addEventListener('input', () => {
    importTextarea.classList.remove('textarea-error');
    updateImportPreview();
  });

  // 确认导入
  btnConfirmImport.addEventListener('click', () => {
    const text = importTextarea.value;
    if (!text.trim()) {
      showToast('请输入邮箱信息', 'warning');
      importTextarea.classList.add('textarea-error');
      shakeElement(importTextarea);
      importTextarea.focus();
      return;
    }

    // 按钮加载状态
    setButtonLoading(btnConfirmImport, true);

    // 模拟微小延迟让用户感知处理过程
    setTimeout(() => {
      const { accounts: newAccounts, errors } = parseImportText(text);

      if (errors.length > 0) {
        errors.forEach(err => showToast(err, 'error', 5000));
      }

      if (newAccounts.length > 0) {
        const existing = getAccounts();
        // 去重（基于邮箱地址）
        const existingEmails = new Set(existing.map(a => a.email.toLowerCase()));
        const unique = newAccounts.filter(a => {
          const email = a.email.toLowerCase();
          if (existingEmails.has(email)) return false;
          existingEmails.add(email);
          return true;
        });
        const duplicates = newAccounts.length - unique.length;

        if (unique.length === 0 && duplicates > 0) {
          // 全部重复
          setButtonLoading(btnConfirmImport, false);
          showToast(`所有 ${duplicates} 个邮箱都已存在，跳过导入`, 'warning');
          return;
        }

        saveAccounts([...existing, ...unique]);
        
        // 记录新导入的ID用于高亮
        const newIds = new Set(unique.map(a => a.id));
        renderAccountList(newIds);

        // 清空输入
        importTextarea.value = '';
        updateImportPreview();

        // 关闭弹窗
        importModal.classList.remove('active');
        
        // 恢复按钮
        setButtonLoading(btnConfirmImport, false);

        // 显示成功横幅
        showImportSuccessBanner(unique.length, duplicates);

        // 滚动账号列表到底部以显示新导入的账号
        setTimeout(() => {
          const accountList = document.getElementById('accountList');
          accountList.scrollTop = accountList.scrollHeight;
        }, 100);

      } else if (errors.length === 0) {
        setButtonLoading(btnConfirmImport, false);
        showToast('未解析到有效的邮箱数据', 'warning');
        importTextarea.classList.add('textarea-error');
        shakeElement(importTextarea);
      } else {
        setButtonLoading(btnConfirmImport, false);
      }
    }, 200);
  });

  // 清空全部
  document.getElementById('btnClearAll').addEventListener('click', () => {
    const provider = _activeProvider;
    const accounts = getAccounts();
    if (accounts.length === 0) {
      showToast('没有可清空的邮箱', 'info');
      return;
    }
    if (confirm(`确定要清空全部 ${accounts.length} 个邮箱吗？`)) {
      // 播放移除动画
      document.querySelectorAll('.account-item').forEach(item => {
        item.classList.add('account-item-removing');
      });
      setTimeout(() => {
        saveAccounts([], provider);
        if (_activeProvider === provider) {
          renderAccountList();
          showToast('已清空全部邮箱', 'info');
        }
      }, 300);
    }
  });

  document.getElementById('btnExportAccounts').addEventListener('click', exportAccounts);
  document.getElementById('btnRefreshSelected').addEventListener('click', () => refreshAccountTokens(getSelectedAccounts()));
  document.getElementById('btnRefreshAll').addEventListener('click', () => refreshAccountTokens(getAccounts()));

  // 选中取件
  document.getElementById('btnFetchSelected').addEventListener('click', () => {
    const selected = getSelectedAccounts();
    if (selected.length === 0) {
      showToast('请先在左侧勾选要取件的邮箱', 'warning');
      // 闪烁左侧面板提示
      const sidebar = document.querySelector('.sidebar');
      sidebar.classList.add('sidebar-highlight');
      setTimeout(() => sidebar.classList.remove('sidebar-highlight'), 1500);
      return;
    }
    startFetch(selected);
  });

  // 全部取件
  document.getElementById('btnFetchAll').addEventListener('click', () => {
    const accounts = getAccounts();
    if (accounts.length === 0) {
      showToast('请先导入邮箱', 'warning');
      return;
    }
    startFetch(accounts);
  });

  // Ctrl+Enter 快捷键提交导入
  importTextarea.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      btnConfirmImport.click();
    }
  });

  // ESC 关闭弹窗
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      importModal.classList.remove('active');
      emailDetailModal.classList.remove('active');
    }
  });
});

/**
 * 更新导入预览：实时显示解析了多少行
 */
function updateImportPreview() {
  const textarea = document.getElementById('importTextarea');
  let previewEl = document.getElementById('importPreview');
  
  if (!previewEl) {
    previewEl = document.createElement('div');
    previewEl.id = 'importPreview';
    previewEl.className = 'import-preview';
    textarea.parentElement.appendChild(previewEl);
  }

  const text = textarea.value;
  if (!text.trim()) {
    previewEl.innerHTML = '<span class="preview-hint">💡 粘贴邮箱数据后将自动预览，支持 Ctrl+Enter 快捷导入</span>';
    return;
  }

  const lines = text.split(_activeProvider === 'mail' ? /\r\n|\n|\r/ : /\n/).filter(l => l.trim());
  const { accounts, errors } = parseImportText(text);
  
  // 检查与现有邮箱的重复
  const existing = getAccounts();
  const existingEmails = new Set(existing.map(a => a.email.toLowerCase()));
  const duplicates = accounts.filter(a => existingEmails.has(a.email.toLowerCase()));
  const unique = accounts.length - duplicates.length;

  let html = `<span class="preview-count">📋 识别 ${lines.length} 行`;
  
  if (accounts.length > 0) {
    html += ` → <span class="preview-valid">✅ ${accounts.length} 个有效</span>`;
  }
  
  if (unique < accounts.length && duplicates.length > 0) {
    html += ` <span class="preview-dup">⚠️ ${duplicates.length} 个重复</span>`;
  }
  
  if (errors.length > 0) {
    html += ` <span class="preview-error">❌ ${errors.length} 个错误</span>`;
  }
  
  html += '</span>';
  previewEl.innerHTML = html;
}

/**
 * 启动取件流程
 */
function startFetch(accounts) {
  return runMailOperation('fetch', () => {
    const keyword = document.getElementById('searchKeyword').value.trim();
    const sender = document.getElementById('searchSender').value.trim();
    const limit = parseInt(document.getElementById('fetchLimit').value) || 1;
    const requestedIds = new Set(accounts.map(account => account.id));
    const latest = getAccounts().filter(account => requestedIds.has(account.id));
    // 勾选的预设过滤条件随请求下发，服务端过滤后再返回
    return fetchEmails(latest, { keyword, sender, limit, filters: getActiveFilters() });
  });
}

/* ==================== 过滤条件（服务器持久化） ==================== */

let _savedFilters = [];

function getActiveFilters() {
  const checked = new Set(
    Array.from(document.querySelectorAll('.filter-option input:checked'), el => el.dataset.id)
  );
  return _savedFilters.filter(item => checked.has(item.id));
}

async function loadFilters() {
  try {
    const res = await fetch('/api/filters', { method: 'GET', cache: 'no-store' });
    const data = await res.json();
    _savedFilters = Array.isArray(data.filters) ? data.filters : [];
  } catch {
    _savedFilters = [];
  }
  renderFilters();
}

async function persistFilters(nextFilters) {
  const res = await fetch('/api/filters', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ filters: nextFilters }),
  });
  const data = await res.json();
  if (!data.success) throw new Error(data.error || '保存过滤条件失败');
  _savedFilters = data.filters || [];
  renderFilters();
}

function renderFilters() {
  const savedEl = document.getElementById('filterSaved');
  const chipsEl = document.getElementById('filterChips');
  const emptyEl = document.getElementById('filterEmpty');
  if (!savedEl || !chipsEl) return;

  const scopeLabel = { all: '标题+正文', subject: '仅标题', body: '仅正文' };

  if (_savedFilters.length === 0) {
    savedEl.innerHTML = '<p class="filter-empty">暂无保存的过滤条件，添加后会存到服务器，下次打开可直接勾选。</p>';
    chipsEl.innerHTML = '';
    updateFilterCount();
    return;
  }

  // 记住勾选状态，重绘后保持
  const checkedIds = new Set(
    Array.from(document.querySelectorAll('.filter-option input:checked'), el => el.dataset.id)
  );

  savedEl.innerHTML = _savedFilters.map(item => `
    <label class="filter-option ${checkedIds.has(item.id) ? 'checked' : ''}">
      <input type="checkbox" data-id="${escapeAttr(item.id)}" ${checkedIds.has(item.id) ? 'checked' : ''}>
      <span class="filter-option-text" title="${escapeAttr(item.text)}">${escapeHtml(item.text)}</span>
      <span class="filter-option-scope">${scopeLabel[item.scope] || '标题+正文'}</span>
      <button type="button" class="filter-option-del" data-id="${escapeAttr(item.id)}" title="删除">×</button>
    </label>
  `).join('');

  savedEl.querySelectorAll('.filter-option input').forEach(input => {
    input.addEventListener('change', () => {
      input.closest('.filter-option')?.classList.toggle('checked', input.checked);
      renderActiveChips();
      updateFilterCount();
    });
  });

  savedEl.querySelectorAll('.filter-option-del').forEach(btn => {
    btn.addEventListener('click', async event => {
      event.preventDefault();
      event.stopPropagation();
      const id = btn.dataset.id;
      try {
        await persistFilters(_savedFilters.filter(item => item.id !== id));
        showToast('已删除过滤条件', 'info', 1800);
      } catch (error) {
        showToast(error.message, 'error');
      }
    });
  });

  renderActiveChips();
  updateFilterCount();
}

function renderActiveChips() {
  const chipsEl = document.getElementById('filterChips');
  if (!chipsEl) return;
  const active = getActiveFilters();
  const scopeLabel = { all: '', subject: '标题', body: '正文' };
  chipsEl.innerHTML = active.map(item => `
    <span class="filter-chip">
      <span class="filter-chip-text" title="${escapeAttr(item.text)}">${escapeHtml(item.text)}</span>
      ${scopeLabel[item.scope] ? `<span class="filter-chip-scope">${scopeLabel[item.scope]}</span>` : ''}
    </span>
  `).join('');
}

function updateFilterCount() {
  const badge = document.getElementById('filterActiveCount');
  if (!badge) return;
  const count = getActiveFilters().length;
  badge.textContent = String(count);
  badge.hidden = count === 0;
}

function initializeFilters() {
  const toggle = document.getElementById('btnToggleFilters');
  const panel = document.getElementById('filterPanel');
  const input = document.getElementById('filterInput');
  const scope = document.getElementById('filterScope');
  const addBtn = document.getElementById('btnAddFilter');

  toggle?.addEventListener('click', () => {
    panel.hidden = !panel.hidden;
    if (!panel.hidden) input.focus();
  });

  async function addFilter() {
    const text = input.value.trim();
    if (!text) {
      shakeElement(input);
      input.focus();
      return;
    }
    const duplicate = _savedFilters.some(
      item => item.text.toLowerCase() === text.toLowerCase() && item.scope === scope.value
    );
    if (duplicate) {
      showToast('该过滤条件已存在', 'warning');
      return;
    }
    try {
      await persistFilters([..._savedFilters, { text, scope: scope.value }]);
      input.value = '';
      input.focus();
      showToast(`已保存过滤条件：${text}`, 'success', 2000);
    } catch (error) {
      showToast(error.message, 'error');
    }
  }

  addBtn?.addEventListener('click', addFilter);
  input?.addEventListener('keydown', event => {
    if (event.key === 'Enter') {
      event.preventDefault();
      addFilter();
    }
  });

  loadFilters();
}
