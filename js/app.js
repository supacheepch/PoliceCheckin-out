/**
 * app.js — Logic หลักของระบบ Check-in / Check-out
 * ==================================================
 */

// ── State ──────────────────────────────────────────
const State = {
  officers:        [],        // รายชื่อตำรวจทั้งหมด
  selectedOfficer: null,      // ตำรวจที่เลือก
  logs:            [],        // ประวัติการ Check-in/out (today)
  dataSource:      'demo',    // 'sheet' | 'demo'
  loading:         false,
};

// ── LocalStorage ──────────────────────────────────
const Store = {
  load() {
    try {
      const raw = localStorage.getItem(CONFIG.STORAGE_KEY);
      if (!raw) return [];
      const all = JSON.parse(raw);
      // เก็บเฉพาะวันนี้
      const today = new Date().toDateString();
      return all.filter(e => new Date(e.timestamp).toDateString() === today);
    } catch { return []; }
  },
  save(logs) {
    try {
      // Load ทั้งหมด แล้ว merge (เพื่อไม่ลบข้อมูลเก่า)
      const raw = localStorage.getItem(CONFIG.STORAGE_KEY);
      const existing = raw ? JSON.parse(raw) : [];
      const today = new Date().toDateString();
      const oldLogs = existing.filter(e => new Date(e.timestamp).toDateString() !== today);
      const merged = [...oldLogs, ...logs].slice(-CONFIG.MAX_LOGS);
      localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(merged));
    } catch (e) {
      console.warn('Storage save failed:', e);
    }
  }
};

// ── Toast ──────────────────────────────────────────
const Toast = {
  _container: null,

  init() {
    this._container = document.getElementById('toastContainer');
  },

  show(type, title, msg, duration = 3500) {
    const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.innerHTML = `
      <div class="toast-icon">${icons[type] || 'ℹ️'}</div>
      <div class="toast-body">
        <div class="toast-title">${title}</div>
        ${msg ? `<div class="toast-msg">${msg}</div>` : ''}
      </div>`;
    this._container.appendChild(el);
    setTimeout(() => {
      el.classList.add('hiding');
      el.addEventListener('animationend', () => el.remove());
    }, duration);
  }
};

// ── Clock ──────────────────────────────────────────
function startClock() {
  const timeEl = document.getElementById('clockTime');
  const dateEl = document.getElementById('clockDate');

  const thDays = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  const thMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                    'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

  function tick() {
    const now = new Date();
    const hh  = String(now.getHours()).padStart(2, '0');
    const mm  = String(now.getMinutes()).padStart(2, '0');
    const ss  = String(now.getSeconds()).padStart(2, '0');
    timeEl.textContent = `${hh}:${mm}:${ss}`;

    const day   = thDays[now.getDay()];
    const date  = now.getDate();
    const month = thMonths[now.getMonth()];
    const year  = now.getFullYear() + 543; // พ.ศ.
    dateEl.textContent = `วัน${day}ที่ ${date} ${month} ${year}`;
  }

  tick();
  setInterval(tick, 1000);
}

// ── Officers ──────────────────────────────────────
async function loadOfficers() {
  setLoadingState(true);
  document.getElementById('officerListWrapper').innerHTML = `
    <div class="loading-row" style="text-align:center;padding:30px;color:var(--text-muted);">
      <span class="spinner"></span> กำลังโหลดข้อมูล...
    </div>`;

  const { officers, source, error } = await GoogleSheet.loadOfficers();

  State.officers   = officers;
  State.dataSource = source;

  renderOfficerSelect();
  renderOfficerList();
  updateSourceBadge(source, error);
  setLoadingState(false);
}

function updateSourceBadge(source, error) {
  const badge = document.getElementById('sourceBadge');
  if (source === 'sheet') {
    badge.className = 'source-badge';
    badge.innerHTML = `<span>●</span> Google Sheet (${State.officers.length} คน)`;
  } else {
    badge.className = 'source-badge' + (error ? ' error' : '');
    badge.innerHTML = error
      ? `<span>●</span> Demo Mode (ไม่พบ Sheet)`
      : `<span>●</span> Demo Mode`;
  }
}

function renderOfficerSelect() {
  const sel = document.getElementById('officerSelect');
  sel.innerHTML = '<option value="">-- เลือกตำรวจ --</option>';
  State.officers.forEach((o, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `${o.rank} ${o.name}`;
    sel.appendChild(opt);
  });
}

function renderOfficerList() {
  const wrapper = document.getElementById('officerListWrapper');
  if (State.officers.length === 0) {
    wrapper.innerHTML = `<div class="empty-state"><div class="icon">👮</div><p>ไม่พบรายชื่อ</p></div>`;
    return;
  }

  wrapper.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'officer-list';

  State.officers.forEach((o, i) => {
    const initials = (o.name || '??').split(' ').map(w => w[0]).slice(0, 2).join('');
    const card = document.createElement('div');
    card.className = 'officer-card';
    card.style.animationDelay = `${i * 30}ms`;
    card.innerHTML = `
      <div class="officer-avatar">${initials}</div>
      <div class="officer-details">
        <div class="name">${o.rank} ${o.name}</div>
        <div class="rank">${o.dept || 'ไม่ระบุสังกัด'} ${o.id ? '• ' + o.id : ''}</div>
      </div>`;
    card.addEventListener('click', () => selectOfficerByIndex(i));
    list.appendChild(card);
  });

  wrapper.appendChild(list);
}

function selectOfficerByIndex(index) {
  const o = State.officers[index];
  if (!o) return;
  State.selectedOfficer = { ...o, index };

  // อัปเดต select dropdown
  document.getElementById('officerSelect').value = index;

  // อัปเดต info box
  const box = document.getElementById('officerInfo');
  box.classList.add('active');
  document.getElementById('infoName').textContent = `${o.rank} ${o.name}`;
  document.getElementById('infoRank').textContent = `${o.dept || ''} ${o.id ? '• รหัส: ' + o.id : ''}`.trim();

  // Highlight officer card
  document.querySelectorAll('.officer-card').forEach((c, i) => {
    c.style.background = i === index
      ? 'rgba(59,130,246,0.12)'
      : 'rgba(255,255,255,0.03)';
    c.style.borderColor = i === index
      ? 'rgba(59,130,246,0.4)'
      : 'var(--border)';
  });
}

function onSelectChange() {
  const idx = document.getElementById('officerSelect').value;
  if (idx === '') {
    State.selectedOfficer = null;
    document.getElementById('officerInfo').classList.remove('active');
    document.querySelectorAll('.officer-card').forEach(c => {
      c.style.background = '';
      c.style.borderColor = '';
    });
    return;
  }
  selectOfficerByIndex(parseInt(idx));
}

// ── Check In / Out ────────────────────────────────
function getThaiDateTime() {
  const now = new Date();
  const hh  = String(now.getHours()).padStart(2, '0');
  const mm  = String(now.getMinutes()).padStart(2, '0');
  const ss  = String(now.getSeconds()).padStart(2, '0');
  const date = now.getDate();
  const months = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return {
    time:      `${hh}:${mm}:${ss}`,
    date:      `${date} ${months[now.getMonth()]} ${now.getFullYear() + 543}`,
    timestamp: now.toISOString(),
  };
}

function recordAction(type) {
  if (!State.selectedOfficer) {
    Toast.show('warning', 'กรุณาเลือกตำรวจ', 'เลือกชื่อจากรายการก่อนกดปุ่ม');
    return;
  }

  const dt = getThaiDateTime();
  const o  = State.selectedOfficer;

  // ตรวจสอบ: ถ้ายัง Check-in อยู่ ห้าม Check-in ซ้ำ (ตรวจจากล่าสุด)
  const lastLog = State.logs.filter(l => l.officerId === o.id || l.officerName === o.name)
                            .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))[0];

  if (type === 'in' && lastLog && lastLog.type === 'in') {
    Toast.show('warning', 'เช็คอินซ้ำ', `${o.rank} ${o.name} เช็คอินแล้ว ยังไม่ได้เช็คเอาต์`);
    return;
  }
  if (type === 'out' && (!lastLog || lastLog.type === 'out')) {
    Toast.show('warning', 'ยังไม่ได้เช็คอิน', `${o.rank} ${o.name} ยังไม่ได้เช็คอินวันนี้`);
    return;
  }

  const entry = {
    id:          Date.now(),
    type,
    officerName: o.name,
    officerRank: o.rank,
    officerDept: o.dept,
    officerId:   o.id,
    time:        dt.time,
    date:        dt.date,
    timestamp:   dt.timestamp,
  };

  State.logs.unshift(entry);
  Store.save(State.logs);

  renderLogTable();
  updateStats();

  const label = type === 'in' ? 'เช็คอินสำเร็จ ✔' : 'เช็คเอาต์สำเร็จ ✔';
  const color = type === 'in' ? 'success' : 'error';
  Toast.show(color, label, `${o.rank} ${o.name} — ${dt.time}`);

  // Flash animation on button
  const btnId = type === 'in' ? 'btnCheckIn' : 'btnCheckOut';
  const btn = document.getElementById(btnId);
  btn.style.transform = 'scale(0.95)';
  setTimeout(() => btn.style.transform = '', 150);
}

// ── Stats ─────────────────────────────────────────
function updateStats() {
  const total   = State.officers.length;
  const inNames = new Set(
    State.logs.filter(l => l.type === 'in').map(l => l.officerName)
  );
  const outNames = new Set(
    State.logs.filter(l => {
      const logs = State.logs.filter(x => x.officerName === l.officerName)
                             .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      return logs[0] && logs[0].type === 'out';
    }).map(l => l.officerName)
  );

  // คนที่ยัง active (check-in ล่าสุด = 'in')
  const activeSet = new Set();
  const officerNames = [...new Set(State.logs.map(l => l.officerName))];
  officerNames.forEach(name => {
    const sorted = State.logs.filter(l => l.officerName === name)
                             .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    if (sorted[0] && sorted[0].type === 'in') activeSet.add(name);
  });

  document.getElementById('statTotal').textContent  = total;
  document.getElementById('statActive').textContent = activeSet.size;
  document.getElementById('statOff').textContent    = Math.max(0, inNames.size - activeSet.size);
}

// ── Log Table ─────────────────────────────────────
function renderLogTable() {
  const tbody = document.getElementById('logTbody');

  if (State.logs.length === 0) {
    tbody.innerHTML = `
      <tr class="loading-row">
        <td colspan="5">
          <div class="empty-state">
            <div class="icon">📋</div>
            <p>ยังไม่มีการเช็คอิน/เช็คเอาต์วันนี้</p>
          </div>
        </td>
      </tr>`;
    return;
  }

  tbody.innerHTML = State.logs.map(e => `
    <tr>
      <td>
        <span class="status-badge ${e.type}">
          <span class="dot"></span>
          ${e.type === 'in' ? 'เช็คอิน' : 'เช็คเอาต์'}
        </span>
      </td>
      <td class="td-name">${e.officerRank} ${e.officerName}</td>
      <td class="td-rank">${e.officerDept || '-'}</td>
      <td style="font-family:monospace;font-size:0.85rem;">${e.time}</td>
      <td style="color:var(--text-muted);font-size:0.8rem;">${e.date}</td>
    </tr>`).join('');
}

// ── Loading State ─────────────────────────────────
function setLoadingState(on) {
  State.loading = on;
  const btn = document.getElementById('btnReload');
  if (on) btn.classList.add('spinning');
  else    btn.classList.remove('spinning');
}

// ── Init ──────────────────────────────────────────
async function init() {
  Toast.init();
  startClock();

  // โหลด log จาก LocalStorage
  State.logs = Store.load();
  renderLogTable();
  updateStats();

  // โหลดรายชื่อจาก Google Sheet
  await loadOfficers();
}

// Start!
document.addEventListener('DOMContentLoaded', init);
