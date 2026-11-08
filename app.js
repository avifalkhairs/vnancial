/* Vnancial – phone app (PWA). Data: snapshot.json from Excel via Apps Script; new entries go to the Drive inbox. */
'use strict';
const APP_VER = '1.3.0';
const $ = (s, el = document) => el.querySelector(s);
const LS = {
  get(k, d) { try { const v = localStorage.getItem('vn.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('vn.' + k, JSON.stringify(v)); } catch (e) {} }
};

/* ---------------------------------------------------------------- language */
const T = {
  id: {
    home: 'Beranda', tx: 'Transaksi', wallet: 'Dompet', more: 'Lainnya', add: 'Tambah', save: 'SIMPAN',
    totalBalance: 'Total saldo', incMonth: 'Pemasukan bln ini', expMonth: 'Pengeluaran',
    safeToday: 'Aman dibelanjakan hari ini', safe: 'Aman', careful: 'Hati-hati', over: 'Lewat batas',
    usedOf: (a, b) => `Terpakai ${a} dari batas ${b}`, budget: 'Anggaran', detail: 'Detail', latest: 'Terakhir', all: 'Semua',
    expense: 'Keluar', income: 'Masuk', transfer: 'Transfer', debt: 'Utang', plan: 'Rencana', investment: 'Investasi',
    net: 'Bersih', today: 'Hari ini', yesterday: 'Kemarin', nothing: 'Belum ada transaksi.',
    amount: 'Jumlah', desc: 'Keterangan', date: 'Tanggal', account: 'Akun', category: 'Kategori', subcat: 'Sub-kategori',
    paidTo: 'Dibayar ke (opsional)', from: 'Dari (opsional)', notes: 'Catatan (opsional)', fromPlan: 'Dari rencana (opsional)',
    none: '—', fromAcc: 'Dari akun', toAcc: 'Ke akun', fee: 'Biaya admin', action: 'Aksi', person: 'Nama orang',
    item: 'Nama rencana', kind: 'Jenis', repeat: 'Ulang', once: 'Sekali', monthly: 'Bulanan',
    ticker: 'Kode', units: 'Jumlah unit', price: 'Harga per unit (dividen: total)',
    afterSafe: v => `Sisa aman hari ini setelah ini: ${v}`,
    savedSent: 'Tersimpan · Excel membacanya saat dibuka', savedQueued: 'Tersimpan di HP · dikirim saat online',
    needAmount: 'Isi jumlah', needAcc: 'Pilih akun', needCat: 'Pilih kategori', needPerson: 'Isi nama orang', needItem: 'Isi nama rencana',
    needTicker: 'Isi kode', sameAcc: 'Akun asal dan tujuan sama',
    bank: 'Bank', ewallet: 'Dompet digital', cash: 'Tunai', other: 'Lainnya', totalCash: 'Total kas', netWorth: 'Kekayaan bersih',
    iOwe: 'Saya berutang', owesMe: 'Orang berutang', people: n => `${n} orang`, ongoing: 'Berjalan', settled: 'Lunas',
    pay: 'Bayar', receive: 'Terima', debtNote: 'Bayar / Terima mencatat cicilan ke akun yang dipilih.',
    planQ: 'Apakah uang saya cukup? · 12 bulan', enough: 'CUKUP', short: 'KURANG',
    lowest: (v, m) => `Saldo terendah ${v} di ${m}`, projection: 'Proyeksi saldo', upcoming: 'Akan datang', open: n => `${n} terbuka`,
    noPlans: 'Belum ada rencana.', holdings: 'Posisi', record: 'Catat', invValue: 'Nilai pasar', pl: 'Untung/rugi',
    sync: 'Sinkron dengan Excel', connected: 'Tersambung', offline: 'Offline', notSet: 'Belum diatur',
    lastSync: t => `Diperbarui: ${t}`, waiting: n => `${n} entri menunggu dibaca Excel`, queued: n => `${n} entri belum terkirim`,
    syncNow: 'Sinkronkan sekarang', language: 'Bahasa · Language', langNote: 'Berlaku untuk seluruh aplikasi',
    settings: 'Sambungan', url: 'URL Web App (Apps Script)', token: 'Token', test: 'Simpan & coba', demo: 'Lihat demo',
    setupTitle: 'Sambungkan ke Excel', setupNote: 'Tempel URL Web App dan token dari langkah Apps Script di panduan.',
    okConn: 'Tersambung', failConn: 'Gagal tersambung: ', pending: 'menunggu', queuedTag: 'di HP',
    debts: 'Utang & piutang', plans: 'Rencana', budgets: 'Anggaran', invest: 'Investasi', catAcc: 'Kategori & akun',
    catAccNote: 'Mengikuti Settings di Excel', back: 'Kembali', photo: 'Foto struk', photoOn: 'Foto terlampir',
    demoOn: 'Mode demo · data contoh', snapshotOld: 'Data Excel belum ada. Buka Excel sekali agar snapshot terkirim.',
    simple: 'Perkiraan dari saldo saat ini dan rencana terbuka', borrow: 'Pinjam', repay: 'Bayar utang', lend: 'Meminjamkan', collect: 'Terima bayaran',
    buy: 'Beli', sell: 'Jual', dividend: 'Dividen', search: 'Cari', months: ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'],
    days: ['Min','Sen','Sel','Rab','Kam','Jum','Sab'], disconnect: 'Putuskan sambungan', version: 'Versi',
    edit: 'Ubah', del: 'Hapus', cancel: 'Batal', delQ: 'Hapus transaksi ini? Excel ikut menghapusnya.', deletedT: 'Dihapus · Excel ikut memperbarui',
    editedT: 'Perubahan disimpan · Excel ikut memperbarui', noRef: 'Transaksi ini diketik langsung di Excel tanpa Ref, ubah di Excel.', tapEdit: 'Ketuk transaksi untuk mengubah atau menghapus',
    setTitle: 'Pengaturan', prefs: 'Preferensi', alloc: 'Alokasi anggaran', accountsT: 'Akun', expCats: 'Kategori pengeluaran', incCats: 'Kategori pemasukan', subcats: 'Sub-kategori',
    addAcc: '+ Tambah akun', addCat: '+ Tambah kategori', addSub: '+ Tambah sub-kategori', name: 'Nama', type: 'Jenis', opening: 'Saldo awal', bucket: 'Kelompok anggaran', fixed: 'Tagihan tetap?',
    yes: 'Ya', no: 'Tidak', share: 'Porsi', total: 'Total', mustTotal: 'Total harus 100%', savedSet: 'Disimpan · Excel ikut memperbarui',
    renameNote: 'Transaksi lama tetap memakai nama lama.', lastItem: 'Minimal harus ada satu.', exists: 'Nama sudah ada.', needName: 'Isi nama',
    connection: 'Sambungan ke Excel', emptyAuto: 'kosong = otomatis', color: 'Warna', hexPh: 'atau kode warna, mis. #F15A23',
    pl_BasisMethod: 'Basis anggaran', pl_BasisManual: 'Jumlah basis manual', pl_OpenDate: 'Tanggal saldo awal', pl_DailyManual: 'Batas harian', pl_EmergMonths: 'Dana darurat (bulan)',
    pl_DebtOn: 'Utang mengubah saldo', pl_FeeCat: 'Kategori biaya transfer', pl_FeeSub: 'Sub-kategori biaya transfer', pl_LocalInbox: 'Folder inbox lokal', pl_CloudInbox: 'Folder inbox Google Drive',
    pl_Owner: 'Pemilik', pl_InvOn: 'Investasi mengubah saldo', pl_SafeMargin: 'Margin aman belanja', pl_SyncFolder: 'Folder sinkron HP',
    bm: { 'Last Month': 'Bulan lalu', 'Average 3 Months': 'Rata-rata 3 bulan', 'This Month': 'Bulan ini', 'Manual': 'Manual' }
  },
  en: {
    home: 'Home', tx: 'Activity', wallet: 'Wallet', more: 'More', add: 'Add', save: 'SAVE',
    totalBalance: 'Total balance', incMonth: 'Income this month', expMonth: 'Spending',
    safeToday: 'Safe to spend today', safe: 'Safe', careful: 'Careful', over: 'Over limit',
    usedOf: (a, b) => `Used ${a} of ${b}`, budget: 'Budget', detail: 'Details', latest: 'Latest', all: 'All',
    expense: 'Expense', income: 'Income', transfer: 'Transfer', debt: 'Debt', plan: 'Plan', investment: 'Investment',
    net: 'Net', today: 'Today', yesterday: 'Yesterday', nothing: 'No transactions yet.',
    amount: 'Amount', desc: 'Description', date: 'Date', account: 'Account', category: 'Category', subcat: 'Sub-category',
    paidTo: 'Paid to (optional)', from: 'From (optional)', notes: 'Notes (optional)', fromPlan: 'From plan (optional)',
    none: '—', fromAcc: 'From account', toAcc: 'To account', fee: 'Fee', action: 'Action', person: 'Person',
    item: 'Plan name', kind: 'Kind', repeat: 'Repeat', once: 'Once', monthly: 'Monthly',
    ticker: 'Ticker', units: 'Units', price: 'Price per unit (dividend: total)',
    afterSafe: v => `Safe left today after this: ${v}`,
    savedSent: 'Saved · Excel picks it up when opened', savedQueued: 'Saved on phone · sent when online',
    needAmount: 'Enter an amount', needAcc: 'Choose an account', needCat: 'Choose a category', needPerson: 'Enter a name', needItem: 'Enter a plan name',
    needTicker: 'Enter a ticker', sameAcc: 'From and to are the same account',
    bank: 'Bank', ewallet: 'E-wallet', cash: 'Cash', other: 'Other', totalCash: 'Total cash', netWorth: 'Net worth',
    iOwe: 'I owe', owesMe: 'Owed to me', people: n => `${n} ${n === 1 ? 'person' : 'people'}`, ongoing: 'Open', settled: 'Settled',
    pay: 'Pay', receive: 'Receive', debtNote: 'Pay / Receive records an installment to the chosen account.',
    planQ: 'Will my money be enough? · 12 months', enough: 'ENOUGH', short: 'SHORT',
    lowest: (v, m) => `Lowest balance ${v} in ${m}`, projection: 'Projected balance', upcoming: 'Upcoming', open: n => `${n} open`,
    noPlans: 'No plans yet.', holdings: 'Holdings', record: 'Record', invValue: 'Market value', pl: 'Profit/loss',
    sync: 'Sync with Excel', connected: 'Connected', offline: 'Offline', notSet: 'Not set up',
    lastSync: t => `Updated: ${t}`, waiting: n => `${n} entr${n === 1 ? 'y' : 'ies'} waiting for Excel`, queued: n => `${n} not sent yet`,
    syncNow: 'Sync now', language: 'Language · Bahasa', langNote: 'Applies to the whole app',
    settings: 'Connection', url: 'Web App URL (Apps Script)', token: 'Token', test: 'Save & test', demo: 'View demo',
    setupTitle: 'Connect to Excel', setupNote: 'Paste the Web App URL and token from the Apps Script step in the guide.',
    okConn: 'Connected', failConn: 'Could not connect: ', pending: 'waiting', queuedTag: 'on phone',
    debts: 'Debts', plans: 'Plans', budgets: 'Budget', invest: 'Investments', catAcc: 'Categories & accounts',
    catAccNote: 'Follows Settings in Excel', back: 'Back', photo: 'Receipt photo', photoOn: 'Photo attached',
    demoOn: 'Demo mode · sample data', snapshotOld: 'No Excel data yet. Open Excel once so it sends a snapshot.',
    simple: 'Estimate from current balance and open plans', borrow: 'Borrow', repay: 'Repay', lend: 'Lend', collect: 'Collect',
    buy: 'Buy', sell: 'Sell', dividend: 'Dividend', search: 'Search', months: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
    days: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'], disconnect: 'Disconnect', version: 'Version',
    edit: 'Edit', del: 'Delete', cancel: 'Cancel', delQ: 'Delete this transaction? Excel deletes it too.', deletedT: 'Deleted · Excel will update',
    editedT: 'Changes saved · Excel will update', noRef: 'This entry was typed in Excel without a Ref; edit it in Excel.', tapEdit: 'Tap a transaction to edit or delete it',
    setTitle: 'Settings', prefs: 'Preferences', alloc: 'Budget allocation', accountsT: 'Accounts', expCats: 'Expense categories', incCats: 'Income categories', subcats: 'Sub-categories',
    addAcc: '+ Add account', addCat: '+ Add category', addSub: '+ Add sub-category', name: 'Name', type: 'Type', opening: 'Opening balance', bucket: 'Budget bucket', fixed: 'Fixed bill?',
    yes: 'Yes', no: 'No', share: 'Share', total: 'Total', mustTotal: 'Total must be 100%', savedSet: 'Saved · Excel will update',
    renameNote: 'Older transactions keep the old name.', lastItem: 'At least one is needed.', exists: 'That name already exists.', needName: 'Enter a name',
    connection: 'Connection to Excel', emptyAuto: 'empty = automatic', color: 'Color', hexPh: 'or a colour code, e.g. #F15A23',
    pl_BasisMethod: 'Budget basis', pl_BasisManual: 'Manual basis amount', pl_OpenDate: 'Opening balance date', pl_DailyManual: 'Daily limit', pl_EmergMonths: 'Emergency fund (months)',
    pl_DebtOn: 'Debt moves balances', pl_FeeCat: 'Transfer fee category', pl_FeeSub: 'Transfer fee sub-category', pl_LocalInbox: 'Local inbox folder', pl_CloudInbox: 'Google Drive inbox folder',
    pl_Owner: 'Owner', pl_InvOn: 'Investments move balances', pl_SafeMargin: 'Safe-spend margin', pl_SyncFolder: 'Phone sync folder',
    bm: { 'Last Month': 'Last month', 'Average 3 Months': '3-month average', 'This Month': 'This month', 'Manual': 'Manual' }
  }
};
let lang = LS.get('lang', 'id');
const t = k => (T[lang][k] !== undefined ? T[lang][k] : T.id[k] !== undefined ? T.id[k] : k);

/* ---------------------------------------------------------------- icons */
const I = {
  home: '<path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/>',
  list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  wallet: '<rect x="3" y="6" width="18" height="14" rx="3"/><path d="M3 10h18M16 15h2"/>',
  more: '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>',
  sync: '<path d="M21 12a9 9 0 0 1-15.5 6.2M3 12A9 9 0 0 1 18.5 5.8M18 2v4h-4M6 22v-4h4"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>', back: '<path d="M15 5l-7 7 7 7"/>', chev: '<path d="M9 5l7 7-7 7"/>',
  cam: '<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
  swap: '<path d="M4 8h14l-3-3M20 16H6l3 3"/>', search: '<circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/>',
  left: '<path d="M15 6l-6 6 6 6"/>', right: '<path d="M9 6l6 6-6 6"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'
};
const ic = (n, sw = 2) => `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round">${I[n]}</svg>`;

/* ---------------------------------------------------------------- formatting */
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => { const n = Number(v); return isFinite(n) ? n : 0; };
const grp = n => Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const rp = (n, sign) => (sign && n > 0 ? '+' : n < 0 ? '−' : '') + 'Rp ' + grp(n);
const short = (n, sign) => (sign && n > 0 ? '+' : n < 0 ? '−' : '') + grp(n);
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const pdate = s => { if (!s) return null; const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const todayIso = () => iso(new Date());
const fdate = s => { const d = pdate(s); return d ? `${d.getDate()} ${t('months')[d.getMonth()]} ${d.getFullYear()}` : ''; };
const dayLabel = s => {
  const d = pdate(s); if (!d) return '';
  const td = new Date(); td.setHours(0, 0, 0, 0); const diff = Math.round((td - d) / 864e5);
  const base = `${t('days')[d.getDay()]}, ${d.getDate()} ${t('months')[d.getMonth()]}`;
  return diff === 0 ? `${t('today')} · ${base}` : diff === 1 ? `${t('yesterday')} · ${base}` : base;
};
const initials = s => (String(s || '?').replace(/[^A-Za-z0-9 &]/g, '').split(/[\s&]+/).filter(Boolean).map(w => w[0]).join('').slice(0, 2) || '?').toUpperCase();
const vib = () => { try { navigator.vibrate && navigator.vibrate(8); } catch (e) {} };

/* ---------------------------------------------------------------- state */
let cfg = LS.get('cfg', { url: '', token: '', demo: false });
let snap = LS.get('snap', null);               // snapshot from Excel
let pending = LS.get('pending', []);            // entries in the Drive inbox not yet read by Excel
let queue = LS.get('queue', []);                // entries not yet sent from this phone
let lastSync = LS.get('lastSync', null);
let online = navigator.onLine;
let syncing = false;

const TYPES = ['Expense', 'Income', 'Transfer', 'Debt', 'Plan'];
const DEBT_ACTS = ['Borrow', 'Repay', 'Lend', 'Collect'];
const INV_ACTS = ['Buy', 'Sell', 'Dividend'];

/* ---------------------------------------------------------------- sync */
async function api(method, body) {
  if (!cfg.url || !cfg.token) throw new Error(t('notSet'));
  const ctrl = new AbortController(); const tm = setTimeout(() => ctrl.abort(), 25000);
  try {
    const r = method === 'GET'
      ? await fetch(cfg.url + (cfg.url.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(cfg.token) + '&_=' + Date.now(), { signal: ctrl.signal })
      : await fetch(cfg.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ token: cfg.token }, body)), signal: ctrl.signal });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'error');
    return j;
  } finally { clearTimeout(tm); }
}

async function sync(manual) {
  if (cfg.demo) { if (manual) toast(t('demoOn')); return; }
  if (syncing || !cfg.url) return;
  syncing = true; setSyncIcon(true);
  try {
    if (queue.length) {
      const send = queue.slice();
      const photos = send.filter(r => r._photo).map(r => ({ ref: r.Ref, data: r._photo }));
      const rows = send.map(r => { const o = Object.assign({}, r); delete o._photo; delete o._at; return o; });
      await api('POST', { rows, photos });
      const sent = new Set(send.map(r => r.Ref));
      queue = queue.filter(r => !sent.has(r.Ref)); LS.set('queue', queue);
    }
    const j = await api('GET');
    if (j.snapshot) { snap = j.snapshot; LS.set('snap', snap); }
    pending = j.pending || []; LS.set('pending', pending);
    // a deletion is finished once its Delete row has been read by Excel (no longer waiting anywhere)
    const waitingDel = new Set(pending.concat(queue).filter(r => String(r.Type).toLowerCase() === 'delete').map(r => String(r.Description)));
    Object.keys(deleted).forEach(ref => { if (!waitingDel.has(ref)) delete deleted[ref]; }); LS.set('deleted', deleted);
    lastSync = new Date().toISOString(); LS.set('lastSync', lastSync);
    online = true;
    if (manual) toast(t('okConn'));
  } catch (e) {
    online = navigator.onLine && false;
    if (manual) toast(t('failConn') + e.message);
  } finally {
    syncing = false; setSyncIcon(false); rerender();
  }
}
function setSyncIcon(on) { document.querySelectorAll('[data-sync] svg').forEach(s => s.classList.toggle('spin', on)); }

/* ---------------------------------------------------------------- data model */
function accHex(name, raw) {  // same rules as Excel: palette name, #RRGGBB, or a colour from the account name
  raw = String(raw || '').trim();
  const hx = raw.replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(hx)) return '#' + hx.toUpperCase();
  const p = PALETTE.find(x => x[0].toLowerCase() === raw.toLowerCase()); if (p) return p[1];
  const a = String(name || '').toLowerCase();
  const rules = [['bca', '#005BAC'], ['bri', '#00529B'], ['bni', '#F15A23'], ['tapcash', '#F15A23'], ['mandiri', '#FFB700'], ['e-money', '#FFB700'], ['emoney', '#FFB700'],
    ['gojek', '#00AA13'], ['gopay', '#00AA13'], ['dana', '#118EEA'], ['shopee', '#EE4D2D'], ['ovo', '#4C3494'], ['cash', '#2E7D32'], ['tunai', '#2E7D32']];
  const m = rules.find(([k]) => a.includes(k)); return m ? m[1] : '#6B6B6B';
}
function inkOn(hex) { const n = parseInt(String(hex).slice(1), 16); const l = 0.2126 * (n >> 16) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255); return l > 150 ? '#262626' : '#FFFFFF'; }
let setOv = LS.get('setOv', null);              // settings saved on the phone, until Excel sends them back
let deleted = LS.get('deleted', {});            // ref -> entry deleted on the phone, until Excel removed it
function S() {
  const base = cfg.demo ? DEMO() : snap;
  if (!setOv || !base) return base;
  const gen = base.generated ? new Date(base.generated).getTime() : 0;
  const waiting = pending.concat(queue).some(r => String(r.Type).toLowerCase() === 'setting');
  if (!cfg.demo && gen > setOv.at && !waiting) { setOv = null; LS.set('setOv', null); return base; }
  const o = Object.assign({}, base), d = setOv.data;
  o.prefs = Object.assign({}, base.prefs || {}, d.prefs);
  const bal = {}; (base.accounts || []).forEach(a => bal[a.name] = num(a.balance));
  o.accounts = d.accounts.map(a => ({ name: a.name, type: a.type, opening: a.opening, colorName: a.colorName || '', color: accHex(a.name, a.colorName), balance: a.name in bal ? bal[a.name] : num(a.opening) }));
  o.buckets = (base.buckets || []).map((b, i) => Object.assign({}, b, d.buckets[i] ? { name: d.buckets[i].name, share: d.buckets[i].share } : {}));
  o.lists = { subcats: d.subs.map(x => ({ cat: x.cat, sub: x.sub, bucket: x.bucket, fixed: x.fixed })),
              expenseCats: d.expCats.map(c => ({ cat: c })), incomeCats: d.incCats.map(c => ({ cat: c })) };
  return o;
}
function rowsOf(type) {
  const s = S(); if (!s) return [];
  const key = { Expense: 'expense', Income: 'income', Transfer: 'transfer', Debt: 'debt', Investment: 'investment' }[type];
  return (s[key] || []);
}
// one list of everything: Excel rows + inbox rows + phone queue, with edits/deletes made on the phone applied on top
const clean = v => (v == null ? '' : String(v));
function excelRaw() { // Excel table rows in the same shape as an inbox row
  const out = [];
  rowsOf('Expense').forEach(r => out.push({ Ref: r.Ref, Type: 'Expense', Date: r.Date, Description: clean(r.Description), Amount: num(r.Amount), Category: clean(r.Category), SubCategory: clean(r['Sub-Category']), Account: clean(r.Account), ToAccount: '', Party: clean(r['Paid To']), Notes: clean(r.Notes), Fee: '' }));
  rowsOf('Income').forEach(r => out.push({ Ref: r.Ref, Type: 'Income', Date: r.Date, Description: clean(r.Description), Amount: num(r.Amount), Category: clean(r.Category), SubCategory: '', Account: clean(r.Account), ToAccount: '', Party: clean(r.From), Notes: clean(r.Notes), Fee: '' }));
  rowsOf('Transfer').forEach(r => out.push({ Ref: r.Ref, Type: 'Transfer', Date: r.Date, Description: '', Amount: num(r.Amount), Category: '', SubCategory: '', Account: clean(r.From), ToAccount: clean(r.To), Party: '', Notes: clean(r.Notes), Fee: num(r.Fee) || '' }));
  rowsOf('Debt').forEach(r => out.push({ Ref: r.Ref, Type: 'Debt', Date: r.Date, Description: clean(r.Description), Amount: num(r.Amount), Category: clean(r.Action), SubCategory: '', Account: clean(r.Account), ToAccount: '', Party: clean(r.Person), Notes: clean(r.Notes), Fee: '' }));
  return out;
}
function fromRow(r, st) {
  const ty = String(r.Type || '').toLowerCase(); const amt = num(r.Amount);
  if (!['expense', 'income', 'transfer', 'debt'].includes(ty)) return null;
  const T0 = ty.charAt(0).toUpperCase() + ty.slice(1);
  const base = { type: T0, date: r.Date, acc: r.Account, amt, ref: r.Ref || '', st, raw: Object.assign({}, r, { Type: T0 }) };
  if (T0 === 'Transfer') return Object.assign(base, { desc: `${r.Account} → ${r.ToAccount}`, cat: num(r.Fee) ? `${t('fee')} ${grp(r.Fee)}` : t('transfer'), acc2: r.ToAccount, fee: num(r.Fee) });
  if (T0 === 'Debt') return Object.assign(base, { desc: r.Description || `${actLabel(r.Category)} · ${r.Party}`, cat: `${actLabel(r.Category)} · ${r.Party}`, act: r.Category });
  return Object.assign(base, { desc: r.Description || r.SubCategory || r.Category, cat: r.SubCategory || r.Category });
}
function allTxRaw() {
  const out = [], seen = new Set();
  const push = o => { if (!o) return; if (o.ref && seen.has(o.ref)) return; if (o.ref) seen.add(o.ref); out.push(o); };
  excelRaw().forEach(r => push(fromRow(r, 'ok')));
  const excelRefs = new Set(out.map(o => o.ref).filter(Boolean));
  const later = pending.map(r => [r, 'pending']).concat(queue.map(r => [r, 'queued']));
  later.forEach(([r, st]) => { const ty = String(r.Type).toLowerCase(); if (!ty.startsWith('edit:') && !excelRefs.has(r.Ref)) push(fromRow(r, st)); });
  const del = delSet(); out.forEach(o => { o.del = !!(o.ref && del.has(o.ref)); });
  // edits: the old version stops counting, the new one counts until Excel has it
  later.forEach(([r, st]) => {
    const ty = String(r.Type).toLowerCase(); if (!ty.startsWith('edit:')) return;
    const nw = fromRow(Object.assign({}, r, { Type: r.Type.slice(5) }), st); if (!nw) return;
    const i = out.findIndex(o => o.ref === r.Ref && !o.del && !o.replaced);
    if (i >= 0) { const old = out[i]; if (old.st === 'ok') old.del = true; else out.splice(i, 1); }
    nw.replaced = false; nw.del = del.has(r.Ref) && !excelRefs.has(r.Ref) ? true : false;
    out.push(nw);
  });
  out.sort((a, b) => (b.date || '').localeCompare(a.date || '') || (a.st === 'ok') - (b.st === 'ok'));
  return out;
}
function delSet() {
  const d = new Set(Object.keys(deleted));
  pending.concat(queue).forEach(r => { if (String(r.Type).toLowerCase() === 'delete') d.add(String(r.Description)); });
  return d;
}
function allTx() { return allTxRaw().filter(o => !o.del); }
// entries Excel has not counted yet (they move balances on the phone right away)
function unseen() {
  const out = [];
  allTxRaw().forEach(o => {
    if (o.del && o.st === 'ok') out.push(Object.assign({}, o, { sign: -1 }));
    else if (!o.del && o.st !== 'ok') out.push(Object.assign({}, o, { sign: 1 }));
  });
  return out;
}
function effect(o) { // per-account change
  const d = {};
  const add = (a, v) => { if (a) d[a] = (d[a] || 0) + v; };
  if (o.type === 'Expense') add(o.acc, -o.amt);
  else if (o.type === 'Income') add(o.acc, o.amt);
  else if (o.type === 'Transfer') { add(o.acc, -o.amt - (o.fee || 0)); add(o.acc2, o.amt); }
  else if (o.type === 'Debt') add(o.acc, (o.act === 'Borrow' || o.act === 'Collect') ? o.amt : -o.amt);
  if (o.sign === -1) Object.keys(d).forEach(k => d[k] = -d[k]);
  return d;
}
function accounts() {
  const s = S(); const list = (s && s.accounts || []).map(a => ({ name: a.name, type: a.type || '', balance: num(a.balance), color: a.color || accHex(a.name, a.colorName) }));
  unseen().forEach(o => { const e = effect(o); Object.keys(e).forEach(k => { const a = list.find(x => x.name === k); if (a) a.balance += e[k]; }); });
  return list;
}
function kpi() {
  const s = S(); const k = Object.assign({}, s && s.kpi || {});
  const ms = todayIso().slice(0, 7), td = todayIso();
  let dBal = 0, dInc = 0, dExp = 0, dToday = 0;
  unseen().forEach(o => {
    const e = effect(o); Object.values(e).forEach(v => dBal += v);
    const sg = o.sign || 1;
    if ((o.date || '').slice(0, 7) === ms) { if (o.type === 'Income') dInc += sg * o.amt; if (o.type === 'Expense') dExp += sg * o.amt; }
    if (o.type === 'Expense' && o.date === td) dToday += sg * o.amt;
  });
  k.TotalBalance = num(k.TotalBalance) + dBal; k.IncThis = num(k.IncThis) + dInc; k.ExpThis = num(k.ExpThis) + dExp;
  k.SafeToday = num(k.SafeToday) - dToday; k.TodayUnpl = num(k.TodayUnpl) + dToday;
  return k;
}
function lists() {
  const s = S() || {}; const l = s.lists || {};
  const sub = (l.subcats || []).filter(x => x.cat);
  const expCats = (l.expenseCats || []).map(x => x.cat).filter(Boolean);
  const incCats = (l.incomeCats || []).map(x => x.cat).filter(Boolean);
  return { sub, expCats: expCats.length ? expCats : [...new Set(sub.map(x => x.cat))], incCats };
}
function openPlans() {
  const s = S() || {};
  const ps = (s.plans || []).filter(p => p.Item && num(p.Amount) > 0 && !/^(done|off)$/i.test(String(p.Status || '')))
    .map(p => ({ date: p.Date, item: p.Item, type: p.Type || 'Expense', cat: p.Category, amt: num(p.Amount), repeat: p.Repeat || 'Once', until: p.Until, st: 'ok' }));
  pending.concat(queue).filter(r => String(r.Type).toLowerCase() === 'plan')
    .forEach(r => ps.push({ date: r.Date, item: r.Description, type: r.SubCategory || 'Expense', cat: r.Category, amt: num(r.Amount), repeat: r.ToAccount || 'Once', until: '', st: 'pending' }));
  return ps.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
}
const actLabel = a => ({ Borrow: t('borrow'), Repay: t('repay'), Lend: t('lend'), Collect: t('collect'), Buy: t('buy'), Sell: t('sell'), Dividend: t('dividend') }[a] || a || '');
const typeColor = ty => ({ Expense: ['#FCE4DE', '#C0392B'], Income: ['#DDEBFA', '#1D63B8'], Transfer: ['#E8E8E3', '#404040'], Debt: ['#E6F2DC', '#3F7A1E'], Plan: ['#EEE6F7', '#5B3BA6'] }[ty] || ['#EEE', '#333']);

/* ---------------------------------------------------------------- router */
const tabs = ['home', 'tx', 'wallet', 'more'];
let stack = [LS.get('tab', 'home')];
let curView = null;
function go(name, opts = {}) {
  vib();
  const prev = stack[stack.length - 1];
  let dir = 'fade';
  if (opts.tab) { stack = [name]; dir = tabs.indexOf(name) >= tabs.indexOf(prev) ? 'right' : 'left'; LS.set('tab', name); }
  else if (opts.back) { stack.pop(); dir = 'left'; }
  else { stack.push(name); dir = 'right'; }
  render(dir);
}
function render(dir = 'fade') {
  const name = stack[stack.length - 1];
  const host = $('#views');
  const v = document.createElement('section');
  v.className = 'view enter-' + dir;
  REG = [];
  v.innerHTML = (VIEWS[name] || VIEWS.home)();
  host.appendChild(v);
  const old = curView; curView = v;
  requestAnimationFrame(() => requestAnimationFrame(() => {
    v.classList.remove('enter-' + dir);
    if (old) { old.classList.add(dir === 'left' ? 'leave-right' : dir === 'right' ? 'leave-left' : 'leave-fade'); setTimeout(() => old.remove(), 380); }
    animateIn(v);
  }));
  bind(v);
  renderNav();
}
function rerender() { // refresh the current view in place (no slide)
  if (!curView) return render();
  const name = stack[stack.length - 1];
  const y = curView.scrollTop;
  REG = [];
  curView.innerHTML = (VIEWS[name] || VIEWS.home)();
  curView.scrollTop = y; bind(curView); animateIn(curView, true); renderNav();
}
function animateIn(v, quiet) {
  v.querySelectorAll('.bar span[data-w]').forEach(s => requestAnimationFrame(() => { s.style.width = s.dataset.w; }));
  if (!quiet) v.querySelectorAll('.stack > *').forEach((el, i) => { el.classList.add('rise'); el.style.animationDelay = (i * 45) + 'ms'; });
  v.querySelectorAll('[data-count]').forEach(el => countUp(el, num(el.dataset.count), el.dataset.sign === '1'));
}
function countUp(el, to, sign) {
  const key = el.dataset.key; const from = key ? (countUp.mem[key] ?? to * 0.9) : 0;
  if (key) countUp.mem[key] = to;
  if (from === to || window.matchMedia('(prefers-reduced-motion: reduce)').matches) { el.textContent = rp(to, sign); return; }
  const t0 = performance.now(), dur = 650;
  const step = now => { const p = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - p, 3); el.textContent = rp(from + (to - from) * e, sign); if (p < 1) requestAnimationFrame(step); };
  requestAnimationFrame(step);
}
countUp.mem = {};

function renderNav() {
  const cur = stack[0];
  const a = (k, icon) => `<a href="#" data-tab="${k}" class="${cur === k ? 'on' : ''}">${ic(icon)}${t(k)}</a>`;
  $('#nav').innerHTML = a('home', 'home') + a('tx', 'list') +
    `<button class="fab press" data-add="Expense" aria-label="${t('add')}">${ic('plus', 2.4)}</button>` + a('wallet', 'wallet') + a('more', 'more');
  bind($('#nav'));
}

/* ---------------------------------------------------------------- views */
const head = (title, opts = {}) => `<header class="top">
  ${opts.back ? `<button class="ibtn press" data-back aria-label="${t('back')}">${ic('back')}</button>` : ''}
  <h1 class="h1" style="flex:1">${title}</h1>${opts.right || ''}</header>`;
const syncBtn = () => `<button class="ibtn press" data-sync aria-label="${t('syncNow')}">${ic('sync')}</button>`;
let REG = [];
const txRow = o => {
  const [bg, fg] = typeColor(o.type);
  const cls = o.type === 'Expense' ? 'out' : o.type === 'Income' ? 'in' : '';
  const sign = o.type === 'Expense' ? '−' : o.type === 'Income' ? '+' : '';
  const tag = o.st === 'pending' ? `<span class="tag">${t('pending')}</span>` : o.st === 'queued' ? `<span class="tag q">${t('queuedTag')}</span>` : '';
  REG.push(o);
  return `<button class="tx press" data-open="${REG.length - 1}"><div class="dot" style="background:${bg};color:${fg}">${esc(initials(o.cat))}</div>
    <div class="grow"><b>${esc(o.desc)}${tag}</b><div class="muted">${esc(o.cat || '')}${o.acc ? ' · ' + esc(o.acc) : ''}</div></div>
    <span class="amt ${cls}">${sign}${grp(o.amt)}</span></button>`;
};
const banner = () => {
  if (cfg.demo) return `<div class="card" style="background:#FFF3C4;padding:10px 14px"><span style="font-weight:600;font-size:13px;color:#6B5000">${t('demoOn')}</span></div>`;
  if (!S()) return `<div class="card" style="background:#FFF3C4"><span style="font-size:13px;color:#6B5000">${t('snapshotOld')}</span></div>`;
  return '';
};

const VIEWS = {
  setup() {
    return `${stack.length > 1 ? head(t('settings'), { back: true }) : ''}<div class="stack" style="padding-top:${stack.length > 1 ? 0 : 20}px">
      <div style="text-align:center;margin:10px 0 6px"><img src="icons/icon-192.png" alt="" width="72" height="72" style="border-radius:50%"><h1 class="h1 brand" style="font-size:30px;margin-top:8px"><span class="v">V</span>NANCIAL</h1></div>
      <section class="card stack" style="gap:12px">
        <div><p class="lab">${t('setupTitle')}</p><p class="muted" style="margin:4px 0 0">${t('setupNote')}</p></div>
        <label class="field"><span class="lab">${t('url')}</span><input class="input" id="s-url" inputmode="url" autocomplete="off" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(cfg.url)}"></label>
        <label class="field"><span class="lab">${t('token')}</span><input class="input" id="s-token" autocomplete="off" value="${esc(cfg.token)}"></label>
        <p class="err" id="s-err"></p>
        <button class="btn press" id="s-save">${t('test')}</button>
        <button class="btn dark press" id="s-demo">${t('demo')}</button>
      </section>
      ${langCard()}
    </div>`;
  },
  home() {
    const k = kpi(); const s = S();
    const spent = num(k.TodayUnpl), safe = num(k.SafeToday), lim = Math.max(spent + Math.max(safe, 0), 1);
    const used = spent / lim;
    const st = safe <= 0 && spent > 0 ? 'bad' : used > 0.8 ? 'warn' : 'ok';
    const stTxt = { ok: t('safe'), warn: t('careful'), bad: t('over') }[st];
    const bcol = { ok: 'var(--ok)', warn: 'var(--warn)', bad: 'var(--out)' }[st];
    const buckets = (s && s.buckets || []).filter(b => b.name && num(b.target) > 0);
    const bcols = ['#1D63B8', '#7A4FB5', '#D98E04', '#2E8B57', '#C0392B'];
    const recent = allTx().slice(0, 4);
    const d = new Date();
    return `${head(`<span class="brand"><span class="v">V</span>NANCIAL</span>`, { right: syncBtn() })}
    <p class="sub" style="margin:-10px 0 12px">${t('days')[d.getDay()]}, ${d.getDate()} ${t('months')[d.getMonth()]} ${d.getFullYear()}</p>
    <div class="stack">
      ${banner()}
      <section class="card dark press" data-go="wallet-tab">
        <p class="lab gold">${t('totalBalance')}</p>
        <p class="big" style="font-size:36px" data-count="${k.TotalBalance}" data-key="tb">${rp(k.TotalBalance)}</p>
        <div class="row" style="margin-top:10px">
          <div><p class="lab">${t('incMonth')}</p><p class="in" style="margin:3px 0 0;font-weight:700">${rp(k.IncThis, 1)}</p></div>
          <div style="text-align:right"><p class="lab">${t('expMonth')}</p><p class="out" style="margin:3px 0 0;font-weight:700">${rp(-k.ExpThis)}</p></div>
        </div>
      </section>
      <section class="card status-${st}">
        <div class="row"><p class="lab c-${st}">${t('safeToday')}</p><span class="c-${st}" style="font-size:12px;font-weight:700">${stTxt}</span></div>
        <p class="big" data-count="${Math.max(safe, 0)}" data-key="safe">${rp(Math.max(safe, 0))}</p>
        <div class="bar" style="margin-top:8px"><span data-w="${Math.min(100, used * 100).toFixed(1)}%" style="background:${bcol}"></span></div>
        <p class="muted" style="margin:7px 0 0">${lim > 1 ? t('usedOf')(rp(spent), rp(lim)) : '—'}</p>
      </section>
      ${buckets.length ? `<section class="card">
        <div class="row"><p class="lab">${t('budget')} ${esc(s.period || '')}</p><button class="link press" data-go="budget">${t('detail')}</button></div>
        <div class="stack" style="gap:10px;margin-top:4px">${buckets.slice(0, 4).map((b, i) => {
          const p = num(b.actual) / num(b.target);
          return `<div><div class="row" style="font-size:14px;margin-bottom:5px"><b>${esc(b.name)}</b><span class="muted">${Math.round(p * 100)}%</span></div>
          <div class="bar"><span data-w="${Math.min(100, p * 100).toFixed(1)}%" style="background:${p > 1 ? 'var(--out)' : bcols[i % 5]}"></span></div></div>`; }).join('')}</div>
      </section>` : ''}
      <section class="card">
        <div class="row"><p class="lab">${t('latest')}</p><button class="link press" data-go="tx-tab">${t('all')}</button></div>
        ${recent.length ? recent.map(txRow).join('') : `<p class="empty">${t('nothing')}</p>`}
      </section>
    </div>`;
  },
  tx() {
    const f = VIEWS.tx.f || 'all'; const m = VIEWS.tx.m || todayIso().slice(0, 7);
    const [yy, mm] = m.split('-').map(Number);
    let list = allTx().filter(o => (o.date || '').slice(0, 7) === m);
    const inc = list.filter(o => o.type === 'Income').reduce((a, o) => a + o.amt, 0);
    const exp = list.filter(o => o.type === 'Expense').reduce((a, o) => a + o.amt, 0);
    if (f !== 'all') list = list.filter(o => o.type === f);
    const q = (VIEWS.tx.q || '').toLowerCase();
    if (q) list = list.filter(o => (o.desc + ' ' + o.cat + ' ' + o.acc).toLowerCase().includes(q));
    let html = '', day = '';
    list.forEach(o => { if (o.date !== day) { day = o.date; html += `<p class="lab daylab">${esc(dayLabel(day))}</p>`; } html += txRow(o); });
    const chip = (k, l) => `<button class="chip press ${f === k ? 'on' : ''}" data-f="${k}">${l}</button>`;
    return `${head(t('tx'), { right: `<button class="ibtn press" data-search aria-label="${t('search')}">${ic('search')}</button>` })}
    <div class="stack">
      ${VIEWS.tx.searching ? `<input class="input" id="q" placeholder="${t('search')}…" value="${esc(VIEWS.tx.q || '')}" autocomplete="off">` : ''}
      <div class="chips scroll">${chip('all', t('all'))}${chip('Expense', t('expense'))}${chip('Income', t('income'))}${chip('Transfer', t('transfer'))}${chip('Debt', t('debt'))}</div>
      <div class="row"><button class="ibtn press" data-month="-1" aria-label="prev">${ic('left')}</button><b style="font-size:16px">${t('months')[mm - 1]} ${yy}</b><button class="ibtn press" data-month="1" aria-label="next">${ic('right')}</button></div>
      <div class="card row" style="padding:12px 16px"><div><p class="lab">${t('income')}</p><b class="in">${rp(inc)}</b></div><div><p class="lab">${t('expense')}</p><b class="out">${rp(exp)}</b></div><div style="text-align:right"><p class="lab">${t('net')}</p><b>${rp(inc - exp)}</b></div></div>
      <section class="card" style="padding-top:4px">${html || `<p class="empty">${t('nothing')}</p>`}</section>
    </div>`;
  },
  wallet() {
    const acc = accounts(); const k = kpi();
    const groups = [['Bank', t('bank')], ['E-Wallet', t('ewallet')], ['Cash', t('cash')]];
    const other = acc.filter(a => !groups.some(g => g[0].toLowerCase() === a.type.toLowerCase()));
    const row = a => `<div class="tx"><div class="dot" style="background:${a.color};color:${inkOn(a.color)}">${esc(initials(a.name))}</div><div class="grow"><b>${esc(a.name)}</b></div><span class="amt ${a.balance < 0 ? 'out' : ''}">${rp(a.balance)}</span></div>`;
    const sec = (lab, arr) => arr.length ? `<p class="lab daylab">${lab}</p>${arr.map(row).join('')}` : '';
    const total = acc.reduce((s, a) => s + a.balance, 0);
    return `${head(t('wallet'), { right: `<button class="ibtn press" data-add="Transfer" aria-label="${t('transfer')}">${ic('swap')}</button>` })}
    <div class="stack">
      ${banner()}
      <section class="card dark"><p class="lab gold">${t('totalCash')}</p><p class="big" data-count="${total}" data-key="tc">${rp(total)}</p>
        <div class="row" style="margin-top:8px;font-size:14px"><span style="color:#BDBDBD">${t('invest')}</span><b>${rp(num(k.InvValue))}</b></div>
        <div class="row" style="font-size:14px"><span style="color:#BDBDBD">${t('netWorth')}</span><b style="color:var(--gold)">${rp(total + num(k.InvValue) + num(k.RecvTotal) - num(k.DebtTotal))}</b></div></section>
      <section class="card" style="padding-top:4px">${groups.map(g => sec(g[1], acc.filter(a => a.type.toLowerCase() === g[0].toLowerCase()))).join('')}${sec(t('other'), other)}
      ${acc.length ? '' : `<p class="empty">—</p>`}</section>
    </div>`;
  },
  more() {
    const n = pending.length, q = queue.length;
    const st = cfg.demo ? ['warn', t('demoOn')] : !cfg.url ? ['bad', t('notSet')] : online ? ['ok', t('connected')] : ['warn', t('offline')];
    const lst = lastSync ? new Date(lastSync) : null;
    const item = (go, c, lab, sub) => `<button class="tx press" data-go="${go}" style="min-height:56px"><div class="dot" style="background:${c[0]};color:${c[1]}">${esc(initials(lab))}</div><div class="grow"><b>${lab}</b><div class="muted">${sub}</div></div><span style="width:20px;color:#8A8A8A">${ic('chev')}</span></button>`;
    const k = kpi(); const s = S() || {};
    const nd = (s.debts || []).filter(d => num(d.owesMe) || num(d.iOwe)).length;
    return `${head(t('more'), { right: `<button class="ibtn press" data-go="settings" aria-label="${t('setTitle')}">${ic('gear')}</button>` })}
    <div class="stack">
      <section class="card status-${st[0]}">
        <div class="row"><p class="lab c-${st[0]}">${t('sync')}</p><span class="c-${st[0]}" style="font-size:12px;font-weight:700">${st[1]}</span></div>
        <p style="margin:6px 0 2px;font-size:14px">${lst ? t('lastSync')(`${lst.getDate()} ${t('months')[lst.getMonth()]} ${String(lst.getHours()).padStart(2, '0')}.${String(lst.getMinutes()).padStart(2, '0')}`) : '—'}</p>
        ${n ? `<p class="muted" style="margin:0">${t('waiting')(n)}</p>` : ''}${q ? `<p class="muted" style="margin:0">${t('queued')(q)}</p>` : ''}
        <div class="chips" style="margin-top:10px"><button class="chip on press" data-sync>${t('syncNow')}</button><button class="chip press" data-go="setup">${t('settings')}</button></div>
      </section>
      <section class="card" style="padding-top:4px;padding-bottom:4px">
        ${item('debt', ['#E6F2DC', '#3F7A1E'], t('debts'), `${rp(num(k.DebtTotal))} · ${rp(num(k.RecvTotal))}`)}
        ${item('plan', ['#EEE6F7', '#5B3BA6'], t('plans'), t('planQ').split(' · ')[0])}
        ${item('invest', ['#FDEBDD', '#B4521A'], t('invest'), rp(num(k.InvValue)))}
        ${item('budget', ['#DDEBFA', '#1D63B8'], t('budgets'), (s.buckets || []).map(b => b.name).filter(Boolean).join(' · ') || '—')}
      </section>
      
      <p class="muted" style="text-align:center;margin:4px 0">${t('version')} ${APP_VER}${s.version ? ' · Excel ' + esc(s.version) : ''}</p>
    </div>`;
  },
  debt() {
    const s = S() || {};
    const rows = (s.debts || []).filter(d => d.person);
    const showSettled = VIEWS.debt.settled;
    const open = rows.filter(d => num(d.owesMe) > 0 || num(d.iOwe) > 0);
    const done = rows.filter(d => !(num(d.owesMe) > 0 || num(d.iOwe) > 0));
    const iOwe = open.reduce((a, d) => a + num(d.iOwe), 0), owes = open.reduce((a, d) => a + num(d.owesMe), 0);
    const list = showSettled ? done : open;
    const row = d => {
      const mine = num(d.iOwe) > 0;
      return `<div class="tx"><div class="dot" style="background:#EDEDE8;color:#404040">${esc(initials(d.person))}</div>
      <div class="grow"><b>${esc(d.person)}</b><div class="muted">${mine ? t('iOwe') : num(d.owesMe) > 0 ? t('owesMe') : t('settled')}</div></div>
      <div style="text-align:right"><span class="amt ${mine ? 'out' : 'in'}">${rp(mine ? num(d.iOwe) : num(d.owesMe))}</span>
      ${showSettled ? '' : `<div><button class="chip press" style="margin-top:5px;min-height:32px;padding:4px 12px;font-size:12px" data-add="Debt" data-act="${mine ? 'Repay' : 'Collect'}" data-person="${esc(d.person)}" data-amt="${mine ? num(d.iOwe) : num(d.owesMe)}">${mine ? t('pay') : t('receive')}</button></div>`}</div></div>`;
    };
    return `${head(t('debts'), { back: true, right: `<button class="ibtn press" data-add="Debt" aria-label="${t('add')}">${ic('plus')}</button>` })}
    <div class="stack">
      <div class="grid2"><section class="card"><p class="lab">${t('iOwe')}</p><p class="big out" style="font-size:24px">${rp(iOwe)}</p><p class="muted" style="margin:0">${t('people')(open.filter(d => num(d.iOwe) > 0).length)}</p></section>
      <section class="card"><p class="lab">${t('owesMe')}</p><p class="big in" style="font-size:24px">${rp(owes)}</p><p class="muted" style="margin:0">${t('people')(open.filter(d => num(d.owesMe) > 0).length)}</p></section></div>
      <div class="chips"><button class="chip press ${showSettled ? '' : 'on'}" data-settled="0">${t('ongoing')}</button><button class="chip press ${showSettled ? 'on' : ''}" data-settled="1">${t('settled')}</button></div>
      <section class="card" style="padding-top:4px">${list.length ? list.map(row).join('') : `<p class="empty">—</p>`}</section>
      <p class="muted" style="margin:0 4px">${t('debtNote')}</p>
    </div>`;
  },
  plan() {
    const k = kpi(); const plans = openPlans();
    const start = new Date(); start.setDate(1);
    const months = []; let bal = num(k.TotalBalance), low = bal, lowM = 0;
    for (let i = 0; i < 12; i++) {
      const m0 = new Date(start.getFullYear(), start.getMonth() + i, 1), key = iso(m0).slice(0, 7);
      let delta = 0;
      plans.forEach(p => {
        const d = pdate(p.date); if (!d) return;
        const pk = iso(d).slice(0, 7); const until = p.until ? String(p.until).slice(0, 7) : '9999-12';
        const hit = /month|bulan/i.test(p.repeat) ? (pk <= key && key <= until) : pk === key;
        if (hit && !(i === 0 && d < new Date(new Date().setHours(0, 0, 0, 0)) && !/month|bulan/i.test(p.repeat))) delta += (/income/i.test(p.type) ? 1 : -1) * p.amt;
      });
      bal += delta; months.push({ m0, bal });
      if (bal < low) { low = bal; lowM = i; }
    }
    const ok = low >= 0;
    const max = Math.max(...months.map(x => x.bal), 1), min = Math.min(...months.map(x => x.bal), 0);
    const y = v => 92 - ((v - min) / (max - min || 1)) * 80;
    const pts = months.map((x, i) => `${(i * 30).toFixed(0)},${y(x.bal).toFixed(1)}`).join(' ');
    const ml = d => `${t('months')[d.getMonth()]} ${d.getFullYear()}`;
    const upcoming = plans.filter(p => (p.date || '') >= todayIso() || /month|bulan/i.test(p.repeat));
    return `${head(t('plans'), { back: true, right: `<button class="ibtn press" data-add="Plan" aria-label="${t('add')}">${ic('plus')}</button>` })}
    <div class="stack">
      <section class="card dark" style="text-align:center">
        <p class="lab gold">${t('planQ')}</p>
        <p class="big" style="font-size:38px;letter-spacing:1px;color:${ok ? '#8FD19E' : '#FF9A8C'}">${ok ? t('enough') : t('short')}</p>
        <p style="margin:0;font-size:13px;color:#D0D0D0">${t('lowest')(rp(low), ml(months[lowM].m0))}</p>
      </section>
      <section class="card"><p class="lab">${t('projection')}</p>
        <svg role="img" aria-label="${t('projection')}" width="100%" height="118" viewBox="0 0 330 112" preserveAspectRatio="none" style="margin-top:6px">
          ${min < 0 ? `<line x1="0" y1="${y(0)}" x2="330" y2="${y(0)}" stroke="#C0392B" stroke-dasharray="5 5" stroke-width="1.5"/>` : ''}
          <polyline class="draw" fill="none" stroke="#1D63B8" stroke-width="3" stroke-linejoin="round" stroke-linecap="round" points="${pts}"/>
        </svg>
        <div class="row muted" style="font-size:11px"><span>${ml(months[0].m0)}</span><span>${ml(months[11].m0)}</span></div>
        <p class="muted" style="margin:6px 0 0;font-size:12px">${t('simple')}</p></section>
      <section class="card"><div class="row"><p class="lab">${t('upcoming')}</p><span class="muted">${t('open')(upcoming.length)}</span></div>
        ${upcoming.length ? upcoming.slice(0, 20).map(p => { const d = pdate(p.date); const inc = /income/i.test(p.type);
          return `<div class="tx"><div style="width:46px;text-align:center;flex:none"><b style="font-size:19px">${d ? d.getDate() : ''}</b><span class="muted" style="font-size:11px">${d ? t('months')[d.getMonth()] : ''}</span></div>
          <div class="grow"><b>${esc(p.item)}${p.st !== 'ok' ? `<span class="tag">${t('pending')}</span>` : ''}</b><div class="muted">${esc(p.cat || '')} · ${/month|bulan/i.test(p.repeat) ? t('monthly') : t('once')}</div></div>
          <span class="amt ${inc ? 'in' : 'out'}">${inc ? '+' : '−'}${grp(p.amt)}</span></div>`; }).join('') : `<p class="empty">${t('noPlans')}</p>`}
      </section>
    </div>`;
  },
  invest() {
    const s = S() || {}; const k = kpi();
    const h = (s.holdings || []).filter(x => x.ticker);
    return `${head(t('invest'), { back: true, right: `<button class="ibtn press" data-add="Investment" aria-label="${t('record')}">${ic('plus')}</button>` })}
    <div class="stack">
      <section class="card dark"><p class="lab gold">${t('invValue')}</p><p class="big">${rp(num(k.InvValue))}</p></section>
      <section class="card" style="padding-top:4px"><p class="lab daylab">${t('holdings')}</p>
        ${h.length ? h.map(x => `<div class="tx"><div class="dot" style="background:#FDEBDD;color:#B4521A">${esc(initials(x.ticker))}</div>
          <div class="grow"><b>${esc(x.ticker)}</b><div class="muted">${esc(x.name || x.type || '')}${num(x.units) ? ' · ' + grp(x.units) + ' unit' : ''}</div></div>
          <div style="text-align:right"><span class="amt">${rp(num(x.value))}</span><div class="muted ${num(x.pl) < 0 ? 'out' : 'in'}" style="font-weight:600">${rp(num(x.pl), 1)}</div></div></div>`).join('') : `<p class="empty">—</p>`}
      </section>
    </div>`;
  },
  budget() {
    const s = S() || {}; const b = (s.buckets || []).filter(x => x.name);
    const k = kpi();
    return `${head(t('budgets'), { back: true })}
    <div class="stack">
      <section class="card dark"><p class="lab gold">${esc(s.period || t('budget'))}</p><p class="big">${rp(num(k.BudgetLeft))}</p><p style="margin:2px 0 0;color:#BDBDBD;font-size:13px">${lang === 'id' ? 'sisa anggaran belanja' : 'spending budget left'} · ${num(k.DaysLeft)} ${lang === 'id' ? 'hari lagi' : 'days left'}</p></section>
      ${b.map(x => { const p = num(x.target) ? num(x.actual) / num(x.target) : 0;
        return `<section class="card"><div class="row"><b style="font-size:16px">${esc(x.name)}</b><span class="muted">${Math.round(num(x.share) * 100)}%</span></div>
        <div class="bar" style="margin:9px 0 7px"><span data-w="${Math.min(100, p * 100).toFixed(1)}%" style="background:${p > 1 ? 'var(--out)' : 'var(--in)'}"></span></div>
        <div class="row muted"><span>${rp(num(x.actual))}</span><span>${rp(num(x.target))}</span></div></section>`; }).join('') || `<p class="empty">—</p>`}
    </div>`;
  }
};
VIEWS.tx.f = 'all';
function langCard() {
  return `<section class="card"><div class="row"><div><p class="lab">${t('language')}</p><p class="muted" style="margin:2px 0 0">${t('langNote')}</p></div>
    <div class="switch" role="group" aria-label="${t('language')}"><button class="chip press ${lang === 'id' ? 'on' : ''}" data-lang="id">Indonesia</button><button class="chip press ${lang === 'en' ? 'on' : ''}" data-lang="en">English</button></div></div></section>`;
}

/* ---------------------------------------------------------------- events */
function bind(root) {
  root.querySelectorAll('[data-tab]').forEach(el => el.onclick = e => { e.preventDefault(); const k = el.dataset.tab; if (stack.length === 1 && stack[0] === k) { curView && curView.scrollTo({ top: 0, behavior: 'smooth' }); return; } go(k, { tab: true }); });
  root.querySelectorAll('[data-go]').forEach(el => el.onclick = () => { const g = el.dataset.go; if (g.endsWith('-tab')) go(g.slice(0, -4), { tab: true }); else go(g); });
  root.querySelectorAll('[data-back]').forEach(el => el.onclick = () => history.back());
  root.querySelectorAll('[data-sync]').forEach(el => el.onclick = () => { vib(); sync(true); });
  root.querySelectorAll('[data-open]').forEach(el => el.onclick = () => openEdit(REG[+el.dataset.open]));
  root.querySelectorAll('[data-set]').forEach(el => el.onclick = () => openSetting(el.dataset.set, el.dataset.i));
  root.querySelectorAll('[data-add]').forEach(el => el.onclick = () => openAdd(el.dataset.add, { act: el.dataset.act, person: el.dataset.person, amt: el.dataset.amt }));
  root.querySelectorAll('[data-lang]').forEach(el => el.onclick = () => { vib(); lang = el.dataset.lang; LS.set('lang', lang); document.documentElement.lang = lang; rerender(); renderNav(); });
  root.querySelectorAll('[data-f]').forEach(el => el.onclick = () => { vib(); VIEWS.tx.f = el.dataset.f; rerender(); });
  root.querySelectorAll('[data-month]').forEach(el => el.onclick = () => { vib(); const m = VIEWS.tx.m || todayIso().slice(0, 7); const d = new Date(+m.slice(0, 4), +m.slice(5) - 1 + Number(el.dataset.month), 1); VIEWS.tx.m = iso(d).slice(0, 7); rerender(); });
  root.querySelectorAll('[data-search]').forEach(el => el.onclick = () => { VIEWS.tx.searching = !VIEWS.tx.searching; if (!VIEWS.tx.searching) VIEWS.tx.q = ''; rerender(); const q = $('#q'); q && q.focus(); });
  const q = $('#q', root); if (q) q.oninput = () => { VIEWS.tx.q = q.value; clearTimeout(q._t); q._t = setTimeout(() => { const pos = q.selectionStart; rerender(); const n = $('#q'); if (n) { n.focus(); n.setSelectionRange(pos, pos); } }, 250); };
  root.querySelectorAll('[data-settled]').forEach(el => el.onclick = () => { vib(); VIEWS.debt.settled = el.dataset.settled === '1'; rerender(); });
  const sv = $('#s-save', root); if (sv) sv.onclick = async () => {
    const url = $('#s-url').value.trim(), tok = $('#s-token').value.trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) { $('#s-err').textContent = t('url') + '?'; return; }
    cfg = { url, token: tok, demo: false }; LS.set('cfg', cfg);
    sv.disabled = true; sv.textContent = '…';
    try { await api('GET'); LS.set('cfg', cfg); stack = ['home']; await sync(false); toast(t('okConn')); render('right'); }
    catch (e) { $('#s-err').textContent = t('failConn') + e.message; sv.disabled = false; sv.textContent = t('test'); }
  };
  const dm = $('#s-demo', root); if (dm) dm.onclick = () => { cfg = Object.assign({}, cfg, { demo: true }); LS.set('cfg', cfg); stack = ['home']; render('right'); };
}
// hardware back button closes sheets / pops sub-pages
window.addEventListener('popstate', () => {
  if ($('.mini.on')) { closeMini(); history.pushState({}, ''); return; }
  if ($('#sheet').classList.contains('on')) { closeAdd(); history.pushState({}, ''); return; }
  if (stack.length > 1) { go(null, { back: true }); history.pushState({}, ''); return; }
});

/* ---------------------------------------------------------------- toast */
function toast(msg) { const el = $('#toast'); el.textContent = msg; el.classList.add('on'); clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('on'), 2600); }

/* ---------------------------------------------------------------- add form (full-height sheet) */
let F = null; // form state
function openAdd(type, pre = {}) {
  vib();
  const accs = accounts().map(a => a.name);
  F = { type: type || 'Expense', amt: pre.amt ? Math.round(num(pre.amt)) : 0, desc: '', date: todayIso(), acc: LS.get('lastAcc', accs[0] || ''), acc2: '', cat: '', sub: '', party: pre.person || '', notes: '', plan: '',
        act: pre.act || 'Borrow', kind: 'Expense', repeat: 'Once', fee: 0, units: '', ticker: '', iact: 'Buy', photo: null };
  if (pre.edit) {
    const r = pre.edit.raw;
    Object.assign(F, { type: pre.edit.type, amt: num(r.Amount), desc: r.Description || '', date: r.Date || todayIso(), acc: r.Account || '', acc2: r.ToAccount || '',
      cat: pre.edit.type === 'Debt' ? '' : (r.Category || ''), sub: r.SubCategory || '', party: r.Party || '', notes: r.Notes || '', fee: num(r.Fee),
      act: pre.edit.type === 'Debt' ? (r.Category || 'Borrow') : 'Borrow', plan: pre.edit.type === 'Expense' ? (r.ToAccount || '') : '', acc2x: '',
      editRef: pre.edit.ref, editSt: pre.edit.st });
    if (pre.edit.type === 'Expense') F.acc2 = '';
  }
  drawAdd();
  $('#sheet').classList.add('on'); $('#sheet-backdrop').classList.add('on'); $('#nav').classList.add('hide');
  setTimeout(() => { const a = $('#f-amt'); if (a && !F.amt) a.focus(); }, 430);
}
function openEdit(o) {
  if (!o) return;
  if (!o.ref) { toast(t('noRef')); return; }
  openAdd(o.type, { edit: o });
}
function closeAdd() { $('#sheet').classList.remove('on'); $('#sheet-backdrop').classList.remove('on'); $('#nav').classList.remove('hide'); closeMini(); }
function drawAdd() {
  const L = lists(); const accs = accounts().map(a => a.name);
  const ty = F.type;
  const typeBtns = (F.editRef ? TYPES.filter(x => x !== 'Plan') : TYPES).map(x => `<button class="chip press ${ty === x ? 'on' : ''}" data-ty="${x}" style="padding:8px 4px;${ty === x && x === 'Expense' ? 'background:var(--out);border-color:var(--out)' : ty === x && x === 'Income' ? 'background:var(--in);border-color:var(--in)' : ''}">${t(x.toLowerCase())}</button>`).join('');
  const amtCls = ty === 'Expense' ? '' : ty === 'Income' ? 'in' : 'neutral';
  const k = kpi();
  const safeAfter = ty === 'Expense' && F.date === todayIso() ? `<p class="muted" style="margin:6px 0 0">${t('afterSafe')(rp(num(k.SafeToday) - F.amt))}</p>` : '';
  const sel = (id, lab, val, ph) => `<div class="field"><span class="lab">${lab}</span><button class="input press" data-pick="${id}" style="text-align:left;display:flex;align-items:center;justify-content:space-between"><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(val || ph || t('none'))}</span><span style="width:18px;color:#777">${ic('chev')}</span></button></div>`;
  const inp = (id, lab, val, attrs = '') => `<label class="field"><span class="lab">${lab}</span><input class="input" id="${id}" value="${esc(val)}" ${attrs}></label>`;
  const chips = (arr, cur, key) => `<div class="chips">${arr.map(c => `<button class="chip press ${cur === c ? 'on' : ''}" data-${key}="${esc(c)}">${esc(key === 'act' || key === 'iact' ? actLabel(c) : c)}</button>`).join('')}</div>`;
  let body = '';
  if (ty === 'Expense' || ty === 'Income') {
    const cats = ty === 'Expense' ? L.expCats : L.incCats;
    const subs = ty === 'Expense' ? L.sub.filter(s => s.cat === F.cat).map(s => s.sub) : [];
    const plans = openPlans().filter(p => !/income/i.test(p.type)).map(p => p.item);
    body = `${inp('f-desc', t('desc'), F.desc, 'autocomplete="off"')}
      <div class="grid2">${inp('f-date', t('date'), F.date, 'type="date"')}${sel('acc', t('account'), F.acc)}</div>
      <div class="field"><span class="lab">${t('category')}</span>${chips(cats, F.cat, 'cat')}</div>
      ${subs.length ? `<div class="field"><span class="lab">${t('subcat')}</span>${chips(subs, F.sub, 'sub')}</div>` : ''}
      ${inp('f-party', ty === 'Expense' ? t('paidTo') : t('from'), F.party, 'autocomplete="off"')}
      ${ty === 'Expense' && plans.length ? sel('plan', t('fromPlan'), F.plan) : ''}
      ${inp('f-notes', t('notes'), F.notes, 'autocomplete="off"')}`;
  } else if (ty === 'Transfer') {
    body = `<div class="grid2">${sel('acc', t('fromAcc'), F.acc)}${sel('acc2', t('toAcc'), F.acc2)}</div>
      <div class="grid2">${inp('f-date', t('date'), F.date, 'type="date"')}${inp('f-fee', t('fee'), F.fee ? grp(F.fee) : '', 'inputmode="numeric"')}</div>
      ${inp('f-notes', t('notes'), F.notes, 'autocomplete="off"')}`;
  } else if (ty === 'Debt') {
    const people = [...new Set(((S() || {}).debts || []).map(d => d.person).filter(Boolean))];
    body = `<div class="field"><span class="lab">${t('action')}</span>${chips(DEBT_ACTS, F.act, 'act')}</div>
      ${inp('f-party', t('person'), F.party, 'autocomplete="off" list="people"')}<datalist id="people">${people.map(p => `<option value="${esc(p)}">`).join('')}</datalist>
      <div class="grid2">${inp('f-date', t('date'), F.date, 'type="date"')}${sel('acc', t('account'), F.acc)}</div>
      ${inp('f-desc', t('desc'), F.desc, 'autocomplete="off"')}${inp('f-notes', t('notes'), F.notes, 'autocomplete="off"')}`;
  } else if (ty === 'Plan') {
    const cats = F.kind === 'Income' ? L.incCats : L.expCats;
    body = `${inp('f-desc', t('item'), F.desc, 'autocomplete="off"')}
      <div class="field"><span class="lab">${t('kind')}</span><div class="chips"><button class="chip press ${F.kind === 'Expense' ? 'on' : ''}" data-kind="Expense">${t('expense')}</button><button class="chip press ${F.kind === 'Income' ? 'on' : ''}" data-kind="Income">${t('income')}</button></div></div>
      <div class="grid2">${inp('f-date', t('date'), F.date, 'type="date"')}<div class="field"><span class="lab">${t('repeat')}</span><div class="chips"><button class="chip press ${F.repeat === 'Once' ? 'on' : ''}" data-repeat="Once">${t('once')}</button><button class="chip press ${F.repeat === 'Monthly' ? 'on' : ''}" data-repeat="Monthly">${t('monthly')}</button></div></div></div>
      <div class="field"><span class="lab">${t('category')}</span>${chips(cats, F.cat, 'cat')}</div>
      ${inp('f-notes', t('notes'), F.notes, 'autocomplete="off"')}`;
  } else if (ty === 'Investment') {
    body = `<div class="field"><span class="lab">${t('action')}</span>${chips(INV_ACTS, F.iact, 'iact')}</div>
      <div class="grid2">${inp('f-ticker', t('ticker'), F.ticker, 'autocomplete="off" style="text-transform:uppercase"')}${F.iact === 'Dividend' ? '' : inp('f-units', t('units'), F.units, 'inputmode="decimal"')}</div>
      <div class="grid2">${inp('f-date', t('date'), F.date, 'type="date"')}${sel('acc', t('account'), F.acc)}</div>
      ${inp('f-fee', t('fee'), F.fee ? grp(F.fee) : '', 'inputmode="numeric"')}`;
  }
  $('#sheet').innerHTML = `<div class="scroll"><div class="stack">
    <header class="top" style="margin-bottom:0"><button class="ibtn press" id="f-close" aria-label="${t('back')}">${ic('close', 2.2)}</button><h1 class="h1">${F.editRef ? t('edit') : ty === 'Investment' ? t('invest') : t('add')}</h1><span style="width:44px"></span></header>
    ${ty === 'Investment' ? '' : `<div class="seg" style="grid-template-columns:repeat(${F.editRef ? 4 : 5},1fr)">${typeBtns}</div>`}
    <section class="card" style="text-align:center;padding:16px">
      <p class="lab">${ty === 'Investment' ? t('price') : t('amount')}</p>
      <div style="display:flex;align-items:center;justify-content:center;gap:6px"><span class="amount ${amtCls}" style="width:auto;font-size:30px">Rp</span>
      <input class="amount ${amtCls}" id="f-amt" inputmode="numeric" autocomplete="off" placeholder="0" value="${F.amt ? grp(F.amt) : ''}" style="max-width:250px;text-align:left"></div>
      <span id="f-safe">${safeAfter}</span>
    </section>
    <section class="card stack" style="gap:13px">${body}</section>
    <p class="err" id="f-err" style="text-align:center"></p>
  </div></div>
  <div class="foot">
    ${F.editRef ? `<button class="ibtn press" id="f-del" style="width:56px;height:56px;color:var(--out)" aria-label="${t('del')}">${ic('trash')}</button>` : ''}
    ${ty === 'Expense' && !F.editRef ? `<button class="ibtn press" id="f-photo" style="width:56px;height:56px;${F.photo ? 'background:#E6F2DC;color:#22704A' : ''}" aria-label="${t('photo')}">${ic('cam')}</button><input type="file" id="f-file" accept="image/*" capture="environment" hidden>` : ''}
    <button class="btn press" id="f-save">${t('save')}</button>
  </div>`;
  bindAdd();
}
function readForm() {
  const v = id => { const el = $('#' + id); return el ? el.value : undefined; };
  if (v('f-desc') !== undefined) F.desc = v('f-desc').trim();
  if (v('f-date') !== undefined) F.date = v('f-date') || todayIso();
  if (v('f-party') !== undefined) F.party = v('f-party').trim();
  if (v('f-notes') !== undefined) F.notes = v('f-notes').trim();
  if (v('f-fee') !== undefined) F.fee = num(v('f-fee').replace(/\D/g, ''));
  if (v('f-ticker') !== undefined) F.ticker = v('f-ticker').trim().toUpperCase();
  if (v('f-units') !== undefined) F.units = v('f-units').replace(',', '.').trim();
}
function bindAdd() {
  const sh = $('#sheet');
  $('#f-close').onclick = () => { vib(); closeAdd(); };
  const a = $('#f-amt');
  a.oninput = () => {
    const digits = a.value.replace(/\D/g, '').slice(0, 13); F.amt = num(digits);
    a.value = digits ? grp(F.amt) : '';
    const fs = $('#f-safe'); if (fs && F.type === 'Expense' && F.date === todayIso()) fs.innerHTML = `<p class="muted" style="margin:6px 0 0">${t('afterSafe')(rp(num(kpi().SafeToday) - F.amt))}</p>`;
  };
  const fee = $('#f-fee'); if (fee) fee.oninput = () => { const d = fee.value.replace(/\D/g, ''); fee.value = d ? grp(num(d)) : ''; };
  const redraw = () => { readForm(); const y = $('.scroll', sh).scrollTop; drawAdd(); $('.scroll', sh).scrollTop = y; };
  sh.querySelectorAll('[data-ty]').forEach(el => el.onclick = () => { vib(); readForm(); F.type = el.dataset.ty; F.cat = ''; F.sub = ''; redraw(); });
  sh.querySelectorAll('[data-cat]').forEach(el => el.onclick = () => { vib(); readForm(); F.cat = F.cat === el.dataset.cat ? '' : el.dataset.cat; F.sub = ''; redraw(); });
  sh.querySelectorAll('[data-sub]').forEach(el => el.onclick = () => { vib(); readForm(); F.sub = F.sub === el.dataset.sub ? '' : el.dataset.sub; redraw(); });
  sh.querySelectorAll('[data-act]').forEach(el => el.onclick = () => { vib(); readForm(); F.act = el.dataset.act; redraw(); });
  sh.querySelectorAll('[data-iact]').forEach(el => el.onclick = () => { vib(); readForm(); F.iact = el.dataset.iact; redraw(); });
  sh.querySelectorAll('[data-kind]').forEach(el => el.onclick = () => { vib(); readForm(); F.kind = el.dataset.kind; F.cat = ''; redraw(); });
  sh.querySelectorAll('[data-repeat]').forEach(el => el.onclick = () => { vib(); readForm(); F.repeat = el.dataset.repeat; redraw(); });
  sh.querySelectorAll('[data-pick]').forEach(el => el.onclick = () => {
    vib(); readForm();
    const what = el.dataset.pick;
    const opts = what === 'plan' ? [''].concat(openPlans().filter(p => !/income/i.test(p.type)).map(p => p.item)) : accounts().map(a => a.name);
    openMini(opts, F[what], val => { F[what] = val; if (what === 'acc') LS.set('lastAcc', val); if (what === 'plan' && val) { const p = openPlans().find(x => x.item === val); if (p) { if (!F.amt) F.amt = p.amt; if (!F.desc) F.desc = p.item; if (!F.cat && p.cat) F.cat = p.cat; } } redraw(); });
  });
  const ph = $('#f-photo'); if (ph) { ph.onclick = () => $('#f-file').click(); $('#f-file').onchange = e => { const f = e.target.files[0]; if (f) shrink(f).then(b64 => { F.photo = b64; readForm(); drawAdd(); toast(t('photoOn')); }); }; }
  $('#f-save').onclick = save;
  const fd = $('#f-del'); if (fd) fd.onclick = () => openConfirm(t('delQ'), t('del'), () => { deleteEntry(F.editRef, F.editSt); closeAdd(); rerender(); toast(t('deletedT')); });
}
function shrink(file) {
  return new Promise(res => {
    const img = new Image(); const url = URL.createObjectURL(file);
    img.onload = () => { const s = Math.min(1, 1280 / Math.max(img.width, img.height)); const c = document.createElement('canvas'); c.width = img.width * s; c.height = img.height * s;
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url); res(c.toDataURL('image/jpeg', 0.75).split(',')[1]); };
    img.onerror = () => res(null); img.src = url;
  });
}
function save() {
  readForm();
  const err = m => { const e = $('#f-err'); e.textContent = m; e.animate && e.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-6px)' }, { transform: 'translateX(6px)' }, { transform: 'translateX(0)' }], { duration: 260 }); try { navigator.vibrate && navigator.vibrate([10, 40, 10]); } catch (x) {} };
  const ty = F.type;
  if (!(F.amt > 0)) return err(t('needAmount'));
  if (ty !== 'Plan' && !F.acc) return err(t('needAcc'));
  if ((ty === 'Expense' || ty === 'Income') && !F.cat) return err(t('needCat'));
  if (ty === 'Transfer' && (!F.acc2)) return err(t('needAcc'));
  if (ty === 'Transfer' && F.acc === F.acc2) return err(t('sameAcc'));
  if (ty === 'Debt' && !F.party) return err(t('needPerson'));
  if (ty === 'Plan' && !F.desc) return err(t('needItem'));
  if (ty === 'Plan' && !F.cat) return err(t('needCat'));
  if (ty === 'Investment' && !F.ticker) return err(t('needTicker'));
  const ref = F.editRef || newRef();
  const r = { Ref: ref, Type: ty, Date: F.date, Description: F.desc, Amount: F.amt, Category: '', SubCategory: '', Account: F.acc, ToAccount: '', Party: F.party, Notes: F.notes, Fee: '' };
  if (ty === 'Expense') { r.Category = F.cat; r.SubCategory = F.sub; r.ToAccount = F.plan || ''; if (!r.Description) r.Description = F.sub || F.cat; }
  if (ty === 'Income') { r.Category = F.cat; if (!r.Description) r.Description = F.cat; }
  if (ty === 'Transfer') { r.ToAccount = F.acc2; r.Fee = F.fee || ''; r.Party = ''; }
  if (ty === 'Debt') { r.Category = F.act; }
  if (ty === 'Plan') { r.Category = F.cat; r.SubCategory = F.kind; r.ToAccount = F.repeat; r.Account = ''; r.Party = ''; }
  if (ty === 'Investment') { r.Description = F.ticker; r.Category = F.iact; r.SubCategory = F.iact === 'Dividend' ? '' : F.units; r.Fee = F.fee || ''; r.Party = ''; }
  if (F.photo) { r._photo = F.photo; r.Notes = (r.Notes ? r.Notes + ' · ' : '') + 'foto: Vnancial Sync/Struk/' + ref + '.jpg'; }
  r._at = Date.now();
  if (F.editRef) {
    // still on this phone: change it in place; otherwise send an Edit (Excel replaces the row with the same Ref)
    const qi = queue.findIndex(q => q.Ref === ref && !/^(edit:|delete|setting)/i.test(String(q.Type)));
    queue = queue.filter(q => !(q.Ref === ref && /^edit:/i.test(String(q.Type))));
    if (F.editSt === 'queued' && qi >= 0) { const at = queue.findIndex(q => q.Ref === ref); queue[at] = r; }
    else { r.Type = 'Edit:' + ty; queue.push(r); }
  } else queue.push(r);
  LS.set('queue', queue);
  const wasEdit = !!F.editRef;
  const b = $('#f-save'); b.textContent = '✓'; b.style.background = '#8FD19E';
  try { navigator.vibrate && navigator.vibrate(18); } catch (x) {}
  setTimeout(() => {
    closeAdd(); rerender();
    if (cfg.demo) { toast(wasEdit ? t('editedT') : t('savedQueued')); return; }
    if (navigator.onLine && cfg.url) sync(false).then(() => toast(queue.some(x => x.Ref === ref) ? t('savedQueued') : wasEdit ? t('editedT') : t('savedSent')));
    else toast(t('savedQueued'));
  }, 260);
}


function newRef() { return 'HP-' + iso(new Date()).replace(/-/g, '').slice(2) + new Date().toTimeString().slice(0, 8).replace(/:/g, '') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase(); }
function cmdRow(type, extra) {
  return Object.assign({ Ref: newRef(), Type: type, Date: todayIso(), Description: '', Amount: 1, Category: '', SubCategory: '', Account: '', ToAccount: '', Party: '', Notes: '', Fee: '', _at: Date.now() }, extra);
}
function deleteEntry(ref, st) {
  vib();
  const own = queue.findIndex(q => q.Ref === ref && !/^(edit:|delete|setting)/i.test(String(q.Type)));
  queue = queue.filter(q => !(q.Ref === ref && /^edit:/i.test(String(q.Type))));
  if (st === 'queued' && own >= 0) queue = queue.filter(q => q.Ref !== ref);          // never left the phone
  else { queue.push(cmdRow('Delete', { Description: ref })); deleted[ref] = Date.now(); LS.set('deleted', deleted); }
  LS.set('queue', queue);
  if (!cfg.demo && navigator.onLine && cfg.url) sync(false);
}
function openConfirm(msg, okLabel, cb) {
  openPanel(`<p style="margin:4px 4px 16px;font-size:16px;font-weight:600;line-height:1.4">${esc(msg)}</p>
    <div class="grid2"><button class="btn dark press" id="cf-no" style="background:#EDEDE8;color:var(--ink)">${t('cancel')}</button>
    <button class="btn press" id="cf-ok" style="background:var(--out);color:#fff">${esc(okLabel)}</button></div>`, m => {
    $('#cf-no', m).onclick = () => { vib(); closeMini(); };
    $('#cf-ok', m).onclick = () => { closeMini(); cb(); };
  });
}
function openPanel(html, bindFn) {
  let m = $('#mini'); if (!m) { m = document.createElement('div'); m.id = 'mini'; m.className = 'mini'; document.body.appendChild(m); }
  m.innerHTML = `<div class="grab"></div>` + html;
  const bd = $('#mini-bd') || (() => { const d = document.createElement('div'); d.id = 'mini-bd'; d.className = 'backdrop'; d.style.zIndex = 22; d.onclick = closeMini; document.body.appendChild(d); return d; })();
  bindFn && bindFn(m);
  requestAnimationFrame(() => { m.classList.add('on'); bd.classList.add('on'); });
}

/* ---------------------------------------------------------------- settings (same as the Settings sheet in Excel) */
const PREFS = [['BasisMethod', 'select', ['Last Month', 'Average 3 Months', 'This Month', 'Manual']], ['BasisManual', 'num'], ['OpenDate', 'date'], ['DailyManual', 'num'],
  ['EmergMonths', 'num'], ['DebtOn', 'yesno'], ['InvOn', 'yesno'], ['SafeMargin', 'pct'], ['FeeCat', 'expcat'], ['FeeSub', 'feesub'], ['Owner', 'text'],
  ['LocalInbox', 'text'], ['CloudInbox', 'text'], ['SyncFolder', 'text']];
const ACC_TYPES = ['Cash', 'Bank', 'E-Wallet', 'Other'];
const PALETTE = [['BCA Blue', '#005BAC'], ['BRI Blue', '#00529B'], ['BNI Orange', '#F15A23'], ['Mandiri Yellow', '#FFB700'], ['Mandiri Navy', '#003D79'],
  ['GoPay Green', '#00AA13'], ['DANA Blue', '#118EEA'], ['Shopee Orange', '#EE4D2D'], ['OVO Purple', '#4C3494'], ['Cash Green', '#2E7D32'],
  ['Red', '#C0392B'], ['Teal', '#00897B'], ['Gray', '#6B6B6B'], ['Black', '#262626']];
function curSettings() {
  const s = S() || {}; const l = s.lists || {};
  return {
    prefs: Object.assign({}, s.prefs || {}),
    accounts: (s.accounts || []).filter(a => a.name).map(a => ({ name: a.name, type: a.type || 'Other', opening: num(a.opening), colorName: a.colorName || '' })),
    buckets: (s.buckets || []).filter(b => b.name).map(b => ({ name: b.name, share: num(b.share) })),
    subs: (l.subcats || []).filter(x => x.cat && x.sub).map(x => ({ cat: x.cat, sub: x.sub, bucket: x.bucket || '', fixed: x.fixed || 'No' })),
    expCats: (l.expenseCats || []).map(x => x.cat).filter(Boolean),
    incCats: (l.incomeCats || []).map(x => x.cat).filter(Boolean)
  };
}
function fmtPref(k, v, type) {
  if (v === null || v === undefined || v === '') return type === 'num' && k === 'DailyManual' ? t('emptyAuto') : '—';
  if (type === 'num') return grp(num(v));
  if (type === 'pct') return Math.round(num(v) * 100) + '%';
  if (type === 'date') return fdate(String(v));
  if (type === 'yesno') return v === 'Yes' ? t('yes') : t('no');
  if (type === 'select') return (t('bm')[v]) || v;
  return String(v);
}
// send changed sections to Excel; the phone shows them at once
function commitSettings(d, sections) {
  const rows = [];
  sections.forEach(sec => {
    if (sec === 'pref') Object.keys(d._prefChanged || {}).forEach(k => rows.push(cmdRow('Setting', { Description: 'pref', Category: k, SubCategory: d.prefs[k] == null ? '' : String(d.prefs[k]) })));
    if (sec === 'bucket') d.buckets.forEach((b, i) => rows.push(cmdRow('Setting', { Description: 'bucket', Party: String(i + 1), SubCategory: String(b.share) })));
    if (sec === 'account') d.accounts.forEach((a, i) => rows.push(cmdRow('Setting', { Description: 'account', Party: String(i + 1), Category: a.name, SubCategory: a.type, Account: String(a.opening || 0), ToAccount: a.colorName || '' })));
    if (sec === 'expcat') d.expCats.forEach((c, i) => rows.push(cmdRow('Setting', { Description: 'expcat', Party: String(i + 1), Category: c })));
    if (sec === 'inccat') d.incCats.forEach((c, i) => rows.push(cmdRow('Setting', { Description: 'inccat', Party: String(i + 1), Category: c })));
    if (sec === 'sub') {
      const order = []; d.expCats.forEach(c => d.subs.filter(x => x.cat === c).forEach(x => order.push(x)));
      d.subs.filter(x => !d.expCats.includes(x.cat)).forEach(x => order.push(x));          // keep sub-categories grouped by category
      d.subs = order;
      order.forEach((x, i) => rows.push(cmdRow('Setting', { Description: 'sub', Party: String(i + 1), Category: x.cat, SubCategory: x.sub, Account: x.bucket, ToAccount: x.fixed })));
    }
  });
  delete d._prefChanged;
  setOv = { at: Date.now(), data: d }; LS.set('setOv', setOv);
  if (cfg.demo) { DEMO.c = null; }
  queue.push(...rows); LS.set('queue', queue);
  rerender(); toast(t('savedSet'));
  if (!cfg.demo && navigator.onLine && cfg.url) sync(false);
}
// one small form in the bottom panel
function openForm(title, fields, onSave, onDelete, note) {
  const f = fields.map((x, i) => {
    const id = 'pf' + i;
    if (x.type === 'color') {
      const cur = String(x.value || ''); const isHex = /^#?[0-9a-f]{6}$/i.test(cur);
      return `<div class="field"><span class="lab">${esc(x.label)}</span><div class="chips" data-chips="${id}" data-color="1">${PALETTE.map(([n, h]) =>
        `<button class="chip press swatch ${cur.toLowerCase() === n.toLowerCase() ? 'on' : ''}" data-v="${esc(n)}" aria-label="${esc(n)}" title="${esc(n)}" style="background:${h};border-color:${h}"></button>`).join('')}</div>
        <input class="input" id="${id}-hex" placeholder="${t('hexPh')}" autocomplete="off" value="${isHex ? esc(cur.startsWith('#') ? cur : '#' + cur) : ''}" style="margin-top:6px"></div>`;
    }
    if (x.type === 'select' || x.type === 'yesno') {
      const opts = x.type === 'yesno' ? [['Yes', t('yes')], ['No', t('no')]] : x.options.map(o => Array.isArray(o) ? o : [o, o]);
      return `<div class="field"><span class="lab">${esc(x.label)}</span><div class="chips" data-chips="${id}">${opts.map(([v, l]) => `<button class="chip press ${String(x.value) === String(v) ? 'on' : ''}" data-v="${esc(v)}">${esc(l)}</button>`).join('')}</div></div>`;
    }
    const val = x.type === 'pct' ? (x.value === '' || x.value == null ? '' : Math.round(num(x.value) * 1000) / 10) : x.type === 'num' ? (x.value === '' || x.value == null ? '' : grp(num(x.value))) : (x.value == null ? '' : x.value);
    const attrs = x.type === 'num' ? 'inputmode="numeric"' : x.type === 'pct' ? 'inputmode="decimal"' : x.type === 'date' ? 'type="date"' : 'autocomplete="off"';
    return `<label class="field"><span class="lab">${esc(x.label)}${x.type === 'pct' ? ' (%)' : ''}</span><input class="input" id="${id}" ${attrs} value="${esc(val)}"></label>`;
  }).join('');
  openPanel(`<h2 class="h1" style="font-size:21px;margin:0 2px 12px">${esc(title)}</h2><div class="stack" style="gap:12px">${f}
    ${note ? `<p class="muted" style="margin:0">${esc(note)}</p>` : ''}<p class="err" id="pf-err"></p>
    <div style="display:flex;gap:10px">${onDelete ? `<button class="ibtn press" id="pf-del" style="width:56px;height:56px;background:#FCE4DE;color:var(--out)" aria-label="${t('del')}">${ic('trash')}</button>` : ''}
    <button class="btn press" id="pf-save">${t('save')}</button></div></div>`, m => {
    m.querySelectorAll('[data-chips]').forEach(g => g.querySelectorAll('.chip').forEach(b => b.onclick = () => { vib(); g.querySelectorAll('.chip').forEach(c => c.classList.remove('on')); b.classList.add('on'); if (g.dataset.color) { const hx = $('#' + g.dataset.chips + '-hex', m); if (hx) hx.value = ''; } }));
    m.querySelectorAll('input[inputmode="numeric"]').forEach(inp => inp.oninput = () => { const dg = inp.value.replace(/\D/g, ''); inp.value = dg ? grp(num(dg)) : ''; });
    $('#pf-save', m).onclick = () => {
      const vals = fields.map((x, i) => {
        const id = 'pf' + i;
        if (x.type === 'color') { const hx = $('#' + id + '-hex', m).value.trim().replace(/^#/, ''); if (/^[0-9a-f]{6}$/i.test(hx)) return '#' + hx.toUpperCase(); const on = m.querySelector(`[data-chips="${id}"] .chip.on`); return on ? on.dataset.v : ''; }
        if (x.type === 'select' || x.type === 'yesno') { const on = m.querySelector(`[data-chips="${id}"] .chip.on`); return on ? on.dataset.v : ''; }
        const v = $('#' + id, m).value.trim();
        if (x.type === 'num') return v === '' ? '' : num(v.replace(/\D/g, ''));
        if (x.type === 'pct') return v === '' ? '' : num(v.replace(',', '.')) / 100;
        return v;
      });
      const e = onSave(vals);
      if (e) { $('#pf-err', m).textContent = e; return; }
      vib(); closeMini();
    };
    const dl = $('#pf-del', m); if (dl) dl.onclick = () => { const e = onDelete(); if (e) { $('#pf-err', m).textContent = e; return; } vib(); closeMini(); };
  });
}
function openSetting(kind, i) {
  const d = curSettings(); i = i === undefined ? -1 : +i;
  if (kind === 'pref') {
    const [k, type, opts] = PREFS[i];
    let options = opts, ty = type;
    if (type === 'expcat') { ty = 'select'; options = d.expCats; }
    if (type === 'feesub') { ty = 'select'; options = d.subs.filter(x => x.cat === d.prefs.FeeCat).map(x => x.sub); if (!options.length) ty = 'text'; }
    if (type === 'select' && k === 'BasisMethod') options = opts.map(o => [o, t('bm')[o]]);
    openForm(t('pl_' + k), [{ label: t('pl_' + k), type: ty, value: d.prefs[k], options }], ([v]) => {
      if (ty === 'pct' && v !== '' && (v <= 0 || v > 1)) return '1–100%';
      d.prefs[k] = v === '' ? null : v; d._prefChanged = { [k]: 1 }; commitSettings(d, ['pref']);
    }, null, k === 'DailyManual' ? t('emptyAuto') : '');
  } else if (kind === 'bucket') {
    const b = d.buckets[i];
    openForm(b.name, [{ label: t('share'), type: 'pct', value: b.share }], ([v]) => { b.share = num(v); commitSettings(d, ['bucket']); },
      null, `${t('total')}: ${Math.round(d.buckets.reduce((a, x, j) => a + (j === i ? 0 : x.share), 0) * 100)}% + ${b.name}`);
  } else if (kind === 'account') {
    const a = i >= 0 ? d.accounts[i] : { name: '', type: 'Bank', opening: 0, colorName: '' };
    openForm(i >= 0 ? a.name : t('addAcc').replace('+ ', ''), [{ label: t('name'), type: 'text', value: a.name }, { label: t('type'), type: 'select', value: a.type, options: ACC_TYPES.map(x => [x, x === 'Cash' ? t('cash') : x === 'Bank' ? t('bank') : x === 'E-Wallet' ? t('ewallet') : t('other')]) },
      { label: t('opening'), type: 'num', value: a.opening }, { label: t('color'), type: 'color', value: a.colorName }], ([n, ty, op, col]) => {
        if (!n) return t('needName');
        if (d.accounts.some((x, j) => j !== i && x.name.toLowerCase() === n.toLowerCase())) return t('exists');
        const nv = { name: n, type: ty || 'Other', opening: num(op), colorName: col || '' };
        if (i >= 0) d.accounts[i] = nv; else d.accounts.push(nv);
        commitSettings(d, ['account']);
      }, i >= 0 ? () => { if (d.accounts.length <= 1) return t('lastItem'); d.accounts.splice(i, 1); commitSettings(d, ['account']); } : null, i >= 0 ? t('renameNote') : '');
  } else if (kind === 'expcat' || kind === 'inccat') {
    const arr = kind === 'expcat' ? d.expCats : d.incCats; const old = i >= 0 ? arr[i] : '';
    openForm(i >= 0 ? old : t('addCat').replace('+ ', ''), [{ label: t('name'), type: 'text', value: old }], ([n]) => {
      if (!n) return t('needName');
      if (arr.some((x, j) => j !== i && x.toLowerCase() === n.toLowerCase())) return t('exists');
      const secs = [kind];
      if (i >= 0) { arr[i] = n; if (kind === 'expcat' && old !== n) { d.subs.forEach(x => { if (x.cat === old) x.cat = n; }); secs.push('sub'); if (d.prefs.FeeCat === old) { d.prefs.FeeCat = n; d._prefChanged = { FeeCat: 1 }; secs.push('pref'); } } }
      else arr.push(n);
      commitSettings(d, secs);
    }, i >= 0 ? () => { if (arr.length <= 1) return t('lastItem'); arr.splice(i, 1); const secs = [kind];
      if (kind === 'expcat') { const n0 = d.subs.length; d.subs = d.subs.filter(x => x.cat !== old); if (d.subs.length !== n0) secs.push('sub'); }
      commitSettings(d, secs); } : null, i >= 0 ? t('renameNote') : '');
  } else if (kind === 'sub') {
    const x = i >= 0 ? d.subs[i] : { cat: d.expCats[0] || '', sub: '', bucket: (d.buckets[0] || {}).name || 'Needs', fixed: 'No' };
    openForm(i >= 0 ? x.sub : t('addSub').replace('+ ', ''), [{ label: t('category'), type: 'select', value: x.cat, options: d.expCats }, { label: t('name'), type: 'text', value: x.sub },
      { label: t('bucket'), type: 'select', value: x.bucket, options: d.buckets.map(b => b.name) }, { label: t('fixed'), type: 'yesno', value: x.fixed }], ([c, n, b, f]) => {
        if (!n) return t('needName'); if (!c) return t('needCat');
        if (d.subs.some((y, j) => j !== i && y.cat === c && y.sub.toLowerCase() === n.toLowerCase())) return t('exists');
        const nv = { cat: c, sub: n, bucket: b, fixed: f || 'No' };
        if (i >= 0) d.subs[i] = nv; else d.subs.push(nv);
        commitSettings(d, ['sub']);
      }, i >= 0 ? () => { if (d.subs.length <= 1) return t('lastItem'); d.subs.splice(i, 1); commitSettings(d, ['sub']); } : null, i >= 0 ? t('renameNote') : '');
  }
}
VIEWS.settings = function () {
  const d = curSettings();
  const row = (kind, i, title, sub, right, dot) => `<button class="tx press" data-set="${kind}" ${i !== undefined ? `data-i="${i}"` : ''} style="min-height:52px">
    ${dot ? `<span class="dot" style="width:28px;height:28px;border-radius:9px;background:${dot}"></span>` : ''}<div class="grow"><b>${esc(title)}</b>${sub ? `<div class="muted">${esc(sub)}</div>` : ''}</div>${right ? `<span class="amt" style="font-size:15px">${esc(right)}</span>` : ''}<span style="width:18px;color:#8A8A8A;flex:none">${ic('chev')}</span></button>`;
  const addBtn = (kind, label) => `<button class="link press" data-set="${kind}" style="padding:12px 0 4px;color:var(--orange)">${label}</button>`;
  const tot = Math.round(d.buckets.reduce((a, b) => a + b.share, 0) * 100);
  const subsBy = d.expCats.map(c => [c, d.subs.map((x, i) => [x, i]).filter(([x]) => x.cat === c)]).filter(([, a]) => a.length);
  return `${head(t('setTitle'), { back: true })}
  <div class="stack">
    ${langCard()}
    <section class="card" style="padding-top:4px;padding-bottom:4px"><button class="tx press" data-go="setup" style="min-height:52px"><div class="grow"><b>${t('connection')}</b><div class="muted">${cfg.demo ? t('demoOn') : cfg.url ? t('connected') : t('notSet')}</div></div><span style="width:18px;color:#8A8A8A">${ic('chev')}</span></button></section>
    ${S() ? `
    <section class="card" style="padding-top:4px"><p class="lab daylab">${t('prefs')}</p>${PREFS.map(([k, type], i) => row('pref', i, t('pl_' + k), '', fmtPref(k, d.prefs[k], type === 'expcat' || type === 'feesub' ? 'text' : type))).join('')}</section>
    <section class="card" style="padding-top:4px"><div class="row"><p class="lab daylab">${t('alloc')}</p><span class="${tot === 100 ? 'c-ok' : 'c-bad'}" style="font-size:12px;font-weight:700;padding-top:12px">${t('total')} ${tot}%${tot === 100 ? '' : ' · ' + t('mustTotal')}</span></div>
      ${d.buckets.map((b, i) => row('bucket', i, b.name, '', Math.round(b.share * 100) + '%')).join('')}</section>
    <section class="card" style="padding-top:4px"><p class="lab daylab">${t('accountsT')}</p>${d.accounts.map((a, i) => row('account', i, a.name, `${a.type} · ${t('opening').toLowerCase()} ${rp(a.opening)}`, '', accHex(a.name, a.colorName))).join('')}${addBtn('account', t('addAcc'))}</section>
    <section class="card" style="padding-top:4px"><p class="lab daylab">${t('expCats')}</p>${d.expCats.map((c, i) => row('expcat', i, c)).join('')}${addBtn('expcat', t('addCat'))}</section>
    <section class="card" style="padding-top:4px"><p class="lab daylab">${t('incCats')}</p>${d.incCats.map((c, i) => row('inccat', i, c)).join('')}${addBtn('inccat', t('addCat'))}</section>
    <section class="card" style="padding-top:4px"><p class="lab daylab">${t('subcats')}</p>
      ${subsBy.map(([c, arr]) => `<p class="lab" style="padding-top:12px;color:var(--ink)">${esc(c)}</p>${arr.map(([x, i]) => row('sub', i, x.sub, `${x.bucket}${x.fixed === 'Yes' ? ' · ' + t('fixed').replace('?', '') : ''}`)).join('')}`).join('')}
      ${addBtn('sub', t('addSub'))}</section>` : banner()}
  </div>`;
};

/* ---------------------------------------------------------------- picker (small bottom sheet) */
function openMini(opts, cur, cb) {
  let m = $('#mini'); if (!m) { m = document.createElement('div'); m.id = 'mini'; m.className = 'mini'; document.body.appendChild(m); }
  m.innerHTML = `<div class="grab"></div>` + opts.map(o => `<button class="opt press ${o === cur ? 'on' : ''}" data-v="${esc(o)}">${esc(o || t('none'))}</button>`).join('');
  m.querySelectorAll('.opt').forEach(b => b.onclick = () => { vib(); cb(b.dataset.v); closeMini(); });
  const bd = $('#mini-bd') || (() => { const d = document.createElement('div'); d.id = 'mini-bd'; d.className = 'backdrop'; d.style.zIndex = 22; d.onclick = closeMini; document.body.appendChild(d); return d; })();
  requestAnimationFrame(() => { m.classList.add('on'); bd.classList.add('on'); });
}
function closeMini() { const m = $('#mini'), bd = $('#mini-bd'); m && m.classList.remove('on'); bd && bd.classList.remove('on'); }

/* ---------------------------------------------------------------- demo data */
function DEMO() {
  if (DEMO.c) return DEMO.c;
  const d = n => { const x = new Date(); x.setDate(x.getDate() - n); return iso(x); };
  const ahead = (m, day) => { const x = new Date(); x.setMonth(x.getMonth() + m, day); return iso(x); };
  DEMO.c = {
    app: 'Vnancial', version: 'demo', period: t('months')[new Date().getMonth()] + ' ' + new Date().getFullYear(),
    kpi: { TotalBalance: 12450000, IncThis: 6500000, ExpThis: 3120000, InvValue: 4200000, DebtTotal: 300000, RecvTotal: 450000, SafeToday: 145000, TodayUnpl: 55000, BudgetLeft: 2380000, DaysLeft: 22 },
    accounts: [{ name: 'BCA', type: 'Bank', balance: 7820000 }, { name: 'BRI', type: 'Bank', balance: 2150000 }, { name: 'GOJEK', type: 'E-Wallet', balance: 410000 },
      { name: 'DANA', type: 'E-Wallet', balance: 380000 }, { name: 'ShopeePay', type: 'E-Wallet', balance: 140000 }, { name: 'OVO', type: 'E-Wallet', balance: 50000 }, { name: 'Cash', type: 'Cash', balance: 1500000 }],
    buckets: [{ name: 'Needs', share: .5, target: 3000000, actual: 1860000 }, { name: 'Family & Giving', share: .1, target: 600000, actual: 240000 }, { name: 'Wants', share: .2, target: 1200000, actual: 1090000 }, { name: 'Savings', share: .2, target: 1200000, actual: 600000 }],
    debts: [{ person: 'Raka', owesMe: 0, iOwe: 300000, status: 'Open' }, { person: 'Dinda', owesMe: 250000, iOwe: 0, status: 'Open' }, { person: 'Fajar', owesMe: 200000, iOwe: 0, status: 'Open' }, { person: 'Bima', owesMe: 0, iOwe: 0, status: 'Settled' }],
    holdings: [{ ticker: 'BBCA', name: 'Bank Central Asia', units: 200, value: 1960000, pl: 160000 }, { ticker: 'XAU', name: 'Emas', units: 1, value: 1540000, pl: 90000 }, { ticker: 'RDPU', name: 'Reksa dana pasar uang', units: 0, value: 700000, pl: 12000 }],
    expense: [{ Date: d(0), Description: 'Makan siang', Category: 'Food & Beverages', 'Sub-Category': 'Eat Out', Account: 'GOJEK', Amount: 35000, Ref: 'D1' },
      { Date: d(0), Description: 'Ojek ke kampus', Category: 'Transportation', 'Sub-Category': 'Public Transportation', Account: 'GOJEK', Amount: 20000, Ref: 'D2' },
      { Date: d(1), Description: 'Token listrik', Category: 'House & Utilities', 'Sub-Category': 'Electricity', Account: 'BRI', Amount: 150000, Ref: 'D3' },
      { Date: d(2), Description: 'Kopi', Category: 'Personal Lifestyle', 'Sub-Category': 'Coffee', Account: 'ShopeePay', Amount: 28000, Ref: 'D4' }],
    income: [{ Date: d(1), Description: 'Honor asisten', Category: 'Side Income', Account: 'BCA', Amount: 1500000, Ref: 'D5' }],
    transfer: [{ Date: d(1), From: 'BCA', To: 'DANA', Amount: 200000, Fee: 2500, Ref: 'D6' }],
    debt: [{ Date: d(2), Person: 'Raka', Action: 'Borrow', Amount: 300000, Account: 'Cash', Description: 'Pinjam ke Raka', Ref: 'D7' }],
    plans: [{ Date: ahead(0, 20), Item: 'Servis motor', Type: 'Expense', Category: 'Transportation', Amount: 450000, Repeat: 'Once', Status: 'Open' },
      { Date: ahead(1, 1), Item: 'Honor asisten', Type: 'Income', Category: 'Side Income', Amount: 1500000, Repeat: 'Monthly', Until: ahead(7, 30), Status: 'Open' },
      { Date: ahead(2, 15), Item: 'UKT semester genap', Type: 'Expense', Category: 'Education & Health', Amount: 5000000, Repeat: 'Once', Status: 'Open' },
      { Date: ahead(0, 5), Item: 'Kos', Type: 'Expense', Category: 'House & Utilities', Amount: 1200000, Repeat: 'Monthly', Status: 'Open' }],
    lists: { subcats: [{ cat: 'Food & Beverages', sub: 'Groceries' }, { cat: 'Food & Beverages', sub: 'Eat Out' }, { cat: 'Food & Beverages', sub: 'Snack & Drink' }, { cat: 'Food & Beverages', sub: 'Online Food Delivery' },
        { cat: 'Transportation', sub: 'Fuel' }, { cat: 'Transportation', sub: 'Public Transportation' }, { cat: 'Transportation', sub: 'Parking' },
        { cat: 'House & Utilities', sub: 'Electricity' }, { cat: 'House & Utilities', sub: 'Internet' }, { cat: 'Personal Lifestyle', sub: 'Coffee' }],
      expenseCats: [{ cat: 'Food & Beverages' }, { cat: 'House & Utilities' }, { cat: 'Transportation' }, { cat: 'Education & Health' }, { cat: 'Personal Lifestyle' }, { cat: 'Entertainment' }, { cat: 'Social & Travel' }, { cat: 'Financial' }, { cat: 'Others' }],
      incomeCats: [{ cat: 'Salary' }, { cat: 'Side Income' }, { cat: 'Passive Income' }, { cat: 'Gift' }, { cat: 'Others' }] }
  };
  return DEMO.c;
}


/* ---------------------------------------------------------------- splash */
const QUOTES = {
  id: ['Catat dulu, baru atur.', 'Uang kecil yang dicatat lebih berharga daripada uang besar yang terlupa.', 'Sisihkan untuk tabungan sebelum belanja, bukan sesudahnya.',
    'Anggaran bukan pembatas, melainkan peta.', 'Bedakan yang dibutuhkan dari yang diinginkan.', 'Dana darurat adalah ketenangan yang bisa ditabung.',
    'Utang terasa ringan saat diambil, berat saat dibayar.', 'Konsisten yang kecil mengalahkan semangat besar yang sesaat.',
    'Kekayaan dibangun dari kebiasaan, bukan keberuntungan.', 'Hemat hari ini, leluasa esok hari.'],
  en: ['Track it first, then manage it.', 'A small amount tracked beats a large amount forgotten.', 'Save before you spend, not after.',
    'A budget is not a cage, it is a map.', 'Know the difference between a need and a want.', 'An emergency fund is peace of mind you can save.',
    'Debt feels light when taken and heavy when repaid.', 'Small and steady beats big and brief.',
    'Wealth is built from habits, not luck.', 'Spend wisely today, live freely tomorrow.']
};
const splashT0 = performance.now();
(function () {
  const q = QUOTES[lang] || QUOTES.id; const el = document.getElementById('sp-quote');
  if (el) el.textContent = q[Math.floor(Math.random() * q.length)];
  requestAnimationFrame(() => requestAnimationFrame(() => { const f = document.getElementById('sp-fill'); if (f) f.style.width = '72%'; }));
})();
function hideSplash() {
  const sp = document.getElementById('splash'); if (!sp || sp.dataset.done) return; sp.dataset.done = 1;
  const wait = Math.max(0, 1400 - (performance.now() - splashT0));
  setTimeout(() => { const f = document.getElementById('sp-fill'); if (f) { f.style.transition = 'width .35s ease'; f.style.width = '100%'; }
    setTimeout(() => { sp.classList.add('out'); setTimeout(() => sp.remove(), 500); }, 380); }, wait);
}

/* ---------------------------------------------------------------- start */
document.documentElement.lang = lang;
history.pushState({}, '');
if (!cfg.url && !cfg.demo) stack = ['setup'];
render('fade');
if (cfg.url && !cfg.demo) { sync(false).finally(hideSplash); setTimeout(hideSplash, 4000); } else hideSplash();
window.addEventListener('online', () => { online = true; sync(false); });
window.addEventListener('offline', () => { online = false; rerender(); });
document.addEventListener('visibilitychange', () => { if (!document.hidden && cfg.url && !cfg.demo) sync(false); });
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
