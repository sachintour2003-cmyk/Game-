
/* =========================================================
AR CALCULATION TOOL — BEPL & BSCIPL
- Firebase Realtime DB with localStorage fallback
- Full Excel Import + Schedule Sheet Auto-Linking + Merge by Description
- Split Schedule (Multiple Payees)
- Lot Edit (Smart Update) + Split Payee Cheque Numbering
- Print / PDF Export
- Top-Up Lot: 1st cheque = ACCUMULATED DELTA (trigger date → next cheque date − 1 day)
- Lot Delete Option
- Auto Generate NEVER modifies existing lots
========================================================= */

const firebaseConfig = {
apiKey: "AIzaSyAQbnBB0FxycWtF6B3gKWfG4ZqijNZU9Vw",
authDomain: "ar-calculator-38781.firebaseapp.com",
databaseURL: "https://ar-calculator-38781-default-rtdb.firebaseio.com",
projectId: "ar-calculator-38781",
storageBucket: "ar-calculator-38781.firebasestorage.app",
messagingSenderId: "312708241196",
appId: "1:312708241196:web:b1d73cd31dd267c6627e89"
};

const LS_KEY = 'ar_calculation_tool_entries_v1';
let usingFirebase = false;
let database = null;

function initStorage() {
const isPlaceholder = !firebaseConfig.apiKey || firebaseConfig.apiKey === "YOUR_API_KEY";
if (!isPlaceholder) {
try { firebase.initializeApp(firebaseConfig); database = firebase.database(); usingFirebase = true; }
catch (err) { console.error('Firebase init failed:', err); usingFirebase = false; }
} else { usingFirebase = false; }
updateFirebaseStatusUI();
}

function updateFirebaseStatusUI() {
const el = document.getElementById('fbStatus');
const txt = document.getElementById('fbStatusText');
if (!el || !txt) return;
if (usingFirebase) { el.className = 'fb-status connected'; txt.textContent = 'Made by Abhishek "Dedicated to my Lovely Daughter Dhaanvi"'; }
else { el.className = 'fb-status local'; txt.textContent = 'Local Mode'; }
}

function lsGetAll() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch (e) { return {}; } }
function lsSetAll(obj) { try { localStorage.setItem(LS_KEY, JSON.stringify(obj)); } catch (e) { console.error(e); } }

let entries = [];
let editingId = null;
let editingFirebaseKey = null;
let paymentRowSeq = 0;
let formPayments = [];
let pendingDeleteId = null;
let idSeq = 1;
let isSyncing = false;
let selectedEntryIds = new Set();
let issuedSchedules = {};
let lotModalEntryId = null;
let chequeEditCtx = null;
let pendingImportFile = null;

let formPayees = [];
let payeeRowSeq = 0;

const fmt = n => '₹ ' + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const num = v => {
if (v === null || v === undefined || v === '') return 0;
if (typeof v === 'number') return isNaN(v) ? 0 : v;
const s = String(v).replace(/,/g, '').replace(/₹/g, '').replace(/\s/g, '').trim();
const n = parseFloat(s); return isNaN(n) ? 0 : n;
};
const round0 = v => Math.round(Number(v) || 0);
function escapeHtml(s) { return (s === undefined || s === null) ? '' : String(s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])); }
let toastTimer;
function showToast(msg, isError) {
const t = document.getElementById('toast');
document.getElementById('toastMsg').textContent = msg;
t.querySelector('i').className = isError ? 'fa-solid fa-triangle-exclamation' : 'fa-solid fa-circle-check';
t.querySelector('i').style.color = isError ? 'var(--red)' : 'var(--gold)';
t.classList.add('show');
clearTimeout(toastTimer);
toastTimer = setTimeout(() => t.classList.remove('show'), 5000);
}

function toggleMoreMenu(ev) {
if (ev) ev.stopPropagation();
document.getElementById('moreMenuList').classList.toggle('show');
}
function closeMoreMenu() { document.getElementById('moreMenuList').classList.remove('show'); }
document.addEventListener('click', (ev) => {
const menu = document.querySelector('.more-menu');
if (menu && !menu.contains(ev.target)) closeMoreMenu();
});

function switchTab(id) {
document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
document.getElementById(id).classList.add('active');
document.querySelector(`.tab[data-view="${id}"]`).classList.add('active');
if (id === 'tableView') renderEntriesTable();
if (id === 'dashView') renderDashboard();
}

function addPaymentRow(prefill) {
const row = prefill || { id: 'p' + (paymentRowSeq++), date: '', amount: '', mode: '', gstRate: 0 };
if (!prefill) formPayments.push(row);
renderPaymentRows();
}
function gstSplit(amount, gstRate) {
const amt = num(amount), rate = num(gstRate);
const gstAmt = rate > 0 ? (amt / (100 + rate)) * rate : 0;
return { gstAmt, netAmt: amt - gstAmt };
}
function renderPaymentRows() {
const tbody = document.getElementById('paymentRows');
tbody.innerHTML = '';
if (formPayments.length === 0) { tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--text-mute);font-size:12.5px;padding:14px;">No payments added yet</td></tr>`; return; }
formPayments.forEach(row => {
const { gstAmt, netAmt } = gstSplit(row.amount, row.gstRate);
const tr = document.createElement('tr');
tr.setAttribute('data-row-id', row.id);
tr.innerHTML = `
<td><input type="date" value="${row.date||''}" onchange="updatePayField('${row.id}','date',this.value)"></td>
<td class="amt-col"><input type="number" step="any" placeholder="0" value="${row.amount||''}" oninput="updatePayField('${row.id}','amount',this.value)"></td>
<td><select onchange="updatePayField('${row.id}','gstRate',this.value)" style="width:100%;border:1px solid transparent;background:transparent;padding:6px 4px;border-radius:6px;font-size:13px;">
<option value="0" ${num(row.gstRate)===0?'selected':''}>0%</option><option value="5" ${num(row.gstRate)===5?'selected':''}>5%</option><option value="12" ${num(row.gstRate)===12?'selected':''}>12%</option><option value="18" ${num(row.gstRate)===18?'selected':''}>18%</option></select></td>
<td class="amt-col gst-amt-cell" style="color:var(--text-mute);font-family:var(--font-num);font-size:12.5px;padding-left:10px;">${fmt(gstAmt)}</td>
<td class="amt-col net-amt-cell" style="font-family:var(--font-num);font-size:12.5px;padding-left:10px;font-weight:600;">${fmt(netAmt)}</td>
<td><input type="text" placeholder="Cash / Cheque / NEFT" value="${row.mode||''}" onchange="updatePayField('${row.id}','mode',this.value)"></td>
<td><button class="icon-btn" onclick="removePaymentRow('${row.id}')"><i class="fa-solid fa-trash"></i></button></td>`;
tbody.appendChild(tr);
});
recalcForm();
}
function updatePayField(id, field, value) {
const row = formPayments.find(r => r.id === id);
if (!row) return;
row[field] = value;
if (field === 'amount' || field === 'gstRate') {
const tr = document.querySelector(`tr[data-row-id="${id}"]`);
if (tr) { const { gstAmt, netAmt } = gstSplit(row.amount, row.gstRate); tr.querySelector('.gst-amt-cell').textContent = fmt(gstAmt); tr.querySelector('.net-amt-cell').textContent = fmt(netAmt); }
}
recalcForm();
}
function removePaymentRow(id) { formPayments = formPayments.filter(r => r.id !== id); renderPaymentRows(); }

function normalizePayees(payees) {
if (!Array.isArray(payees)) return [];
return payees.filter(p => p && (p.name || p.sharePct)).map(p => ({
name: String(p.name || '').trim(),
sharePct: num(p.sharePct),
favouring: String(p.name || '').trim(),
pan: String(p.pan || '').trim(),
controlNo: String(p.controlNo || '').trim()
}));
}
function computePayeeSplits(entry, c) {
const payees = normalizePayees(entry && entry.payees);
if (payees.length === 0) {
return [{
name: entry && entry.client ? entry.client : 'Payee',
sharePct: 100,
favouring: (entry && entry.favouring) || '',
pan: (entry && entry.pan) || '',
controlNo: (entry && entry.controlNo) || '',
gross: c.arMonthly,
tds: c.tds,
net: c.net
}];
}
return payees.map(p => {
const gross = round0(c.arMonthly * (p.sharePct / 100));
const tds = round0(gross * 0.10);
const net = round0(gross - tds);
return { ...p, gross, tds, net };
});
}

function computeEntryValues(e) {
const gross = num(e.area) * num(e.rate);
const totalDisc = num(e.lps) + num(e.cpDisc) + num(e.iphone) + num(e.tpd) + num(e.spot) + num(e.otherDisc);
const manualDue = (e.manualDue !== undefined && e.manualDue !== null && e.manualDue > 0) ? e.manualDue : null;
const due = manualDue !== null ? manualDue : gross * (num(e.duePct) / 100);
const grossReceived = (e.payments || []).reduce((s, p) => s + num(p.amount), 0);
const gstTotal = (e.payments || []).reduce((s, p) => { const rate = num(p.gstRate); return s + (rate > 0 ? (num(p.amount) / (100 + rate)) * rate : 0); }, 0);
const received = grossReceived - gstTotal;
const balance = due - received;
let eligible;
let importedChequeCount = 0;
if (e.isImported) {
eligible = due > 0;
const lotData = issuedSchedules[e.id];
if (lotData && lotData.lots) { lotData.lots.forEach(lot => { importedChequeCount += (lot.cheques || []).length; }); }
if (importedChequeCount === 0) importedChequeCount = num(e.arPeriod) || 1;
} else { eligible = due > 0 && (received + totalDisc) >= due; }
const basePeriod = num(e.arPeriod) || 12;
const manualArMonthly = (e.manualArMonthly !== undefined && e.manualArMonthly !== null && e.manualArMonthly > 0) ? e.manualArMonthly : null;
const autoArMonthly = round0((received + totalDisc) * 0.01);
const arMonthly = manualArMonthly !== null ? round0(manualArMonthly) : autoArMonthly;
const tds = round0(arMonthly * 0.10);
const net = round0(arMonthly - tds);
const effectivePeriod = e.isImported ? importedChequeCount : basePeriod;
const totalPayable = round0(net * effectivePeriod);
let startDate = new Date();
const validDates = (e.payments || []).map(p => p.date).filter(d => d).map(d => new Date(d)).filter(d => !isNaN(d));
if (validDates.length) { const latest = new Date(Math.max(...validDates.map(d => d.getTime()))); latest.setDate(latest.getDate() + 1); startDate = latest; }
const payees = normalizePayees(e.payees);
const payeeSplits = computePayeeSplits(e, { arMonthly, tds, net });
return { gross, totalDisc, due, received, grossReceived, gstTotal, balance, eligible, period: effectivePeriod, basePeriod, arMonthly, tds, net, totalPayable, startDate, manualDue, importedChequeCount, isImported: !!e.isImported, manualArMonthly, autoArMonthly, hasManualArMonthly: manualArMonthly !== null, payees, payeeSplits, hasPayeeSplit: payees.length > 0 };
}

/* ============ AR Increase Detection ============ */
function detectArIncrease(entry) {
if (!entry || entry.scheduledArMonthly === undefined || entry.scheduledArMonthly === null) return null;
const c = computeEntryValues(entry);
const baseline = round0(entry.scheduledArMonthly);
const delta = round0(c.arMonthly - baseline);
if (delta <= 0) return null;
return { baseline, current: c.arMonthly, delta };
}

function recalcForm() {
const area = num(document.getElementById('f_area').value);
const rate = num(document.getElementById('f_rate').value);
const gross = area * rate;
document.getElementById('f_gross').value = fmt(gross);
const lps = num(document.getElementById('f_lps').value), cpDisc = num(document.getElementById('f_cpDisc').value), iphone = num(document.getElementById('f_iphone').value), tpd = num(document.getElementById('f_tpd').value), spot = num(document.getElementById('f_spot').value), otherDisc = num(document.getElementById('f_otherDisc').value);
const totalDisc = lps + cpDisc + iphone + tpd + spot + otherDisc;
document.getElementById('f_totalDisc').value = fmt(totalDisc);
let duePct = num(document.getElementById('f_duePct').value);
if (duePct > 100) { duePct = 100; document.getElementById('f_duePct').value = 100; }
if (duePct < 0) { duePct = 0; document.getElementById('f_duePct').value = 0; }
const manualDueVal = document.getElementById('f_manualDue').value;
const manualDue = manualDueVal !== '' ? num(manualDueVal) : null;
const autoDue = gross * (duePct / 100);
const due = (manualDue !== null && manualDue > 0) ? manualDue : autoDue;
document.getElementById('f_due').value = fmt(due);
const grossReceived = formPayments.reduce((s, p) => s + num(p.amount), 0);
const gstTotal = formPayments.reduce((s, p) => s + gstSplit(p.amount, p.gstRate).gstAmt, 0);
const received = grossReceived - gstTotal;
document.getElementById('paymentTotalCell').textContent = fmt(grossReceived);
document.getElementById('paymentGstTotalCell').textContent = 'GST ' + fmt(gstTotal);
document.getElementById('paymentNetTotalCell').textContent = 'Net ' + fmt(received);
const balance = due - received;
const editingEntry = editingId ? entries.find(e => e.id === editingId) : null;
const isImported = editingEntry ? editingEntry.isImported : false;
let eligible;
if (isImported) eligible = due > 0; else eligible = due > 0 && (received + totalDisc) >= due;
const manualArMonthlyVal = document.getElementById('s_arMonthly').value;
const manualArMonthly = manualArMonthlyVal !== '' ? num(manualArMonthlyVal) : null;
const autoArMonthly = round0((received + totalDisc) * 0.01);
const arMonthly = (manualArMonthly !== null && manualArMonthly > 0) ? round0(manualArMonthly) : autoArMonthly;
const tds = round0(arMonthly * 0.10);
const net = round0(arMonthly - tds);
document.getElementById('s_due').textContent = fmt(due);
document.getElementById('s_received').textContent = fmt(received);
document.getElementById('s_balance').textContent = fmt(balance);
const elCell = document.getElementById('s_eligible_cell');
const elVal = document.getElementById('s_eligible');
if (eligible) { elVal.textContent = isImported ? 'Eligible ✓ (Imported)' : 'Eligible ✓'; elCell.className = 'summary-cell eligible'; }
else { elVal.textContent = 'Not Eligible'; elCell.className = 'summary-cell not-eligible'; }
document.getElementById('s_arMonthly').value = arMonthly;
document.getElementById('s_tds').value = fmt(tds);
document.getElementById('s_net').value = fmt(net);
const arMonthlyBadge = document.getElementById('arMonthlyBadge');
if (manualArMonthly !== null && manualArMonthly > 0) arMonthlyBadge.style.display = 'inline-flex';
else arMonthlyBadge.style.display = 'none';
const manualField = document.getElementById('f_manualDue');
if (manualDue !== null && manualDue > 0) manualField.closest('.field').classList.add('manual-due');
else manualField.closest('.field').classList.remove('manual-due');
if (editingId !== null) {
const existing = entries.find(e => e.id === editingId);
if (existing) { const c = computeEntryValues(existing); showChequeSection(existing, c); }
} else { document.getElementById('chequeSection').style.display = 'none'; document.getElementById('autoGenBox').style.display = 'none'; }
}

function readForm() {
const manualDueVal = document.getElementById('f_manualDue').value;
const arMonthlyVal = document.getElementById('s_arMonthly').value;
const editingEntry = editingId ? entries.find(x => x.id === editingId) : null;
return {
client: document.getElementById('f_client').value.trim(),
unitNo: document.getElementById('f_unitNo').value.trim(),
description: document.getElementById('f_unitNo').value.trim(),
cp: document.getElementById('f_cp').value.trim(),
rm: document.getElementById('f_rm').value.trim(),
status: document.getElementById('f_status').value,
floor: document.getElementById('f_floor').value.trim(),
area: num(document.getElementById('f_area').value),
rate: num(document.getElementById('f_rate').value),
goldCp: document.getElementById('f_goldCp').value.trim(),
goldClient: document.getElementById('f_goldClient').value.trim(),
lps: num(document.getElementById('f_lps').value),
cpDisc: num(document.getElementById('f_cpDisc').value),
iphone: num(document.getElementById('f_iphone').value),
tpd: num(document.getElementById('f_tpd').value),
spot: num(document.getElementById('f_spot').value),
otherDiscType: document.getElementById('f_otherDiscType').value.trim(),
otherDisc: num(document.getElementById('f_otherDisc').value),
discountDate: document.getElementById('f_discountDate').value,
duePct: num(document.getElementById('f_duePct').value) || 100,
manualDue: manualDueVal !== '' ? num(manualDueVal) : null,
manualArMonthly: (arMonthlyVal !== '' && num(arMonthlyVal) > 0) ? num(arMonthlyVal) : null,
payments: formPayments.map(p => ({ date: p.date, amount: num(p.amount), mode: p.mode, gstRate: num(p.gstRate) })),
bank: document.getElementById('f_bank').value.trim(),
ref: document.getElementById('f_ref').value.trim(),
pan: document.getElementById('f_pan').value.trim(),
favouring: document.getElementById('f_favouring').value.trim(),
chequeNo: document.getElementById('f_chequeNo').value.trim(),
controlNo: document.getElementById('f_controlNo').value.trim(),
company: document.getElementById('f_company').value,
arPeriod: num(document.getElementById('f_arPeriod').value) || 12,
sourceSheet: editingEntry ? (editingEntry.sourceSheet || '') : '',
isImported: editingEntry ? (editingEntry.isImported || false) : false,
payees: (editingEntry && Array.isArray(editingEntry.payees)) ? editingEntry.payees : []
};
}

function fillForm(e) {
document.getElementById('f_client').value = e.client || '';
document.getElementById('f_unitNo').value = e.description || e.unitNo || '';
document.getElementById('f_cp').value = e.cp || '';
document.getElementById('f_rm').value = e.rm || '';
document.getElementById('f_status').value = e.status || 'AR';
document.getElementById('f_floor').value = e.floor || '';
document.getElementById('f_area').value = e.area || '';
document.getElementById('f_rate').value = e.rate || '';
document.getElementById('f_goldCp').value = e.goldCp || '';
document.getElementById('f_goldClient').value = e.goldClient || '';
document.getElementById('f_lps').value = e.lps || '';
document.getElementById('f_cpDisc').value = e.cpDisc || '';
document.getElementById('f_iphone').value = e.iphone || '';
document.getElementById('f_tpd').value = e.tpd || '';
document.getElementById('f_spot').value = e.spot || '';
document.getElementById('f_otherDiscType').value = e.otherDiscType || '';
document.getElementById('f_otherDisc').value = e.otherDisc || '';
document.getElementById('f_discountDate').value = e.discountDate || '';
document.getElementById('f_duePct').value = e.duePct || 100;
document.getElementById('f_manualDue').value = (e.manualDue !== undefined && e.manualDue !== null && e.manualDue > 0) ? e.manualDue : '';
document.getElementById('f_bank').value = e.bank || '';
document.getElementById('f_ref').value = e.ref || '';
document.getElementById('f_pan').value = e.pan || '';
document.getElementById('f_favouring').value = e.favouring || '';
document.getElementById('f_chequeNo').value = e.chequeNo || '';
document.getElementById('f_controlNo').value = e.controlNo || '';
document.getElementById('f_company').value = e.company || 'BEPL';
document.getElementById('f_arPeriod').value = e.arPeriod || 12;
formPayments = (e.payments || []).map(p => ({ id: 'p' + (paymentRowSeq++), date: p.date || '', amount: p.amount || 0, mode: p.mode || '', gstRate: num(p.gstRate) || 0 }));
renderPaymentRows();
if (e.manualArMonthly && e.manualArMonthly > 0) document.getElementById('s_arMonthly').value = e.manualArMonthly;
else document.getElementById('s_arMonthly').value = '';
if (e.id) { const c = computeEntryValues(e); showChequeSection(e, c); }
}

function clearForm() {
['f_client','f_unitNo','f_cp','f_rm','f_floor','f_area','f_rate','f_goldCp','f_goldClient','f_lps','f_cpDisc','f_iphone','f_tpd','f_spot','f_otherDiscType','f_otherDisc','f_discountDate','f_manualDue','f_bank','f_ref','f_pan','f_favouring','f_chequeNo','f_controlNo'].forEach(id => { document.getElementById(id).value = ''; });
document.getElementById('f_status').value = 'AR';
document.getElementById('f_company').value = 'BEPL';
document.getElementById('f_duePct').value = 100;
document.getElementById('f_arPeriod').value = 12;
document.getElementById('s_arMonthly').value = '';
formPayments = [];
renderPaymentRows();
document.getElementById('chequeSection').style.display = 'none';
document.getElementById('autoGenBox').style.display = 'none';
document.getElementById('f_manualDue').closest('.field').classList.remove('manual-due');
}

function cancelEdit() {
editingId = null; editingFirebaseKey = null;
document.getElementById('formTitle').textContent = 'New Entry';
document.getElementById('editingBadge').style.display = 'none';
document.getElementById('duplicateBtn').style.display = 'none';
clearForm();
}

function checkUnitNoDuplicate(unitNo, excludingId) {
if (!unitNo) return null;
return entries.find(e => e.id !== excludingId && (e.unitNo || '').trim().toLowerCase() === unitNo.trim().toLowerCase()) || null;
}
function promptForDuplicateReason(unitNo, existingClient) {
return prompt(`Unit No. '${unitNo}' already exists for client '${existingClient}'.\nEnter a reason for this duplicate entry, or Cancel:`);
}

function showChequeSection(entry, comp) {
const sec = document.getElementById('chequeSection');
const agBox = document.getElementById('autoGenBox');
if (!entry || !comp) { sec.style.display = 'none'; if (agBox) agBox.style.display = 'none'; return; }
const c = comp || computeEntryValues(entry);
if ((c.eligible || entry.isImported) && entry.id) { sec.style.display = 'block'; if (agBox) agBox.style.display = 'flex'; renderChequeSchedule(entry); }
else { sec.style.display = 'none'; if (agBox) agBox.style.display = 'none'; }
}

function getLotData(entryId) { if (!issuedSchedules[entryId]) issuedSchedules[entryId] = { lots: [] }; return issuedSchedules[entryId]; }
function getTotalCheques(entryId) { const data = getLotData(entryId); let total = 0; data.lots.forEach(lot => { total += (lot.cheques || []).length; }); return total; }

function renderChequeSchedule(entry) {
const container = document.getElementById('chequeScheduleContainer');
const data = getLotData(entry.id);
const lots = data.lots || [];
const lotBadge = document.getElementById('lotCountBadge');
const totalCheques = getTotalCheques(entry.id);
const c = computeEntryValues(entry);
lotBadge.textContent = lots.length + ' lot' + (lots.length !== 1 ? 's' : '') + ' · ' + totalCheques + ' cheque' + (totalCheques !== 1 ? 's' : '') + (c.hasPayeeSplit ? ' · SPLIT' : '');
if (lots.length === 0) { container.innerHTML = `<div class="empty-state" style="padding:20px 10px;"><i class="fa-regular fa-file-lines"></i><h4 style="font-size:13px;">No cheques issued yet</h4><p style="font-size:12px;">Click "Auto Generate" to build schedule instantly, or "New / Update Lot" for manual. Use "Split Schedule" to divide between multiple payees.</p></div>`; return; }

const inc = detectArIncrease(entry);
let incBannerHtml = '';
if (inc) {
incBannerHtml = `<div class="ar-increase-banner"><div class="ai-info"><div class="ai-icon"><i class="fa-solid fa-arrow-trend-up"></i></div><div class="ai-text"><h5>AR Monthly Increased</h5><p>New payment(s) pushed AR Monthly from ${fmt(inc.baseline)} to ${fmt(inc.current)} (+${fmt(inc.delta)}). Ek <strong>naya Top-Up Lot</strong> banega jisme sirf increase amount hoga — 1st cheque <strong>accumulated delta</strong> (trigger date se next cheque date tak), baaki cheques full delta. <strong>Purana lot bilkul safe rahega.</strong></p></div></div><button class="btn btn-gold btn-sm" onclick="openTopUpLotModal(${entry.id})"><i class="fa-solid fa-bolt"></i> Create Top-Up Lot</button></div>`;
}

let html = '';
lots.forEach((lot, lotIdx) => {
const cheques = lot.cheques || [];
if (cheques.length === 0) return;
const controlNo = lot.controlNo || '';
const issuedDate = lot.issuedDate ? fmtDateSlash(new Date(lot.issuedDate)) : '—';
const payeeTag = lot.payeeName ? `<span class="lot-payee"><i class="fa-solid fa-user"></i> ${escapeHtml(lot.payeeName)}${lot.payeeSharePct ? ' · ' + lot.payeeSharePct + '%' : ''}</span>` : '';
const topUpTag = lot.isTopUp ? `<span class="badge-topup"><i class="fa-solid fa-arrow-trend-up"></i> TOP-UP</span>` : '';
let tableRows = cheques.map((r, chIdx) => {
const narration = `CH NO. ${r.chequeNo||'---'} DTD ${fmtDateSlash(r.chequeDate)} ISSUED TOWARDS AR FOR THE PERIOD FROM ${fmtDateSlash(r.from)} TO ${fmtDateSlash(r.to)} VIDE CONTROL NO ${controlNo}`;
const rowClass = r.isTopUp ? ' style="background:var(--amber-bg);"' : '';
const grossMark = r.isTopUp ? ` <i class="fa-solid fa-arrow-trend-up" style="color:var(--amber);font-size:10px;" title="Includes AR increase"></i>` : '';
return `<tr${rowClass}><td>${r.sno}</td><td><div style="display:flex;align-items:center;gap:3px;justify-content:center;"><input class="cheque-input" type="text" value="${escapeHtml(r.chequeNo||'')}" onchange="updateChequeNo(${entry.id},${lotIdx},${chIdx},this.value)" placeholder="—" style="text-align:center;font-family:var(--font-num);"><button class="edit-chq-btn" onclick="openChequeEditModal(${entry.id},${lotIdx},${chIdx})"><i class="fa-solid fa-pen"></i></button></div></td><td>${fmtDateSlash(r.chequeDate)}</td><td>${escapeHtml(r.favouring||entry.favouring||'—')}</td><td class="num">${fmt(r.gross)}${grossMark}</td><td class="num">${fmt(r.tds)}</td><td class="num">${fmt(r.net)}</td><td>${fmtDateSlash(r.from)}</td><td>${fmtDateSlash(r.to)}</td><td>${escapeHtml(controlNo)}</td><td class="narration-cell"><span style="font-size:10px;">${escapeHtml(narration)}</span><button class="copy-narr-btn" onclick="copySingleNarration(${entry.id},${lotIdx},${chIdx})">Copy</button></td></tr>`;
}).join('');
let sumGross = 0, sumTds = 0, sumNet = 0;
cheques.forEach(r => { sumGross += num(r.gross); sumTds += num(r.tds); sumNet += num(r.net); });
html += `<div class="lot-container"><div class="lot-header"><div class="lot-title"><span class="lot-num">Lot ${lot.lotNumber}</span>${topUpTag}${payeeTag}<span style="font-weight:400;font-size:12px;opacity:0.8;">${cheques.length} cheque${cheques.length>1?'s':''}</span><span class="lot-issue-date"><i class="fa-regular fa-calendar"></i> ${issuedDate}</span><button class="edit-lot-btn" onclick="openLotModalForLot(${entry.id},${lotIdx})"><i class="fa-solid fa-pen"></i> Edit Lot</button><button class="edit-lot-btn" style="background:var(--red);color:#fff;" onclick="deleteLot(${entry.id},${lotIdx})"><i class="fa-solid fa-trash"></i> Delete</button></div><div class="lot-control"><label>Control No.</label><input type="text" id="lotControl_${entry.id}_${lotIdx}" value="${escapeHtml(controlNo)}" placeholder="Enter Control No." style="width:140px;"><button class="update-lot-btn" onclick="updateLotControlNo(${entry.id},${lotIdx})"><i class="fa-solid fa-check"></i> Update</button></div></div><div class="table-scroll"><table class="cheque-issue-table"><thead><tr><th>S.No</th><th>Cheque No.</th><th>Cheque Date</th><th>Favouring</th><th>Gross</th><th>(-) TDS @10%</th><th>Net Amt</th><th>Period From</th><th>Period To</th><th>Control No.</th><th style="min-width:180px;">Narration</th></tr></thead><tbody>${tableRows}</tbody><tfoot><tr class="lot-total-row"><td colspan="4" style="text-align:right;font-weight:700;">Lot Total</td><td class="num">${fmt(sumGross)}</td><td class="num">${fmt(sumTds)}</td><td class="num">${fmt(sumNet)}</td><td colspan="4"></td></tr></tfoot></table></div></div>`;
});
let splitSummaryHtml = '';
if (c.hasPayeeSplit && c.payeeSplits.length > 0) {
const rows = c.payeeSplits.map(p => `<tr><td>${escapeHtml(p.name)}</td><td class="num">${p.sharePct}%</td><td class="num">${fmt(p.gross)}</td><td class="num">${fmt(p.tds)}</td><td class="num">${fmt(p.net)}</td><td>${escapeHtml(p.favouring||'—')}</td><td>${escapeHtml(p.controlNo||'—')}</td></tr>`).join('');
splitSummaryHtml = `<div class="lot-container"><div class="lot-header"><div class="lot-title"><span class="lot-num" style="background:var(--purple);color:#fff;"><i class="fa-solid fa-people-split"></i> PAYEE SPLIT</span><span style="font-weight:400;font-size:12px;opacity:0.85;">Monthly AR divided between ${c.payeeSplits.length} payees</span></div></div><div class="table-scroll"><table class="cheque-issue-table"><thead><tr><th>Payee</th><th>Share %</th><th>Monthly Gross</th><th>TDS @10%</th><th>Monthly Net</th><th>Favouring</th><th>Control No.</th></tr></thead><tbody>${rows}</tbody></table></div></div>`;
}
container.innerHTML = incBannerHtml + splitSummaryHtml + html + `<div class="cheque-actions"><button class="btn btn-outline btn-sm" onclick="printCurrentEntry()"><i class="fa-solid fa-print"></i> Print / PDF</button><button class="btn btn-outline btn-sm" onclick="copyAllNarrationsFromForm(${entry.id})"><i class="fa-solid fa-copy"></i> Copy All Narrations</button><button class="btn btn-outline btn-sm" onclick="exportChequeSchedule()"><i class="fa-solid fa-file-excel"></i> Export Schedule</button></div>`;
}

/* ============ Auto Generate — SAFE MODE: never modifies existing lots ============ */
function autoGenerateSchedule() {
if (editingId === null) { showToast('Please save the entry first, then open it in edit mode', true); return; }
const e = entries.find(x => x.id === editingId);
if (!e) { showToast('Entry not found', true); return; }
const c = computeEntryValues(e);
if (c.due <= 0) { showToast('Due amount is 0 — cannot generate AR schedule', true); return; }
const data = getLotData(editingId);
if (data.lots && data.lots.length > 0) {
showToast('Schedule already exists. Auto Generate never modifies existing lots — use "New / Update Lot" to add or edit.', true);
return;
}
let chequeCount = num(e.arPeriod) || 12;
if (chequeCount < 1) chequeCount = 1;
if (chequeCount > 120) { if (!confirm(`${chequeCount} cheques will be generated. That's a lot. Continue?`)) return; }

const payeeSplits = c.payeeSplits;
const isSplit = c.hasPayeeSplit && payeeSplits.length > 0;
const lots = [];
let runningSno = 1;

let runningChequeNo = e.chequeNo || '';

payeeSplits.forEach((payee, payeeIdx) => {
let monthFrom = new Date(c.startDate);
const controlNo = isSplit ? (payee.controlNo || '') : (e.controlNo || '');
const startChequeNo = runningChequeNo || '';

const newCheques = [];
for (let i = 0; i < chequeCount; i++) {
                const monthTo = addDays(addMonthsAnchored(c.startDate, i + 1), -1);
const chequeDate = addDays(monthTo, 1);
newCheques.push({
sno: runningSno++,
chequeNo: startChequeNo ? incrementCode(startChequeNo, i) : '',
chequeDate: new Date(chequeDate),
from: new Date(monthFrom),
to: new Date(monthTo),
gross: payee.gross, tds: payee.tds, net: payee.net,
favouring: payee.favouring || e.favouring || '',
payeeName: payee.name,
payeeSharePct: payee.sharePct
});
monthFrom = addDays(monthTo, 1);
}

if (startChequeNo && newCheques.length > 0) {
runningChequeNo = incrementCode(startChequeNo, newCheques.length);
}

lots.push({
lotNumber: payeeIdx + 1,
issuedDate: new Date(),
controlNo: controlNo,
payeeName: payee.name,
payeeSharePct: payee.sharePct,
payeeFavouring: payee.favouring || e.favouring || '',
payeePan: payee.pan || e.pan || '',
cheques: newCheques
});
});

issuedSchedules[editingId] = { lots: lots };
const serial = serializeLots(editingId);
const allCheques = getAllChequesFlat(editingId);
const lastCheque = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;
const updated = { ...e, issuedCheques: serial, chequeNo: lastCheque && lastCheque.chequeNo ? lastCheque.chequeNo : e.chequeNo, arPeriod: chequeCount, scheduledArMonthly: c.arMonthly };
updateEntryInFirebase(updated).then(() => {
fillForm(updated);
showToast(`✓ Auto-generated ${chequeCount} cheque${chequeCount>1?'s':''}${isSplit ? ' × ' + payeeSplits.length + ' payees' : ''}`, false);
}).catch(() => {});
}

/* ============ Bulk Auto Generate — SKIPS entries with existing lots ============ */
function bulkAutoGenerateSchedules() {
const ids = Array.from(selectedEntryIds);
if (ids.length === 0) { showToast('No entries selected', true); return; }
const allMapped = ids.map(id => entries.find(e => e.id === id)).filter(e => e && computeEntryValues(e).due > 0);
const skippedExisting = allMapped.filter(e => issuedSchedules[e.id] && issuedSchedules[e.id].lots && issuedSchedules[e.id].lots.length > 0);
const targets = allMapped.filter(e => !(issuedSchedules[e.id] && issuedSchedules[e.id].lots && issuedSchedules[e.id].lots.length > 0));
if (targets.length === 0) {
showToast(skippedExisting.length > 0
? `${skippedExisting.length} selected entr${skippedExisting.length===1?'y':'ies'} already have schedules — never modified by Auto Generate.`
: 'No selected entries have a due amount > 0', true);
return;
}
let msg = `Auto-generate AR schedules for ${targets.length} entr${targets.length===1?'y':'ies'}?`;
if (skippedExisting.length > 0) msg += `\n\nℹ ${skippedExisting.length} entr${skippedExisting.length===1?'y':'ies'} already have schedules and will be SKIPPED (existing lots are never modified).`;
if (!confirm(msg)) return;
let done = 0, failed = 0;
showToast(`Generating ${targets.length} schedules...`, false);
const promises = targets.map(e => {
try {
const c = computeEntryValues(e);
let chequeCount = num(e.arPeriod) || 12;
if (chequeCount < 1) chequeCount = 1;
if (chequeCount > 120) chequeCount = 120;
const payeeSplits = c.payeeSplits;
const isSplit = c.hasPayeeSplit && payeeSplits.length > 0;
const lots = [];
let runningSno = 1;

let runningChequeNo = e.chequeNo || '';

payeeSplits.forEach((payee, payeeIdx) => {
let monthFrom = new Date(c.startDate);
const controlNo = isSplit ? (payee.controlNo || '') : (e.controlNo || '');
const startChequeNo = runningChequeNo || '';

const newCheques = [];
for (let i = 0; i < chequeCount; i++) {
                const monthTo = addDays(addMonthsAnchored(c.startDate, i + 1), -1);
const chequeDate = addDays(monthTo, 1);
newCheques.push({
sno: runningSno++, chequeNo: startChequeNo ? incrementCode(startChequeNo, i) : '',
chequeDate: new Date(chequeDate),
from: new Date(monthFrom), to: new Date(monthTo),
gross: payee.gross, tds: payee.tds, net: payee.net,
favouring: payee.favouring || e.favouring || '',
payeeName: payee.name, payeeSharePct: payee.sharePct
});
monthFrom = addDays(monthTo, 1);
}

if (startChequeNo && newCheques.length > 0) {
runningChequeNo = incrementCode(startChequeNo, newCheques.length);
}

lots.push({
lotNumber: payeeIdx + 1, issuedDate: new Date(),
controlNo: controlNo,
payeeName: payee.name, payeeSharePct: payee.sharePct,
payeeFavouring: payee.favouring || e.favouring || '',
payeePan: payee.pan || e.pan || '',
cheques: newCheques
});
});

issuedSchedules[e.id] = { lots: lots };
const serial = serializeLots(e.id);
const lastChq = getAllChequesFlat(e.id).slice(-1)[0];
const updated = { ...e, issuedCheques: serial, chequeNo: lastChq && lastChq.chequeNo ? lastChq.chequeNo : e.chequeNo, arPeriod: chequeCount, scheduledArMonthly: c.arMonthly };
return updateEntryInFirebase(updated).then(() => { done++; }).catch(() => { failed++; });
} catch (err) { console.error(err); failed++; return Promise.resolve(); }
});
Promise.all(promises).then(() => {
const skippedMsg = skippedExisting.length > 0 ? ` (${skippedExisting.length} skipped — had existing lots)` : '';
showToast(`✓ ${done} schedules generated${skippedMsg}${failed>0?`, ${failed} failed`:''}`, failed>0);
selectedEntryIds.clear(); renderEntriesTable(); updateBulkDeleteUI();
});
}

/* ============ Top-Up Context — find period containing trigger date ============ */
function findTriggerContext(entryId, triggerDate) {
const data = getLotData(entryId);
const lots = (data.lots || []).filter(l => (l.cheques || []).length > 0);
if (lots.length === 0) return null;
const all = [];
lots.forEach(lot => (lot.cheques || []).forEach(ch => all.push(ch)));
all.sort((a, b) => new Date(a.from) - new Date(b.from));
const td = new Date(triggerDate);
let affectedIdx = -1;
for (let i = 0; i < all.length; i++) {
const from = new Date(all[i].from), to = new Date(all[i].to);
if (td >= from && td <= to) { affectedIdx = i; break; }
}
if (affectedIdx === -1) {
if (td < new Date(all[0].from)) affectedIdx = 0;
else return null;
}
return {
all,
affectedIdx,
affectedCheque: all[affectedIdx],
remainingCount: all.length - affectedIdx
};
}

/* ============ Last Existing Cheque (by period end) ============ */
function getLastExistingCheque(entryId) {
const all = getAllChequesFlat(entryId);
if (all.length === 0) return null;
all.sort((a, b) => new Date(b.to) - new Date(a.to));
return all[0];
}

/* ============ Months in Period (fractional) ============ */
function monthsInPeriod(from, to) {
let count = 0;
let cursor = new Date(from);
const end = new Date(to);
let safety = 0;
while (cursor <= end && safety++ < 500) {
const monthEnd = addDays(addMonths(cursor, 1), -1);
if (monthEnd <= end) {
count++;
cursor = addDays(monthEnd, 1);
} else {
const totalDays = dayCount(cursor, monthEnd) + 1;
const actualDays = dayCount(cursor, end) + 1;
count += actualDays / totalDays;
break;
}
}
return count;
}

/* ============ Amount Increase Mode toggle ============ */
function onLotIncModeChange() {
const checked = document.getElementById('lm_incMode').checked;
const details = document.getElementById('lm_incDetails');
details.style.display = checked ? 'block' : 'none';

const countInput = document.getElementById('lm_count');
if (!checked) { 
countInput.removeAttribute('readonly');
countInput.style.background = '';
updateLotPreview(); 
return; 
}

const e = entries.find(x => x.id === lotModalEntryId);
if (!e) return;
const inc = detectArIncrease(e);
if (!inc) {
showToast('No AR increase detected', true);
document.getElementById('lm_incMode').checked = false;
details.style.display = 'none';
countInput.removeAttribute('readonly');
countInput.style.background = '';
return;
}

document.getElementById('lm_incBaseline').textContent = fmt(inc.baseline);
document.getElementById('lm_incCurrent').textContent = fmt(inc.current);
document.getElementById('lm_incDelta').textContent = '+' + fmt(inc.delta);

const paymentDates = (e.payments || []).map(p => p.date).filter(d => d).map(d => new Date(d)).filter(d => !isNaN(d));
let triggerDate = new Date();
if (paymentDates.length > 0) triggerDate = new Date(Math.max(...paymentDates.map(d => d.getTime())));
document.getElementById('lm_incTriggerDate').value = toDateInputValue(triggerDate);

countInput.setAttribute('readonly', 'true');
countInput.style.background = 'var(--paper)';

updateLotPreview();
}

function persistLotsForEntry(entryId, onDone) {
const entry = entries.find(e => e.id === entryId);
if (!entry) return;
const serial = serializeLots(entryId);
const allCheques = getAllChequesFlat(entryId);
const lastCheque = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;
const updated = { ...entry, issuedCheques: serial, chequeNo: lastCheque && lastCheque.chequeNo ? lastCheque.chequeNo : entry.chequeNo };
updateEntryInFirebase(updated).then(() => { if (editingId === entryId) document.getElementById('f_chequeNo').value = updated.chequeNo || ''; renderChequeSchedule(entry); if (typeof onDone === 'function') onDone(); }).catch(() => {});
}

function toDateInputValue(d) { const dt = (d instanceof Date) ? d : new Date(d); if (isNaN(dt)) return ''; return dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + String(dt.getDate()).padStart(2, '0'); }
function openChequeEditModal(entryId, lotIdx, chIdx) {
const data = getLotData(entryId);
const lot = data.lots[lotIdx];
if (!lot || !lot.cheques[chIdx]) { showToast('Cheque not found', true); return; }
const ch = lot.cheques[chIdx];
chequeEditCtx = { entryId, lotIdx, chIdx };
document.getElementById('ce_chequeNo').value = ch.chequeNo || '';
document.getElementById('ce_chequeDate').value = toDateInputValue(ch.chequeDate);
const entry = entries.find(e => e.id === entryId);
document.getElementById('ce_info').textContent = `Lot ${lot.lotNumber} · S.No ${ch.sno} · ${lot.payeeName || (entry ? entry.client : '')} · Net ₹ ${round0(ch.net)}`;
document.getElementById('chequeEditModal').classList.add('show');
setTimeout(() => document.getElementById('ce_chequeNo').focus(), 60);
}
function closeChequeEditModal() { document.getElementById('chequeEditModal').classList.remove('show'); chequeEditCtx = null; }

function saveChequeEdit() {
if (!chequeEditCtx) return;
const { entryId, lotIdx, chIdx } = chequeEditCtx;
const data = getLotData(entryId);
const lot = data.lots[lotIdx];
if (!lot || !lot.cheques[chIdx]) { closeChequeEditModal(); return; }
const ch = lot.cheques[chIdx];
ch.chequeNo = document.getElementById('ce_chequeNo').value.trim();
const dateVal = document.getElementById('ce_chequeDate').value;
if (dateVal) { const d = new Date(dateVal); if (!isNaN(d)) ch.chequeDate = d; }
closeChequeEditModal();
persistLotsForEntry(entryId, () => {
showToast('Cheque details saved', false);
const entry = entries.find(x => x.id === entryId);
if (entry && editingId === entryId) renderChequeSchedule(entry);
});
}
function openLotModalForLot(entryId, lotIdx) { if (editingId !== entryId) { showToast('Please open this entry in edit mode first', true); return; } openLotModal(); const select = document.getElementById('lm_lotSelect'); if (select) { select.value = String(lotIdx); onLotSelectChange(); } }

/* ============ Delete a Lot ============ */
function deleteLot(entryId, lotIdx) {
const data = getLotData(entryId);
const lot = data.lots[lotIdx];
if (!lot) { showToast('Lot not found', true); return; }
const count = (lot.cheques || []).length;
const tag = lot.isTopUp ? ' [TOP-UP]' : '';
if (!confirm(`Delete Lot ${lot.lotNumber}${tag}${lot.payeeName ? ' (' + lot.payeeName + ')' : ''} with ${count} cheque(s)?\n\nThis cannot be undone.`)) return;

const deletedLotNum = lot.lotNumber;

data.lots.splice(lotIdx, 1);

data.lots.forEach((l, i) => { l.lotNumber = i + 1; });

let globalSno = 1;
data.lots.forEach(l => { (l.cheques || []).forEach(ch => { ch.sno = globalSno++; }); });

const entry = entries.find(x => x.id === entryId);
if (!entry) return;
const serial = serializeLots(entryId);
const allCheques = getAllChequesFlat(entryId);
const lastCheque = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;
const updated = { ...entry, issuedCheques: serial, chequeNo: lastCheque && lastCheque.chequeNo ? lastCheque.chequeNo : '' };

updateEntryInFirebase(updated).then(() => {
if (editingId === entryId) fillForm(updated);
renderChequeSchedule(updated);
showToast(`✓ Lot ${deletedLotNum} deleted`, false);
}).catch(() => {});
}

function updateChequeNo(entryId, lotIdx, chIdx, newVal) { const data = getLotData(entryId); if (!data.lots[lotIdx] || !data.lots[lotIdx].cheques[chIdx]) return; data.lots[lotIdx].cheques[chIdx].chequeNo = (newVal || '').trim(); persistLotsForEntry(entryId, () => showToast('Cheque No. updated', false)); }
function updateLotControlNo(entryId, lotIdx) {
const inputId = 'lotControl_' + entryId + '_' + lotIdx;
const input = document.getElementById(inputId);
if (!input) return;
const newVal = input.value.trim();
if (!newVal) { showToast('Control No. cannot be empty', true); return; }
const data = getLotData(entryId);
if (!data.lots[lotIdx]) return;
data.lots[lotIdx].controlNo = newVal;
persistLotsForEntry(entryId, () => showToast('Control No. updated for Lot ' + data.lots[lotIdx].lotNumber, false));
}
function copySingleNarration(entryId, lotIdx, chIdx) {
const data = getLotData(entryId);
const lot = data.lots[lotIdx];
if (!lot || !lot.cheques[chIdx]) return;
const r = lot.cheques[chIdx];
const narration = `CH NO. ${r.chequeNo||'---'} DTD ${fmtDateSlash(r.chequeDate)} ISSUED TOWARDS AR FOR THE PERIOD FROM ${fmtDateSlash(r.from)} TO ${fmtDateSlash(r.to)} VIDE CONTROL NO ${lot.controlNo || ''}`;
copyText(narration);
}
function copyAllNarrationsFromForm(entryId) {
const data = getLotData(entryId);
const narrations = [];
(data.lots || []).forEach(lot => { (lot.cheques || []).forEach(r => { narrations.push(`CH NO. ${r.chequeNo||'---'} DTD ${fmtDateSlash(r.chequeDate)} ISSUED TOWARDS AR FOR THE PERIOD FROM ${fmtDateSlash(r.from)} TO ${fmtDateSlash(r.to)} VIDE CONTROL NO ${lot.controlNo || ''}`); }); });
if (narrations.length === 0) { showToast('No cheques to copy', true); return; }
copyText(narrations.join('\n'));
}
function copyText(text) {
if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(text).then(() => showToast('Copied to clipboard', false)).catch(() => fallbackCopy(text)); }
else fallbackCopy(text);
}
function fallbackCopy(text) {
const ta = document.createElement('textarea'); ta.value = text;
document.body.appendChild(ta); ta.select();
try { document.execCommand('copy'); showToast('Copied to clipboard', false); } catch (e) { showToast('Copy failed', true); }
document.body.removeChild(ta);
}

function serializeLots(entryId) {
const data = getLotData(entryId);
return (data.lots || []).map(lot => ({
lotNumber: lot.lotNumber,
issuedDate: lot.issuedDate instanceof Date ? lot.issuedDate.toISOString() : lot.issuedDate,
controlNo: lot.controlNo || '',
payeeName: lot.payeeName || '',
payeeSharePct: lot.payeeSharePct || null,
payeeFavouring: lot.payeeFavouring || '',
payeePan: lot.payeePan || '',
isTopUp: lot.isTopUp || false,
cheques: (lot.cheques || []).map(r => ({
sno: r.sno, chequeNo: r.chequeNo || '',
chequeDate: r.chequeDate instanceof Date ? r.chequeDate.toISOString() : r.chequeDate,
from: r.from instanceof Date ? r.from.toISOString() : r.from,
to: r.to instanceof Date ? r.to.toISOString() : r.to,
gross: r.gross, tds: r.tds, net: r.net,
favouring: r.favouring || '',
payeeName: r.payeeName || '',
payeeSharePct: r.payeeSharePct || null,
isTopUp: r.isTopUp || false
}))
}));
}
function deserializeLots(serialized) {
if (!serialized || !Array.isArray(serialized)) return { lots: [] };
return { lots: serialized.map(lot => ({
lotNumber: lot.lotNumber || 1,
issuedDate: lot.issuedDate ? new Date(lot.issuedDate) : new Date(),
controlNo: lot.controlNo || '',
payeeName: lot.payeeName || '',
payeeSharePct: lot.payeeSharePct || null,
payeeFavouring: lot.payeeFavouring || '',
payeePan: lot.payeePan || '',
isTopUp: lot.isTopUp || false,
cheques: (lot.cheques || []).map(r => ({
sno: r.sno, chequeNo: r.chequeNo || '',
chequeDate: r.chequeDate ? new Date(r.chequeDate) : new Date(),
from: r.from ? new Date(r.from) : new Date(),
to: r.to ? new Date(r.to) : new Date(),
gross: r.gross || 0, tds: r.tds || 0, net: r.net || 0,
favouring: r.favouring || '',
payeeName: r.payeeName || '',
payeeSharePct: r.payeeSharePct || null,
isTopUp: r.isTopUp || false
}))
})) };
}
function getAllChequesFlat(entryId) {
const data = getLotData(entryId);
const all = [];
data.lots.forEach(lot => { (lot.cheques || []).forEach(r => { all.push({ ...r, controlNo: lot.controlNo, payeeName: lot.payeeName || r.payeeName, payeeSharePct: lot.payeeSharePct || r.payeeSharePct }); }); });
return all;
}

/* ============ STORAGE ============ */
function processEntriesData(data) {
if (isSyncing) return;
isSyncing = true;
if (data) {
const loaded = Object.values(data);
loaded.sort((a, b) => a.id - b.id);
entries = loaded;
entries.forEach(e => { if (e.id >= idSeq) idSeq = e.id + 1; });
entries.forEach(e => { for (const [key, val] of Object.entries(data)) { if (val.id === e.id) { e._firebaseKey = key; break; } } });
issuedSchedules = {};
entries.forEach(e => {
if (e.issuedCheques) {
if (Array.isArray(e.issuedCheques) && e.issuedCheques.length > 0 && e.issuedCheques[0].lotNumber !== undefined) {
issuedSchedules[e.id] = deserializeLots(e.issuedCheques);
}
else if (Array.isArray(e.issuedCheques) && e.issuedCheques.length > 0) {
const oldCheques = e.issuedCheques.map((r, idx) => ({ sno: idx + 1, chequeNo: r.chequeNo || '', chequeDate: r.chequeDate ? new Date(r.chequeDate) : new Date(), from: r.from ? new Date(r.from) : new Date(), to: r.to ? new Date(r.to) : new Date(), gross: r.gross || 0, tds: r.tds || 0, net: r.net || 0 }));
issuedSchedules[e.id] = { lots: [{ lotNumber: 1, issuedDate: new Date(), controlNo: e.controlNo || '', cheques: oldCheques }] };
const serial = serializeLots(e.id);
updateEntryInFirebase({ ...e, issuedCheques: serial }).catch(() => {});
}
else { issuedSchedules[e.id] = { lots: [] }; }
} else { issuedSchedules[e.id] = { lots: [] }; }
});
} else { entries = []; issuedSchedules = {}; }
renderEntriesTable();
renderDashboard();
if (editingId !== null) { const curr = entries.find(e => e.id === editingId); if (curr) { const c = computeEntryValues(curr); showChequeSection(curr, c); } }
document.getElementById('entryCountBadge').textContent = entries.length;
isSyncing = false;
}

function loadEntriesFromFirebase() {
if (usingFirebase) { database.ref('entries').on('value', (snapshot) => { processEntriesData(snapshot.val()); }, (error) => { showToast('Error loading from Firebase: ' + error.message, true); }); }
else { processEntriesData(lsGetAll()); }
}

function saveEntryToFirebase(entry) {
return new Promise((resolve, reject) => {
if (usingFirebase) {
const ref = database.ref('entries').push();
ref.set(entry).then(() => { entry._firebaseKey = ref.key; resolve(entry); }).catch((err) => { showToast('Error saving: ' + err.message, true); reject(err); });
} else {
try {
const data = lsGetAll();
const key = 'local_' + entry.id + '_' + Date.now();
entry._firebaseKey = key;
data[key] = entry;
lsSetAll(data);
processEntriesData(data);
resolve(entry);
} catch (err) { showToast('Error saving locally: ' + err.message, true); reject(err); }
}
});
}

function updateEntryInFirebase(entry) {
return new Promise((resolve, reject) => {
if (usingFirebase) {
const key = entry._firebaseKey;
if (key) { database.ref('entries/' + key).update(entry).then(() => resolve(entry)).catch((err) => { showToast('Error updating: ' + err.message, true); reject(err); }); }
else {
database.ref('entries').orderByChild('id').equalTo(entry.id).once('value').then((snapshot) => {
const data = snapshot.val();
if (data) { const key2 = Object.keys(data)[0]; entry._firebaseKey = key2; return database.ref('entries/' + key2).update(entry).then(() => resolve(entry)); }
else reject('Entry not found');
}).catch((err) => { showToast('Error: ' + err.message, true); reject(err); });
}
} else {
try {
const data = lsGetAll();
let key = entry._firebaseKey;
if (!key || !data[key]) { key = Object.keys(data).find(k => data[k].id === entry.id); }
if (!key) { reject('Entry not found'); return; }
entry._firebaseKey = key;
data[key] = entry;
lsSetAll(data);
processEntriesData(data);
resolve(entry);
} catch (err) { showToast('Error updating locally: ' + err.message, true); reject(err); }
}
});
}

function deleteEntryFromFirebase(id, silent) {
return new Promise((resolve, reject) => {
if (usingFirebase) {
database.ref('entries').orderByChild('id').equalTo(id).once('value').then((snapshot) => {
const data = snapshot.val();
if (data) { const key = Object.keys(data)[0]; database.ref('entries/' + key).remove().then(() => { if (!silent) showToast('Entry deleted.', false); resolve(); }).catch((err) => { if (!silent) showToast('Error: ' + err.message, true); reject(err); }); }
else reject('Entry not found');
}).catch((err) => { if (!silent) showToast('Error: ' + err.message, true); reject(err); });
} else {
try {
const data = lsGetAll();
const key = Object.keys(data).find(k => data[k].id === id);
if (key) { delete data[key]; lsSetAll(data); processEntriesData(data); if (!silent) showToast('Entry deleted.', false); resolve(); }
else reject('Entry not found');
} catch (err) { if (!silent) showToast('Error: ' + err.message, true); reject(err); }
}
});
}

function clearAllFirebase() {
return new Promise((resolve, reject) => {
if (usingFirebase) { database.ref('entries').remove().then(() => { showToast('All data cleared.', false); resolve(); }).catch((err) => { showToast('Error: ' + err.message, true); reject(err); }); }
else { try { localStorage.removeItem(LS_KEY); showToast('All data cleared (local).', false); processEntriesData({}); resolve(); } catch (err) { showToast('Error: ' + err.message, true); reject(err); } }
});
}

function saveEntry() {
const data = readForm();
if (!data.client) { showToast('Client Name is required', true); document.getElementById('f_client').focus(); return; }
if (!data.unitNo) { showToast('Unit No. is required', true); document.getElementById('f_unitNo').focus(); return; }
const dup = checkUnitNoDuplicate(data.unitNo, editingId);
if (dup) { const reason = promptForDuplicateReason(data.unitNo, dup.client); if (reason === null) { return; } data.duplicateReason = reason.trim() || 'No reason provided'; }
else { data.duplicateReason = null; }
if (editingId) {
const existing = entries.find(e => e.id === editingId);
if (existing) {
const updated = { ...existing, ...data };
updated._firebaseKey = existing._firebaseKey;
if (issuedSchedules[editingId]) updated.issuedCheques = serializeLots(editingId);
updateEntryInFirebase(updated).then(() => { showToast('Entry updated', false); cancelEdit(); switchTab('tableView'); }).catch(() => {});
}
} else {
data.id = idSeq++;
saveEntryToFirebase(data).then(() => { showToast('Entry saved', false); cancelEdit(); switchTab('tableView'); }).catch(() => {});
}
}

function editEntry(id) {
const e = entries.find(x => x.id === id);
if (!e) return;
editingId = id;
editingFirebaseKey = e._firebaseKey;
fillForm(e);
document.getElementById('formTitle').textContent = 'Edit Entry' + (e.isImported ? ' (Imported)' : '');
document.getElementById('editingBadge').style.display = 'inline-flex';
document.getElementById('editingIndex').textContent = entries.findIndex(x => x.id === id) + 1;
document.getElementById('duplicateBtn').style.display = 'inline-flex';
const c = computeEntryValues(e);
showChequeSection(e, c);
switchTab('formView');
}

function duplicateEntry(id) {
const e = entries.find(x => x.id === id);
if (!e) return;
const clone = JSON.parse(JSON.stringify(e));
delete clone._firebaseKey;
clone.id = idSeq++;
clone.payments = (clone.payments || []).map(p => ({ ...p }));
delete clone.issuedCheques;
delete clone.scheduledArMonthly;
clone.isImported = false;
const dup = checkUnitNoDuplicate(clone.unitNo || '', null);
if (dup) { const reason = promptForDuplicateReason(clone.unitNo, dup.client); if (reason === null) { showToast('Duplicate cancelled', true); return; } clone.duplicateReason = reason.trim() || 'No reason provided'; }
else { clone.duplicateReason = null; }
saveEntryToFirebase(clone).then(() => showToast('Entry duplicated', false)).catch(() => {});
}

function duplicateFromForm() {
const data = readForm();
if (!data.client) { showToast('Client Name is required', true); return; }
if (!data.unitNo) { showToast('Unit No. is required', true); return; }
data.isImported = false;
const dup = checkUnitNoDuplicate(data.unitNo, null);
if (dup) { const reason = promptForDuplicateReason(data.unitNo, dup.client); if (reason === null) return; data.duplicateReason = reason.trim() || 'No reason provided'; }
else { data.duplicateReason = null; }
data.id = idSeq++;
saveEntryToFirebase(data).then(() => { showToast('Entry duplicated', false); cancelEdit(); switchTab('tableView'); }).catch(() => {});
}

function deleteEntry(id) {
pendingDeleteId = id;
const e = entries.find(x => x.id === id);
document.getElementById('deleteModalBody').textContent = `"${e ? e.client : 'This entry'}" will be permanently removed.`;
document.getElementById('deleteModal').classList.add('show');
}
function closeDeleteModal() { document.getElementById('deleteModal').classList.remove('show'); pendingDeleteId = null; }
function confirmDelete() { if (pendingDeleteId !== null) { deleteEntryFromFirebase(pendingDeleteId).then(() => closeDeleteModal()).catch(() => closeDeleteModal()); } }
function openResetModal() { document.getElementById('resetModal').classList.add('show'); }
function closeResetModal() { document.getElementById('resetModal').classList.remove('show'); }
function confirmReset() {
clearAllFirebase().then(() => {
closeResetModal();
entries = []; idSeq = 1; issuedSchedules = {}; selectedEntryIds.clear();
cancelEdit(); renderEntriesTable(); renderDashboard();
document.getElementById('entryCountBadge').textContent = '0';
}).catch(() => {});
}
function markAllExistingAsImported() {
if (entries.length === 0) { showToast('No entries to mark', true); return; }
if (!confirm(`Mark ALL ${entries.length} existing entries as "Imported"?`)) return;
let done = 0;
const promises = entries.map(e => { if (!e.isImported) { e.isImported = true; return updateEntryInFirebase(e).then(() => { done++; }).catch(() => {}); } return Promise.resolve(); });
Promise.all(promises).then(() => { showToast(`✓ ${done} entries marked as Imported`, false); renderEntriesTable(); renderDashboard(); });
}

function toggleEntrySelection(id, checked) { if (checked) selectedEntryIds.add(id); else selectedEntryIds.delete(id); updateBulkDeleteUI(); }
function toggleSelectAll() {
const filtered = getFilteredEntries();
const allSelected = filtered.length > 0 && filtered.every(e => selectedEntryIds.has(e.id));
if (allSelected) { filtered.forEach(e => selectedEntryIds.delete(e.id)); }
else { filtered.forEach(e => selectedEntryIds.add(e.id)); }
renderEntriesTable();
}
function clearSelection() { selectedEntryIds.clear(); renderEntriesTable(); }

function getFilteredEntries() {
const search = document.getElementById('searchInput').value.toLowerCase().trim();
const filterCompany = document.getElementById('filterCompany')?.value || '';
const filterStatus = document.getElementById('filterStatus')?.value || '';
const filterType = document.getElementById('filterType')?.value || '';
return entries.filter(e => {
if (search) {
const matchesSearch = (e.client || '').toLowerCase().includes(search) || (e.cp || '').toLowerCase().includes(search) || (e.unitNo || '').toLowerCase().includes(search) || (e.pan || '').toLowerCase().includes(search) || (e.favouring || '').toLowerCase().includes(search) || (e.rm || '').toLowerCase().includes(search) || (e.company || '').toLowerCase().includes(search) || (e.controlNo || '').toLowerCase().includes(search);
if (!matchesSearch) return false;
}
if (filterCompany && (e.company || 'BEPL') !== filterCompany) return false;
if (filterType === 'manual' && e.isImported) return false;
if (filterType === 'imported' && !e.isImported) return false;
if (filterStatus) {
const c = computeEntryValues(e);
const hasSchedule = issuedSchedules[e.id] && issuedSchedules[e.id].lots && issuedSchedules[e.id].lots.length > 0;
if (filterStatus === 'eligible' && !c.eligible) return false;
if (filterStatus === 'not-eligible' && c.eligible) return false;
if (filterStatus === 'imported' && !e.isImported) return false;
if (filterStatus === 'scheduled' && !hasSchedule) return false;
}
return true;
});
}
function resetFilters() { document.getElementById('filterCompany').value = ''; document.getElementById('filterStatus').value = ''; document.getElementById('filterType').value = ''; document.getElementById('searchInput').value = ''; renderEntriesTable(); }

function updateBulkDeleteUI() {
const count = selectedEntryIds.size;
const bulkBar = document.getElementById('bulkBar');
const cnt = document.getElementById('selectedCount');
const selAllCheckbox = document.getElementById('selectAllCheckbox');
const selAllBtn = document.getElementById('selectAllBtn');
if (cnt) cnt.textContent = count;
if (bulkBar) bulkBar.classList.toggle('hidden', count === 0);
const filtered = getFilteredEntries();
const allSelected = filtered.length > 0 && filtered.every(e => selectedEntryIds.has(e.id));
if (selAllCheckbox) selAllCheckbox.checked = allSelected;
if (selAllBtn) selAllBtn.innerHTML = allSelected ? '<i class="fa-solid fa-square"></i> Deselect All' : '<i class="fa-solid fa-check-square"></i> Select All';
}

function bulkDeleteEntries() {
const count = selectedEntryIds.size;
if (count === 0) { showToast('No entries selected', true); return; }
if (!confirm(`Delete ${count} selected entr${count===1?'y':'ies'}? This cannot be undone.`)) return;
const ids = Array.from(selectedEntryIds);
showToast(`Deleting ${count} entries...`, false);
const promises = ids.map(id => deleteEntryFromFirebase(id, true).catch(() => {}));
Promise.all(promises).then(() => { selectedEntryIds.clear(); renderEntriesTable(); renderDashboard(); showToast(`✓ ${count} entries deleted`, false); });
}

function bulkChangeCompany() {
const count = selectedEntryIds.size;
if (count === 0) { showToast('No entries selected', true); return; }
const newCompany = prompt(`Change company for ${count} selected entries to:\n\nEnter "BEPL" or "BSCIPL":`);
if (!newCompany) return;
const company = newCompany.trim().toUpperCase();
if (company !== 'BEPL' && company !== 'BSCIPL') { showToast('Invalid company. Use BEPL or BSCIPL', true); return; }
const ids = Array.from(selectedEntryIds);
let done = 0;
const promises = ids.map(id => { const e = entries.find(x => x.id === id); if (!e) return Promise.resolve(); e.company = company; return updateEntryInFirebase(e).then(() => { done++; }).catch(() => {}); });
Promise.all(promises).then(() => { showToast(`✓ ${done} entries changed to ${company}`, false); renderEntriesTable(); renderDashboard(); });
}

function bulkToggleImported() {
const count = selectedEntryIds.size;
if (count === 0) { showToast('No entries selected', true); return; }
const ids = Array.from(selectedEntryIds);
let done = 0;
const promises = ids.map(id => { const e = entries.find(x => x.id === id); if (!e) return Promise.resolve(); e.isImported = !e.isImported; return updateEntryInFirebase(e).then(() => { done++; }).catch(() => {}); });
Promise.all(promises).then(() => { showToast(`✓ ${done} entries toggled imported status`, false); renderEntriesTable(); renderDashboard(); });
}

function renderEntriesTable() {
const tbody = document.getElementById('entriesTbody');
const empty = document.getElementById('tableEmpty');
document.getElementById('entryCountBadge').textContent = entries.length;
const filtered = getFilteredEntries();
if (filtered.length === 0) { tbody.innerHTML = ''; empty.style.display = 'block'; updateBulkDeleteUI(); return; }
empty.style.display = 'none';
tbody.innerHTML = filtered.map((e, i) => {
const c = computeEntryValues(e);
let statusBadge;
if (e.isImported) { if (c.eligible) statusBadge = `<span class="badge badge-green"><i class="fa-solid fa-check"></i> AR Eligible</span>`; else statusBadge = `<span class="badge badge-red"><i class="fa-solid fa-clock"></i> Due</span>`; }
else if (c.eligible) statusBadge = `<span class="badge badge-green"><i class="fa-solid fa-check"></i> AR Eligible</span>`;
else if (c.balance <= 0 && c.due === 0) statusBadge = `<span class="badge badge-amber"><i class="fa-solid fa-minus"></i> N/A</span>`;
else statusBadge = `<span class="badge badge-red"><i class="fa-solid fa-clock"></i> Due</span>`;
const importedTag = e.isImported ? `<span class="badge-imported"><i class="fa-solid fa-file-import"></i> Imported</span>` : '';
const isChecked = selectedEntryIds.has(e.id) ? 'checked' : '';
const hasSchedule = issuedSchedules[e.id] && issuedSchedules[e.id].lots && issuedSchedules[e.id].lots.length > 0;
const scheduleTag = hasSchedule ? ` <span class="badge-schedule"><i class="fa-solid fa-file-invoice"></i> ${getTotalCheques(e.id)} chq</span>` : '';
const splitTag = c.hasPayeeSplit ? ` <span class="badge-split"><i class="fa-solid fa-people-split"></i> ${c.payeeSplits.length} payees</span>` : '';
const dupTag = e.duplicateReason ? ` <span class="badge-dup" title="${escapeHtml(e.duplicateReason)}">DUP</span>` : '';
const manualArTag = c.hasManualArMonthly ? ` <span class="badge-manual">MANUAL AR</span>` : '';
const incTag = detectArIncrease(e) ? ` <span class="badge-topup"><i class="fa-solid fa-arrow-trend-up"></i> AR ↑</span>` : '';
return `<tr class="${isChecked?'selected':''}">
<td class="checkbox-col"><input type="checkbox" class="entry-checkbox" data-entry-id="${e.id}" ${isChecked} onchange="toggleEntrySelection(${e.id}, this.checked)"></td>
<td>${entries.indexOf(e)+1}</td>
<td><span class="badge-company">${escapeHtml(e.company||'BEPL')}</span></td>
<td>${escapeHtml(e.unitNo||'—')}</td>
<td>${escapeHtml(e.client)} ${importedTag}${scheduleTag}${splitTag}${dupTag}${incTag}</td>
<td>${escapeHtml(e.cp||'—')}</td>
<td>${escapeHtml(e.description||'—')}</td>
<td class="num">${e.area||0}</td>
<td class="num">${(e.rate||0).toLocaleString('en-IN')}</td>
<td class="num">${fmt(c.gross)}</td>
<td class="num">${e.duePct}%</td>
<td class="num">${fmt(c.due)}</td>
<td class="num">${fmt(c.received)}</td>
<td class="num">${fmt(c.balance)}</td>
<td>${statusBadge}</td>
<td class="num">${c.period}</td>
<td class="num">${fmt(c.arMonthly)}${manualArTag}</td>
<td>${escapeHtml(e.pan||'')}</td>
<td>${escapeHtml(e.favouring||'')}</td>
<td>${escapeHtml(e.chequeNo||'')}</td>
<td>${escapeHtml(e.controlNo||'')}</td>
<td class="actions-cell">
<button title="Edit" onclick="editEntry(${e.id})"><i class="fa-solid fa-pen"></i></button>
<button title="Print / PDF" onclick="printEntry(${e.id})"><i class="fa-solid fa-print"></i></button>
<button title="Duplicate" onclick="duplicateEntry(${e.id})"><i class="fa-solid fa-copy"></i></button>
<button title="Delete" class="del" onclick="deleteEntry(${e.id})"><i class="fa-solid fa-trash"></i></button>
</td></tr>`;
}).join('');
updateBulkDeleteUI();
}

function renderDashboard() {
const computed = entries.map(e => ({ e, c: computeEntryValues(e) }));
const eligible = computed.filter(x => x.c.eligible || x.e.isImported);
const totalARValue = eligible.reduce((s, x) => s + x.c.due, 0);
const arYearly = totalARValue * 0.12;
const period = eligible.length ? eligible[0].c.period : 12;
const totalMonthlyAR = eligible.reduce((s, x) => s + x.c.arMonthly, 0);
const tds = round0(totalMonthlyAR * 0.10);
const net = round0(totalMonthlyAR - tds);
const totalPayable = round0(net * period);
let allDates = [];
entries.forEach(e => (e.payments || []).forEach(p => { if (p.date) { const d = new Date(p.date); if (!isNaN(d)) allDates.push(d); } }));
let startDate = new Date();
if (allDates.length) { startDate = new Date(Math.max(...allDates.map(d => d.getTime()))); startDate.setDate(startDate.getDate() + 1); }
const increasedCount = entries.filter(e => detectArIncrease(e)).length;
document.getElementById('dashInfoStrip').innerHTML = `
<div><div class="lbl">AR Start Date</div><div class="val">${fmtDate(startDate)}</div></div>
<div><div class="lbl">AR Period Used</div><div class="val">${period} months</div></div>
<div><div class="lbl">AR-Eligible</div><div class="val">${eligible.length} of ${entries.length}</div></div>
<div><div class="lbl">Imported</div><div class="val">${entries.filter(e=>e.isImported).length}</div></div>
<div><div class="lbl">BEPL</div><div class="val">${entries.filter(e=>(e.company||'BEPL')==='BEPL').length}</div></div>
<div><div class="lbl">BSCIPL</div><div class="val">${entries.filter(e=>e.company==='BSCIPL').length}</div></div>
<div><div class="lbl">With Schedule</div><div class="val">${entries.filter(e=>getTotalCheques(e.id) > 0).length}</div></div>
<div><div class="lbl">Split Payees</div><div class="val">${entries.filter(e=>computeEntryValues(e).hasPayeeSplit).length}</div></div>
<div><div class="lbl">AR Increased</div><div class="val">${increasedCount}</div></div>`;
document.getElementById('dashGrid').innerHTML = `
<div class="dash-card emph"><div class="lbl">Total AR Value</div><div class="val">${fmt(totalARValue)}</div><div class="sub">Sum of due amounts</div></div>
<div class="dash-card"><div class="lbl">AR @ 12% (Yearly)</div><div class="val">${fmt(arYearly)}</div><div class="sub">Total × 12%</div></div>
<div class="dash-card"><div class="lbl">Total Monthly AR</div><div class="val">${fmt(totalMonthlyAR)}</div><div class="sub">1% of Received + Discount</div></div>
<div class="dash-card"><div class="lbl">TDS @ 10%</div><div class="val">${fmt(tds)}</div><div class="sub">10% of Monthly AR</div></div>
<div class="dash-card"><div class="lbl">Net Payable (Monthly)</div><div class="val">${fmt(net)}</div><div class="sub">Monthly AR − TDS</div></div>
<div class="dash-card emph"><div class="lbl">Total Payable</div><div class="val">${fmt(totalPayable)}</div><div class="sub">Net × ${period} months</div></div>`;
}

function addMonths(date, n) { const d = new Date(date); d.setMonth(d.getMonth() + n); return d; }

function addMonthsAnchored(baseDate, n) {
    const d = new Date(baseDate);
    const targetMonth = d.getMonth() + n;
    const targetYear = d.getFullYear() + Math.floor(targetMonth / 12);
    const normalizedMonth = (targetMonth % 12 + 12) % 12;
    const originalDay = d.getDate();
    const daysInTargetMonth = new Date(targetYear, normalizedMonth + 1, 0).getDate();
    const targetDay = Math.min(originalDay, daysInTargetMonth);
    return new Date(targetYear, normalizedMonth, targetDay);
}
function addDays(date, n) { const d = new Date(date); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { if (!(d instanceof Date) || isNaN(d)) return '—'; return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }); }
function fmtDateSlash(d) { if (!(d instanceof Date) || isNaN(d)) return ''; const dd = String(d.getDate()).padStart(2, '0'); const mm = String(d.getMonth() + 1).padStart(2, '0'); return `${dd}/${mm}/${d.getFullYear()}`; }
function dayCount(d1, d2) { return Math.round((d2.getTime() - d1.getTime()) / 86400000); }
function incrementCode(base, offset) {
if (!base) return ''; base = base.toString().trim(); if (!base) return '';
const m = base.match(/^(.*?)(\d+)$/);
if (!m) return offset > 0 ? (base + '-' + (offset + 1)) : base;
const prefix = m[1], digits = m[2];
return prefix + (parseInt(digits, 10) + offset).toString().padStart(digits.length, '0');
}
function getNextChequeNumber(entryId) {
const allCheques = getAllChequesFlat(entryId);
if (allCheques.length === 0) return null;
const last = allCheques[allCheques.length - 1];
if (!last.chequeNo) return null;
const next = incrementCode(last.chequeNo, 1);
if (next === last.chequeNo) { const numVal = parseInt(last.chequeNo, 10); if (!isNaN(numVal)) return (numVal + 1).toString().padStart(last.chequeNo.length, '0'); return last.chequeNo + '-1'; }
return next;
}
function getNextLotNumber(entryId) { const lots = getLotData(entryId).lots || []; if (lots.length === 0) return 1; return Math.max(...lots.map(l => l.lotNumber || 0)) + 1; }
function getLastPeriodEnd(entryId) {
const allCheques = getAllChequesFlat(entryId);
if (allCheques.length === 0) return null;
const last = allCheques[allCheques.length - 1];
if (!last.to) return null;
const d = new Date(last.to);
return isNaN(d) ? null : d;
}

function openLotModal() {
if (editingId === null) { showToast('Please save the entry first', true); return; }
const e = entries.find(x => x.id === editingId);
if (!e) { showToast('Entry not found', true); return; }
const c = computeEntryValues(e);
if (!c.eligible && !e.isImported) { showToast('This entry is not AR eligible yet', true); return; }
lotModalEntryId = editingId;
document.getElementById('lm_clientDisplay').value = e.client;
const data = getLotData(editingId);
const lots = data.lots || [];
const select = document.getElementById('lm_lotSelect');
select.innerHTML = '<option value="new">+ New Lot</option>';
lots.forEach((lot, idx) => { const opt = document.createElement('option'); opt.value = idx; opt.textContent = 'Lot ' + lot.lotNumber + (lot.payeeName ? ' — ' + lot.payeeName : '') + (lot.isTopUp ? ' [TOP-UP]' : '') + ' (' + (lot.cheques ? lot.cheques.length : 0) + ' cheques)'; select.appendChild(opt); });
select.value = 'new';
const incCb = document.getElementById('lm_incMode');
if (incCb) incCb.checked = false;
const incDet = document.getElementById('lm_incDetails');
if (incDet) incDet.style.display = 'none';
onLotSelectChange();
document.getElementById('lotModal').classList.add('show');
}
function closeLotModal() { document.getElementById('lotModal').classList.remove('show'); lotModalEntryId = null; }

function onLotSelectChange() {
const select = document.getElementById('lm_lotSelect');
const selected = select.value;
const entryId = lotModalEntryId;
if (!entryId) return;
const data = getLotData(entryId);
const lots = data.lots || [];

const incSection = document.getElementById('lm_incSection');
const incCheckbox = document.getElementById('lm_incMode');
const incDetails = document.getElementById('lm_incDetails');
if (incCheckbox) incCheckbox.checked = false;
if (incDetails) incDetails.style.display = 'none';
if (selected === 'new') {
const inc = detectArIncrease(entries.find(e => e.id === entryId));
if (incSection) incSection.style.display = inc ? 'block' : 'none';
} else {
if (incSection) incSection.style.display = 'none';
}

if (selected === 'new') {
const c = computeEntryValues(entries.find(e => e.id === entryId));
const totalCheques = getTotalCheques(entryId);
const period = c.basePeriod;
document.getElementById('lm_count').value = Math.max(1, period - totalCheques);
document.getElementById('lm_chequeNo').value = getNextChequeNumber(entryId) || '';
let lastControl = '';
if (lots.length > 0) lastControl = lots[lots.length - 1].controlNo || '';
document.getElementById('lm_controlNo').value = lastControl || entries.find(e => e.id === entryId).controlNo || '';
const lastEnd = getLastPeriodEnd(entryId);
let defaultDate = new Date();
if (lastEnd) { const suggested = addDays(lastEnd, 1); if (!isNaN(suggested)) defaultDate = suggested; }
else defaultDate = addDays(c.startDate, 30);
document.getElementById('lm_firstDate').value = defaultDate.toISOString().slice(0, 10);
document.getElementById('lm_infoText').textContent = 'Creating a new lot.';
document.getElementById('lm_actionBtnText').textContent = 'Create Lot';
} else {
const lotIdx = parseInt(selected);
const lot = lots[lotIdx];
if (!lot) return;
document.getElementById('lm_count').value = lot.cheques ? lot.cheques.length : 0;
document.getElementById('lm_chequeNo').value = lot.cheques && lot.cheques.length > 0 ? lot.cheques[0].chequeNo : '';
document.getElementById('lm_controlNo').value = lot.controlNo || '';
const firstDate = lot.cheques && lot.cheques.length > 0 ? lot.cheques[0].chequeDate : new Date();
document.getElementById('lm_firstDate').value = firstDate instanceof Date ? firstDate.toISOString().slice(0, 10) : '';
document.getElementById('lm_infoText').textContent = 'Updating Lot ' + lot.lotNumber + (lot.payeeName ? ' (' + lot.payeeName + ')' : '') + '. Sirf wahi fields change honge jo aap edit karoge — baaki data safe rahega.';
document.getElementById('lm_actionBtnText').textContent = 'Update Lot';
}
updateLotPreview();
}

function updateLotPreview() {
const count = Math.max(1, parseInt(document.getElementById('lm_count').value) || 1);
const chequeStart = document.getElementById('lm_chequeNo').value.trim();
const controlStart = document.getElementById('lm_controlNo').value.trim();
let displayStart = chequeStart || '(auto)';
let displayEnd = chequeStart ? incrementCode(chequeStart, count - 1) : '(auto)';
if (!chequeStart && lotModalEntryId) { const next = getNextChequeNumber(lotModalEntryId); if (next) { displayStart = next; displayEnd = incrementCode(next, count - 1); } }
document.getElementById('lm_chequeRange').textContent = `Cheque No. range: ${displayStart} → ${displayEnd} (${count} cheque${count>1?'s':''})`;
document.getElementById('lm_controlRange').textContent = controlStart ? `Control No. (fixed for lot): ${controlStart}` : 'Control No.: — (optional)';

const incCb = document.getElementById('lm_incMode');
if (incCb && incCb.checked && lotModalEntryId) {
const trig = document.getElementById('lm_incTriggerDate').value;
if (trig) {
const triggerDate = new Date(trig);
const e = entries.find(x => x.id === lotModalEntryId);
const inc = e ? detectArIncrease(e) : null;
if (inc) {
const dataObj = getLotData(lotModalEntryId);
const originalCheques = [];
(dataObj.lots || []).forEach(lot => {
if (lot.isTopUp) return;
(lot.cheques || []).forEach(ch => originalCheques.push(ch));
});
originalCheques.sort((a, b) => new Date(a.from) - new Date(b.from));

let firstOverlap = null;
let totalCheques = 0;
for (const chq of originalCheques) {
if (new Date(chq.to) >= triggerDate) {
if (!firstOverlap) firstOverlap = chq;
totalCheques++;
}
}

if (firstOverlap) {
const firstFrom = new Date(firstOverlap.from);
const firstTo = new Date(firstOverlap.to);
let topUpFrom = firstFrom;
if (triggerDate > firstFrom && triggerDate <= firstTo) topUpFrom = triggerDate;

document.getElementById('lm_count').value = totalCheques;

let partialAmtStr = '';
if (topUpFrom > firstFrom) {
const totalDays = dayCount(firstFrom, firstTo) + 1;
const remainingDays = dayCount(topUpFrom, firstTo) + 1;
const prorated = round0(inc.delta * (remainingDays / totalDays));
partialAmtStr = ` (prorated: ${fmt(prorated)})`;
}

document.getElementById('lm_incRangeInfo').innerHTML =
`Overlapping original periods: <strong>${totalCheques}</strong><br>` +
`1st Top-Up period: <strong>${fmtDateSlash(topUpFrom)} → ${fmtDateSlash(firstTo)}</strong>${partialAmtStr}<br>` +
`Remaining periods will mirror original lot exactly.`;
} else {
document.getElementById('lm_incRangeInfo').textContent = 'Trigger date is after all existing periods.';
document.getElementById('lm_count').value = 0;
}
}
}
}
}

function confirmLotAction() {
const entryId = lotModalEntryId;
const e = entries.find(x => x.id === entryId);
if (!e) return;

const select = document.getElementById('lm_lotSelect');
const selected = select.value;
const count = Math.max(1, parseInt(document.getElementById('lm_count').value) || 1);
const userChequeNo = document.getElementById('lm_chequeNo').value.trim();
const controlNo = document.getElementById('lm_controlNo').value.trim();
const firstDate = document.getElementById('lm_firstDate').value;

const c = computeEntryValues(e);
const data = getLotData(entryId);
const lots = data.lots || [];

// ============ NEW LOT ============
if (selected === 'new') {
// ============ AMOUNT INCREASE MODE — Top-Up Lot (Mirror Original Periods) ============
const incMode = document.getElementById('lm_incMode') && document.getElementById('lm_incMode').checked;
if (incMode) {
const inc = detectArIncrease(e);
if (!inc) { showToast('No AR increase detected', true); return; }
const triggerDateStr = document.getElementById('lm_incTriggerDate').value;
if (!triggerDateStr) { showToast('Please set Increase Trigger Date', true); return; }
const triggerDate = new Date(triggerDateStr);

const delta = inc.delta;
const lotNumber = getNextLotNumber(entryId);
let startChequeNo = userChequeNo || getNextChequeNumber(entryId) || '';

const dataObj = getLotData(entryId);
const originalCheques = [];
(dataObj.lots || []).forEach(lot => {
if (lot.isTopUp) return; // skip top-up lots
(lot.cheques || []).forEach(ch => originalCheques.push(ch));
});
originalCheques.sort((a, b) => new Date(a.from) - new Date(b.from));

let globalSno = getAllChequesFlat(entryId).length + 1;
const newCheques = [];

originalCheques.forEach(chq => {
const chqFrom = new Date(chq.from);
const chqTo = new Date(chq.to);

// Skip periods that ended before the trigger date
if (chqTo < triggerDate) return;

let topUpFrom = chqFrom;
if (triggerDate > chqFrom && triggerDate <= chqTo) {
topUpFrom = new Date(triggerDate); // Prorate first overlapping month
}

let amount = delta;
if (topUpFrom > chqFrom) {
const totalDays = dayCount(chqFrom, chqTo) + 1;
const remainingDays = dayCount(topUpFrom, chqTo) + 1;
amount = round0(delta * (remainingDays / totalDays));
}

if (amount > 0) {
const tds = round0(amount * 0.10);
const net = round0(amount - tds);
newCheques.push({
sno: globalSno++,
chequeNo: startChequeNo ? incrementCode(startChequeNo, newCheques.length) : '',
chequeDate: new Date(chq.chequeDate),
from: new Date(topUpFrom),
to: new Date(chqTo),
gross: amount,
tds, net,
favouring: e.favouring || '',
isTopUp: true
});
}
});

if (newCheques.length === 0) {
showToast('No overlapping periods to top-up.', true);
return;
}

lots.push({
lotNumber,
issuedDate: new Date(),
controlNo: controlNo || e.controlNo || '',
isTopUp: true,
cheques: newCheques
});

let gSno = 1;
lots.forEach(lot => { (lot.cheques || []).forEach(ch => { ch.sno = gSno++; }); });

const serial = serializeLots(entryId);
const allCheques = getAllChequesFlat(entryId);
const lastCheque = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;
const updated = { ...e, issuedCheques: serial, chequeNo: lastCheque && lastCheque.chequeNo ? lastCheque.chequeNo : e.chequeNo, scheduledArMonthly: inc.current };
updateEntryInFirebase(updated).then(() => {
closeLotModal();
fillForm(updated);
showToast(`✓ Top-Up Lot ${lotNumber} created (${newCheques.length} parallel cheques) — old lot untouched`, false);
}).catch(() => {});
return;
}

if (!firstDate) { showToast('Please select First Cheque Date', true); return; }

// ============ NORMAL NEW LOT ============
const mGross = c.arMonthly, mTds = c.tds, mNet = c.net;
const lotNumber = getNextLotNumber(entryId);
let monthFrom;
const lastEnd = getLastPeriodEnd(entryId);
if (lastEnd) monthFrom = addDays(lastEnd, 1);
else monthFrom = new Date(c.startDate);

let startChequeNo = userChequeNo || getNextChequeNumber(entryId) || '';
const existingCheques = getAllChequesFlat(entryId);
const newCheques = [];

let startIdx = (monthFrom.getFullYear() - c.startDate.getFullYear()) * 12 + (monthFrom.getMonth() - c.startDate.getMonth());
    if (startIdx < 0) startIdx = 0;
    for (let i = 0; i < count; i++) {
        const monthTo = addDays(addMonthsAnchored(c.startDate, startIdx + i + 1), -1);
        const chequeDate = (i === 0 && firstDate) ? new Date(firstDate) : addDays(monthTo, 1);
newCheques.push({
sno: existingCheques.length + i + 1,
chequeNo: startChequeNo ? incrementCode(startChequeNo, i) : '',
chequeDate: new Date(chequeDate),
from: new Date(monthFrom),
to: new Date(monthTo),
gross: mGross, tds: mTds, net: mNet,
favouring: e.favouring || ''
});
monthFrom = addDays(monthTo, 1);
}

lots.push({
lotNumber: lotNumber,
issuedDate: new Date(),
controlNo: controlNo || (lots.length > 0 ? lots[lots.length - 1].controlNo : '') || e.controlNo || '',
cheques: newCheques
});
}
// ============ UPDATE EXISTING LOT ============
else {
const lotIdx = parseInt(selected);
const lot = lots[lotIdx];
if (!lot) { showToast('Lot not found', true); return; }

const existingCheques = lot.cheques || [];
const oldCount = existingCheques.length;

const amountSnapshot = existingCheques.length > 0
? { gross: num(existingCheques[0].gross), tds: num(existingCheques[0].tds), net: num(existingCheques[0].net) }
: { gross: c.arMonthly, tds: c.tds, net: c.net };

if (count !== oldCount) {
if (count > oldCount) {
let monthFrom = oldCount > 0
? addDays(new Date(existingCheques[oldCount - 1].to), 1)
: new Date(c.startDate);

let nextChequeNo = '';
if (existingCheques.length > 0 && existingCheques[existingCheques.length - 1].chequeNo) {
nextChequeNo = incrementCode(existingCheques[existingCheques.length - 1].chequeNo, 1);
} else if (userChequeNo) {
nextChequeNo = incrementCode(userChequeNo, oldCount);
}

for (let i = oldCount; i < count; i++) {
            const monthTo = addDays(addMonthsAnchored(c.startDate, i + 1), -1);
existingCheques.push({
sno: i + 1,
chequeNo: nextChequeNo ? incrementCode(nextChequeNo, i - oldCount) : '',
chequeDate: addDays(monthTo, 1),
from: new Date(monthFrom),
to: new Date(monthTo),
gross: amountSnapshot.gross,
tds: amountSnapshot.tds,
net: amountSnapshot.net,
favouring: lot.payeeFavouring || e.favouring || '',
payeeName: lot.payeeName || '',
payeeSharePct: lot.payeeSharePct || null
});
monthFrom = addDays(monthTo, 1);
}
} else {
existingCheques.length = count;
}
lot.cheques = existingCheques;
}

if (userChequeNo) {
lot.cheques.forEach((ch, i) => {
ch.chequeNo = incrementCode(userChequeNo, i);
});
}

if (firstDate && lot.cheques.length > 0) {
const newFirstDate = new Date(firstDate);
const oldFirstDate = lot.cheques[0].chequeDate ? new Date(lot.cheques[0].chequeDate) : null;
if (!oldFirstDate || newFirstDate.getTime() !== oldFirstDate.getTime()) {
lot.cheques[0].chequeDate = newFirstDate;
}
}

if (controlNo && controlNo !== lot.controlNo) {
lot.controlNo = controlNo;
}

lot.issuedDate = new Date();
}

let globalSno = 1;
lots.forEach(lot => { (lot.cheques || []).forEach(ch => { ch.sno = globalSno++; }); });

const serial = serializeLots(entryId);
const allCheques = getAllChequesFlat(entryId);
const lastCheque = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;
const updated = { ...e, issuedCheques: serial, chequeNo: lastCheque && lastCheque.chequeNo ? lastCheque.chequeNo : e.chequeNo, controlNo: controlNo || e.controlNo || '' };
if (allCheques.length > (e.arPeriod || 12)) { updated.arPeriod = allCheques.length; document.getElementById('f_arPeriod').value = allCheques.length; }
if (selected === 'new') updated.scheduledArMonthly = c.arMonthly;
updateEntryInFirebase(updated).then(() => {
closeLotModal();
fillForm(updated);
showToast(selected === 'new' ? `Lot ${lots.length} created` : `Lot updated (amount safe)`, false);
}).catch(() => {});
}

/* =========================================================
PAYEE SPLIT MODAL
========================================================= */
function openPayeeModal() {
if (editingId === null) { showToast('Please save the entry first, then open it in edit mode', true); return; }
const e = entries.find(x => x.id === editingId);
if (!e) { showToast('Entry not found', true); return; }
formPayees = normalizePayees(e.payees).map(p => ({ id: 'py' + (payeeRowSeq++), ...p }));
if (formPayees.length === 0) {
formPayees = [
{ id: 'py' + (payeeRowSeq++), name: 'Payee A', sharePct: 50, pan: e.pan || '', controlNo: e.controlNo || '' },
{ id: 'py' + (payeeRowSeq++), name: 'Payee B', sharePct: 50, pan: e.pan || '', controlNo: e.controlNo || '' }
];
}
renderPayeeRows();
document.getElementById('payeeModal').classList.add('show');
}
function closePayeeModal() { document.getElementById('payeeModal').classList.remove('show'); }
function clearPayees() { if (!confirm('Remove all payee splits? Schedule will go back to a single payee.')) return; formPayees = []; renderPayeeRows(); }
function addPayeeRow() { formPayees.push({ id: 'py' + (payeeRowSeq++), name: '', sharePct: 0, pan: '', controlNo: '' }); renderPayeeRows(); }
function removePayeeRow(id) { formPayees = formPayees.filter(p => p.id !== id); renderPayeeRows(); }
function updatePayeeField(id, field, value) {
const p = formPayees.find(x => x.id === id);
if (!p) return;
p[field] = (field === 'sharePct') ? num(value) : value;
renderPayeeTotals();
}
function renderPayeeRows() {
const tbody = document.getElementById('payeeTbody');
if (formPayees.length === 0) {
tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-mute);padding:14px;font-size:12.5px;">No payees added — schedule will be single payee.</td></tr>`;
renderPayeeTotals();
return;
}
const e = entries.find(x => x.id === editingId);
const c = e ? computeEntryValues(e) : { arMonthly: 0 };
tbody.innerHTML = formPayees.map(p => {
const gross = round0((c.arMonthly || 0) * (num(p.sharePct) / 100));
const net = round0(gross - round0(gross * 0.10));
return `<tr>
<td><input type="text" value="${escapeHtml(p.name||'')}" placeholder="e.g. A" oninput="updatePayeeField('${p.id}','name',this.value)"></td>
<td><input type="number" step="any" min="0" max="100" value="${p.sharePct||''}" placeholder="0" oninput="updatePayeeField('${p.id}','sharePct',this.value)" style="text-align:right;font-family:var(--font-num);"></td>
<td><input type="text" value="${escapeHtml(p.pan||'')}" placeholder="PAN (required)" oninput="updatePayeeField('${p.id}','pan',this.value)"></td>
<td><input type="text" value="${escapeHtml(p.controlNo||'')}" placeholder="Control No." oninput="updatePayeeField('${p.id}','controlNo',this.value)"></td>
<td class="num" id="payeeNet_${p.id}">${fmt(net)}</td>
<td><button class="icon-btn" onclick="removePayeeRow('${p.id}')"><i class="fa-solid fa-trash"></i></button></td>
</tr>`;
}).join('');
renderPayeeTotals();
}
function renderPayeeTotals() {
const e = entries.find(x => x.id === editingId);
const c = e ? computeEntryValues(e) : { arMonthly: 0 };
let totalPct = 0, totalNet = 0;
formPayees.forEach(p => {
totalPct += num(p.sharePct);
const gross = round0((c.arMonthly || 0) * (num(p.sharePct) / 100));
const net = round0(gross - round0(gross * 0.10));
totalNet += net;
const cell = document.getElementById('payeeNet_' + p.id);
if (cell) cell.textContent = fmt(net);
});
const pctEl = document.getElementById('payeeTotalPct');
const netEl = document.getElementById('payeeTotalNet');
if (pctEl) { pctEl.textContent = totalPct.toFixed(2).replace(/\.00$/, '') + '%'; pctEl.style.color = (Math.abs(totalPct - 100) < 0.01) ? 'var(--green)' : 'var(--red)'; }
if (netEl) netEl.textContent = fmt(totalNet);
}
function savePayees() {
if (editingId === null) { closePayeeModal(); return; }
const e = entries.find(x => x.id === editingId);
if (!e) { closePayeeModal(); return; }
const cleaned = formPayees.map(p => ({
name: String(p.name || '').trim(),
sharePct: num(p.sharePct),
pan: String(p.pan || '').trim(),
controlNo: String(p.controlNo || '').trim()
})).filter(p => p.name || p.sharePct);
if (cleaned.length > 0) {
const totalPct = cleaned.reduce((s, p) => s + num(p.sharePct), 0);
if (Math.abs(totalPct - 100) > 0.01) { showToast(`Total share must equal 100% (currently ${totalPct}%)`, true); return; }
for (const p of cleaned) { if (!p.name) { showToast('Every payee needs a name', true); return; } }
for (const p of cleaned) { if (!p.pan) { showToast(`PAN No. is required for every payee (missing for "${p.name}")`, true); return; } }
}
const updated = { ...e, payees: cleaned };
updateEntryInFirebase(updated).then(() => {
closePayeeModal();
fillForm(updated);
if (cleaned.length > 0) { showToast(`✓ ${cleaned.length} payee(s) saved. Click Auto Generate to rebuild schedule.`, false); }
else { showToast('Payee split cleared — schedule is single payee again.', false); }
}).catch(() => {});
}

/* =========================================================
EXPORT ALL ENTRIES
========================================================= */
function exportExcel() {
if (entries.length === 0) { showToast('No entries to export', true); return; }
const data = entries.map((e, i) => {
const c = computeEntryValues(e);
const allPayments = (e.payments || []).map(p => `${p.date || ''} | ₹${num(p.amount).toFixed(2)} | ${p.mode || ''}`).join('\n');
const allModes = (e.payments || []).map(p => p.mode).filter(Boolean).join(', ');
const allPayDates = (e.payments || []).map(p => p.date).filter(Boolean).join(', ');
const allPayAmounts = (e.payments || []).map(p => num(p.amount)).join(', ');
const payeeSummary = c.hasPayeeSplit ? c.payeeSplits.map(p => `${p.name}: ${p.sharePct}% (Net ₹${p.net})`).join(' | ') : '';
return {
'S.no': i + 1, 'Company': e.company || 'BEPL', 'Unit No.': e.unitNo || '', 'Duplicate Reason': e.duplicateReason || '',
'Imported': e.isImported ? 'Yes' : 'No', 'Source Sheet': e.sourceSheet || '',
'Name of the client': e.client, 'Description': e.description, 'CP': e.cp, 'RM': e.rm, 'Floor': e.floor,
'Area': e.area, 'Rate': e.rate, 'Gold for CP': e.goldCp, 'Gold for client': e.goldClient,
'LPS Dis.': e.lps, 'CP Dis.': e.cpDisc, 'I Phone Adjustment': e.iphone, 'TPD (within 30 days)': e.tpd,
'On Spot Discount': e.spot, 'Other Discount Type': e.otherDiscType || '', 'Other Discount Amount': e.otherDisc || 0, 'Total Discount': c.totalDisc, 'Discount Date': e.discountDate || '', 'Due %': e.duePct,
'Manual Due Override': e.manualDue || '', 'Manual AR Monthly': e.manualArMonthly || '', 'Status': e.status,
'Payee Split': payeeSummary,
'Total Payment Count': (e.payments || []).length, 'All Payment Dates': allPayDates, 'All Payment Amounts': allPayAmounts,
'All Payment Modes': allModes, 'All Payments (Detail)': allPayments,
'Bank': e.bank, 'Ref No.': e.ref, 'PAN No.': e.pan, 'Favouring': e.favouring,
'Last Issued Cheque No.': e.chequeNo, 'Control No.': e.controlNo,
'Gross Amount': c.gross, 'Due Amount': c.due, 'Total Received (Gross)': c.grossReceived,
'Amount Received (Net)': c.received, 'Balance': c.balance, 'AR Eligible': c.eligible ? 'Yes' : 'No',
'AR Period (Months)': c.period, 'AR Monthly (1% of Received + Discount)': c.arMonthly, 'TDS (10%)': c.tds,
'Net Monthly': c.net, 'Total Payable': c.totalPayable,
};
});
const ws = XLSX.utils.json_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'AR Entries');
XLSX.writeFile(wb, `AR_Calculation_Tool_${new Date().toISOString().slice(0,10)}.xlsx`);
showToast('All entries exported', false);
}

/* ============ EXCEL SCHEDULE EXPORT ============ */
function addScheduleSheetToWorkbook_orig(wb, entry, usedNames) {
const allCheques = getAllChequesFlat(entry.id);
if (allCheques.length === 0) return false;
const c = computeEntryValues(entry);
const NAVY = 'FF0A1E3D', LINE = 'FFB9C2D6', LABEL_BG = 'FFEFF2F7';
const BAND_TEXT = 'FFFFFFFF', GREEN = 'FF1F8A55', GREEN_BG = 'FFE4F5EC';
const RED = 'FFC1432F', RED_BG = 'FFFBE9E6', AMBER = 'FFA8720F', AMBER_BG = 'FFFAF1DE';
const TOTAL_BG = 'FFE7ECF3', INR_FMT = '###,##0', INR_FMT2 = '###,##0.00';
const thin = { style: 'thin', color: { argb: LINE } };
const allBorders = { top: thin, left: thin, bottom: thin, right: thin };
let baseName = (entry.description || entry.client || 'Sheet').replace(/[\\\/\?\*\[\]:]/g, ' ').trim().slice(0, 26) || 'Sheet';
let sheetName = baseName, suffix = 1;
while (usedNames.has(sheetName.toLowerCase())) { suffix++; sheetName = (baseName.slice(0, 22) + '_' + suffix).trim(); }
usedNames.add(sheetName.toLowerCase());
const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
ws.columns = [27, 23, 16, 40, 14, 11, 16, 19, 12, 11, 55].map(w => ({ width: w }));
const setCell = (row, col, value, opts = {}) => {
const cell = ws.getCell(row, col);
cell.value = value;
cell.font = { name: 'Segoe UI', size: 10, bold: !!opts.bold, italic: !!opts.italic, color: { argb: opts.color || 'FF18233A' } };
cell.alignment = { vertical: 'middle', horizontal: opts.align || 'left', wrapText: !!opts.wrap };
if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
if (opts.border !== false) cell.border = allBorders;
if (opts.numFmt) cell.numFmt = opts.numFmt;
return cell;
};
const band = (row, fromCol, toCol, text, opts = {}) => {
ws.mergeCells(row, fromCol, row, toCol);
setCell(row, fromCol, text, { bold: true, align: opts.align || 'center', fill: opts.fill || NAVY, color: opts.color || BAND_TEXT });
ws.getRow(row).height = opts.height || 20;
};
const lv = (row, labelCol, valueCol, label, value, opts = {}) => {
setCell(row, labelCol, label, { bold: opts.boldLabel !== false, fill: LABEL_BG });
const vc = setCell(row, valueCol, value, { align: opts.align || 'left', numFmt: opts.numFmt, bold: opts.boldValue, color: opts.color });
if (opts.fill) vc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
return vc;
};
let r = 1;
band(r, 1, 10, `${entry.company || 'BEPL'} — UNIT ${entry.unitNo || '—'} — ASSURED RETURN (AR) CALCULATION SHEET`, { height: 24 });
ws.getRow(r).font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: BAND_TEXT } };
r += 2;
band(r, 1, 4, 'CUSTOMER & PARTY DETAILS');
band(r, 5, 9, 'PROPERTY & AR SUMMARY');
r++;
const pairs = [
[['Client Name', entry.client || ''], ['PAN No.', entry.pan || ''], ['Floor / Type', entry.floor || '-'], ['Due Amount', c.due]],
[['Unit No.', entry.description || '-'], ['Favouring Name', entry.favouring || ''], ['Area (sq.ft)', entry.area || 0], ['Total Received (Net)', c.received]],
[['Channel Partner (CP)', entry.cp || '-'], ['Control No.', entry.controlNo || ''], ['Rate (₹/sq.ft)', entry.rate || 0], ['Balance', c.balance]],
[['Relationship Manager', entry.rm || '-'], ['Company ', `${entry.company||'-'} / ${entry.unitNo||'-'}`], ['Gross Amount', c.gross], ['Due %', entry.duePct || 100]],
];
for (let i = 0; i < 4; i++) {
lv(r + i, 1, 2, pairs[i][0][0], pairs[i][0][1]);
lv(r + i, 3, 4, pairs[i][1][0], pairs[i][1][1]);
const p2 = pairs[i][2]; const isMoney = ['Rate (₹/sq.ft)','Gross Amount'].includes(p2[0]);
lv(r + i, 5, 6, p2[0], p2[1], { align: isMoney?'right':'left', numFmt: isMoney?INR_FMT2:undefined });
const p3 = pairs[i][3]; const isPct = p3[0] === 'Due %'; const isBal = p3[0] === 'Balance';
lv(r + i, 8, 9, p3[0], p3[1], { align: 'right', numFmt: isPct ? '0.00"%"' : INR_FMT2, boldValue: true, color: isBal ? (c.balance > 0 ? RED : GREEN) : undefined });
}
r += 5;
band(r, 1, 9, 'DISCOUNT DETAILS', { height: 20 });
r++;
lv(r, 1, 2, 'LPS Discount', num(entry.lps), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
lv(r, 4, 5, 'TPD (within 30 days)', num(entry.tpd), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
setCell(r, 7, `Discount Date: ${entry.discountDate ? fmtDateSlash(new Date(entry.discountDate)) : '—'}`, {});
r++;
lv(r, 1, 2, 'CP Discount', num(entry.cpDisc), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
lv(r, 4, 5, 'On Spot Discount', num(entry.spot), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
setCell(r, 7, 'Total Discount', { bold: true, fill: TOTAL_BG });
setCell(r, 9, c.totalDisc, { bold: true, align: 'right', numFmt: INR_FMT2, fill: TOTAL_BG });
r++;
lv(r, 1, 2, 'I Phone Adjustment', num(entry.iphone), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
lv(r, 4, 5, entry.otherDiscType ? `Other Discount (${entry.otherDiscType})` : 'Other Discount', num(entry.otherDisc), { align: 'right', numFmt: INR_FMT2, boldLabel: false });
r += 3;

if (c.hasPayeeSplit && c.payeeSplits.length > 0) {
band(r, 1, 10, 'PAYEE SPLIT — MONTHLY AR DIVISION', { height: 20, fill: 'FF7C3AED' });
r++;
['Payee', 'Share %', 'Monthly Gross', 'TDS @10%', 'Monthly Net', 'Favouring', 'PAN', 'Control No.'].forEach((h, i) => setCell(r, 1 + i, h, { bold: true, align: 'center', fill: NAVY, color: BAND_TEXT }));
r++;
let spG = 0, spT = 0, spN = 0;
c.payeeSplits.forEach(p => {
setCell(r, 1, p.name || '', {});
setCell(r, 2, p.sharePct, { align: 'right', numFmt: '0.00"%"' });
setCell(r, 3, p.gross, { align: 'right', numFmt: INR_FMT });
setCell(r, 4, p.tds, { align: 'right', numFmt: INR_FMT });
setCell(r, 5, p.net, { align: 'right', numFmt: INR_FMT });
setCell(r, 6, p.favouring || '', {});
setCell(r, 7, p.pan || '', {});
setCell(r, 8, p.controlNo || '', {});
spG += p.gross; spT += p.tds; spN += p.net;
r++;
});
ws.mergeCells(r, 1, r, 2);
setCell(r, 1, 'Total', { bold: true, align: 'right', fill: TOTAL_BG });
setCell(r, 3, spG, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
setCell(r, 4, spT, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
setCell(r, 5, spN, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
r += 2;
}

band(r, 1, 3, 'AR CALCULATION TERMS');
const ph = r;
['Date', 'Amount', 'GST %', 'GST Amt', 'Net Amt'].forEach((h, i) => setCell(r, 4 + i, h, { bold: true, align: 'center', fill: NAVY, color: BAND_TEXT }));
r++;
const termRows = [
['AR Eligibility', c.eligible ? (entry.isImported?'Eligible (Imported)':'Eligible') : 'Not Eligible', false, c.eligible?GREEN_BG:RED_BG, c.eligible?GREEN:RED],
['AR Period / Cheques', c.period, false, null, null],
['AR Monthly (1% of Received + Discount)', c.arMonthly, true, null, null],
['TDS @ 10% (Monthly)', c.tds, true, null, null],
['Net Monthly Payable', c.net, true, null, null],
['Total Payable', c.totalPayable, true, AMBER_BG, AMBER],
];
const tsr = r;
termRows.forEach((t, i) => {
lv(tsr + i, 1, 2, t[0], t[1], { align: 'right', numFmt: t[2]?INR_FMT:undefined, boldValue: true, fill: t[3]||undefined, color: t[4]||undefined });
});
const payments = entry.payments || [];
let grossReceived = 0, gstTotal = 0;
payments.forEach((p, i) => {
const { gstAmt, netAmt } = gstSplit(p.amount, p.gstRate);
grossReceived += num(p.amount);
gstTotal += gstAmt;
const pr = ph + 1 + i;
setCell(pr, 4, p.date ? fmtDateSlash(new Date(p.date)) : '', { align: 'center' });
setCell(pr, 5, num(p.amount), { align: 'right', numFmt: INR_FMT2 });
setCell(pr, 6, num(p.gstRate), { align: 'center', numFmt: '0"%"' });
setCell(pr, 7, gstAmt, { align: 'right', numFmt: INR_FMT2 });
setCell(pr, 8, netAmt, { align: 'right', numFmt: INR_FMT2 });
});
const totalRowIdx = ph + 1 + payments.length;
setCell(totalRowIdx, 4, 'Total Received', { bold: true, fill: TOTAL_BG });
setCell(totalRowIdx, 5, grossReceived, { bold: true, align: 'right', numFmt: INR_FMT2, fill: TOTAL_BG });
setCell(totalRowIdx, 6, '', { fill: TOTAL_BG });
setCell(totalRowIdx, 7, gstTotal, { bold: true, align: 'right', numFmt: INR_FMT2, fill: TOTAL_BG });
setCell(totalRowIdx, 8, grossReceived - gstTotal, { bold: true, align: 'right', numFmt: INR_FMT2, fill: TOTAL_BG });
r = Math.max(tsr + termRows.length - 1, totalRowIdx) + 2;
band(r, 1, 10, 'AR CHEQUE ISSUANCE SCHEDULE', { height: 22 });
r++;
['S.No', 'Cheque No.', 'Cheque Date', 'Favouring', 'Gross Amount', 'TDS @10%', 'Net Amount', 'Period From', 'Period To', 'Control No.', 'Narration'].forEach((h, i) => setCell(r, 1 + i, h, { bold: true, align: 'center', fill: NAVY, color: BAND_TEXT }));
r++;
let sumGross = 0, sumTds = 0, sumNet = 0;
allCheques.forEach(row => {
setCell(r, 1, row.sno, { align: 'center' });
setCell(r, 2, row.chequeNo || '', { align: 'center' });
setCell(r, 3, fmtDateSlash(row.chequeDate), { align: 'center' });
setCell(r, 4, row.favouring || entry.favouring || '', { align: 'center' });
setCell(r, 5, row.gross, { align: 'right', numFmt: INR_FMT });
setCell(r, 6, row.tds, { align: 'right', numFmt: INR_FMT });
setCell(r, 7, row.net, { align: 'right', numFmt: INR_FMT });
setCell(r, 8, fmtDateSlash(row.from), { align: 'center' });
setCell(r, 9, fmtDateSlash(row.to), { align: 'center' });
setCell(r, 10, row.controlNo || '', { align: 'center' });
const narration = `CH NO. ${row.chequeNo || '---'} DTD ${fmtDateSlash(row.chequeDate)} ISSUED TOWARDS AR FOR THE PERIOD FROM ${fmtDateSlash(row.from)} TO ${fmtDateSlash(row.to)} VIDE CONTROL NO ${row.controlNo || ''}`;
setCell(r, 11, narration, { align: 'left', wrap: true });
sumGross += num(row.gross); sumTds += num(row.tds); sumNet += num(row.net);
r++;
});
ws.mergeCells(r, 1, r, 4);
setCell(r, 1, 'Grand Total', { bold: true, align: 'center', fill: TOTAL_BG });
setCell(r, 5, sumGross, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
setCell(r, 6, sumTds, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
setCell(r, 7, sumNet, { bold: true, align: 'right', numFmt: INR_FMT, fill: TOTAL_BG });
setCell(r, 11, '', { fill: TOTAL_BG });
return true;
}

async function exportChequeSchedule_orig() {
if (editingId === null) { showToast('Please save the entry first', true); return; }
const entry = entries.find(e => e.id === editingId);
if (!entry) { showToast('Entry not found', true); return; }
if (getTotalCheques(editingId) === 0) { showToast('No cheques issued for this entry', true); return; }
if (typeof ExcelJS === 'undefined') { showToast('Excel styling library failed to load', true); return; }
const wb = new ExcelJS.Workbook();
wb.creator = 'AR Calculation Tool';
wb.created = new Date();
const usedNames = new Set();
addScheduleSheetToWorkbook(wb, entry, usedNames);
wb.xlsx.writeBuffer().then(buf => {
const blob = new Blob([buf], { type: 'application/octet-stream' });
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = `AR_Schedule_${(entry.client||'client').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`;
document.body.appendChild(a); a.click(); document.body.removeChild(a);
URL.revokeObjectURL(url);
showToast('AR Calculation Sheet exported', false);
}).catch(err => { console.error(err); showToast('Export failed: ' + err.message, true); });
}

async function exportAllSchedules_orig() {
if (typeof ExcelJS === 'undefined') { showToast('Excel styling library failed to load', true); return; }
const withSchedules = entries.filter(e => getTotalCheques(e.id) > 0);
if (withSchedules.length === 0) { showToast('No entries have a cheque schedule to export', true); return; }
const wb = new ExcelJS.Workbook();
wb.creator = 'AR Calculation Tool';
wb.created = new Date();
const usedNames = new Set();
let count = 0;
withSchedules.forEach(entry => { if (addScheduleSheetToWorkbook(wb, entry, usedNames)) count++; });
if (count === 0) { showToast('No schedules to export', true); return; }
wb.xlsx.writeBuffer().then(buf => {
const blob = new Blob([buf], { type: 'application/octet-stream' });
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = `All_AR_Schedules_${new Date().toISOString().slice(0,10)}.xlsx`;
document.body.appendChild(a); a.click(); document.body.removeChild(a);
URL.revokeObjectURL(url);
showToast(`✓ ${count} schedule sheet(s) exported`, false);
}).catch(err => { console.error(err); showToast('Export failed: ' + err.message, true); });
}

/* ============ IMPORT EXCEL ============ */
document.getElementById('importFile').addEventListener('change', function(e) {
const file = e.target.files[0];
if (!file) return;
pendingImportFile = file;
e.target.value = '';
document.getElementById('importModeModal').classList.add('show');
});

function closeImportModeModal() { document.getElementById('importModeModal').classList.remove('show'); pendingImportFile = null; }

function executeImport(mode) {
if (!pendingImportFile) { showToast('No file selected', true); closeImportModeModal(); return; }
document.getElementById('importModeModal').classList.remove('show');
const file = pendingImportFile;
pendingImportFile = null;
const reader = new FileReader();
reader.onload = async function(evt) {
try {
const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true });
if (mode === 'all') await importAllRows(wb);
else await importMergedByClient(wb);
} catch (err) {
console.error(err);
showToast('Import failed: ' + err.message, true);
}
};
reader.readAsArrayBuffer(file);
}

function parseImportedDate(v) {
if (!v && v !== 0) return '';
if (v instanceof Date && !isNaN(v)) { const y = v.getFullYear(); const m = String(v.getMonth() + 1).padStart(2, '0'); const d = String(v.getDate()).padStart(2, '0'); return `${y}-${m}-${d}`; }
if (typeof v === 'number' && v > 1000) { const d = XLSX.SSF ? XLSX.SSF.parse_date_code(v) : null; if (d && d.y) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`; }
const parsed = new Date(v);
if (!isNaN(parsed)) { const y = parsed.getFullYear(); const m = String(parsed.getMonth() + 1).padStart(2, '0'); const d = String(parsed.getDate()).padStart(2, '0'); return `${y}-${m}-${d}`; }
return '';
}

function isDateLike(v) {
if (!v && v !== 0) return false;
if (v instanceof Date && !isNaN(v)) return true;
if (typeof v === 'number' && v > 20000 && v < 60000) return true;
const s = String(v);
return /^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(s);
}

function findHeaderRow(rows) {
for (let i = 0; i < Math.min(rows.length, 25); i++) {
const row = rows[i];
if (!row) continue;
const strs = row.map(c => String(c || '').toLowerCase().trim());
const hasName = strs.some(c => c.includes('name of') || c === 'particulars' || c.includes('client name'));
const hasAmount = strs.some(c => c.includes('amount'));
const hasDate = strs.some(c => c === 'date' || c.includes(' date'));
const hasPayment = strs.some(c => /(payment|pay|p)\s*\d+/.test(c));
if (hasName && (hasAmount || hasDate || hasPayment)) return i;
}
return -1;
}

function buildColMap(headers) {
const findCol = (...keywords) => {
for (let i = 0; i < headers.length; i++) {
const h = headers[i] || '';
for (const kw of keywords) { if (h.includes(kw)) return i; }
}
return -1;
};
return {
name: findCol('name of the client', 'name of client', 'client name', 'particulars'),
desc: findCol('description'), cp: findCol('cp'), rm: findCol('rm'), floor: findCol('floor'),
area: findCol('area'), rate: findCol('rate'), goldCp: findCol('gold for cp'), goldClient: findCol('gold for client'),
gross: findCol('gross amount', 'gross'), duePct: findCol('due %', 'due%'), dueAmt: findCol('due amount'),
status: findCol('status'), mode: findCol('mode'), date: findCol('date'), amount: findCol('amount'),
pan: findCol('pan no', 'pan'), favouring: findCol('favouring'), controlNo: findCol('control no'),
chequeNo: findCol('cheque no'), company: findCol('company'), unitNo: findCol('unit no', 'unit number', 'unit'),
lps: findCol('lps dis', 'lps'), cpDisc: findCol('cp dis'), iphone: findCol('i phone', 'iphone'),
tpd: findCol('tpd'), spot: findCol('on spot', 'spot'), totalDisc: findCol('total discount'),
otherDiscType: findCol('other discount type'), otherDisc: findCol('other discount amount', 'other discount'),
period: findCol('no. of cheque', 'no of cheque', 'no. of cheques', 'no of cheques', 'total cheque', 'total cheques issued', 'cheques issued', 'cheque to be issued', 'ar period', 'period'),
sourceSheet: findCol('source sheet'),
};
}

function extractPaymentsFromWideColumns(headers, row) {
const found = {};
for (let i = 0; i < headers.length; i++) {
const h = (headers[i] || '').toLowerCase().trim();
if (!h) continue;
const m = h.match(/^(payment|pay|p)\s*(\d+)\s*[\.\-_\s]*(date|dt|amount|amt|mode|type|method|value|received)?/);
if (!m) continue;
const payNum = m[2];
const subField = (m[3] || '').trim();
if (!found[payNum]) found[payNum] = { date: '', amount: 0, mode: '' };
if (subField === 'date' || subField === 'dt') found[payNum].date = row[i];
else if (subField === 'amount' || subField === 'amt' || subField === 'value' || subField === 'received') found[payNum].amount = num(row[i]);
else if (subField === 'mode' || subField === 'type' || subField === 'method') found[payNum].mode = String(row[i] || '').trim();
else {
const val = row[i];
if (isDateLike(val)) { if (!found[payNum].date) found[payNum].date = val; }
else { const n = num(val); if (n !== 0) found[payNum].amount = n; }
}
}
const payments = [];
Object.keys(found).sort((a, b) => parseInt(a) - parseInt(b)).forEach(payNum => {
const p = found[payNum];
if (p.amount > 0) payments.push({ date: parseImportedDate(p.date), amount: p.amount, mode: p.mode, gstRate: 0 });
});
return payments;
}

async function importChequeScheduleFromWorkbook(wb, savedEntries) {
let schedSheet = null;
for (const sheetName of wb.SheetNames) {
const sheet = wb.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });
if (rows.length < 3) continue;
for (let i = 0; i < Math.min(rows.length, 5); i++) {
const row = rows[i];
if (!row) continue;
const strs = row.map(c => String(c || '').toLowerCase().trim());
const hasSourceSheet = strs.some(s => s === 'source sheet');
const hasChequeDate = strs.some(s => s.includes('cheque date'));
const hasSno = strs.some(s => s === 's.no' || s === 'sno' || s === 'sr no');
if (hasSourceSheet && (hasChequeDate || hasSno)) {
schedSheet = { name: sheetName, rows, headerIdx: i, headers: strs };
break;
}
}
if (schedSheet) break;
}
if (!schedSheet) return { updated: 0 };

const H = schedSheet.headers;
const getIdx = (...kws) => {
for (let i = 0; i < H.length; i++) {
for (const kw of kws) { if (H[i] === kw || H[i].includes(kw)) return i; }
}
return -1;
};
const colIdx = {
sourceSheet: getIdx('source sheet'),
client: getIdx('client'),
lotCount: getIdx('lot no'),
sno: getIdx('s.no'),
chequeNo: getIdx('cheque no'),
chequeDate: getIdx('cheque date'),
gross: getIdx('gross amount', 'gross'),
tds: getIdx('tds'),
net: getIdx('net amount'),
from: getIdx('period from'),
to: getIdx('period to'),
control: getIdx('control no'),
payee: getIdx('payee'),
};
if (colIdx.sourceSheet === -1) return { updated: 0 };

const bySource = {};
for (let r = schedSheet.headerIdx + 1; r < schedSheet.rows.length; r++) {
const row = schedSheet.rows[r];
if (!row || row.length === 0) continue;
const src = String(row[colIdx.sourceSheet] || '').trim();
if (!src) continue;
if (!bySource[src]) bySource[src] = [];
bySource[src].push(row);
}

const entriesBySource = {};
savedEntries.forEach(e => {
const srcStr = (e.sourceSheet || '').trim();
if (!srcStr) return;
const srcs = srcStr.split(',').map(s => s.trim()).filter(s => s);
srcs.forEach(src => {
if (!entriesBySource[src]) entriesBySource[src] = [];
entriesBySource[src].push(e);
});
});

let updatedCount = 0;
const updatePromises = [];

for (const [srcSheet, rows] of Object.entries(bySource)) {
const matchingEntries = entriesBySource[srcSheet];
if (!matchingEntries || matchingEntries.length === 0) continue;

const lots = [];
let currentLot = null;
let lastLotCount = null;
let lastSno = 0;

for (const row of rows) {
const lotCount = num(row[colIdx.lotCount]);
const sno = num(row[colIdx.sno]);
const controlNo = String(row[colIdx.control] || '').trim();
const payeeName = colIdx.payee >= 0 ? String(row[colIdx.payee] || '').trim() : '';

let isNewLot = false;
if (currentLot === null) isNewLot = true;
else if (lotCount !== lastLotCount) isNewLot = true;
else if (sno > 0 && sno <= lastSno) isNewLot = true;

if (isNewLot) {
if (currentLot) lots.push(currentLot);
currentLot = { lotNumber: lots.length + 1, issuedDate: new Date().toISOString(), controlNo: controlNo, payeeName: payeeName, cheques: [] };
lastLotCount = lotCount;
lastSno = 0;
}
if (controlNo && !currentLot.controlNo) currentLot.controlNo = controlNo;
if (payeeName && !currentLot.payeeName) currentLot.payeeName = payeeName;

const chequeDateStr = parseImportedDate(row[colIdx.chequeDate]);
const fromStr = parseImportedDate(row[colIdx.from]);
const toStr = parseImportedDate(row[colIdx.to]);

currentLot.cheques.push({
sno: sno || currentLot.cheques.length + 1,
chequeNo: String(row[colIdx.chequeNo] || '').trim().replace(/^'/, ''),
chequeDate: chequeDateStr ? new Date(chequeDateStr).toISOString() : null,
from: fromStr ? new Date(fromStr).toISOString() : null,
to: toStr ? new Date(toStr).toISOString() : null,
gross: num(row[colIdx.gross]),
tds: num(row[colIdx.tds]),
net: num(row[colIdx.net]),
payeeName: payeeName,
});
lastSno = sno;
}
if (currentLot) lots.push(currentLot);

const targetEntry = matchingEntries[0];
const allCheques = lots.flatMap(l => l.cheques);
const lastChq = allCheques.length > 0 ? allCheques[allCheques.length - 1] : null;

const updated = { ...targetEntry, issuedCheques: lots };
if (lastChq && lastChq.chequeNo) updated.chequeNo = lastChq.chequeNo;

updatePromises.push(
updateEntryInFirebase(updated).then(() => { updatedCount++; }).catch(() => {})
);
}

await Promise.all(updatePromises);
return { updated: updatedCount };
}

async function importAllRows(wb) {
const summary = {};
const entriesToSave = [];

for (const sheetName of wb.SheetNames) {
const sheet = wb.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });
if (rows.length === 0) continue;

const headerIndices = [];
let i = 0;
while (i < rows.length) {
const idx = findHeaderRow(rows.slice(i));
if (idx === -1) break;
headerIndices.push(i + idx);
i = i + idx + 1;
}
if (headerIndices.length === 0) continue;

let sheetCount = 0;

for (let blockIdx = 0; blockIdx < headerIndices.length; blockIdx++) {
const headerIdx = headerIndices[blockIdx];
const headers = rows[headerIdx].map(c => String(c || '').toLowerCase().trim());
const colMap = buildColMap(headers);
const nextHeaderIdx = headerIndices[blockIdx + 1] || rows.length;
let lastClient = '';

for (let r = headerIdx + 1; r < nextHeaderIdx; r++) {
const row = rows[r];
if (!row || row.length === 0) continue;

const nameCell = colMap.name >= 0 ? String(row[colMap.name] || '').trim() : '';
const isNewName = nameCell && nameCell.length > 2 && !nameCell.startsWith('=') && !/^\d+$/.test(nameCell);
if (isNewName) lastClient = nameCell;
if (!lastClient) continue;
if (/^\*+/.test(nameCell)) continue;

let payments = extractPaymentsFromWideColumns(headers, row);
if (payments.length === 0) {
const dateCell = colMap.date >= 0 ? row[colMap.date] : '';
const amountCell = colMap.amount >= 0 ? row[colMap.amount] : '';
const amountVal = amountCell !== '' ? num(amountCell) : 0;
if (isDateLike(dateCell) && amountVal > 0) {
payments = [{ date: parseImportedDate(dateCell), amount: amountVal, mode: colMap.mode >= 0 ? String(row[colMap.mode] || '').trim() : '', gstRate: 0 }];
}
}

const lps = colMap.lps >= 0 ? num(row[colMap.lps]) : 0;
const cpDisc = colMap.cpDisc >= 0 ? num(row[colMap.cpDisc]) : 0;
const iphone = colMap.iphone >= 0 ? num(row[colMap.iphone]) : 0;
const tpd = colMap.tpd >= 0 ? num(row[colMap.tpd]) : 0;
const spot = colMap.spot >= 0 ? num(row[colMap.spot]) : 0;
const otherDisc = colMap.otherDisc >= 0 ? num(row[colMap.otherDisc]) : 0;
const otherDiscType = colMap.otherDiscType >= 0 ? String(row[colMap.otherDiscType] || '').trim() : '';
const hasDiscounts = (lps + cpDisc + iphone + tpd + spot + otherDisc) !== 0;
const area = colMap.area >= 0 ? num(row[colMap.area]) : 0;
const rate = colMap.rate >= 0 ? num(row[colMap.rate]) : 0;
const duePct = colMap.duePct >= 0 ? num(row[colMap.duePct]) : 0;
const dueAmt = colMap.dueAmt >= 0 ? num(row[colMap.dueAmt]) : 0;

if (payments.length === 0 && !hasDiscounts && !area && !rate && !isNewName) continue;

const companyVal = colMap.company >= 0 ? String(row[colMap.company] || '').trim().toUpperCase() : '';
const descVal = colMap.desc >= 0 ? String(row[colMap.desc] || '').trim() : '';
const unitVal = colMap.unitNo >= 0 ? String(row[colMap.unitNo] || '').trim() : '';

const entry = {
id: idSeq++,
client: lastClient,
description: descVal || unitVal,
cp: colMap.cp >= 0 ? String(row[colMap.cp] || '').trim() : '',
rm: colMap.rm >= 0 ? String(row[colMap.rm] || '').trim() : '',
status: colMap.status >= 0 ? String(row[colMap.status] || 'AR').trim() : 'AR',
floor: colMap.floor >= 0 ? String(row[colMap.floor] || '').trim() : '',
area: area, rate: rate,
goldCp: colMap.goldCp >= 0 ? String(row[colMap.goldCp] || '').trim() : '',
goldClient: colMap.goldClient >= 0 ? String(row[colMap.goldClient] || '').trim() : '',
lps: lps, cpDisc: cpDisc, iphone: iphone, tpd: tpd, spot: spot, otherDisc: otherDisc, otherDiscType: otherDiscType,
discountDate: '',
duePct: duePct > 0 ? duePct : 100,
manualDue: dueAmt > 0 ? dueAmt : null,
manualArMonthly: null,
payments: payments,
bank: '', ref: '',
pan: colMap.pan >= 0 ? String(row[colMap.pan] || '').trim() : '',
favouring: colMap.favouring >= 0 ? String(row[colMap.favouring] || '').trim() : '',
chequeNo: colMap.chequeNo >= 0 ? String(row[colMap.chequeNo] || '').trim() : '',
controlNo: colMap.controlNo >= 0 ? String(row[colMap.controlNo] || '').trim() : '',
company: (companyVal === 'BSCIPL') ? 'BSCIPL' : 'BEPL',
unitNo: unitVal || descVal,
arPeriod: colMap.period >= 0 ? (num(row[colMap.period]) || 12) : 12,
sourceSheet: colMap.sourceSheet >= 0 ? String(row[colMap.sourceSheet] || '').trim() : '',
isImported: true,
duplicateReason: null,
payees: [],
_sourceSheet: sheetName,
};
entriesToSave.push(entry);
sheetCount++;
}
}
if (sheetCount > 0) summary[sheetName] = sheetCount;
}

if (entriesToSave.length === 0) { showToast('No entries found. Check headers.', true); return; }

const savedEntries = [];
for (const entry of entriesToSave) {
try { await saveEntryToFirebase(entry); savedEntries.push(entry); }
catch (e) { console.error(e); }
}

const schedResult = await importChequeScheduleFromWorkbook(wb, savedEntries);

const parts = Object.entries(summary).slice(0, 5).map(([s, c]) => `${s}: ${c}`);
const more = Object.keys(summary).length > 5 ? ` +${Object.keys(summary).length - 5} more` : '';
const schedMsg = schedResult.updated > 0 ? ` · ${schedResult.updated} schedule(s) linked` : '';
showToast(`✓ ${savedEntries.length} entries imported${schedMsg} (${parts.join(', ')}${more})`, false);
}

async function importMergedByClient(wb) {
const summary = {};
const entriesToSave = [];

for (const sheetName of wb.SheetNames) {
const sheet = wb.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });
if (rows.length === 0) continue;

const headerIndices = [];
let i = 0;
while (i < rows.length) {
const idx = findHeaderRow(rows.slice(i));
if (idx === -1) break;
headerIndices.push(i + idx);
i = i + idx + 1;
}
if (headerIndices.length === 0) continue;

let sheetCount = 0;

for (let blockIdx = 0; blockIdx < headerIndices.length; blockIdx++) {
const headerIdx = headerIndices[blockIdx];
const headers = rows[headerIdx].map(c => String(c || '').toLowerCase().trim());
const colMap = buildColMap(headers);
const nextHeaderIdx = headerIndices[blockIdx + 1] || rows.length;
const groups = {};
let lastGroupKey = null;
let lastClient = '';

for (let r = headerIdx + 1; r < nextHeaderIdx; r++) {
const row = rows[r];
if (!row || row.length === 0) continue;

const nameCell = colMap.name >= 0 ? String(row[colMap.name] || '').trim() : '';
const isNewName = nameCell && nameCell.length > 2 && !nameCell.startsWith('=') && !/^\d+$/.test(nameCell);

if (isNewName) {
lastClient = nameCell;
const descVal = colMap.desc >= 0 ? String(row[colMap.desc] || '').trim() : '';
const unitVal = colMap.unitNo >= 0 ? String(row[colMap.unitNo] || '').trim() : '';
const srcVal = colMap.sourceSheet >= 0 ? String(row[colMap.sourceSheet] || '').trim() : '';
const groupKey = (descVal || unitVal || lastClient).toLowerCase();

if (!groups[groupKey]) {
const companyVal = colMap.company >= 0 ? String(row[colMap.company] || '').trim().toUpperCase() : '';
groups[groupKey] = {
clientName: lastClient,
meta: {
desc: descVal,
cp: colMap.cp >= 0 ? String(row[colMap.cp] || '').trim() : '',
rm: colMap.rm >= 0 ? String(row[colMap.rm] || '').trim() : '',
floor: colMap.floor >= 0 ? String(row[colMap.floor] || '').trim() : '',
area: colMap.area >= 0 ? num(row[colMap.area]) : 0,
rate: colMap.rate >= 0 ? num(row[colMap.rate]) : 0,
goldCp: colMap.goldCp >= 0 ? String(row[colMap.goldCp] || '').trim() : '',
goldClient: colMap.goldClient >= 0 ? String(row[colMap.goldClient] || '').trim() : '',
duePct: colMap.duePct >= 0 ? (num(row[colMap.duePct]) || 100) : 100,
dueAmt: colMap.dueAmt >= 0 ? num(row[colMap.dueAmt]) : 0,
status: colMap.status >= 0 ? String(row[colMap.status] || 'AR').trim() : 'AR',
pan: colMap.pan >= 0 ? String(row[colMap.pan] || '').trim() : '',
favouring: colMap.favouring >= 0 ? String(row[colMap.favouring] || '').trim() : '',
controlNo: colMap.controlNo >= 0 ? String(row[colMap.controlNo] || '').trim() : '',
chequeNo: colMap.chequeNo >= 0 ? String(row[colMap.chequeNo] || '').trim() : '',
company: (companyVal === 'BSCIPL') ? 'BSCIPL' : 'BEPL',
unitNo: unitVal || descVal,
period: colMap.period >= 0 ? (num(row[colMap.period]) || 12) : 12,
sourceSheets: srcVal ? [srcVal] : [],
otherDiscType: colMap.otherDiscType >= 0 ? String(row[colMap.otherDiscType] || '').trim() : '',
},
payments: [],
discounts: { lps: 0, cpDisc: 0, iphone: 0, tpd: 0, spot: 0, otherDisc: 0 }
};
} else {
if (srcVal && !groups[groupKey].meta.sourceSheets.includes(srcVal)) {
groups[groupKey].meta.sourceSheets.push(srcVal);
}
if (!groups[groupKey].meta.otherDiscType && colMap.otherDiscType >= 0) {
const t = String(row[colMap.otherDiscType] || '').trim();
if (t) groups[groupKey].meta.otherDiscType = t;
}
}
lastGroupKey = groupKey;
const g = groups[groupKey];
g.discounts.lps += colMap.lps >= 0 ? num(row[colMap.lps]) : 0;
g.discounts.cpDisc += colMap.cpDisc >= 0 ? num(row[colMap.cpDisc]) : 0;
g.discounts.iphone += colMap.iphone >= 0 ? num(row[colMap.iphone]) : 0;
g.discounts.tpd += colMap.tpd >= 0 ? num(row[colMap.tpd]) : 0;
g.discounts.spot += colMap.spot >= 0 ? num(row[colMap.spot]) : 0;
g.discounts.otherDisc += colMap.otherDisc >= 0 ? num(row[colMap.otherDisc]) : 0;
}

if (!lastGroupKey || !groups[lastGroupKey]) continue;

let widePayments = extractPaymentsFromWideColumns(headers, row);
if (widePayments.length === 0) {
const dateCell = colMap.date >= 0 ? row[colMap.date] : '';
const amountCell = colMap.amount >= 0 ? row[colMap.amount] : '';
const amountVal = amountCell !== '' ? num(amountCell) : 0;
if (isDateLike(dateCell) && amountVal > 0) {
widePayments = [{ date: parseImportedDate(dateCell), amount: amountVal, mode: colMap.mode >= 0 ? String(row[colMap.mode] || '').trim() : '', gstRate: 0 }];
}
}
widePayments.forEach(p => groups[lastGroupKey].payments.push(p));
}

Object.values(groups).forEach(g => {
const hasAnyData = g.payments.length > 0 || g.discounts.lps || g.discounts.cpDisc || g.discounts.iphone || g.discounts.tpd || g.discounts.spot || g.discounts.otherDisc || g.meta.area || g.meta.rate || g.meta.dueAmt;
if (!hasAnyData) return;
const entry = {
id: idSeq++,
client: g.clientName,
description: g.meta.desc || g.meta.unitNo,
cp: g.meta.cp, rm: g.meta.rm, status: g.meta.status,
floor: g.meta.floor, area: g.meta.area, rate: g.meta.rate,
goldCp: g.meta.goldCp, goldClient: g.meta.goldClient,
lps: g.discounts.lps, cpDisc: g.discounts.cpDisc,
iphone: g.discounts.iphone, tpd: g.discounts.tpd, spot: g.discounts.spot,
otherDisc: g.discounts.otherDisc, otherDiscType: g.meta.otherDiscType || '',
discountDate: '', duePct: g.meta.duePct,
manualDue: g.meta.dueAmt > 0 ? g.meta.dueAmt : null,
manualArMonthly: null,
payments: g.payments, bank: '', ref: '',
pan: g.meta.pan, favouring: g.meta.favouring, chequeNo: g.meta.chequeNo,
controlNo: g.meta.controlNo, company: g.meta.company,
unitNo: g.meta.unitNo,
arPeriod: g.meta.period,
sourceSheet: g.meta.sourceSheets.join(', '),
isImported: true, duplicateReason: null, payees: [], _sourceSheet: sheetName,
};
entriesToSave.push(entry);
sheetCount++;
});
}
if (sheetCount > 0) summary[sheetName] = sheetCount;
}

if (entriesToSave.length === 0) { showToast('No entries found. Check headers.', true); return; }

const savedEntries = [];
for (const entry of entriesToSave) {
try { await saveEntryToFirebase(entry); savedEntries.push(entry); }
catch (e) { console.error(e); }
}

const schedResult = await importChequeScheduleFromWorkbook(wb, savedEntries);

const parts = Object.entries(summary).slice(0, 5).map(([s, c]) => `${s}: ${c}`);
const more = Object.keys(summary).length > 5 ? ` +${Object.keys(summary).length - 5} more` : '';
const schedMsg = schedResult.updated > 0 ? ` · ${schedResult.updated} schedule(s) linked` : '';
showToast(`✓ ${savedEntries.length} entries imported (merged by DESCRIPTION)${schedMsg} (${parts.join(', ')}${more})`, false);
}

/* =========================================================
PRINT / PDF EXPORT
========================================================= */
let _printTimer = null;

function printCleanup() {
document.body.classList.remove('printing');
const area = document.getElementById('printArea');
if (area) area.innerHTML = '';
}
window.addEventListener('afterprint', () => { clearTimeout(_printTimer); printCleanup(); });

function doPrint(html) {
const area = document.getElementById('printArea');
if (!area) { showToast('Print area missing', true); return; }
area.innerHTML = html;
document.body.classList.add('printing');
clearTimeout(_printTimer);
setTimeout(() => {
try { window.print(); } catch (e) { console.error(e); showToast('Print failed', true); }
_printTimer = setTimeout(printCleanup, 1500);
}, 80);
}

function prVal(v, fallback) { return (v === undefined || v === null || v === '') ? (fallback || '—') : String(v); }

function buildPrintBrand() {
const now = new Date();
const timeStr = now.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' });
return `<div class="pr-brand">
<div class="pr-mark"><img src="logo.png" class="pr-logo" alt="Logo" onerror="this.remove()"><i class="fa-solid fa-chart-line"></i></div>
<div class="pr-brandtxt">
<h1>AR Calculation Tool</h1>
<div class="pr-ded">Made by DRE</div>
</div>
<div class="pr-meta">Printed On<br><strong>${fmtDate(now)} ${timeStr}</strong></div>
</div>`;
}

function buildEntryPrintPage(entry) {
const c = computeEntryValues(entry);
const lots = (issuedSchedules[entry.id] && issuedSchedules[entry.id].lots) || [];
const allCheques = getAllChequesFlat(entry.id);
const payments = entry.payments || [];

let h = '<div class="pr-page">';
h += buildPrintBrand();
h += `<div class="pr-title">${escapeHtml(entry.company || 'BEPL')} — UNIT ${escapeHtml(entry.unitNo || '—')} — ASSURED RETURN (AR) CALCULATION SHEET</div>`;

h += `<div class="pr-sec">CUSTOMER &amp; PARTY DETAILS</div>`;
h += `<table class="pr-tbl">
<tr>
<th>Client Name</th><td>${escapeHtml(prVal(entry.client))}</td>
<th>Unit No.</th><td>${escapeHtml(prVal(entry.unitNo))}</td>
<th>PAN No.</th><td>${escapeHtml(prVal(entry.pan))}</td>
</tr>
<tr>
<th>Channel Partner</th><td>${escapeHtml(prVal(entry.cp))}</td>
<th>Relationship Manager</th><td>${escapeHtml(prVal(entry.rm))}</td>
<th>Favouring Name</th><td>${escapeHtml(prVal(entry.favouring))}</td>
</tr>
<tr>
<th>Company</th><td>${escapeHtml(prVal(entry.company, 'BEPL'))}</td>
<th>Control No.</th><td>${escapeHtml(prVal(entry.controlNo))}</td>
<th>Status</th><td>${escapeHtml(prVal(entry.status, 'AR'))}</td>
</tr>
<tr>
<th>Floor</th><td>${escapeHtml(prVal(entry.floor))}</td>
<th>Last Issued Cheque No.</th><td>${escapeHtml(prVal(entry.chequeNo))}</td>
<th>Entry Type</th><td>${entry.isImported ? 'Imported' : 'Manual'}${entry.duplicateReason ? ' · DUP: ' + escapeHtml(entry.duplicateReason) : ''}</td>
</tr>
</table>`;

h += `<div class="pr-sec">PROPERTY &amp; AR SUMMARY</div>`;
h += `<table class="pr-tbl">
<tr>
<th>Area (sq.ft)</th><td>${num(entry.area).toLocaleString('en-IN')}</td>
<th>Rate (₹/sq.ft)</th><td>${num(entry.rate).toLocaleString('en-IN')}</td>
<th>Gross Amount</th><td>${fmt(c.gross)}</td>
</tr>
<tr>
<th>Due %</th><td>${num(entry.duePct)}%</td>
<th>Due Amount</th><td>${fmt(c.due)}</td>
<th>Received (Net)</th><td>${fmt(c.received)}</td>
</tr>
<tr>
<th>Balance</th><td>${fmt(c.balance)}</td>
<th>Total Discount</th><td>${fmt(c.totalDisc)}</td>
<th>AR Eligibility</th><td>${c.eligible ? 'Eligible' : 'Not Eligible'}</td>
</tr>
</table>`;

h += `<div class="pr-sec">DISCOUNT DETAILS</div>`;
h += `<table class="pr-tbl">
<tr>
<th>CP Dis.</th><td>${fmt(num(entry.cpDisc))}</td>
<th>I Phone Adjustment</th><td>${fmt(num(entry.iphone))}</td>
<th>TPD (within 30 days)</th><td>${fmt(num(entry.tpd))}</td>
</tr>
<tr>
<th>On Spot Discount</th><td>${fmt(num(entry.spot))}</td>
<th>Other Discount${entry.otherDiscType ? ' (' + escapeHtml(entry.otherDiscType) + ')' : ''}</th><td>${fmt(num(entry.otherDisc))}</td>
<th>Discount Date</th><td>${entry.discountDate ? fmtDateSlash(new Date(entry.discountDate)) : '—'}</td>
</tr>
<tr>
<th>Total Discount</th><td colspan="5"><strong>${fmt(c.totalDisc)}</strong></td>
</tr>
</table>`;

if (payments.length > 0) {
const rows = payments.map((p, i) => {
const { gstAmt, netAmt } = gstSplit(p.amount, p.gstRate);
return `<tr>
<td class="c">${i + 1}</td>
<td class="c">${p.date ? fmtDateSlash(new Date(p.date)) : '—'}</td>
<td class="r">${fmt(p.amount)}</td>
<td class="c">${num(p.gstRate)}%</td>
<td class="r">${fmt(gstAmt)}</td>
<td class="r">${fmt(netAmt)}</td>
<td>${escapeHtml(prVal(p.mode))}</td>
</tr>`;
}).join('');
h += `<div class="pr-sec">PAYMENT &amp; RECEIPT DETAILS</div>
<table class="pr-data">
<thead><tr><th>S.No</th><th>Date</th><th>Amount</th><th>GST %</th><th>GST Amt</th><th>Net Amt</th><th>Mode</th></tr></thead>
<tbody>${rows}</tbody>
<tfoot><tr><td colspan="2" class="r">Total Received</td><td class="r">${fmt(c.grossReceived)}</td><td></td><td class="r">${fmt(c.gstTotal)}</td><td class="r">${fmt(c.received)}</td><td></td></tr></tfoot>
</table>`;
}

h += `<div class="pr-sec">AR CALCULATION TERMS</div>`;
h += `<table class="pr-tbl">
<tr>
<th>AR Period / Cheques</th><td>${c.period}</td>
<th>AR Monthly (1% of Received + Discount)</th><td>${fmt(c.arMonthly)}${c.hasManualArMonthly ? ' (Manual)' : ''}</td>
<th>TDS @ 10% (Monthly)</th><td>${fmt(c.tds)}</td>
</tr>
<tr>
<th>Net Monthly Payable</th><td>${fmt(c.net)}</td>
<th>Total Payable</th><td colspan="3"><strong>${fmt(c.totalPayable)}</strong></td>
</tr>
</table>`;

if (c.hasPayeeSplit && c.payeeSplits.length > 0) {
const rows = c.payeeSplits.map(p => `<tr>
<td>${escapeHtml(p.name)}</td>
<td class="r">${p.sharePct}%</td>
<td class="r">${fmt(p.gross)}</td>
<td class="r">${fmt(p.tds)}</td>
<td class="r">${fmt(p.net)}</td>
<td>${escapeHtml(prVal(p.favouring))}</td>
<td>${escapeHtml(prVal(p.pan))}</td>
<td>${escapeHtml(prVal(p.controlNo))}</td>
</tr>`).join('');
h += `<div class="pr-sec">PAYEE SPLIT — MONTHLY AR DIVISION</div>
<table class="pr-data">
<thead><tr><th>Payee</th><th>Share %</th><th>Monthly Gross</th><th>TDS @10%</th><th>Monthly Net</th><th>Favouring</th><th>PAN</th><th>Control No.</th></tr></thead>
<tbody>${rows}</tbody>
</table>`;
}

h += `<div class="pr-sec">AR CHEQUE ISSUANCE SCHEDULE</div>`;
if (lots.length === 0 || allCheques.length === 0) {
h += `<div class="pr-empty">No cheques issued yet for this entry.</div>`;
} else {
lots.forEach(lot => {
const cheques = lot.cheques || [];
if (cheques.length === 0) return;
let sumG = 0, sumT = 0, sumN = 0;
const rows = cheques.map(r => {
sumG += num(r.gross); sumT += num(r.tds); sumN += num(r.net);
const narration = `CH NO. ${r.chequeNo || '---'} DTD ${fmtDateSlash(r.chequeDate)} ISSUED TOWARDS AR FOR THE PERIOD FROM ${fmtDateSlash(r.from)} TO ${fmtDateSlash(r.to)} VIDE CONTROL NO ${lot.controlNo || ''}`;
return `<tr>
<td class="c">${r.sno}</td>
<td class="c">${escapeHtml(r.chequeNo || '—')}</td>
<td class="c">${fmtDateSlash(r.chequeDate) || '—'}</td>
<td>${escapeHtml(prVal(r.favouring || entry.favouring))}</td>
<td class="r">${fmt(r.gross)}</td>
<td class="r">${fmt(r.tds)}</td>
<td class="r">${fmt(r.net)}</td>
<td class="c">${fmtDateSlash(r.from) || '—'}</td>
<td class="c">${fmtDateSlash(r.to) || '—'}</td>
<td class="c">${escapeHtml(prVal(lot.controlNo))}</td>
<td class="narr">${escapeHtml(narration)}</td>
</tr>`;
}).join('');
h += `<div class="pr-lot">
<div class="pr-lothead">
<span class="pr-pill">Lot ${lot.lotNumber}</span>
${lot.isTopUp ? '<span class="pr-pill" style="background:#a8720f;">TOP-UP</span>' : ''}
${lot.payeeName ? `<span class="pr-pill payee">${escapeHtml(lot.payeeName)}${lot.payeeSharePct ? ' · ' + lot.payeeSharePct + '%' : ''}</span>` : ''}
<span>Control No.: ${escapeHtml(prVal(lot.controlNo))}</span>
<span>${cheques.length} Cheque${cheques.length > 1 ? 's' : ''}</span>
</div>
<table class="pr-data">
<thead><tr>
<th>S.No</th><th>Cheque No.</th><th>Cheque Date</th><th>Favouring</th>
<th>Gross</th><th>TDS @10%</th><th>Net Amt</th>
<th>Period From</th><th>Period To</th><th>Control No.</th><th style="width:26%">Narration</th>
</tr></thead>
<tbody>${rows}</tbody>
<tfoot><tr><td colspan="4" class="r">Lot Total</td><td class="r">${fmt(sumG)}</td><td class="r">${fmt(sumT)}</td><td class="r">${fmt(sumN)}</td><td colspan="4"></td></tr></tfoot>
</table>
</div>`;
});

let gG = 0, gT = 0, gN = 0;
allCheques.forEach(r => { gG += num(r.gross); gT += num(r.tds); gN += num(r.net); });
h += `<table class="pr-data" style="margin-top:6px;">
<tfoot><tr><td colspan="4" class="r">GRAND TOTAL</td><td class="r">${fmt(gG)}</td><td class="r">${fmt(gT)}</td><td class="r">${fmt(gN)}</td><td colspan="4"></td></tr></tfoot>
</table>`;
}

h += `<div class="pr-note">Note: This is a system generated AR calculation sheet. Amounts shown are as per the data entered in the AR Calculation Tool.</div>`;
h += `<div class="pr-sign"><div>Prepared By</div><div>Checked By</div><div>Authorised Signatory</div></div>`;
h += '</div>';
return h;
}

function printEntry(id) {
const e = entries.find(x => x.id === id);
if (!e) { showToast('Entry not found', true); return; }
doPrint(buildEntryPrintPage(e));
}

function printCurrentEntry() {
if (editingId === null) { showToast('Please save the entry first, then open it in edit mode', true); return; }
printEntry(editingId);
}

function printAllEntries() {
if (entries.length === 0) { showToast('No entries to print', true); return; }
if (entries.length > 25 && !confirm(`Print ${entries.length} entries? This may take a while.`)) return;
doPrint(entries.map(buildEntryPrintPage).join(''));
}

/* ============ Top-Up Lot Modal Opener (from banner) ============ */
function openTopUpLotModal(entryId) {
if (editingId !== entryId) { showToast('Please open this entry in edit mode first', true); return; }
const e = entries.find(x => x.id === entryId);
if (!e) return;
const inc = detectArIncrease(e);
if (!inc) { showToast('No AR increase detected', true); return; }
const data = getLotData(entryId);
if (!data.lots || data.lots.length === 0) { showToast('Create the main schedule first', true); return; }

openLotModal();
const select = document.getElementById('lm_lotSelect');
if (select) { select.value = 'new'; onLotSelectChange(); }
setTimeout(() => {
const cb = document.getElementById('lm_incMode');
if (cb) {
cb.checked = true;
onLotIncModeChange();
}
}, 60);
}

/*__AR_PATCH_V2__*/
/* =====================================================================
   PATCH V2 — (1) Excel export choice + formulas  (2) ER Register
   ===================================================================== */
const EXPORT_POPUP_TITLE = 'Oops! Something Went Wrong';
const EXPORT_POPUP_TEXT = 'Please choose how you would like to proceed.';
const EXPORT_LABEL_FORMULA = 'Delete';  // Excel WITH formulas, LPS shown
const EXPORT_LABEL_PLAIN = 'Block';     // Excel WITHOUT formulas, LPS hidden
const EXPORT_LABEL_BACK = 'Back';       // returns to the Entries register
let _expKind = null;
window.__exportOpts = { formulas: true, hideLps: false };

function openExportModal(kind) { _expKind = kind; document.getElementById('exportModal').classList.add('show'); const b = document.getElementById('opsBack'); if (b) b.focus(); }
function exportBack() { closeExportModal(); _expKind = null; switchTab('tableView'); }
document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && document.getElementById('exportModal') && document.getElementById('exportModal').classList.contains('show')) exportBack(); });
function closeExportModal() { document.getElementById('exportModal').classList.remove('show'); }
function runExport(withFormula) {
    if (!_expKind) return;

    // Admin Password Prompt
    const pwd = prompt('Enter Admin Password to proceed with export:');
    if (pwd !== 'admin123') {
        if (pwd !== null) { showToast('Invalid Admin Password! Access Denied.', true); }
        return;
    }

    closeExportModal();
    window.__exportOpts = { formulas: withFormula, hideLps: !withFormula };
    if (_expKind === 'all') exportEntriesWorkbook(window.__exportOpts);
    else if (_expKind === 'sched') exportChequeSchedule_orig();
    else if (_expKind === 'allsched') exportAllSchedules_orig();
}
function exportExcel() { if (entries.length === 0) { showToast('No entries to export', true); return; } openExportModal('all'); }
function exportChequeSchedule() {
    if (editingId === null) { showToast('Please save the entry first', true); return; }
    if (getTotalCheques(editingId) === 0) { showToast('No cheques issued for this entry', true); return; }
    openExportModal('sched');
}
function exportAllSchedules() {
    if (!entries.some(e => getTotalCheques(e.id) > 0)) { showToast('No entries have a cheque schedule to export', true); return; }
    openExportModal('allsched');
}
function addScheduleSheetToWorkbook(wb, entry, usedNames) {
    const ok = addScheduleSheetToWorkbook_orig(wb, entry, usedNames);
    if (!ok) return ok;
    try { schedPost(wb.worksheets[wb.worksheets.length - 1], entry, window.__exportOpts || { formulas: true, hideLps: false }); }
    catch (err) { console.error('schedPost failed', err); }
    return ok;
}

/* Post-process a schedule sheet: hide LPS, or convert values into formulas */
function schedPost(ws, entry, o) {
    const T = (r, c) => { const v = ws.getCell(r, c).value; return (v !== null && typeof v === 'object') ? '' : v; };
    let band = 0;
    for (let r = 1; r <= ws.rowCount; r++) { if (T(r, 1) === 'AR CHEQUE ISSUANCE SCHEDULE') { band = r; break; } }
    if (!band) return;
    const lab = (t, off) => { for (let r = 1; r < band; r++) for (let c = 1; c <= 9; c++) if (T(r, c) === t) return [r, c + (off || 1)]; return null; };
    const labP = p => { for (let r = 1; r < band; r++) for (let c = 1; c <= 9; c++) { const v = T(r, c); if (typeof v === 'string' && v.indexOf(p) === 0) return [r, c + 1]; } return null; };
    const ad = p => ws.getCell(p[0], p[1]).address;
    const setF = (p, f) => { if (!p || !f) return; const cell = ws.getCell(p[0], p[1]); cell.value = { formula: f, result: cell.value }; };

    const lps = lab('LPS Discount');
    if (o.hideLps) {
        if (lps) [[lps[0], lps[1] - 1], [lps[0], lps[1]]].forEach(p => { const cell = ws.getCell(p[0], p[1]); cell.value = null; cell.style = {}; });
        return;
    }
    if (!o.formulas) return;

    // payments table
    let hr = 0;
    for (let r = 1; r < band; r++) { if (T(r, 4) === 'Date' && T(r, 5) === 'Amount') { hr = r; break; } }
    let recvP = lab('Total Received (Net)');
    if (hr) {
        let r = hr + 1;
        while (r < band && T(r, 4) !== 'Total Received') {
            setF([r, 7], `IF(F${r}>0,E${r}/(100+F${r})*F${r},0)`);
            setF([r, 8], `E${r}-G${r}`);
            r++;
        }
        const f = hr + 1, l = r - 1;
        if (l >= f) ['E', 'G', 'H'].forEach((col, i) => setF([r, [5, 7, 8][i]], `SUM(${col}${f}:${col}${l})`));
        setF(recvP, `H${r}`);
    }
    // property / due
    const area = lab('Area (sq.ft)'), rate = lab('Rate (₹/sq.ft)'), gross = lab('Gross Amount'), dpct = lab('Due %'), due = lab('Due Amount'), bal = lab('Balance');
    if (area && rate && gross) setF(gross, `${ad(rate)}*${ad(area)}`);
    if (gross && dpct && due && !(num(entry.manualDue) > 0)) setF(due, `${ad(gross)}*${ad(dpct)}/100`);
    if (due && recvP && bal) setF(bal, `${ad(due)}-${ad(recvP)}`);
    // discount
    const tdisc = lab('Total Discount', 2);
    const parts = [lps, lab('CP Discount'), lab('TPD (within 30 days)'), lab('On Spot Discount'), lab('I Phone Adjustment'), labP('Other Discount')].filter(Boolean).map(ad);
    if (tdisc && parts.length) setF(tdisc, `SUM(${parts.join(',')})`);
    // AR terms
    const per = lab('AR Period / Cheques'), arm = lab('AR Monthly (1% of Received + Discount)'), tds = lab('TDS @ 10% (Monthly)'), net = lab('Net Monthly Payable'), tp = lab('Total Payable');
    if (arm && recvP && tdisc && !(num(entry.manualArMonthly) > 0)) setF(arm, `ROUND((${ad(recvP)}+${ad(tdisc)})*0.01,0)`);
    if (arm && tds) setF(tds, `ROUND(${ad(arm)}*0.1,0)`);
    if (arm && tds && net) setF(net, `${ad(arm)}-${ad(tds)}`);
    if (net && per && tp) setF(tp, `ROUND(${ad(net)}*${ad(per)},0)`);
    // cheque table
    let ch = 0;
    for (let r = band; r <= ws.rowCount; r++) { if (T(r, 1) === 'S.No' && T(r, 2) === 'Cheque No.') { ch = r; break; } }
    if (ch) {
        let r = ch + 1;
        while (r <= ws.rowCount && T(r, 1) !== 'Grand Total') {
            const g = num(T(r, 5)), t = num(T(r, 6)), n = num(T(r, 7));
            if (Math.round(g * 0.1) === t) setF([r, 6], `ROUND(E${r}*0.1,0)`);
            if (g - t === n) setF([r, 7], `E${r}-F${r}`);
            r++;
        }
        const f = ch + 1, l = r - 1;
        if (l >= f && T(r, 1) === 'Grand Total') ['E', 'F', 'G'].forEach((col, i) => setF([r, 5 + i], `SUM(${col}${f}:${col}${l})`));
    }
}

function expDownload(wb, name) {
    return wb.xlsx.writeBuffer().then(buf => {
        const url = URL.createObjectURL(new Blob([buf], { type: 'application/octet-stream' }));
        const a = document.createElement('a'); a.href = url; a.download = name;
        document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
        showToast('Excel exported', false);
    }).catch(err => { console.error(err); showToast('Export failed: ' + err.message, true); });
}

/* Export All Entries — real Excel formulas (or plain values without LPS) */
function exportEntriesWorkbook(o) {
    if (typeof ExcelJS === 'undefined') { showToast('Excel library failed to load', true); return; }
    const C = [], add = (k, h, v, x) => C.push(Object.assign({ k, h, v }, x || {}));
    const pays = e => e.payments || [];
    add('sno', 'S.no', (e, c, i) => i + 1);
    add('co', 'Company', e => e.company || 'BEPL');
    add('unit', 'Unit No.', e => e.unitNo || '');
    add('dup', 'Duplicate Reason', e => e.duplicateReason || '');
    add('imp', 'Imported', e => e.isImported ? 'Yes' : 'No');
    add('src', 'Source Sheet', e => e.sourceSheet || '');
    add('client', 'Name of the client', e => e.client || '');
    add('desc', 'Description', e => e.description || '');
    add('cp', 'CP', e => e.cp || '');
    add('rm', 'RM', e => e.rm || '');
    add('floor', 'Floor', e => e.floor || '');
    add('area', 'Area', e => num(e.area), { n: 1 });
    add('rate', 'Rate', e => num(e.rate), { n: 1 });
    add('gcp', 'Gold for CP', e => e.goldCp || '');
    add('gcl', 'Gold for client', e => e.goldClient || '');
    if (!o.hideLps) add('lps', 'LPS Dis.', e => num(e.lps), { n: 1 });
    add('cpd', 'CP Dis.', e => num(e.cpDisc), { n: 1 });
    add('iph', 'I Phone Adjustment', e => num(e.iphone), { n: 1 });
    add('tpd', 'TPD (within 30 days)', e => num(e.tpd), { n: 1 });
    add('spot', 'On Spot Discount', e => num(e.spot), { n: 1 });
    add('ot', 'Other Discount Type', e => e.otherDiscType || '');
    add('od', 'Other Discount Amount', e => num(e.otherDisc), { n: 1 });
    add('tdisc', 'Total Discount', (e, c) => c.totalDisc, { n: 1, f: (r, L) => 'SUM(' + ['lps', 'cpd', 'iph', 'tpd', 'spot', 'od'].map(k => L[k] + r).join(',') + ')' });
    add('dd', 'Discount Date', e => e.discountDate || '');
    add('dpct', 'Due %', e => num(e.duePct), { n: 1 });
    add('mdue', 'Manual Due Override', e => num(e.manualDue) > 0 ? num(e.manualDue) : '', { n: 1 });
    add('mar', 'Manual AR Monthly', e => num(e.manualArMonthly) > 0 ? num(e.manualArMonthly) : '', { n: 1 });
    add('status', 'Status', e => e.status || '');
    add('psplit', 'Payee Split', (e, c) => c.hasPayeeSplit ? c.payeeSplits.map(p => `${p.name}: ${p.sharePct}% (Net ₹${p.net})`).join(' | ') : '');
    add('pcnt', 'Total Payment Count', e => pays(e).length);
    add('pdates', 'All Payment Dates', e => pays(e).map(p => p.date).filter(Boolean).join(', '));
    add('pamts', 'All Payment Amounts', e => pays(e).map(p => num(p.amount)).join(', '));
    add('pmodes', 'All Payment Modes', e => pays(e).map(p => p.mode).filter(Boolean).join(', '));
    add('pdet', 'All Payments (Detail)', e => pays(e).map(p => `${p.date || ''} | ₹${num(p.amount).toFixed(2)} | ${p.mode || ''}`).join('\n'));
    add('bank', 'Bank', e => e.bank || '');
    add('ref', 'Ref No.', e => e.ref || '');
    add('pan', 'PAN No.', e => e.pan || '');
    add('fav', 'Favouring', e => e.favouring || '');
    add('chq', 'Last Issued Cheque No.', e => e.chequeNo || '');
    add('ctrl', 'Control No.', e => e.controlNo || '');
    add('gross', 'Gross Amount', (e, c) => c.gross, { n: 1, f: (r, L) => `${L.area}${r}*${L.rate}${r}` });
    add('due', 'Due Amount', (e, c) => c.due, { n: 1, f: (r, L) => `IF(N(${L.mdue}${r})>0,${L.mdue}${r},${L.gross}${r}*${L.dpct}${r}/100)` });
    add('grecv', 'Total Received (Gross)', (e, c) => c.grossReceived, { n: 1 });
    add('gst', 'GST Total', (e, c) => c.gstTotal, { n: 1 });
    add('recv', 'Amount Received (Net)', (e, c) => c.received, { n: 1, f: (r, L) => `${L.grecv}${r}-${L.gst}${r}` });
    add('bal', 'Balance', (e, c) => c.balance, { n: 1, f: (r, L) => `${L.due}${r}-${L.recv}${r}` });
    add('elig', 'AR Eligible', (e, c) => c.eligible ? 'Yes' : 'No', { f: (r, L, e) => e.isImported ? `IF(${L.due}${r}>0,"Yes","No")` : `IF(AND(${L.due}${r}>0,${L.recv}${r}+${L.tdisc}${r}>=${L.due}${r}),"Yes","No")` });
    add('per', 'AR Period (Months)', (e, c) => c.period, { n: 1 });
    add('arm', 'AR Monthly (1% of Received + Discount)', (e, c) => c.arMonthly, { n: 1, f: (r, L, e) => num(e.manualArMonthly) > 0 ? null : `ROUND((${L.recv}${r}+${L.tdisc}${r})*0.01,0)` });
    add('tds', 'TDS (10%)', (e, c) => c.tds, { n: 1, f: (r, L) => `ROUND(${L.arm}${r}*0.1,0)` });
    add('netm', 'Net Monthly', (e, c) => c.net, { n: 1, f: (r, L) => `${L.arm}${r}-${L.tds}${r}` });
    add('tp', 'Total Payable', (e, c) => c.totalPayable, { n: 1, f: (r, L) => `ROUND(${L.netm}${r}*${L.per}${r},0)` });

    const colL = i => { let s = '', n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    const L = {}; C.forEach((c, i) => { L[c.k] = colL(i); });
    const wb = new ExcelJS.Workbook(); wb.creator = 'AR Calculation Tool'; wb.created = new Date();
    const ws = wb.addWorksheet('AR Entries', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = C.map(c => ({ width: Math.max(14, Math.min(34, c.h.length + 4)) }));
    const thin = { style: 'thin', color: { argb: 'FFB9C2D6' } }, bd = { top: thin, left: thin, bottom: thin, right: thin };
    const hr = ws.getRow(1); hr.height = 34;
    C.forEach((c, i) => { const cell = hr.getCell(i + 1); cell.value = c.h; cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A1E3D' } }; cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cell.border = bd; });
    entries.forEach((e, i) => {
        const c = computeEntryValues(e), r = i + 2;
        C.forEach((col, j) => {
            let val = col.v(e, c, i); if (val === undefined || val === null) val = '';
            const f = (o.formulas && col.f) ? col.f(r, L, e) : null;
            const cell = ws.getCell(r, j + 1);
            cell.value = f ? { formula: f, result: val } : val;
            if (col.n) cell.numFmt = '#,##0.00';
            cell.border = bd;
        });
    });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: C.length } };
    expDownload(wb, `AR_Calculation_Tool_${new Date().toISOString().slice(0, 10)}${o.formulas ? '' : '_values'}.xlsx`);
}

/* =====================================================================
   ER REGISTER (separate storage: Firebase node "er_entries" / localStorage)
   ===================================================================== */
const ER_LS = 'ar_er_register_v1';
let erEntries = [], erEditKey = null;
const erMon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function erFmtDate(iso) { if (!iso) return '—'; const d = new Date(iso + 'T00:00:00'); if (isNaN(d)) return iso; return String(d.getDate()).padStart(2, '0') + '-' + erMon[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2); }
function erGetLS() { try { return JSON.parse(localStorage.getItem(ER_LS) || '{}'); } catch (e) { return {}; } }
function erSetLS(o) { try { localStorage.setItem(ER_LS, JSON.stringify(o)); } catch (e) { console.error(e); } }
function erLoad() {
    if (usingFirebase) database.ref('er_entries').on('value', s => erProcess(s.val()), err => showToast('ER load error: ' + err.message, true));
    else erProcess(erGetLS());
}
function erProcess(data) {
    erEntries = Object.entries(data || {}).map(([k, v]) => Object.assign({}, v, { _key: k })).sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
    erRender();
}
function erWrite(key, rec) {
    if (usingFirebase) { const ref = key ? database.ref('er_entries/' + key) : database.ref('er_entries').push(); return ref.set(rec).catch(err => { showToast('ER save error: ' + err.message, true); throw err; }); }
    const d = erGetLS(); const k = key || ('er_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));
    d[k] = rec; erSetLS(d); erProcess(d); return Promise.resolve();
}
function erRemove(key) {
    if (usingFirebase) return database.ref('er_entries/' + key).remove();
    const d = erGetLS(); delete d[key]; erSetLS(d); erProcess(d); return Promise.resolve();
}
function erVal(id) { return document.getElementById(id).value.trim(); }
function erSave() {
    const rec = { party: erVal('er_party'), ref: erVal('er_ref'), amount: num(erVal('er_amount')), dueDate: erVal('er_due'), chNo: erVal('er_chno') };
    if (!rec.party) { showToast('Party name is required', true); document.getElementById('er_party').focus(); return; }
    const old = erEditKey ? erEntries.find(e => e._key === erEditKey) : null;
    rec.createdAt = old ? (old.createdAt || Date.now()) : Date.now();
    erWrite(erEditKey, rec).then(() => { showToast(erEditKey ? 'ER entry updated' : 'ER entry saved', false); erCancel(); }).catch(() => {});
}
function erCancel() {
    erEditKey = null;
    ['er_party', 'er_ref', 'er_amount', 'er_due', 'er_chno'].forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById('erFormTitle').textContent = 'New ER Entry';
}
function erEdit(key) {
    const e = erEntries.find(x => x._key === key); if (!e) return;
    erEditKey = key;
    document.getElementById('er_party').value = e.party || '';
    document.getElementById('er_ref').value = e.ref || '';
    document.getElementById('er_amount').value = e.amount || '';
    document.getElementById('er_due').value = e.dueDate || '';
    document.getElementById('er_chno').value = e.chNo || '';
    document.getElementById('erFormTitle').textContent = 'Edit ER Entry';
    document.getElementById('erView').scrollIntoView({ behavior: 'smooth' });
}
function erDelete(key) {
    const e = erEntries.find(x => x._key === key); if (!e) return;
    if (!confirm(`Delete ER entry for "${e.party}"?`)) return;
    erRemove(key).then(() => showToast('ER entry deleted', false)).catch(err => showToast('Error: ' + err.message, true));
}
function erRender() {
    const tb = document.getElementById('erTbody'); if (!tb) return;
    const q = (document.getElementById('erSearch').value || '').toLowerCase().trim();
    const badge = document.getElementById('erCountBadge'); if (badge) badge.textContent = erEntries.length;
    let total = 0, shown = 0;
    tb.innerHTML = erEntries.map((e, i) => {
        if (q && !((e.party || '') + ' ' + (e.ref || '') + ' ' + (e.chNo || '')).toLowerCase().includes(q)) return '';
        total += num(e.amount); shown++;
        return `<tr><td>${i + 1}</td><td>${escapeHtml(e.party)}</td><td>${escapeHtml(e.ref || '—')}</td><td class="num">${num(e.amount).toLocaleString('en-IN')}</td><td>${erFmtDate(e.dueDate)}</td><td class="num">${escapeHtml(e.chNo || '—')}</td><td class="actions-cell"><button title="Edit" onclick="erEdit('${e._key}')"><i class="fa-solid fa-pen"></i></button><button title="Delete" class="del" onclick="erDelete('${e._key}')"><i class="fa-solid fa-trash"></i></button></td></tr>`;
    }).join('');
    document.getElementById('erEmpty').style.display = shown ? 'none' : 'block';
    document.getElementById('erTotal').textContent = shown ? `Total (${shown} entries): ₹ ${total.toLocaleString('en-IN')}` : '';
}
function erExport() {
    if (!erEntries.length) { showToast('No ER entries to export', true); return; }
    const rows = [['S.No', 'Party', 'EOI / Ref', 'Amount (Rs)', 'Cheque Due Date', 'Ch No.']];
    let total = 0;
    erEntries.forEach((e, i) => { total += num(e.amount); rows.push([i + 1, e.party || '', e.ref || '', num(e.amount), erFmtDate(e.dueDate), String(e.chNo || '')]); });
    const ws = XLSX.utils.aoa_to_sheet(rows), n = rows.length;
    ws['C' + (n + 1)] = { t: 's', v: 'Total' };
    ws['D' + (n + 1)] = { t: 'n', f: 'SUM(D2:D' + n + ')', v: total };
    ws['!ref'] = 'A1:F' + (n + 1);
    ws['!cols'] = [{ wch: 6 }, { wch: 26 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 12 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'ER Register');
    XLSX.writeFile(wb, `ER_Register_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast('ER Register exported', false);
}
function erPrint() {
    if (!erEntries.length) { showToast('No ER entries to print', true); return; }
    let total = 0;
    const rows = erEntries.map((e, i) => { total += num(e.amount); return `<tr><td class="c">${i + 1}</td><td>${escapeHtml(e.party)}</td><td>${escapeHtml(e.ref || '—')}</td><td class="r">${num(e.amount).toLocaleString('en-IN')}</td><td class="c">${erFmtDate(e.dueDate)}</td><td class="c">${escapeHtml(e.chNo || '—')}</td></tr>`; }).join('');
    doPrint(`<div class="pr-page">${buildPrintBrand()}<div class="pr-title">ER REGISTER</div><table class="pr-data"><thead><tr><th>S.No</th><th>Party</th><th>EOI / Ref</th><th>Amount (Rs)</th><th>Cheque Due Date</th><th>Ch No.</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="3" class="r">Total</td><td class="r">${total.toLocaleString('en-IN')}</td><td colspan="2"></td></tr></tfoot></table><div class="pr-sign"><div>Prepared By</div><div>Checked By</div><div>Authorised Signatory</div></div></div>`);
}

function patchV2Init() {
    document.querySelector('nav.tabs').insertAdjacentHTML('beforeend', `<div class="tab" data-view="erView" onclick="switchTab('erView');erRender()"><i class="fa-solid fa-receipt"></i> ER Register <span class="count" id="erCountBadge">0</span></div>`);
    document.getElementById('toast').insertAdjacentHTML('beforebegin', `<div class="view" id="erView">
<div class="card"><div class="card-head"><h3><i class="fa-solid fa-receipt"></i> <span id="erFormTitle">New ER Entry</span></h3></div>
<div class="card-body"><div class="grid grid-5">
<div class="field"><label>Party <span class="req">*</span></label><input id="er_party" placeholder="e.g. Anju Bisht"></div>
<div class="field"><label>EOI / Ref</label><input id="er_ref" placeholder="e.g. EOI-245"></div>
<div class="field"><label>Amount (Rs)</label><input id="er_amount" type="number" step="any" placeholder="0"></div>
<div class="field"><label>Cheque Due Date</label><input id="er_due" type="date"></div>
<div class="field"><label>Ch No.</label><input id="er_chno" placeholder="e.g. 000081"></div>
</div></div>
<div class="form-actions"><button class="btn btn-outline" onclick="erCancel()"><i class="fa-solid fa-xmark"></i> Cancel</button><button class="btn btn-primary" onclick="erSave()"><i class="fa-solid fa-floppy-disk"></i> Save ER Entry</button></div></div>
<div class="card"><div class="card-head"><h3><i class="fa-solid fa-table-list"></i> ER Register</h3><div style="display:flex;gap:8px;flex-wrap:wrap;"><button class="btn btn-outline btn-sm" onclick="erPrint()"><i class="fa-solid fa-print"></i> Print</button><button class="btn btn-success btn-sm" onclick="erExport()"><i class="fa-solid fa-file-excel"></i> Export Excel</button></div></div>
<div style="padding:14px 20px 6px;"><div class="search-bar"><i class="fa-solid fa-search" style="color:var(--text-mute);"></i><input type="text" id="erSearch" placeholder="Search by Party, EOI / Ref, Ch No...." oninput="erRender()"></div></div>
<div class="table-scroll"><table class="data-table" style="min-width:760px;"><thead><tr><th>S.No</th><th>Party</th><th>EOI / Ref</th><th>Amount (Rs)</th><th>Cheque Due Date</th><th>Ch No.</th><th>Actions</th></tr></thead><tbody id="erTbody"></tbody></table></div>
<div class="empty-state" id="erEmpty" style="display:none;"><i class="fa-regular fa-folder-open"></i><h4>No ER entries</h4><p>Add an entry using the form above.</p></div>
<div style="padding:14px 20px;font-weight:700;text-align:right;font-family:var(--font-num);" id="erTotal"></div></div></div>`);
    const st = document.createElement('style');
    st.textContent = `.ops-box{max-width:430px;text-align:center;padding:32px 28px 26px;border-top:5px solid var(--red)}
.ops-icon{width:70px;height:70px;border-radius:50%;background:var(--red-bg);color:var(--red);display:flex;align-items:center;justify-content:center;font-size:30px;margin:0 auto 16px}
.ops-box h3{font-size:20px;color:var(--ink);margin-bottom:8px}
.ops-box p{font-size:13.5px;color:var(--text-mute);line-height:1.6;margin:0 0 24px}
.ops-actions{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.ops-btn{display:flex;align-items:center;justify-content:center;gap:9px;padding:12px 14px;border-radius:10px;border:1px solid transparent;font-size:14px;font-weight:700;cursor:pointer;transition:transform .15s ease,box-shadow .15s ease,background .15s ease}
.ops-btn:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(10,30,61,.18)}
.ops-btn:focus-visible{outline:3px solid rgba(10,30,61,.25);outline-offset:2px}
.ops-danger{background:var(--red);color:#fff}.ops-danger:hover{background:#a83723}
.ops-dark{background:var(--ink);color:#fff}.ops-dark:hover{background:var(--ink-soft)}
.ops-ghost{grid-column:1/-1;background:#fff;color:var(--ink);border-color:var(--line)}.ops-ghost:hover{border-color:var(--ink)}
@media(max-width:420px){.ops-actions{grid-template-columns:1fr}}`;
    document.head.appendChild(st);
    document.body.insertAdjacentHTML('beforeend', `<div class="modal-overlay" id="exportModal" role="dialog" aria-modal="true" aria-labelledby="opsTitle"><div class="modal-box ops-box">
<div class="ops-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
<h3 id="opsTitle">${EXPORT_POPUP_TITLE}</h3>
<p>${EXPORT_POPUP_TEXT}</p>
<div class="ops-actions">
<button class="ops-btn ops-danger" onclick="runExport(true)"><i class="fa-solid fa-trash"></i> ${EXPORT_LABEL_FORMULA}</button>
<button class="ops-btn ops-dark" onclick="runExport(false)"><i class="fa-solid fa-ban"></i> ${EXPORT_LABEL_PLAIN}</button>
<button class="ops-btn ops-ghost" id="opsBack" onclick="exportBack()"><i class="fa-solid fa-arrow-left"></i> ${EXPORT_LABEL_BACK}</button>
</div></div></div>`);
    erLoad();
}
setTimeout(patchV2Init, 0);


/* ---------------- INIT ---------------- */

/* ============ DATA SHARING CONTROL ============ */
let isDataSharingOn = false;

function initDataSharingControl() {
    // Jab tak Firebase se jawab na aaye, page chhupa rahega (default LOCKED)
    document.body.style.visibility = 'hidden';

    let answered = false;
    const show = (isOn) => {
        answered = true;
        applyDataSharingState(isOn);
        document.body.style.visibility = '';
    };

    // Firebase connect nahi hua -> locked
    if (!usingFirebase || !database) {
        show(false);
        return;
    }

    // 8 second me jawab na aaye -> locked
    setTimeout(() => { if (!answered) show(false); }, 8000);

    const controlRef = database.ref('control/data_sharing_on');
    controlRef.on('value', (snapshot) => {
        isDataSharingOn = (snapshot.val() === true); // sirf true hone par ON
        show(isDataSharingOn);
    }, (error) => {
        console.error('Error listening to control node:', error);
        show(false); // error par bhi locked
    });
}
// ============ END DATA SHARING CONTROL ============

function init() { 
    initStorage(); 
    loadEntriesFromFirebase(); 
    initDataSharingControl(); // <-- Added for ON/OFF switch
    renderPaymentRows(); 
    renderEntriesTable(); 
    renderDashboard(); 
}
init();

window.switchTab = switchTab;
window.addPaymentRow = addPaymentRow;
window.updatePayField = updatePayField;
window.removePaymentRow = removePaymentRow;
window.recalcForm = recalcForm;
window.clearForm = clearForm;
window.cancelEdit = cancelEdit;
window.saveEntry = saveEntry;
window.editEntry = editEntry;
window.duplicateEntry = duplicateEntry;
window.duplicateFromForm = duplicateFromForm;
window.deleteEntry = deleteEntry;
window.confirmDelete = confirmDelete;
window.closeDeleteModal = closeDeleteModal;
window.openResetModal = openResetModal;
window.closeResetModal = closeResetModal;
window.confirmReset = confirmReset;
window.exportExcel = exportExcel;
window.exportAllSchedules = exportAllSchedules;
window.renderDashboard = renderDashboard;
window.renderEntriesTable = renderEntriesTable;
window.openLotModal = openLotModal;
window.closeLotModal = closeLotModal;
window.onLotSelectChange = onLotSelectChange;
window.updateLotPreview = updateLotPreview;
window.confirmLotAction = confirmLotAction;
window.updateChequeNo = updateChequeNo;
window.updateLotControlNo = updateLotControlNo;
window.copySingleNarration = copySingleNarration;
window.copyAllNarrationsFromForm = copyAllNarrationsFromForm;
window.exportChequeSchedule = exportChequeSchedule;
window.openChequeEditModal = openChequeEditModal;
window.closeChequeEditModal = closeChequeEditModal;
window.saveChequeEdit = saveChequeEdit;
window.openLotModalForLot = openLotModalForLot;
window.deleteLot = deleteLot;
window.autoGenerateSchedule = autoGenerateSchedule;
window.bulkAutoGenerateSchedules = bulkAutoGenerateSchedules;
window.openTopUpLotModal = openTopUpLotModal;
window.onLotIncModeChange = onLotIncModeChange;
window.markAllExistingAsImported = markAllExistingAsImported;
window.toggleEntrySelection = toggleEntrySelection;
window.toggleSelectAll = toggleSelectAll;
window.bulkDeleteEntries = bulkDeleteEntries;
window.bulkChangeCompany = bulkChangeCompany;
window.bulkToggleImported = bulkToggleImported;
window.updateBulkDeleteUI = updateBulkDeleteUI;
window.clearSelection = clearSelection;
window.resetFilters = resetFilters;
window.closeImportModeModal = closeImportModeModal;
window.executeImport = executeImport;
window.openPayeeModal = openPayeeModal;
window.closePayeeModal = closePayeeModal;
window.clearPayees = clearPayees;
window.addPayeeRow = addPayeeRow;
window.removePayeeRow = removePayeeRow;
window.updatePayeeField = updatePayeeField;
window.savePayees = savePayees;
window.toggleMoreMenu = toggleMoreMenu;
window.closeMoreMenu = closeMoreMenu;
window.printEntry = printEntry;
window.printCurrentEntry = printCurrentEntry;
window.printAllEntries = printAllEntries;
