'use strict';

// ── CONFIG ──────────────────────────────────────────────────────────────────
const CLIENT_ID = '75e627e8-2de0-4ec6-bec9-311757b89e08';
const TENANT_ID = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES    = ['User.Read', 'Sites.ReadWrite.All'];
const SP_SITE   = 'dihag.sharepoint.com:/sites/gruppe_shb';
const SP_LIST   = 'Bedarfsanfrage';
const API       = 'https://graph.microsoft.com/v1.0';

// ── BPMN-KONFORME FORMFELD-DEFINITION ───────────────────────────────────────
// key        = interner SP-Spaltenname (wird beim POST verwendet)
// alsoTry    = alternative interne Namen falls key nicht gefunden
// step       = Wizard-Schritt (1-3)
const FORM_FIELDS = [
  // Step 1: Bedarf
  { key:'Title',             label:'Bezeichnung',                  step:1, required:true  },
  { key:'Beschreibung',      label:'Beschreibung / Begründung',    step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung'] },
  { key:'Warengruppe',       label:'Warengruppe',                  step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',        label:'Priorität',                    step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  // Step 2: Menge
  { key:'Menge',             label:'Menge',                        step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',     label:'Mengeneinheit',                step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge', label:'Mindestlagermenge',            step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',            label:'Benötigt bis',                 step:2, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis'] },
  // Step 3: Beschaffung
  { key:'Beschaffungslogik', label:'Beschaffungslogik',            step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Artikelnummer',     label:'Artikelnummer / Nummernangaben',step:3,alsoTry:['MaterialNumber','ItemNumber'] },
  { key:'Lieferant',         label:'Lieferant (Lieferanten-Logik)',step:3, alsoTry:['Vendor','Supplier'] },
  { key:'GeschaetzterPreis', label:'Geschätzter Preis netto (€)',  step:3, alsoTry:['EstimatedPrice','Preis','Price'] },
  { key:'Kostenstelle',      label:'Kostenstelle',                 step:3, alsoTry:['CostCenter'] },
];

// Felder die Einkauf nach der Einreichung befüllt
const EINKAUF_FIELDS = [
  { key:'Bestellnummer',  label:'Bestellnummer',        alsoTry:['OrderNumber','PO_Number'] },
  { key:'Lieferdatum',    label:'Lieferdatum',          alsoTry:['DeliveryDate'] },
  { key:'TatsaechlicherPreis', label:'Tatsächlicher Preis (€)', alsoTry:['ActualPrice','FinalPrice'] },
];

// Status-Werte → Darstellung (kommt von Power Automate)
const STATUS_STYLES = {
  'eingereicht':   { bg:'#eff6ff', color:'#1d4ed8' },
  'in prüfung':    { bg:'#fffbeb', color:'#b45309' },
  'freigegeben':   { bg:'#f0fdf4', color:'#15803d' },
  'abgelehnt':     { bg:'#fef2f2', color:'#b91c1c' },
  'bestellt':      { bg:'#faf5ff', color:'#7e22ce' },
  'erledigt':      { bg:'#f3f4f6', color:'#374151' },
};

const SYSTEM_FIELDS = new Set([
  '@odata.etag','@odata.id','id','ContentType','Modified','Created',
  'AuthorLookupId','EditorLookupId','Attachments','Edit','LinkTitleNoMenu','LinkTitle',
  'ItemChildCount','FolderChildCount','_UIVersionString','ComplianceAssetId',
  'OData__ColorTag','AppAuthorLookupId','AppEditorLookupId'
]);

// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp, account;
let siteId = null, listId = null;
let allItems = [];
let colByKey   = {};  // internal name → column definition (from SP)
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let currentView = 'dashboard';
let prevView    = 'dashboard';
let wizardData  = {};

// ── AUTH ────────────────────────────────────────────────────────────────────
async function initAuth() {
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth: { clientId:CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation:'localStorage', storeAuthStateInCookie:true }
  });
  await msalApp.initialize();
  await msalApp.handleRedirectPromise();
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) { account = accounts[0]; return true; }
  return false;
}

async function doLogin() {
  $id('boot-btn').style.display = 'none';
  $id('boot-sub').textContent = 'Anmeldung läuft…';
  $id('boot-spinner').style.display = 'block';
  try {
    const r = await msalApp.loginPopup({ scopes: SCOPES });
    account = r.account;
    bootDone();
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
  }
}

async function bootDone() {
  $id('boot-sub').textContent = 'Daten werden geladen…';
  try {
    await discoverSP();
    await loadItems(false);
    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';
    // Set user info in sidebar
    const name = account?.name || account?.username || '?';
    $id('hdr-av').textContent   = name.split(' ').map(p=>p[0]||'').join('').substring(0,2).toUpperCase();
    $id('hdr-name').textContent = name;
    $id('hdr-mail').textContent = account?.username || '';
    navigate('dashboard');
  } catch(e) {
    $id('boot-sub').textContent = 'Fehler beim Laden: ' + e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-err').textContent = e.message;
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
    $id('boot-btn').onclick = bootDone;
  }
}

function doLogout() {
  msalApp?.logoutPopup({ account }).catch(()=>{});
  location.reload();
}

async function getToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try   { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account})).accessToken; }
  catch { return (await msalApp.acquireTokenPopup ({scopes:SCOPES, account})).accessToken; }
}

// ── GRAPH API ────────────────────────────────────────────────────────────────
async function gGet(path) {
  const tok = await getToken();
  const r   = await fetch(API + path, { headers:{ Authorization:'Bearer '+tok } });
  if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
async function gPost(path, body) {
  const tok = await getToken();
  const r   = await fetch(API + path, {
    method:'POST', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (!r.ok) throw new Error(`Graph POST ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
async function gPatch(path, body) {
  const tok = await getToken();
  const r   = await fetch(API + path, {
    method:'PATCH', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (r.status === 204) return {};
  if (!r.ok) throw new Error(`Graph PATCH ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json().catch(()=>({}));
}

// ── SP DISCOVERY ─────────────────────────────────────────────────────────────
async function discoverSP() {
  // 1. Site ID
  const site = await gGet(`/sites/${SP_SITE}`);
  siteId = site.id;

  // 2. List ID
  const list = await gGet(`/sites/${siteId}/lists/${encodeURIComponent(SP_LIST)}`);
  listId = list.id;

  // 3. Columns
  const colsRes = await gGet(`/sites/${siteId}/lists/${listId}/columns`);
  colByKey = {};
  for (const c of colsRes.value) {
    colByKey[c.name] = c;  // by internal name
  }

  // 4. Resolve each form field to actual SP internal name
  resolvedFields = {};
  const missing = [];
  for (const fd of [...FORM_FIELDS, ...EINKAUF_FIELDS]) {
    const found = resolveColName(fd);
    resolvedFields[fd.key] = found;
    if (!found && fd.key !== 'Title') missing.push(fd.label);
  }
  // Title always exists
  resolvedFields['Title'] = 'Title';

  // Show warning for missing columns
  if (missing.length) {
    const w = $id('col-warning');
    w.innerHTML = `<strong>⚠ Folgende Spalten fehlen noch in der SharePoint-Liste</strong> – bitte in der SP-Liste anlegen, damit alle Felder gespeichert werden können:<br>
      <em>${missing.join(' · ')}</em>`;
    w.style.display = 'block';
  }

  // Populate Choice selects from actual SP column definitions
  for (const fd of FORM_FIELDS) {
    const colName = resolvedFields[fd.key];
    if (!colName) continue;
    const col = colByKey[colName];
    if (!col?.choice?.choices?.length) continue;
    const el = document.getElementById('f-' + fd.key);
    if (!el || el.tagName !== 'SELECT') continue;
    el.innerHTML = fd.required ? '' : '<option value="">– bitte wählen –</option>';
    for (const c of col.choice.choices) el.add(new Option(c, c));
  }
}

function resolveColName(fd) {
  if (fd.key === 'Title') return 'Title';
  // Only consider writable, non-sealed columns
  const ok = c => c && !c.readOnly && !c.sealed;
  if (ok(colByKey[fd.key])) return fd.key;
  for (const alt of (fd.alsoTry || [])) {
    if (ok(colByKey[alt])) return alt;
  }
  // Case-insensitive display-name match (writable only)
  const labelNorm = (fd.label || '').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
  for (const [k, c] of Object.entries(colByKey)) {
    if (!ok(c)) continue;
    const dn = (c.displayName||'').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
    if (dn === labelNorm) return k;
  }
  return null;
}

// ── LOAD ITEMS ───────────────────────────────────────────────────────────────
async function loadItems(showToast = true) {
  const btn = $id('btn-reload');
  if (btn) btn.disabled = true;
  try {
    const data = await gGet(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500&$orderby=createdDateTime desc`
    );
    allItems = data.value || [];
    if (showToast) toast('Daten aktualisiert', 'success');
    // Re-render current view
    if (currentView === 'dashboard') renderDashboard();
    else if (currentView === 'mine')  renderList('mine');
    else if (currentView === 'all')   renderList('all');
  } catch(e) {
    toast('Fehler beim Laden: ' + e.message, 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ── ROUTING ──────────────────────────────────────────────────────────────────
const VIEW_TITLES = { dashboard:'Dashboard', new:'Neue Bedarfsanfrage',
  mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details' };

function navigate(view, id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item[data-view]').forEach(n => n.classList.remove('active'));

  const el = $id('view-' + view);
  if (!el) return;
  el.classList.add('active');

  const nav = document.querySelector(`.nav-item[data-view="${view}"]`);
  if (nav) nav.classList.add('active');

  $id('page-title').textContent = VIEW_TITLES[view] || view;
  prevView = currentView;
  currentView = view;

  if      (view === 'dashboard') renderDashboard();
  else if (view === 'mine')      renderList('mine');
  else if (view === 'all')       renderList('all');
  else if (view === 'new')       initWizard();
  else if (view === 'detail' && id) renderDetail(id);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Nav clicks
  document.querySelectorAll('.nav-item[data-view]').forEach(a => {
    a.addEventListener('click', e => { e.preventDefault(); navigate(a.dataset.view); });
  });
  // Dashboard "Neue Anfrage" button
  document.querySelectorAll('[data-view="new"]').forEach(b => {
    b.addEventListener('click', e => { e.preventDefault(); navigate('new'); });
  });
  // Detail back button
  $id('detail-back').addEventListener('click', () => navigate(prevView));
  // Menu toggle
  $id('menu-toggle').addEventListener('click', () => $id('sidebar').classList.toggle('open'));

  // Boot
  $id('boot-spinner').style.display = 'block';
  try {
    const loggedIn = await initAuth();
    if (loggedIn) await bootDone();
    else {
      $id('boot-sub').textContent = 'Bitte melden Sie sich an.';
      $id('boot-spinner').style.display = 'none';
      $id('boot-btn').style.display = 'block';
    }
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
  }
});

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function renderDashboard() {
  const total    = allItems.length;
  const open     = allItems.filter(i => isOpenStatus(getField(i,'Status'))).length;
  const approved = allItems.filter(i => {
    const s = (getField(i,'Status')||'').toLowerCase();
    return s.includes('freigegeben') || s.includes('bestellt') || s.includes('erledigt');
  }).length;
  const volume   = allItems.reduce((s,i) => s + (parseFloat(getField(i,'GeschaetzterPreis') ||
    getField(i,resolvedFields['GeschaetzterPreis'])) || 0), 0);

  $id('st-total').textContent    = total;
  $id('st-open').textContent     = open;
  $id('st-approved').textContent = approved;
  $id('st-volume').textContent   = fmtEuro(volume);

  // Recent 8 items
  const recent = allItems.slice(0, 8);
  $id('dash-recent').innerHTML = recent.length
    ? recent.map(i => compactRow(i)).join('')
    : emptyState('Noch keine Anfragen vorhanden');
}

function isOpenStatus(s) {
  if (!s) return true; // no status = freshly submitted, counts as open
  const sl = s.toLowerCase();
  return sl.includes('eingereicht') || sl.includes('prüfung') || sl.includes('offen');
}

// ── LIST VIEWS ────────────────────────────────────────────────────────────────
function renderList(type) {
  // Populate status filter options once
  const sel = $id(`filter-${type}-status`);
  if (sel.options.length <= 1) {
    const statuses = [...new Set(allItems.map(i => getField(i,'Status')).filter(Boolean))];
    statuses.forEach(s => sel.add(new Option(s, s)));
  }
  filterView(type);
}

function filterView(type) {
  const search = ($id(`search-${type}`)?.value || '').toLowerCase();
  const status = $id(`filter-${type}-status`)?.value || '';
  const userName = (account?.name || account?.username || '').toLowerCase();

  let items = [...allItems];
  if (type === 'mine') {
    items = items.filter(i => {
      const author = (getField(i,'Author') || getField(i,'AuthorLookupId') || '').toString().toLowerCase();
      const creator = (i.createdBy?.user?.displayName || i.createdBy?.user?.email || '').toLowerCase();
      return author.includes(userName.split(' ')[0]) || creator.includes(userName.split(' ')[0]);
    });
  }
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) ||
    String(i.id||'').includes(search)
  );
  if (status) items = items.filter(i => (getField(i,'Status')||'') === status);

  const container = $id(`list-${type}`);
  container.innerHTML = items.length
    ? items.map(i => itemCard(i)).join('')
    : emptyState(type === 'mine' ? 'Sie haben noch keine Anfragen erstellt.' : 'Keine Anfragen gefunden.');
}

// ── DETAIL VIEW ───────────────────────────────────────────────────────────────
function renderDetail(id) {
  const item = allItems.find(i => String(i.id) === String(id));
  if (!item) { $id('detail-content').innerHTML = '<p>Eintrag nicht gefunden.</p>'; return; }
  const f = item.fields || {};

  // Group known fields
  const bedarfKeys   = FORM_FIELDS.filter(fd => fd.step === 1).map(fd => resolvedFields[fd.key] || fd.key);
  const mengeKeys    = FORM_FIELDS.filter(fd => fd.step === 2).map(fd => resolvedFields[fd.key] || fd.key);
  const beschKeys    = FORM_FIELDS.filter(fd => fd.step === 3).map(fd => resolvedFields[fd.key] || fd.key);
  const einkaufKeys  = EINKAUF_FIELDS.map(fd => resolvedFields[fd.key] || fd.key);
  const knownKeys    = new Set([...bedarfKeys,...mengeKeys,...beschKeys,...einkaufKeys,'Status','Title']);

  // Remaining SP fields not in our definition
  const extraFields  = Object.entries(f)
    .filter(([k]) => !SYSTEM_FIELDS.has(k) && !knownKeys.has(k))
    .filter(([,v]) => v !== null && v !== undefined && v !== '');

  const createdBy    = item.createdBy?.user?.displayName || item.createdBy?.user?.email || getField(item,'Author') || '–';
  const createdAt    = item.createdDateTime ? fmtDateTime(item.createdDateTime) : '–';
  const statusVal    = getField(item,'Status') || 'Eingereicht';
  const preis        = parseFloat(getField(item,'GeschaetzterPreis') || getField(item, resolvedFields['GeschaetzterPreis'])) || null;
  const menge        = getField(item,'Menge') || getField(item, resolvedFields['Menge']);
  const isEinkauf    = true; // for now, all logged-in users can add order info (adjust if needed)

  $id('detail-content').innerHTML = `
    <div class="detail-header">
      <h2>${esc(getField(item,'Title') || '–')}</h2>
      <div class="detail-meta">
        <span class="item-id">ID ${item.id}</span>
        ${statusBadge(statusVal)}
        ${prioTag(getField(item,'Prioritaet') || getField(item, resolvedFields['Prioritaet']))}
      </div>
      <div class="detail-byline">Erstellt von <strong>${esc(createdBy)}</strong> am ${createdAt}</div>
    </div>

    <div class="detail-grid">
      <div class="detail-left">
        ${detailSection('Bedarf', [
          ...FORM_FIELDS.filter(fd=>fd.step===1).map(fd => detailRow(fd.label, getField(item, resolvedFields[fd.key] || fd.key)))
        ])}
        ${detailSection('Mengengaben', [
          ...FORM_FIELDS.filter(fd=>fd.step===2).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${detailSection('Beschaffungsdetails', [
          ...FORM_FIELDS.filter(fd=>fd.step===3).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${extraFields.length ? detailSection('Weitere Felder aus SharePoint', extraFields.map(([k,v]) => detailRow(k, v))) : ''}
      </div>

      <div class="detail-right">
        ${detailSidebar(item, isEinkauf)}
      </div>
    </div>`;

  // Bind update button
  const btnOrder = document.getElementById('btn-add-order');
  if (btnOrder) btnOrder.onclick = () => openOrderModal(item.id);
}

function detailSection(title, rows) {
  const content = rows.filter(r => r).join('');
  if (!content) return '';
  return `<div class="detail-card"><div class="detail-card-header">${title}</div><div class="detail-card-body">${content}</div></div>`;
}
function detailRow(label, value) {
  if (!value && value !== 0) return '';
  return `<div class="detail-row"><span class="detail-label">${esc(String(label))}</span><span class="detail-value">${esc(String(value))}</span></div>`;
}
function formatFieldValue(fd, v) {
  if (!v && v !== 0) return '';
  if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') return fmtEuro(v);
  if (fd.key === 'Termin' || fd.key === 'Lieferdatum') return v ? fmtDate(v) : '';
  return v;
}

function detailSidebar(item, isEinkauf) {
  const orderNr   = getField(item,'Bestellnummer')      || getField(item, resolvedFields['Bestellnummer'])      || '';
  const lieferd   = getField(item,'Lieferdatum')        || getField(item, resolvedFields['Lieferdatum'])        || '';
  const tatPreis  = getField(item,'TatsaechlicherPreis')|| getField(item, resolvedFields['TatsaechlicherPreis'])|| '';

  return `
    ${renderApprovalCard(item)}

    <div class="detail-card">
      <div class="detail-card-header">Bestellung (Einkauf)
        ${isEinkauf ? `<button class="btn btn-sm btn-outline" id="btn-add-order">Bearbeiten</button>` : ''}
      </div>
      <div class="detail-card-body">
        ${orderNr  ? detailRow('Bestellnummer', orderNr) : '<p class="no-order">Noch keine Bestellnummer eingetragen.</p>'}
        ${lieferd  ? detailRow('Lieferdatum', fmtDate(lieferd)) : ''}
        ${tatPreis ? detailRow('Tatsächlicher Preis', fmtEuro(tatPreis)) : ''}
      </div>
    </div>`;
}

// Keywords to detect approval-related SP columns by displayName or internal name
const APPROVAL_RE   = /genehmig|freigab|entscheid|ablehn|kommentar.*genehm|genehm.*kommentar/i;
const STAGE_MAP = [
  { label: 'Einkauf',          re: /einkauf/i },
  { label: 'Verwaltung',       re: /verwaltung/i },
  { label: 'Geschäftsführung', re: /\bgf\b|geschäftsführ|geschaeftsfuehr/i },
];

function approvalStyle(val) {
  const v = (val || '').toLowerCase();
  if (/freigegeben|genehmigt|approved|ja\b/.test(v)) return { bg:'#f0fdf4', color:'#15803d', dot:'✓', cls:'ap-ok' };
  if (/abgelehnt|rejected|nein\b/.test(v))           return { bg:'#fef2f2', color:'#b91c1c', dot:'✗', cls:'ap-no' };
  return { bg:'#fffbeb', color:'#b45309', dot:'…', cls:'ap-pending' };
}

function renderApprovalCard(item) {
  const statusVal = getField(item,'Status') || 'Eingereicht';

  // Find all approval-related columns that have a value on this item
  const found = Object.entries(colByKey)
    .filter(([k, c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
    .map(([k, c]) => ({ key: k, label: c.displayName || k, val: getField(item, k) }))
    .filter(c => c.val !== null && c.val !== undefined && c.val !== '');

  // Group into stages; ungrouped = shown below without stage header
  const stages = STAGE_MAP.map(s => ({
    label: s.label,
    cols:  found.filter(c => s.re.test(c.label) || s.re.test(c.key)),
  })).filter(s => s.cols.length);
  const assigned = new Set(stages.flatMap(s => s.cols.map(c => c.key)));
  const ungrouped = found.filter(c => !assigned.has(c.key));

  const stagesHtml = stages.map(s => {
    // Pick the "decision" field (contains genehmig/entscheid/freigab/ablehn)
    const decisionCol = s.cols.find(c => /genehmig|entscheid|freigab|ablehn/i.test(c.label));
    const otherCols   = s.cols.filter(c => c !== decisionCol);
    const st = decisionCol ? approvalStyle(decisionCol.val) : { bg:'#f3f4f6', color:'#6b7280', dot:'○', cls:'ap-neutral' };
    return `
      <div class="approval-stage">
        <div class="ap-dot ${st.cls}">${st.dot}</div>
        <div class="ap-body">
          <div class="ap-stage-label">${esc(s.label)}</div>
          ${decisionCol ? `<span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(decisionCol.val))}</span>` : ''}
          ${otherCols.map(c => `<div class="ap-meta">${esc(c.label)}: ${esc(String(c.val))}</div>`).join('')}
        </div>
      </div>`;
  }).join('');

  const ungroupedHtml = ungrouped.map(c => {
    const st = approvalStyle(c.val);
    const isComment = /kommentar|ablehn|grund/i.test(c.label);
    if (isComment) return `<div class="ap-comment-box"><strong>${esc(c.label)}:</strong> ${esc(String(c.val))}</div>`;
    return `<div class="approval-stage">
      <div class="ap-dot ${st.cls}">${st.dot}</div>
      <div class="ap-body">
        <div class="ap-stage-label">${esc(c.label)}</div>
        <span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(c.val))}</span>
      </div>
    </div>`;
  }).join('');

  const noData = !stages.length && !ungrouped.length;

  return `
    <div class="detail-card">
      <div class="detail-card-header">Status &amp; Genehmigung</div>
      <div class="detail-card-body">
        <div class="ap-current-status">${statusBadge(statusVal)}</div>
        ${noData
          ? `<p class="ap-empty">Noch keine Genehmigungsdaten — Power Automate aktualisiert diesen Bereich automatisch.</p>`
          : `<div class="approval-stages">${stagesHtml}${ungroupedHtml}</div>`
        }
      </div>
    </div>`;
}

// ── EINKAUF ORDER MODAL ───────────────────────────────────────────────────────
function openOrderModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  const orderNr  = getField(item,'Bestellnummer')      || getField(item, resolvedFields['Bestellnummer'])      || '';
  const lieferd  = getField(item,'Lieferdatum')        || getField(item, resolvedFields['Lieferdatum'])        || '';
  const tatPreis = getField(item,'TatsaechlicherPreis')|| getField(item, resolvedFields['TatsaechlicherPreis'])|| '';

  $id('modal-title').textContent = 'Einkauf-Daten eintragen';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellnummer</label>
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. PO-2025-1234"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferdatum</label>
      <input type="date" id="m-delivery" value="${lieferd ? lieferd.slice(0,10) : ''}"/>
    </div>
    <div class="form-group">
      <label>Tatsächlicher Preis netto (€)</label>
      <input type="number" id="m-price" value="${tatPreis}" min="0" step="0.01" placeholder="0,00"/>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveOrderData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

async function saveOrderData(itemId) {
  const orderNr  = $id('m-ordernr').value.trim();
  const delivery = $id('m-delivery').value;
  const price    = $id('m-price').value;

  // Build patch using resolved names; fall back to raw key if column was never resolved
  const patch = {};
  const add = (key, val) => {
    if (!val) return;
    const col = resolvedFields[key] || key;
    if (col) patch[col] = val;
  };
  add('Bestellnummer', orderNr);
  if (delivery) add('Lieferdatum', toSpDate(delivery, resolvedFields['Lieferdatum'] || 'Lieferdatum'));
  if (price) add('TatsaechlicherPreis', parseFloat(price));

  if (!Object.keys(patch).length) { closeModal(); return; }

  // Retry PATCH removing unrecognized fields
  const skipped = [];
  for (let i = 0; i < 10; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Einkauf-Daten gespeichert ✓', 'success');
      closeModal();
      await loadItems(false);
      renderDetail(itemId);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); return; }
      skipped.push(m[1]);
      delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele nicht erkannte Felder.', 'error');
}

// ── WIZARD ────────────────────────────────────────────────────────────────────
function initWizard() {
  wizardData = {};
  showStep(1);
  // Show field availability hints
  for (const fd of FORM_FIELDS) {
    const hint = $id('hint-' + fd.key);
    if (!hint) continue;
    if (resolvedFields[fd.key]) {
      hint.textContent = '';
      hint.className = 'field-hint';
    } else if (fd.key !== 'Title') {
      hint.textContent = '⚠ Spalte „' + fd.label + '" existiert noch nicht in der SP-Liste – wird beim Speichern übersprungen.';
      hint.className = 'field-hint warn';
    }
  }
  // Reset fields
  ['Title','Beschreibung','Warengruppe','Prioritaet','Menge','Mengeneinheit',
   'Mindestlagermenge','Termin','Artikelnummer','Lieferant','GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  document.querySelector('input[name=Beschaffungslogik][value="Bestandsmaterial (bestandsgeführt)"]').checked = true;
  $id('preis-route-hint').style.display = 'none';
}

function showStep(n) {
  [1,2,3,4].forEach(i => {
    $id('wstep-'+i).classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-step="${i}"]`);
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function wNext(step) {
  if (step === 1) {
    const title = $id('f-Title').value.trim();
    const wg    = $id('f-Warengruppe').value;
    if (!title) { toast('Bitte Bezeichnung angeben.', 'error'); return; }
    if (!wg)    { toast('Bitte Warengruppe wählen.', 'error'); return; }
    wizardData.step1 = {
      Title:        title,
      Beschreibung: $id('f-Beschreibung').value.trim(),
      Warengruppe:  wg,
      Prioritaet:   $id('f-Prioritaet').value,
    };
  } else if (step === 2) {
    const menge = $id('f-Menge').value;
    const me    = $id('f-Mengeneinheit').value;
    if (!menge || parseFloat(menge) <= 0) { toast('Bitte gültige Menge eingeben.', 'error'); return; }
    if (!me)  { toast('Bitte Mengeneinheit wählen.', 'error'); return; }
    wizardData.step2 = {
      Menge:             menge,
      Mengeneinheit:     me,
      Mindestlagermenge: $id('f-Mindestlagermenge').value || null,
      Termin:            $id('f-Termin').value || null,
    };
  } else if (step === 3) {
    wizardData.step3 = {
      Beschaffungslogik: document.querySelector('input[name=Beschaffungslogik]:checked')?.value || '',
      Artikelnummer:     $id('f-Artikelnummer').value.trim(),
      Lieferant:         $id('f-Lieferant').value.trim(),
      GeschaetzterPreis: $id('f-GeschaetzterPreis').value ? parseFloat($id('f-GeschaetzterPreis').value) : null,
      Kostenstelle:      $id('f-Kostenstelle').value.trim(),
    };
    buildReview();
  }
  showStep(step + 1);
}

function wBack(step) { showStep(step - 1); }

function updatePreisHint() {
  const preis  = parseFloat($id('f-GeschaetzterPreis').value) || 0;
  const menge  = parseFloat($id('f-Menge')?.value) || 1;
  const gesamt = preis * menge;
  const hint   = $id('preis-route-hint');
  if (gesamt > 0) {
    hint.textContent = genehmigungsweg(gesamt);
    hint.style.display = 'block';
  } else {
    hint.style.display = 'none';
  }
}

function genehmigungsweg(gesamt) {
  if (gesamt < 250)   return `Geschätztes Volumen: ${fmtEuro(gesamt)} – kein Angebot nötig`;
  if (gesamt < 750)   return `Geschätztes Volumen: ${fmtEuro(gesamt)} – mind. 2 Angebote (Einkauf)`;
  if (gesamt < 10000) return `Geschätztes Volumen: ${fmtEuro(gesamt)} – Freigabe: Einkauf + Verwaltung`;
  return `Geschätztes Volumen: ${fmtEuro(gesamt)} – Freigabe: Einkauf + Verwaltung + Geschäftsführung`;
}

function buildReview() {
  const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
  const missingCols = FORM_FIELDS
    .filter(fd => fd.key !== 'Title' && !resolvedFields[fd.key] && d[fd.key])
    .map(fd => fd.label);

  $id('review-content').innerHTML = `
    <div class="review-grid">
      ${reviewSection('Bedarf', [
        ['Bezeichnung', d.Title],
        ['Beschreibung', d.Beschreibung],
        ['Warengruppe', d.Warengruppe],
        ['Priorität', d.Prioritaet],
      ])}
      ${reviewSection('Mengengaben', [
        ['Menge', d.Menge && d.Mengeneinheit ? `${d.Menge} ${d.Mengeneinheit}` : d.Menge],
        ['Mindestlagermenge', d.Mindestlagermenge],
        ['Benötigt bis', d.Termin ? fmtDate(d.Termin) : null],
      ])}
      ${reviewSection('Beschaffungsdetails', [
        ['Beschaffungslogik', d.Beschaffungslogik],
        ['Artikelnummer', d.Artikelnummer],
        ['Lieferant', d.Lieferant],
        ['Gesch. Preis (netto)', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
        ['Kostenstelle', d.Kostenstelle],
      ])}
    </div>
    ${d.GeschaetzterPreis ? `<div class="info-box info" style="margin-top:12px">${genehmigungsweg(d.GeschaetzterPreis * (d.Menge||1))}</div>` : ''}
    <div class="info-box info" style="margin-top:10px">
      Nach dem Einreichen wird <strong>Power Automate</strong> automatisch den Genehmigungsprozess starten und den zuständigen Genehmiger per E-Mail / Teams benachrichtigen.
    </div>`;

  const mw = $id('missing-cols-warn');
  if (missingCols.length) {
    mw.innerHTML = `⚠ Folgende Felder werden <strong>nicht gespeichert</strong> (Spalten fehlen in SP-Liste): <em>${missingCols.join(', ')}</em>`;
    mw.style.display = 'block';
  } else {
    mw.style.display = 'none';
  }
}

function reviewSection(title, rows) {
  const content = rows.filter(([,v]) => v).map(([l,v]) =>
    `<div class="review-row"><span class="review-label">${esc(l)}</span><span class="review-value">${esc(String(v))}</span></div>`
  ).join('');
  if (!content) return '';
  return `<div class="review-section"><h3>${title}</h3>${content}</div>`;
}

// ── SUBMIT ───────────────────────────────────────────────────────────────────
const NUMBER_FIELDS = new Set(['GeschaetzterPreis','Menge','Mindestlagermenge','TatsaechlicherPreis']);
const DATE_FIELDS   = new Set(['Termin','Lieferdatum']);

// SP Graph requires ISO 8601 for DateTime columns: "2026-05-10T00:00:00Z"
// "Date Only" columns want just the date: "2026-05-10"
// <input type="date"> yields "2026-05-10" → append time only if column is dateTime
function toSpDate(val, spColName) {
  if (!val) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(val)) return val; // already has time component
  const col = spColName ? colByKey[spColName] : null;
  const isDateOnly = col?.dateTime?.format === 'dateOnly';
  return isDateOnly ? val : val + 'T00:00:00Z';
}

function buildFields(data, fieldDefs) {
  const fields = {};
  for (const fd of fieldDefs) {
    const spCol = resolvedFields[fd.key];
    const val   = data[fd.key];
    if (!spCol || val === null || val === undefined || val === '') continue;
    if (DATE_FIELDS.has(fd.key)) {
      const d = toSpDate(val, spCol);
      if (d) fields[spCol] = d;
    } else if (NUMBER_FIELDS.has(fd.key)) {
      const col = colByKey[spCol];
      const n   = parseFloat(val);
      if (!isNaN(n)) {
        // Send as number only if SP column is actually numeric; otherwise string
        fields[spCol] = col?.number ? n : String(n);
      }
    } else {
      fields[spCol] = val;
    }
  }
  return fields;
}

// Retry POST removing offending fields until SP accepts the payload.
// Handles 400 "not recognized", generic 400 "invalidRequest", and 500 "generalException".
async function postRetry(path, fields) {
  const skipped = [];
  // Drop order: optional fields first, then non-Title required fields
  const optionalKeys = Object.entries(resolvedFields)
    .filter(([k]) => !FORM_FIELDS.find(f => f.key === k && f.required))
    .map(([, v]) => v).filter(Boolean);
  const requiredKeys = Object.entries(resolvedFields)
    .filter(([k]) => FORM_FIELDS.find(f => f.key === k && f.required && f.key !== 'Title'))
    .map(([, v]) => v).filter(Boolean);
  const dropQueue = [...optionalKeys, ...requiredKeys];

  for (let i = 0; i < dropQueue.length + 5; i++) {
    try {
      await gPost(path, { fields });
      return skipped;
    } catch(e) {
      // 400 with named field: SP says exactly which field is wrong
      const m400 = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (m400) {
        const bad = m400[1];
        skipped.push(bad);
        delete fields[bad];
        for (const [k, v] of Object.entries(resolvedFields)) {
          if (v === bad) resolvedFields[k] = null;
        }
        continue;
      }
      // 500 or generic 400 "invalidRequest": SP won't say which field is wrong
      // → drop the next field from the queue and retry
      const isRetryable = e.message.includes('500') || e.message.includes('generalException') ||
        e.message.includes('invalidRequest');
      if (isRetryable) {
        const dropKey = dropQueue.find(k => k && fields[k] !== undefined);
        if (dropKey) {
          const code = e.message.includes('400') ? '400' : '500';
          console.warn(`[postRetry] ${code} – dropping field:`, dropKey, 'value:', fields[dropKey]);
          skipped.push(dropKey + `(${code})`);
          delete fields[dropKey];
          continue;
        }
      }
      throw e;
    }
  }
  throw new Error('Einreichen fehlgeschlagen – zu viele problematische Felder.');
}

async function submitRequest() {
  const btn = $id('btn-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';

  try {
    const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
    const fields = buildFields(d, FORM_FIELDS);

    if (!fields['Title']) { toast('Titel fehlt.', 'error'); btn.disabled=false; btn.textContent='✓ Anfrage einreichen'; return; }

    const skipped = await postRetry(`/sites/${siteId}/lists/${listId}/items`, fields);

    if (skipped.length) {
      // Map SP column names back to human-readable field labels
      const labels = skipped.map(s => {
        const colName = s.replace(/\(\d+\)$/, '');
        const fd = FORM_FIELDS.find(f => resolvedFields[f.key] === colName || f.key === colName);
        return fd ? fd.label : colName;
      });
      toast(`Eingereicht. Folgende Felder wurden von SharePoint abgelehnt und nicht gespeichert: ${labels.join(', ')}`, 'info');
    } else {
      toast('Anfrage eingereicht! Power Automate startet den Genehmigungsprozess.', 'success');
    }
    await loadItems(false);
    navigate('mine');
  } catch(e) {
    toast('Fehler: ' + e.message, 'error');
    btn.disabled = false;
    btn.textContent = '✓ Anfrage einreichen';
  }
}

// ── CARD / ROW TEMPLATES ─────────────────────────────────────────────────────
function itemCard(item) {
  const title   = esc(getField(item,'Title') || '–');
  const status  = getField(item,'Status') || '';
  const wg      = getField(item, resolvedFields['Warengruppe']||'Warengruppe') || '';
  const preis   = parseFloat(getField(item, resolvedFields['GeschaetzterPreis']||'GeschaetzterPreis')) || null;
  const prio    = getField(item, resolvedFields['Prioritaet']||'Prioritaet') || '';
  const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
  const creator = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '';

  return `
    <div class="item-card" onclick="navigate('detail','${item.id}')">
      <div class="ic-left">
        <div class="ic-title">${prioDot(prio)}${title}</div>
        <div class="ic-meta">
          ${wg ? `<span>📦 ${esc(wg)}</span>` : ''}
          ${creator ? `<span>👤 ${esc(creator)}</span>` : ''}
          ${created ? `<span>📅 ${created}</span>` : ''}
        </div>
      </div>
      <div class="ic-right">
        ${preis ? `<div class="ic-price">${fmtEuro(preis)}</div>` : ''}
        ${statusBadge(status)}
      </div>
    </div>`;
}

function compactRow(item) {
  const title  = esc(getField(item,'Title') || '–');
  const status = getField(item,'Status') || '';
  const created = item.createdDateTime ? fmtRelative(item.createdDateTime) : '';
  return `
    <div class="compact-row" onclick="navigate('detail','${item.id}')">
      <div class="compact-title">${title}</div>
      <div class="compact-meta">ID ${item.id} · ${created}</div>
      ${statusBadge(status)}
    </div>`;
}

function emptyState(msg) {
  return `<div class="empty-state">
    <svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M5 4a3 3 0 00-3 3v6a3 3 0 003 3h10a3 3 0 003-3V7a3 3 0 00-3-3H5zm-1 9v-1h5v2H5a1 1 0 01-1-1zm7 1h4a1 1 0 001-1v-1h-5v2zm0-4h5V8h-5v2zM9 8H4v2h5V8z" clip-rule="evenodd"/></svg>
    <p>${msg}</p></div>`;
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
function closeModal() { $id('modal-overlay').classList.add('hidden'); }
document.addEventListener('DOMContentLoaded', () => {
  $id('modal-overlay')?.addEventListener('click', e => {
    if (e.target === $id('modal-overlay')) closeModal();
  });
});

// ── TOAST ─────────────────────────────────────────────────────────────────────
function toast(msg, type='info') {
  const icons = { success:'✅', error:'❌', info:'ℹ️' };
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.innerHTML = `<span>${icons[type]||'ℹ️'}</span><span>${esc(msg)}</span>`;
  $id('toast-c').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(()=>t.remove(),260); }, 4000);
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
const $id = id => document.getElementById(id);
const esc = s  => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function getField(item, key) {
  if (!key || !item?.fields) return null;
  return item.fields[key] ?? null;
}

function statusBadge(s) {
  if (!s) return '';
  const sl  = s.toLowerCase().trim();
  const st  = STATUS_STYLES[sl] || { bg:'#f3f4f6', color:'#374151' };
  return `<span class="status-badge" style="background:${st.bg};color:${st.color}">${esc(s)}</span>`;
}

function prioTag(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const color = p.toLowerCase() === 'dringend' ? '#b91c1c' : '#b45309';
  const bg    = p.toLowerCase() === 'dringend' ? '#fef2f2' : '#fffbeb';
  return `<span class="status-badge" style="background:${bg};color:${color}">${esc(p)}</span>`;
}

function prioDot(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const c = p.toLowerCase() === 'dringend' ? '#ef4444' : '#f59e0b';
  return `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${c};margin-right:5px;vertical-align:middle"></span>`;
}

function fmtEuro(v) {
  if (!v && v !== 0) return '';
  return new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(v);
}
function fmtDate(s) {
  if (!s) return '';
  return new Date(s).toLocaleDateString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric'});
}
function fmtDateTime(s) {
  if (!s) return '';
  return new Date(s).toLocaleString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
function fmtRelative(s) {
  const d = Date.now() - new Date(s).getTime();
  if (d < 60000)     return 'gerade eben';
  if (d < 3600000)   return Math.floor(d/60000) + ' Min.';
  if (d < 86400000)  return Math.floor(d/3600000) + ' Std.';
  if (d < 2592000000)return Math.floor(d/86400000) + ' Tage';
  return fmtDate(s);
}
