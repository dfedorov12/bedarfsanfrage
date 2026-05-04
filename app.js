'use strict';

// ── Data Store ─────────────────────────────────────────────────────────────
const STORE_KEY = 'dihag_bedarfsanfragen_v2';

function load() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY)) || []; } catch { return []; }
}
function save(data) {
  localStorage.setItem(STORE_KEY, JSON.stringify(data));
}

let requests = load();

// Seed demo data on first load
if (requests.length === 0) {
  const now = Date.now();
  const day = 86400000;
  requests = [
    {
      id: 'BA-0001', createdAt: now - 10 * day, createdBy: 'Anna Bauer',
      bezeichnung: 'Ergonomische Bürostühle', beschreibung: 'Neue Stühle für Büro 3. OG nach Umbau',
      warengruppe: 'Bürobedarf & Büroausstattung', kostenstelle: '4200',
      menge: 8, einheit: 'Stück', termin: fmtDate(now + 14 * day), prioritaet: 'hoch',
      materialtyp: 'neues_material', artikelnummer: '', lieferant: 'Interstuhl', spezifikation: 'Netzrücken, höhenverstellbar',
      preis: 450, gesamt: 3600, budget: 'ja', kostennotiz: '',
      status: 'in_pruefung',
      approvals: [
        { role: 'Einkauf', person: 'Denis Fedorov', decision: 'approved', comment: 'Angebot liegt vor', date: now - 5 * day },
        { role: 'Verwaltung', person: 'Klaus Meier', decision: 'pending', comment: '', date: null }
      ],
      bestellnummer: '', lieferdatum: ''
    },
    {
      id: 'BA-0002', createdAt: now - 3 * day, createdBy: 'Denis Fedorov',
      bezeichnung: 'Schrauben M8x50 verzinkt', beschreibung: 'Nachbestellung Lagermaterial',
      warengruppe: 'Rohstoffe & Materialien', kostenstelle: '3100',
      menge: 500, einheit: 'Stück', termin: fmtDate(now + 7 * day), prioritaet: 'normal',
      materialtyp: 'bestandsmaterial', artikelnummer: 'M-4711', lieferant: 'Würth GmbH', spezifikation: 'ISO 4017, galv. verzinkt',
      preis: 0.18, gesamt: 90, budget: 'ja', kostennotiz: '',
      status: 'freigegeben',
      approvals: [
        { role: 'Einkauf', person: 'Denis Fedorov', decision: 'approved', comment: '', date: now - 2 * day }
      ],
      bestellnummer: 'PO-2025-1042', lieferdatum: fmtDate(now + 5 * day)
    },
    {
      id: 'BA-0003', createdAt: now - 1 * day, createdBy: 'Maria Schmidt',
      bezeichnung: 'IT-Audit Dienstleistung', beschreibung: 'Externer IT-Security Audit gemäß ISO 27001',
      warengruppe: 'Dienstleistungen', kostenstelle: '5500',
      menge: 5, einheit: 'Stunden', termin: fmtDate(now + 30 * day), prioritaet: 'normal',
      materialtyp: 'dienstleistung', artikelnummer: '', lieferant: '', spezifikation: '',
      preis: 1800, gesamt: 9000, budget: 'nein', kostennotiz: 'Budgetfreigabe durch Controlling notwendig',
      status: 'eingereicht',
      approvals: [
        { role: 'Einkauf', person: 'Denis Fedorov', decision: 'pending', comment: '', date: null },
        { role: 'Verwaltung', person: 'Klaus Meier', decision: 'pending', comment: '', date: null },
        { role: 'Controlling', person: 'Sandra Koch', decision: 'pending', comment: '', date: null }
      ],
      bestellnummer: '', lieferdatum: ''
    },
    {
      id: 'BA-0004', createdAt: now - 6 * day, createdBy: 'Thomas Lang',
      bezeichnung: 'Gabelstapler gebraucht', beschreibung: 'Ersatz für defekten Gabelstapler Halle 2',
      warengruppe: 'Fahrzeuge & Transport', kostenstelle: '2100',
      menge: 1, einheit: 'Stück', termin: fmtDate(now + 45 * day), prioritaet: 'dringend',
      materialtyp: 'neues_material', artikelnummer: '', lieferant: 'Linde Material Handling', spezifikation: '2t Tragkraft, Elektroantrieb',
      preis: 18500, gesamt: 18500, budget: 'ja', kostennotiz: 'Investitionsplan 2025 Pos. 7',
      status: 'in_pruefung',
      approvals: [
        { role: 'Einkauf', person: 'Denis Fedorov', decision: 'approved', comment: 'Angebote eingeholt', date: now - 4 * day },
        { role: 'Verwaltung', person: 'Klaus Meier', decision: 'approved', comment: '', date: now - 3 * day },
        { role: 'Geschäftsführung', person: 'Dr. Fischer', decision: 'pending', comment: '', date: null }
      ],
      bestellnummer: '', lieferdatum: ''
    },
    {
      id: 'BA-0005', createdAt: now - 15 * day, createdBy: 'Denis Fedorov',
      bezeichnung: 'Druckerpapier A4 80g', beschreibung: 'Bürobedarf Quartalsnachbestellung',
      warengruppe: 'Bürobedarf & Büroausstattung', kostenstelle: '4200',
      menge: 20, einheit: 'Karton', termin: fmtDate(now - 5 * day), prioritaet: 'normal',
      materialtyp: 'bestandsmaterial', artikelnummer: 'B-0012', lieferant: 'Staples', spezifikation: '',
      preis: 35, gesamt: 700, budget: 'ja', kostennotiz: '',
      status: 'bestellt',
      approvals: [
        { role: 'Einkauf', person: 'Denis Fedorov', decision: 'approved', comment: '', date: now - 14 * day }
      ],
      bestellnummer: 'PO-2025-0987', lieferdatum: fmtDate(now - 8 * day)
    }
  ];
  save(requests);
}

// ── Current Role / User ─────────────────────────────────────────────────────
const ROLES = {
  beschaffer:       { label: 'Beschaffer',         name: 'Max Mustermann' },
  einkauf:          { label: 'Einkauf / Lead Buyer', name: 'Denis Fedorov' },
  verwaltung:       { label: 'Verwaltung',          name: 'Klaus Meier' },
  controlling:      { label: 'Controlling',         name: 'Sandra Koch' },
  geschaeftsfuehrung: { label: 'Geschäftsführung',  name: 'Dr. Fischer' }
};
let currentRole = 'einkauf';

function getCurrentUser() { return ROLES[currentRole]; }

// ── Routing ──────────────────────────────────────────────────────────────────
let currentView = 'dashboard';
let prevView = 'dashboard';

function navigate(view, id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const el = document.getElementById('view-' + view);
  if (!el) return;
  el.classList.add('active');

  const navEl = document.querySelector(`.nav-item[data-view="${view}"]`);
  if (navEl) navEl.classList.add('active');

  const titles = {
    dashboard: 'Dashboard',
    new: 'Neue Bedarfsanfrage',
    mine: 'Meine Anfragen',
    approvals: 'Zur Genehmigung',
    all: 'Alle Anfragen',
    detail: 'Anfrage Details'
  };
  document.getElementById('page-title').textContent = titles[view] || view;

  prevView = currentView;
  currentView = view;

  if (view === 'dashboard') renderDashboard();
  else if (view === 'mine') renderList('mine');
  else if (view === 'approvals') renderList('approvals');
  else if (view === 'all') renderList('all');
  else if (view === 'new') initWizard();
  else if (view === 'detail' && id) renderDetail(id);

  updateApprovalBadge();
}

// ── Event Wiring ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Nav clicks
  document.querySelectorAll('.nav-item[data-view]').forEach(a => {
    a.addEventListener('click', e => { e.preventDefault(); navigate(a.dataset.view); });
  });
  // Dashboard "Neu" button
  document.querySelectorAll('[data-view="new"]').forEach(b => {
    b.addEventListener('click', e => { e.preventDefault(); navigate('new'); });
  });
  // Role switcher
  document.getElementById('role-switcher').addEventListener('change', e => {
    currentRole = e.target.value;
    const u = getCurrentUser();
    document.getElementById('user-name-sidebar').textContent = u.name;
    document.getElementById('user-role-sidebar').textContent = u.label;
    document.getElementById('user-avatar-sidebar').textContent = u.name.split(' ').map(p => p[0]).join('').slice(0, 2).toUpperCase();
    navigate(currentView);
  });
  // Back button in detail
  document.getElementById('detail-back').addEventListener('click', () => navigate(prevView));
  // Menu toggle
  document.getElementById('menu-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('open');
  });

  navigate('dashboard');
});

// ── Dashboard ─────────────────────────────────────────────────────────────────
function renderDashboard() {
  const total = requests.length;
  const pending = requests.filter(r => ['eingereicht', 'in_pruefung'].includes(r.status)).length;
  const approved = requests.filter(r => r.status === 'freigegeben' || r.status === 'bestellt').length;
  const volume = requests.reduce((s, r) => s + (r.gesamt || 0), 0);

  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-pending').textContent = pending;
  document.getElementById('stat-approved').textContent = approved;
  document.getElementById('stat-volume').textContent = fmtEuro(volume);

  // Recent list
  const sorted = [...requests].sort((a, b) => b.createdAt - a.createdAt).slice(0, 6);
  const el = document.getElementById('dashboard-list');
  el.innerHTML = sorted.length ? sorted.map(r => `
    <div class="request-list-item" onclick="navigate('detail','${r.id}')">
      <div style="flex:1">
        <div class="rli-title">${r.bezeichnung}</div>
        <div class="rli-meta">${r.id} · ${r.createdBy} · ${fmtRelative(r.createdAt)}</div>
      </div>
      ${statusBadge(r.status)}
    </div>`).join('') : emptyState('Keine Anfragen vorhanden');

  // Pending approvals for current role
  const myApprovals = getMyPendingApprovals();
  const el2 = document.getElementById('dashboard-approvals');
  el2.innerHTML = myApprovals.length ? myApprovals.map(r => `
    <div class="request-list-item" onclick="navigate('detail','${r.id}')">
      <div style="flex:1">
        <div class="rli-title">${r.bezeichnung}</div>
        <div class="rli-meta">${r.id} · ${fmtEuro(r.gesamt)}</div>
      </div>
      <span class="status-badge status-in_pruefung">Warte auf Sie</span>
    </div>`).join('') : emptyState('Keine Genehmigungen ausstehend');
}

// ── List Views ────────────────────────────────────────────────────────────────
function renderList(type) {
  let items = [...requests].sort((a, b) => b.createdAt - a.createdAt);

  if (type === 'mine') {
    const user = getCurrentUser();
    items = items.filter(r => r.createdBy === user.name);
  } else if (type === 'approvals') {
    items = getMyPendingApprovals();
  }

  const container = document.getElementById('list-' + type);
  if (!container) return;

  // Populate warengruppe filter once
  if (type === 'all') {
    const wgSel = document.getElementById('filter-all-wg');
    if (wgSel.options.length <= 1) {
      const wgs = [...new Set(requests.map(r => r.warengruppe))].filter(Boolean).sort();
      wgs.forEach(wg => wgSel.add(new Option(wg, wg)));
    }
  }

  renderFilteredList(type, items, container);
}

function filterList(type) {
  let items = [...requests].sort((a, b) => b.createdAt - a.createdAt);
  if (type === 'mine') {
    const user = getCurrentUser();
    items = items.filter(r => r.createdBy === user.name);
  } else if (type === 'approvals') {
    items = getMyPendingApprovals();
  }
  const container = document.getElementById('list-' + type);
  renderFilteredList(type, items, container);
}

function renderFilteredList(type, items, container) {
  const search = (document.getElementById(`search-${type}`)?.value || '').toLowerCase();
  const status = document.getElementById(`filter-${type}-status`)?.value || '';
  const wg     = document.getElementById(`filter-all-wg`)?.value || '';

  let filtered = items;
  if (search) filtered = filtered.filter(r =>
    r.bezeichnung.toLowerCase().includes(search) ||
    r.id.toLowerCase().includes(search) ||
    r.createdBy.toLowerCase().includes(search)
  );
  if (status) filtered = filtered.filter(r => r.status === status);
  if (wg)     filtered = filtered.filter(r => r.warengruppe === wg);

  container.innerHTML = filtered.length
    ? filtered.map(requestCard).join('')
    : `<div class="empty-state"><svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M5 4a3 3 0 00-3 3v6a3 3 0 003 3h10a3 3 0 003-3V7a3 3 0 00-3-3H5zm-1 9v-1h5v2H5a1 1 0 01-1-1zm7 1h4a1 1 0 001-1v-1h-5v2zm0-4h5V8h-5v2zM9 8H4v2h5V8z" clip-rule="evenodd"/></svg>Keine Anfragen gefunden</div>`;
}

function requestCard(r) {
  const pendingCount = (r.approvals || []).filter(a => a.decision === 'pending').length;
  return `
    <div class="request-card" onclick="navigate('detail','${r.id}')">
      <div class="rc-left">
        <div class="rc-title">
          <span class="priority-dot ${r.prioritaet}"></span>${r.bezeichnung}
        </div>
        <div class="rc-sub">
          <span>📋 ${r.id}</span>
          <span>👤 ${r.createdBy}</span>
          <span>📦 ${r.warengruppe}</span>
          <span>🗓 ${fmtDate2(r.createdAt)}</span>
          ${pendingCount > 0 ? `<span>⏳ ${pendingCount} Genehmigung${pendingCount > 1 ? 'en' : ''} ausstehend</span>` : ''}
        </div>
      </div>
      <div class="rc-right">
        <div class="rc-amount">${fmtEuro(r.gesamt)}</div>
        ${statusBadge(r.status)}
      </div>
    </div>`;
}

// ── Detail View ───────────────────────────────────────────────────────────────
function renderDetail(id) {
  const r = requests.find(x => x.id === id);
  if (!r) return;

  const approvalRoute = getApprovalRoute(r.gesamt);
  const canApprove = canCurrentRoleApprove(r);
  const myApproval = canApprove ? getCurrentApproval(r) : null;

  document.getElementById('detail-content').innerHTML = `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:18px;flex-wrap:wrap;">
      <h2 style="font-size:1.2rem;font-weight:700;flex:1">${r.bezeichnung}</h2>
      ${statusBadge(r.status)}
      <span class="status-badge" style="background:#f3f4f6;color:#374151">${r.id}</span>
      ${r.prioritaet !== 'normal' ? `<span class="status-badge ${r.prioritaet === 'dringend' ? 'status-abgelehnt' : 'status-in_pruefung'}">${r.prioritaet === 'dringend' ? '🔴 Dringend' : '🟡 Hoch'}</span>` : ''}
    </div>

    <div class="detail-grid">
      <div>
        <!-- Grunddaten -->
        <div class="detail-card">
          <div class="detail-card-header">Grunddaten</div>
          <div class="detail-card-body">
            <div class="detail-row"><span class="detail-label">Bezeichnung</span><span class="detail-value">${r.bezeichnung}</span></div>
            ${r.beschreibung ? `<div class="detail-row"><span class="detail-label">Beschreibung</span><span class="detail-value">${r.beschreibung}</span></div>` : ''}
            <div class="detail-row"><span class="detail-label">Warengruppe</span><span class="detail-value">${r.warengruppe}</span></div>
            <div class="detail-row"><span class="detail-label">Kostenstelle</span><span class="detail-value">${r.kostenstelle || '–'}</span></div>
            <div class="detail-row"><span class="detail-label">Menge</span><span class="detail-value">${r.menge} ${r.einheit}</span></div>
            ${r.termin ? `<div class="detail-row"><span class="detail-label">Benötigt bis</span><span class="detail-value">${r.termin}</span></div>` : ''}
            <div class="detail-row"><span class="detail-label">Erstellt von</span><span class="detail-value">${r.createdBy} am ${fmtDate2(r.createdAt)}</span></div>
          </div>
        </div>

        <!-- Material -->
        <div class="detail-card">
          <div class="detail-card-header">Materialinformationen</div>
          <div class="detail-card-body">
            <div class="detail-row"><span class="detail-label">Materialtyp</span><span class="detail-value">${mtLabel(r.materialtyp)}</span></div>
            ${r.artikelnummer ? `<div class="detail-row"><span class="detail-label">Artikelnummer</span><span class="detail-value">${r.artikelnummer}</span></div>` : ''}
            ${r.lieferant ? `<div class="detail-row"><span class="detail-label">Lieferant</span><span class="detail-value">${r.lieferant}</span></div>` : ''}
            ${r.spezifikation ? `<div class="detail-row"><span class="detail-label">Spezifikation</span><span class="detail-value">${r.spezifikation}</span></div>` : ''}
          </div>
        </div>

        <!-- Kosten -->
        <div class="detail-card">
          <div class="detail-card-header">Kosten & Budget</div>
          <div class="detail-card-body">
            <div class="detail-row"><span class="detail-label">Einzelpreis (netto)</span><span class="detail-value">${fmtEuro(r.preis)}</span></div>
            <div class="detail-row"><span class="detail-label">Gesamtpreis (netto)</span><span class="detail-value" style="font-weight:700;font-size:1rem">${fmtEuro(r.gesamt)}</span></div>
            <div class="detail-row"><span class="detail-label">Budget vorhanden</span><span class="detail-value">${r.budget === 'ja' ? '✅ Ja' : r.budget === 'nein' ? '❌ Nein' : '❓ Unbekannt'}</span></div>
            <div class="detail-row"><span class="detail-label">Genehmigungsweg</span><span class="detail-value">${approvalRoute.label}</span></div>
            ${r.kostennotiz ? `<div class="detail-row"><span class="detail-label">Notiz</span><span class="detail-value">${r.kostennotiz}</span></div>` : ''}
          </div>
        </div>

        <!-- Bestellung -->
        ${r.bestellnummer ? `
        <div class="detail-card">
          <div class="detail-card-header">Bestellinformationen</div>
          <div class="detail-card-body">
            <div class="detail-row"><span class="detail-label">Bestellnummer</span><span class="detail-value" style="font-weight:700">${r.bestellnummer}</span></div>
            ${r.lieferdatum ? `<div class="detail-row"><span class="detail-label">Lieferdatum</span><span class="detail-value">${r.lieferdatum}</span></div>` : ''}
          </div>
        </div>` : ''}

        <!-- Approval actions -->
        ${canApprove && myApproval && myApproval.decision === 'pending' ? `
        <div class="detail-card" style="border-color:#c7d2fe">
          <div class="detail-card-header" style="background:#eff3ff;color:#1a56db">Ihre Genehmigung ist erforderlich</div>
          <div class="detail-card-body">
            <p style="font-size:.88rem;color:#374151;margin-bottom:12px">Diese Anfrage wartet auf Ihre Entscheidung als <strong>${getCurrentUser().label}</strong>.</p>
            <div class="approval-actions">
              <button class="btn btn-success" onclick="openApprovalModal('${r.id}','approved')">✓ Freigeben</button>
              <button class="btn btn-danger" onclick="openApprovalModal('${r.id}','rejected')">✗ Ablehnen</button>
              ${r.status === 'freigegeben' && currentRole === 'einkauf' ? `<button class="btn btn-outline" onclick="openOrderModal('${r.id}')">📋 Bestellung anlegen</button>` : ''}
            </div>
          </div>
        </div>` : ''}

        ${r.status === 'freigegeben' && currentRole === 'einkauf' && !r.bestellnummer ? `
        <div class="detail-card" style="border-color:#bbf7d0">
          <div class="detail-card-header" style="background:#f0fdf4;color:#16a34a">Bereit zur Bestellung</div>
          <div class="detail-card-body">
            <div class="approval-actions">
              <button class="btn btn-primary" onclick="openOrderModal('${r.id}')">📋 Bestellnummer eintragen</button>
            </div>
          </div>
        </div>` : ''}
      </div>

      <!-- Right col: Timeline -->
      <div>
        <div class="detail-card">
          <div class="detail-card-header">Genehmigungsprozess</div>
          <div class="detail-card-body">
            <div class="timeline">
              <div class="timeline-item">
                <div class="tl-dot approved">✓</div>
                <div class="tl-body">
                  <div class="tl-title">Anfrage erstellt</div>
                  <div class="tl-sub">${r.createdBy} · ${fmtDate2(r.createdAt)}</div>
                </div>
              </div>
              ${(r.approvals || []).map(a => timelineItem(a)).join('')}
              ${r.bestellnummer ? `
              <div class="timeline-item">
                <div class="tl-dot approved">📦</div>
                <div class="tl-body">
                  <div class="tl-title">Bestellung aufgegeben</div>
                  <div class="tl-sub">${r.bestellnummer}</div>
                </div>
              </div>` : ''}
            </div>
          </div>
        </div>
      </div>
    </div>`;
}

function timelineItem(a) {
  const dotClass = a.decision === 'approved' ? 'approved' : a.decision === 'rejected' ? 'rejected' : 'pending';
  const icon = a.decision === 'approved' ? '✓' : a.decision === 'rejected' ? '✗' : '…';
  return `
    <div class="timeline-item">
      <div class="tl-dot ${dotClass}">${icon}</div>
      <div class="tl-body">
        <div class="tl-title">${a.role}${a.person ? ` – ${a.person}` : ''}</div>
        <div class="tl-sub">${a.decision === 'pending' ? 'Ausstehend' : fmtDate2(a.date)}</div>
        ${a.comment ? `<div class="tl-comment">${a.comment}</div>` : ''}
      </div>
    </div>`;
}

// ── Wizard ────────────────────────────────────────────────────────────────────
let wizardData = {};

function initWizard() {
  wizardData = {};
  showWizardStep(1);
  document.querySelectorAll('.wstep').forEach(s => {
    s.classList.remove('active', 'done');
    if (s.dataset.step === '1') s.classList.add('active');
  });
  // Reset form fields
  ['f-bezeichnung','f-beschreibung','f-warengruppe','f-kostenstelle','f-menge','f-termin',
   'f-artikelnummer','f-lieferant','f-spezifikation','f-preis','f-gesamt','f-kostennotiz'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.querySelector('input[name=materialtyp][value=bestandsmaterial]').checked = true;
  document.querySelector('input[name=budget][value=ja]').checked = true;
  document.getElementById('approval-route-info').classList.add('hidden');
}

function showWizardStep(n) {
  [1,2,3,4].forEach(i => {
    document.getElementById('step-' + i).classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-step="${i}"]`);
    s.classList.remove('active', 'done');
    if (i < n) s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function wizardNext(step) {
  if (step === 1) {
    const bz = document.getElementById('f-bezeichnung').value.trim();
    const wg = document.getElementById('f-warengruppe').value;
    const mg = document.getElementById('f-menge').value;
    if (!bz || !wg || !mg) { showToast('Bitte alle Pflichtfelder ausfüllen.', 'error'); return; }
    wizardData.step1 = {
      bezeichnung: bz,
      beschreibung: document.getElementById('f-beschreibung').value.trim(),
      warengruppe: wg,
      kostenstelle: document.getElementById('f-kostenstelle').value.trim(),
      menge: parseFloat(mg),
      einheit: document.getElementById('f-einheit').value,
      termin: document.getElementById('f-termin').value,
      prioritaet: document.getElementById('f-prioritaet').value
    };
  } else if (step === 2) {
    wizardData.step2 = {
      materialtyp: document.querySelector('input[name=materialtyp]:checked').value,
      artikelnummer: document.getElementById('f-artikelnummer').value.trim(),
      lieferant: document.getElementById('f-lieferant').value.trim(),
      spezifikation: document.getElementById('f-spezifikation').value.trim()
    };
  } else if (step === 3) {
    const preis = parseFloat(document.getElementById('f-preis').value);
    if (!preis || preis <= 0) { showToast('Bitte einen Preis angeben.', 'error'); return; }
    wizardData.step3 = {
      preis,
      gesamt: preis * (wizardData.step1?.menge || 1),
      budget: document.querySelector('input[name=budget]:checked').value,
      kostennotiz: document.getElementById('f-kostennotiz').value.trim()
    };
    buildReviewPage();
  }
  showWizardStep(step + 1);
}

function wizardBack(step) {
  showWizardStep(step - 1);
}

function updateGesamtpreis() {
  const p = parseFloat(document.getElementById('f-preis').value) || 0;
  const m = parseFloat(document.getElementById('f-menge')?.value) || 1;
  const g = p * m;
  document.getElementById('f-gesamt').value = g > 0 ? g.toFixed(2) : '';

  const info = document.getElementById('approval-route-info');
  if (g > 0) {
    const route = getApprovalRoute(g);
    info.className = 'info-box info';
    info.innerHTML = `<strong>Genehmigungsweg:</strong> ${route.label}<br><small>${route.desc}</small>`;
    info.classList.remove('hidden');
  } else {
    info.classList.add('hidden');
  }
}

function buildReviewPage() {
  const d1 = wizardData.step1 || {};
  const d2 = wizardData.step2 || {};
  const d3 = wizardData.step3 || {};
  const route = getApprovalRoute(d3.gesamt || 0);

  document.getElementById('review-content').innerHTML = `
    <div class="review-section">
      <h3>Grunddaten</h3>
      ${reviewRow('Bezeichnung', d1.bezeichnung)}
      ${d1.beschreibung ? reviewRow('Beschreibung', d1.beschreibung) : ''}
      ${reviewRow('Warengruppe', d1.warengruppe)}
      ${reviewRow('Menge', d1.menge + ' ' + d1.einheit)}
      ${d1.kostenstelle ? reviewRow('Kostenstelle', d1.kostenstelle) : ''}
      ${d1.termin ? reviewRow('Benötigt bis', d1.termin) : ''}
      ${reviewRow('Priorität', d1.prioritaet)}
    </div>
    <div class="review-section">
      <h3>Material</h3>
      ${reviewRow('Typ', mtLabel(d2.materialtyp))}
      ${d2.artikelnummer ? reviewRow('Artikelnummer', d2.artikelnummer) : ''}
      ${d2.lieferant ? reviewRow('Lieferant', d2.lieferant) : ''}
    </div>
    <div class="review-section">
      <h3>Kosten</h3>
      ${reviewRow('Einzelpreis', fmtEuro(d3.preis))}
      ${reviewRow('Gesamtpreis', fmtEuro(d3.gesamt))}
      ${reviewRow('Budget', d3.budget === 'ja' ? 'Vorhanden' : d3.budget === 'nein' ? 'Nicht vorhanden' : 'Unbekannt')}
    </div>
    <div class="approval-route-card">
      <div class="route-title">Genehmigungsweg</div>
      <div class="route-steps">
        ${route.steps.map(s => `<div class="route-step">${s}</div>`).join('<div class="route-arrow">→</div>')}
      </div>
      <div style="margin-top:8px;font-size:.8rem;color:#3730a3">${route.desc}</div>
    </div>`;
}

function reviewRow(label, value) {
  return `<div class="review-row"><span class="review-label">${label}</span><span class="review-value">${value}</span></div>`;
}

function submitRequest() {
  const d1 = wizardData.step1 || {};
  const d2 = wizardData.step2 || {};
  const d3 = wizardData.step3 || {};
  const user = getCurrentUser();
  const route = getApprovalRoute(d3.gesamt || 0);

  const id = 'BA-' + String(requests.length + 1).padStart(4, '0');
  const newReq = {
    id, createdAt: Date.now(), createdBy: user.name,
    ...d1, ...d2, ...d3,
    status: 'eingereicht',
    approvals: route.steps.map(role => ({ role, person: '', decision: 'pending', comment: '', date: null })),
    bestellnummer: '', lieferdatum: ''
  };

  requests.unshift(newReq);
  save(requests);
  showToast('Anfrage ' + id + ' wurde eingereicht!', 'success');
  navigate('mine');
}

// ── Approval Logic ────────────────────────────────────────────────────────────
function getApprovalRoute(amount) {
  if (amount < 750) return {
    label: 'Einkauf (automatisch)',
    steps: ['Einkauf'],
    desc: 'Bestellvolumen < 750 € – Freigabe durch Einkauf'
  };
  if (amount < 10000) return {
    label: 'Einkauf → Verwaltung',
    steps: ['Einkauf', 'Verwaltung'],
    desc: 'Bestellvolumen 750 € – 10.000 € – Freigabe durch Einkauf und Verwaltung'
  };
  return {
    label: 'Einkauf → Verwaltung → Geschäftsführung',
    steps: ['Einkauf', 'Verwaltung', 'Geschäftsführung'],
    desc: 'Bestellvolumen > 10.000 € – mehrstufige Freigabe erforderlich'
  };
}

function getMyPendingApprovals() {
  const roleLabel = getCurrentUser().label.split(' /')[0];
  return requests.filter(r => {
    const myStep = (r.approvals || []).find(a => {
      if (currentRole === 'einkauf') return a.role === 'Einkauf';
      if (currentRole === 'verwaltung') return a.role === 'Verwaltung';
      if (currentRole === 'controlling') return a.role === 'Controlling';
      if (currentRole === 'geschaeftsfuehrung') return a.role === 'Geschäftsführung';
      return false;
    });
    if (!myStep || myStep.decision !== 'pending') return false;
    // previous steps must be approved
    const idx = (r.approvals || []).indexOf(myStep);
    return idx === 0 || (r.approvals || []).slice(0, idx).every(a => a.decision === 'approved');
  });
}

function canCurrentRoleApprove(r) {
  return getMyPendingApprovals().some(x => x.id === r.id);
}

function getCurrentApproval(r) {
  return (r.approvals || []).find(a => {
    if (currentRole === 'einkauf') return a.role === 'Einkauf';
    if (currentRole === 'verwaltung') return a.role === 'Verwaltung';
    if (currentRole === 'controlling') return a.role === 'Controlling';
    if (currentRole === 'geschaeftsfuehrung') return a.role === 'Geschäftsführung';
    return false;
  });
}

function updateApprovalBadge() {
  const count = getMyPendingApprovals().length;
  const badge = document.getElementById('approval-badge');
  badge.textContent = count;
  badge.classList.toggle('visible', count > 0);
}

// ── Modals ────────────────────────────────────────────────────────────────────
function openApprovalModal(id, decision) {
  const r = requests.find(x => x.id === id);
  const label = decision === 'approved' ? 'Freigabe' : 'Ablehnung';
  const btnClass = decision === 'approved' ? 'btn-success' : 'btn-danger';

  document.getElementById('modal-title').textContent = label + ': ' + r.bezeichnung;
  document.getElementById('modal-body').innerHTML = `
    <p style="margin-bottom:12px;color:#374151">Möchten Sie diese Anfrage ${decision === 'approved' ? '<strong>freigeben</strong>' : '<strong>ablehnen</strong>'}?</p>
    <div class="form-group">
      <label>Kommentar ${decision === 'rejected' ? '<span class="req">*</span>' : '(optional)'}</label>
      <textarea id="approval-comment" rows="3" placeholder="Begründung …"></textarea>
    </div>`;
  document.getElementById('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn ${btnClass}" onclick="confirmApproval('${id}','${decision}')">${label} bestätigen</button>`;

  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmApproval(id, decision) {
  const comment = document.getElementById('approval-comment').value.trim();
  if (decision === 'rejected' && !comment) { showToast('Bitte eine Begründung angeben.', 'error'); return; }

  const r = requests.find(x => x.id === id);
  const approval = getCurrentApproval(r);
  if (!approval) return;

  approval.decision = decision;
  approval.comment = comment;
  approval.date = Date.now();
  approval.person = getCurrentUser().name;

  if (decision === 'rejected') {
    r.status = 'abgelehnt';
  } else {
    const allDone = (r.approvals || []).every(a => a.decision === 'approved');
    r.status = allDone ? 'freigegeben' : 'in_pruefung';
  }

  save(requests);
  closeModal();
  showToast(decision === 'approved' ? 'Anfrage freigegeben ✓' : 'Anfrage abgelehnt', decision === 'approved' ? 'success' : 'error');
  renderDetail(id);
  updateApprovalBadge();
}

function openOrderModal(id) {
  const r = requests.find(x => x.id === id);
  document.getElementById('modal-title').textContent = 'Bestellnummer eintragen';
  document.getElementById('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellnummer <span class="req">*</span></label>
      <input type="text" id="order-number" value="${r.bestellnummer || ''}" placeholder="z. B. PO-2025-1234" />
    </div>
    <div class="form-group">
      <label>Voraussichtliches Lieferdatum</label>
      <input type="date" id="order-delivery" value="${r.lieferdatum || ''}" />
    </div>`;
  document.getElementById('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveOrder('${id}')">Speichern</button>`;
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function saveOrder(id) {
  const nr = document.getElementById('order-number').value.trim();
  if (!nr) { showToast('Bitte Bestellnummer angeben.', 'error'); return; }
  const r = requests.find(x => x.id === id);
  r.bestellnummer = nr;
  r.lieferdatum = document.getElementById('order-delivery').value;
  r.status = 'bestellt';
  save(requests);
  closeModal();
  showToast('Bestellnummer gespeichert ✓', 'success');
  renderDetail(id);
}

function closeModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}
document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay')) closeModal();
});

// ── Toast ─────────────────────────────────────────────────────────────────────
function showToast(msg, type = 'info') {
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.textContent = msg;
  document.getElementById('toast-container').appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtEuro(v) {
  if (!v && v !== 0) return '–';
  return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(v);
}
function fmtDate(ts) {
  return new Date(ts).toISOString().slice(0, 10);
}
function fmtDate2(ts) {
  if (!ts) return '–';
  return new Date(ts).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
function fmtRelative(ts) {
  const d = Date.now() - ts;
  if (d < 60000) return 'gerade eben';
  if (d < 3600000) return Math.floor(d / 60000) + ' Min.';
  if (d < 86400000) return Math.floor(d / 3600000) + ' Std.';
  return Math.floor(d / 86400000) + ' Tage';
}
function mtLabel(v) {
  return { bestandsmaterial: '📦 Bestandsmaterial', neues_material: '🆕 Neues Material', dienstleistung: '🔧 Dienstleistung' }[v] || v;
}
function statusBadge(s) {
  const labels = {
    entwurf: 'Entwurf', eingereicht: 'Eingereicht', in_pruefung: 'In Prüfung',
    freigegeben: 'Freigegeben', abgelehnt: 'Abgelehnt', bestellt: 'Bestellt'
  };
  return `<span class="status-badge status-${s}">${labels[s] || s}</span>`;
}
function emptyState(msg) {
  return `<div class="empty-state">${msg}</div>`;
}
