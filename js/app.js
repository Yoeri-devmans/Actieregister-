/* Actieregister — iPhone-webapp. Zie README.md voor uitleg. */
(function () {
  'use strict';
  const R = window.Register;
  const S = window.Store;

  // ---------- toestand ----------
  const state = {
    data: null,          // { meta, actions, fileName, importedAt, exportedAt }
    tab: 'lijst',
    quick: S.pref('quick') || 'actief',
    search: '',
    filters: S.pref('filters') || { prioriteit: '', domein: '', eigenaar: '', status: '' },
    groupBy: S.pref('groupBy') || 'prioriteit',
    me: S.pref('me') || '',
  };

  const $ = (sel, el = document) => el.querySelector(sel);
  const main = $('#main');
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const today = () => R.todayISO();

  const DONE = R.DONE;
  const isActive = (a) => !R.NOT_ACTIVE.has(a.status);
  const owners = (a) => String(a.eigenaar || '').split(/\s*[\/,&+]\s*|\s+en\s+/i).map((s) => s.trim()).filter(Boolean);
  const hasOwner = (a, name) => owners(a).some((o) => o.toLowerCase() === name.toLowerCase());
  const pendingCount = () => state.data ? state.data.actions.filter((a) => a._changed || a._new).length : 0;

  const QUICK = [
    { id: 'mijn', label: 'Mijn acties', test: (a) => state.me && hasOwner(a, state.me) && isActive(a) },
    { id: 'top', label: 'Top 10', test: (a) => a.rang !== '' && a.rang != null && a.status !== DONE, sort: 'rang' },
    { id: 'actief', label: 'Actief', test: isActive },
    { id: 'achterstallig', label: 'Achterstallig', test: (a) => R.isOverdue(a) },
    { id: 'week', label: 'Deze week', test: (a) => { const d = R.daysUntil(a.deadline); return a.status !== DONE && d != null && d <= 7; } },
    { id: 'wacht', label: 'Wacht op', test: (a) => a.status === 'Wacht op' },
    { id: 'verifieren', label: 'Te verifiëren', test: (a) => a.status === 'Te verifiëren' },
    { id: 'ideeen', label: 'Ideeën & plannen', test: (a) => a.status === 'Idee' || a.status === 'Plan' },
    { id: 'klaar', label: 'Klaar', test: (a) => a.status === DONE },
    { id: 'alles', label: 'Alles', test: () => true },
  ];

  const GROUPS = {
    prioriteit: 'Prioriteit', deadline: 'Deadline', domein: 'Domein', eigenaar: 'Eigenaar', status: 'Status', geen: 'Geen',
  };

  const PRIO_LABEL = { P1: 'P1 — nu', P2: 'P2 — deze maand', P3: 'P3 — agenderen' };

  // ---------- opslag ----------
  let saveTimer;
  function persist() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => S.set('data', state.data).catch((e) => toast('Opslaan mislukt: ' + e.message)), 150);
    updatePendingDot();
  }

  async function init() {
    try { state.data = await S.get('data'); } catch (e) { state.data = null; }
    bindChrome();
    render();
    if ('serviceWorker' in navigator && location.protocol !== 'file:') {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }

  // ---------- importeren ----------
  async function importFile(file) {
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(buf);
    } catch (e) {
      throw new Error('Dit bestand kon niet worden gelezen als Excel (.xlsx). ' + e.message);
    }
    const parsed = R.parseWorkbook(wb);
    const { actions, ...meta } = parsed;
    let next = actions.map((a) => ({ ...a }));

    const pending = state.data ? state.data.actions.filter((a) => a._changed || a._new) : [];
    if (pending.length) {
      const choice = await ask('Niet-opgeslagen wijzigingen',
        `Je hebt ${pending.length} wijziging${pending.length === 1 ? '' : 'en'} in de app die nog niet in Excel staan. Wil je die meenemen naar het nieuwe bestand?`,
        [{ id: 'merge', label: 'Meenemen', primary: true }, { id: 'drop', label: 'Weggooien', destructive: true }, { id: 'cancel', label: 'Annuleren' }]);
      if (choice === 'cancel' || !choice) return false;
      if (choice === 'merge') next = mergePending(next, pending);
    }

    await S.set('file', buf);
    state.data = { meta, actions: next, fileName: file.name, importedAt: new Date().toISOString(), exportedAt: null };
    await S.set('data', state.data);
    return true;
  }

  function mergePending(fresh, pending) {
    // wijzigingen opnieuw toepassen op het nieuwe bestand, gekoppeld op actienummer (of tekst)
    const key = (a) => (a.nr !== '' && a.nr != null ? 'nr:' + a.nr : 'actie:' + a.actie);
    const byKey = new Map(fresh.map((a) => [key(a), a]));
    let maxNr = fresh.reduce((m, a) => Math.max(m, Number(a.nr) || 0), 0);
    for (const p of pending) {
      const target = !p._new && byKey.get(key(p));
      if (target) {
        for (const f of p._fields || []) target[f] = p[f];
        target._changed = true;
        target._fields = [...new Set([...(target._fields || []), ...(p._fields || [])])];
      } else if (p._new || !target) {
        const n = { ...p, _new: true, _changed: false, row: null };
        if (fresh.some((a) => String(a.nr) === String(n.nr))) n.nr = ++maxNr;
        fresh.push(n);
      }
    }
    return fresh;
  }

  async function loadDemo() {
    const res = await fetch('tests/fixture.xlsx');
    if (!res.ok) throw new Error('Voorbeeldbestand niet gevonden');
    const blob = await res.blob();
    return importFile(new File([blob], 'Voorbeeld_actieregister.xlsx'));
  }

  // ---------- exporteren ----------
  async function buildExport() {
    const buf = await S.get('file');
    if (!buf) throw new Error('Het originele Excel-bestand ontbreekt. Importeer het opnieuw.');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const parsed = R.parseWorkbook(wb);
    // werk op het actuele bestand: rijnummers opnieuw koppelen via de opgeslagen rij
    const result = R.applyActions(wb, parsed, state.data.actions);
    const out = await wb.xlsx.writeBuffer();
    return { out, result };
  }

  async function doExport() {
    if (!pendingCount()) { toast('Geen wijzigingen om op te slaan'); return; }
    toast('Excel-bestand maken…', 10000);
    let built;
    try { built = await buildExport(); } catch (e) { toast('Mislukt: ' + e.message); return; }
    hideToast();
    const name = exportName();
    const file = new File([built.out], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const canShare = navigator.canShare && navigator.canShare({ files: [file] });

    // iOS vereist dat 'delen' direct op een tik volgt; daarom een tweede knop
    const choice = await ask('Excel-bestand klaar',
      `${built.result.changed} gewijzigd, ${built.result.added} nieuw. Bewaar het bestand in OneDrive/SharePoint (via "Bewaar in Bestanden") of deel het via Teams.`,
      [canShare ? { id: 'share', label: 'Delen / bewaren…', primary: true, onTap: () => navigator.share({ files: [file], title: name }).catch(() => {}) } : null,
        { id: 'download', label: canShare ? 'Downloaden' : 'Downloaden', primary: !canShare, onTap: () => download(file) },
        { id: 'cancel', label: 'Sluiten' }].filter(Boolean));
    if (choice === 'cancel' || !choice) return;

    const done = await ask('Is het bestand opgeslagen?',
      'Als je het bestand hebt bewaard, worden de wijzigingen in de app als opgeslagen gemarkeerd. Volgende keer begin je dan vanaf dit nieuwe bestand.',
      [{ id: 'yes', label: 'Ja, opgeslagen', primary: true }, { id: 'no', label: 'Nog niet' }]);
    if (done === 'yes') {
      await S.set('file', built.out);
      for (const a of state.data.actions) { delete a._changed; delete a._new; delete a._fields; delete a._prevStatus; }
      // rijnummers van nieuwe acties staan nu vast in het bestand
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(built.out);
      const parsed = R.parseWorkbook(wb);
      const { actions, ...meta } = parsed;
      state.data.actions = actions;
      state.data.meta = meta;
      state.data.fileName = name;
      state.data.exportedAt = new Date().toISOString();
      persist();
      render();
      toast('Wijzigingen opgeslagen');
    }
  }

  function exportName() {
    const base = (state.data.fileName || 'Actieregister.xlsx').replace(/\.xlsx$/i, '').replace(/_\d{4}-\d{2}-\d{2}( \d+)?$/, '');
    return `${base}.xlsx`;
  }

  function download(file) {
    const url = URL.createObjectURL(file);
    const a = document.createElement('a');
    a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  // ---------- wijzigen ----------
  function update(a, field, value) {
    if (a[field] === value) return;
    a[field] = value;
    if (!a._new) {
      a._changed = true;
      a._fields = [...new Set([...(a._fields || []), field])];
    }
    persist();
  }

  function toggleDone(a) {
    if (a.status === DONE) {
      update(a, 'status', a._prevStatus || 'Open');
      delete a._prevStatus;
      toast('Weer open gezet');
    } else {
      const prev = a.status;
      a._prevStatus = prev;
      update(a, 'status', DONE);
      toast('Afgevinkt', 4000, { label: 'Ongedaan maken', fn: () => { update(a, 'status', prev || 'Open'); delete a._prevStatus; persist(); render(); } });
    }
    render();
  }

  function newAction() {
    const acts = state.data.actions;
    const maxNr = acts.reduce((m, a) => Math.max(m, Number(a.nr) || 0), 0);
    return {
      _new: true, row: null, extra: {}, nr: maxNr + 1, actie: '', domein: '', eigenaar: state.me || '',
      prioriteit: 'P2', deadline: '', status: 'Open', rang: '', waarom: '', aanpak: '', bron: 'App', besluitdatum: today(), toelichting: '',
    };
  }

  // ---------- weergave: algemeen ----------
  function bindChrome() {
    $('#tabbar').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-tab]');
      if (!b) return;
      state.tab = b.dataset.tab;
      window.scrollTo(0, 0);
      render();
    });
    $('#search').addEventListener('input', (e) => { state.search = e.target.value; renderList(); });
    $('#filter-btn').addEventListener('click', openFilterSheet);
    $('#add-btn').addEventListener('click', () => openDetail(newAction(), true));
    $('#file-input').addEventListener('change', async (e) => {
      const f = e.target.files[0];
      e.target.value = '';
      if (!f) return;
      try {
        toast('Inlezen…', 10000);
        const ok = await importFile(f);
        hideToast();
        if (ok) { state.tab = 'lijst'; render(); toast(`${state.data.actions.length} acties ingelezen`); }
      } catch (err) { toast(err.message, 6000); }
    });
    $('#quick').addEventListener('click', (e) => {
      const c = e.target.closest('[data-quick]');
      if (!c) return;
      state.quick = c.dataset.quick;
      S.pref('quick', state.quick);
      renderList();
      c.scrollIntoView({ inline: 'center', block: 'nearest', behavior: 'smooth' });
    });
  }

  function render() {
    const has = !!(state.data && state.data.actions);
    $('#tabbar').hidden = !has;
    $('#add-btn').hidden = !has || state.tab !== 'lijst';
    $('#list-controls').hidden = !has || state.tab !== 'lijst';
    document.querySelectorAll('#tabbar button').forEach((b) => b.classList.toggle('active', b.dataset.tab === state.tab));
    updatePendingDot();
    if (!has) return renderWelcome();
    if (state.tab === 'lijst') renderList();
    else if (state.tab === 'overzicht') renderOverview();
    else renderFile();
  }

  function updatePendingDot() {
    const dot = $('#pending-dot');
    if (dot) dot.hidden = !pendingCount();
  }

  function setTitle(t) { $('#page-title').textContent = t; document.title = t; }

  // ---------- welkom ----------
  function renderWelcome() {
    setTitle('Actieregister');
    main.innerHTML = `
      <section class="welcome">
        <div class="welcome-icon"><img src="icons/icon-192.png" alt="" width="88" height="88"></div>
        <h2>Je actieregister als to-do-lijst</h2>
        <p>Open het Excel-bestand (bijvoorbeeld <em>Actieregister_Marsdijk.xlsx</em> uit SharePoint/OneDrive via de app Bestanden). De gegevens blijven alleen op deze telefoon.</p>
        <button class="btn primary" data-act="import">Excel-bestand openen</button>
        <button class="btn plain" data-act="demo">Probeer met voorbeelddata</button>
        <details class="howto">
          <summary>Zet de app op je beginscherm</summary>
          <ol>
            <li>Open deze pagina in <strong>Safari</strong>.</li>
            <li>Tik op de deelknop <span aria-hidden="true">⬆︎</span> onderin.</li>
            <li>Kies <strong>Zet op beginscherm</strong>.</li>
          </ol>
        </details>
      </section>`;
    main.onclick = async (e) => {
      const b = e.target.closest('[data-act]');
      if (!b) return;
      if (b.dataset.act === 'import') $('#file-input').click();
      if (b.dataset.act === 'demo') {
        try { await loadDemo(); render(); } catch (err) { toast(err.message); }
      }
    };
  }

  // ---------- lijst ----------
  function activeFilterCount() {
    return Object.values(state.filters).filter(Boolean).length;
  }

  function filtered() {
    const q = QUICK.find((x) => x.id === state.quick) || QUICK[2];
    const f = state.filters;
    const terms = state.search.toLowerCase().split(/\s+/).filter(Boolean);
    const list = state.data.actions.filter((a) => {
      if (!q.test(a)) return false;
      if (f.prioriteit && a.prioriteit !== f.prioriteit) return false;
      if (f.domein && a.domein !== f.domein) return false;
      if (f.status && a.status !== f.status) return false;
      if (f.eigenaar && !hasOwner(a, f.eigenaar)) return false;
      if (terms.length) {
        const hay = [a.nr, a.actie, a.eigenaar, a.domein, a.toelichting, a.aanpak, a.waarom, a.bron, a.status].join(' ').toLowerCase();
        if (!terms.every((t) => hay.includes(t))) return false;
      }
      return true;
    });
    return { list, sortRang: q.sort === 'rang' };
  }

  const prioRank = (p) => ({ P1: 1, P2: 2, P3: 3 }[p] || 9);

  function compare(a, b) {
    const ra = Number(a.rang) || 999, rb = Number(b.rang) || 999;
    if (ra !== rb) return ra - rb;
    const pa = prioRank(a.prioriteit), pb = prioRank(b.prioriteit);
    if (pa !== pb) return pa - pb;
    const da = R.isISODate(a.deadline) ? a.deadline : '9999', db = R.isISODate(b.deadline) ? b.deadline : '9999';
    if (da !== db) return da < db ? -1 : 1;
    return (Number(a.nr) || 0) - (Number(b.nr) || 0);
  }

  function deadlineBucket(a) {
    if (R.isOverdue(a)) return [0, 'Achterstallig'];
    const d = R.daysUntil(a.deadline);
    if (d == null) return [5, a.deadline ? 'Overig' : 'Geen deadline'];
    if (d <= 7) return [1, 'Deze week'];
    if (d <= 31) return [2, 'Binnen een maand'];
    if (d <= 92) return [3, 'Binnen drie maanden'];
    return [4, 'Later'];
  }

  function groupsOf(list, sortRang) {
    const by = sortRang ? 'geen' : state.groupBy;
    const sorted = [...list].sort(compare);
    if (by === 'geen') return [{ key: '', label: '', items: sorted }];
    const map = new Map();
    for (const a of sorted) {
      let keys;
      if (by === 'prioriteit') keys = [[prioRank(a.prioriteit), PRIO_LABEL[a.prioriteit] || a.prioriteit || 'Geen prioriteit']];
      else if (by === 'deadline') keys = [deadlineBucket(a)];
      else if (by === 'domein') keys = [[a.domein || 'zzz', a.domein || 'Geen domein']];
      else if (by === 'status') keys = [[state.data.meta.lists.status.indexOf(a.status), a.status || 'Geen status']];
      else if (by === 'eigenaar') keys = (owners(a).length ? owners(a) : ['Niemand']).map((o) => [o === 'Niemand' ? 'zzz' : o.toLowerCase(), o]);
      for (const [sortKey, label] of keys) {
        if (!map.has(label)) map.set(label, { key: sortKey, label, items: [] });
        map.get(label).items.push(a);
      }
    }
    return [...map.values()].sort((x, y) => (x.key < y.key ? -1 : x.key > y.key ? 1 : 0));
  }

  function renderQuick() {
    const counts = {};
    for (const q of QUICK) counts[q.id] = state.data.actions.filter(q.test).length;
    $('#quick').innerHTML = QUICK
      .filter((q) => q.id !== 'mijn' || state.me)
      .map((q) => `<button class="chip ${q.id === state.quick ? 'on' : ''} ${q.id === 'achterstallig' && counts[q.id] ? 'warn' : ''}" data-quick="${q.id}">${esc(q.label)} <span>${counts[q.id]}</span></button>`)
      .join('');
    const n = activeFilterCount();
    const fc = $('#filter-count');
    fc.hidden = !n; fc.textContent = n;
    $('#filter-btn').classList.toggle('on', n > 0 || state.groupBy !== 'prioriteit');
  }

  function renderList() {
    setTitle(state.data.meta.title.replace(/\s+—.*$/, '') || 'Actieregister');
    if (state.quick === 'mijn' && !state.me) state.quick = 'actief';
    renderQuick();
    const { list, sortRang } = filtered();
    const groups = groupsOf(list, sortRang);
    if (!list.length) {
      main.innerHTML = `<div class="empty"><p>${state.search ? 'Niets gevonden voor deze zoekopdracht.' : 'Geen acties in deze weergave.'}</p>${activeFilterCount() ? '<button class="btn plain" data-clear>Filters wissen</button>' : ''}</div>`;
    } else {
      main.innerHTML = groups.map((g) => `
        <section class="group">
          ${g.label ? `<h3 class="group-title">${esc(g.label)} <span>${g.items.length}</span></h3>` : ''}
          <ul class="cards">${g.items.map(cardHTML).join('')}</ul>
        </section>`).join('') + `<p class="footnote">${list.length} van ${state.data.actions.length} acties</p>`;
    }
    main.onclick = (e) => {
      if (e.target.closest('[data-clear]')) { state.filters = { prioriteit: '', domein: '', eigenaar: '', status: '' }; S.pref('filters', state.filters); renderList(); return; }
      const li = e.target.closest('[data-idx]');
      if (!li) return;
      const a = state.data.actions[Number(li.dataset.idx)];
      if (e.target.closest('.check')) toggleDone(a);
      else openDetail(a, false);
    };
  }

  function fmtDate(iso, opts = { weekday: 'short', day: 'numeric', month: 'short' }) {
    if (!R.isISODate(iso)) return iso || '';
    const [y, m, d] = iso.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    const o = { ...opts };
    if (y !== new Date().getFullYear()) o.year = 'numeric';
    return dt.toLocaleDateString('nl-NL', o);
  }

  function relDeadline(a) {
    if (!a.deadline) return '';
    const d = R.daysUntil(a.deadline);
    if (d == null) return esc(a.deadline);
    if (a.status === DONE) return fmtDate(a.deadline);
    if (d < 0) return `${-d} ${-d === 1 ? 'dag' : 'dagen'} te laat`;
    if (d === 0) return 'vandaag';
    if (d === 1) return 'morgen';
    if (d < 7) return fmtDate(a.deadline, { weekday: 'long' });
    return fmtDate(a.deadline);
  }

  function statusClass(s) {
    return 's-' + String(s || 'leeg').toLowerCase().normalize('NFD').replace(/[^a-z]/g, '');
  }

  function cardHTML(a) {
    const idx = state.data.actions.indexOf(a);
    const overdue = R.isOverdue(a);
    const done = a.status === DONE;
    const dl = relDeadline(a);
    return `
      <li class="card ${done ? 'done' : ''}" data-idx="${idx}">
        <button class="check" aria-label="${done ? 'Weer openzetten' : 'Afvinken'}"><span></span></button>
        <div class="card-body">
          <div class="card-title">${a.rang ? `<span class="rang">#${esc(a.rang)}</span>` : ''}${esc(a.actie)}</div>
          <div class="card-meta">
            ${a.prioriteit ? `<span class="prio p-${esc(a.prioriteit).toLowerCase()}">${esc(a.prioriteit)}</span>` : ''}
            ${a.status && a.status !== 'Open' ? `<span class="status ${statusClass(a.status)}">${esc(a.status)}</span>` : ''}
            ${dl ? `<span class="due ${overdue ? 'late' : ''}">${dl}</span>` : ''}
            ${a.eigenaar ? `<span class="owner">${esc(a.eigenaar)}</span>` : ''}
            ${a._changed || a._new ? '<span class="dirty" title="Nog niet in Excel opgeslagen"></span>' : ''}
          </div>
        </div>
      </li>`;
  }

  // ---------- detail ----------
  function openDetail(a, isNew) {
    const meta = state.data.meta;
    const lists = meta.lists;
    const cols = meta.columns || {};
    const has = (f) => isNew || f in cols || f in a;
    const allOwners = [...new Set(state.data.actions.flatMap(owners))].sort((x, y) => x.localeCompare(y, 'nl'));
    const draft = isNew ? a : null;
    const days = R.daysOpen(a);

    const opt = (list, val, empty) => (empty ? `<option value="">${empty}</option>` : '') +
      [...new Set([...list, ...(val && !list.includes(val) ? [val] : [])])].map((v) => `<option ${v === val ? 'selected' : ''}>${esc(v)}</option>`).join('');

    const text = (f, label, rows = 3) => has(f) ? `
      <label class="field"><span>${label}</span><textarea name="${f}" rows="${rows}">${esc(a[f])}</textarea></label>` : '';

    const extra = Object.entries(a.extra || {}).map(([k, v]) => `<div class="field ro"><span>${esc(k)}</span><p>${esc(v)}</p></div>`).join('');

    const html = `
      <div class="sheet-head">
        <button class="link" data-close>${isNew ? 'Annuleer' : 'Sluit'}</button>
        <strong>${isNew ? 'Nieuwe actie' : 'Actie ' + esc(a.nr)}</strong>
        ${isNew ? '<button class="link bold" data-save>Voeg toe</button>' : `<button class="link bold" data-done>${a.status === DONE ? 'Heropen' : 'Klaar'}</button>`}
      </div>
      <form class="sheet-body" autocomplete="off" onsubmit="return false">
        <label class="field big"><span>Actie</span><textarea name="actie" rows="3" placeholder="Wat moet er gebeuren?">${esc(a.actie)}</textarea></label>

        <div class="field"><span>Prioriteit</span>
          <div class="segmented" data-seg="prioriteit">${lists.prioriteit.map((p) => `<button type="button" class="${a.prioriteit === p ? 'on' : ''}" data-v="${esc(p)}">${esc(p)}</button>`).join('')}</div>
        </div>

        <div class="grid2">
          <label class="field"><span>Status</span><select name="status">${opt(lists.status, a.status)}</select></label>
          <label class="field"><span>Deadline</span><input type="date" name="deadline" value="${R.isISODate(a.deadline) ? a.deadline : ''}"></label>
        </div>
        ${a.deadline && !R.isISODate(a.deadline) ? `<p class="hint">Deadline in Excel: “${esc(a.deadline)}”</p>` : ''}

        <label class="field"><span>Eigenaar</span><input name="eigenaar" list="owners" value="${esc(a.eigenaar)}" placeholder="Bijv. Anna / Bert"></label>
        <datalist id="owners">${allOwners.map((o) => `<option value="${esc(o)}">`).join('')}</datalist>

        ${has('domein') ? `<label class="field"><span>Domein</span><select name="domein">${opt(lists.domein, a.domein, '—')}</select></label>` : ''}

        ${(R.isOverdue(a) || days != null) ? `<div class="facts">
          ${R.isOverdue(a) ? '<span class="late">Achterstallig</span>' : ''}
          ${days != null ? `<span>${days} dagen open</span>` : ''}
        </div>` : ''}

        ${text('waarom', 'Waarom nu', 3)}
        ${text('aanpak', 'Aanpak — eerste stap', 4)}
        ${text('toelichting', 'Toelichting', 3)}

        <details class="more" ${isNew ? '' : ''}>
          <summary>Meer velden</summary>
          <div class="grid2">
            ${has('rang') ? `<label class="field"><span>Rang</span><input type="number" inputmode="numeric" name="rang" value="${esc(a.rang)}"></label>` : ''}
            ${has('besluitdatum') ? `<label class="field"><span>Besluitdatum</span><input type="date" name="besluitdatum" value="${R.isISODate(a.besluitdatum) ? a.besluitdatum : ''}"></label>` : ''}
          </div>
          ${text('bron', 'Bron', 2)}
          ${extra}
        </details>
        ${a._changed ? `<p class="hint">Gewijzigd in de app (${esc((a._fields || []).join(', '))}) — nog niet in Excel opgeslagen.</p>` : ''}
        ${a._new && !isNew ? '<p class="hint">Nieuwe actie — nog niet in Excel opgeslagen.</p><button type="button" class="btn danger" data-delete>Verwijder deze nieuwe actie</button>' : ''}
      </form>`;

    const sheet = openSheet(html, { full: true });
    const form = $('form', sheet);
    const target = draft || a;

    const setField = (name, value) => {
      if (name === 'rang') value = value === '' ? '' : Number(value);
      if (isNew) target[name] = value; else update(target, name, value);
    };
    form.addEventListener('change', (e) => {
      const el = e.target;
      if (!el.name) return;
      setField(el.name, el.value.trim ? el.value.trim() : el.value);
    });
    form.addEventListener('input', (e) => {
      const el = e.target;
      if (el.tagName === 'TEXTAREA') { el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px'; }
    });
    form.querySelectorAll('textarea').forEach((t) => { requestAnimationFrame(() => { t.style.height = 'auto'; t.style.height = t.scrollHeight + 'px'; }); });
    form.addEventListener('click', (e) => {
      const b = e.target.closest('[data-seg] button');
      if (!b) return;
      const seg = b.parentElement;
      seg.querySelectorAll('button').forEach((x) => x.classList.toggle('on', x === b));
      setField(seg.dataset.seg, b.dataset.v);
    });

    sheet.addEventListener('click', (e) => {
      if (e.target.closest('[data-close]')) { closeSheet(sheet); render(); }
      else if (e.target.closest('[data-save]')) {
        // laatste wijziging in een veld dat nog focus heeft ook meenemen
        form.querySelectorAll('[name]').forEach((el) => { target[el.name] = el.name === 'rang' ? (el.value === '' ? '' : Number(el.value)) : el.value.trim(); });
        if (!target.actie) { toast('Vul eerst de actie in'); return; }
        state.data.actions.push(target);
        persist();
        closeSheet(sheet);
        render();
        toast('Actie toegevoegd');
      } else if (e.target.closest('[data-done]')) {
        closeSheet(sheet);
        toggleDone(a);
      } else if (e.target.closest('[data-delete]')) {
        state.data.actions.splice(state.data.actions.indexOf(a), 1);
        persist(); closeSheet(sheet); render(); toast('Verwijderd');
      }
    });
    if (isNew) setTimeout(() => $('textarea[name=actie]', sheet).focus(), 350);
  }

  // ---------- filters ----------
  function openFilterSheet() {
    const lists = state.data.meta.lists;
    const allOwners = [...new Set(state.data.actions.flatMap(owners))].sort((x, y) => x.localeCompare(y, 'nl'));
    const sel = (name, list, label) => `
      <label class="field"><span>${label}</span><select name="${name}"><option value="">Alle</option>${list.map((v) => `<option ${state.filters[name] === v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select></label>`;
    const html = `
      <div class="sheet-head">
        <button class="link" data-reset>Wis</button>
        <strong>Filters</strong>
        <button class="link bold" data-close>Gereed</button>
      </div>
      <form class="sheet-body" onsubmit="return false">
        <div class="field"><span>Groeperen op</span>
          <div class="segmented wrap" data-seg="groupBy">${Object.entries(GROUPS).map(([k, v]) => `<button type="button" class="${state.groupBy === k ? 'on' : ''}" data-v="${k}">${v}</button>`).join('')}</div>
        </div>
        ${sel('prioriteit', lists.prioriteit, 'Prioriteit')}
        ${sel('eigenaar', allOwners, 'Eigenaar')}
        ${sel('domein', lists.domein, 'Domein')}
        ${sel('status', lists.status, 'Status')}
        <p class="hint">Filters werken bovenop de gekozen weergave (bijv. “Actief” of “Alles”).</p>
      </form>`;
    const sheet = openSheet(html);
    sheet.addEventListener('change', (e) => {
      if (!e.target.name) return;
      state.filters[e.target.name] = e.target.value;
      S.pref('filters', state.filters);
      renderList();
    });
    sheet.addEventListener('click', (e) => {
      const b = e.target.closest('[data-seg] button');
      if (b) {
        b.parentElement.querySelectorAll('button').forEach((x) => x.classList.toggle('on', x === b));
        state.groupBy = b.dataset.v; S.pref('groupBy', state.groupBy); renderList();
      }
      if (e.target.closest('[data-reset]')) {
        state.filters = { prioriteit: '', domein: '', eigenaar: '', status: '' };
        state.groupBy = 'prioriteit';
        S.pref('filters', state.filters); S.pref('groupBy', state.groupBy);
        closeSheet(sheet); renderList();
      }
      if (e.target.closest('[data-close]')) closeSheet(sheet);
    });
  }

  // ---------- overzicht ----------
  function renderOverview() {
    setTitle('Overzicht');
    const acts = state.data.actions;
    const active = acts.filter(isActive);
    const count = (fn, list = acts) => list.filter(fn).length;
    const tiles = [
      { label: 'Actief', n: active.length, go: { quick: 'actief' } },
      { label: 'Achterstallig', n: count((a) => R.isOverdue(a)), go: { quick: 'achterstallig' }, warn: true },
      { label: 'Deze week', n: count(QUICK.find((q) => q.id === 'week').test), go: { quick: 'week' } },
      { label: 'Wacht op anderen', n: count((a) => a.status === 'Wacht op'), go: { quick: 'wacht' } },
      { label: 'Te verifiëren', n: count((a) => a.status === 'Te verifiëren'), go: { quick: 'verifieren' } },
      { label: 'Afgerond', n: count((a) => a.status === DONE), go: { quick: 'klaar' }, ok: true },
    ];
    const prio = ['P1', 'P2', 'P3'].map((p) => ({ label: PRIO_LABEL[p], key: p, n: count((a) => a.prioriteit === p, active) }));
    const domains = state.data.meta.lists.domein.map((d) => ({ label: d, n: count((a) => a.domein === d, active) })).filter((d) => d.n);
    const ownerCounts = {};
    for (const a of active) for (const o of owners(a)) ownerCounts[o] = (ownerCounts[o] || 0) + 1;
    const ownerList = Object.entries(ownerCounts).sort((a, b) => b[1] - a[1]).map(([label, n]) => ({ label, n }));
    const maxD = Math.max(1, ...domains.map((d) => d.n));
    const maxO = Math.max(1, ...ownerList.map((d) => d.n));

    const bars = (rows, max, kind) => rows.map((r) => `
      <button class="bar-row" data-kind="${kind}" data-v="${esc(r.key || r.label)}">
        <span class="bar-label">${esc(r.label)}</span>
        <span class="bar"><i style="width:${r.n ? Math.max(4, (r.n / max) * 100) : 0}%"></i></span>
        <span class="bar-n">${r.n}</span>
      </button>`).join('');

    main.innerHTML = `
      ${state.data.meta.subtitle ? `<p class="subtitle">${esc(state.data.meta.subtitle)}</p>` : ''}
      <div class="tiles">${tiles.map((t, i) => `
        <button class="tile ${t.warn && t.n ? 'warn' : ''} ${t.ok ? 'ok' : ''}" data-tile="${i}">
          <strong>${t.n}</strong><span>${t.label}</span>
        </button>`).join('')}</div>
      <section class="panel"><h3>Actief per prioriteit</h3>${bars(prio, Math.max(1, ...prio.map((p) => p.n)), 'prioriteit')}</section>
      <section class="panel"><h3>Actief per domein</h3>${bars(domains, maxD, 'domein')}</section>
      <section class="panel"><h3>Actief per eigenaar</h3>${bars(ownerList, maxO, 'eigenaar')}</section>
      <p class="footnote">“Actief” = alles behalve Klaar, Idee en Plan — net als het tabblad Overzicht in Excel.</p>`;

    main.onclick = (e) => {
      const t = e.target.closest('[data-tile]');
      const b = e.target.closest('.bar-row');
      if (t) {
        state.quick = tiles[Number(t.dataset.tile)].go.quick;
        state.filters = { prioriteit: '', domein: '', eigenaar: '', status: '' };
      } else if (b) {
        state.quick = 'actief';
        state.filters = { prioriteit: '', domein: '', eigenaar: '', status: '', [b.dataset.kind]: b.dataset.v };
      } else return;
      S.pref('quick', state.quick); S.pref('filters', state.filters);
      state.tab = 'lijst';
      window.scrollTo(0, 0);
      render();
    };
  }

  // ---------- bestand / instellingen ----------
  function renderFile() {
    setTitle('Bestand');
    const d = state.data;
    const n = pendingCount();
    const allOwners = [...new Set(d.actions.flatMap(owners))].sort((x, y) => x.localeCompare(y, 'nl'));
    const fmtTime = (iso) => iso ? new Date(iso).toLocaleString('nl-NL', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—';
    const pending = d.actions.filter((a) => a._changed || a._new);

    main.innerHTML = `
      <section class="panel">
        <h3>Wijzigingen</h3>
        ${n ? `<p><strong>${n}</strong> actie${n === 1 ? '' : 's'} gewijzigd of toegevoegd in de app, nog niet in Excel.</p>
          <ul class="changes">${pending.slice(0, 8).map((a) => `<li><span class="${a._new ? 'tag-new' : 'tag-mod'}">${a._new ? 'nieuw' : 'gewijzigd'}</span> ${esc(a.nr)}. ${esc(a.actie)}</li>`).join('')}${pending.length > 8 ? `<li>… en ${pending.length - 8} meer</li>` : ''}</ul>
          <button class="btn primary" data-act="export">Opslaan als Excel</button>
          <p class="hint">Maakt een nieuwe versie van het originele bestand: opmaak, formules, dropdowns en de tabbladen Overzicht en Toelichting blijven behouden. Vervang daarna het bestand in SharePoint.</p>`
        : '<p>Alles in de app staat ook in het Excel-bestand.</p>'}
      </section>

      <section class="panel">
        <h3>Ik ben</h3>
        <label class="field"><select name="me"><option value="">— kies je naam —</option>${allOwners.map((o) => `<option ${o === state.me ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></label>
        <p class="hint">Hiermee werkt de weergave “Mijn acties” (ook acties die je deelt, zoals “Anna / Bert”).</p>
      </section>

      <section class="panel">
        <h3>Excel-bestand</h3>
        <dl class="kv">
          <dt>Bestand</dt><dd>${esc(d.fileName || '—')}</dd>
          <dt>Ingelezen</dt><dd>${fmtTime(d.importedAt)}</dd>
          <dt>Laatst opgeslagen</dt><dd>${fmtTime(d.exportedAt)}</dd>
          <dt>Acties</dt><dd>${d.actions.length}</dd>
        </dl>
        <button class="btn" data-act="import">Nieuwere versie inlezen</button>
        <p class="hint">Na elk overleg: open hier de nieuwste versie uit SharePoint. Wijzigingen uit de app die nog niet zijn opgeslagen kun je meenemen.</p>
      </section>

      <section class="panel">
        <h3>Privacy</h3>
        <p class="hint">Het register bevat vertrouwelijke gegevens. Deze app stuurt niets naar internet: alles staat alleen op dit toestel.</p>
        <button class="btn danger" data-act="wipe">Wis gegevens van dit toestel</button>
      </section>`;

    main.onchange = (e) => {
      if (e.target.name === 'me') { state.me = e.target.value; S.pref('me', state.me); if (state.me) { state.quick = 'mijn'; S.pref('quick', 'mijn'); } toast(state.me ? `“Mijn acties” staat nu op ${state.me}` : 'Naam gewist'); }
    };
    main.onclick = async (e) => {
      const b = e.target.closest('[data-act]');
      if (!b) return;
      if (b.dataset.act === 'export') doExport();
      if (b.dataset.act === 'import') $('#file-input').click();
      if (b.dataset.act === 'wipe') {
        const extra = n ? ` Let op: ${n} wijziging${n === 1 ? '' : 'en'} ${n === 1 ? 'is' : 'zijn'} nog niet opgeslagen en ${n === 1 ? 'gaat' : 'gaan'} verloren.` : '';
        const ok = await ask('Gegevens wissen?', 'Het register wordt van deze telefoon verwijderd.' + extra,
          [{ id: 'yes', label: 'Wis', destructive: true }, { id: 'no', label: 'Annuleer' }]);
        if (ok === 'yes') { await S.clear(); state.data = null; state.tab = 'lijst'; render(); }
      }
    };
  }

  // ---------- sheets, dialogen, toasts ----------
  function openSheet(html, { full = false } = {}) {
    const root = $('#sheet-root');
    const wrap = document.createElement('div');
    wrap.className = 'sheet-wrap';
    wrap.innerHTML = `<div class="sheet-backdrop"></div><div class="sheet ${full ? 'full' : ''}" role="dialog" aria-modal="true">${html}</div>`;
    root.appendChild(wrap);
    document.body.classList.add('locked');
    requestAnimationFrame(() => wrap.classList.add('open'));
    $('.sheet-backdrop', wrap).addEventListener('click', () => { closeSheet(wrap); if (state.tab === 'lijst') renderList(); });
    return wrap;
  }

  function closeSheet(wrap) {
    if (document.activeElement) document.activeElement.blur();
    wrap.classList.remove('open');
    setTimeout(() => {
      wrap.remove();
      if (!$('#sheet-root').children.length) document.body.classList.remove('locked');
    }, 280);
  }

  function ask(title, text, buttons) {
    return new Promise((resolve) => {
      const wrap = document.createElement('div');
      wrap.className = 'dialog-wrap';
      wrap.innerHTML = `<div class="dialog" role="alertdialog" aria-modal="true">
        <h4>${esc(title)}</h4><p>${esc(text)}</p>
        <div class="dialog-btns">${buttons.map((b) => `<button data-id="${b.id}" class="${b.primary ? 'primary' : ''} ${b.destructive ? 'destructive' : ''}">${esc(b.label)}</button>`).join('')}</div>
      </div>`;
      document.body.appendChild(wrap);
      requestAnimationFrame(() => wrap.classList.add('open'));
      wrap.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-id]');
        if (!btn) return;
        const def = buttons.find((b) => b.id === btn.dataset.id);
        if (def && def.onTap) def.onTap(); // synchroon binnen de tik (nodig voor iOS-delen)
        wrap.remove();
        resolve(btn.dataset.id);
      });
    });
  }

  let toastTimer;
  function toast(msg, ms = 2600, action) {
    const t = $('#toast');
    t.innerHTML = `<span>${esc(msg)}</span>${action ? `<button>${esc(action.label)}</button>` : ''}`;
    t.hidden = false;
    requestAnimationFrame(() => t.classList.add('show'));
    if (action) $('button', t).onclick = () => { action.fn(); hideToast(); };
    clearTimeout(toastTimer);
    toastTimer = setTimeout(hideToast, ms);
  }
  function hideToast() {
    const t = $('#toast');
    t.classList.remove('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 250);
  }

  init();
})();
