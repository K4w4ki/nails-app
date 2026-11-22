'use strict';
/* =========================================================
   CLIENTES — busca por nome/WhatsApp + histórico completo
   Requer login de admin, igual às outras telas internas.
   ========================================================= */

let allClientesCache = null;
let allAgendamentosCache = null;

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('clientSearch')) return;

  // Mesmo padrão de proteção de página usado no dashboard.js:
  // sem login de admin válido, volta para admin.html.
  auth.onAuthStateChanged(async (user) => {
    if (!user) { location.replace('admin.html'); return; }
    try {
      const adm = await db.collection('admins').doc(user.uid).get();
      if (!adm.exists) { await auth.signOut(); location.replace('admin.html'); }
    } catch (e) {
      console.error(e);
      location.replace('admin.html');
    }
  });

  document.getElementById('clientSearchBtn').addEventListener('click', searchClients);
  document.getElementById('clientSearch').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') searchClients();
  });
});

async function loadAllData() {
  if (allClientesCache && allAgendamentosCache) return;
  const [cs, ag] = await Promise.all([
    db.collection('clientes').get(),
    db.collection('agendamentos').get()
  ]);
  allClientesCache = cs.docs.map(d => ({ id: d.id, ...d.data() }));
  allAgendamentosCache = ag.docs.map(d => ({ id: d.id, ...d.data() }));
}

async function searchClients() {
  const raw = document.getElementById('clientSearch').value.trim();
  const resultsEl = document.getElementById('clientResults');
  document.getElementById('clientDetail').classList.add('hidden');

  if (!raw) { toast('Digite um nome ou WhatsApp para buscar.'); return; }

  resultsEl.innerHTML = '<div class="empty">Buscando...</div>';

  try {
    await loadAllData();
    const q = raw.toLowerCase();
    const qDigits = digits(raw);

    const matches = allClientesCache.filter((c) => {
      const nameMatch = (c.nome || '').toLowerCase().includes(q);
      const phoneMatch = qDigits.length >= 3 && (c.telefone || '').includes(qDigits);
      return nameMatch || phoneMatch;
    });

    if (!matches.length) {
      resultsEl.innerHTML = '<div class="empty">Nenhuma cliente encontrada com esse termo.</div>';
      return;
    }

    resultsEl.innerHTML = '';
    matches.forEach((c) => {
      const hist = historyFor(c);
      const concluidos = hist.filter(a => a.status === 'concluido');
      const totalGasto = concluidos.reduce((s, a) => s + (+a.preco || 0), 0);
      const isLoyal = hist.some(a => a.status === 'confirmado' || a.status === 'concluido');

      const row = document.createElement('article');
      row.className = 'client-row card';
      row.innerHTML = `
        <div class="client-row-main">
          <h3>${esc(c.nome || 'Sem nome')} ${isLoyal ? '<span class="badge loyal">Fidelizada</span>' : ''}</h3>
          <p>${formatPhoneDisplay(c.telefone)} · ${hist.length} agendamento${hist.length === 1 ? '' : 's'} · ${concluidos.length} concluído${concluidos.length === 1 ? '' : 's'}</p>
        </div>
        <div class="client-row-stats">
          <b>${money(totalGasto)}</b>
          <span>total gasto</span>
        </div>`;
      row.addEventListener('click', () => renderClientDetail(c, hist));
      resultsEl.appendChild(row);
    });
  } catch (e) {
    console.error(e);
    resultsEl.innerHTML = '<div class="empty">Erro ao buscar clientes.</div>';
  }
}

function historyFor(cliente) {
  return allAgendamentosCache.filter(a => a.clienteId === cliente.id || a.telefone === cliente.telefone);
}

function renderClientDetail(c, hist) {
  const el = document.getElementById('clientDetail');
  el.classList.remove('hidden');
  el.scrollIntoView({ behavior: 'smooth', block: 'start' });

  const sorted = [...hist].sort((a, b) => (b.data || '').localeCompare(a.data || ''));
  const concluidos = hist.filter(a => a.status === 'concluido');
  const totalGasto = concluidos.reduce((s, a) => s + (+a.preco || 0), 0);
  const canceladosCliente = hist.filter(a => a.status === 'cancelado' && a.canceladoPor === 'cliente').length;
  const reagendamentos = hist.reduce((s, a) => s + (+a.reagendamentos || 0), 0);

  el.innerHTML = `
    <article class="card">
      <div class="client-detail-head">
        <div>
          <span class="eyebrow">Histórico</span>
          <h2>${esc(c.nome || 'Sem nome')}</h2>
          <p class="lead" style="font-size:13px">${formatPhoneDisplay(c.telefone)}</p>
        </div>
        <button class="btn btn-ghost" id="closeDetail">Fechar</button>
      </div>

      <div class="client-summary-grid">
        <div><b>${hist.length}</b><span>Agendamentos no total</span></div>
        <div><b>${concluidos.length}</b><span>Concluídos</span></div>
        <div><b>${money(totalGasto)}</b><span>Total gasto</span></div>
        <div><b>${reagendamentos}</b><span>Reagendamentos</span></div>
        <div><b>${canceladosCliente}</b><span>Cancelados por ela</span></div>
      </div>

      <hr style="border:0;border-top:1px solid var(--line);margin:20px 0">

      <div class="client-history-list">
        ${sorted.length ? sorted.map(a => `
          <div class="client-history-row">
            <div>
              <strong>${esc(a.servico || '—')}</strong>
              <span>${br(a.data)} às ${a.horario || '—'}</span>
            </div>
            <div>
              <span class="badge ${a.status || ''}">${(a.status || '—').replace('-', ' ')}</span>
              <b>${money(a.preco)}</b>
            </div>
          </div>`).join('') : '<div class="empty">Nenhum agendamento registrado.</div>'}
      </div>
    </article>`;

  document.getElementById('closeDetail').addEventListener('click', () => el.classList.add('hidden'));
}
