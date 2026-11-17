'use strict';
/* =========================================================
   DASHBOARD EM TEMPO REAL — Área da Carol
   Depende de: firebase-config.js (db, auth), app-core.js
   (money, keyDate, br, toast) e Chart.js 4.
   ========================================================= */

const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const MESES_CURTO = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const REDUCE = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/* Cores (mesma família do styles.css) */
const C = { ink:'#282828', ink2:'#3d3d3d', gold:'#b08a4a', violet:'#8b7fa3', wine:'#7a3d40' };

/* Status exibidos no gráfico de rosca (ordem = ordem da legenda) */
const STATUS_META = [
  { k:'pendente',          label:'Pendentes',            color:'#b08a4a' },
  { k:'aguardando-sinal',  label:'Aguardando sinal',     color:'#c9763a' },
  { k:'confirmado',        label:'Confirmados',          color:'#46654d' },
  { k:'concluido',         label:'Concluídos',           color:'#42597a' },
  { k:'reagendado',        label:'Reagendados',          color:'#8b7fa3' },
  { k:'recusado',          label:'Recusados',            color:'#b56a6d' },
  { k:'cancelado-carol',   label:'Cancelados (Carol)',   color:'#7a3d40' },
  { k:'cancelado-cliente', label:'Cancelados (cliente)', color:'#a8a29a' }
];

/* Indicadores do topo */
const KPIS = [
  { id:'revenue',     label:'Faturamento do mês',      icon:'💰', fmt:'money', wide:true },
  { id:'done',        label:'Atendimentos realizados', icon:'✅' },
  { id:'pend',        label:'Pendentes',               icon:'⏳' },
  { id:'wait',        label:'Aguardando sinal',        icon:'💳' },
  { id:'today',       label:'Confirmados hoje',        icon:'📅' },
  { id:'resched',     label:'Reagendamentos',          icon:'🔁' },
  { id:'cancelCarol', label:'Cancelamentos (Carol)',   icon:'🚫' }
];

const $ = id => document.getElementById(id);
let all = [];                                   // todos os agendamentos (atualizado em tempo real)
let view = { y: new Date().getFullYear(), m: new Date().getMonth() };
let charts = {}, svcRev = [], unsub = null, ready = false;

/* ===== Datas ===== */
const ymKey = (y, m) => `${y}-${String(m + 1).padStart(2, '0')}`;
function shift(v, delta) { const d = new Date(v.y, v.m + delta, 1); return { y: d.getFullYear(), m: d.getMonth() }; }

/* =========================================================
   AGREGAÇÃO (função pura — recebe a lista e devolve os números)
   Regras:
   - Faturamento = soma do preço dos agendamentos "concluido"
   - Mês = mês da data do atendimento (campo "data")
   - "Confirmados hoje" ignora o mês selecionado
   - Cancelado pela Carol = status "cancelado" + canceladoPor "carol"
   - Cancelado com reagendamentos > 0 conta como "reagendado"
   ========================================================= */
function aggregate(list, y, m, today) {
  const ym = ymKey(y, m), dim = new Date(y, m + 1, 0).getDate();
  const A = {
    dim, total: 0, revenue: 0, done: 0, pend: 0, wait: 0, today: 0, resched: 0, cancelCarol: 0,
    daily: Array(dim).fill(0),
    weeks: { resched: [0,0,0,0,0], cancel: [0,0,0,0,0] },
    status: Object.fromEntries(STATUS_META.map(s => [s.k, 0])),
    services: {}, monthly: {}
  };
  for (const a of list) {
    if (!a.data) continue;
    const key = a.data.slice(0, 7), price = +a.preco || 0;

    if (a.status === 'concluido') {                       // série histórica (todos os meses)
      const mm = A.monthly[key] || (A.monthly[key] = { rev: 0, n: 0 });
      mm.rev += price; mm.n++;
    }
    if (a.status === 'confirmado' && a.data === today) A.today++;
    if (key !== ym) continue;

    A.total++;
    const day = +a.data.slice(8, 10), wk = Math.min(Math.ceil(day / 7), 5) - 1;
    let sk = a.status;
    if (sk === 'cancelado') {
      sk = a.canceladoPor === 'carol' ? 'cancelado-carol'
         : (a.canceladoPor === 'reagendamento' || a.reagendamentos > 0) ? 'reagendado'
         : 'cancelado-cliente';
    }
    if (sk in A.status) A.status[sk]++;

    const rs = +a.reagendamentos || 0;
    if (rs) { A.resched += rs; A.weeks.resched[wk] += rs; }
    if (sk === 'cancelado-carol') { A.cancelCarol++; A.weeks.cancel[wk]++; }
    if (a.status === 'pendente') A.pend++;
    if (a.status === 'aguardando-sinal') A.wait++;
    if (a.status === 'concluido') {
      A.done++; A.revenue += price;
      if (day >= 1 && day <= dim) A.daily[day - 1] += price;
      const s = A.services[a.servico] || (A.services[a.servico] = { n: 0, rev: 0 });
      s.n++; s.rev += price;
    }
  }
  return A;
}

/* =========================================================
   KPIs — números que "sobem e descem" até o novo valor
   ========================================================= */
function buildKpis() {
  $('kpis').innerHTML = KPIS.map((k, i) => `
    <article class="kpi${k.wide ? ' wide' : ''}" style="--i:${i}">
      <div class="kpi-top"><span class="kpi-icon">${k.icon}</span><span class="kpi-label">${k.label}</span></div>
      <b class="kpi-val" id="k-${k.id}" data-fmt="${k.fmt || 'int'}">0</b>
      <small class="kpi-hint" id="h-${k.id}"></small>
    </article>`).join('');
}

function countTo(el, to) {
  const from = el._cur || 0, money_ = el.dataset.fmt === 'money';
  const show = v => { el._cur = v; el.textContent = money_ ? money(v) : Math.round(v).toLocaleString('pt-BR'); };
  cancelAnimationFrame(el._raf);
  if (REDUCE || from === to) { show(to); return; }
  const t0 = performance.now(), dur = 900;
  const tick = now => {
    const p = Math.min((now - t0) / dur, 1), e = 1 - Math.pow(1 - p, 3);   // easeOutCubic
    show(p < 1 ? from + (to - from) * e : to);
    if (p < 1) el._raf = requestAnimationFrame(tick);
  };
  el._raf = requestAnimationFrame(tick);
}

function setKpi(id, value, hint) { countTo($('k-' + id), value); $('h-' + id).innerHTML = hint || ''; }

function deltaHTML(cur, prev) {
  if (!prev) return cur ? '<span class="delta flat">sem mês anterior para comparar</span>' : '';
  const p = (cur - prev) / prev * 100;
  if (Math.abs(p) < 0.5) return '<span class="delta flat">= igual ao mês anterior</span>';
  return `<span class="delta ${p > 0 ? 'up' : 'down'}">${p > 0 ? '▲' : '▼'} ${Math.abs(p).toFixed(0)}% vs mês anterior</span>`;
}

/* =========================================================
   GRÁFICOS (Chart.js) — criados uma vez, só os dados mudam,
   por isso as barras sobem e descem a cada atualização.
   ========================================================= */
const kfmt = v => 'R$ ' + (v >= 1000 ? (v / 1000).toFixed(1).replace('.0', '') + 'k' : v);
const stagger = ms => ({ delay: c => (c.type === 'data' && c.mode === 'default') ? c.dataIndex * ms : 0 });
const noGrid = { grid: { display: false }, border: { display: false } };
const gridY = { beginAtZero: true, grid: { color: '#00000010' }, border: { display: false } };
const grad = (top, bottom, horizontal) => ctx => {
  const a = ctx.chart.chartArea; if (!a) return top;
  const g = horizontal ? ctx.chart.ctx.createLinearGradient(a.left, 0, a.right, 0)
                       : ctx.chart.ctx.createLinearGradient(0, a.bottom, 0, a.top);
  g.addColorStop(0, bottom); g.addColorStop(1, top); return g;
};

/* Total no centro da rosca */
const centerPlugin = {
  id: 'center',
  afterDraw(c) {
    const a = c.chartArea, g = c.ctx, x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
    const total = c.data.datasets[0].data.reduce((s, v) => s + v, 0);
    g.save(); g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillStyle = C.ink; g.font = "700 38px 'Cormorant Garamond',serif"; g.fillText(total, x, y - 6);
    g.fillStyle = '#737373'; g.font = "600 10px 'Poppins',sans-serif"; g.fillText('SOLICITAÇÕES', x, y + 20);
    g.restore();
  }
};

function setupCharts() {
  const d = Chart.defaults;
  d.font.family = "'Poppins',system-ui,sans-serif"; d.font.size = 12; d.color = '#737373';
  d.animation.duration = REDUCE ? 0 : 1100; d.animation.easing = 'easeOutQuart';
  d.plugins.legend.labels.usePointStyle = true; d.plugins.legend.labels.boxWidth = 8;
  Object.assign(d.plugins.tooltip, { backgroundColor: C.ink, padding: 12, cornerRadius: 10, boxPadding: 4 });
  const base = { responsive: true, maintainAspectRatio: false };

  charts.daily = new Chart($('cDaily'), {
    data: { labels: [], datasets: [
      { type: 'bar',  label: 'Faturamento do dia', data: [], backgroundColor: grad(C.ink2, '#9a9a9a'), borderRadius: 6, borderSkipped: false, maxBarThickness: 22, order: 2, yAxisID: 'y' },
      { type: 'line', label: 'Acumulado do mês', data: [], borderColor: C.gold, backgroundColor: '#b08a4a1f', fill: true, tension: .35, pointRadius: 0, pointHoverRadius: 5, borderWidth: 2.5, order: 1, yAxisID: 'y1' }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false },
      animation: stagger(14),
      plugins: { legend: { position: 'top', align: 'end' },
        tooltip: { callbacks: { title: i => 'Dia ' + i[0].label, label: c => ` ${c.dataset.label}: ${money(c.parsed.y)}` } } },
      scales: { x: noGrid, y: { ...gridY, ticks: { callback: kfmt } },
        y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, border: { display: false }, ticks: { callback: kfmt } } } }
  });

  charts.status = new Chart($('cStatus'), {
    type: 'doughnut',
    data: { labels: STATUS_META.map(s => s.label), datasets: [{ data: [], backgroundColor: STATUS_META.map(s => s.color), borderColor: '#fff', borderWidth: 3, hoverOffset: 8 }] },
    options: { ...base, cutout: '70%', animation: { animateRotate: true, animateScale: true },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.label}: ${c.parsed}` } } } },
    plugins: [centerPlugin]
  });

  charts.trend = new Chart($('cTrend'), {
    data: { labels: [], datasets: [
      { type: 'bar',  label: 'Faturamento', data: [], backgroundColor: grad(C.ink2, '#9a9a9a'), borderRadius: 8, borderSkipped: false, maxBarThickness: 34, order: 2, yAxisID: 'y' },
      { type: 'line', label: 'Atendimentos', data: [], borderColor: C.gold, backgroundColor: C.gold, tension: .35, pointRadius: 4, pointHoverRadius: 6, borderWidth: 2.5, order: 1, yAxisID: 'y1' }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false }, animation: stagger(90),
      plugins: { legend: { position: 'top', align: 'end' },
        tooltip: { callbacks: { label: c => c.datasetIndex === 0 ? ` Faturamento: ${money(c.parsed.y)}` : ` Atendimentos: ${c.parsed.y}` } } },
      scales: { x: noGrid, y: { ...gridY, ticks: { callback: kfmt } },
        y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, border: { display: false }, ticks: { precision: 0 } } } }
  });

  charts.services = new Chart($('cServices'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Atendimentos', data: [], backgroundColor: grad(C.ink2, '#9a9a9a', true), borderRadius: 8, borderSkipped: false, maxBarThickness: 26 }] },
    options: { ...base, indexAxis: 'y', animation: stagger(90),
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: c => ` ${c.parsed.x} atendimento${c.parsed.x === 1 ? '' : 's'}`, afterLabel: c => ` Receita: ${money(svcRev[c.dataIndex] || 0)}` } } },
      scales: { x: { ...gridY, ticks: { precision: 0 } }, y: noGrid } }
  });

  charts.occ = new Chart($('cOcc'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Reagendamentos',        data: [], backgroundColor: C.violet, borderRadius: 6, borderSkipped: false, maxBarThickness: 24 },
      { label: 'Cancelamentos (Carol)', data: [], backgroundColor: C.wine,   borderRadius: 6, borderSkipped: false, maxBarThickness: 24 }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false }, animation: stagger(90),
      plugins: { legend: { position: 'top', align: 'end' }, tooltip: { callbacks: { title: i => 'Dias ' + i[0].label } } },
      scales: { x: noGrid, y: { ...gridY, ticks: { precision: 0 } } } }
  });
}

/* =========================================================
   RENDER — recalcula tudo e atualiza (anima) os gráficos
   ========================================================= */
function render() {
  const today = keyDate(new Date());
  const A = aggregate(all, view.y, view.m, today);
  const prevV = shift(view, -1);
  const prev = A.monthly[ymKey(prevV.y, prevV.m)] || { rev: 0, n: 0 };
  const isCurrent = view.y === new Date().getFullYear() && view.m === new Date().getMonth();
  const todayDay = +today.slice(8, 10);
  const plural = (n, s, p) => n === 1 ? s : p;

  $('mLabel').textContent = `${MESES[view.m]} ${view.y}`;

  /* KPIs */
  setKpi('revenue',     A.revenue,     deltaHTML(A.revenue, prev.rev));
  setKpi('done',        A.done,        deltaHTML(A.done, prev.n));
  setKpi('pend',        A.pend,        `de ${A.total} ${plural(A.total, 'solicitação', 'solicitações')} no mês`);
  setKpi('wait',        A.wait,        'pagamento do sinal em aberto');
  setKpi('today',       A.today,       `hoje · ${br(today)}`);
  setKpi('resched',     A.resched,     'no mês selecionado');
  setKpi('cancelCarol', A.cancelCarol, `${A.status.recusado} ${plural(A.status.recusado, 'recusa', 'recusas')} à parte`);

  /* Faturamento por dia + acumulado (linha para no dia de hoje no mês atual) */
  let run = 0;
  charts.daily.data.labels = A.daily.map((_, i) => String(i + 1));
  charts.daily.data.datasets[0].data = A.daily;
  charts.daily.data.datasets[1].data = A.daily.map((v, i) => { run += v; return (isCurrent && i + 1 > todayDay) ? null : run; });
  $('cDaily').parentNode.dataset.empty = A.revenue ? '' : '1';
  charts.daily.update();

  /* Rosca de status + legenda */
  const vals = STATUS_META.map(s => A.status[s.k]);
  charts.status.data.datasets[0].data = vals;
  $('cStatus').parentNode.dataset.empty = A.total ? '' : '1';
  $('statusLegend').innerHTML = STATUS_META.map((s, i) => `<li><i style="background:${s.color}"></i>${s.label}<b>${vals[i]}</b></li>`).join('');
  charts.status.update();

  /* Últimos 6 meses */
  const months = [5, 4, 3, 2, 1, 0].map(i => shift(view, -i));
  const series = months.map(x => A.monthly[ymKey(x.y, x.m)] || { rev: 0, n: 0 });
  charts.trend.data.labels = months.map(x => `${MESES_CURTO[x.m]}/${String(x.y).slice(2)}`);
  charts.trend.data.datasets[0].data = series.map(s => s.rev);
  charts.trend.data.datasets[1].data = series.map(s => s.n);
  $('cTrend').parentNode.dataset.empty = series.some(s => s.n) ? '' : '1';
  charts.trend.update();

  /* Serviços mais realizados (top 6) */
  const svc = Object.entries(A.services).sort((a, b) => b[1].n - a[1].n || b[1].rev - a[1].rev).slice(0, 6);
  svcRev = svc.map(([, v]) => v.rev);
  charts.services.data.labels = svc.map(([n]) => n);
  charts.services.data.datasets[0].data = svc.map(([, v]) => v.n);
  $('cServices').parentNode.dataset.empty = svc.length ? '' : '1';
  charts.services.update();

  /* Reagendamentos e cancelamentos por semana */
  const nW = A.dim > 28 ? 5 : 4;
  charts.occ.data.labels = ['1–7', '8–14', '15–21', '22–28', `29–${A.dim}`].slice(0, nW);
  charts.occ.data.datasets[0].data = A.weeks.resched.slice(0, nW);
  charts.occ.data.datasets[1].data = A.weeks.cancel.slice(0, nW);
  $('cOcc').parentNode.dataset.empty = (A.resched + A.cancelCarol) ? '' : '1';
  charts.occ.update();
}

/* =========================================================
   TEMPO REAL (Firestore onSnapshot) + acesso restrito
   ========================================================= */
function setLive(on, text) {
  $('live').classList.toggle('off', !on);
  $('liveText').textContent = text;
}

function startRealtime() {
  if (unsub) return;
  unsub = db.collection('agendamentos').onSnapshot(snap => {
    all = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    ready = true;
    $('dash').classList.remove('is-loading');
    render();
    setLive(!snap.metadata.fromCache, snap.metadata.fromCache
      ? 'Sem conexão — exibindo dados salvos'
      : 'Ao vivo · atualizado às ' + new Date().toLocaleTimeString('pt-BR'));
  }, err => {
    console.error(err);
    setLive(false, 'Conexão perdida');
    toast('Não foi possível receber atualizações em tempo real.');
  });
}

document.addEventListener('DOMContentLoaded', () => {
  if (!$('dash')) return;
  if (typeof Chart === 'undefined') { toast('Não foi possível carregar a biblioteca de gráficos.'); return; }

  buildKpis();
  setupCharts();
  $('mLabel').textContent = `${MESES[view.m]} ${view.y}`;

  const go = d => { view = d === 0 ? { y: new Date().getFullYear(), m: new Date().getMonth() } : shift(view, d); if (ready) render(); };
  $('mPrev').addEventListener('click', () => go(-1));
  $('mNext').addEventListener('click', () => go(1));
  $('mToday').addEventListener('click', () => go(0));

  auth.onAuthStateChanged(async user => {
    if (!user) { location.replace('admin.html'); return; }
    try {
      const adm = await db.collection('admins').doc(user.uid).get();
      if (!adm.exists) { await auth.signOut(); location.replace('admin.html'); return; }
      startRealtime();
    } catch (e) { console.error(e); toast('Erro ao validar o acesso.'); }
  });
});

window.addEventListener('pagehide', () => { if (unsub) unsub(); });
