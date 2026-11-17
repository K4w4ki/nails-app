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

/* Indicadores do topo (ícones SVG de traço — ver ICON) */
const KPIS = [
  { id:'revenue',     label:'Faturamento do mês',      icon:'money', fmt:'money', wide:true },
  { id:'done',        label:'Atendimentos realizados', icon:'done' },
  { id:'pend',        label:'Pendentes',               icon:'clock' },
  { id:'wait',        label:'Aguardando sinal',        icon:'card' },
  { id:'today',       label:'Confirmados hoje',        icon:'today' },
  { id:'resched',     label:'Reagendamentos',          icon:'swap' },
  { id:'cancelCarol', label:'Cancelamentos (Carol)',   icon:'ban' }
];

/* Ícones de traço 24×24 (estilo Lucide) */
const ICON = {
  money: '<rect width="20" height="12" x="2" y="6" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01M18 12h.01"/>',
  done:  '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
  clock: '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>',
  card:  '<rect width="20" height="14" x="2" y="5" rx="2"/><path d="M2 10h20M6 15h4"/>',
  today: '<path d="M8 2v4M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18M9 16l2 2 4-4"/>',
  swap:  '<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>',
  ban:   '<circle cx="12" cy="12" r="10"/><path d="m4.9 4.9 14.2 14.2"/>',
  up:    '<path d="M7 17 17 7M7 7h10v10"/>',
  down:  '<path d="m7 7 10 10M17 7v10H7"/>',
  flat:  '<path d="M5 12h14"/>',
  chart: '<path d="M3 3v18h18M18 17V9M13 17V5M8 17v-3"/>'
};
const svg = (name, size = 18, w = 1.75) =>
  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[name]}</svg>`;


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
      <div class="kpi-top"><span class="kpi-label">${k.label}</span><span class="kpi-icon">${svg(k.icon, 18)}</span></div>
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
  const note = t => `<span class="delta-note">${t}</span>`;
  if (!prev) return cur ? note('sem mês anterior para comparar') : '';
  const p = (cur - prev) / prev * 100;
  if (Math.abs(p) < 0.5) return `<span class="delta flat">${svg('flat', 12, 2.25)}0%</span>${note('igual ao mês anterior')}`;
  const up = p > 0;
  return `<span class="delta ${up ? 'up' : 'down'}">${svg(up ? 'up' : 'down', 12, 2.25)}${Math.abs(p).toFixed(0)}%</span>${note('vs mês anterior')}`;
}

/* =========================================================
   GRÁFICOS (Chart.js) — criados uma vez; só os dados mudam,
   por isso as barras sobem e descem a cada atualização.
   ========================================================= */
const GRID = '#e9e5df', MUTED = '#8a847c', GOLD_TXT = '#8a6531';
let trendTitles = [];

const kfmt = v => 'R$ ' + (v >= 1000 ? (v / 1000).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' mil' : v.toLocaleString('pt-BR'));
const stagger = ms => ({ delay: c => (c.type === 'data' && c.mode === 'default') ? c.dataIndex * ms : 0 });
const nz = v => v || null;                       // zero vira "sem barra" (nada de traço sobre o eixo)

/* Escala de 4 intervalos "redondos": eixo esquerdo e direito ficam com as mesmas linhas de grade */
function niceScale(max, int = false, n = 4) {
  const need = Math.max(max, n) / n, p = Math.pow(10, Math.floor(Math.log10(need)));
  const step = (int ? [1, 2, 5, 10] : [1, 2, 2.5, 5, 10]).map(f => f * p).find(s => s >= need * 0.9999) || 10 * p;
  return { step, max: step * n };
}
const fitAxis = (ax, s) => { ax.max = s.max; ax.ticks.stepSize = s.step; };

const axisX = (ticks = {}) => ({
  grid: { display: false }, border: { display: false },
  ticks: { color: MUTED, maxRotation: 0, minRotation: 0, autoSkip: false, padding: 8, ...ticks }
});
const axisY = (opts = {}, ticks = {}) => ({
  min: 0, border: { display: false },
  grid: { color: c => c.tick.value === 0 ? '#d6d0c8' : GRID, drawTicks: false },
  ticks: { color: MUTED, padding: 10, ...ticks }, ...opts
});
const barStyle = (color, r = 4) => ({ backgroundColor: color, borderRadius: { topLeft: r, topRight: r }, borderSkipped: 'bottom' });
const areaGrad = chart => {
  const a = chart.chartArea; if (!a) return 'transparent';
  const g = chart.ctx.createLinearGradient(0, a.top, 0, a.bottom);
  g.addColorStop(0, 'rgba(176,138,74,.22)'); g.addColorStop(1, 'rgba(176,138,74,0)'); return g;
};

/* Total no centro da rosca */
const centerPlugin = {
  id: 'center',
  afterDraw(c) {
    const a = c.chartArea, g = c.ctx, x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
    const total = c.data.datasets[0].data.reduce((s, v) => s + v, 0);
    g.save(); g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillStyle = C.ink; g.font = "600 32px 'Poppins',sans-serif"; g.fillText(total, x, y - 8);
    g.fillStyle = MUTED; g.font = "500 11.5px 'Poppins',sans-serif"; g.fillText(total === 1 ? 'solicitação' : 'solicitações', x, y + 18);
    g.restore();
  }
};

/* Valor ao fim de cada barra horizontal */
const valueLabels = {
  id: 'valueLabels',
  afterDatasetsDraw(c) {
    const g = c.ctx, data = c.data.datasets[0].data;
    g.save(); g.fillStyle = C.ink; g.font = "600 12px 'Poppins',sans-serif"; g.textAlign = 'left'; g.textBaseline = 'middle';
    c.getDatasetMeta(0).data.forEach((bar, i) => { if (data[i] != null) g.fillText(data[i], bar.x + 8, bar.y); });
    g.restore();
  }
};

/* Linha vertical fina no ponto sob o cursor */
const crosshair = {
  id: 'crosshair',
  beforeDatasetsDraw(c) {
    const t = c.tooltip;
    if (!t || !t.opacity || !t.dataPoints || !t.dataPoints.length) return;
    const x = t.dataPoints[0].element.x, a = c.chartArea, g = c.ctx;
    g.save(); g.strokeStyle = '#28282822'; g.lineWidth = 1; g.beginPath(); g.moveTo(x, a.top); g.lineTo(x, a.bottom); g.stroke(); g.restore();
  }
};

/* Legenda em HTML (clicável) no lugar da legenda do Chart.js */
function addLegend(chart, canvasId, items) {
  const el = document.createElement('div');
  el.className = 'chart-legend';
  el.innerHTML = items.map((it, i) => `<button type="button" class="lg" data-i="${i}" aria-pressed="true"><i class="sw ${it.type}" style="--c:${it.color}"></i>${it.label}</button>`).join('');
  $(canvasId).parentNode.before(el);
  el.addEventListener('click', e => {
    const b = e.target.closest('.lg'); if (!b) return;
    const i = +b.dataset.i, on = !chart.isDatasetVisible(i);
    chart.setDatasetVisibility(i, on);
    b.classList.toggle('off', !on); b.setAttribute('aria-pressed', on);
    chart.update();
  });
}

/* Estado vazio: o gráfico some e entra uma mensagem (sem véu sobre os eixos) */
const EMPTY = {
  cDaily:    'Nenhum atendimento concluído neste mês',
  cStatus:   'Nenhuma solicitação neste mês',
  cTrend:    'Nenhum atendimento concluído no período',
  cServices: 'Nenhum serviço concluído neste mês',
  cOcc:      'Nenhum reagendamento ou cancelamento neste mês'
};
const setEmpty = (id, on) => { $(id).closest('.dash-card').dataset.empty = on ? '1' : ''; };

function setupCharts() {
  const d = Chart.defaults;
  d.font.family = "'Poppins',system-ui,sans-serif"; d.font.size = 11.5; d.color = MUTED;
  d.animation.duration = REDUCE ? 0 : 900; d.animation.easing = 'easeOutQuart';
  d.plugins.legend.display = false;
  Object.assign(d.plugins.tooltip, {
    backgroundColor: C.ink, titleFont: { weight: 600, size: 12 }, bodyFont: { size: 12 },
    padding: { top: 10, right: 12, bottom: 10, left: 12 }, cornerRadius: 10, boxPadding: 5,
    usePointStyle: true, caretSize: 0, caretPadding: 10
  });
  const base = { responsive: true, maintainAspectRatio: false, layout: { padding: { top: 6, right: 2 } } };
  const right = { position: 'right', grid: { drawOnChartArea: false, drawTicks: false } };

  charts.daily = new Chart($('cDaily'), {
    data: { labels: [], datasets: [
      { type: 'bar', label: 'Faturamento do dia', data: [], ...barStyle(C.ink2), maxBarThickness: 16, categoryPercentage: .9, barPercentage: .8, order: 2, yAxisID: 'y' },
      { type: 'line', label: 'Acumulado do mês', data: [], borderColor: C.gold, borderWidth: 2, backgroundColor: c => areaGrad(c.chart), fill: 'origin',
        cubicInterpolationMode: 'monotone', pointRadius: 0, pointHoverRadius: 5, pointHoverBackgroundColor: '#fff', pointHoverBorderColor: C.gold, pointHoverBorderWidth: 2, order: 1, yAxisID: 'y1' }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false }, animation: stagger(14),
      plugins: { tooltip: { callbacks: {
        title: i => `${i[0].label} de ${MESES[view.m].toLowerCase()}`,
        label: c => ` ${c.dataset.label}: ${money(c.parsed.y)}` } } },
      scales: {
        x: axisX({ callback(v) { const n = +this.getLabelForValue(v); return (n === 1 || n % 5 === 0) ? String(n) : ''; } }),
        y: axisY({}, { callback: kfmt }),
        y1: axisY(right, { callback: kfmt, color: GOLD_TXT })
      } },
    plugins: [crosshair]
  });

  charts.status = new Chart($('cStatus'), {
    type: 'doughnut',
    data: { labels: STATUS_META.map(s => s.label), datasets: [{ data: [], backgroundColor: STATUS_META.map(s => s.color), borderColor: '#fff', borderWidth: 2, hoverOffset: 4 }] },
    options: { ...base, layout: { padding: 8 }, cutout: '74%',
      plugins: { tooltip: { callbacks: { label: c => ` ${c.label}: ${c.parsed}` } } } },
    plugins: [centerPlugin]
  });

  charts.trend = new Chart($('cTrend'), {
    data: { labels: [], datasets: [
      { type: 'bar', label: 'Faturamento', data: [], ...barStyle(C.ink2), maxBarThickness: 28, categoryPercentage: .7, barPercentage: .9, order: 2, yAxisID: 'y' },
      { type: 'line', label: 'Atendimentos', data: [], borderColor: C.gold, borderWidth: 2, cubicInterpolationMode: 'monotone', pointRadius: 4, pointHoverRadius: 6,
        pointBackgroundColor: '#fff', pointBorderColor: C.gold, pointBorderWidth: 2, order: 1, yAxisID: 'y1' }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false }, animation: stagger(90),
      plugins: { tooltip: { callbacks: {
        title: i => trendTitles[i[0].dataIndex] || i[0].label,
        label: c => c.datasetIndex === 0 ? ` Faturamento: ${money(c.parsed.y)}` : ` Atendimentos: ${c.parsed.y}` } } },
      scales: {
        x: axisX({ font: { size: 11 } }),
        y: axisY({}, { callback: kfmt }),
        y1: axisY(right, { precision: 0, color: GOLD_TXT })
      } },
    plugins: [crosshair]
  });

  charts.services = new Chart($('cServices'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Atendimentos', data: [], backgroundColor: c => c.dataIndex === 0 ? C.gold : C.ink2,
      borderRadius: { topRight: 4, bottomRight: 4 }, borderSkipped: 'start', barThickness: 14 }] },
    options: { ...base, indexAxis: 'y', animation: stagger(80), layout: { padding: { top: 4, right: 34 } },
      plugins: { tooltip: { displayColors: false, callbacks: {
        label: c => ` ${c.parsed.x} atendimento${c.parsed.x === 1 ? '' : 's'}`,
        afterLabel: c => ` Receita: ${money(svcRev[c.dataIndex] || 0)}` } } },
      scales: {
        x: { display: false, min: 0 },
        y: { grid: { display: false }, border: { display: false },
          ticks: { color: C.ink, font: { weight: 500 }, padding: 8,
            callback(v) { const s = this.getLabelForValue(v); return s.length > 22 ? s.slice(0, 21) + '…' : s; } } }
      } },
    plugins: [valueLabels]
  });

  charts.occ = new Chart($('cOcc'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Reagendamentos',        data: [], ...barStyle(C.violet), maxBarThickness: 18 },
      { label: 'Cancelamentos (Carol)', data: [], ...barStyle(C.wine),   maxBarThickness: 18 }
    ]},
    options: { ...base, interaction: { mode: 'index', intersect: false }, animation: stagger(80),
      plugins: { tooltip: { callbacks: { title: i => 'Dias ' + i[0].label } } },
      scales: { x: axisX(), y: axisY({}, { precision: 0 }) } }
  });

  addLegend(charts.daily, 'cDaily', [{ label: 'Faturamento do dia', color: C.ink2, type: 'bar' }, { label: 'Acumulado do mês', color: C.gold, type: 'line' }]);
  addLegend(charts.trend, 'cTrend', [{ label: 'Faturamento', color: C.ink2, type: 'bar' }, { label: 'Atendimentos', color: C.gold, type: 'line' }]);
  addLegend(charts.occ, 'cOcc', [{ label: 'Reagendamentos', color: C.violet, type: 'bar' }, { label: 'Cancelamentos (Carol)', color: C.wine, type: 'bar' }]);

  Object.entries(EMPTY).forEach(([id, msg]) => $(id).parentNode.insertAdjacentHTML('beforeend',
    `<div class="empty-state">${svg('chart', 22, 1.5)}<span>${msg}</span></div>`));

  /* A rosca desenha texto no canvas: redesenha quando a fonte terminar de carregar */
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => Object.values(charts).forEach(c => c.update('none')));
}

/* =========================================================
   RENDER — recalcula tudo e atualiza (anima) os gráficos
   ========================================================= */
function render() {
  const today = keyDate(new Date());
  const A = aggregate(all, view.y, view.m, today);
  const prevV = shift(view, -1);
  const prev = A.monthly[ymKey(prevV.y, prevV.m)] || { rev: 0, n: 0 };
  const now = new Date(), isCurrent = view.y === now.getFullYear() && view.m === now.getMonth();
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

  /* Faturamento por dia + acumulado (a linha para no dia de hoje no mês atual) */
  let run = 0;
  const dc = charts.daily;
  dc.data.labels = A.daily.map((_, i) => String(i + 1));
  dc.data.datasets[0].data = A.daily.map(nz);
  dc.data.datasets[1].data = A.daily.map((v, i) => { run += v; return (isCurrent && i + 1 > todayDay) ? null : run; });
  fitAxis(dc.options.scales.y, niceScale(Math.max(...A.daily)));
  fitAxis(dc.options.scales.y1, niceScale(A.revenue));
  setEmpty('cDaily', !A.revenue);
  dc.update();

  /* Rosca de status + legenda */
  const vals = STATUS_META.map(s => A.status[s.k]);
  charts.status.data.datasets[0].data = vals;
  setEmpty('cStatus', !vals.some(Boolean));
  $('statusLegend').innerHTML = STATUS_META.map((s, i) =>
    `<li class="${vals[i] ? '' : 'zero'}"><i style="background:${s.color}"></i>${s.label}<b>${vals[i]}</b></li>`).join('');
  charts.status.update();

  /* Últimos 6 meses */
  const months = [5, 4, 3, 2, 1, 0].map(i => shift(view, -i));
  const series = months.map(x => A.monthly[ymKey(x.y, x.m)] || { rev: 0, n: 0 });
  const tc = charts.trend;
  tc.data.labels = months.map((x, i) => MESES_CURTO[x.m] + (i === 0 || x.m === 0 ? ` ’${String(x.y).slice(2)}` : ''));
  trendTitles = months.map(x => `${MESES[x.m]} de ${x.y}`);
  tc.data.datasets[0].data = series.map(s => s.rev || null);
  tc.data.datasets[1].data = series.map(s => s.n);
  fitAxis(tc.options.scales.y, niceScale(Math.max(...series.map(s => s.rev))));
  fitAxis(tc.options.scales.y1, niceScale(Math.max(...series.map(s => s.n)), true));
  setEmpty('cTrend', !series.some(s => s.n));
  tc.update();

  /* Serviços mais realizados (top 6) */
  const svc = Object.entries(A.services).sort((a, b) => b[1].n - a[1].n || b[1].rev - a[1].rev).slice(0, 6);
  svcRev = svc.map(([, v]) => v.rev);
  const sc = charts.services;
  sc.data.labels = svc.map(([n]) => n);
  sc.data.datasets[0].data = svc.map(([, v]) => v.n);
  sc.options.scales.x.max = Math.max(3, svc.length ? svc[0][1].n : 0);
  setEmpty('cServices', !svc.length);
  sc.update();

  /* Reagendamentos e cancelamentos por semana */
  const nW = A.dim > 28 ? 5 : 4, rw = A.weeks.resched.slice(0, nW), cw = A.weeks.cancel.slice(0, nW);
  const oc = charts.occ;
  oc.data.labels = ['1–7', '8–14', '15–21', '22–28', `29–${A.dim}`].slice(0, nW);
  oc.data.datasets[0].data = rw.map(nz);
  oc.data.datasets[1].data = cw.map(nz);
  fitAxis(oc.options.scales.y, niceScale(Math.max(...rw, ...cw), true));
  setEmpty('cOcc', !(A.resched + A.cancelCarol));
  oc.update();
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
