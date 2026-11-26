'use strict';
/* =========================================================
   EXPORTAR DASHBOARD PARA EXCEL — Área da Carol
   Inclui IMAGENS dos gráficos (SVG -> PNG) e tabelas detalhadas.
   ========================================================= */
(function () {
  const EXCELJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';

  /* Paleta do site (styles.css) em formato ARGB */
  const X = {
    ink: 'FF282828', ink2: 'FF3D3D3D', gold: 'FFB08A4A', mist: 'FFF6F6F5', soft: 'FFECECEB',
    line: 'FFDEDEDE', muted: 'FF737373', ok: 'FF46654D', bad: 'FF965154', white: 'FFFFFFFF'
  };
  const FONT = 'Calibri', SERIF = 'Georgia';
  const MONEY = '"R$" #,##0.00';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const argb = hex => 'FF' + hex.replace('#', '').toUpperCase();
  const tint = (hex, a = 0.16) => {
    const h = hex.replace('#', ''), n = i => parseInt(h.slice(i, i + 2), 16);
    const m = v => Math.round(255 - (255 - v) * a).toString(16).padStart(2, '0');
    return ('FF' + m(n(0)) + m(n(2)) + m(n(4))).toUpperCase();
  };
  const utcDate = k => new Date(Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10)));

  function plainDelta(cur, prev, t) {
    if (!prev) return cur ? t.none : '';
    const p = (cur - prev) / prev * 100;
    if (Math.abs(p) < 0.5) return '0% · ' + t.same;
    return `${p > 0 ? '▲' : '▼'} ${Math.abs(p).toFixed(0)}% ${t.vs}`;
  }
  const MES_T = { none: 'sem mês anterior para comparar', same: 'igual ao mês anterior', vs: 'vs mês anterior' };
  const SEM_T = { none: 'sem semana anterior para comparar', same: 'igual à semana anterior', vs: 'vs semana anterior' };

  function statusKey(a) {
    let sk = a.status;
    if (sk === 'cancelado') {
      sk = a.canceladoPor === 'carol' ? 'cancelado-carol'
        : (a.canceladoPor === 'reagendamento' || a.reagendamentos > 0) ? 'reagendado'
        : 'cancelado-cliente';
    }
    return sk;
  }

  /* =======================================================
     Helper: Transforma um SVG da tela em Imagem Base64
     ======================================================= */
  async function svgToPngBase64(containerId) {
    return new Promise((resolve) => {
      const container = document.getElementById(containerId);
      if (!container) return resolve(null);
      const svg = container.querySelector('svg');
      if (!svg) return resolve(null);

      const clone = svg.cloneNode(true);
      // Define tamanhos explícitos para não perder qualidade
      const w = svg.clientWidth || 800;
      const h = svg.clientHeight || 300;
      clone.setAttribute('width', w);
      clone.setAttribute('height', h);
      
      // Fundo branco para o Excel não ficar transparente
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.setAttribute('width', '100%');
      rect.setAttribute('height', '100%');
      rect.setAttribute('fill', '#ffffff');
      clone.insertBefore(rect, clone.firstChild);

      const svgString = new XMLSerializer().serializeToString(clone);
      const svg64 = btoa(unescape(encodeURIComponent(svgString)));
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        // Resolução 2x para ficar nítido
        canvas.width = w * 2;
        canvas.height = h * 2;
        const ctx = canvas.getContext('2d');
        ctx.scale(2, 2);
        ctx.drawImage(img, 0, 0, w, h);
        resolve({ base64: canvas.toDataURL('image/png'), width: w, height: h });
      };
      img.onerror = () => resolve(null);
      img.src = 'data:image/svg+xml;base64,' + svg64;
    });
  }

  /* =======================================================
     Monta a planilha.
     ======================================================= */
  async function buildWorkbook(ExcelJS, ctx, imgs) {
    const { all, view, A, today, now } = ctx;
    const U = window.DashUtils;
    const { MESES, STATUS_META, ymKey, shift, sumWeek, weekRangeOffset } = U;

    const mes = MESES[view.m], ym = ymKey(view.y, view.m);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Nails by Nascimento';
    wb.created = now;
    wb.title = `Dashboard — ${mes} ${view.y}`;

    /* ---------- Aba 1: Dashboard ---------- */
    const ws = wb.addWorksheet('Dashboard', {
      views: [{ showGridLines: false }],
      properties: { tabColor: { argb: X.gold } },
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .4, right: .4, top: .5, bottom: .5, header: .2, footer: .2 } }
    });
    ws.columns = [{ width: 3 }, { width: 36 }, { width: 20 }, { width: 20 }, { width: 20 }];
    let r = 2;

    const mergeBE = (row) => ws.mergeCells(`B${row}:E${row}`);

    /* Faixa de marca */
    mergeBE(r);
    Object.assign(ws.getCell(r, 2), {
      value: 'Nails by Nascimento  ·  Carol Nascimento',
      font: { name: SERIF, size: 12, bold: true, color: { argb: X.white } },
      fill: solid(X.ink), alignment: { vertical: 'middle', indent: 1 }
    });
    for (let c = 3; c <= 5; c++) ws.getCell(r, c).fill = solid(X.ink);
    ws.getRow(r).height = 30; r += 2;

    Object.assign(ws.getCell(r, 2), { value: 'DASHBOARD', font: { name: FONT, size: 9, bold: true, color: { argb: X.gold } } });
    r++; mergeBE(r);
    Object.assign(ws.getCell(r, 2), { value: 'Visão geral do estúdio', font: { name: SERIF, size: 24, bold: true, color: { argb: X.ink } }, alignment: { vertical: 'middle' } });
    ws.getRow(r).height = 36; r++; mergeBE(r);
    Object.assign(ws.getCell(r, 2), {
      value: `${mes} ${view.y}  ·  gerado em ${now.toLocaleDateString('pt-BR')} às ${now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`,
      font: { name: FONT, size: 10, color: { argb: X.muted } }
    });
    r += 2;

    const section = (eyebrow, title, note) => {
      Object.assign(ws.getCell(r, 2), { value: eyebrow.toUpperCase(), font: { name: FONT, size: 9, bold: true, color: { argb: X.gold } } });
      r++; mergeBE(r);
      Object.assign(ws.getCell(r, 2), { value: title, font: { name: SERIF, size: 17, bold: true, color: { argb: X.ink } }, alignment: { vertical: 'middle' } });
      ws.getRow(r).height = 26; r++;
      if (note) {
        mergeBE(r);
        Object.assign(ws.getCell(r, 2), { value: note, font: { name: FONT, size: 9, italic: true, color: { argb: X.muted } } });
        r++;
      }
    };

    const tr = (vals, o = {}) => {
      const kind = o.kind || 'body', n = o.span3 ? 4 : vals.length;
      const bg = kind === 'head' || kind === 'hero' ? X.ink : kind === 'total' ? X.mist : null;
      const fg = kind === 'head' || kind === 'hero' ? X.white : X.ink;
      for (let i = 0; i < n; i++) {
        const c = ws.getCell(r, 2 + i), v = vals[i];
        if (v !== undefined) c.value = v;
        const isText = i === 0 || (o.span3 && i >= 2);
        c.font = {
          name: FONT, size: kind === 'head' ? 10 : 10.5,
          bold: kind !== 'body' || (!!o.boldFirst && i === 0),
          color: { argb: (!isText && kind === 'body' && o.muted) ? X.muted : (isText && i > 0 && kind === 'body') ? X.muted : fg }
        };
        c.alignment = { vertical: 'middle', horizontal: isText ? 'left' : 'right', indent: 1 };
        if (bg) c.fill = solid(bg);
        if (o.fmts && o.fmts[i]) c.numFmt = o.fmts[i];
        c.border = kind === 'total'
          ? { top: { style: 'thin', color: { argb: X.ink } }, bottom: { style: 'thin', color: { argb: X.line } } }
          : { bottom: { style: 'thin', color: { argb: X.line } } };
      }
      if (o.span3) ws.mergeCells(`D${r}:E${r}`);
      if (o.swatch) ws.getCell(r, 1).fill = solid(o.swatch);
      ws.getRow(r).height = kind === 'hero' ? 30 : 21;
      r++;
    };
    const bars = (ref, color) => ws.addConditionalFormatting({
      ref, rules: [{ type: 'dataBar', priority: 1, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: color } }]
    });
    const gap = () => { r += 2; };

    /* --- Resumo da semana --- */
    const wCur = weekRangeOffset(0), wPrev = weekRangeOffset(-1);
    const sCur = sumWeek(all, wCur.start, wCur.end), sPrev = sumWeek(all, wPrev.start, wPrev.end);
    section('Resumo da semana', `${br(wCur.start)} – ${br(wCur.end)}`, 'Independe do mês selecionado · de segunda a domingo');
    tr(['Indicador', 'Valor', 'Comparação'], { kind: 'head', span3: true });
    tr(['Faturamento da semana', sCur.revenue, plainDelta(sCur.revenue, sPrev.revenue, SEM_T)], { span3: true, fmts: [null, MONEY], boldFirst: true });
    tr(['Atendimentos concluídos', sCur.done, plainDelta(sCur.done, sPrev.done, SEM_T)], { span3: true, fmts: [null, '0'], boldFirst: true });
    gap();

    /* --- Indicadores do mês --- */
    const prevV = shift(view, -1);
    const prev = A.monthly[ymKey(prevV.y, prevV.m)] || { rev: 0, n: 0 };
    const plural = (n, s, p) => n === 1 ? s : p;
    section('Indicadores', `${mes} ${view.y}`, 'Os mesmos cartões exibidos no topo do dashboard');
    tr(['Indicador', 'Valor', 'Detalhe'], { kind: 'head', span3: true });
    tr(['Faturamento do mês', A.revenue, plainDelta(A.revenue, prev.rev, MES_T)], { kind: 'hero', span3: true, fmts: [null, MONEY] });
    tr(['Atendimentos realizados', A.done, plainDelta(A.done, prev.n, MES_T)], { span3: true, fmts: [null, '0'], boldFirst: true });
    tr(['Pendentes', A.pend, `de ${A.total} ${plural(A.total, 'solicitação', 'solicitações')} no mês`], { span3: true, fmts: [null, '0'], boldFirst: true });
    tr(['Aguardando sinal', A.wait, 'pagamento do sinal em aberto'], { span3: true, fmts: [null, '0'], boldFirst: true });
    tr(['Confirmados hoje', A.today, `hoje · ${br(today)}`], { span3: true, fmts: [null, '0'], boldFirst: true });
    tr(['Reagendamentos', A.resched, 'no mês selecionado'], { span3: true, fmts: [null, '0'], boldFirst: true });
    tr(['Cancelamentos (Carol)', A.cancelCarol, `${A.status.recusado} ${plural(A.status.recusado, 'recusa', 'recusas')} à parte`], { span3: true, fmts: [null, '0'], boldFirst: true });
    gap();

    /* --- Inserir Imagem do Gráfico 1: Faturamento por dia --- */
    if (imgs.cDaily) {
      section('Receita', 'Faturamento por dia', 'Receita dos atendimentos concluídos, dia a dia e acumulada no mês');
      const imgId = wb.addImage({ base64: imgs.cDaily.base64, extension: 'png' });
      ws.addImage(imgId, {
        tl: { col: 1, row: r - 1 },
        ext: { width: 600, height: 300 }
      });
      r += 16; // Pula as linhas que a imagem está ocupando
    }

    /* --- Inserir Imagem do Gráfico 2: Rosca --- */
    if (imgs.cStatus) {
      section('Agenda', 'Situação das solicitações', 'Todas as solicitações do mês, por status');
      const imgId = wb.addImage({ base64: imgs.cStatus.base64, extension: 'png' });
      ws.addImage(imgId, {
        tl: { col: 1, row: r - 1 },
        ext: { width: 350, height: 300 }
      });
      r += 16;
    }

    /* --- Situação das solicitações (Tabela) --- */
    tr(['Status', 'Solicitações', '% do mês'], { kind: 'head' });
    const st0 = r;
    STATUS_META.forEach(s => {
      const v = A.status[s.k];
      tr([s.label, v, A.total ? v / A.total : 0], { fmts: [null, '0', '0%'], swatch: argb(s.color) });
    });
    bars(`C${st0}:C${r - 1}`, 'FFB9B9B9');
    tr(['Total do mês', A.total, A.total ? 1 : 0], { kind: 'total', fmts: [null, '0', '0%'] });
    gap();

    /* --- Inserir Imagem do Gráfico 3: Tendência --- */
    if (imgs.cTrend) {
      section('Tendência', 'Últimos 6 meses', 'Faturamento e atendimentos, mês a mês');
      const imgId = wb.addImage({ base64: imgs.cTrend.base64, extension: 'png' });
      ws.addImage(imgId, {
        tl: { col: 1, row: r - 1 },
        ext: { width: 600, height: 300 }
      });
      r += 16;
    }

    /* --- Inserir Imagem do Gráfico 4: Serviços --- */
    if (imgs.cServices) {
      section('Catálogo', 'Serviços mais realizados', 'Atendimentos concluídos no mês');
      const imgId = wb.addImage({ base64: imgs.cServices.base64, extension: 'png' });
      ws.addImage(imgId, {
        tl: { col: 1, row: r - 1 },
        ext: { width: 600, height: 300 }
      });
      r += 16;
    }

    /* --- Inserir Imagem do Gráfico 5: Ocorrências --- */
    if (imgs.cOcc) {
      section('Ocorrências', 'Reagendamentos e cancelamentos', 'Por semana do mês');
      const imgId = wb.addImage({ base64: imgs.cOcc.base64, extension: 'png' });
      ws.addImage(imgId, {
        tl: { col: 1, row: r - 1 },
        ext: { width: 600, height: 300 }
      });
      r += 16;
    }

    /* ---------- Aba 2: Faturamento por dia ---------- */
    const wd = wb.addWorksheet('Faturamento por dia', {
      views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }],
      properties: { tabColor: { argb: X.ink } }
    });
    wd.columns = [{ width: 3 }, { width: 16 }, { width: 18 }, { width: 22 }, { width: 24 }];
    wd.mergeCells('B1:E1');
    Object.assign(wd.getCell('B1'), { value: `Faturamento por dia — ${mes} ${view.y}`, font: { name: SERIF, size: 18, bold: true, color: { argb: X.ink } }, alignment: { vertical: 'middle' } });
    wd.getRow(1).height = 32;
    wd.mergeCells('B2:E2');
    Object.assign(wd.getCell('B2'), { value: 'Receita dos atendimentos concluídos, dia a dia e acumulada no mês', font: { name: FONT, size: 9, italic: true, color: { argb: X.muted } } });
    ['Data', 'Dia da semana', 'Faturamento do dia', 'Acumulado do mês'].forEach((h, i) => {
      const c = wd.getCell(4, 2 + i);
      c.value = h;
      c.font = { name: FONT, size: 10, bold: true, color: { argb: X.white } };
      c.fill = solid(X.ink);
      c.alignment = { vertical: 'middle', horizontal: i < 2 ? 'left' : 'right', indent: 1 };
    });
    wd.getRow(4).height = 22;
    const nowD = new Date(), isCurrent = view.y === nowD.getFullYear() && view.m === nowD.getMonth(), todayDay = +today.slice(8, 10);
    let run = 0;
    A.daily.forEach((v, i) => {
      const row = 5 + i, k = `${ym}-${String(i + 1).padStart(2, '0')}`;
      run += v;
      const vals = [utcDate(k), utcDate(k).toLocaleDateString('pt-BR', { weekday: 'long', timeZone: 'UTC' }), v, (isCurrent && i + 1 > todayDay) ? null : run];
      vals.forEach((val, j) => {
        const c = wd.getCell(row, 2 + j);
        if (val !== null) c.value = val;
        c.font = { name: FONT, size: 10.5, color: { argb: j === 1 ? X.muted : X.ink }, bold: j === 2 && v > 0 };
        c.alignment = { vertical: 'middle', horizontal: j < 2 ? 'left' : 'right', indent: 1 };
        c.border = { bottom: { style: 'thin', color: { argb: X.line } } };
        if (j === 0) c.numFmt = 'dd/mm/yyyy';
        if (j >= 2) c.numFmt = MONEY;
      });
      wd.getRow(row).height = 20;
    });
    const last = 4 + A.dim, tot = last + 1;
    wd.addConditionalFormatting({ ref: `D5:D${last}`, rules: [{ type: 'dataBar', priority: 1, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: argb('#b08a4a') } }] });
    [['Total do mês', null], ['', null], [{ formula: `SUM(D5:D${last})`, result: A.revenue }, MONEY], ['', null]].forEach(([val, fmt], j) => {
      const c = wd.getCell(tot, 2 + j);
      if (val !== '') c.value = val;
      c.font = { name: FONT, size: 10.5, bold: true, color: { argb: X.ink } };
      c.fill = solid(X.mist);
      c.alignment = { vertical: 'middle', horizontal: j < 2 ? 'left' : 'right', indent: 1 };
      c.border = { top: { style: 'thin', color: { argb: X.ink } } };
      if (fmt) c.numFmt = fmt;
    });
    wd.getRow(tot).height = 22;

    /* ---------- Aba 3: Agendamentos do mês ---------- */
    const wa = wb.addWorksheet('Agendamentos do mês', {
      views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }],
      properties: { tabColor: { argb: X.ok } }
    });
    wa.columns = [{ width: 3 }, { width: 13 }, { width: 10 }, { width: 28 }, { width: 17 }, { width: 26 }, { width: 15 }, { width: 22 }, { width: 16 }, { width: 17 }];
    wa.mergeCells('B1:J1');
    Object.assign(wa.getCell('B1'), { value: `Agendamentos — ${mes} ${view.y}`, font: { name: SERIF, size: 18, bold: true, color: { argb: X.ink } }, alignment: { vertical: 'middle' } });
    wa.getRow(1).height = 32;
    const list = all.filter(a => a.data && a.data.slice(0, 7) === ym)
      .sort((a, b) => (a.data + (a.horario || '')).localeCompare(b.data + (b.horario || '')));
    wa.mergeCells('B2:J2');
    Object.assign(wa.getCell('B2'), { value: `${list.length} ${list.length === 1 ? 'solicitação' : 'solicitações'} no mês · base dos números do dashboard`, font: { name: FONT, size: 9, italic: true, color: { argb: X.muted } } });
    const heads = ['Data', 'Horário', 'Cliente', 'WhatsApp', 'Serviço', 'Valor', 'Status', 'Sinal', 'Reagendamentos'];
    heads.forEach((h, i) => {
      const c = wa.getCell(4, 2 + i);
      c.value = h;
      c.font = { name: FONT, size: 10, bold: true, color: { argb: X.white } };
      c.fill = solid(X.ink);
      c.alignment = { vertical: 'middle', horizontal: i === 5 || i === 8 ? 'right' : 'left', indent: 1 };
    });
    wa.getRow(4).height = 22;
    const DEP = { pendente: 'Pendente', pago: 'Pago', 'nao-aplicavel': 'Não se aplica' };
    list.forEach((a, i) => {
      const row = 5 + i, sk = statusKey(a), meta = STATUS_META.find(s => s.k === sk);
      const vals = [utcDate(a.data), a.horario || '', a.nome || '', formatPhoneDisplay(a.telefone), a.servico || '', +a.preco || 0,
        meta ? meta.label : (a.status || ''), DEP[a.deposito] || a.deposito || '', +a.reagendamentos || 0];
      vals.forEach((val, j) => {
        const c = wa.getCell(row, 2 + j);
        c.value = val;
        c.font = { name: FONT, size: 10.5, color: { argb: X.ink }, bold: j === 6 };
        c.alignment = { vertical: 'middle', horizontal: j === 5 || j === 8 ? 'right' : 'left', indent: 1 };
        c.border = { bottom: { style: 'thin', color: { argb: X.line } } };
        if (j === 0) c.numFmt = 'dd/mm/yyyy';
        if (j === 5) c.numFmt = MONEY;
        if (j === 6 && meta) { c.fill = solid(tint(meta.color)); c.font = { name: FONT, size: 10.5, bold: true, color: { argb: argb(meta.color) } }; }
      });
      wa.getRow(row).height = 20;
    });
    if (list.length) wa.autoFilter = { from: { row: 4, column: 2 }, to: { row: 4 + list.length, column: 10 } };
    else {
      wa.mergeCells('B5:J5');
      Object.assign(wa.getCell('B5'), { value: 'Nenhuma solicitação neste mês', font: { name: FONT, size: 10.5, color: { argb: X.muted } }, alignment: { horizontal: 'center', vertical: 'middle' } });
      wa.getRow(5).height = 28;
    }

    return wb;
  }

  /* =======================================================
     Download e carregamento da Lib
     ======================================================= */
  let libPromise = null;
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!libPromise) {
      libPromise = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = EXCELJS_URL; s.async = true;
        s.onload = () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS indisponível'));
        s.onerror = () => { libPromise = null; s.remove(); reject(new Error('Falha ao carregar ExcelJS')); };
        document.head.appendChild(s);
      });
    }
    return libPromise;
  }

  async function exportDashboard() {
    const btn = document.getElementById('mExport');
    if (!btn || btn.disabled) return;
    if (!window.DashState || !window.DashUtils) return;
    if (!window.DashState.ready) { toast('Aguarde os dados carregarem.'); return; }

    const U = window.DashUtils;
    const label = btn.querySelector('.exp-label'), old = label.textContent;
    btn.disabled = true; label.textContent = 'Preparando...';

    try {
      // 1. Carrega a Lib
      const ExcelJS = await loadExcelJS();
      
      // 2. Transforma os SVGs da tela em Imagens
      label.textContent = 'Desenhando...';
      const imgs = {
        cDaily: await svgToPngBase64('cDaily'),
        cStatus: await svgToPngBase64('cStatus'),
        cTrend: await svgToPngBase64('cTrend'),
        cServices: await svgToPngBase64('cServices'),
        cOcc: await svgToPngBase64('cOcc')
      };

      // 3. Monta o arquivo
      label.textContent = 'Gerando...';
      const now = new Date(), today = keyDate(now);
      const snapshotAll  = window.DashState.all.slice();
      const snapshotView = { ...window.DashState.view };
      const A = U.aggregate(snapshotAll, snapshotView.y, snapshotView.m, today);

      const wb = await buildWorkbook(ExcelJS, { all: snapshotAll, view: snapshotView, A, today, now }, imgs);
      const buf = await wb.xlsx.writeBuffer();

      // 4. Baixa
      const url = URL.createObjectURL(new Blob([buf], { type: MIME }));
      const a = document.createElement('a');
      a.href = url;
      a.download = `dashboard-nails-by-nascimento-${U.ymKey(snapshotView.y, snapshotView.m)}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      toast(`Planilha de ${U.MESES[snapshotView.m]} ${snapshotView.y} gerada.`);
    } catch (e) {
      console.error(e);
      toast('Não foi possível gerar a planilha. Verifique a conexão e tente de novo.');
    } finally {
      btn.disabled = false; label.textContent = old;
    }
  }

  /* Exposto para testes manuais, se precisar */
  globalThis.DashboardExport = { buildWorkbook };

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', () => {
      const btn = document.getElementById('mExport');
      if (btn) btn.addEventListener('click', exportDashboard);
    });
  }
})();
