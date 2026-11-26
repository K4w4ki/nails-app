'use strict';
/* =========================================================
   EXPORTAR DASHBOARD PARA EXCEL VISUAL (XLSX)
   Gera tabelas formatadas com a identidade visual do site
   e insere as imagens dos gráficos no arquivo.
   ========================================================= */
(function () {
  const EXCELJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  /* Paleta do site (styles.css) em formato ARGB */
  const X = {
    ink: 'FF282828', ink2: 'FF3D3D3D', gold: 'FFB08A4A', mist: 'FFF6F6F5', soft: 'FFECECEB',
    line: 'FFDEDEDE', muted: 'FF737373', ok: 'FF46654D', bad: 'FF965154', white: 'FFFFFFFF'
  };
  const FONT = 'Calibri', SERIF = 'Georgia';
  const MONEY = '"R$" #,##0.00';

  // Utilitários de formatação
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
  const MES_T = { none: 'sem mês anterior', same: 'igual ao mês anterior', vs: 'vs mês anterior' };
  const SEM_T = { none: 'sem semana anterior', same: 'igual à semana anterior', vs: 'vs semana anterior' };

  function statusKey(a) {
    let sk = a.status;
    if (sk === 'cancelado') {
      sk = a.canceladoPor === 'carol' ? 'cancelado-carol'
        : (a.canceladoPor === 'reagendamento' || a.reagendamentos > 0) ? 'reagendado'
        : 'cancelado-cliente';
    }
    return sk;
  }

  /* =========================================================
     Transforma os Gráficos (SVG) da tela em Imagem (PNG)
     ========================================================= */
  async function svgToPngBase64(containerId) {
    return new Promise((resolve) => {
      const container = document.getElementById(containerId);
      if (!container) return resolve(null);
      const svg = container.querySelector('svg');
      if (!svg) return resolve(null);

      const clone = svg.cloneNode(true);
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
        canvas.width = w * 2; // 2x para resolução melhor
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

  /* =========================================================
     Monta a planilha com Tabelas + Imagens
     ========================================================= */
  async function buildWorkbook(ExcelJS, ctx, imgs) {
    const { all, view, A, today, now } = ctx;
    const U = window.DashUtils;
    const { MESES, STATUS_META, ymKey, shift, sumWeek, weekRangeOffset } = U;

    const mes = MESES[view.m], ym = ymKey(view.y, view.m);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Nails by Nascimento';
    wb.created = now;
    wb.title = `Dashboard — ${mes} ${view.y}`;

    /* --- Aba 1: Dashboard --- */
    const ws = wb.addWorksheet('Dashboard', {
      views: [{ showGridLines: false }],
      properties: { tabColor: { argb: X.gold } },
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 }
    });
    ws.columns = [{ width: 3 }, { width: 36 }, { width: 20 }, { width: 20 }, { width: 20 }];
    let r = 2;

    const mergeBE = (row) => ws.mergeCells(`B${row}:E${row}`);

    // Faixa de marca
    mergeBE(r);
    Object.assign(ws.getCell(r, 2), { value: 'Nails by Nascimento  ·  Carol Nascimento', font: { name: SERIF, size: 12, bold: true, color: { argb: X.white } }, fill: solid(X.ink), alignment: { vertical: 'middle', indent: 1 } });
    for (let c = 3; c <= 5; c++) ws.getCell(r, c).fill = solid(X.ink);
    ws.getRow(r).height = 30; r += 2;

    Object.assign(ws.getCell(r, 2), { value: 'DASHBOARD', font: { name: FONT, size: 9, bold: true, color: { argb: X.gold } } });
    r++; mergeBE(r);
    Object.assign(ws.getCell(r, 2), { value: 'Visão geral do estúdio', font: { name: SERIF, size: 24, bold: true, color: { argb: X.ink } } });
    ws.getRow(r).height = 36; r++; mergeBE(r);
    Object.assign(ws.getCell(r, 2), { value: `${mes} ${view.y}  ·  gerado em ${now.toLocaleDateString('pt-BR')}`, font: { name: FONT, size: 10, color: { argb: X.muted } } });
    r += 2;

    const section = (eyebrow, title) => {
      Object.assign(ws.getCell(r, 2), { value: eyebrow.toUpperCase(), font: { name: FONT, size: 9, bold: true, color: { argb: X.gold } } });
      r++; mergeBE(r);
      Object.assign(ws.getCell(r, 2), { value: title, font: { name: SERIF, size: 17, bold: true, color: { argb: X.ink } } });
      ws.getRow(r).height = 26; r++;
    };

    const tr = (vals, o = {}) => {
      const kind = o.kind || 'body', n = o.span3 ? 4 : vals.length;
      const bg = kind === 'head' || kind === 'hero' ? X.ink : kind === 'total' ? X.mist : null;
      const fg = kind === 'head' || kind === 'hero' ? X.white : X.ink;
      for (let i = 0; i < n; i++) {
        const c = ws.getCell(r, 2 + i), v = vals[i];
        if (v !== undefined) c.value = v;
        const isText = i === 0 || (o.span3 && i >= 2);
        c.font = { name: FONT, size: kind === 'head' ? 10 : 10.5, bold: kind !== 'body' || (!!o.boldFirst && i === 0), color: { argb: fg } };
        c.alignment = { vertical: 'middle', horizontal: isText ? 'left' : 'right', indent: 1 };
        if (bg) c.fill = solid(bg);
        if (o.fmts && o.fmts[i]) c.numFmt = o.fmts[i];
        c.border = { bottom: { style: 'thin', color: { argb: X.line } } };
      }
      if (o.span3) ws.mergeCells(`D${r}:E${r}`);
      if (o.swatch) ws.getCell(r, 1).fill = solid(o.swatch);
      ws.getRow(r).height = 21;
      r++;
    };
    const gap = () => { r += 2; };

    // Resumo Semanal
    const wCur = weekRangeOffset(0), wPrev = weekRangeOffset(-1);
    const sCur = sumWeek(all, wCur.start, wCur.end), sPrev = sumWeek(all, wPrev.start, wPrev.end);
    section('Resumo da semana', `${window.br(wCur.start)} – ${window.br(wCur.end)}`);
    tr(['Indicador', 'Valor', 'Comparação'], { kind: 'head', span3: true });
    tr(['Faturamento', sCur.revenue, plainDelta(sCur.revenue, sPrev.revenue, SEM_T)], { span3: true, fmts: [null, MONEY] });
    tr(['Atendimentos', sCur.done, plainDelta(sCur.done, sPrev.done, SEM_T)], { span3: true, fmts: [null, '0'] });
    gap();

    // Indicadores do mês
    const prevV = shift(view, -1);
    const prev = A.monthly[ymKey(prevV.y, prevV.m)] || { rev: 0, n: 0 };
    section('Indicadores', `${mes} ${view.y}`);
    tr(['Indicador', 'Valor', 'Detalhe'], { kind: 'head', span3: true });
    tr(['Faturamento do mês', A.revenue, plainDelta(A.revenue, prev.rev, MES_T)], { kind: 'hero', span3: true, fmts: [null, MONEY] });
    tr(['Atendimentos realizados', A.done, plainDelta(A.done, prev.n, MES_T)], { span3: true, fmts: [null, '0'] });
    tr(['Pendentes', A.pend, ''], { span3: true, fmts: [null, '0'] });
    tr(['Aguardando sinal', A.wait, ''], { span3: true, fmts: [null, '0'] });
    tr(['Confirmados hoje', A.today, ''], { span3: true, fmts: [null, '0'] });
    tr(['Reagendamentos', A.resched, ''], { span3: true, fmts: [null, '0'] });
    tr(['Cancelamentos (Carol)', A.cancelCarol, ''], { span3: true, fmts: [null, '0'] });
    gap();

    // Imagem Faturamento
    if (imgs.cDaily) {
      section('Receita', 'Faturamento por dia');
      const imgId = wb.addImage({ base64: imgs.cDaily.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 1, row: r - 1 }, ext: { width: 600, height: 300 } });
      r += 16;
    }

    // Imagem Status + Tabela
    if (imgs.cStatus) {
      section('Agenda', 'Situação das solicitações');
      const imgId = wb.addImage({ base64: imgs.cStatus.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 1, row: r - 1 }, ext: { width: 350, height: 300 } });
      r += 16;
    }
    
    tr(['Status', 'Qtd'], { kind: 'head' });
    STATUS_META.forEach(s => tr([s.label, A.status[s.k]], { fmts: [null, '0'], swatch: argb(s.color) }));
    gap();

    // Imagem Tendência
    if (imgs.cTrend) {
      section('Tendência', 'Últimos 6 meses');
      const imgId = wb.addImage({ base64: imgs.cTrend.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 1, row: r - 1 }, ext: { width: 600, height: 300 } });
      r += 16;
    }

    // Imagem Serviços
    if (imgs.cServices) {
      section('Catálogo', 'Serviços mais realizados');
      const imgId = wb.addImage({ base64: imgs.cServices.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 1, row: r - 1 }, ext: { width: 600, height: 300 } });
      r += 16;
    }

    // Imagem Ocorrências
    if (imgs.cOcc) {
      section('Ocorrências', 'Reagendamentos e cancelamentos');
      const imgId = wb.addImage({ base64: imgs.cOcc.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 1, row: r - 1 }, ext: { width: 600, height: 300 } });
      r += 16;
    }

    /* --- Aba 2: Faturamento por dia --- */
    const wd = wb.addWorksheet('Faturamento por dia', { views: [{ showGridLines: false }] });
    wd.columns = [{ width: 3 }, { width: 16 }, { width: 22 }, { width: 22 }];
    wd.mergeCells('B2:D2');
    Object.assign(wd.getCell('B2'), { value: `Faturamento por dia — ${mes} ${view.y}`, font: { name: SERIF, size: 16, bold: true } });
    ['Data', 'Faturamento do dia', 'Acumulado'].forEach((h, i) => {
      const c = wd.getCell(4, 2 + i);
      c.value = h; c.font = { name: FONT, size: 10, bold: true, color: { argb: X.white } }; c.fill = solid(X.ink);
    });
    let run = 0;
    A.daily.forEach((v, i) => {
      run += v;
      const row = 5 + i;
      wd.getCell(row, 2).value = utcDate(`${ym}-${String(i+1).padStart(2,'0')}`);
      wd.getCell(row, 2).numFmt = 'dd/mm/yyyy';
      wd.getCell(row, 3).value = v; wd.getCell(row, 3).numFmt = MONEY;
      wd.getCell(row, 4).value = run; wd.getCell(row, 4).numFmt = MONEY;
    });

    /* --- Aba 3: Agendamentos --- */
    const wa = wb.addWorksheet('Agendamentos', { views: [{ showGridLines: false }] });
    wa.columns = [{ width: 3 }, { width: 13 }, { width: 10 }, { width: 28 }, { width: 17 }, { width: 26 }, { width: 15 }, { width: 22 }, { width: 16 }];
    const list = all.filter(a => a.data && a.data.slice(0, 7) === ym).sort((a, b) => (a.data + (a.horario || '')).localeCompare(b.data + (b.horario || '')));
    const heads = ['Data', 'Horário', 'Cliente', 'WhatsApp', 'Serviço', 'Valor', 'Status', 'Sinal', 'Reagendamentos'];
    heads.forEach((h, i) => {
      const c = wa.getCell(4, 2 + i);
      c.value = h; c.font = { name: FONT, size: 10, bold: true, color: { argb: X.white } }; c.fill = solid(X.ink);
    });
    list.forEach((a, i) => {
      const row = 5 + i, sk = statusKey(a), meta = STATUS_META.find(s => s.k === sk);
      const vals = [utcDate(a.data), a.horario || '', a.nome || '', window.formatPhoneDisplay(a.telefone), a.servico || '', +a.preco || 0, meta ? meta.label : (a.status || ''), a.deposito || '', +a.reagendamentos || 0];
      vals.forEach((val, j) => {
        const c = wa.getCell(row, 2 + j);
        c.value = val; c.font = { name: FONT, size: 10.5, color: { argb: X.ink } }; c.border = { bottom: { style: 'thin', color: { argb: X.line } } };
        if (j === 0) c.numFmt = 'dd/mm/yyyy';
        if (j === 5) c.numFmt = MONEY;
        if (j === 6 && meta) { c.fill = solid(tint(meta.color)); c.font = { name: FONT, size: 10.5, bold: true, color: { argb: argb(meta.color) } }; }
      });
    });

    return wb;
  }

  /* =========================================================
     Orquestração do Download
     ========================================================= */
  let libPromise = null;
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!libPromise) {
      libPromise = new Promise((resolve, reject) => {
        const s = document.createElement('script'); s.src = EXCELJS_URL; s.async = true;
        s.onload = () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('Erro'));
        document.head.appendChild(s);
      });
    }
    return libPromise;
  }

  async function exportarExcel() {
    const btn = document.getElementById('mExport');
    if (!btn || btn.disabled) return;
    if (!window.DashState || !window.DashState.ready) return window.toast('Aguarde os dados carregarem.');

    const label = btn.querySelector('.exp-label'), old = label.textContent;
    btn.disabled = true; label.textContent = 'Gerando...';

    try {
      const ExcelJS = await loadExcelJS();
      
      label.textContent = 'Desenhando...';
      const imgs = {
        cDaily: await svgToPngBase64('cDaily'), cStatus: await svgToPngBase64('cStatus'),
        cTrend: await svgToPngBase64('cTrend'), cServices: await svgToPngBase64('cServices'), cOcc: await svgToPngBase64('cOcc')
      };
      
      label.textContent = 'Montando...';
      const now = new Date(), today = window.keyDate(now);
      const snapshotAll = window.DashState.all.slice();
      const snapshotView = { ...window.DashState.view };
      const A = window.DashUtils.aggregate(snapshotAll, snapshotView.y, snapshotView.m, today);

      const wb = await buildWorkbook(ExcelJS, { all: snapshotAll, view: snapshotView, A, today, now }, imgs);
      const buf = await wb.xlsx.writeBuffer();

      const url = URL.createObjectURL(new Blob([buf], { type: MIME_XLSX }));
      const a = document.createElement('a'); a.href = url; a.download = `dashboard-visual-${window.DashUtils.ymKey(snapshotView.y, snapshotView.m)}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      window.toast('Excel Visual gerado com sucesso!');
    } catch (e) {
      console.error(e); window.toast('Erro ao gerar o Excel.');
    } finally {
      btn.disabled = false; label.textContent = old;
    }
  }

  /* Inicialização */
  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', () => {
      const bXlsx = document.getElementById('mExport');
      if (bXlsx) bXlsx.addEventListener('click', exportarExcel);
    });
  }
})();
