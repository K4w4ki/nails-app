'use strict';
/* =========================================================
   EXPORTAR DASHBOARD PARA EXCEL — Área da Carol
   Depende de:
     - dashboard.js  → window.DashState, window.DashUtils
     - app-core.js   → br, keyDate, formatPhoneDisplay, toast
   A biblioteca ExcelJS só é baixada no primeiro clique.
   Gera os mesmos números que estão na tela, para o mês selecionado.
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

  /* Mesma regra do aggregate() para o status exibido */
  function statusKey(a) {
    let sk = a.status;
    if (sk === 'cancelado') {
      sk = a.canceladoPor === 'carol' ? 'cancelado-carol'
        : (a.canceladoPor === 'reagendamento' || a.reagendamentos > 0) ? 'reagendado'
        : 'cancelado-cliente';
    }
    return sk;
  }

  /* Gráficos visuais em PNG: ExcelJS não cria gráficos nativos.\n     As tabelas de origem permanecem editáveis nas outras abas. */
  function chartCanvas(title, subtitle) {
    const canvas = document.createElement('canvas');
    canvas.width = 1000; canvas.height = 520;
    const c = canvas.getContext('2d');
    c.fillStyle = '#FFFFFF'; c.fillRect(0, 0, canvas.width, canvas.height);
    c.fillStyle = '#B08A4A'; c.fillRect(48, 42, 5, 52);
    c.fillStyle = '#282828'; c.font = 'bold 30px Georgia'; c.fillText(title, 72, 68);
    c.fillStyle = '#737373'; c.font = '18px Arial'; c.fillText(subtitle || '', 72, 101);
    c.strokeStyle = '#DEDEDE'; c.lineWidth = 1; c.beginPath(); c.moveTo(48, 126); c.lineTo(952, 126); c.stroke();
    return { canvas, c };
  }
  function addCanvasChart(wb, ws, title, subtitle, draw, col, row) {
    const { canvas, c } = chartCanvas(title, subtitle);
    draw(c, canvas);
    const id = wb.addImage({ base64: canvas.toDataURL('image/png').split(',')[1], extension: 'png' });
    ws.addImage(id, { tl: { col, row }, ext: { width: 440, height: 229 } });
  }
  const moneyLabel = v => 'R$ ' + Number(v || 0).toLocaleString('pt-BR', { maximumFractionDigits: 0 });
  function drawBars(c, labels, values, color, suffix = '') {
    const left = 90, top = 165, width = 820, height = 275;
    const max = Math.max(1, ...values), step = width / Math.max(1, values.length);
    c.font = '15px Arial'; c.textAlign = 'right'; c.fillStyle = '#737373';
    for (let i = 0; i <= 4; i++) {
      const y = top + height - height * i / 4;
      c.strokeStyle = '#ECECEB'; c.beginPath(); c.moveTo(left, y); c.lineTo(left + width, y); c.stroke();
      c.fillText(Math.round(max * i / 4).toLocaleString('pt-BR') + suffix, left - 12, y + 5);
    }
    values.forEach((v, i) => {
      const bw = Math.max(4, step * .62), bh = height * v / max, x = left + i * step + (step - bw) / 2;
      c.fillStyle = color; c.fillRect(x, top + height - bh, bw, bh);
      if (labels.length <= 12 || i % Math.ceil(labels.length / 12) === 0 || i === labels.length - 1) {
        c.fillStyle = '#737373'; c.textAlign = 'center'; c.font = '13px Arial';
        c.fillText(labels[i], x + bw / 2, top + height + 25);
      }
    });
  }
  function addChartsSheet(wb, A, view, U) {
    const { MESES, STATUS_META, ymKey, shift } = U;
    const ws = wb.addWorksheet('Gráficos', { views: [{ showGridLines: false }], properties: { tabColor: { argb: X.gold } } });
    ws.columns = Array.from({ length: 10 }, () => ({ width: 12 }));
    ws.mergeCells('B1:J1');
    Object.assign(ws.getCell('B1'), { value: `Painel visual — ${MESES[view.m]} ${view.y}`, font: { name: SERIF, size: 20, bold: true, color: { argb: X.ink } } });
    ws.mergeCells('B2:J2');
    Object.assign(ws.getCell('B2'), { value: 'Gráficos gerados a partir do mesmo recorte de dados do dashboard no momento da exportação.', font: { name: FONT, size: 10, italic: true, color: { argb: X.muted } } });

    addCanvasChart(wb, ws, 'Faturamento por dia', 'Receita diária e acumulada de atendimentos concluídos', (c) => {
      drawBars(c, A.daily.map((_, i) => String(i + 1)), A.daily, '#3D3D3D', '');
      const left=90, top=165, width=820, height=275, cumulative=[]; let run=0;
      const current = view.y === new Date().getFullYear() && view.m === new Date().getMonth();
      const limit = current ? new Date().getDate() : A.dim;
      for(let i=0;i<A.dim;i++){ run += A.daily[i]; cumulative.push(i < limit ? run : null); }
      const max=Math.max(1,...cumulative.filter(v=>v!==null));
      c.strokeStyle='#B08A4A'; c.lineWidth=4; c.beginPath(); let started=false;
      cumulative.forEach((v,i)=>{ if(v===null)return; const x=left+(i+.5)*(width/A.dim), y=top+height-height*v/max; if(!started){c.moveTo(x,y);started=true;}else c.lineTo(x,y); }); c.stroke();
      c.textAlign='left'; c.font='15px Arial'; c.fillStyle='#3D3D3D'; c.fillText('■ Faturamento diário', 680, 151); c.fillStyle='#B08A4A'; c.fillText('━ Acumulado', 825, 151);
    }, 1, 3);

    addCanvasChart(wb, ws, 'Situação das solicitações', 'Distribuição por status no mês selecionado', (c) => {
      const vals = STATUS_META.map(s => A.status[s.k]), total = vals.reduce((a,b) => a+b, 0);
      const cx = 315, cy = 305, radius = 135; let angle = -Math.PI / 2;
      vals.forEach((v, i) => {
        if (!v || !total) return;
        const end = angle + v / total * Math.PI * 2;
        c.beginPath(); c.moveTo(cx, cy); c.arc(cx, cy, radius, angle, end); c.closePath(); c.fillStyle = STATUS_META[i].color; c.fill(); angle = end;
      });
      c.beginPath(); c.arc(cx, cy, 78, 0, Math.PI * 2); c.fillStyle = '#FFFFFF'; c.fill();
      c.fillStyle = '#282828'; c.font = 'bold 34px Arial'; c.textAlign = 'center'; c.fillText(String(total), cx, cy + 10);
      c.font = '14px Arial'; c.fillStyle = '#737373'; c.fillText('solicitações', cx, cy + 34);
      STATUS_META.forEach((s, i) => {
        const y = 178 + i * 35; c.fillStyle = s.color; c.fillRect(555, y - 13, 16, 16);
        c.fillStyle = '#3D3D3D'; c.font = '15px Arial'; c.textAlign = 'left';
        c.fillText(`${s.label} (${vals[i]})`, 582, y);
      });
    }, 6, 3);

    const months = [5,4,3,2,1,0].map(i => shift(view, -i));
    const series = months.map(x => A.monthly[ymKey(x.y, x.m)] || { rev: 0, n: 0 });
    addCanvasChart(wb, ws, 'Tendência dos últimos 6 meses', 'Faturamento e atendimentos concluídos', (c) => {
      drawBars(c, months.map(x => U.MESES_CURTO[x.m]), series.map(x => x.rev), '#3D3D3D');
      const left=90, top=165, width=820, height=275, max=Math.max(1,...series.map(x=>x.n));
      c.strokeStyle='#B08A4A'; c.lineWidth=5; c.beginPath();
      series.forEach((v,i)=>{ const x=left+(i+.5)*(width/6), y=top+height-height*v.n/max; if(i===0)c.moveTo(x,y);else c.lineTo(x,y); }); c.stroke();
      series.forEach((v,i)=>{ const x=left+(i+.5)*(width/6), y=top+height-height*v.n/max; c.beginPath();c.arc(x,y,7,0,Math.PI*2);c.fillStyle='#B08A4A';c.fill(); });
      c.textAlign='left'; c.font='15px Arial'; c.fillStyle='#3D3D3D'; c.fillText('■ Faturamento', 690, 151); c.fillStyle='#B08A4A'; c.fillText('● Atendimentos', 815, 151);
    }, 1, 15);

    const services = Object.entries(A.services).sort((a,b)=>b[1].n-a[1].n||b[1].rev-a[1].rev).slice(0,6);
    addCanvasChart(wb, ws, 'Serviços mais realizados', 'Atendimentos concluídos no mês · top 6', (c) => {
      const max=Math.max(1,...services.map(x=>x[1].n)), x0=330, y0=165, barW=570, rowH=34;
      services.forEach(([name,v],i)=>{ const y=y0+i*rowH; c.fillStyle='#737373';c.font='15px Arial';c.textAlign='right';c.fillText(name.length>30?name.slice(0,27)+'…':name,x0-14,y+18);c.fillStyle='#B08A4A';c.fillRect(x0,y,barW*v.n/max,22);c.fillStyle='#282828';c.textAlign='left';c.fillText(String(v.n),x0+barW*v.n/max+8,y+18); });
      if (!services.length) { c.fillStyle='#737373';c.textAlign='center';c.font='20px Arial';c.fillText('Nenhum atendimento concluído no mês',500,300); }
    }, 6, 15);

    addCanvasChart(wb, ws, 'Ocorrências por semana', 'Reagendamentos e cancelamentos realizados pela Carol', (c) => {
      const nW=A.dim>28?5:4, labels=['1–7','8–14','15–21','22–28',`29–${A.dim}`].slice(0,nW), left=105, top=170, width=800, height=260;
      const max=Math.max(1,...A.weeks.resched.slice(0,nW),...A.weeks.cancel.slice(0,nW)), group=width/nW;
      for(let i=0;i<=4;i++){const y=top+height-height*i/4;c.strokeStyle='#ECECEB';c.beginPath();c.moveTo(left,y);c.lineTo(left+width,y);c.stroke();}
      labels.forEach((l,i)=>{const x=left+i*group+group/2,bw=Math.min(42,group*.25);const a=A.weeks.resched[i],b=A.weeks.cancel[i];c.fillStyle='#8B7FA3';c.fillRect(x-bw-3,top+height-height*a/max,bw,height*a/max);c.fillStyle='#7A3D40';c.fillRect(x+3,top+height-height*b/max,bw,height*b/max);c.fillStyle='#737373';c.font='15px Arial';c.textAlign='center';c.fillText(l,x,top+height+26);});
      c.textAlign='left';c.font='15px Arial';c.fillStyle='#8B7FA3';c.fillText('■ Reagendamentos',680,151);c.fillStyle='#7A3D40';c.fillText('■ Cancelamentos (Carol)',820,151);
    }, 1, 27);
  }

  /* =======================================================
     Monta a planilha. ctx = { all, view, A, today, now }
     ======================================================= */
  function buildWorkbook(ExcelJS, ctx) {
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

    /* Linha de tabela. kind: head | body | total | hero. span3: 3º valor ocupa D:E */
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
      ref, rules: [{
        type: 'dataBar', priority: 1,
        cfvo: [{ type: 'num', value: 0 }, { type: 'max' }],
        color: { argb: color }
      }]
    });
    const gap = () => { r++; };

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

    /* --- Situação das solicitações --- */
    section('Agenda', 'Situação das solicitações', 'Todas as solicitações do mês, por status');
    tr(['Status', 'Solicitações', '% do mês'], { kind: 'head' });
    const st0 = r;
    STATUS_META.forEach(s => {
      const v = A.status[s.k];
      tr([s.label, v, A.total ? v / A.total : 0], { fmts: [null, '0', '0%'], swatch: argb(s.color) });
    });
    bars(`C${st0}:C${r - 1}`, 'FFB9B9B9');
    tr(['Total do mês', A.total, A.total ? 1 : 0], { kind: 'total', fmts: [null, '0', '0%'] });
    gap();

    /* --- Serviços mais realizados --- */
    const svc = Object.entries(A.services).sort((a, b) => b[1].n - a[1].n || b[1].rev - a[1].rev);
    section('Catálogo', 'Serviços mais realizados', 'Atendimentos concluídos no mês, do mais ao menos realizado');
    tr(['Serviço', 'Atendimentos', 'Faturamento'], { kind: 'head' });
    if (svc.length) {
      const sv0 = r;
      svc.forEach(([name, v], i) => tr([name, v.n, v.rev], { fmts: [null, '0', MONEY], boldFirst: i === 0 }));
      bars(`C${sv0}:C${r - 1}`, argb('#b08a4a'));
    } else {
      tr(['Nenhum serviço concluído neste mês', '', ''], { muted: true });
    }
    gap();

    /* --- Últimos 6 meses --- */
    const months = [5, 4, 3, 2, 1, 0].map(i => shift(view, -i));
    const series = months.map(x => A.monthly[ymKey(x.y, x.m)] || { rev: 0, n: 0 });
    section('Tendência', 'Últimos 6 meses', 'Faturamento e atendimentos concluídos, mês a mês');
    tr(['Mês', 'Faturamento', 'Atendimentos'], { kind: 'head' });
    const t0 = r;
    months.forEach((x, i) => tr([`${MESES[x.m]} ${x.y}`, series[i].rev, series[i].n], { fmts: [null, MONEY, '0'], boldFirst: i === 5 }));
    bars(`C${t0}:C${r - 1}`, 'FFB9B9B9');
    gap();

    /* --- Ocorrências por semana --- */
    const nW = A.dim > 28 ? 5 : 4;
    const wk = ['1–7', '8–14', '15–21', '22–28', `29–${A.dim}`].slice(0, nW);
    section('Ocorrências', 'Reagendamentos e cancelamentos', 'Por semana do mês · cancelamentos feitos pela Carol');
    tr(['Semana (dias)', 'Reagendamentos', 'Cancelamentos (Carol)'], { kind: 'head' });
    const o0 = r;
    wk.forEach((l, i) => tr([`Dias ${l}`, A.weeks.resched[i], A.weeks.cancel[i]], { fmts: [null, '0', '0'] }));
    bars(`C${o0}:C${r - 1}`, argb('#8b7fa3'));
    bars(`D${o0}:D${r - 1}`, argb('#7a3d40'));
    tr(['Total', A.resched, A.cancelCarol], { kind: 'total', fmts: [null, '0', '0'] });

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
    wd.addConditionalFormatting({
      ref: `D5:D${last}`,
      rules: [{ type: 'dataBar', priority: 1, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: argb('#b08a4a') } }]
    });
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

    addChartsSheet(wb, A, view, U);
    return wb;
  }

  /* =======================================================
     Navegador: carrega a biblioteca, monta e baixa o arquivo
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

    /* Estado e utils vêm do dashboard.js */
    if (!window.DashState || !window.DashUtils) {
      toast('O dashboard ainda não terminou de carregar. Tente de novo em instantes.');
      return;
    }
    if (!window.DashState.ready) { toast('Aguarde os dados carregarem para exportar.'); return; }

    const U = window.DashUtils;
    const label = btn.querySelector('.exp-label'), old = label.textContent;
    btn.disabled = true; btn.setAttribute('aria-busy', 'true'); label.textContent = 'Gerando…';

    try {
      /* Congela os dados e o mês exatamente no momento do clique,
         inclusive enquanto a biblioteca do Excel é carregada. */
      const now = new Date(), today = keyDate(now);
      const snapshotAll = window.DashState.all.slice();
      const snapshotView = { ...window.DashState.view };
      const A = U.aggregate(snapshotAll, snapshotView.y, snapshotView.m, today);
      const ExcelJS = await loadExcelJS();
      const wb = buildWorkbook(ExcelJS, { all: snapshotAll, view: snapshotView, A, today, now });
      const buf = await wb.xlsx.writeBuffer();

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
      btn.disabled = false; btn.removeAttribute('aria-busy'); label.textContent = old;
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