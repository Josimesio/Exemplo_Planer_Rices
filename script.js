const els = {
  dataAtualizacao: document.getElementById('dataAtualizacao'),
  globalPercent: document.getElementById('globalPercent'),
  ringProgress: document.getElementById('ringProgress'),
  statusExecutivo: document.getElementById('statusExecutivo'),
  totalItens: document.getElementById('totalItens'),
  totalBaseKeys: document.getElementById('totalBaseKeys') || document.getElementById('totalRicesIguais'),
  totalConcluidos: document.getElementById('totalConcluidos'),
  totalAndamento: document.getElementById('totalAndamento'),
  totalNaoIniciado: document.getElementById('totalNaoIniciado'),
  totalCriticos: document.getElementById('totalCriticos'),
  totalSoEmpresa: document.getElementById('totalSoEmpresa'),
  totalSoConsultoria: document.getElementById('totalSoConsultoria'),
  totalRicesEmpresa: document.getElementById('totalRicesEmpresa'),
  totalRicesConsultoria: document.getElementById('totalRicesConsultoria'),
  headlineCallout: document.getElementById('headlineCallout'),
  headlinePill: document.getElementById('headlinePill'),
  empresaMetrics: document.getElementById('empresaMetrics'),
  consultoriaMetrics: document.getElementById('consultoriaMetrics'),
  leaderboard: document.getElementById('leaderboard'),
  statusBars: document.getElementById('statusBars'),
  comparisonBoard: document.getElementById('comparisonBoard'),
  keyBoard: document.getElementById('keyBoard') || document.getElementById('riceBoard'),
  riceFilterSelect: document.getElementById('riceFilterSelect'),
  riceFilterStatus: document.getElementById('riceFilterStatus'),
  btnRiceFilterAll: document.getElementById('btnRiceFilterAll'),
  btnRiceFilterClear: document.getElementById('btnRiceFilterClear'),
  divergenceBoard: document.getElementById('divergenceBoard'),
  divergenceEmpresaList: document.getElementById('divergenceEmpresaList'),
  divergenceConsultoriaList: document.getElementById('divergenceConsultoriaList'),
  divergenceTable: document.getElementById('divergenceTable'),
  focusTable: document.getElementById('focusTable'),
  focusFilterOrigem: document.getElementById('focusFilterOrigem'),
  focusFilterRice: document.getElementById('focusFilterRice'),
  focusFilterFase: document.getElementById('focusFilterFase'),
  focusFilterTarefa: document.getElementById('focusFilterTarefa'),
  focusFilterResponsavel: document.getElementById('focusFilterResponsavel'),
  focusFilterStatus: document.getElementById('focusFilterStatus')
};

const RING_CIRCUMFERENCE = 301.59;
const state = { pairs: [], riceFilter: new Set(), focusRows: [], focusFilters: {} };

const EXECUTIVE_MESSAGES = [
  { threshold: 0, status: 'Leitura inicial', title: 'A base consolidada foi carregada e comparada entre Empresa e Consultoria.' },
  { threshold: 20, status: 'Avanço inicial', title: 'Existe progresso, mas ainda há volume relevante fora da linha de chegada.' },
  { threshold: 40, status: 'Tração consistente', title: 'O programa mostra ritmo real de entrega na base casada.' },
  { threshold: 60, status: 'Execução madura', title: 'O avanço é consistente. Agora o jogo é destravar gargalos.' },
  { threshold: 80, status: 'Alta performance', title: 'A base casada está forte. Falta limpar resíduos críticos.' },
  { threshold: 100, status: 'Ciclo concluído', title: 'Todos os RICEs comparáveis foram concluídos dos dois lados.' }
];

document.addEventListener('DOMContentLoaded', async () => {
  bindRiceFilter();
  bindFocusFilters();

  if (typeof DASH_CONFIG === 'undefined') {
    renderFatal('config.js não foi carregado.');
    return;
  }

  await loadDataFile();
});

async function loadDataFile() {
  const fileName = DASH_CONFIG.dataFile || 'dados_consilidados_rices.csv';
  try {
    setGeneratedAt(`Lendo base consolidada: ${fileName}...`);
    const response = await fetch(`${encodeURI(fileName)}?t=${Date.now()}`, { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const rows = fileName.toLowerCase().endsWith('.json') ? parseJsonRows(text) : parseCsv(text);
    processRows(rows, fileName);
  } catch (error) {
    console.error(error);
    renderFatal(`Falha ao carregar ${fileName}. Rode o converter_rices.py e publique o CSV junto com o site. Detalhe: ${error.message || error}`);
  }
}

function parseJsonRows(text) {
  const payload = JSON.parse(text);
  return Array.isArray(payload) ? payload : payload.rows || [];
}

function parseCsv(text) {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      row.push(field);
      field = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') i += 1;
      row.push(field);
      if (row.some(value => String(value).trim() !== '')) rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }

  if (field || row.length) {
    row.push(field);
    if (row.some(value => String(value).trim() !== '')) rows.push(row);
  }

  if (!rows.length) return [];
  const headers = rows[0].map(h => String(h || '').replace(/^\uFEFF/, '').trim());
  return rows.slice(1).map(values => {
    const obj = {};
    headers.forEach((header, index) => { obj[header] = values[index] || ''; });
    return obj;
  });
}

function detectDelimiter(text) {
  const firstLine = String(text || '').split(/\r?\n/)[0] || '';
  const semicolon = (firstLine.match(/;/g) || []).length;
  const comma = (firstLine.match(/,/g) || []).length;
  const tab = (firstLine.match(/\t/g) || []).length;
  if (semicolon >= comma && semicolon >= tab) return ';';
  if (tab >= comma) return '\t';
  return ',';
}

function processRows(rows, fileName) {
  const normalized = normalizeRows(rows);
  const model = buildComparisonModel(normalized);

  if (!model.totalUnion) {
    renderFatal('O CSV foi carregado, mas não encontrei RICEs válidos. Verifique a coluna RICE ou Nome da tarefa.');
    return;
  }

  renderDashboard(model);
  setGeneratedAt(`Base consolidada carregada: ${formatNumber(model.equalCount)} RICE(s) iguais, ${formatNumber(model.onlyEmpresa.length)} só na Empresa, ${formatNumber(model.onlyConsultoria.length)} só na Consultoria, ${formatNumber(model.totalUnion)} RICE(s) no total sem repetir. Arquivo: ${fileName}.`);
}

function normalizeRows(rows) {
  return rows.map(row => {
    const rice = getValue(row, ['RICE', 'Rice', 'RICE base', DASH_CONFIG.fields?.rice]) || extractRiceCode(getValue(row, DASH_CONFIG.fields.taskName));
    const origem = getValue(row, ['__origem', 'Origem']) || 'Empresa';
    return {
      raw: row,
      origem,
      riceCode: rice,
      compareKey: normalizeRiceKey(rice),
      compareDisplay: rice || '-',
      id: getValue(row, DASH_CONFIG.fields.taskId),
      tarefa: getValue(row, DASH_CONFIG.fields.taskName),
      categoria: getValue(row, DASH_CONFIG.fields.category) || getValue(row, 'Fase consolidada origem') || 'Sem fase',
      statusOriginal: getValue(row, DASH_CONFIG.fields.status) || getValue(row, 'Status consolidado origem') || 'Sem status',
      prioridade: getValue(row, DASH_CONFIG.fields.priority) || 'Sem prioridade',
      responsavel: firstResponsible(getValue(row, DASH_CONFIG.fields.responsible) || getValue(row, 'Responsáveis consolidados')),
      responsaveisTodos: splitPeople(getValue(row, DASH_CONFIG.fields.responsible) || getValue(row, 'Responsáveis consolidados')),
      atrasadoFlag: isDelayedValue(getValue(row, DASH_CONFIG.fields.delayed)),
      qtdRegistros: Number(getValue(row, 'Qtd registros na origem') || 1),
      fasesOrigem: getValue(row, 'Fases encontradas na origem'),
      tipoComparacao: getValue(row, 'Tipo comparação')
    };
  }).filter(row => row.compareKey);
}

function buildComparisonModel(rows) {
  const grouped = new Map();
  rows.forEach(row => {
    if (!grouped.has(row.compareKey)) {
      grouped.set(row.compareKey, { compareKey: row.compareKey, compareDisplay: row.compareDisplay, Empresa: null, Consultoria: null });
    }
    const item = grouped.get(row.compareKey);
    item.compareDisplay = item.compareDisplay || row.compareDisplay;
    item[row.origem] = row;
  });

  const pairs = [];
  const onlyEmpresa = [];
  const onlyConsultoria = [];

  Array.from(grouped.values()).sort((a, b) => String(a.compareDisplay).localeCompare(String(b.compareDisplay), 'pt-BR')).forEach(item => {
    if (item.Empresa && item.Consultoria) {
      const pairStatus = buildPairStatus(item.Empresa, item.Consultoria);
      pairs.push({
        compareKey: item.compareKey,
        compareDisplay: item.compareDisplay,
        empresa: item.Empresa,
        consultoria: item.Consultoria,
        pairStatus,
        progressPercent: buildPairProgressPercent(item.Empresa, item.Consultoria)
      });
    } else if (item.Empresa) {
      onlyEmpresa.push(item.Empresa);
    } else if (item.Consultoria) {
      onlyConsultoria.push(item.Consultoria);
    }
  });

  return {
    pairs,
    onlyEmpresa,
    onlyConsultoria,
    equalCount: pairs.length,
    totalUnion: pairs.length + onlyEmpresa.length + onlyConsultoria.length,
    empresaTotal: pairs.length + onlyEmpresa.length,
    consultoriaTotal: pairs.length + onlyConsultoria.length,
    comparableRows: pairs.flatMap(pair => [pair.empresa, pair.consultoria])
  };
}

function buildPairStatus(empresa, consultoria) {
  if (isCritical(empresa) || isCritical(consultoria)) return 'Crítico';
  if (isConcluded(empresa.statusOriginal) && isConcluded(consultoria.statusOriginal)) return 'Concluído';
  if (isNotStarted(empresa.statusOriginal) && isNotStarted(consultoria.statusOriginal)) return 'Não iniciado';
  return 'Em andamento';
}

function buildPairProgressPercent(empresa, consultoria) {
  const done = (isConcluded(empresa.statusOriginal) ? 1 : 0) + (isConcluded(consultoria.statusOriginal) ? 1 : 0);
  return Math.round((done / 2) * 100);
}

function renderDashboard(model) {
  const summary = buildSummary(model.pairs);
  updateSummary(summary, model);
  renderPlannerMetrics(model.comparableRows.filter(row => row.origem === 'Empresa'), 'Empresa', els.empresaMetrics, model.empresaTotal);
  renderPlannerMetrics(model.comparableRows.filter(row => row.origem === 'Consultoria'), 'Consultoria', els.consultoriaMetrics, model.consultoriaTotal);
  renderLeaderboard(model.comparableRows);
  renderStatusBars(summary);
  renderComparisonBoard(model.comparableRows);
  setupRiceFilter(model.pairs);
  renderDivergenceBoard(model.onlyEmpresa, model.onlyConsultoria);
  renderDivergencePanels(model.onlyEmpresa, model.onlyConsultoria);
  renderDivergenceTable(model.onlyEmpresa, model.onlyConsultoria);
  setupFocusTable(model.pairs);
}

function buildSummary(pairs) {
  const summary = { total: pairs.length, concluded: 0, inProgress: 0, notStarted: 0, critical: 0, percent: 0 };
  pairs.forEach(pair => {
    if (pair.pairStatus === 'Concluído') summary.concluded += 1;
    else if (pair.pairStatus === 'Não iniciado') summary.notStarted += 1;
    else if (pair.pairStatus === 'Crítico') summary.critical += 1;
    else summary.inProgress += 1;
  });
  summary.percent = getPercent(summary.concluded, summary.total);
  return summary;
}

function updateSummary(summary, model) {
  setText(els.totalItens, summary.total);
  setText(els.totalBaseKeys, model.equalCount);
  setText(els.totalConcluidos, summary.concluded);
  setText(els.totalAndamento, summary.inProgress);
  setText(els.totalNaoIniciado, summary.notStarted);
  setText(els.totalCriticos, summary.critical);
  setText(els.totalSoEmpresa, model.onlyEmpresa.length);
  setText(els.totalSoConsultoria, model.onlyConsultoria.length);
  setText(els.totalRicesEmpresa, model.empresaTotal);
  setText(els.totalRicesConsultoria, model.consultoriaTotal);
  if (els.globalPercent) els.globalPercent.textContent = `${summary.percent}%`;
  if (els.ringProgress) els.ringProgress.style.strokeDashoffset = `${RING_CIRCUMFERENCE * (1 - summary.percent / 100)}`;

  const message = [...EXECUTIVE_MESSAGES].reverse().find(item => summary.percent >= item.threshold) || EXECUTIVE_MESSAGES[0];
  if (els.statusExecutivo) els.statusExecutivo.textContent = message.status;
  if (els.headlineCallout) els.headlineCallout.textContent = `${message.title} Total sem repetir nas duas bases: ${formatNumber(model.totalUnion)} RICE(s).`;
  if (els.headlinePill) els.headlinePill.textContent = 'RICE consolidado';
}

function renderPlannerMetrics(rows, origem, targetEl, totalUnico) {
  if (!targetEl) return;
  const total = rows.length;
  const concluded = rows.filter(row => isConcluded(row.statusOriginal)).length;
  const inProgress = rows.filter(row => isInProgress(row.statusOriginal)).length;
  const notStarted = rows.filter(row => isNotStarted(row.statusOriginal)).length;
  const critical = rows.filter(row => isCritical(row)).length;
  targetEl.innerHTML = `
    <div class="planner-header">
      <div class="planner-title"><strong>${escapeHtml(origem)}</strong><span>${formatNumber(totalUnico)} RICE(s) únicos · ${formatNumber(total)} casados</span></div>
      <div class="planner-percent">${getPercent(concluded, total)}%</div>
    </div>
    <div class="planner-status-grid">
      <div class="planner-kpi"><small>Concluídos</small><strong>${formatNumber(concluded)}</strong></div>
      <div class="planner-kpi"><small>Em andamento</small><strong>${formatNumber(inProgress)}</strong></div>
      <div class="planner-kpi"><small>Não iniciados</small><strong>${formatNumber(notStarted)}</strong></div>
      <div class="planner-kpi"><small>Críticos</small><strong>${formatNumber(critical)}</strong></div>
      <div class="planner-kpi"><small>Total único</small><strong>${formatNumber(totalUnico)}</strong></div>
    </div>`;
}

function renderLeaderboard(rows) {
  if (!els.leaderboard) return;
  const grouped = new Map();
  rows.forEach(row => {
    const people = row.responsaveisTodos.length ? row.responsaveisTodos : [row.responsavel];
    people.forEach(person => {
      if (!person || person === 'Sem responsável') return;
      const key = `${person}|||${row.origem}`;
      if (!grouped.has(key)) grouped.set(key, { responsavel: person, origem: row.origem, total: 0, concluded: 0 });
      const item = grouped.get(key);
      item.total += 1;
      if (isConcluded(row.statusOriginal)) item.concluded += 1;
    });
  });
  const ranking = Array.from(grouped.values()).map(item => ({ ...item, percent: getPercent(item.concluded, item.total) }))
    .sort((a, b) => b.percent - a.percent || b.concluded - a.concluded || a.responsavel.localeCompare(b.responsavel, 'pt-BR')).slice(0, 10);
  els.leaderboard.innerHTML = ranking.length ? ranking.map((item, index) => `
    <div class="leader-row">
      <div class="place ${index < 3 ? `top-${index + 1}` : ''}">${index + 1}</div>
      <div><div class="leader-name">${escapeHtml(item.responsavel)}</div><div class="leader-meta">${escapeHtml(item.origem)} · ${formatNumber(item.concluded)} concluídos de ${formatNumber(item.total)}</div></div>
      <div class="leader-score"><strong>${item.percent}%</strong><span>aproveitamento</span></div>
    </div>`).join('') : '<div class="empty-state">Sem responsáveis válidos para exibir.</div>';
}

function renderStatusBars(summary) {
  if (!els.statusBars) return;
  const statuses = [
    { label: 'Concluído', value: summary.concluded },
    { label: 'Em andamento', value: summary.inProgress },
    { label: 'Não iniciado', value: summary.notStarted },
    { label: 'Crítico', value: summary.critical }
  ];
  els.statusBars.innerHTML = statuses.map(item => {
    const percent = getPercent(item.value, summary.total);
    return `<div class="status-item"><div class="status-head"><strong>${item.label}</strong><span>${formatNumber(item.value)} · ${percent}%</span></div><div class="status-track"><div class="status-fill" style="width:${percent}%;"></div></div></div>`;
  }).join('');
}

function renderComparisonBoard(rows) {
  if (!els.comparisonBoard) return;
  const empresa = rows.filter(row => row.origem === 'Empresa');
  const consultoria = rows.filter(row => row.origem === 'Consultoria');
  const statuses = [
    { label: 'Concluído', fn: row => isConcluded(row.statusOriginal) },
    { label: 'Em andamento', fn: row => isInProgress(row.statusOriginal) },
    { label: 'Não iniciado', fn: row => isNotStarted(row.statusOriginal) },
    { label: 'Bloqueado/Atrasado', fn: row => isCritical(row) }
  ];
  els.comparisonBoard.innerHTML = statuses.map(item => {
    const e = empresa.filter(item.fn).length;
    const c = consultoria.filter(item.fn).length;
    const ep = getPercent(e, empresa.length);
    const cp = getPercent(c, consultoria.length);
    return `<div class="comparison-row"><div class="comparison-label">${item.label}</div>
      <div class="comparison-side comparison-side-inline"><span>Empresa</span><div class="comparison-bar-wrap"><div class="status-track"><div class="status-fill" style="width:${ep}%;"></div></div><strong class="comparison-inline-total">${formatNumber(e)}</strong></div></div>
      <div class="comparison-side comparison-side-inline"><span>Consultoria</span><div class="comparison-bar-wrap"><div class="status-track"><div class="status-fill" style="width:${cp}%;"></div></div><strong class="comparison-inline-total">${formatNumber(c)}</strong></div></div>
    </div>`;
  }).join('');
}

function setupRiceFilter(pairs) {
  state.pairs = pairs;
  state.riceFilter = new Set(pairs.map(pair => pair.compareKey));
  if (els.riceFilterSelect) {
    els.riceFilterSelect.innerHTML = pairs.map(pair => `<option value="${escapeHtml(pair.compareKey)}" selected>${escapeHtml(pair.compareDisplay)}</option>`).join('');
  }
  renderRiceBoard();
}

function bindRiceFilter() {
  if (els.riceFilterSelect) els.riceFilterSelect.addEventListener('change', () => {
    state.riceFilter = new Set(Array.from(els.riceFilterSelect.selectedOptions).map(option => option.value));
    renderRiceBoard();
  });
  if (els.btnRiceFilterAll) els.btnRiceFilterAll.addEventListener('click', () => setRiceFilter(true));
  if (els.btnRiceFilterClear) els.btnRiceFilterClear.addEventListener('click', () => setRiceFilter(false));
}

function setRiceFilter(selected) {
  if (!els.riceFilterSelect) return;
  Array.from(els.riceFilterSelect.options).forEach(option => { option.selected = selected; });
  state.riceFilter = selected ? new Set(state.pairs.map(pair => pair.compareKey)) : new Set();
  renderRiceBoard();
}

function renderRiceBoard() {
  if (!els.keyBoard) return;
  const pairs = state.pairs.filter(pair => state.riceFilter.has(pair.compareKey));
  if (els.riceFilterStatus) els.riceFilterStatus.textContent = `${formatNumber(pairs.length)} de ${formatNumber(state.pairs.length)} RICE(s) exibidos.`;
  els.keyBoard.innerHTML = pairs.length ? pairs.map(pair => `
    <div class="area-card">
      <div class="area-top"><div class="area-title">${escapeHtml(pair.compareDisplay)}</div><div class="area-badge">${pair.progressPercent}%</div></div>
      <div class="status-track" style="margin-top:12px;"><div class="status-fill" style="width:${pair.progressPercent}%;"></div></div>
      <div class="area-stats"><span>Empresa: ${escapeHtml(pair.empresa.statusOriginal)}</span><span>Consultoria: ${escapeHtml(pair.consultoria.statusOriginal)}</span></div>
      <div class="area-stats"><span>Status do par</span><span>${escapeHtml(pair.pairStatus)}</span></div>
    </div>`).join('') : '<div class="empty-state">Nenhum RICE selecionado.</div>';
}

function renderDivergenceBoard(onlyEmpresa, onlyConsultoria) {
  if (!els.divergenceBoard) return;
  const groups = [
    { title: 'Só na Empresa', rows: onlyEmpresa, subtitle: 'Existe internamente e não apareceu na Consultoria.' },
    { title: 'Só na Consultoria', rows: onlyConsultoria, subtitle: 'Existe na Consultoria e não apareceu internamente.' }
  ];
  els.divergenceBoard.innerHTML = groups.map(group => `
    <div class="area-card"><div class="area-top"><div class="area-title">${group.title}</div><div class="area-badge">${formatNumber(group.rows.length)}</div></div>
    <div class="area-stats" style="display:block;"><span>${group.subtitle}</span></div>
    <div class="chip-board" style="margin-top:12px;">${group.rows.slice(0, 20).map(row => `<span class="category-chip">${escapeHtml(row.compareDisplay)}</span>`).join('') || '<span class="category-chip">Nenhum</span>'}</div></div>`).join('');
}

function renderDivergencePanels(onlyEmpresa, onlyConsultoria) {
  renderDivergencePanel(els.divergenceEmpresaList, onlyEmpresa, 'Nenhuma RICE exclusiva da Empresa.');
  renderDivergencePanel(els.divergenceConsultoriaList, onlyConsultoria, 'Nenhuma RICE exclusiva da Consultoria.');
}

function renderDivergencePanel(target, rows, emptyMessage) {
  if (!target) return;
  target.innerHTML = rows.length ? rows.slice(0, 50).map(row => `
    <div class="divergence-compact-card"><div class="divergence-main"><div class="divergence-code">${escapeHtml(row.compareDisplay)}</div>
    <div class="divergence-desc">${escapeHtml(row.tarefa || '-')}</div><div class="divergence-owner"><strong>Atribuído:</strong> ${escapeHtml(row.responsavel || '-')}</div></div>
    <div class="divergence-phases"><div class="divergence-phase-row"><span>Registros</span><strong>${formatNumber(row.qtdRegistros)}</strong></div><div class="divergence-phase-row"><span>Status</span><strong>${escapeHtml(row.statusOriginal)}</strong></div></div></div>`).join('') : `<div class="empty-state">${emptyMessage}</div>`;
}

function renderDivergenceTable(onlyEmpresa, onlyConsultoria) {
  if (!els.divergenceTable) return;
  const rows = [...onlyEmpresa, ...onlyConsultoria].sort((a, b) => a.origem.localeCompare(b.origem, 'pt-BR') || a.compareDisplay.localeCompare(b.compareDisplay, 'pt-BR'));
  els.divergenceTable.innerHTML = rows.length ? rows.map(row => `
    <tr><td>${originBadge(row.origem)}</td><td>${escapeHtml(row.compareDisplay)}</td><td>${escapeHtml(row.id || '-')}</td><td>${escapeHtml(row.tarefa || '-')}</td><td>${escapeHtml(row.responsavel || '-')}</td><td>${statusPill(row)}</td><td>${row.origem === 'Empresa' ? 'Existe na Empresa e não apareceu na Consultoria.' : 'Existe na Consultoria e não apareceu na Empresa.'}</td></tr>`).join('') : '<tr><td colspan="7" class="empty-cell">Nenhuma divergência encontrada.</td></tr>';
}

function setupFocusTable(pairs) {
  state.focusRows = pairs.filter(pair => pair.pairStatus === 'Crítico' || pair.pairStatus === 'Não iniciado').flatMap(pair => [
    { ...pair.empresa, pairStatus: pair.pairStatus },
    { ...pair.consultoria, pairStatus: pair.pairStatus }
  ]).slice(0, 200);
  populateFocusFilters();
  renderFocusTable();
}

function bindFocusFilters() {
  const map = { origem: els.focusFilterOrigem, rice: els.focusFilterRice, fase: els.focusFilterFase, tarefa: els.focusFilterTarefa, responsavel: els.focusFilterResponsavel, status: els.focusFilterStatus };
  Object.entries(map).forEach(([key, element]) => {
    if (!element) return;
    element.addEventListener('change', event => {
      state.focusFilters[key] = event.target.value || '';
      renderFocusTable();
    });
  });
}

function populateFocusFilters() {
  const defs = [
    ['origem', els.focusFilterOrigem, row => row.origem],
    ['rice', els.focusFilterRice, row => row.compareDisplay],
    ['fase', els.focusFilterFase, row => row.categoria],
    ['tarefa', els.focusFilterTarefa, row => row.tarefa],
    ['responsavel', els.focusFilterResponsavel, row => row.responsavel],
    ['status', els.focusFilterStatus, row => row.statusOriginal]
  ];
  defs.forEach(([key, element, fn]) => {
    if (!element) return;
    const values = [...new Set(state.focusRows.map(fn).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
    element.innerHTML = '<option value="">Todos</option>' + values.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join('');
  });
}

function renderFocusTable() {
  if (!els.focusTable) return;
  const rows = state.focusRows.filter(row => {
    const f = state.focusFilters;
    if (f.origem && row.origem !== f.origem) return false;
    if (f.rice && row.compareDisplay !== f.rice) return false;
    if (f.fase && row.categoria !== f.fase) return false;
    if (f.tarefa && row.tarefa !== f.tarefa) return false;
    if (f.responsavel && row.responsavel !== f.responsavel) return false;
    if (f.status && row.statusOriginal !== f.status) return false;
    return true;
  });
  els.focusTable.innerHTML = rows.length ? rows.map(row => `<tr><td>${originBadge(row.origem)}</td><td>${escapeHtml(row.compareDisplay)}</td><td>${escapeHtml(row.categoria || '-')}</td><td>${escapeHtml(row.tarefa || '-')}</td><td>${escapeHtml(row.responsavel || '-')}</td><td>${statusPill(row)}</td></tr>`).join('') : '<tr><td colspan="6" class="empty-cell">Nenhum item encontrado.</td></tr>';
}

function getValue(row, key) {
  if (!row) return '';
  const keys = Array.isArray(key) ? key : [key];
  for (const candidate of keys) {
    if (!candidate) continue;
    if (Object.prototype.hasOwnProperty.call(row, candidate)) return row[candidate] ?? '';
    const normalizedCandidate = normalizeFieldName(candidate);
    const found = Object.keys(row).find(k => normalizeFieldName(k) === normalizedCandidate);
    if (found) return row[found] ?? '';
  }
  return '';
}

function normalizeFieldName(value) {
  return normalize(value).replace(/\s+/g, ' ');
}

function extractRiceCode(taskName) {
  const raw = String(taskName || '').trim();
  const match = raw.match(/^([A-Za-z0-9]+(?:[.#_-][A-Za-z0-9]+)*(?:\s+\d+)?(?:#\d+)?)/);
  return match ? match[1].toUpperCase() : '';
}

function normalizeRiceKey(value) {
  return normalize(value).toUpperCase().replace(/[^A-Z0-9#._]/g, '');
}

function splitPeople(value) {
  return String(value || '').split(';').map(item => item.trim()).filter(Boolean);
}

function firstResponsible(value) {
  return splitPeople(value)[0] || 'Sem responsável';
}

function normalize(text = '') {
  return String(text).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function containsAny(text, terms) {
  const normalized = normalize(text);
  return terms.some(term => normalized.includes(normalize(term)));
}

function isConcluded(status) {
  return containsAny(status, DASH_CONFIG.statusRules.concludedContains);
}

function isInProgress(status) {
  return containsAny(status, DASH_CONFIG.statusRules.inProgressContains);
}

function isNotStarted(status) {
  return containsAny(status, DASH_CONFIG.statusRules.notStartedContains);
}

function isCritical(row) {
  return row.atrasadoFlag === true || containsAny(row.statusOriginal, DASH_CONFIG.statusRules.blockedContains);
}

function isDelayedValue(value) {
  return ['true', 'verdadeiro', 'sim', '1', 'yes'].includes(normalize(value));
}

function getPercent(value, total) {
  return total ? Math.round((Number(value || 0) / Number(total || 0)) * 100) : 0;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('pt-BR');
}

function setText(element, value) {
  if (element) element.textContent = formatNumber(value);
}

function originBadge(origin) {
  const className = origin === 'Consultoria' ? 'origin-consultoria' : 'origin-empresa';
  return `<span class="origin-badge ${className}">${escapeHtml(origin)}</span>`;
}

function statusPill(row) {
  let className = 'status-outro';
  if (row.atrasadoFlag) className = 'status-atrasado';
  else if (containsAny(row.statusOriginal, DASH_CONFIG.statusRules.blockedContains)) className = 'status-bloqueado';
  else if (isConcluded(row.statusOriginal)) className = 'status-concluido';
  else if (isInProgress(row.statusOriginal)) className = 'status-andamento';
  else if (isNotStarted(row.statusOriginal)) className = 'status-nao-iniciado';
  return `<span class="status-pill ${className}">${escapeHtml(row.statusOriginal || '-')}</span>`;
}

function setGeneratedAt(message) {
  if (els.dataAtualizacao) els.dataAtualizacao.textContent = message;
}

function renderFatal(message) {
  setGeneratedAt(message);
  if (els.statusExecutivo) els.statusExecutivo.textContent = 'Falha de carga';
  if (els.headlineCallout) els.headlineCallout.textContent = message;
  [els.empresaMetrics, els.consultoriaMetrics, els.leaderboard, els.statusBars, els.comparisonBoard, els.keyBoard, els.divergenceBoard, els.divergenceEmpresaList, els.divergenceConsultoriaList].forEach(el => {
    if (el) el.innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
  });
  if (els.focusTable) els.focusTable.innerHTML = `<tr><td colspan="6" class="empty-cell">${escapeHtml(message)}</td></tr>`;
  if (els.divergenceTable) els.divergenceTable.innerHTML = `<tr><td colspan="7" class="empty-cell">${escapeHtml(message)}</td></tr>`;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
