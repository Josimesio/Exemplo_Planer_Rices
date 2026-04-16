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
  heroDescription: document.getElementById('heroDescription'),
  modeDescription: document.getElementById('modeDescription'),
  baseKeyLabel: document.getElementById('baseKeyLabel'),
  commonKeysTitle: document.getElementById('commonKeysTitle'),
  commonKeys: document.getElementById('commonKeys') || document.getElementById('commonRices'),
  empresaHint: document.getElementById('empresaHint'),
  consultoriaHint: document.getElementById('consultoriaHint'),
  empresaMetrics: document.getElementById('empresaMetrics'),
  consultoriaMetrics: document.getElementById('consultoriaMetrics'),
  leaderboard: document.getElementById('leaderboard'),
  statusBars: document.getElementById('statusBars'),
  comparisonBoard: document.getElementById('comparisonBoard'),
  keyBoardTitle: document.getElementById('keyBoardTitle'),
  keyBoard: document.getElementById('keyBoard') || document.getElementById('riceBoard'),
  riceFilterSelect: document.getElementById('riceFilterSelect'),
  riceFilterStatus: document.getElementById('riceFilterStatus'),
  btnRiceFilterAll: document.getElementById('btnRiceFilterAll'),
  btnRiceFilterClear: document.getElementById('btnRiceFilterClear'),
  focusFilterOrigem: document.getElementById('focusFilterOrigem'),
  focusFilterRice: document.getElementById('focusFilterRice'),
  focusFilterFase: document.getElementById('focusFilterFase'),
  focusFilterTarefa: document.getElementById('focusFilterTarefa'),
  focusFilterResponsavel: document.getElementById('focusFilterResponsavel'),
  focusFilterStatus: document.getElementById('focusFilterStatus'),
  focusTable: document.getElementById('focusTable'),
  divergenceBoard: document.getElementById('divergenceBoard'),
  divergenceTable: document.getElementById('divergenceTable'),
  divergenceEmpresaList: document.getElementById('divergenceEmpresaList'),
  divergenceConsultoriaList: document.getElementById('divergenceConsultoriaList'),
  manualLoaderCard: document.getElementById('manualLoaderCard'),
  manualLoaderMessage: document.getElementById('manualLoaderMessage'),
  fileEmpresa: document.getElementById('fileEmpresa'),
  fileConsultoria: document.getElementById('fileConsultoria'),
  btnCarregarArquivos: document.getElementById('btnCarregarArquivos')
};

const RING_CIRCUMFERENCE = 301.59;
const IS_LOCAL_FILE = window.location.protocol === 'file:';
const focusTableState = {
  rows: [],
  filters: {
    origem: '',
    rice: '',
    fase: '',
    tarefa: '',
    responsavel: '',
    status: ''
  }
};

const dashboardState = {
  ricePairsAll: [],
  focusRowsAll: []
};

const EXECUTIVE_MESSAGES = [
  { threshold: 0, status: 'Leitura inicial', title: 'A comparação foi preparada com base na configuração ativa do painel.' },
  { threshold: 20, status: 'Avanço inicial', title: 'Existe progresso na base comparável, mas ainda há volume relevante fora da linha de chegada.' },
  { threshold: 40, status: 'Tração consistente', title: 'O programa mostra ritmo real de entrega na régua configurada.' },
  { threshold: 60, status: 'Execução madura', title: 'O avanço já é consistente. O foco agora é limpar gargalos e sustentar previsibilidade.' },
  { threshold: 80, status: 'Alta performance', title: 'A base comparável está forte. A prioridade passa a ser eliminar resíduos críticos.' },
  { threshold: 100, status: 'Ciclo concluído', title: 'Todos os itens da base comparável foram concluídos.' }
];


async function ensureXlsxLibrary() {
  if (typeof XLSX !== 'undefined') return true;

  const fallbacks = [
    'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
  ];

  for (const url of fallbacks) {
    try {
      await loadExternalScript(url);
      if (typeof XLSX !== 'undefined') return true;
    } catch (error) {
      console.warn('Falha ao carregar biblioteca XLSX em', url, error);
    }
  }

  return typeof XLSX !== 'undefined';
}

function loadExternalScript(url) {
  return new Promise((resolve, reject) => {
    const existing = Array.from(document.scripts).find(script => script.src === url);
    if (existing) {
      if (existing.dataset.loaded === 'true') {
        resolve();
        return;
      }
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error(`Falha ao carregar ${url}`)), { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = url;
    script.async = true;
    script.onload = () => {
      script.dataset.loaded = 'true';
      resolve();
    };
    script.onerror = () => reject(new Error(`Falha ao carregar ${url}`));
    document.head.appendChild(script);
  });
}

document.addEventListener('DOMContentLoaded', async () => {
  bindManualLoader();
  bindRiceBoardFilter();
  bindFocusFilters();

  if (typeof DASH_CONFIG === 'undefined') {
    handleFatal('config.js não foi carregado. O painel depende desse arquivo para saber quais planilhas ler.');
    return;
  }

  const xlsxReady = await ensureXlsxLibrary();
  if (!xlsxReady) {
    handleFatal('A biblioteca de leitura das planilhas não foi carregada. Verifique se o servidor permite acessar CDNs externas ou publique também a biblioteca XLSX junto com o site.');
    return;
  }

  applyConfigToLayout();

  if (IS_LOCAL_FILE) {
    handleFatal('Este pacote foi preparado para hospedagem online. Publique o dashboard e as duas planilhas no mesmo diretório do site e abra por HTTP/HTTPS.');
    return;
  }

  loadWorkbookData();
});

function bindManualLoader() {
  if (!els.btnCarregarArquivos) return;
  els.btnCarregarArquivos.addEventListener('click', () => {
    loadWorkbookDataFromManualSelection();
  });
}

function bindRiceBoardFilter() {
  if (els.riceFilterSelect) {
    els.riceFilterSelect.addEventListener('change', () => {
      applyRiceBoardFilter();
    });
  }

  if (els.btnRiceFilterAll) {
    els.btnRiceFilterAll.addEventListener('click', () => {
      setRiceFilterSelection(true);
      applyRiceBoardFilter();
    });
  }

  if (els.btnRiceFilterClear) {
    els.btnRiceFilterClear.addEventListener('click', () => {
      setRiceFilterSelection(false);
      applyRiceBoardFilter();
    });
  }
}

function bindFocusFilters() {
  const filterMap = {
    origem: els.focusFilterOrigem,
    rice: els.focusFilterRice,
    fase: els.focusFilterFase,
    tarefa: els.focusFilterTarefa,
    responsavel: els.focusFilterResponsavel,
    status: els.focusFilterStatus
  };

  Object.entries(filterMap).forEach(([key, element]) => {
    if (!element) return;
    element.addEventListener('change', event => {
      focusTableState.filters[key] = String(event.target.value || '');
      renderFilteredFocusTable();
    });
  });
}

function populateFocusFilterOptions() {
  const definitions = [
    {
      key: 'origem',
      element: els.focusFilterOrigem,
      values: focusTableState.rows.map(row => row.origem || '-'),
      sortFn: (a, b) => String(a).localeCompare(String(b), 'pt-BR')
    },
    {
      key: 'rice',
      element: els.focusFilterRice,
      values: focusTableState.rows.map(row => row.compareDisplay || '-'),
      sortFn: (a, b) => String(a).localeCompare(String(b), 'pt-BR')
    },
    {
      key: 'fase',
      element: els.focusFilterFase,
      values: focusTableState.rows.map(row => row.categoria || '-'),
      sortFn: (a, b) => extractStageOrder(a) - extractStageOrder(b) || String(a).localeCompare(String(b), 'pt-BR')
    },
    {
      key: 'tarefa',
      element: els.focusFilterTarefa,
      values: focusTableState.rows.map(row => formatTaskForDisplay(row) || '-'),
      sortFn: (a, b) => String(a).localeCompare(String(b), 'pt-BR')
    },
    {
      key: 'responsavel',
      element: els.focusFilterResponsavel,
      values: focusTableState.rows.map(row => row.responsavel || '-'),
      sortFn: (a, b) => String(a).localeCompare(String(b), 'pt-BR')
    },
    {
      key: 'status',
      element: els.focusFilterStatus,
      values: focusTableState.rows.map(row => row.statusOriginal || '-'),
      sortFn: (a, b) => String(a).localeCompare(String(b), 'pt-BR')
    }
  ];

  definitions.forEach(definition => {
    if (!definition.element) return;

    const selectedValue = focusTableState.filters[definition.key] || '';
    const uniqueValues = [...new Set(definition.values.filter(Boolean))].sort(definition.sortFn);
    const availableSet = new Set(uniqueValues);

    if (selectedValue && !availableSet.has(selectedValue)) {
      focusTableState.filters[definition.key] = '';
    }

    definition.element.innerHTML = [
      '<option value="">Todos</option>',
      ...uniqueValues.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    ].join('');
    definition.element.value = focusTableState.filters[definition.key] || '';
  });
}

function getFilteredFocusRows() {
  return focusTableState.rows.filter(row => {
    if (focusTableState.filters.origem && row.origem !== focusTableState.filters.origem) return false;
    if (focusTableState.filters.rice && row.compareDisplay !== focusTableState.filters.rice) return false;
    if (focusTableState.filters.fase && (row.categoria || '-') !== focusTableState.filters.fase) return false;
    if (focusTableState.filters.tarefa && formatTaskForDisplay(row) !== focusTableState.filters.tarefa) return false;
    if (focusTableState.filters.responsavel && (row.responsavel || '-') !== focusTableState.filters.responsavel) return false;
    if (focusTableState.filters.status && (row.statusOriginal || '-') !== focusTableState.filters.status) return false;
    return true;
  });
}

function renderFilteredFocusTable() {
  if (!els.focusTable) return;

  const rows = getFilteredFocusRows();

  els.focusTable.innerHTML = rows.length
    ? rows.map(row => `
      <tr>
        <td>${originBadge(row.origem)}</td>
        <td>${escapeHtml(row.compareDisplay || '-')}</td>
        <td>${escapeHtml(row.categoria || '-')}</td>
        <td>${escapeHtml(formatTaskForDisplay(row))}</td>
        <td>${escapeHtml(row.responsavel || '-')}</td>
        <td>${statusPill(row)}</td>
      </tr>`).join('')
    : '<tr><td colspan="6" class="empty-cell">Nenhum item encontrado com os filtros aplicados.</td></tr>';
}

function applyConfigToLayout() {
  const mode = getModeConfig();
  if (els.heroDescription) els.heroDescription.textContent = mode.heroDescription;
  if (els.modeDescription) els.modeDescription.innerHTML = `Modo atual: <strong>${mode.label}</strong>.`;
  if (els.baseKeyLabel) els.baseKeyLabel.textContent = mode.keyLabel;
  if (els.commonKeysTitle) els.commonKeysTitle.textContent = mode.commonKeysTitle;
  if (els.keyBoardTitle) els.keyBoardTitle.textContent = mode.keyBoardTitle;
  if (els.empresaHint) els.empresaHint.textContent = DASH_CONFIG.files[0]?.label || 'fonte 1';
  if (els.consultoriaHint) els.consultoriaHint.textContent = DASH_CONFIG.files[1]?.label || 'fonte 2';
}

async function loadWorkbookData() {
  try {
    setGeneratedAt('Lendo planilhas configuradas...');
    hideManualLoader();

    const datasets = [];
    for (const plan of DASH_CONFIG.files) {
      const workbook = await fetchWorkbook(plan.fileName);
      datasets.push({ plan, workbook });
    }

    processDatasets(datasets);
  } catch (error) {
    console.error(error);
    const details = error && error.message ? error.message : 'erro não identificado';
    const msg = `Falha ao carregar as planilhas automaticamente. Verifique se as duas planilhas foram publicadas na mesma pasta do dashboard com os nomes esperados. Detalhe técnico: ${details}`;
    setGeneratedAt(msg);
    renderEmptyStates(msg);
  }
}

async function loadWorkbookDataFromManualSelection() {
  try {
    if (!els.fileEmpresa?.files?.[0] || !els.fileConsultoria?.files?.[0]) {
      showManualLoader('Selecione as duas planilhas antes de carregar: uma da Empresa e uma da Consultoria.');
      setGeneratedAt('Faltam arquivos para a leitura manual.');
      return;
    }

    setGeneratedAt('Lendo planilhas selecionadas manualmente...');

    const datasets = [];
    const manualFiles = [
      { plan: DASH_CONFIG.files[0], file: els.fileEmpresa.files[0] },
      { plan: DASH_CONFIG.files[1], file: els.fileConsultoria.files[0] }
    ];

    for (const item of manualFiles) {
      const workbook = await readWorkbookFromLocalFile(item.file);
      datasets.push({ plan: item.plan, workbook });
    }

    processDatasets(datasets);
    hideManualLoader();
  } catch (error) {
    console.error(error);
    showManualLoader('A leitura manual também falhou. Confirme se os arquivos escolhidos são realmente as duas planilhas .xlsx esperadas.');
    setGeneratedAt('Falha ao ler as planilhas selecionadas manualmente.');
    renderEmptyStates('Falha ao carregar as planilhas selecionadas.');
  }
}

function processDatasets(datasets) {
  const allRows = [];

  datasets.forEach(({ plan, workbook }) => {
    const sheet = workbook.Sheets[DASH_CONFIG.sheetName] || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
      throw new Error(`Nenhuma aba válida encontrada em ${plan.fileName}.`);
    }

    const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    jsonRows.forEach(row => {
      row.__origem = plan.origem;
      row.__arquivo = plan.fileName;
    });
    allRows.push(...jsonRows);
  });

  const normalized = normalizeRows(allRows);
  const consolidated = consolidateCards(normalized);
  const comparisonModel = buildComparisonModel(consolidated);

  renderDashboard(comparisonModel);

  setGeneratedAt(
    `Leitura concluída: ${comparisonModel.pairs.length} RICE(s) comparáveis 1x1, ${comparisonModel.onlyEmpresa.length} só na Empresa e ${comparisonModel.onlyConsultoria.length} só na Consultoria.`
  );
}

async function fetchWorkbook(fileName) {
  const candidates = buildWorkbookCandidates(fileName);
  let lastError = null;

  for (const candidate of candidates) {
    try {
      const response = await fetch(candidate, { cache: 'no-store' });
      if (!response.ok) {
        lastError = new Error(`Falha ao carregar ${candidate}: HTTP ${response.status}`);
        continue;
      }
      const arrayBuffer = await response.arrayBuffer();
      return XLSX.read(arrayBuffer, { type: 'array' });
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error(`Falha ao carregar ${fileName}.`);
}

function buildWorkbookCandidates(fileName) {
  const trimmed = String(fileName || '').trim();
  const stamp = Date.now();
  const candidates = [
    `${encodeURI(trimmed)}?t=${stamp}`,
    `${trimmed}?t=${stamp}`
  ];

  const simplified = trimmed.replace(/'/g, '').replace(/\s+/g, ' ').trim();
  if (simplified && simplified !== trimmed) {
    candidates.push(`${encodeURI(simplified)}?t=${stamp}`);
    candidates.push(`${simplified}?t=${stamp}`);
  }

  return [...new Set(candidates)];
}

function readWorkbookFromLocalFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const arrayBuffer = event.target?.result;
        resolve(XLSX.read(arrayBuffer, { type: 'array' }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error(`Falha ao ler o arquivo local ${file?.name || ''}.`));
    reader.readAsArrayBuffer(file);
  });
}

function normalizeRows(rows) {
  const fields = DASH_CONFIG.fields;

  return rows
    .filter(row => String(getValue(row, fields.taskName)).trim() !== '')
    .map(row => {
      const tarefa = getValue(row, fields.taskName);
      const categoria = getValue(row, fields.category) || 'Sem categoria';
      const riceCode = extractRiceCode(tarefa);
      const cardSignature = buildTaskSignature(tarefa, riceCode);

      return {
        origem: row.__origem || 'Empresa',
        arquivo: row.__arquivo || '',
        id: getValue(row, fields.taskId),
        tarefa,
        riceCode,
        cardSignature,
        cardGroupKey: buildCardGroupKey({ origem: row.__origem || 'Empresa', riceCode, cardSignature, tarefa }),
        categoria,
        stageOrder: extractStageOrder(categoria),
        compareKey: buildCompareKey({ riceCode, categoria }),
        compareDisplay: buildCompareDisplay({ riceCode, categoria }),
        statusOriginal: getValue(row, fields.status) || 'Sem status',
        prioridade: getValue(row, fields.priority) || 'Sem prioridade',
        responsavel: firstResponsible(getValue(row, fields.responsible)),
        responsaveisTodos: splitPeople(getValue(row, fields.responsible)),
        atrasadoFlag: isDelayedValue(getValue(row, fields.delayed))
      };
    })
    .filter(row => row.compareKey !== '');
}

function consolidateCards(rows) {
  const grouped = new Map();

  rows.forEach(row => {
    const key = row.cardGroupKey || `${row.origem}|||${row.compareKey}|||${normalize(row.tarefa)}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  });

  return [...grouped.values()].map(groupRows => consolidateCardGroup(groupRows));
}

function consolidateCardGroup(groupRows) {
  const representative = pickRepresentativeRow(groupRows);
  const responsaveisTodos = mergeResponsibles(groupRows);
  const historicoCategorias = groupRows
    .map(row => row.categoria)
    .filter((value, index, array) => value && array.indexOf(value) === index)
    .sort((a, b) => extractStageOrder(a) - extractStageOrder(b) || String(a).localeCompare(String(b), 'pt-BR'));

  return {
    ...representative,
    responsaveisTodos,
    responsavel: responsaveisTodos[0] || representative.responsavel || 'Sem responsável',
    atrasadoFlag: groupRows.some(row => row.atrasadoFlag === true),
    rawRowsCount: groupRows.length,
    rawCardsCount: 1,
    historicoCategorias
  };
}

function pickRepresentativeRow(rows) {
  const ordered = [...rows].sort((a, b) =>
    (a.stageOrder - b.stageOrder) ||
    (getRepresentativePriority(b) - getRepresentativePriority(a)) ||
    String(a.categoria || '').localeCompare(String(b.categoria || ''), 'pt-BR') ||
    String(a.id || '').localeCompare(String(b.id || ''), 'pt-BR')
  );

  const nonConcluded = ordered.filter(row => !isConcluded(row.statusOriginal));

  if (nonConcluded.length) {
    const currentStage = Math.min(...nonConcluded.map(row => row.stageOrder));
    const candidates = nonConcluded.filter(row => row.stageOrder === currentStage);
    return pickBestCandidate(candidates);
  }

  const latestStage = Math.max(...ordered.map(row => row.stageOrder));
  const candidates = ordered.filter(row => row.stageOrder === latestStage);
  return pickBestCandidate(candidates);
}

function pickBestCandidate(rows) {
  return [...rows].sort((a, b) =>
    (getRepresentativePriority(b) - getRepresentativePriority(a)) ||
    String(a.categoria || '').localeCompare(String(b.categoria || ''), 'pt-BR') ||
    String(a.id || '').localeCompare(String(b.id || ''), 'pt-BR')
  )[0];
}

function getRepresentativePriority(row) {
  if (isCritical(row)) return 4;
  if (isInProgress(row.statusOriginal)) return 3;
  if (isNotStarted(row.statusOriginal)) return 2;
  if (isConcluded(row.statusOriginal)) return 1;
  return 0;
}

function mergeResponsibles(rows) {
  const merged = [];
  const seen = new Set();

  rows.forEach(row => {
    const people = row.responsaveisTodos.length ? row.responsaveisTodos : [row.responsavel];
    people.forEach(person => {
      if (!person || person === 'Sem responsável' || seen.has(person)) return;
      seen.add(person);
      merged.push(person);
    });
  });

  return merged;
}

function buildComparisonModel(rows) {
  const grouped = new Map();

  rows.forEach(row => {
    const key = row.compareKey;
    if (!grouped.has(key)) {
      grouped.set(key, {
        compareKey: key,
        compareDisplay: row.compareDisplay,
        byOrigin: { Empresa: [], Consultoria: [] }
      });
    }

    const item = grouped.get(key);
    if (!item.byOrigin[row.origem]) item.byOrigin[row.origem] = [];
    item.byOrigin[row.origem].push(row);
    if (!item.compareDisplay) item.compareDisplay = row.compareDisplay;
  });

  const pairs = [];
  const onlyEmpresa = [];
  const onlyConsultoria = [];

  [...grouped.values()]
    .sort((a, b) => String(a.compareDisplay || '').localeCompare(String(b.compareDisplay || ''), 'pt-BR'))
    .forEach(item => {
      const empresaRows = item.byOrigin.Empresa || [];
      const consultoriaRows = item.byOrigin.Consultoria || [];

      const empresa = empresaRows.length ? consolidateCompareKeyRows(empresaRows, item.compareDisplay, 'Empresa') : null;
      const consultoria = consultoriaRows.length ? consolidateCompareKeyRows(consultoriaRows, item.compareDisplay, 'Consultoria') : null;

      if (empresa && consultoria) {
        pairs.push({
          compareKey: item.compareKey,
          compareDisplay: item.compareDisplay,
          empresa,
          consultoria,
          pairStatus: buildPairStatus(empresa, consultoria),
          progressPercent: buildPairProgressPercent(empresa, consultoria)
        });
      } else if (empresa) {
        onlyEmpresa.push(empresa);
      } else if (consultoria) {
        onlyConsultoria.push(consultoria);
      }
    });

  return {
    pairs,
    comparableRows: pairs.flatMap(pair => [pair.empresa, pair.consultoria]),
    commonKeys: pairs.map(pair => pair.compareDisplay),
    onlyEmpresa,
    onlyConsultoria
  };
}

function consolidateCompareKeyRows(rows, compareDisplay, origem) {
  const representative = pickRepresentativeRow(rows);
  const responsaveisTodos = mergeResponsibles(rows);
  const historicoCategorias = rows
    .flatMap(row => Array.isArray(row.historicoCategorias) ? row.historicoCategorias : [row.categoria])
    .filter((value, index, array) => value && array.indexOf(value) === index)
    .sort((a, b) => extractStageOrder(a) - extractStageOrder(b) || String(a).localeCompare(String(b), 'pt-BR'));

  return {
    ...representative,
    origem,
    compareDisplay: compareDisplay || representative.compareDisplay,
    responsaveisTodos,
    responsavel: responsaveisTodos[0] || representative.responsavel || 'Sem responsável',
    atrasadoFlag: rows.some(row => row.atrasadoFlag === true),
    rawRowsCount: rows.reduce((total, row) => total + Number(row.rawRowsCount || 1), 0),
    rawCardsCount: rows.length,
    historicoCategorias,
    phaseSummary: buildPhaseSummary(rows)
  };
}

function buildPairStatus(empresa, consultoria) {
  if (isCritical(empresa) || isCritical(consultoria)) return 'Crítico';
  if (isConcluded(empresa.statusOriginal) && isConcluded(consultoria.statusOriginal)) return 'Concluído';
  if (isNotStarted(empresa.statusOriginal) && isNotStarted(consultoria.statusOriginal)) return 'Não iniciado';
  return 'Em andamento';
}

function buildPairProgressPercent(empresa, consultoria) {
  const doneEmpresa = isConcluded(empresa.statusOriginal) ? 1 : 0;
  const doneConsultoria = isConcluded(consultoria.statusOriginal) ? 1 : 0;
  return Math.round(((doneEmpresa + doneConsultoria) / 2) * 100);
}

function buildPairSummary(pairs) {
  const summary = { total: pairs.length, concluded: 0, inProgress: 0, notStarted: 0, critical: 0 };

  pairs.forEach(pair => {
    if (pair.pairStatus === 'Concluído') summary.concluded += 1;
    else if (pair.pairStatus === 'Não iniciado') summary.notStarted += 1;
    else if (pair.pairStatus === 'Crítico') summary.critical += 1;
    else summary.inProgress += 1;
  });

  summary.percent = getPercent(summary.concluded, summary.total);
  return summary;
}

function buildPhaseSummary(rows) {
  const summary = { concluded: 0, inProgress: 0, notStarted: 0 };

  rows.forEach(row => {
    if (isConcluded(row.statusOriginal)) summary.concluded += 1;
    else if (isNotStarted(row.statusOriginal)) summary.notStarted += 1;
    else if (isInProgress(row.statusOriginal)) summary.inProgress += 1;
  });

  return summary;
}

function buildCardGroupKey({ origem, riceCode, cardSignature, tarefa }) {
  const signature = cardSignature || buildTaskSignature(tarefa, riceCode);
  return [origem || 'Origem', riceCode || '', signature || normalize(tarefa)].join('|||');
}

function buildTaskSignature(taskName, riceCode) {
  let signature = normalize(taskName);
  const normalizedRice = normalize(riceCode);

  if (normalizedRice && signature.startsWith(normalizedRice)) {
    signature = signature.slice(normalizedRice.length);
  }

  signature = signature
    .replace(/[\[\]()]/g, ' ')
    .replace(/[|/_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return signature || normalizedRice || normalize(taskName);
}

function extractStageOrder(categoria) {
  const match = String(categoria || '').match(/^\s*(\d+)/);
  return match ? Number(match[1]) : 999;
}

function buildCompareKey({ riceCode, categoria }) {
  const mode = DASH_CONFIG.compareMode;
  if (mode === 'category') return normalize(categoria);
  if (mode === 'rice') return riceCode;
  if (mode === 'rice_category') return riceCode && categoria ? `${riceCode}|||${normalize(categoria)}` : '';
  return riceCode;
}

function buildCompareDisplay({ riceCode, categoria }) {
  const mode = DASH_CONFIG.compareMode;
  if (mode === 'category') return categoria;
  if (mode === 'rice') return riceCode;
  if (mode === 'rice_category') return `${riceCode} · ${categoria}`;
  return riceCode;
}

function getModeConfig() {
  return DASH_CONFIG.compareModeOptions[DASH_CONFIG.compareMode] || DASH_CONFIG.compareModeOptions.rice;
}

function renderDashboard(model) {
  const hasComparable = model.pairs.length > 0;
  const hasDivergences = model.onlyEmpresa.length > 0 || model.onlyConsultoria.length > 0;

  if (!hasComparable && !hasDivergences) {
    renderEmptyStates('As planilhas foram lidas, mas não há itens válidos para comparar nessa configuração.');
    return;
  }

  renderCommonKeys(model.commonKeys);

  const summary = buildPairSummary(model.pairs);

  updateSummary({
    totalComparable: summary.total,
    totalKeys: model.commonKeys.length,
    concluded: summary.concluded,
    inProgress: summary.inProgress,
    notStarted: summary.notStarted,
    critical: summary.critical,
    percent: summary.percent,
    onlyEmpresa: model.onlyEmpresa.length,
    onlyConsultoria: model.onlyConsultoria.length
  });

  renderPlannerMetrics(model.comparableRows.filter(row => row.origem === 'Empresa'), 'Empresa', els.empresaMetrics);
  renderPlannerMetrics(model.comparableRows.filter(row => row.origem === 'Consultoria'), 'Consultoria', els.consultoriaMetrics);
  renderLeaderboard(model.comparableRows);
  renderStatusBars(summary.total, summary.concluded, summary.inProgress, summary.notStarted, summary.critical);
  renderComparisonBoard(model.comparableRows);
  setupRiceBoardFilter(model.pairs);
  renderFocusTable(model.pairs);
  renderDivergenceBoard(model.onlyEmpresa, model.onlyConsultoria);
  renderDivergencePanels(model.onlyEmpresa, model.onlyConsultoria);
  renderDivergenceTable(model.onlyEmpresa, model.onlyConsultoria);
}

function renderCommonKeys(commonKeys) {
  if (!els.commonKeys) return;
  els.commonKeys.innerHTML = commonKeys.length
    ? commonKeys.map(key => `<span class="category-chip">${escapeHtml(key)}</span>`).join('')
    : 'Nenhum RICE casado 1x1 encontrado.';
}

function updateSummary(summary) {
  const mode = getModeConfig();

  if (els.totalItens) els.totalItens.textContent = formatNumber(summary.totalComparable);
  if (els.totalBaseKeys) els.totalBaseKeys.textContent = formatNumber(summary.totalKeys);
  if (els.totalConcluidos) els.totalConcluidos.textContent = formatNumber(summary.concluded);
  if (els.totalAndamento) els.totalAndamento.textContent = formatNumber(summary.inProgress);
  if (els.totalNaoIniciado) els.totalNaoIniciado.textContent = formatNumber(summary.notStarted);
  if (els.totalCriticos) els.totalCriticos.textContent = formatNumber(summary.critical);
  if (els.totalSoEmpresa) els.totalSoEmpresa.textContent = formatNumber(summary.onlyEmpresa);
  if (els.totalSoConsultoria) els.totalSoConsultoria.textContent = formatNumber(summary.onlyConsultoria);
  if (els.totalRicesEmpresa) els.totalRicesEmpresa.textContent = formatNumber(summary.totalComparable + summary.onlyEmpresa);
  if (els.totalRicesConsultoria) els.totalRicesConsultoria.textContent = formatNumber(summary.totalComparable + summary.onlyConsultoria);
  if (els.globalPercent) els.globalPercent.textContent = `${summary.percent}%`;

  if (els.ringProgress) {
    els.ringProgress.style.strokeDashoffset = `${RING_CIRCUMFERENCE * (1 - summary.percent / 100)}`;
  }

  const message = [...EXECUTIVE_MESSAGES].reverse().find(item => summary.percent >= item.threshold) || EXECUTIVE_MESSAGES[0];

  if (els.statusExecutivo) els.statusExecutivo.textContent = message.status;
  if (els.headlineCallout) {
    els.headlineCallout.textContent = summary.totalComparable
      ? `${message.title} Existem ${formatNumber(summary.onlyEmpresa)} RICE(s) só na Empresa e ${formatNumber(summary.onlyConsultoria)} só na Consultoria para saneamento.`
      : `Não há pares 1x1 para comparar. Existem ${formatNumber(summary.onlyEmpresa)} RICE(s) só na Empresa e ${formatNumber(summary.onlyConsultoria)} só na Consultoria.`;
  }
  if (els.headlinePill) els.headlinePill.textContent = mode.label;
}

function renderPlannerMetrics(rows, origem, targetEl) {
  if (!targetEl) return;

  const filtered = rows || [];
  if (!filtered.length) {
    targetEl.innerHTML = '<div class="empty-state">Sem registros dessa origem na base comparável 1x1.</div>';
    return;
  }

  const total = filtered.length;
  const concluded = filtered.filter(row => isConcluded(row.statusOriginal)).length;
  const inProgress = filtered.filter(row => isInProgress(row.statusOriginal)).length;
  const notStarted = filtered.filter(row => isNotStarted(row.statusOriginal)).length;
  const critical = filtered.filter(row => isCritical(row)).length;
  const uniqueKeys = new Set(filtered.map(row => row.compareDisplay)).size;
  const percent = getPercent(concluded, total);

  targetEl.innerHTML = `
    <div class="planner-header">
      <div class="planner-title">
        <strong>${escapeHtml(origem)}</strong>
        <span>${formatNumber(total)} card(s) comparáveis · ${formatNumber(uniqueKeys)} RICE(s) casados</span>
      </div>
      <div class="planner-percent">${percent}%</div>
    </div>
    <div class="planner-status-grid">
      <div class="planner-kpi"><small>Concluídos</small><strong>${formatNumber(concluded)}</strong></div>
      <div class="planner-kpi"><small>Em andamento</small><strong>${formatNumber(inProgress)}</strong></div>
      <div class="planner-kpi"><small>Não iniciados</small><strong>${formatNumber(notStarted)}</strong></div>
      <div class="planner-kpi"><small>Críticos</small><strong>${formatNumber(critical)}</strong></div>
      <div class="planner-kpi"><small>RICEs</small><strong>${formatNumber(uniqueKeys)}</strong></div>
    </div>
  `;
}

function renderLeaderboard(rows) {
  if (!els.leaderboard) return;

  const grouped = new Map();

  rows.forEach(row => {
    const people = row.responsaveisTodos.length ? row.responsaveisTodos : [row.responsavel];
    people.forEach(person => {
      const key = `${person}|||${row.origem}`;
      if (!grouped.has(key)) {
        grouped.set(key, { responsavel: person, origem: row.origem, total: 0, concluded: 0, inProgress: 0 });
      }
      const item = grouped.get(key);
      item.total += 1;
      if (isConcluded(row.statusOriginal)) item.concluded += 1;
      if (isInProgress(row.statusOriginal)) item.inProgress += 1;
    });
  });

  const ranking = [...grouped.values()]
    .filter(item => item.responsavel && item.responsavel !== 'Sem responsável')
    .map(item => ({ ...item, percent: getPercent(item.concluded, item.total) }))
    .sort((a, b) =>
      b.percent - a.percent ||
      b.concluded - a.concluded ||
      a.responsavel.localeCompare(b.responsavel, 'pt-BR')
    )
    .slice(0, 10);

  els.leaderboard.innerHTML = ranking.length
    ? ranking.map((item, index) => `
      <div class="leader-row">
        <div class="place ${index === 0 ? 'top-1' : index === 1 ? 'top-2' : index === 2 ? 'top-3' : ''}">${index + 1}</div>
        <div>
          <div class="leader-name">${escapeHtml(item.responsavel)}</div>
          <div class="leader-meta">${escapeHtml(item.origem)} · ${formatNumber(item.concluded)} concluídos de ${formatNumber(item.total)} · ${formatNumber(item.inProgress)} em andamento</div>
        </div>
        <div class="leader-score"><strong>${item.percent}%</strong><span>aproveitamento</span></div>
      </div>`).join('')
    : 'Sem responsáveis válidos para exibir.';
}

function renderStatusBars(total, concluded, inProgress, notStarted, critical) {
  if (!els.statusBars) return;

  if (!total) {
    els.statusBars.innerHTML = 'Sem pares 1x1 para distribuição.';
    return;
  }

  const other = Math.max(total - concluded - inProgress - notStarted - critical, 0);
  const statuses = [
    { label: 'Concluído', value: concluded, percent: getPercent(concluded, total), color: 'linear-gradient(90deg, #14d3a6, #7dffd8)' },
    { label: 'Em andamento', value: inProgress, percent: getPercent(inProgress, total), color: 'linear-gradient(90deg, #ffb84d, #ffd88d)' },
    { label: 'Não iniciado', value: notStarted, percent: getPercent(notStarted, total), color: 'linear-gradient(90deg, #7c5cff, #b7a6ff)' },
    { label: 'Crítico', value: critical, percent: getPercent(critical, total), color: 'linear-gradient(90deg, #ff4d6d, #ff8fa3)' },
    { label: 'Outros', value: other, percent: getPercent(other, total), color: 'linear-gradient(90deg, #98a7d8, #cad5ff)' }
  ];

  els.statusBars.innerHTML = statuses.map(item => `
    <div class="status-item">
      <div class="status-head"><strong>${item.label}</strong><span>${formatNumber(item.value)} · ${item.percent}%</span></div>
      <div class="status-track"><div class="status-fill" style="width:${item.percent}%; background:${item.color}"></div></div>
    </div>`).join('');
}

function renderComparisonBoard(rows) {
  if (!els.comparisonBoard) return;

  if (!rows.length) {
    els.comparisonBoard.innerHTML = 'Sem pares 1x1 para comparar.';
    return;
  }

  const empresa = rows.filter(r => r.origem === 'Empresa');
  const consultoria = rows.filter(r => r.origem === 'Consultoria');

  const statuses = [
    { label: 'Concluído', fn: row => isConcluded(row.statusOriginal) },
    { label: 'Em andamento', fn: row => isInProgress(row.statusOriginal) },
    { label: 'Não iniciado', fn: row => isNotStarted(row.statusOriginal) },
    { label: 'Bloqueado', fn: row => isBlocked(row) },
    { label: 'Atrasado', fn: row => isDelayed(row) }
  ];

  els.comparisonBoard.innerHTML = statuses.map(item => {
    const empresaValue = empresa.filter(item.fn).length;
    const consultoriaValue = consultoria.filter(item.fn).length;
    const empresaPercent = getPercent(empresaValue, empresa.length);
    const consultoriaPercent = getPercent(consultoriaValue, consultoria.length);

    return `<div class="comparison-row">
      <div class="comparison-label">${item.label}</div>
      <div class="comparison-side comparison-side-inline">
        <span>Empresa</span>
        <div class="comparison-bar-wrap">
          <div class="status-track"><div class="status-fill" style="width:${empresaPercent}%; background: linear-gradient(90deg, #7c5cff, #b7a6ff)"></div></div>
          <strong class="comparison-inline-total">${formatNumber(empresaValue)}</strong>
        </div>
      </div>
      <div class="comparison-side comparison-side-inline">
        <span>Consultoria</span>
        <div class="comparison-bar-wrap">
          <div class="status-track"><div class="status-fill" style="width:${consultoriaPercent}%; background: linear-gradient(90deg, #14d3a6, #7dffd8)"></div></div>
          <strong class="comparison-inline-total">${formatNumber(consultoriaValue)}</strong>
        </div>
      </div>
    </div>`;
  }).join('');
}

function setupRiceBoardFilter(pairs) {
  dashboardState.ricePairsAll = Array.isArray(pairs) ? [...pairs] : [];
  populateRiceBoardFilter(dashboardState.ricePairsAll);
  applyRiceBoardFilter();
}

function populateRiceBoardFilter(pairs) {
  if (!els.riceFilterSelect) return;

  const options = [...pairs]
    .sort((a, b) => String(a.compareDisplay || '').localeCompare(String(b.compareDisplay || ''), 'pt-BR'))
    .map(pair => `<option value="${escapeHtml(pair.compareKey)}" selected>${escapeHtml(pair.compareDisplay)}</option>`);

  els.riceFilterSelect.innerHTML = options.join('');

  if (!pairs.length && els.riceFilterStatus) {
    els.riceFilterStatus.textContent = 'Sem RICEs comparáveis para selecionar.';
  }
}

function setRiceFilterSelection(isSelected) {
  if (!els.riceFilterSelect) return;
  Array.from(els.riceFilterSelect.options).forEach(option => {
    option.selected = isSelected;
  });
}

function getSelectedRiceFilterValues() {
  if (!els.riceFilterSelect) return [];
  return Array.from(els.riceFilterSelect.selectedOptions).map(option => option.value);
}

function applyRiceBoardFilter() {
  const allPairs = Array.isArray(dashboardState.ricePairsAll) ? dashboardState.ricePairsAll : [];

  if (!allPairs.length) {
    renderKeyBoard([]);
    if (els.riceFilterStatus) els.riceFilterStatus.textContent = 'Sem RICEs comparáveis para exibir.';
    return;
  }

  const selectedKeys = getSelectedRiceFilterValues();

  if (!selectedKeys.length) {
    if (els.keyBoard) {
      els.keyBoard.innerHTML = '<div class="empty-state">Nenhum RICE selecionado no filtro.</div>';
    }
    if (els.riceFilterStatus) els.riceFilterStatus.textContent = `0 de ${formatNumber(allPairs.length)} RICE(s) exibidos.`;
    return;
  }

  const selectedSet = new Set(selectedKeys);
  const filteredPairs = allPairs.filter(pair => selectedSet.has(pair.compareKey));
  renderKeyBoard(filteredPairs);
  if (els.riceFilterStatus) {
    els.riceFilterStatus.textContent = `${formatNumber(filteredPairs.length)} de ${formatNumber(allPairs.length)} RICE(s) exibidos.`;
  }
}

function renderKeyBoard(pairs) {
  if (!els.keyBoard) return;

  if (!pairs.length) {
    els.keyBoard.innerHTML = '<div class="empty-state">Nenhum RICE casado 1x1 para exibir.</div>';
    return;
  }

  const statusWeight = { 'Crítico': 4, 'Em andamento': 3, 'Não iniciado': 2, 'Concluído': 1 };

  const cards = [...pairs]
    .sort((a, b) =>
      (statusWeight[b.pairStatus] || 0) - (statusWeight[a.pairStatus] || 0) ||
      b.progressPercent - a.progressPercent ||
      a.compareDisplay.localeCompare(b.compareDisplay, 'pt-BR')
    );

  els.keyBoard.innerHTML = cards.map(pair => `
    <div class="area-card">
      <div class="area-top">
        <div class="area-title">${escapeHtml(pair.compareDisplay)}</div>
        <div class="area-badge">${pair.progressPercent}%</div>
      </div>
      <div class="status-track" style="margin-top:12px;"><div class="status-fill" style="width:${pair.progressPercent}%; background: linear-gradient(90deg, #7c5cff, #14d3a6);"></div></div>
      <div class="area-stats"><span>Empresa: ${escapeHtml(pair.empresa.statusOriginal)}</span><span>Consultoria: ${escapeHtml(pair.consultoria.statusOriginal)}</span></div>
      <div class="area-stats"><span>Status do par</span><span>${escapeHtml(pair.pairStatus)}</span></div>
    </div>`).join('');
}


function renderFocusTable(pairs) {
  focusTableState.rows = pairs
    .filter(pair =>
      pair.pairStatus === 'Crítico' ||
      pair.pairStatus === 'Não iniciado' ||
      isNotStarted(pair.empresa.statusOriginal) ||
      isNotStarted(pair.consultoria.statusOriginal)
    )
    .flatMap(pair => [
      { ...pair.empresa, compareDisplay: pair.compareDisplay, pairStatus: pair.pairStatus },
      { ...pair.consultoria, compareDisplay: pair.compareDisplay, pairStatus: pair.pairStatus }
    ])
    .sort((a, b) =>
      compareCriticality(a, b) ||
      comparePriority(a.prioridade, b.prioridade) ||
      a.compareDisplay.localeCompare(b.compareDisplay, 'pt-BR')
    )
    .slice(0, 20);

  populateFocusFilterOptions();
  renderFilteredFocusTable();
}

function renderDivergenceBoard(onlyEmpresa, onlyConsultoria) {
  if (!els.divergenceBoard) return;

  const groups = [
    {
      title: 'Só na Empresa',
      count: onlyEmpresa.length,
      subtitle: 'RICE(s) encontrados internamente e ausentes na Consultoria.',
      keys: onlyEmpresa.map(row => row.compareDisplay)
    },
    {
      title: 'Só na Consultoria',
      count: onlyConsultoria.length,
      subtitle: 'RICE(s) encontrados na Consultoria e ausentes na Empresa.',
      keys: onlyConsultoria.map(row => row.compareDisplay)
    }
  ];

  els.divergenceBoard.innerHTML = groups.map(group => `
    <div class="area-card">
      <div class="area-top">
        <div class="area-title">${escapeHtml(group.title)}</div>
        <div class="area-badge">${formatNumber(group.count)}</div>
      </div>
      <div class="area-stats" style="display:block; margin-top:12px;">
        <span>${escapeHtml(group.subtitle)}</span>
      </div>
      <div class="chip-board" style="margin-top:12px;">
        ${group.keys.length
          ? group.keys.slice(0, 12).map(key => `<span class="category-chip">${escapeHtml(key)}</span>`).join('')
          : '<span class="category-chip">Nenhum</span>'}
      </div>
    </div>`).join('');
}

function renderDivergencePanels(onlyEmpresa, onlyConsultoria) {
  renderSingleDivergencePanel({
    targetEl: els.divergenceEmpresaList,
    emptyMessage: 'Nenhuma RICE exclusiva da Empresa.',
    rows: onlyEmpresa
  });

  renderSingleDivergencePanel({
    targetEl: els.divergenceConsultoriaList,
    emptyMessage: 'Nenhuma RICE exclusiva da Consultoria.',
    rows: onlyConsultoria
  });
}

function renderSingleDivergencePanel({ targetEl, rows, emptyMessage }) {
  if (!targetEl) return;

  targetEl.classList.add('divergence-compact-list');

  if (!rows.length) {
    targetEl.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
    return;
  }

  targetEl.innerHTML = rows
    .slice()
    .sort((a, b) => String(a.compareDisplay || '').localeCompare(String(b.compareDisplay || ''), 'pt-BR'))
    .map(row => `
      <div class="divergence-compact-card">
        <div class="divergence-main">
          <div class="divergence-code">${escapeHtml(row.compareDisplay || '-')}</div>
          <div class="divergence-desc">${escapeHtml(formatRiceDescription(row))}</div>
          <div class="divergence-owner"><strong>Atribuído:</strong> ${escapeHtml(row.responsavel || 'Sem responsável')}</div>
        </div>
        <div class="divergence-phases">
          <div class="divergence-phase-row"><span>Concluído</span><strong>${formatNumber(row.phaseSummary?.concluded || 0)}</strong></div>
          <div class="divergence-phase-row"><span>Não iniciado</span><strong>${formatNumber(row.phaseSummary?.notStarted || 0)}</strong></div>
          <div class="divergence-phase-row"><span>Em andamento</span><strong>${formatNumber(row.phaseSummary?.inProgress || 0)}</strong></div>
        </div>
      </div>`)
    .join('');
}

function renderDivergenceTable(onlyEmpresa, onlyConsultoria) {
  if (!els.divergenceTable) return;

  const rows = [
    ...onlyEmpresa.map(row => ({ ...row, divergenceSide: 'Só na Empresa' })),
    ...onlyConsultoria.map(row => ({ ...row, divergenceSide: 'Só na Consultoria' }))
  ].sort((a, b) =>
    a.divergenceSide.localeCompare(b.divergenceSide, 'pt-BR') ||
    a.compareDisplay.localeCompare(b.compareDisplay, 'pt-BR')
  );

  els.divergenceTable.innerHTML = rows.length
    ? rows.map(row => `
      <tr>
        <td>${originBadge(row.origem)}</td>
        <td>${escapeHtml(row.compareDisplay || '-')}</td>
        <td>${escapeHtml(row.id || '-')}</td>
        <td>${escapeHtml(formatTaskForDisplay(row))}</td>
        <td>${escapeHtml(row.responsavel || '-')}</td>
        <td>${statusPill(row)}</td>
        <td>${escapeHtml(getDivergenceReason(row))}</td>
      </tr>`).join('')
    : '<tr><td colspan="7" class="empty-cell">Nenhuma divergência encontrada entre as bases.</td></tr>';
}

function getDivergenceReason(row) {
  if (row.origem === 'Empresa') return 'Existe na Empresa e não apareceu na Consultoria.';
  if (row.origem === 'Consultoria') return 'Existe na Consultoria e não apareceu na Empresa.';
  return 'Sem par na outra base.';
}

function renderEmptyStates(message) {
  resetSummary();
  if (els.statusExecutivo) els.statusExecutivo.textContent = 'Falha de carga';
  if (els.headlineCallout) els.headlineCallout.textContent = message;
  if (els.headlinePill) els.headlinePill.textContent = 'Sem leitura';
  if (els.commonKeys) els.commonKeys.innerHTML = message;
  if (els.leaderboard) els.leaderboard.innerHTML = message;
  if (els.statusBars) els.statusBars.innerHTML = message;
  if (els.comparisonBoard) els.comparisonBoard.innerHTML = message;
  if (els.keyBoard) els.keyBoard.innerHTML = message;
  dashboardState.ricePairsAll = [];
  if (els.riceFilterSelect) els.riceFilterSelect.innerHTML = '';
  if (els.riceFilterStatus) els.riceFilterStatus.textContent = message;
  if (els.divergenceBoard) els.divergenceBoard.innerHTML = `<div class="empty-state">${message}</div>`;
  if (els.divergenceEmpresaList) els.divergenceEmpresaList.innerHTML = `<div class="empty-state">${message}</div>`;
  if (els.divergenceConsultoriaList) els.divergenceConsultoriaList.innerHTML = `<div class="empty-state">${message}</div>`;
  if (els.empresaMetrics) els.empresaMetrics.innerHTML = `<div class="empty-state">${message}</div>`;
  if (els.consultoriaMetrics) els.consultoriaMetrics.innerHTML = `<div class="empty-state">${message}</div>`;
  focusTableState.rows = [];
  if (els.focusTable) els.focusTable.innerHTML = `<tr><td colspan="6" class="empty-cell">${message}</td></tr>`;
  populateFocusFilterOptions();
  if (els.divergenceTable) els.divergenceTable.innerHTML = `<tr><td colspan="7" class="empty-cell">${message}</td></tr>`;
}

function resetSummary() {
  if (els.totalItens) els.totalItens.textContent = '0';
  if (els.totalBaseKeys) els.totalBaseKeys.textContent = '0';
  if (els.totalConcluidos) els.totalConcluidos.textContent = '0';
  if (els.totalAndamento) els.totalAndamento.textContent = '0';
  if (els.totalNaoIniciado) els.totalNaoIniciado.textContent = '0';
  if (els.totalCriticos) els.totalCriticos.textContent = '0';
  if (els.totalSoEmpresa) els.totalSoEmpresa.textContent = '0';
  if (els.totalSoConsultoria) els.totalSoConsultoria.textContent = '0';
  if (els.totalRicesEmpresa) els.totalRicesEmpresa.textContent = '0';
  if (els.totalRicesConsultoria) els.totalRicesConsultoria.textContent = '0';
  if (els.globalPercent) els.globalPercent.textContent = '0%';
  if (els.ringProgress) els.ringProgress.style.strokeDashoffset = `${RING_CIRCUMFERENCE}`;
}

function handleFatal(message) {
  console.error(message);
  showManualLoader(message);
  setGeneratedAt(message);
  renderEmptyStates(message);
}

function showManualLoader(message) {
  if (els.manualLoaderCard) els.manualLoaderCard.hidden = false;
  if (els.manualLoaderMessage) els.manualLoaderMessage.textContent = message;
}

function hideManualLoader() {
  if (els.manualLoaderCard) els.manualLoaderCard.hidden = true;
}

function setGeneratedAt(message) {
  if (els.dataAtualizacao) els.dataAtualizacao.textContent = message;
}

function getValue(row, key) {
  if (!row) return '';
  const keys = Array.isArray(key) ? key : [key];

  for (const candidate of keys) {
    if (candidate in row) return row[candidate] ?? '';
  }

  const normalizedEntries = Object.entries(row).map(([entryKey, value]) => [normalizeFieldName(entryKey), value]);

  for (const candidate of keys) {
    const normalizedCandidate = normalizeFieldName(candidate);
    const match = normalizedEntries.find(([entryKey]) => entryKey === normalizedCandidate);
    if (match) return match[1] ?? '';
  }

  return '';
}

function normalizeFieldName(value = '') {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function splitPeople(value) {
  return String(value || '').split(';').map(item => item.trim()).filter(Boolean);
}

function firstResponsible(value) {
  const people = splitPeople(value);
  return people[0] || 'Sem responsável';
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

function isBlocked(row) {
  return containsAny(row.statusOriginal, DASH_CONFIG.statusRules.blockedContains);
}

function isDelayed(row) {
  return row.atrasadoFlag === true;
}

function isDelayedValue(value) {
  const normalized = normalize(value);
  return ['true', 'verdadeiro', 'sim', '1', 'yes'].includes(normalized);
}

function compareCriticality(a, b) {
  const weight = row => isCritical(row) ? 2 : isNotStarted(row.statusOriginal) ? 1 : 0;
  return weight(b) - weight(a);
}

function comparePriority(a, b) {
  const order = DASH_CONFIG.statusRules.priorityOrder || {};
  return (order[normalize(b)] || 0) - (order[normalize(a)] || 0);
}

function extractRiceCode(taskName) {
  const raw = String(taskName || '').trim();
  const match = raw.match(/^([A-Za-z0-9]+(?:[.#-][A-Za-z0-9]+)*(?:\.\d+)*)/);
  return match ? match[1].toUpperCase() : '';
}

function getPercent(value, total) {
  return total ? Math.round((value / total) * 100) : 0;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('pt-BR');
}

function formatRiceDescription(row) {
  const baseTask = String(row?.tarefa || '').trim();
  const riceCode = String(row?.riceCode || row?.compareDisplay || '').trim();

  if (!baseTask) return 'Sem descrição';
  if (!riceCode) return baseTask;

  const pattern = new RegExp(`^\\s*${escapeRegExp(riceCode)}(?:\\s*[-–—:|.#]+)?\\s*`, 'i');
  const description = baseTask.replace(pattern, '').trim();
  return description || 'Sem descrição complementar';
}

function formatTaskForDisplay(row) {
  const baseTask = row?.tarefa || '-';
  const rawCardsCount = Number(row?.rawCardsCount || 0);
  const rawRowsCount = Number(row?.rawRowsCount || 0);

  if (rawCardsCount > 1) return `${baseTask} · consolidado (${rawCardsCount} cards / ${rawRowsCount} réguas)`;
  if (rawRowsCount > 1) return `${baseTask} · consolidado (${rawRowsCount} réguas)`;
  return baseTask;
}

function formatPhaseForDisplay(row) {
  return String(row?.categoria || 'Sem fase').trim() || 'Sem fase';
}

function originBadge(origin) {
  const className = origin === 'Consultoria' ? 'origin-consultoria' : 'origin-empresa';
  return `<span class="origin-badge ${className}">${escapeHtml(origin)}</span>`;
}

function statusPill(row) {
  const label = escapeHtml(row.statusOriginal || '-');
  let className = 'status-outro';

  if (row.atrasadoFlag === true) className = 'status-atrasado';
  else if (containsAny(row.statusOriginal, DASH_CONFIG.statusRules.blockedContains)) className = 'status-bloqueado';
  else if (isConcluded(row.statusOriginal)) className = 'status-concluido';
  else if (isInProgress(row.statusOriginal)) className = 'status-andamento';
  else if (isNotStarted(row.statusOriginal)) className = 'status-nao-iniciado';

  return `<span class="status-pill ${className}">${label}</span>`;
}

function escapeRegExp(value) {
  return String(value ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
