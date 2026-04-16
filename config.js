
const DASH_CONFIG = {
  auth: { username: 'conecta', password: 'pluma@2026' },

  files: [
    {
      fileName: 'Programa Conecta - Acompanhamento interno.xlsx',
      origem: 'Empresa',
      label: 'Programa Conecta - Acompanhamento interno'
    },
    {
      fileName: "AN.020 - Acompanhamento RICE's.xlsx",
      origem: 'Consultoria',
      label: "AN.020 - Acompanhamento RICE's"
    }
  ],

  sheetName: 'Dados Consolidados',

  fields: {
    taskName: 'Nome da tarefa',
    taskId: ['Identificação da tarefa', 'Identificacao da tarefa'],
    status: 'Status',
    responsible: ['Atribuído a', 'Atribuido a'],
    category: 'Categoria',
    priority: 'Prioridade',
    delayed: 'Atrasados'
  },

  // Modos aceitos:
  // 'category'      -> compara categorias iguais
  // 'rice'          -> compara códigos de RICE iguais
  // 'rice_category' -> compara combinação RICE + categoria
  compareMode: 'rice',

  compareModeOptions: {
    category: {
      label: 'Categoria igual',
      keyLabel: 'Categorias iguais',
      heroDescription: 'Leitura executiva consolidada comparando apenas categorias presentes nas duas planilhas.',
      modeDescription: 'Modo atual: base comparável por categoria presente nas duas fontes.',
      commonKeysTitle: 'Categorias usadas na comparação',
      keyBoardTitle: 'Desempenho por categoria comparável'
    },
    rice: {
      label: 'RICE igual',
      keyLabel: 'RICEs iguais',
      heroDescription: 'Leitura executiva consolidada comparando apenas códigos de RICE presentes nas duas planilhas.',
      modeDescription: 'Modo atual: base comparável por código de RICE no início do nome da tarefa.',
      commonKeysTitle: 'RICEs usados na comparação',
      keyBoardTitle: 'Desempenho por código de RICE'
    },
    rice_category: {
      label: 'RICE + Categoria',
      keyLabel: 'RICE+Categoria',
      heroDescription: 'Leitura executiva consolidada comparando apenas combinações de RICE + categoria presentes nas duas planilhas.',
      modeDescription: 'Modo atual: base comparável por combinação de código de RICE e categoria.',
      commonKeysTitle: 'Combinações usadas na comparação',
      keyBoardTitle: 'Desempenho por RICE + categoria'
    }
  },

  statusRules: {
    concludedContains: ['concluida'],
    inProgressContains: ['andamento'],
    notStartedContains: ['nao iniciado'],
    blockedContains: ['bloque'],
    priorityOrder: {
      urgente: 4,
      alta: 3,
      media: 2,
      baixa: 1
    }
  }
};
