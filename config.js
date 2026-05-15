const DASH_CONFIG = {
  auth: { username: 'conecta', password: 'pluma@2026' },
  dataFile: 'dados_consilidados_rices.csv',

  files: [
    {
      fileName: 'Programa Conecta - Acompanhamento interno.xlsx',
      origem: 'Empresa',
      label: 'Programa Conecta - Acompanhamento interno'
    },
    {
      fileName: 'AN.020 - Acompanhamento RICEs.xlsx',
      origem: 'Consultoria',
      label: 'AN.020 - Acompanhamento RICEs'
    }
  ],

  sheetName: 'Dados Consolidados',

  fields: {
    rice: 'RICE',
    taskName: 'Nome da tarefa',
    taskId: 'Identificação da tarefa',
    status: 'Status',
    responsible: 'Atribuído a',
    category: 'Categoria',
    priority: 'Prioridade',
    delayed: 'Atrasados'
  },

  compareMode: 'rice',

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
