/**
 * Script para renomeação automática de pastas no Google Drive
 * baseado nos dados da planilha FUP RENAULT - PADRÃO
 * VERSÃO ENXUGADA - Apenas funcionalidades essenciais
 */

// Configurações
const SPREADSHEET_ID = '1C9ZkOqXBh5IdjXP8K8CuEsmmQgy5C1yZxXHKJFnHj24';
const SHEET_NAME = 'ANDAM.';

// IDs das pastas onde procurar os processos
const PASTA_MPR_2025_ID = '1NjYLk2dt8PB3RpaJ0C9O8K_Gs_8OIHoZ';
const PASTA_PHF_DIV_OHS_ID = '1AXkNs3PHYfhMJ78ZAf656dbIpmY6TQJw';

// Colunas da planilha
const COL_REF_RENAULT = 1; // Coluna A
const COL_REF_GNH = 2;     // Coluna B

// Meses para busca no MPR
const MESES_MPR = [
  '01 - JANEIRO', '02 - FEVEREIRO', '03 - MARÇO', '04 - ABRIL',
  '05 - MAIO', '06 - JUNHO', '07 - JULHO', '08 - AGOSTO',
  '09 - SETEMBRO', '10 - OUTUBRO', '11 - NOVEMBRO', '12 - DEZEMBRO'
];

/**
 * Função principal - executa a renomeação (com controle de horário)
 */
function renomearPastas() {
  try {
    // Verificar se está no horário de trabalho (8h às 18h)
    const agora = new Date();
    const hora = agora.getHours();
    
    if (hora < 8 || hora >= 18) {
      console.log(`⏰ Fora do horário de trabalho (${hora}h). Script pausado até 8h.`);
      return;
    }
    
    console.log('=== INICIANDO RENOMEAÇÃO ===');
    console.log(`⏰ ${agora.toLocaleString()} (${hora}h - Horário válido)`);
    
    const dados = obterDadosPlanilha();
    if (!dados || dados.length === 0) {
      console.log('❌ Nenhum dado encontrado na planilha.');
      return;
    }
    
    console.log(`📊 ${dados.length} linhas encontradas`);
    
    let processados = 0, renomeados = 0, erros = 0, ignoradosX = 0;
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const refRenault = linha[COL_REF_RENAULT - 1];
      const refGnh = linha[COL_REF_GNH - 1];
      
      // Pular cabeçalho
      if (refRenault === 'REF. RENAULT' || refGnh === 'REF. GNH') continue;
      
      // Verificar dados válidos
      if (!refRenault || !refGnh || 
          !refRenault.toString().trim() || 
          !refGnh.toString().trim()) {
        continue;
      }

      // Contar processos com X
      if (refRenault.toString().toUpperCase().trim() === 'X') {
        ignoradosX++;
        continue;
      }

      console.log(`\n--- Linha ${i + 1}: ${refRenault} | ${refGnh} ---`);

      try {
        if (processarPasta(refRenault.toString().trim(), refGnh.toString().trim())) {
          renomeados++;
          console.log('✅ Renomeada');
        }
        processados++;
      } catch (error) {
        erros++;
        console.error(`❌ Erro:`, error.message);
      }
    }

    console.log(`\n=== RESUMO ===`);
    console.log(`✅ Processadas: ${processados} | Renomeadas: ${renomeados}`);
    console.log(`⏸️ Pendentes (X): ${ignoradosX} | ❌ Erros: ${erros}`);
    
  } catch (error) {
    console.error('❌ Erro geral:', error.message);
  }
}

/**
 * Obter dados da planilha
 */
function obterDadosPlanilha() {
  try {
    const planilha = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = planilha.getSheetByName(SHEET_NAME);
    
    if (!aba) {
      console.error(`❌ Aba "${SHEET_NAME}" não encontrada`);
      return [];
    }
    
    if (aba.getLastRow() < 2) {
      console.log('❌ Planilha vazia');
      return [];
    }
    
    return aba.getDataRange().getValues();
    
  } catch (error) {
    console.error('❌ Erro ao acessar planilha:', error.message);
    return [];
  }
}

/**
 * Processar uma pasta específica
 */
function processarPasta(refRenault, refGnh) {
  // Pular se a referência ainda é "X" (processo não digitado)
  if (refRenault.toUpperCase() === 'X') {
    console.log(`⏸️ Processo pendente de digitação (REF: X)`);
    return false;
  }
  
  // Possíveis nomes da pasta atual (incluindo versões com X)
  const possiveisNomes = [
    refGnh.replace('/', '_'),           // P628646_25
    refGnh.replace('/', '-'),           // P628646-25
    refGnh,                             // P628646/25
    refGnh.replace('/', '_') + '_X',    // P628646_25_X
    refGnh.replace('/', '-') + '_X',    // P628646-25_X
    refGnh.replace('/', '_') + '_x',    // P628646_25_x
    refGnh.replace('/', '-') + '_x'     // P628646-25_x
  ];
  
  // Nome final desejado
  const nomeNovo = `${refGnh.replace('/', '-')}_${refRenault}`;
  
  // Buscar pasta
  let pasta = null;
  let nomeEncontrado = '';
  for (const nome of possiveisNomes) {
    pasta = encontrarPasta(nome);
    if (pasta) {
      nomeEncontrado = nome;
      break;
    }
  }
  
  if (!pasta) {
    console.log(`❌ Pasta não encontrada: ${possiveisNomes.slice(0, 3).join(', ')}`);
    return false;
  }
  
  const nomeAtual = pasta.getName();
  console.log(`🔍 Pasta encontrada: ${nomeAtual}`);
  
  // Verificar se já está no formato correto
  if (nomeAtual === nomeNovo) {
    console.log(`ℹ️ Já no formato final: ${nomeAtual}`);
    return false;
  }
  
  // Verificar se já tem a referência (mas não é exatamente igual)
  if (nomeAtual.includes(refRenault) && !nomeAtual.toUpperCase().includes('_X')) {
    console.log(`ℹ️ Já contém a referência: ${nomeAtual}`);
    return false;
  }
  
  // Verificar conflito de nome
  const pastaExistente = encontrarPasta(nomeNovo);
  if (pastaExistente && pastaExistente.getId() !== pasta.getId()) {
    console.log(`⚠️ Conflito: já existe pasta "${nomeNovo}"`);
    return false;
  }
  
  // Verificar se é uma atualização de pasta com X
  const isAtualizacaoX = nomeAtual.toUpperCase().includes('_X');
  
  // Renomear
  pasta.setName(nomeNovo);
  if (isAtualizacaoX) {
    console.log(`🔄 ATUALIZADO: ${nomeAtual} → ${nomeNovo}`);
  } else {
    console.log(`📝 RENOMEADO: ${nomeAtual} → ${nomeNovo}`);
  }
  return true;
}

/**
 * Encontrar pasta no Drive
 */
function encontrarPasta(nomePasta) {
  try {
    // 1. Buscar no MPR/2025
    const pastaMpr = buscarNoMpr(nomePasta);
    if (pastaMpr) return pastaMpr;
    
    // 2. Buscar no PHF-DIV-OHS
    const pastaPhf = buscarNoPhf(nomePasta);
    if (pastaPhf) return pastaPhf;
    
    // 3. Busca geral
    const pastasGeral = DriveApp.getFoldersByName(nomePasta);
    return pastasGeral.hasNext() ? pastasGeral.next() : null;
    
  } catch (error) {
    console.error(`❌ Erro ao buscar "${nomePasta}":`, error.message);
    return null;
  }
}

/**
 * Buscar nas pastas mensais do MPR/2025
 */
function buscarNoMpr(nomePasta) {
  try {
    const pastaMpr2025 = DriveApp.getFolderById(PASTA_MPR_2025_ID);
    
    // Buscar diretamente na pasta 2025
    let pasta = obterSubpasta(pastaMpr2025, nomePasta);
    if (pasta) return pasta;
    
    // Buscar em cada pasta mensal
    for (const mes of MESES_MPR) {
      const pastaMes = obterSubpasta(pastaMpr2025, mes);
      if (pastaMes) {
        pasta = obterSubpasta(pastaMes, nomePasta);
        if (pasta) return pasta;
      }
    }
  } catch (error) {
    console.error(`❌ Erro no MPR: ${error.message}`);
  }
  
  return null;
}

/**
 * Buscar na pasta PHF-DIV-OHS
 */
function buscarNoPhf(nomePasta) {
  try {
    const pastaPhf = DriveApp.getFolderById(PASTA_PHF_DIV_OHS_ID);
    return obterSubpasta(pastaPhf, nomePasta);
  } catch (error) {
    console.error(`❌ Erro no PHF: ${error.message}`);
    return null;
  }
}

/**
 * Obter subpasta por nome
 */
function obterSubpasta(pastaPai, nomeSubpasta) {
  try {
    const subpastas = pastaPai.getFoldersByName(nomeSubpasta);
    return subpastas.hasNext() ? subpastas.next() : null;
  } catch (error) {
    return null;
  }
}

/**
 * Configurar execução automática (8h às 18h, de 4 em 4 horas)
 */
function configurarExecucaoAutomatica() {
  try {
    // Remover todos os triggers existentes
    const triggersExistentes = ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === 'renomearPastas');
    
    triggersExistentes.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    if (triggersExistentes.length > 0) {
      console.log(`🗑️ ${triggersExistentes.length} trigger(s) anterior(es) removido(s)`);
    }
    
    // Criar triggers para horários específicos: 8h, 12h, 16h
    const horarios = [8, 12, 16];
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('renomearPastas')
        .timeBased()
        .everyDays(1)
        .atHour(hora)
        .create();
    });
    
    console.log('✅ Execução automática configurada para:');
    console.log('   • 08:00 (início do expediente)');
    console.log('   • 12:00 (meio do dia)');
    console.log('   • 16:00 (final da tarde)');
    console.log('📝 Script só executa entre 8h e 18h (controle interno)');
    
    // Mostrar próximas execuções
    const agora = new Date();
    const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
    
    console.log('\n🕐 Próximas execuções:');
    horarios.forEach(hora => {
      const proximaExecucao = new Date(hoje);
      proximaExecucao.setHours(hora, 0, 0, 0);
      
      if (proximaExecucao <= agora) {
        proximaExecucao.setDate(proximaExecucao.getDate() + 1);
      }
      
      console.log(`   • ${proximaExecucao.toLocaleString()}`);
    });
    
  } catch (error) {
    console.error('❌ Erro ao configurar automação:', error.message);
  }
}

/**
 * Listar triggers ativos com horários
 */
function listarTriggers() {
  console.log('=== TRIGGERS ATIVOS ===');
  
  try {
    const triggers = ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === 'renomearPastas');
    
    if (triggers.length === 0) {
      console.log('❌ Nenhum trigger do script configurado');
      console.log('💡 Execute: configurarExecucaoAutomatica()');
      return;
    }
    
    console.log(`✅ ${triggers.length} trigger(s) ativo(s):`);
    
    triggers.forEach((trigger, index) => {
      const eventType = trigger.getEventType();
      console.log(`${index + 1}. Tipo: ${eventType}`);
      console.log(`   ID: ${trigger.getUniqueId()}`);
    });
    
    console.log('\n📋 Configuração: Executa diariamente às 8h, 12h e 16h');
    console.log('🕐 Controle interno: Só processa entre 8h e 18h');
    
  } catch (error) {
    console.error('❌ Erro ao listar triggers:', error.message);
  }
}

/**
 * Teste rápido com uma pasta específica
 */
function testeRapido() {
  console.log('=== TESTE RÁPIDO ===');
  
  // Exemplo 1: Processo com X (deve ser ignorado)
  console.log('\n1. Testando processo pendente:');
  try {
    const resultado1 = processarPasta('X', 'P628646/25');
    console.log(`Resultado: ${resultado1 ? 'Processado' : 'Ignorado (correto)'}`);
  } catch (error) {
    console.error('❌ Erro:', error.message);
  }
  
  // Exemplo 2: Processo com referência real (deve ser processado)
  console.log('\n2. Testando processo com referência:');
  try {
    const resultado2 = processarPasta('MQBPYA003325', 'P628646/25');
    console.log(`Resultado: ${resultado2 ? 'Processado' : 'Não processado'}`);
  } catch (error) {
    console.error('❌ Erro:', error.message);
  }
}

/**
 * Função para primeira configuração (executar uma vez apenas)
 */
function primeiraConfiguracao() {
  console.log('=== PRIMEIRA CONFIGURAÇÃO ===');
  
  // Testar acesso básico
  try {
    const planilha = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log(`✅ Planilha acessível: ${planilha.getName()}`);
    
    const pastaMpr = DriveApp.getFolderById(PASTA_MPR_2025_ID);
    console.log(`✅ Pasta MPR acessível: ${pastaMpr.getName()}`);
    
    const pastaPhf = DriveApp.getFolderById(PASTA_PHF_DIV_OHS_ID);
    console.log(`✅ Pasta PHF acessível: ${pastaPhf.getName()}`);
    
    // Mostrar horário atual
    const agora = new Date();
    const hora = agora.getHours();
    console.log(`⏰ Horário atual: ${agora.toLocaleString()} (${hora}h)`);
    
    if (hora >= 8 && hora < 18) {
      console.log('✅ Dentro do horário de trabalho');
    } else {
      console.log('⚠️ Fora do horário de trabalho (8h às 18h)');
    }
    
    console.log('\n🎉 Configuração OK! Próximos passos:');
    console.log('• Execute: renomearPastas() - para testar uma execução manual');
    console.log('• Execute: configurarExecucaoAutomatica() - para ativar automação');
    console.log('• Execute: listarTriggers() - para ver triggers ativos');
    
    console.log('\n📅 Horários de execução automática:');
    console.log('• 08:00 - Início do expediente');
    console.log('• 12:00 - Meio do dia');  
    console.log('• 16:00 - Final da tarde');
    console.log('• Não executa fora do horário 8h-18h');
    
  } catch (error) {
    console.error('❌ Erro na configuração:', error.message);
    console.log('\n🔧 Verifique:');
    console.log('• IDs das pastas estão corretos');
    console.log('• Você tem permissão de acesso');
    console.log('• Execute as autorizações necessárias');
  }
}