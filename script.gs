/**
 * Script para renomeação automática de pastas no Google Drive
 * baseado nos dados de uma planilha do Google Sheets
 * VERSÃO TEMPLATE - Configure suas próprias informações
 */

// ===============================
// CONFIGURAÇÕES - EDITE AQUI
// ===============================

// ID da planilha do Google Sheets (encontre na URL da planilha)
const SPREADSHEET_ID = 'SEU_SPREADSHEET_ID_AQUI';

// Nome da aba da planilha
const SHEET_NAME = 'NOME_DA_ABA';

// IDs das pastas onde procurar os processos
const PASTA_PRINCIPAL_ID = 'ID_DA_PASTA_PRINCIPAL';
const PASTA_SECUNDARIA_ID = 'ID_DA_PASTA_SECUNDARIA';

// Colunas da planilha (números das colunas)
const COL_REFERENCIA_ORIGEM = 1;  // Coluna A - Referência original
const COL_REFERENCIA_DESTINO = 2; // Coluna B - Nova referência

// Subpastas para busca (adapte conforme sua estrutura)
const SUBPASTAS_BUSCA = [
  '01 - JANEIRO', '02 - FEVEREIRO', '03 - MARÇO', '04 - ABRIL',
  '05 - MAIO', '06 - JUNHO', '07 - JULHO', '08 - AGOSTO',
  '09 - SETEMBRO', '10 - OUTUBRO', '11 - NOVEMBRO', '12 - DEZEMBRO'
];

// Horários de trabalho (formato 24h)
const HORA_INICIO = 8;  // 8h
const HORA_FIM = 18;    // 18h

// ===============================
// CÓDIGO PRINCIPAL
// ===============================

/**
 * Função principal - executa a renomeação (com controle de horário)
 */
function renomearPastas() {
  try {
    // Verificar se está no horário de trabalho
    const agora = new Date();
    const hora = agora.getHours();
    
    if (hora < HORA_INICIO || hora >= HORA_FIM) {
      console.log(`⏰ Fora do horário de trabalho (${hora}h). Script pausado até ${HORA_INICIO}h.`);
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
      const refOrigem = linha[COL_REFERENCIA_ORIGEM - 1];
      const refDestino = linha[COL_REFERENCIA_DESTINO - 1];
      
      // Pular cabeçalho (adapte conforme seus dados)
      if (refOrigem === 'REFERENCIA_ORIGEM' || refDestino === 'REFERENCIA_DESTINO') continue;
      
      // Verificar dados válidos
      if (!refOrigem || !refDestino || 
          !refOrigem.toString().trim() || 
          !refDestino.toString().trim()) {
        continue;
      }

      // Contar processos com X (indicando processo pendente)
      if (refOrigem.toString().toUpperCase().trim() === 'X') {
        ignoradosX++;
        continue;
      }

      console.log(`\n--- Linha ${i + 1}: ${refOrigem} | ${refDestino} ---`);

      try {
        if (processarPasta(refOrigem.toString().trim(), refDestino.toString().trim())) {
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
function processarPasta(refOrigem, refDestino) {
  // Pular se a referência ainda é "X" (processo não digitado)
  if (refOrigem.toUpperCase() === 'X') {
    console.log(`⏸️ Processo pendente de digitação (REF: X)`);
    return false;
  }
  
  // Possíveis nomes da pasta atual (adapte conforme seu padrão)
  const possiveisNomes = [
    refDestino.replace('/', '_'),           // Formato com underscore
    refDestino.replace('/', '-'),           // Formato com hífen
    refDestino,                             // Formato original
    refDestino.replace('/', '_') + '_X',    // Com indicador X
    refDestino.replace('/', '-') + '_X',    
    refDestino.replace('/', '_') + '_x',    
    refDestino.replace('/', '-') + '_x'     
  ];
  
  // Nome final desejado (adapte o padrão conforme necessário)
  const nomeNovo = `${refDestino.replace('/', '-')}_${refOrigem}`;
  
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
  if (nomeAtual.includes(refOrigem) && !nomeAtual.toUpperCase().includes('_X')) {
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
    // 1. Buscar na pasta principal
    const pastaPrincipal = buscarNaPastaPrincipal(nomePasta);
    if (pastaPrincipal) return pastaPrincipal;
    
    // 2. Buscar na pasta secundária
    const pastaSecundaria = buscarNaPastaSecundaria(nomePasta);
    if (pastaSecundaria) return pastaSecundaria;
    
    // 3. Busca geral
    const pastasGeral = DriveApp.getFoldersByName(nomePasta);
    return pastasGeral.hasNext() ? pastasGeral.next() : null;
    
  } catch (error) {
    console.error(`❌ Erro ao buscar "${nomePasta}":`, error.message);
    return null;
  }
}

/**
 * Buscar nas subpastas da pasta principal
 */
function buscarNaPastaPrincipal(nomePasta) {
  try {
    const pastaPrincipal = DriveApp.getFolderById(PASTA_PRINCIPAL_ID);
    
    // Buscar diretamente na pasta principal
    let pasta = obterSubpasta(pastaPrincipal, nomePasta);
    if (pasta) return pasta;
    
    // Buscar em cada subpasta
    for (const subpasta of SUBPASTAS_BUSCA) {
      const pastaSubpasta = obterSubpasta(pastaPrincipal, subpasta);
      if (pastaSubpasta) {
        pasta = obterSubpasta(pastaSubpasta, nomePasta);
        if (pasta) return pasta;
      }
    }
  } catch (error) {
    console.error(`❌ Erro na pasta principal: ${error.message}`);
  }
  
  return null;
}

/**
 * Buscar na pasta secundária
 */
function buscarNaPastaSecundaria(nomePasta) {
  try {
    const pastaSecundaria = DriveApp.getFolderById(PASTA_SECUNDARIA_ID);
    return obterSubpasta(pastaSecundaria, nomePasta);
  } catch (error) {
    console.error(`❌ Erro na pasta secundária: ${error.message}`);
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
 * Configurar execução automática
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
    
    // Criar triggers para horários específicos
    const horarios = [HORA_INICIO, 12, HORA_FIM - 2]; // Ex: 8h, 12h, 16h
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('renomearPastas')
        .timeBased()
        .everyDays(1)
        .atHour(hora)
        .create();
    });
    
    console.log('✅ Execução automática configurada para:');
    horarios.forEach(hora => {
      console.log(`   • ${hora.toString().padStart(2, '0')}:00`);
    });
    console.log(`📝 Script só executa entre ${HORA_INICIO}h e ${HORA_FIM}h (controle interno)`);
    
  } catch (error) {
    console.error('❌ Erro ao configurar automação:', error.message);
  }
}

/**
 * Listar triggers ativos
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
    
    console.log(`\n📋 Configuração: Executa diariamente em horários específicos`);
    console.log(`🕐 Controle interno: Só processa entre ${HORA_INICIO}h e ${HORA_FIM}h`);
    
  } catch (error) {
    console.error('❌ Erro ao listar triggers:', error.message);
  }
}

/**
 * Teste rápido com dados de exemplo
 */
function testeRapido() {
  console.log('=== TESTE RÁPIDO ===');
  
  // Exemplo 1: Processo com X (deve ser ignorado)
  console.log('\n1. Testando processo pendente:');
  try {
    const resultado1 = processarPasta('X', 'EXEMPLO123/25');
    console.log(`Resultado: ${resultado1 ? 'Processado' : 'Ignorado (correto)'}`);
  } catch (error) {
    console.error('❌ Erro:', error.message);
  }
  
  // Exemplo 2: Processo com referência real (deve ser processado)
  console.log('\n2. Testando processo com referência:');
  try {
    const resultado2 = processarPasta('REF123456', 'EXEMPLO123/25');
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
    
    const pastaPrincipal = DriveApp.getFolderById(PASTA_PRINCIPAL_ID);
    console.log(`✅ Pasta principal acessível: ${pastaPrincipal.getName()}`);
    
    const pastaSecundaria = DriveApp.getFolderById(PASTA_SECUNDARIA_ID);
    console.log(`✅ Pasta secundária acessível: ${pastaSecundaria.getName()}`);
    
    // Mostrar horário atual
    const agora = new Date();
    const hora = agora.getHours();
    console.log(`⏰ Horário atual: ${agora.toLocaleString()} (${hora}h)`);
    
    if (hora >= HORA_INICIO && hora < HORA_FIM) {
      console.log('✅ Dentro do horário de trabalho');
    } else {
      console.log(`⚠️ Fora do horário de trabalho (${HORA_INICIO}h às ${HORA_FIM}h)`);
    }
    
    console.log('\n🎉 Configuração OK! Próximos passos:');
    console.log('• Execute: renomearPastas() - para testar uma execução manual');
    console.log('• Execute: configurarExecucaoAutomatica() - para ativar automação');
    console.log('• Execute: listarTriggers() - para ver triggers ativos');
    
  } catch (error) {
    console.error('❌ Erro na configuração:', error.message);
    console.log('\n🔧 Verifique:');
    console.log('• IDs das pastas estão corretos');
    console.log('• Você tem permissão de acesso');
    console.log('• Execute as autorizações necessárias');
  }
}
