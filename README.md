# 🔄 Renomeador Automático de Pastas - Google Drive

## 📋 Descrição

Script automatizado para renomeação de pastas no Google Drive baseado em dados de planilha Google Sheets. O sistema realiza a atualização automática de nomes de pastas conforme os dados são atualizados na planilha, ideal para organizações que precisam manter estruturas de pastas sincronizadas com sistemas de controle.

## ✨ Funcionalidades

- **Renomeação Automática**: Atualiza nomes de pastas baseado em dados da planilha
- **Controle de Horário**: Executa apenas durante horário comercial (8h às 18h)
- **Busca Inteligente**: Procura pastas em locais específicos configuráveis
- **Tratamento de Estados**: Gerencia pastas com status pendente e com dados completos
- **Execução Programada**: Roda automaticamente 3x por dia (8h, 12h, 16h)
- **Logs Detalhados**: Relatórios completos de todas as operações

## 🏗️ Estrutura do Projeto

```
📦 renomeador-pastas-drive/
├── 📄 script.gs              # Script principal do Google Apps Script
├── 📄 README.md              # Documentação em português
├── 📄 README-EN.md           # Documentação em inglês
└── 📄 LICENSE                # Licença do projeto
```

## ⚙️ Configuração

### 1. Configurações Iniciais

Edite as constantes no início do script:

```javascript
// ID da planilha Google Sheets
const SPREADSHEET_ID = 'SEU_ID_DA_PLANILHA_AQUI';

// Nome da aba da planilha
const SHEET_NAME = 'NOME_DA_ABA';

// IDs das pastas do Google Drive
const PASTA_PRINCIPAL_ID = 'ID_DA_PASTA_PRINCIPAL';
const PASTA_SECUNDARIA_ID = 'ID_DA_PASTA_SECUNDARIA';
```

### 2. Estrutura da Planilha

| Coluna A | Coluna B |
|----------|----------|
| CÓDIGO PRINCIPAL | CÓDIGO SECUNDÁRIO |
| ABC123XYZ | DOC001/25 |
| X | DOC002/25 |

- **Coluna A**: Código principal (ou "X" para processos pendentes)
- **Coluna B**: Código secundário (formato personalizável)

### 3. Instalação

1. Acesse [Google Apps Script](https://script.google.com)
2. Crie um novo projeto
3. Cole o código do arquivo `script.gs`
4. Configure as permissões necessárias
5. Execute `primeiraConfiguracao()` para validar setup

## 🚀 Como Usar

### Execução Manual

```javascript
// Teste de configuração
primeiraConfiguracao();

// Execução única
renomearPastas();

// Teste com casos específicos
testeRapido();
```

### Execução Automática

```javascript
// Configurar automação (executar uma vez)
configurarExecucaoAutomatica();

// Verificar triggers ativos
listarTriggers();
```

## 📊 Funcionamento

### Fluxo de Renomeação

1. **Leitura da Planilha**: Carrega dados das colunas A e B
2. **Filtro de Dados**: Ignora linhas vazias e processos com "X"
3. **Busca de Pastas**: Procura pastas nos locais configurados
4. **Validação**: Verifica se renomeação é necessária
5. **Renomeação**: Atualiza nome para formato padrão
6. **Log**: Registra resultado da operação

### Padrões de Nome

- **Antes**: `DOC001_25_X` ou `DOC001-25`
- **Depois**: `DOC001-25_ABC123XYZ`

### Locais de Busca

1. **Pasta Principal**: Pasta principal e subpastas configuradas
2. **Pasta Secundária**: Pasta específica do departamento/setor
3. **Busca Geral**: Fallback para todo o Drive

## ⏰ Horários de Execução

- **08:00**: Início do expediente
- **12:00**: Meio do dia
- **16:00**: Final da tarde

> **Nota**: Script só executa entre 8h e 18h, mesmo com triggers configurados

## 📈 Relatórios

### Exemplo de Saída

```
=== INICIANDO RENOMEAÇÃO ===
⏰ 17/06/2025 08:00:00 (8h - Horário válido)
📊 150 linhas encontradas

--- Linha 5: ABC123XYZ | DOC001/25 ---
🔍 Pasta encontrada: DOC001_25_X
🔄 ATUALIZADO: DOC001_25_X → DOC001-25_ABC123XYZ
✅ Renomeada

=== RESUMO ===
✅ Processadas: 45 | Renomeadas: 12
⏸️ Pendentes (X): 98 | ❌ Erros: 0
```

## 🔒 Permissões Necessárias

- **Google Sheets**: Leitura da planilha de dados
- **Google Drive**: Busca e renomeação de pastas
- **Google Apps Script**: Execução de triggers temporais

## 🐛 Troubleshooting

### Problemas Comuns

**Erro de Permissão**
```
❌ Erro: Exception: You do not have permission to call DriveApp.getFolderById
```
- Solução: Executar `primeiraConfiguracao()` e autorizar permissões

**Aba Não Encontrada**
```
❌ Aba "NOME_DA_ABA" não encontrada
```
- Solução: Verificar `SPREADSHEET_ID` e `SHEET_NAME`

**Pasta Não Encontrada**
```
❌ Pasta não encontrada: DOC001_25, DOC001-25, DOC001/25
```
- Solução: Verificar se pasta existe nos locais configurados

### Debug

```javascript
// Listar triggers ativos
listarTriggers();

// Teste com dados específicos
testeRapido();

// Verificar configuração
primeiraConfiguracao();
```

## 📝 Changelog

### v1.0.0 (Atual)
- ✅ Renomeação automática baseada em planilha
- ✅ Controle de horário comercial
- ✅ Busca em múltiplos locais
- ✅ Tratamento de estados X
- ✅ Execução programada
- ✅ Sistema de logs detalhados

## 🤝 Contribuindo

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 👥 Autores

- **Herberth Goldan Cardoso Junior** - *Desenvolvimento inicial* - https://github.com/HerberthGoldanJr

## 📞 Suporte

- 📧 Email: herberthgoldanjr@gmail.com


---

⭐ **Gostou do projeto? Deixe uma estrela!**
