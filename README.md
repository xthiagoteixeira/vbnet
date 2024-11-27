# Análise do Sistema Escolar

Este é um sistema de gestão escolar desenvolvido em VB.NET usando o framework DevExpress, focado no gerenciamento de notas, frequências e boletins escolares.

## 🏗️ Arquitetura do Sistema

O sistema é composto por diferentes módulos:

### 📊 Módulo de Importação (frmImportarExcel.vb)
- Permite importar dados de planilhas Excel (.xls, .xlsx, .xltx)
- Converte os arquivos para formato CSV intermediário 
- Realiza validações dos dados importados
- Suporta importação em lote de múltiplos arquivos
- Atualiza banco de dados MySQL com notas e frequências

### 📝 Módulo do Professor (frmMaisProfessor.vb)
- Interface principal para professores lançarem notas e frequências
- Gerenciamento de turmas, disciplinas e bimestres  
- Cálculo automático de médias e estatísticas
- Suporte a avaliações finais e recuperações
- Sistema de backup local dos dados

### 📋 Módulo de Relatórios (frmRpt.vb)
- Geração de boletins e relatórios usando DevExpress
- Visualização e impressão de documentos
- Filtros por período, turma e disciplina
- Formatação personalizada de relatórios

### 🌐 Web Service (WebService_Cadastro.asmx)
- Sincronização de dados com servidor
- Consultas de notas e frequências
- Deliberações e gestão de fichas de alunos
- Validação de usuários e permissões

## 🔧 Tecnologias Utilizadas

- VB.NET Windows Forms
- DevExpress Components
- MySQL Database
- Web Services ASMX
- Crystal Reports
- Office Automation

## 📦 Recursos Principais

1. **Gestão de Notas**
   - Lançamento por bimestre
   - Cálculo automático de médias
   - Suporte a recuperações
   - Validações de regras de negócio

2. **Controle de Frequência** 
   - Registro de presenças/faltas
   - Cálculo de percentuais
   - Alertas de frequência mínima
   - Compensação de ausências

3. **Importação de Dados**
   - Suporte a múltiplos formatos Excel
   - Validação na importação
   - Processamento em lote
   - Log de operações

4. **Relatórios**
   - Boletins individuais
   - Mapas de notas por turma
   - Estatísticas de desempenho 
   - Exportação em diversos formatos

5. **Segurança**
   - Controle de acesso por perfil
   - Backup automático 
   - Log de alterações
   - Sincronização segura

## 👥 Módulos Administrativos

### 📚 Gestão de Turmas (frmAdmTurma.vb)
- Cadastro e manutenção de turmas
- Controle de períodos (Manhã, Tarde, Noite, etc.)
- Níveis de ensino configuráveis (EJA, Fundamental, Médio, etc.)
- Sistema de bloqueio/desbloqueio
- Validações de exclusão com verificação de boletins
- Sincronização com boletim web

### 👤 Gestão de Usuários (frmAdmUsuarios.vb)
- Gerenciamento de credenciais
- Sistema de bloqueio/desbloqueio de acesso
- Alteração segura de senhas
- Validação de duplicidade
- Interface grid com status visual

## 📦 Módulos Adicionais

### 🔄 Módulo de Suporte (frmF10ChamaSuporte.vb)
- Atalho global Alt+F10 para acesso rápido
- Integração com TeamViewer QuickSupport
- Sistema de manutenção remota
- Monitoramento de processos ativos

### 💰 Módulo Financeiro (frmF2b_RegistroMensal.vb)
- Gerenciamento de cobranças mensais
- Geração automática de boletos
- Controle de status de pagamentos
- Histórico de transações
- Integração com F2B para boletos

### 🔍 Módulo de Consultas (frmF2b_SituacaoCobranca_Historico.vb) 
- Histórico detalhado de cobranças
- Filtros por período
- Status de pagamentos
- Visualização de valores e vencimentos

### ⚙️ Módulo de Gestão (frmExcluirTurmas.vb)
- Gerenciamento de turmas ativas/inativas
- Exclusão controlada de registros
- Visualização de boletins por turma
- Validações de integridade

## 💻 Recursos Técnicos Adicionais

### Integração com TeamViewer
```vb
' Exemplo de inicialização do TeamViewer
Process.Start(String.Format("{0}\{1}\TeamViewerQS_pt.exe", 
    meucaminhorelatorio, NomeDoProjeto))
