# 🔐 Identity Manager & AD Sync - Magalog Suporte Infra (v1.3.0)

## 🚀 Visão Geral
O **Identity Manager** evoluiu de uma ferramenta simples de reset para um ecossistema completo de gestão de identidades e automação de Active Directory. O sistema utiliza uma arquitetura híbrida para unir a agilidade da nuvem (Google Cloud/Apps Script) com a autoridade de execução local (PowerShell/Active Directory).

## 🏗️ Arquitetura do Sistema
O ecossistema é composto por três pilares fundamentais:

1.  **Orquestrador Web (Middleware/Backend)**: Desenvolvido em **Google Apps Script**, gerencia a lógica de negócio, autenticação (SSO), fila de solicitações e auditoria.
2.  **Interface de Governança (Frontend)**: Uma Single Page Application (SPA) moderna em **Vue.js**, oferecendo uma experiência premium (Magalu Style) para analistas e usuários, com busca em tempo real via **BigQuery**.
3.  **Daemon de Execução (Worker)**: Um serviço **PowerShell** resiliente (`Unified_AD_Daemon.ps1`) que atua como o braço operacional no domínio, processando tarefas da fila e realizando as alterações diretamente no Active Directory.

---

## 💎 Pilares de Funcionalidade

### 1. Fila Unificada de Atendimento
Centralização absoluta de demandas. O sistema não distingue apenas resets; ele gerencia fluxos complexos em uma única esteira:
*   **Reset de Senha**: Automação total (Reset + Desbloqueio + Troca Obrigatória).
*   **Account Unlock**: Desbloqueio técnico sem alteração de credenciais.
*   **User Mirroring**: Clonagem inteligente de grupos de segurança entre usuários modelo e alvos.

### 2. Motor de SLA & Governança (v1.3.0)
Garantia de atendimento e conformidade:
*   **Monitoramento Ativo**: Alerta automático para qualquer chamado pendente há mais de **2 horas**.
*   **Lembretes Recorrentes**: Notificações horárias aos analistas com templates premium.
*   **Cessação Inteligente**: O motor de alerta interrompe os disparos instantaneamente após a ação técnica.

### 3. Compliance & Auditoria
*   **Identificação SSO**: Cada ação é atrelada à sessão Google do analista responsavél, garantindo o "quem" e o "quando".
*   **Schema Dinâmico**: O backend gerencia o próprio banco de dados, corrigindo cabeçalhos e garantindo a integridade dos 17 campos de dados.

---

## 📊 Schema de Dados (Aba: Solicitações)
A estrutura de dados é otimizada para performance e histórico:

| ID | Col | Campo | Descrição |
|:---:|:---:|:---|:---|
| 1 | **A** | `ID` | Identificador único auto-incremental (Primary Key). |
| 2 | **B** | `DATA_HORA` | Timestamp de criação da demanda. |
| 3 | **C** | `FILIAL` | Unidade de negócio de origem. |
| 4 | **D** | `USER_NAME` | Login de rede do colaborador alvo. |
| 5 | **E** | `NOME` | Nome completo do colaborador. |
| 6 | **F** | `EMAIL_COLAB` | E-mail para recebimento de credenciais. |
| 7 | **G** | `CENTRO_CUSTO` | Dados de lotação orçamentária. |
| 8 | **H** | `ANALISTA_RESPONSAVEL` | E-mail corporativo do analista que atendeu/gerou a demanda. |
| 9 | **I** | `SOLICITANTE` | E-mail de quem abriu a demanda via portal. |
| 10 | **J** | `STATUS_PROCESSAMENTO` | Status técnico no AD (`PENDENTE`, `CONCLUIDO`, `ERRO`). |
| 11 | **K** | `STATUS_APROVACAO` | Fluxo humano (`PENDENTE`, `APROVADO`, `REPROVADO`). |
| 12 | **L** | `TIPO_TAREFA` | Categoria da ação (`RESET`, `UNLOCK`, `MIRROR`). |
| 13 | **M** | `DETALHES_ADICIONAIS` | Logs técnicos e mensagens de erro do Daemon. |
| 14 | **N** | `MODELO` | Usuário de referência (Exclusivo para Mirror). |
| 15 | **O** | `DESTINOS` | Lista de usuários alvos (JSON - Exclusivo para Mirror). |
| 16 | **P** | `GRUPOS` | Lista de grupos a serem sincronizados (JSON). |
| 17 | **Q** | `ULTIMO_LEMBRETE` | Timestamp de controle do motor de SLA. |

---

## 🛠️ Manutenção e Deploy

### Backend/Frontend (Apps Script)
Utilizamos o **Clasp CLI** para versionamento e deploy seguro. Mantendo sempre o mesmo `Deployment ID` para não quebrar a URL fixa de produção.
```bash
clasp push
clasp deploy -i <ProdID> -d "Release v1.3.0"
```

### Daemon (Local)
O Daemon deve rodar como **Tarefa Agendada (GPO/Task Scheduler)** em um servidor com acesso ao módulo ActiveDirectory.
*   **Configuração**: `$API_URL` deve apontar para o Web App publicado.
*   **Logs**: Localizados em `C:\ProgramData\ADResetTool\Logs`.

---

## 📜 Histórico de Versões Relevantes

### v1.3.0 (Atividade Atual)
- Implementação de **Sistema de SLA** com alertas dinâmicos.
- Template HTML de e-mail dedicado para monitoramento.
- Autovigência de cabeçalhos (`ensureHeaders`).

### v1.2.0 (Unificação)
- Consolidação de abas. O sistema abandonou as tabelas separadas para operar em um modelo unificado de 17 colunas.
- Badging de interface para visualização clara de tipos de tarefa.

### v1.1.0/v1.1.6
- Introdução de **Google SSO** para auditoria.
- Adição de fluxos de **Espelhamento de Grupos**.
- Refatoração total da UI para Magalu Style.

---
**Responsável Técnico**: Leandro Araújo - Suporte Infra CDs Magalog
