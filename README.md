# 🔐 AD Reset Tool v1.0.1

## Sobre o Projeto
Ferramenta desenvolvida em PowerShell com interface gráfica (Windows Forms) para automatizar o processo de reset de senhas de usuários do Active Directory e criações de conta no Turia.

O sistema integra-se com uma planilha Google Sheets (via Apps Script API) para buscar demandas de resets pendentes e auditar as ações executadas.

## Funcionalidades Principais
- **Listagem de Pendências:** Busca solicitações de reset via API.
- **Reset Automático:** Reseta senha, desbloqueia conta, força troca no próximo logon e ativa a conta.
- **Envio de Emails:**
  - Envia credenciais para o colaborador ou gestor (via SMTP Interno).
  - Envia instruções de criação de conta no Turia se usuário não existir no AD.
- **Auditoria:** Registra todas as ações em planilha na nuvem e logs locais.
- **Resiliência:** Sistema de retentativa automática (Retry) para falhas de rede.
- **Web Interface (Frontend):**
  - Solicitação de acesso e reset de senha pelo usuário.
  - Busca por **Nome**, **ID Magalu**, **Usuário de Rede** ou **Email**.
  - Funcionalidade **"Lembrar-me"** para salvar credenciais locais.

## Pré-Requisitos
1. **Sistema Operacional:** Windows 10/11 ou Server (com PowerShell 5.1+).
2. **Permissões:** Usuario deve ter permissão de reset de senha no AD.
3. **Módulo Active Directory:** RSAT instalado (`Import-Module ActiveDirectory`).
4. **Acesso à Rede:** 
   - Acesso à Internet (Google Apps Script).
   - Acesso ao SMTP Interno (`smtpml.magazineluiza.intranet`, Porta 25).

## Como Executar
1. Clone ou baixe este repositório.
2. Execute o arquivo `Iniciar_Reset_users_Infra_cds.bat` (ou execute o `.ps1` via PowerShell).
3. Selecione seu nome na lista de analistas.
4. Digite a filial desejada ou use `*` para todas.
5. Clique em **Carregar Demandas**.
6. Clique em **EXECUTAR PROCESSO**.

## Estrutura de Arquivos
- `Reset_users_Infra_cds.ps1`: Script principal (Core).
- `Iniciar_Reset_users_Infra_cds.bat`: Launcher para execução fácil.
- `AppsScript_Backend_v1.0.0.txt`: Código do backend (Google Apps Script).
- `AppsScript_Web_Index_v1.0.0.html`: Interface Web (Frontend) v1.0.1.
- `Logs/`: Diretório onde são salvos os logs de execução (`C:\ProgramData\ADResetTool\Logs`).

## Solução de Problemas
- **Erro de Módulo AD:** Instale o RSAT (Remote Server Administration Tools).
- **Tela travada:** O script usa `DoEvents` para manter a interface responsiva, mas operações pesadas de AD podem causar leve delay.
- **Falha de API:** Verifique sua conexão com a internet. O sistema tentará 3 vezes antes de falhar.

## Histórico de Versões
- **v1.0.1 (Atual):**
  - [Frontend] Adicionado busca por ID Magalu.
  - [Frontend] Adicionado checkbox "Lembrar-me".
  - [Backend] Atualizações de segurança e versão API.
- **v1.0.0:** Release inicial.
