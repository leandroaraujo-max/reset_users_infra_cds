# 🔐 Gerenciamento de Usuários - Suporte Infra CDs v1.0.5

## Sobre o Projeto
Ferramenta desenvolvida em PowerShell com interface gráfica (Windows Forms) para automatizar o processo de reset de senhas de usuários do Active Directory e criações de conta no Turia.

O sistema integra-se com uma planilha Google Sheets (via Apps Script API) para buscar demandas de resets pendentes e auditar as ações executadas.
 
 ## URL de Acesso (Produção)
 **Link Fixo:** [Acessar Ferramenta Web](https://script.google.com/macros/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec)
 *Use este link para acessar a interface web e para atualizações futuras.*

> [!IMPORTANT]
> **DEPLOYMENT VIA CLASP:**
> Toda atualização via `clasp` DEVE manter a URL Fixa acima.
> Consulte o arquivo `Arquivos App Script/Url_Fixa.txt` para conferir a URL.
> Para deploy mantendo a URL, use o comando: `clasp deploy -i <DeploymentID> -d "Descrição"`


## Funcionalidades Principais
- **Listagem de Pendências:** Busca solicitações de reset via API.
- **Reset Automático:** Reseta senha, desbloqueia conta, força troca no próximo logon e ativa a conta.
- **Envio de Emails:**
  - Envia credenciais para o colaborador ou gestor (via SMTP Interno).
  - Envia instruções de criação de conta no Turia se usuário não existir no AD.
- **Auditoria:** Registra todas as ações em planilha na nuvem e logs locais.
- **Resiliência:** Sistema de retentativa automática (Retry) para falhas de rede.
- **Web Interface (Frontend):**
  - Sistema de autenticação com login/senha e opção **"Esqueci Minha Senha"**.
  - Navegação intuitiva: Clique no título para voltar à Home.
  - Solicitação de acesso e recuperação de senha.
  - Busca avançada por **Nome**, **ID Magalu**, **Usuário de Rede** ou **Email** com **ordenação de colunas**.
  - Busca flexível: Pode pesquisar **sem selecionar filial** (Placeholder: "Digite sua filial Magalog").
  - Funcionalidade **"Lembrar-me"** para salvar credenciais locais.
  - Fila de acompanhamento completa (sem limites) com **Filtro por Filial** e **ID sequencial**.

## Pré-Requisitos
1. **Sistema Operacional:** Windows 10/11 ou Server (com PowerShell 5.1+).
2. **Permissões:** Usuário deve ter permissão de reset de senha no AD.
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
| Arquivo | Descrição |
|---------|-----------|
| `Reset_users_Infra_cds.ps1` | Script principal (Core PowerShell) |
| `Iniciar_Reset_users_Infra_cds.bat` | Launcher para execução fácil |
| `AppsScript_Backend.js` | Backend Google Apps Script |
| `AppsScript_Web_Index.html` | Interface Web (Frontend Vue.js) |
| `Frontend.js` | Scripts complementares (Relatórios BigQuery) |
| `.clasp.json` | Configuração do Clasp CLI |
| `appsscript.json` | Manifesto do projeto Apps Script |

## Deploy via Clasp
O projeto utiliza [Clasp](https://github.com/google/clasp) para deploy automatizado:

```bash
# Instalar Clasp (requer Node.js)
npm install -g @google/clasp

# Login
clasp login

# Push (enviar código)
clasp push

# Deploy (atualizar produção)
clasp deploy -i <DEPLOYMENT_ID> -d "Descrição"
```

## Solução de Problemas
| Problema | Solução |
|----------|---------|
| Erro de Módulo AD | Instale o RSAT (Remote Server Administration Tools) |
| Tela travada | Operações pesadas de AD podem causar leve delay |
| Falha de API | Verifique conexão com internet (3 retentativas automáticas) |
| ID não aparece na fila | Verifique se a coluna ID existe na aba "Solicitações" |

## Histórico de Versões

### v1.0.6 (Atual)
- [Daemon] Corrigido loop infinito de "ID não encontrado" (parâmetro `requestId` vs `id`).
- [Backend] Corrigido erro "Parâmetros inválidos" no link de aprovação por email.
- [Backend] Atualizado para aceitar status `GRUPOS_ENCONTRADOS` do Daemon como sucesso.
- [UI] Logo LuizaLabs agora branco e sem fundo no header para melhor contraste.
- [Deploy] Deployment URL fixada e sincronizada em todos os arquivos.

### v1.0.5
- [PowerShell] Refinamento de layout: Ordem correta do Header e Faixa Rainbow
- [PowerShell] Correção de sobreposição de textos no cabeçalho
- [PowerShell] Fix SSL/CRL: Adicionado bypass de revogação para conexão estável com a API
- [Frontend] Novo link "Esqueci Minha Senha" na tela de login
- [Frontend] Navegação "Voltar para Home" ao clicar no título
- [Frontend] Placeholder de filial atualizado para "Digite sua filial Magalog"
- [Frontend] Tabelas com ordenação de colunas (sortable columns)
- [Frontend] Fila de Atendimento sem limite de linhas e com filtro por Filial
- [Backend] ID auto-incremental nas abas **Auditoria** e **Solicitações**
- [UI] Refinamento estético geral (Look & Feel Magalu)
- [UI] Título do sistema unificado como "Reset de Usuários - Suporte Infra CDs"

### v1.0.4
- [Backend] Adicionado campo ID na Auditoria
- [Backend] Função `NUMERAR_AUDITORIA_EXISTENTE()`

### v1.0.3
- [Frontend] Corrigido alinhamento dos botões no modal de confirmação
- [Frontend] Filial preenchida automaticamente ao buscar por ID
- [Backend] Retorna filial do colaborador no resultado da busca

### v1.0.2
- [Frontend] Adicionada coluna **ID** na tabela de resultados de busca
- [Backend] Query SQL atualizada para retornar `t2.ID`

### v1.0.1
- [Frontend] Adicionado busca por ID Magalu
- [Frontend] Adicionado checkbox "Lembrar-me"
- [Backend] Atualizações de segurança e versão API

### v1.0.0
- Release inicial

---
**Desenvolvido por:** Leandro Araújo - Suporte Infra CDs
