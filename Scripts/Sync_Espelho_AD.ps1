# ==============================================================================
# AGENTE DE SINCRONIZAÇÃO - ESPELHO DE USUÁRIOS (AD BRIDGE)
# ==============================================================================
# Este script deve ser agendado no Task Scheduler (ex: a cada 1 minuto)
# Requisito: RSAT Active Directory instalado e acesso à Internet.

# URL da sua API (Fixa)
$API_URL = "https://script.google.com/macros/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec"

$LogDir = "$PSScriptRoot\..\Logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$LOG_FILE = "$LogDir\Log_Espelho_AD_$(Get-Date -Format 'yyyy-MM-dd').txt"

function Write-Log {
    param($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LOG_FILE -Value "[$TimeStamp] $Message"
}

# 1. Verifica Módulo AD
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ERRO: Módulo ActiveDirectory não encontrado."
    exit
}
Import-Module ActiveDirectory

try {
    # 2. Consulta a Fila na Nuvem (Google Sheets)
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Iniciando sicronização de Espelho AD..." -ForegroundColor White
    Write-Log "Iniciando ciclo de verificação..."

    Write-Host "Conectando API..." -NoNewline -ForegroundColor Gray
    
    # Força TLS 1.2 (Essencial para conexões modernas Google/AWS)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    try {
        $response = Invoke-RestMethod -Uri "$API_URL?mode=check_mirror_queue" -Method Get -TimeoutSec 30 -ErrorAction Stop
        Write-Host " [OK]" -ForegroundColor Green
    } catch {
        Write-Host " [FALHA]" -ForegroundColor Red
        Write-Host "DETALHE DO ERRO: $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.InnerException) {
            Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor DarkYellow
        }
        Write-Log "Falha de conexão: $($_.Exception.Message)"
        exit
    }
    
    # Verifica se há solicitações
    if (-not $response -or $response.Count -eq 0) {
        Write-Host "`nNenhuma solicitação pendente encontrada na fila." -ForegroundColor Yellow
        Write-Log "Fila vazia."
        exit
    }
    
    # Se não for array, converte
    if ($response -isnot [Array]) { $response = @($response) }
    
    $count = $response.Count
    Write-Host "$count solicitação(ões) pendente(s) encontrada(s)." -ForegroundColor Green
    Write-Log "$count solicitações encontradas."

    # 3. Processa cada solicitação
    foreach ($item in $response) {
        if ($item.error) {
             continue
        }

        $idSolicitacao = $item.id
        $usuarioModelo = $item.user_modelo
        Write-Log "Processando ID #$idSolicitacao - Usuário: $usuarioModelo"

        try {
            # 4. Consulta o AD Local
            $adUser = Get-ADUser -Identity $usuarioModelo -ErrorAction Stop
            
            # Pega os grupos (apenas nomes)
            $grupos = Get-ADPrincipalGroupMembership -Identity $usuarioModelo | Select-Object -ExpandProperty Name | Sort-Object

            # 5. Envia Resultado de volta para o Google
            $payload = @{
                action        = "update_mirror_result" # Action para o doPost
                id            = $idSolicitacao
                status        = "SUCESSO"
                grupos        = ($grupos -join ";") # Lista separada por ponto e vírgula
                msg_erro      = ""
            }

            # Converte para JSON e envia
            $jsonPayload = $payload | ConvertTo-Json -Depth 5
            
            $updateResp = Invoke-RestMethod -Uri $API_URL -Method Post -Body $jsonPayload -ContentType "application/json"
            Write-Log "Sucesso! Grupos enviados para ID #$idSolicitacao"

        } catch {
            $erroMsg = $_.Exception.Message
            Write-Log "ERRO ao consultar AD para $($usuarioModelo): $erroMsg"

            # Envia erro para o Google para atualizar status
            $payload = @{
                action   = "update_mirror_result"
                id       = $idSolicitacao
                status   = "ERRO"
                grupos   = ""
                msg_erro = $erroMsg
            }
            $jsonPayload = $payload | ConvertTo-Json
            Invoke-RestMethod -Uri $API_URL -Method Post -Body $jsonPayload -ContentType "application/json"
        }
    }

} catch {
    Write-Log "FALHA CRÍTICA NA EXECUÇÃO: $($_.Exception.Message)"
}
