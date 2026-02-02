# ==============================================================================
# AGENTE DE SINCRONIZAÇÃO - ESPELHO DE USUÁRIOS (AD BRIDGE) - V4 (DAEMON)
# ==============================================================================
# Este script agora roda continuamente (Modo Daemon).
# Ele consulta a API a cada 10 segundos em busca de novas solicitações de grupos.

$API_URL = "https://script.google.com/macros/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec".Trim()
$LogDir = "$PSScriptRoot\..\Logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LOG_FILE = "$LogDir\Log_Espelho_AD_Sync_$(Get-Date -Format 'yyyy-MM-dd').txt"
$POLLING_INTERVAL = 10 # Segundos entre verificações

function Write-Log {
    param($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LOG_FILE -Value "[$TimeStamp] $Message"
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ERRO: Módulo ActiveDirectory não encontrado."
    exit
}
Import-Module ActiveDirectory

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "AGENTE DE SINCRONIZAÇÃO INICIADO (MODO CONTÍNUO)" -ForegroundColor White
Write-Host "Pressione CTRL+C para encerrar." -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "Agente iniciado em modo Daemon."

# --- LOOP INFINITO ---
while ($true) {
    try {
        $dataHora = Get-Date -Format "HH:mm:ss"
        Write-Host "[$dataHora] Verificando fila..." -NoNewline -ForegroundColor Gray
        
        $FullUriString = $API_URL + "?mode=check_mirror_queue"
        $UriObj = [System.Uri]$FullUriString
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $response = Invoke-RestMethod -Uri $UriObj -Method Get -ErrorAction Stop
        
        if (-not $response -or $response.Count -eq 0) {
            Write-Host " [Vazio]" -ForegroundColor DarkGray
        }
        else {
            Write-Host " [$($response.Count) pendência(s)!]" -ForegroundColor Green
            
            if ($response -isnot [Array]) { $response = @($response) }

            foreach ($item in $response) {
                $idSolicitacao = $item.id
                $tipo = $item.type 
                
                # O script de sync só processa FETCH_GROUPS (Etapa 1)
                if ($tipo -ne "FETCH_GROUPS") { continue }
                
                Write-Host "  -> Processando ID #$idSolicitacao..." -ForegroundColor Cyan
                
                try {
                    $usuarioModelo = $item.user_modelo
                    Write-Host "     Buscando grupos para: $usuarioModelo" -ForegroundColor White
                    
                    $grupos = Get-ADPrincipalGroupMembership -Identity $usuarioModelo | Select-Object -ExpandProperty Name | Sort-Object
                    
                    $payload = @{
                        action = "update_mirror_result"
                        id     = $idSolicitacao
                        type   = "FETCH_GROUPS"
                        status = "SUCESSO"
                        grupos = ($grupos -join ";")
                    }

                    $jsonPayload = $payload | ConvertTo-Json -Depth 5
                    $upResp = Invoke-RestMethod -Uri $API_URL -Method Post -Body $jsonPayload -ContentType "application/json"
                    
                    Write-Host "     Status API: $upResp" -ForegroundColor Green
                    Write-Log "Sucesso na captura de grupos para usuário: $usuarioModelo (ID #${idSolicitacao})"

                }
                catch {
                    $erroMsg = $_.Exception.Message
                    Write-Host "     ERRO: $erroMsg" -ForegroundColor Red
                    Write-Log "ERRO no ID #${idSolicitacao}: $erroMsg"
                    
                    $payload = @{
                        action    = "update_mirror_result"
                        id        = $idSolicitacao
                        type      = "FETCH_GROUPS"
                        status    = "ERRO"
                        msg_error = $erroMsg
                    }
                    $jsonPayload = $payload | ConvertTo-Json
                    Invoke-RestMethod -Uri $API_URL -Method Post -Body $jsonPayload -ContentType "application/json"
                }
            }
        }

    }
    catch {
        $err = $_.Exception.Message
        Write-Host " [FALHA NA CONEXÃO]" -ForegroundColor Red
        Write-Host "   Detalhe: $err" -ForegroundColor Yellow
        Write-Log "Falha de conexão ou erro no ciclo: $err"
    }

    # Aguarda o intervalo definido antes de tentar novamente
    Start-Sleep -Seconds $POLLING_INTERVAL
}
