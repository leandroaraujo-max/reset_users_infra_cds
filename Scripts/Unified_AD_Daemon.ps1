# ==============================================================================
# IDENTITY MANAGER - SUPORTE INFRA CDS - v6.0 (WMS PRINT QUEUES SUPPORT)
# ==============================================================================

# --- CONFIGURAÇÃO ---
# --- CONFIGURAÇÃO MULTI-AMBIENTE ---
$Environments = @(
    @{
        Name   = "PROD"
        Prefix = "[PROD] "
        Url    = "https://script.google.com/macros/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec".Trim()
    },
    @{
        Name   = "STAGING"
        Prefix = "[STAGING] "
        Url    = "https://script.google.com/a/macros/luizalabs.com/s/AKfycbziGqnkYDXS4oI2nDqPlrk2epJjN8boCVcjSdZ-kgZWqIPgBfprE9vCTIRCDggllC_aKg/exec".Trim()
    }
)

$LoopIntervalSeconds = 5 
$LogDir = "C:\ProgramData\ADResetTool\Logs"
$global:smtpServer = "smtpml.magazineluiza.intranet"

# --- PREPARAÇÃO ---
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir "UnifiedDaemon_$(Get-Date -Format 'yyyy-MM-dd').log"

# Garante o módulo de AD
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
}

# Configuração UTF-8 para console e saída (silencia erro se não suportado)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [string]$Prefix = "")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Prefix$Message"
    $color = switch ($Level) { "INFO" { "Cyan" } "WARN" { "Yellow" } "ERROR" { "Red" } "SUCCESS" { "Green" } default { "Gray" } }
    Write-Host $logEntry -ForegroundColor $color
    $logEntry | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

# Função para enviar requisições POST com encoding UTF-8
function Send-ApiRequest {
    param([string]$Uri, [hashtable]$Body)
    $json = ConvertTo-Json $Body -Depth 5
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    return Invoke-RestMethod -Uri $Uri -Method Post -Body $bytes -ContentType "application/json; charset=utf-8"
}

# --- FUNÇÕES DE EMAIL (ESTILIZADAS MAGALU) ---

function Send-ResetEmail {
    param($Para, $CC, $Usuario, $NomeColaborador, $NovaSenha, $FromEmail)
    $primeiroNome = ($NomeColaborador -split ' ')[0]
    $remetente = if ($FromEmail -and $FromEmail -match "^[\w\.-]+@([\w-]+\.)+[\w-]{2,4}$") { $FromEmail } else { "suporte-infra-cds@luizalabs.com" }
    
    $corpoHtml = "
    <div style='font-family: Arial; padding: 20px;'>
        <h2 style='color: #0284c7;'>Olá, $primeiroNome!</h2>
        <p>Sua senha de rede foi resetada conforme solicitado.</p>
        <div style='background: #f1f5f9; padding: 15px; border-radius: 8px; font-family: monospace;'>
            <b>Usuário:</b> $Usuario<br>
            <b>Nova Senha:</b> $NovaSenha
        </div>
        <p><i>Nota: Você deverá alterar esta senha no primeiro login.</i></p>
    </div>"
    
    try {
        $msg = New-Object System.Net.Mail.MailMessage -ArgumentList $remetente, $Para, "Senha Resetada - $Usuario", $corpoHtml
        if ($CC) { $CC -split ";" | ForEach-Object { if ($_) { $msg.CC.Add($_.Trim()) } } }
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($global:smtpServer, 25)
        $smtp.Send($msg)
        return $true
    }
    catch { Write-Log "Erro SMTP Reset: $($_.Exception.Message)" "ERROR"; return $false }
}

function Send-BitlockerEmail {
    param($Para, $Hostname, $RecoveryKey, $KeyId, $ReqId, $ApiUrl)
    $linkFinalizar = "$ApiUrl`?action=finalizar&id=$ReqId&analista=$Para"
    if ($RecoveryKey -eq "NOT_FOUND") {
        $assunto = "⚠️ Alerta: Chave BitLocker Não Encontrada - $Hostname"
        $corpoHtml = "
        <div style='font-family: Arial; padding: 20px; background-color: #fffbeb;'>
            <h2>Chave não localizada</h2>
            <p>Estação: <b>$Hostname</b></p>
            <p>Nenhuma chave de recuperação foi encontrada no AD para esta máquina.</p><br>
            <a href='$linkFinalizar' style='background: #4b5563; color: white; padding: 10px; text-decoration: none; border-radius: 4px;'>Confirmar Leitura e Fechar Chamado</a>
        </div>"
    }
    else {
        $assunto = "🔐 RECUPERAÇÃO: Chave BitLocker - $Hostname"
        $corpoHtml = "
        <div style='font-family: Arial; padding: 25px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;'>
            <h2 style='color: #1e40af; margin-top: 0;'>Chave de Recuperação Encontrada</h2>
            <p style='color: #64748b;'>Abaixo estão os detalhes para a estação: <b style='color: #0f172a;'>$Hostname</b></p>
            
            <div style='background: #ffffff; padding: 20px; border: 2px solid #3b82f6; border-radius: 8px; margin: 20px 0;'>
                <label style='display: block; font-size: 10px; color: #3b82f6; font-weight: bold; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 5px;'>Recovery Key (Numérica)</label>
                <div style='font-family: monospace; font-size: 24px; color: #1e293b; font-weight: bold; letter-spacing: 2px; text-align: center;'>
                    $RecoveryKey
                </div>
            </div>

            <div style='background: #f1f5f9; padding: 12px; border-radius: 6px; font-size: 13px;'>
                <b style='color: #475569;'>ID da Senha:</b> <span style='font-family: monospace;'>$KeyId</span><br>
                <b style='color: #475569;'>Solicitação:</b> #$ReqId
            </div>

            <div style='margin-top: 25px; text-align: center;'>
                <a href='$linkFinalizar' style='display: inline-block; background: #2563eb; color: white; padding: 12px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 14px;'>FINALIZAR ATENDIMENTO</a>
                <p style='font-size: 11px; color: #94a3b8; margin-top: 15px;'>Ao clicar acima, você confirma que entregou a chave ao colaborador e registrou o atendimento.</p>
            </div>
        </div>"
    }
    try {
        $msg = New-Object System.Net.Mail.MailMessage -ArgumentList "suporte-infra-cds@luizalabs.com", $Para, $assunto, $corpoHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($global:smtpServer, 25)
        $smtp.Send($msg)
        return $true
    }
    catch { Write-Log "Erro SMTP BitLocker: $($_.Exception.Message)" "ERROR"; return $false }
}

function Send-WMSEmail {
    param($Para, $QueueName, $ReqId, $StatusMsg, $Sucesso)
    
    $subjectStatus = if ($Sucesso) { "CONCLUÍDO" } else { "ERRO" }
    $colorStatus = if ($Sucesso) { "#16a34a" } else { "#dc2626" }
    $assunto = "🖨️ WMS: Limpeza de Fila $subjectStatus - $QueueName"
    
    $corpoHtml = "
    <div style='font-family: Arial; padding: 25px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;'>
        <h2 style='color: #1e40af; margin-top: 0;'>Notificação WMS</h2>
        <p style='color: #334155; font-size: 14px;'>A solicitação de limpeza de fila de impressão foi executada pelo Agente Automático.</p>
        
        <div style='background: white; padding: 15px; border-left: 4px solid $colorStatus; margin: 20px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1);'>
            <p style='margin: 5px 0;'><b style='color: #475569;'>ID Solicitação:</b> #$ReqId</p>
            <p style='margin: 5px 0;'><b style='color: #475569;'>Fila:</b> $QueueName</p>
            <p style='margin: 5px 0;'><b style='color: #475569;'>Status:</b> <span style='color: $colorStatus; font-weight: bold;'>$subjectStatus</span></p>
        </div>
        
        <div style='background: #f1f5f9; padding: 12px; border-radius: 6px; font-size: 13px; color: #334155;'>
            <b>Detalhes da Execução:</b><br>
            $StatusMsg
        </div>
        
        <p style='font-size: 11px; color: #94a3b8; margin-top: 20px; text-align: center;'>Este é um e-mail automático. Não responda.</p>
    </div>"

    try {
        $msg = New-Object System.Net.Mail.MailMessage -ArgumentList "suporte-infra-cds@luizalabs.com", $Para, $assunto, $corpoHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($global:smtpServer, 25)
        $smtp.Send($msg)
        return $true
    }
    catch { Write-Log "Erro SMTP WMS: $($_.Exception.Message)" "ERROR"; return $false }
}

# --- FUNÇÕES WMS (v1.7.0) ---

function Get-WmsClusterConfig {
    param([string]$ApiUrl)
    
    try {
        $response = Invoke-RestMethod -Uri "$ApiUrl`?mode=get_wms_cluster" -Method Get -ErrorAction Stop
        
        # Valida se o retorno é um array de strings válidas
        if ($response -and $response.Count -gt 0) {
            $validServers = @()
            foreach ($server in $response) {
                if ($server -and $server.ToString().Trim().Length -gt 0) {
                    $validServers += $server.ToString().Trim()
                }
            }
            
            if ($validServers.Count -gt 0) {
                Write-Log "Cluster WMS: $($validServers.Count) servidores obtidos" "INFO"
                return $validServers
            }
        }
        
        Write-Log "Cluster WMS: Lista vazia ou inválida retornada" "WARN"
        return @()
    }
    catch {
        Write-Log "Erro ao obter cluster WMS: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Sync-PrintQueues {
    param([string]$ApiUrl)
    
    
    # Obtém lista de servidores WMS da API (Failover Coluna C - Planilha Servidores)
    $wmsServers = Get-WmsClusterConfig -Url $ApiUrl
    
    if (-not $wmsServers -or $wmsServers.Count -eq 0) {
        Write-Log "Usando fallback local para Cluster WMS" "WARN"
        $wmsServers = @(
            "ml-ibm-wms-01.magazineluiza.intranet",
            "ml-ibm-wms-02.magazineluiza.intranet"
        )
    }
    
    $printers = $null
    $lastError = $null
    
    foreach ($server in $wmsServers) {
        try {
            Write-Log "Tentando conectar ao servidor WMS: $server..." "INFO"
            
            # Testa conexão primeiro
            if (-not (Test-Connection -ComputerName $server -Count 1 -Quiet)) {
                Write-Log "Servidor $server não responde ao ping" "WARN"
                continue
            }
            
            # Lista todas as filas de impressão do servidor
            $printers = Get-Printer -ComputerName $server -ErrorAction Stop | 
            Select-Object -ExpandProperty Name
            
            if ($printers -and $printers.Count -gt 0) {
                Write-Log "Conectado a $server - Encontradas $($printers.Count) filas" "SUCCESS"
                break
            }
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "Falha em $server : $lastError" "WARN"
        }
    }
    
    if (-not $printers -or $printers.Count -eq 0) {
        throw "Nao foi possivel obter impressoras de nenhum servidor WMS. Ultimo erro: $lastError"
    }
    
    # Ordena alfabeticamente
    $printers = $printers | Sort-Object
    Write-Log "Impressoras ordenadas alfabeticamente" "INFO"
    
    # Envia para o Backend com encoding UTF-8
    try {
        $payload = @{
            action   = "update_printer_cache"
            printers = $printers
        } | ConvertTo-Json -Depth 3
        
        # Converte para UTF-8
        $payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        
        $result = Invoke-RestMethod -Uri $ApiUrl -Method Post -Body $payloadBytes -ContentType "application/json; charset=utf-8" -ErrorAction Stop
        Write-Log "Cache de impressoras atualizado: $result" "SUCCESS"
    }
    catch {
        throw "Erro ao enviar cache para API: $($_.Exception.Message)"
    }
}

# --- PROCESSAMENTO ---

function Invoke-TaskExecution {
    param($Task, $EnvContext)
    
    $ApiUrl = $EnvContext.Url
    $LogPfx = $EnvContext.Prefix

    # Mapeamento Dinâmico v5.2
    $id = ($Task.id_solicitacao, $Task.id, $Task.ID | Where-Object { $_ } | Select-Object -First 1)
    $user = ($Task.user_name, $Task.USER_NAME, $Task.usuario | Where-Object { $_ } | Select-Object -First 1)
    $type = ($Task.task_type, $Task.TIPO_TAREFA, $Task.type | Where-Object { $_ } | Select-Object -First 1)
    $clearPwd = ($Task.nova_senha, $Task.NOVA_SENHA, $Task.senha | Where-Object { $_ } | Select-Object -First 1)
    $nome = ($Task.nome, $Task.NOME_COLABORADOR, $Task.NOME | Where-Object { $_ } | Select-Object -First 1)
    $emailColab = ($Task.email_colaborador, $Task.EMAIL_COLABORADOR, $Task.email_colab | Where-Object { $_ } | Select-Object -First 1)
    $emailGestor = ($Task.email_gestor, $Task.EMAIL_GESTOR | Where-Object { $_ } | Select-Object -First 1)
    $analista = ($Task.analista, $Task.ANALISTA_EMAIL, $Task.solicitante | Where-Object { $_ } | Select-Object -First 1)
    $filial = ($Task.filial, $Task.FILIAL | Where-Object { $_ } | Select-Object -First 1)
    $prefix = $Task.password_id_prefix

    if (-not $id) { return }

    try {
        Write-Log "Executando Tarefa ID #$id ($type) para $user" "INFO" $LogPfx
        $payload = @{ action = "report_status"; id = $id; status = "ERRO"; message = "Iniciado" }

        switch ($type) {
            "FETCH_GROUPS" {
                $userModelo = ($Task.user_modelo, $Task.USER_MODELO | Where-Object { $_ } | Select-Object -First 1)
                if (-not $userModelo) { throw "Usuário Modelo não informado." }

                Write-Log "Buscando grupos para modelo: $userModelo" "INFO" $LogPfx
                
                # Se for matrícula (apenas números), busca o username
                if ($userModelo -match '^\d+$') {
                    Write-Log "Matrícula detectada ($userModelo), buscando username..." "INFO" $LogPfx
                    $adUser = Get-ADUser -Filter "EmployeeID -eq '$userModelo'" -Properties SamAccountName -ErrorAction Stop
                    if (-not $adUser) {
                        throw "Usuário com matrícula $userModelo não encontrado no AD"
                    }
                    $userModelo = $adUser.SamAccountName
                    Write-Log "Username encontrado: $userModelo" "INFO" $LogPfx
                }
                
                $groups = (Get-ADPrincipalGroupMembership -Identity $userModelo -ErrorAction Stop | Select-Object -ExpandProperty Name) -join ";"
                
                # Para FETCH_GROUPS, usa action update_mirror_result
                $payload = @{ 
                    action    = "update_mirror_result"
                    type      = "FETCH_GROUPS"
                    id        = $id
                    requestId = $id
                    status    = "CONCLUIDO"
                    msg_error = ""
                    grupos    = $groups
                }
            }

            "RESET" {
                if (-not $clearPwd) { throw "Senha ausente para Reset." }
                $securePwd = ConvertTo-SecureString $clearPwd -AsPlainText -Force
                Set-ADAccountPassword -Identity $user -NewPassword $securePwd -Reset -ErrorAction Stop
                Unlock-ADAccount -Identity $user -ErrorAction Stop
                Set-ADUser -Identity $user -ChangePasswordAtLogon $true -ErrorAction Stop
                
                $msgFinal = "Reset executado."
                if ($emailColab) {
                    if (Send-ResetEmail -Para $emailColab -CC $emailGestor -Usuario $user -NomeColaborador $nome -NovaSenha $clearPwd -FromEmail $analista) {
                        $msgFinal += " E-mail enviado para o colaborador."
                    }
                }
                $payload.status = "CONCLUIDO"; $payload.message = $msgFinal
            }

            "BITLOCKER" {
                # v5.7: Busca Inteligente Resiliente via RDN
                $recovery = $null
                $hostnameResolved = $user

                if ($prefix) {
                    Write-Log "Busca GLOBAL BitLocker (v5.7 - RDN Filter) por ID: $prefix" "INFO" $LogPfx
                    # Buscamos todos os objetos da classe e filtramos pelo DN/Name para evitar erro de atributos sintéticos.
                    $recovery = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -Properties msFVE-RecoveryPassword | 
                    Where-Object { $_.DistinguishedName -like "*$prefix*" -or $_.Name -like "*$prefix*" } | Select-Object -First 1
                    
                    if ($recovery) {
                        if ($recovery.DistinguishedName -match "CN=([^,]+),CN=([^,]+),") {
                            $hostnameResolved = $matches[2]
                        }
                        else {
                            $parentDN = $recovery.DistinguishedName -replace '^CN=[^,]+,', ''
                            $compParent = Get-ADComputer -Identity $parentDN -ErrorAction SilentlyContinue
                            if ($compParent) { $hostnameResolved = $compParent.Name }
                        }
                        Write-Log "Chave localizada no DN (ID: $prefix). Hostname: $hostnameResolved" "SUCCESS" $LogPfx
                    }
                }

                if (-not $recovery) {
                    Write-Log "Fallback: Buscando computador por Hostname: $user" "INFO" $LogPfx
                    $comp = Get-ADComputer -Filter "Name -eq '$user'" -ErrorAction SilentlyContinue
                    if ($comp) {
                        $allRecoveries = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $comp.DistinguishedName -Properties msFVE-RecoveryPassword
                        $recovery = $allRecoveries | Sort-Object whenCreated -Descending | Select-Object -First 1
                        $hostnameResolved = $comp.Name
                    }
                }

                if ($recovery) {
                    $realKey = $recovery."msFVE-RecoveryPassword"
                    $fullId = if ($recovery.Name -match "{([A-F0-9-]+)}") { $matches[1] } else { $recovery.Name }
                    
                    Send-BitlockerEmail -Para $analista -Hostname $hostnameResolved -RecoveryKey $realKey -KeyId $fullId -ReqId $id -ApiUrl $ApiUrl
                    
                    $payload = @{ 
                        action        = "update_bitlocker_result"
                        id            = $id
                        status        = "SUCESSO"
                        recoveryKey   = $realKey
                        recoveryKeyId = $fullId
                        hostname      = $hostnameResolved
                        filial        = $filial
                    }
                }
                else {
                    if ($analista) { Send-BitlockerEmail -Para $analista -Hostname $hostnameResolved -RecoveryKey "NOT_FOUND" -ReqId $id -ApiUrl $ApiUrl }
                    throw "Chave não localizada no AD para $hostnameResolved."
                }
            }

            "DESBLOQUEIO_CONTA" {
                Unlock-ADAccount -Identity $user -ErrorAction Stop
                $payload.status = "CONCLUIDO"; $payload.message = "Conta desbloqueada com sucesso."
            }

            "WMS_PRINT_CLEAN" {
                # v6.0: Limpeza de Fila de Impressão WMS em Cluster
                $queueName = ($Task.queue_name, $Task.QUEUE_NAME, $Task.fila | Where-Object { $_ } | Select-Object -First 1)
                
                if (-not $queueName) {
                    throw "Nome da fila de impressao nao especificado."
                }

                Write-Log "Iniciando limpeza da fila: $queueName" "INFO" $LogPfx
                
                # Obtém lista de servidores do cluster
                $clusterServers = Get-WmsClusterConfig -ApiUrl $ApiUrl
                
                if ($clusterServers.Count -eq 0) {
                    throw "Cluster WMS vazio ou inacessivel."
                }

                $successCount = 0
                $errorCount = 0
                $errors = @()

                foreach ($server in $clusterServers) {
                    try {
                        Write-Log "Limpando fila '$queueName' no servidor $server..." "INFO" $LogPfx
                        
                        # Remove todos os trabalhos da fila
                        Get-PrintJob -ComputerName $server -PrinterName $queueName -ErrorAction Stop | 
                        Remove-PrintJob -ErrorAction Stop
                        
                        $successCount++
                        Write-Log "Fila limpa com sucesso em $server" "SUCCESS" $LogPfx
                    }
                    catch {
                        $errorCount++
                        $errorMsg = "Erro em ${server}: $($_.Exception.Message)"
                        $errors += $errorMsg
                        Write-Log $errorMsg "ERROR" $LogPfx
                    }
                }

                if ($successCount -gt 0) {
                    $msgFinal = "Fila '$queueName' limpa em $successCount/$($clusterServers.Count) servidores."
                    if ($errorCount -gt 0) {
                        $msgFinal += " Erros: $errorCount. Detalhes: $($errors -join '; ')"
                    }
                    $payload.status = "CONCLUIDO"
                    $payload.message = $msgFinal
                    
                    if ($analista) {
                        Send-WMSEmail -Para $analista -QueueName $queueName -ReqId $id -StatusMsg $msgFinal -Sucesso $true
                    }
                }
                else {
                    $msgErro = "Falha ao limpar fila em todos os servidores. Erros: $($errors -join '; ')"
                    if ($analista) {
                        Send-WMSEmail -Para $analista -QueueName $queueName -ReqId $id -StatusMsg $msgErro -Sucesso $false
                    }
                    throw $msgErro
                }
            }
            
            default {
                throw "Tipo de tarefa não suportado: $type"
            }
        }

        # Update Final
        Send-ApiRequest -Uri $ApiUrl -Body $payload
        Write-Log "ID #$id OK!" "SUCCESS" $LogPfx

    }
    catch {
        Write-Log "FALHA ID #$id : $($_.Exception.Message)" "ERROR" $LogPfx
        $errPayload = @{ action = "report_status"; id = $id; status = "ERRO"; message = $_.Exception.Message }
        Send-ApiRequest -Uri $ApiUrl -Body $errPayload
    }
}

# --- LOOP PRINCIPAL ---
Write-Log "Daemon v6.0 ATIVO - MULTI-AMBIENTE (PROD + STAGING) + WMS SUPPORT" "SUCCESS"

# Controle de sync de impressoras WMS (a cada 2 horas)
$lastPrinterSync = [DateTime]::MinValue
$printerSyncIntervalMinutes = 120

while ($true) {
    # Sync de impressoras WMS periodicamente
    $timeSinceSync = (Get-Date) - $lastPrinterSync
    if ($timeSinceSync.TotalMinutes -ge $printerSyncIntervalMinutes) {
        Write-Log "Iniciando sincronização de impressoras WMS..." "INFO"
        foreach ($env in $Environments) {
            try {
                Sync-PrintQueues -ApiUrl $env.Url
                Write-Log "Sync WMS concluído para $($env.Name)" "SUCCESS" $env.Prefix
            }
            catch {
                Write-Log "Erro no sync WMS para $($env.Name): $($_.Exception.Message)" "WARN" $env.Prefix
            }
        }
        $lastPrinterSync = Get-Date
    }

    foreach ($env in $Environments) {
        try {
            # Ocasionalmente loga heartbeat
            # Write-Log "Consultando..." "INFO" $env.Prefix
            
            $response = Invoke-RestMethod -Uri "$($env.Url)?mode=get_daemon_queue" -Method Get -ErrorAction Stop
            $tasks = @($response) | Where-Object { $_.ID -or $_.id_solicitacao }
            
            foreach ($t in $tasks) { 
                if ($t.status_aprovacao -eq "REPROVADO") {
                    $rid = ($t.id_solicitacao, $t.id, $t.ID | Where-Object { $_ } | Select-Object -First 1)
                    Write-Log "Limpando ID #$rid (Status: REPROVADO)" "WARN" $env.Prefix
                    $p = @{ action = "report_status"; id = $rid; status = "REPROVADO" }
                    Send-ApiRequest -Uri $env.Url -Body $p
                    continue
                }
                Invoke-TaskExecution -Task $t -EnvContext $env
                Start-Sleep -Milliseconds 200 
            }
        }
        catch { 
            # Erros de conexão silenciosos ou log de erro simples para não spammar
            # Write-Log "Erro API ($($env.Name)): $($_.Exception.Message)" "ERROR" $env.Prefix
        }
    }
    Start-Sleep -Seconds $LoopIntervalSeconds
}
