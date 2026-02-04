# ==============================================================================
# GERENCIAMENTO DE USUÁRIOS - SUPORTE INFRA CDS - v4.7 (UNIFICADO & ROBUSTO)
# ==============================================================================

# --- CONFIGURAÇÃO ---
$API_URL = "https://script.google.com/macros/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec".Trim()
$LoopIntervalSeconds = 5 
$LogDir = "C:\ProgramData\ADResetTool\Logs"
$global:smtpServer = "smtpml.magazineluiza.intranet"

# --- PREPARAÇÃO ---
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir "UnifiedDaemon_$(Get-Date -Format 'yyyy-MM-dd').log"

# Importa módulos necessários
if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) { "INFO"{"Cyan"} "WARN"{"Yellow"} "ERROR"{"Red"} "SUCCESS"{"Green"} default{"Gray"} }
    Write-Host $logEntry -ForegroundColor $color
    Add-Content -Path $LogFile -Value $logEntry -ErrorAction SilentlyContinue
}

# --- FUNÇÕES DE EMAIL ---

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
    } catch { Write-Log "Erro envio email reset: $_" "ERROR"; return $false }
}

function Send-BitlockerEmail {
    param($Para, $Hostname, $RecoveryKey, $KeyId, $ReqId)
    $linkFinalizar = "$API_URL`?action=finalizar&id=$ReqId&analista=$Para"
    if ($RecoveryKey -eq "NOT_FOUND") {
        $assunto = "⚠️ Alerta: Chave BitLocker Não Encontrada - $Hostname"
        $corpoHtml = "
        <div style='font-family: Arial; padding: 20px; background-color: #fffbeb;'>
            <h2>Chave não localizada</h2>
            <p>Estação: <b>$Hostname</b></p>
            <p>Nenhuma chave de recuperação foi encontrada no AD para esta máquina.</p><br>
            <a href='$linkFinalizar' style='background: #4b5563; color: white; padding: 10px; text-decoration: none; border-radius: 4px;'>Confirmar Leitura e Fechar Chamado</a>
        </div>"
    } else {
        $assunto = "🔐 CUSTÓDIA: Chave BitLocker - $Hostname"
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
    } catch { Write-Log "Erro envio email bitlocker: $_" "ERROR" }
}

# --- PROCESSAMENTO DE TAREFAS ---

function Invoke-TaskExecution {
    param($Task)
    
    # Normalização robusta de nomes de campos (Mapeamento flexível)
    $id = if ($Task.id_solicitacao) { $Task.id_solicitacao } else { $Task.ID }
    $user = if ($Task.user_name) { $Task.user_name } else { $Task.USER_NAME }
    $type = if ($Task.task_type) { $Task.task_type } else { $Task.TIPO_TAREFA }
    $pwd = if ($Task.nova_senha) { $Task.nova_senha } else { $Task.NOVA_SENHA }
    $nome = if ($Task.nome) { $Task.nome } else { $Task.NOME_COLABORADOR }
    $emailColab = if ($Task.email_colaborador) { $Task.email_colaborador } else { $Task.EMAIL_COLABORADOR }
    $emailGestor = if ($Task.email_gestor) { $Task.email_gestor } else { $Task.EMAIL_GESTOR }
    $analista = if ($Task.analista) { $Task.analista } else { $Task.ANALISTA_EMAIL }

    if (-not $id) { return }

    try {
        Write-Log "Processando ID #$id ($type) para: $user" "INFO"

        switch ($type) {
            "RESET" {
                if (-not $pwd) { throw "Senha não fornecida pela API" }
                $securePwd = ConvertTo-SecureString $pwd -AsPlainText -Force
                Set-ADAccountPassword -Identity $user -NewPassword $securePwd -Reset -ErrorAction Stop
                Unlock-ADAccount -Identity $user -ErrorAction Stop
                Set-ADUser -Identity $user -ChangePasswordAtLogon $true -ErrorAction Stop
                
                # Envio de E-mail (Se houver dados)
                if ($emailColab) {
                    Send-ResetEmail -Para $emailColab -CC $emailGestor -Usuario $user -NomeColaborador $nome -NovaSenha $pwd -FromEmail $analista
                    $msgFinal = "Reset executado e e-mail enviado."
                } else {
                    $msgFinal = "Reset executado (E-mail não enviado - destinatário ausente)."
                }
                $payload = @{ action="report_status"; id=$id; status="CONCLUIDO"; message=$msgFinal }
            }
            "DESBLOQUEIO_CONTA" {
                Unlock-ADAccount -Identity $user -ErrorAction Stop
                $payload = @{ action="report_status"; id=$id; status="CONCLUIDO"; message="Conta desbloqueada com sucesso." }
            }
            "BITLOCKER" {
                # v1.5.5: Busca Inteligente (Global por ID ou Local por Hostname)
                $prefix = $req.password_id_prefix
                $recovery = $null
                $hostnameResolved = $user

                if ($prefix) {
                    Write-Log "Iniciando busca GLOBAL BitLocker por ID: $prefix"
                    # v1.5.6: Filtro LDAP resiliente (igual ao dsa.msc). 
                    # Usamos '*' no início e fim para garantir o match independente de '{' no esquema.
                    $recovery = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation' -and msFVE-RecoveryPasswordID -like '*$prefix*'" -Properties msFVE-RecoveryPassword, msFVE-RecoveryPasswordID | Select-Object -First 1
                    
                    if ($recovery) {
                        # Extrai o nome do computador do DN do objeto pai
                        $parentDN = $recovery.DistinguishedName -replace '^CN=[^,]+,',''
                        $compParent = Get-ADComputer -Identity $parentDN -ErrorAction SilentlyContinue
                        if ($compParent) { $hostnameResolved = $compParent.Name }
                        Write-Log "Chave localizada via ID Global. Hostname detectado: $hostnameResolved"
                    }
                }

                # Fallback ou Busca por Hostname se não encontrou por ID ou ID não foi fornecido
                if (-not $recovery) {
                    Write-Log "Buscando computador por Hostname: $user"
                    $comp = Get-ADComputer -Filter "Name -eq '$user'" -ErrorAction SilentlyContinue
                    if (-not $comp) {
                        if ($analista) { Send-BitlockerEmail -Para $analista -Hostname $user -RecoveryKey "NOT_FOUND" -ReqId $id }
                        throw "Computador/ID '$user' não encontrado no AD."
                    }

                    Write-Log "Buscando chaves BitLocker na estação: $($comp.Name)"
                    $allRecoveries = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $comp.DistinguishedName -Properties msFVE-RecoveryPassword, msFVE-RecoveryPasswordID
                    
                    if ($prefix) {
                        # Filtra pelo prefixo dentro da máquina específica (caso a busca global tenha falhado ou retornado múltiplos)
                        $recovery = $allRecoveries | Where-Object { 
                            $cleanId = $_."msFVE-RecoveryPasswordID".Replace("{", "").Replace("}", "")
                            $cleanId.StartsWith($prefix) 
                        } | Sort-Object whenCreated -Descending | Select-Object -First 1
                    } else {
                        # Sem prefixo, usamos a chave mais recente vinculada à máquina
                        $recovery = $allRecoveries | Sort-Object whenCreated -Descending | Select-Object -First 1
                    }
                    $hostnameResolved = $comp.Name
                }

                if (-not $recovery) {
                    if ($analista) { Send-BitlockerEmail -Para $analista -Hostname $hostnameResolved -RecoveryKey "NOT_FOUND" -ReqId $id }
                    throw "Nenhuma chave BitLocker encontrada para '$hostnameResolved' (Filtro ID: $prefix)."
                }
                
                $realKey = $recovery."msFVE-RecoveryPassword"
                $fullId = $recovery."msFVE-RecoveryPasswordID"

                if ($analista) {
                    Send-BitlockerEmail -Para $analista -Hostname $hostnameResolved -RecoveryKey $realKey -KeyId $fullId -ReqId $id
                }
                
                $payload = @{ 
                    action = "update_bitlocker_result"
                    id = $id
                    status = "SUCESSO"
                    recoveryKey = $realKey
                    recoveryKeyId = $fullId
                    hostname = $hostnameResolved
                }
            }
            default {
                Write-Log "Tipo de tarefa não suportado: $type" "WARN"
                return
            }
        }

        # Notifica Sucesso
        Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $payload) -ContentType "application/json"
        Write-Log "ID #$id finalizado e reportado." "SUCCESS"
    } catch {
        Write-Log "FALHA no ID #${id}: $($_.Exception.Message)" "ERROR"
        $payload = @{ action="report_status"; id=$id; status="ERRO"; message=$_.Exception.Message }
        Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $payload) -ContentType "application/json"
    }
}

# --- LOOP PRINCIPAL ---

Write-Log "Daemon v4.7 Iniciado - Unificado & Robusto (Varredura 5s)" "SUCCESS"

while ($true) {
    try {
        # Busca a fila (Trata sempre como Array)
        $response = Invoke-RestMethod -Uri "$API_URL`?mode=get_daemon_queue" -Method Get -ErrorAction Stop
        $tasks = @($response) | Where-Object { $_.id_solicitacao -or $_.ID }

        if ($tasks.Count -gt 0) {
            Write-Log "Varredura: $($tasks.Count) solicitações reais encontradas." "WARN"
            
            foreach ($task in $tasks) {
                # Se na planilha estiver REPROVADO, apenas limpa
                if ($task.status_aprovacao -eq "REPROVADO") {
                    $id = if ($task.id_solicitacao) { $task.id_solicitacao } else { $task.ID }
                    Write-Log "Limpando ID #$id (Reprovado manualmente)." "WARN"
                    $p = @{ action="report_status"; id=$id; status="REPROVADO" }
                    Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $p) -ContentType "application/json"
                    continue
                }

                # EXECUTA TAREFA
                Invoke-TaskExecution -Task $task
                Start-Sleep -Milliseconds 300
            }
        }
    }
    catch {
        Write-Log "Erro de conexão ou processamento: $_" "ERROR"
    }

    Start-Sleep -Seconds $LoopIntervalSeconds
}