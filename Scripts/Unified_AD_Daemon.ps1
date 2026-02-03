# ==============================================================================
# GERENCIAMENTO DE USUÁRIOS - SUPORTE INFRA CDS - v4.1 (DAEMON)
# ==============================================================================
# Este script unifica as funções de RESET DE SENHA e ESPELHAMENTO DE AD.
# Ele roda em loop infinito, consultando a fila unificada no Google Apps Script.
#
# REQUISITOS:
# - Módulo ActiveDirectory
# - Acesso à internet (Google Scripts)
# - Permissão de envio de e-mail (SMTP Interno)
# ==============================================================================

# --- CONFIGURAÇÃO ---
# Configurações da API (Apps Script)
$API_URL = "https://script.google.com/a/macros/luizalabs.com/s/AKfycbwcwKziwn37TfZgEJcHA_37l9aG6prf73CL-8JZ9pMgO9igU6mEC9iTrdNI1FbtI4Kr/exec"
$LogDir = "C:\ProgramData\ADResetTool\Logs"
$LoopIntervalSeconds = 5
$global:smtpServer = "smtpml.magazineluiza.intranet"

# --- PREPARAÇÃO DO AMBIENTE ---
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir "UnifiedDaemon_$(Get-Date -Format 'yyyy-MM-dd').log"

# Verifica AD
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERRO CRÍTICO: Módulo ActiveDirectory não instalado." -ForegroundColor Red
    # Em produção, descomentar: exit
}
else {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
}

# --- FUNÇÕES DE UTILIDADE ---

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-ConsoleMessage "$logEntry" $Level
    Add-Content -Path $LogFile -Value $logEntry -ErrorAction SilentlyContinue
}

function Write-ConsoleMessage {
    param([string]$Msg, [string]$Level)
    $color = switch ($Level) {
        "INFO" { "Cyan" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
        default { "Gray" }
    }
    Write-Host $Msg -ForegroundColor $color
}

function Invoke-Retry {
    param(
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$Delay = 2
    )
    $count = 0
    while ($true) {
        try {
            return & $ScriptBlock
        }
        catch {
            $count++
            if ($count -ge $MaxRetries) { throw $_ }
            Start-Sleep -Seconds $Delay
        }
    }
}

# --- FUNÇÕES DE EMAIL ---

function Send-ResetEmail {
    param($Para, $CC, $Usuario, $NomeColaborador, $NovaSenha, $Executor, $FromEmail)
    $smtpServer = "smtpml.magazineluiza.intranet"
    $assunto = "Senha Resetada - $Usuario"
    $primeiroNome = ($NomeColaborador -split ' ')[0]

    # Validação do Remetente (Analista)
    $remetente = "suporte-infra-cds@luizalabs.com"
    if ($FromEmail -and $FromEmail -match "^[\w\.-]+@([\w-]+\.)+[\w-]{2,4}$") {
        $remetente = $FromEmail
    }

    $templatePath = Join-Path $PSScriptRoot "Template_Reset_Email.html"
    
    if (Test-Path $templatePath) {
        $corpoHtml = Get-Content $templatePath -Raw -Encoding UTF8
        $corpoHtml = $corpoHtml -replace "{PRIMEIRO_NOME}", $primeiroNome
        $corpoHtml = $corpoHtml -replace "{USUARIO}", $Usuario
        $corpoHtml = $corpoHtml -replace "{NOVA_SENHA}", $NovaSenha
    }
    else {
        Write-Log "Template de email não encontrado em: $templatePath. Usando fallback simples." "WARN"
        $corpoHtml = "<body><h2>Senha Resetada</h2><p>Usuario: $Usuario</p><p>Senha: $NovaSenha</p></body>"
    }

    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $remetente
        $msg.To.Add($Para)
        if ($CC) { $CC -split ";" | ForEach-Object { if ($_) { $msg.CC.Add($_.Trim()) } } }
        $msg.Subject = $assunto
        $msg.Body = $corpoHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, 25)
        $smtp.EnableSsl = $false
        $smtp.Send($msg)
        return $true
    }
    catch {
        Write-Log "Erro envio email: $_" "ERROR"
        return $false
    }
}

# --- FUNÇÕES DE NEGÓCIO ---

function Send-UnlockEmail {
    param($Para, $CC, $Usuario, $NomeColaborador, $FromEmail)
    $primeiroNome = $NomeColaborador.Split(" ")[0]
    $assunto = "✅ Conta Desbloqueada - Suporte Infra CDs"
    $remetente = "suporte-infra-cds@luizalabs.com"
    if ($FromEmail -and $FromEmail -match "^[\w\.-]+@([\w-]+\.)+[\w-]{2,4}$") { $remetente = $FromEmail }

    $corpoHtml = "
    <div style='font-family: Arial; padding: 20px;'>
        <h2 style='color: #059669;'>Olá, $primeiroNome!</h2>
        <p>Sua conta de rede (<b>$Usuario</b>) foi desbloqueada e já está pronta para uso.</p>
        <p>Caso ainda tenha problemas de acesso, retorne o contato com o analista: $remetente.</p>
        <hr>
        <p style='font-size: 11px; color: #666;'>Atenciosamente,<br>Equipe de Infraestrutura - Magalu</p>
    </div>"
    
    try {
        $msg = New-Object System.Net.Mail.MailMessage -ArgumentList $remetente, $Para, $assunto, $corpoHtml
        if ($CC) { $CC -split ";" | ForEach-Object { if ($_) { $msg.CC.Add($_.Trim()) } } }
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, 25)
        $smtp.Send($msg)
    }
    catch { Write-Log "Erro envio email desbloqueio: $_" "ERROR" }
}

function Send-MirrorEmail {
    param($Para, $Usuario, $NomeColaborador, $ModelUser, $FromEmail)
    $primeiroNome = $NomeColaborador.Split(" ")[0]
    $assunto = "🔄 Acessos Atualizados (Espelhamento) - Suporte Infra CDs"
    $remetente = "suporte-infra-cds@luizalabs.com"
    if ($FromEmail -and $FromEmail -match "^[\w\.-]+@([\w-]+\.)+[\w-]{2,4}$") { $remetente = $FromEmail }

    $corpoHtml = "
    <div style='font-family: Arial; padding: 20px;'>
        <h2 style='color: #7c3aed;'>Olá, $primeiroNome!</h2>
        <p>Seus acessos foram atualizados com base no modelo <b>$ModelUser</b>.</p>
        <p>Os novos grupos e permissões já estão disponíveis em sua conta (<b>$Usuario</b>).</p>
        <p>Qualquer dúvida, contate o analista responsável: $remetente.</p>
        <hr>
        <p style='font-size: 11px; color: #666;'>Atenciosamente,<br>Equipe de Infraestrutura - Magalu</p>
    </div>"
    
    try {
        $msg = New-Object System.Net.Mail.MailMessage -ArgumentList $remetente, $Para, $assunto, $corpoHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, 25)
        $smtp.Send($msg)
    }
    catch { Write-Log "Erro envio email mirror: $_" "ERROR" }
}

function Send-RejectEmail {
    param($Para, $NomeSolicitante, $IdSolicitacao, $Tipo, $AnalistaEmail, $AnalistaNome)
    
    $subject = "🚫 Solicitação Reprovada: #$IdSolicitacao"
    
    $remetente = "suporte-infra-cds@luizalabs.com"
    if ($AnalistaEmail -and $AnalistaEmail -match "^[\w\.-]+@([\w-]+\.)+[\w-]{2,4}$") {
        $remetente = $AnalistaEmail
    }

    $corpoHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f3f4f6; margin: 0; padding: 0; }
        .container { max-width: 600px; margin: 20px auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; border-left: 5px solid #dc2626; }
        .header { background-color: #ffffff; padding: 20px; border-bottom: 1px solid #e5e7eb; }
        .title { color: #dc2626; font-size: 24px; font-weight: bold; margin: 0; }
        .content { padding: 30px; color: #374151; line-height: 1.6; }
        .footer { background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #6b7280; border-top: 1px solid #e5e7eb; }
        .info-box { background-color: #fee2e2; border: 1px solid #fca5a5; color: #b91c1c; padding: 15px; border-radius: 6px; margin: 20px 0; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1 class="title">Solicitação Reprovada</h1>
        </div>
        <div class="content">
            <p>Olá,</p>
            <p>Informamos que sua solicitação <strong>#$IdSolicitacao</strong> ($Tipo) foi <strong>REPROVADA</strong>.</p>
            
            <div class="info-box">
                <strong>Ação:</strong> Para entender o motivo ou ajustar sua solicitação, entre em contato com o analista responsável.
            </div>

            <p><strong>Analista Responsável:</strong> $AnalistaEmail</p>
        </div>
        <div class="footer">
            Sistema de Gerenciamento de Identidades &bull; Suporte Infra CDs
        </div>
    </div>
</body>
</html>
"@

    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $remetente
        $msg.To.Add($Para)
        $msg.Subject = $subject
        $msg.Body = $corpoHtml
        $msg.IsBodyHtml = $true
        
        $smtp = New-Object System.Net.Mail.SmtpClient("smtpml.magazineluiza.intranet", 25) # Servidor interno padrão Magalu
        $smtp.EnableSsl = $false # Internal relay usually no SSL on port 25
        $smtp.Send($msg)
        return $true
    }
    catch {
        Write-Log "Erro envio email reprovação: $_" "ERROR"
        return $false
    }
}

function Invoke-RejectUser {
    param($Task)
    $id = $Task.id_solicitacao
    $user = $Task.user_name
    Write-Log "Processando REJEIÇÃO para solicitação #$id ($user)..." "WARN"

    # Envia Email
    $aprovadorEmail = $Task.analista # ou $Task.aprovador em alguns contextos
    # Se analista não vier, tentar pegar de outra prop ou usar default
    if (-not $aprovadorEmail) { $aprovadorEmail = $Task.aprovador }

    # Solicitante email logic could be 'solicitante' or 'email_colaborador' depending on task logic
    # Reset tasks have 'solicitante' mapped now in Backend
    $solicitanteEmail = $Task.solicitante
    if (-not $solicitanteEmail) { $solicitanteEmail = $Task.email_colaborador } # Fallback for old tasks

    if ($solicitanteEmail) {
        Send-RejectEmail -Para $solicitanteEmail -IdSolicitacao $id -Tipo $Task.task_type -AnalistaEmail $aprovadorEmail
        Write-Log "Email de reprovação enviado para $solicitanteEmail (De: $aprovadorEmail)" "SUCCESS"
    }
    else {
        Write-Log "Não foi possível enviar email de reprovação: Email solicitante não encontrado." "WARN"
    }

    # Finaliza no Backend
    if ($Task.task_type -eq "MIRROR" -or $Task.task_type -eq "FETCH_GROUPS") {
        # Mirror usa a função genérica Send-Result
        Send-Result -Id $id -Type $Task.task_type -Status "REPROVADO" -Msg "Solicitação recusada pelo analista." -Task $Task
    }
    else {
        # Reset usa endpoint padrão audit
        Send-Result -Id $id -Type "RESET" -Status "REPROVADO" -Msg "REPROVADO PELO ANALISTA" -Task $Task
    }
}

function Invoke-ResetUser {
    param($Task)
    $id = $Task.id_solicitacao
    $user = $Task.user_name
    $newPassword = $Task.nova_senha
    
    Write-Log "Iniciando RESET para $user (Solicitação #$id)..." "INFO"

    try {
        # 1. Verificar usuário no AD
        $adUser = Get-ADUser -Identity $user -Properties mail, displayName, Enabled -ErrorAction Stop

        # 2. Resetar Senha
        $securePwd = ConvertTo-SecureString $newPassword -AsPlainText -Force
        Set-ADAccountPassword -Identity $user -NewPassword $securePwd -Reset -ErrorAction Stop
        
        # 3. Desbloquear e Ativar
        Unlock-ADAccount -Identity $user -ErrorAction Stop
        if (-not $adUser.Enabled) { Enable-ADAccount -Identity $user -ErrorAction Stop }
        
        # 4. Forçar troca (pwdLastSet = 0)
        Set-ADUser -Identity $user -ChangePasswordAtLogon $true -ErrorAction Stop

        # 5. Enviar Email (Personificado com E-mail do Analista)
        Send-ResetEmail -Para $Task.email_colaborador -CC $Task.email_gestor -Usuario $user -NomeColaborador $Task.nome -NovaSenha $newPassword -Executor "AUTOMACAO_DAEMON" -FromEmail $Task.analista

        # 6. Reportar Sucesso
        Send-Result -Id $id -Type "RESET" -Status "CONCLUIDO" -Msg "Senha resetada e email enviado." -Task $Task
        Write-Log "RESET #$id concluído com sucesso." "SUCCESS"
    }
    catch {
        Write-Log "Falha no RESET #$id para ${user}: $_" "ERROR"
        Send-Result -Id $id -Type "RESET" -Status "ERRO" -Msg "Falha tecnica: $_" -Task $Task
    }
}

function Invoke-UnlockUser {
    param($Task)
    $id = $Task.id_solicitacao
    $user = $Task.user_name
    
    Write-Log "Iniciando DESBLOQUEIO para $user (Solicitação #$id)..." "INFO"

    try {
        Unlock-ADAccount -Identity $user -ErrorAction Stop
        
        # Envia Email de Conclusão Personificado
        Send-UnlockEmail -Para $Task.email_colaborador -CC $Task.email_gestor -Usuario $user -NomeColaborador $Task.nome -FromEmail $Task.analista

        Send-Result -Id $id -Type "UNLOCK" -Status "CONCLUIDO" -Msg "Conta desbloqueada com sucesso." -Task $Task
        Write-Log "DESBLOQUEIO #$id concluído." "SUCCESS"
    }
    catch {
        Write-Log "Falha no DESBLOQUEIO #$id para ${user}: $_" "ERROR"
        Send-Result -Id $id -Type "UNLOCK" -Status "ERRO" -Msg "Falha tecnica: $_" -Task $Task
    }
}

function Invoke-MirrorFetch {
    param($Task)
    $id = $Task.id_solicitacao
    $modelOne = $Task.user_modelo
    
    Write-Log "Buscando grupos para espelho: $modelOne (ID: $id)" "INFO"
    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $modelOne -ErrorAction Stop | Select-Object -ExpandProperty Name
        $jsonGroups = $groups -join ";"
        
        # Payload ajustado para a função updateMirrorResult do Backend
        $payload = @{
            action = "update_mirror_result"
            id     = $id
            groups = $jsonGroups
            status = "GRUPOS_ENCONTRADOS"
        }
        
        Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $payload) -ContentType "application/json"
        Write-Log "Grupos enviados para ID #$id." "SUCCESS"
    }
    catch {
        $payload = @{
            action  = "update_mirror_result"
            id      = $id
            status  = "ERRO"
            message = "$_"
        }
        Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $payload) -ContentType "application/json"
        Write-Log "Erro ao buscar grupos de ${modelOne}: $_" "ERROR"
    }
}

function Invoke-MirrorExecute {
    param($Task)
    $id = $Task.id_solicitacao
    $model = $Task.user_modelo
    $targets = $Task.targets
    $groups = $Task.grupos
    
    Write-Log "Executando espelhamento #$id ($model -> $targets)..." "INFO"
    
    foreach ($target in $targets) {
        foreach ($grp in $groups) {
            try {
                Add-ADGroupMember -Identity $grp -Members $target -ErrorAction Stop
                Write-Log " + Adicionado $target em $grp" "INFO"
            }
            catch {
                if ($_.Exception.Message -like "*already a member*") {
                    Write-Log " . $target já está em $grp" "INFO"
                }
                else {
                    Write-Log " ! Falha ao adicionar $target em ${grp}: $_" "ERROR"
                }
            }
        }
    }
    
    # Envia Email de Conclusão Personificado para cada alvo
    foreach ($target in $targets) {
        # Busca AD para pegar o email do destino
        try {
            $destUser = Get-ADUser -Identity $target -Properties mail, displayName -ErrorAction SilentlyContinue
            if ($destUser.mail) {
                Send-MirrorEmail -Para $destUser.mail -Usuario $target -NomeColaborador $destUser.displayName -ModelUser $model -FromEmail $Task.analista
            }
        }
        catch { }
    }

    # Reportar Sucesso Final
    Send-Result -Id $id -Type "MIRROR" -Status "CONCLUIDO" -Msg "Espelhamento executado." -Task $Task
}

function Send-Result {
    param($Id, $Type, $Status, $Msg, $Task)
    
    $body = @{}
    
    if ($Type -eq "RESET") {
        # Payload completo para logar na Auditoria e Atualizar Fila
        $body = @{
            id_solicitacao    = $Id
            data_hora         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            filial            = $Task.aba
            user_name         = $Task.user_name
            nova_senha        = $Task.nova_senha
            status            = $Status
            executor          = "DAEMON_V4"
            email_colaborador = $Task.email_colaborador
            email_gestor      = $Task.email_gestor
            centro_custo      = $Task.centro_custo
            email_status      = "ENVIADO"
            aprovador         = $Task.analista
        }
    }
    elseif ($Type -eq "MIRROR") {
        # Payload para EXECUTE_MIRROR
        $body = @{
            type           = "EXECUTE_MIRROR"
            id_solicitacao = $Id
            status         = $Status
            message        = $Msg
        }
    }
    elseif ($Type -eq "UNLOCK") {
        $body = @{
            id_solicitacao    = $Id
            data_hora         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            filial            = $Task.aba
            user_name         = $Task.user_name
            nova_senha        = "CONTA_DESBLOQUEADA"
            status            = $Status
            executor          = "DAEMON_V4"
            email_colaborador = $Task.email_colaborador
            email_gestor      = $Task.email_gestor
            centro_custo      = $Task.centro_custo
            email_status      = "SOLICITACAO_PROCESSADA"
            aprovador         = $Task.analista
            type              = "UNLOCK"
            message           = $Msg
        }
    }
    
    try {
        Invoke-RestMethod -Uri $API_URL -Method Post -Body (ConvertTo-Json $body) -ContentType "application/json"
    }
    catch {
        Write-Log "Falha ao reportar status para API: $_" "ERROR"
    }
}


function Test-APIConnection {
    try {
        # Como não temos mode=ping, vamos usar get_daemon_queue mesmo, o teste é se conecta.
        $request = Invoke-WebRequest -Uri $API_URL -Method Get -ErrorAction Stop
        if ($request.StatusCode -eq 200) { return $true }
        return $false
    }
    catch {
        Write-Log "Falha de conectividade com API: $_" "ERROR"
        return $false
    }
}

# --- LOOP PRINCIPAL ---

Write-Log "Unified Daemon v4.2 Iniciado. Aguardando tarefas..." "SUCCESS"
Write-Log "API URL: $API_URL" "INFO"

if (Test-APIConnection) {
    Write-Log "Conexão com API estabelecida com sucesso." "SUCCESS"
}
else {
    Write-Log "ERRO: Não foi possível conectar na API. Verifique internet/proxy." "ERROR"
}

while ($true) {
    try {
        $url = "$API_URL" + "?mode=get_daemon_queue"
        # Debug detalhado
        # Write-Log "Consultando: $url" "DEBUG"
        
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        
        # Logar o que veio (somente se não for array vazio para não poluir)
        if ($response -is [System.Array] -and $response.Count -gt 0) {
            $tasks = $response
            Write-Log "Processando $($tasks.Count) tarefa(s) da fila." "INFO"
            
            foreach ($task in $tasks) {
                
                # VERIFICAÇÃO DE REPROVAÇÃO (NOVO v2.1.6)
                if ($task.status_aprovacao -eq "REPROVADO") {
                    Invoke-RejectUser -Task $task
                    continue # Pula para próxima tarefa
                }

                $id = $task.id_solicitacao
                Write-Log "Iniciando Atendimento Solicitação #$id ($($task.task_type))..." "INFO"
                switch ($task.task_type) {
                    "RESET" { Invoke-ResetUser -Task $task }
                    "RESET_SENHA" { Invoke-ResetUser -Task $task }
                    "DESBLOQUEIO_CONTA" { Invoke-UnlockUser -Task $task }
                    "FETCH_GROUPS" { Invoke-MirrorFetch -Task $task }
                    "MIRROR" { Invoke-MirrorExecute -Task $task }
                    "ESPELHO_USUARIO" { Invoke-MirrorExecute -Task $task }
                    default { Write-Log "Tipo de tarefa desconhecido: $($task.task_type)" "WARN" }
                }
            }
        }
        else {
            # Fila vazia - Aguardando próximo ciclo
        }
    }
    catch {
        Write-Log "Erro no loop principal: $_" "ERROR"
    }
    
    Start-Sleep -Seconds $LoopIntervalSeconds
}
