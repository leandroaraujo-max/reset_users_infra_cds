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
$LoopIntervalSeconds = 10

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
    param($Para, $CC, $Usuario, $NomeColaborador, $NovaSenha, $Executor)
    $smtpServer = "smtpml.magazineluiza.intranet"
    $assunto = "Senha Resetada - $Usuario"
    $primeiroNome = ($NomeColaborador -split ' ')[0]


    $templatePath = Join-Path $PSScriptRoot "..\Templates\Template_Reset_Email.html"
    
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
        $msg.From = "suporte-infra-cds@luizalabs.com"
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

function Invoke-ResetUser {
    param($Task)
    $id = $Task.id_solicitacao
    $user = $Task.user_name
    $newPassword = $Task.nova_senha
    
    Write-Log "Iniciando RESET para $user (Solicitação #$id)..." "INFO"

    try {
        # 1. Verificar usuário no AD
        try {
            $adUser = Get-ADUser -Identity $user -Properties mail, displayName, Enabled -ErrorAction Stop
        }
        catch {
            Write-Log "Usuário $user não encontrado no AD." "WARN"
            Send-Result -Id $id -Type "RESET" -Status "ERRO" -Msg "Usuario nao encontrado no AD" -Task $Task
            return
        }

        # 2. Resetar Senha
        $securePwd = ConvertTo-SecureString $newPassword -AsPlainText -Force
        Set-ADAccountPassword -Identity $user -NewPassword $securePwd -Reset -ErrorAction Stop
        
        # 3. Desbloquear e Ativar
        Unlock-ADAccount -Identity $user -ErrorAction Stop
        if (-not $adUser.Enabled) { Enable-ADAccount -Identity $user -ErrorAction Stop }
        
        # 4. Forçar troca (pwdLastSet = 0)
        Set-ADUser -Identity $user -ChangePasswordAtLogon $true -ErrorAction Stop

        # 5. Enviar Email
        Send-ResetEmail -Para $Task.email_colaborador -CC $Task.email_gestor -Usuario $user -NomeColaborador $Task.nome -NovaSenha $newPassword -Executor "AUTOMACAO_DAEMON"

        # 6. Reportar Sucesso
        Send-Result -Id $id -Type "RESET" -Status "CONCLUIDO" -Msg "Senha resetada e email enviado." -Task $Task
        Write-Log "RESET #$id concluído com sucesso." "SUCCESS"

    }
    catch {
        Write-Log "Falha no RESET #$id para ${user}: $_" "ERROR"
        Send-Result -Id $id -Type "RESET" -Status "ERRO" -Msg "Falha tecnica: $_" -Task $Task
    }
}

function Invoke-MirrorFetch {
    param($Task)
    $id = $Task.id_solicitacao
    $modelOne = $Task.user_modelo
    
    Write-Log "Buscando grupos para espelho: $modelOne (ID: $id)" "INFO"
    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $modelOne -ErrorAction Stop | Select-Object -ExpandProperty Name
        $jsonGroups = $groups | ConvertTo-Json -Compress
        
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
            Write-Log "Recebido(s) $($tasks.Count) tarefa(s). Payload: $($tasks | ConvertTo-Json -Depth 1 -Compress)" "INFO"
            
            foreach ($task in $tasks) {
                switch ($task.task_type) {
                    "RESET" { Invoke-ResetUser -Task $task }
                    "FETCH_GROUPS" { Invoke-MirrorFetch -Task $task }
                    "MIRROR" { Invoke-MirrorExecute -Task $task }
                    default { Write-Log "Tipo de tarefa desconhecido: $($task.task_type)" "WARN" }
                }
            }
        }
        else {
            $hora = Get-Date -Format 'HH:mm:ss'
            Write-Host "[$hora] [DEBUG] Fila vazia. Nenhuma tarefa pendente e aprovada..." -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Log "Erro no loop principal: $_" "ERROR"
    }
    
    Start-Sleep -Seconds $LoopIntervalSeconds
}
