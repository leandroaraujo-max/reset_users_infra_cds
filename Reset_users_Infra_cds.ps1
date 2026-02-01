# --- VERSÃO 1.0.1 (RELEASE OFICIAL - SISTEMA COMPLETO) ---
Write-Host "--- VERSÃO 1.0.1 CARREGADA ---" -ForegroundColor Cyan

# ==========================================================================
# CONFIGURAÇÃO CENTRAL
# ==========================================================================
# COLE AQUI A SUA URL DA IMPLANTAÇÃO (EXEC):
$API_URL = "https://script.google.com/a/macros/luizalabs.com/s/AKfycbyramIj-2zInesvtZbkUFA5BiEp-__bdGsa7yOv_G5xys8EZqG-c2zvkxi4JpjsZkjb/exec"

# --- PRÉ-REQUISITOS DE SEGURANÇA ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Verifica módulo AD (Ignorar se for teste fora do ambiente de domínio, mas ideal ter instalado)
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "AVISO: Módulo ActiveDirectory não encontrado. O script rodará apenas em modo de visualização." -ForegroundColor Yellow
    # Em produção, descomente a linha abaixo para bloquear sem RSAT:
    # [System.Windows.Forms.MessageBox]::Show("Instale o RSAT.", "Erro", "OK", "Error"); exit 
}
else {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
}

# Carrega bibliotecas visuais
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================================================
# FUNÇÕES UTILITÁRIAS (LOGS E RETRY) - DEFINIDAS NO INÍCIO
# ==========================================================================
# --- CONFIGURAÇÃO DE LOGS ---
$logDir = "C:\ProgramData\ADResetTool\Logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force -Path $logDir | Out-Null }
$logFile = Join-Path $logDir "Log_$(Get-Date -Format 'yyyy-MM-dd').txt"

function global:Write-Console($msg) {
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $logMsg = "[$timestamp] $msg"
    
    # Escreve no Arquivo
    try {
        Add-Content -Path $logFile -Value $logMsg -ErrorAction SilentlyContinue
    }
    catch {}

    # Escreve na GUI (se existir)
    if ($global:txtConsole -and !$global:txtConsole.IsDisposed) {
        $global:txtConsole.AppendText("$logMsg`r`n")
        $global:txtConsole.SelectionStart = $global:txtConsole.Text.Length
        $global:txtConsole.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

Write-Host "Carregando funções utilitárias..." -ForegroundColor Gray

function global:Invoke-Retry {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 2,
        [int]$DelaySeconds = 2,
        [string]$ErrorMessage = "Operação falhou"
    )
    
    $retryCount = 0
    $completed = $false
    $lastError = $null
    
    while (-not $completed) {
        try {
            # Executa o bloco de código
            $result = & $ScriptBlock
            $completed = $true
            return $result
        }
        catch {
            $lastError = $_
            $retryCount++
            
            if ($retryCount -ge $MaxRetries) {
                Write-Console " > ERRO FATAL: $ErrorMessage após $retryCount tentativas. Detalhe: $($lastError.Exception.Message)"
                throw $lastError
            }
            else {
                Write-Console " > AVISO: Tentativa $retryCount de $MaxRetries falhou. Retentando em $DelaySeconds segundos..."
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
}

# --- CONFIGURAÇÃO DA INTERFACE (GUI) ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "AD Reset Tool v1.0.0 - Gestão Centralizada"
$mainForm.Size = New-Object System.Drawing.Size(1200, 850)
$mainForm.StartPosition = "CenterScreen"
$mainForm.BackColor = [System.Drawing.Color]::WhiteSmoke
$mainForm.WindowState = [System.Windows.Forms.FormWindowState]::Maximized

# Barra de Status (Criada antes para usar no log de carregamento)
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Iniciando..."
$statusStrip.Items.Add($statusLabel) | Out-Null
$mainForm.Controls.Add($statusStrip)

# ==========================================================================
# SEÇÃO 1: SELEÇÃO DO ANALISTA (DINÂMICO)
# ==========================================================================
$grpSeguranca = New-Object System.Windows.Forms.GroupBox
$grpSeguranca.Location = New-Object System.Drawing.Point(20, 10)
$grpSeguranca.Size = New-Object System.Drawing.Size(300, 80)
$grpSeguranca.Text = "Identificação do Analista"
$grpSeguranca.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($grpSeguranca)

$lblAnalista = New-Object System.Windows.Forms.Label
$lblAnalista.Location = New-Object System.Drawing.Point(10, 25); $lblAnalista.Size = New-Object System.Drawing.Size(280, 20)
$lblAnalista.Text = "Selecione seu nome:"
$lblAnalista.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$grpSeguranca.Controls.Add($lblAnalista)

$cbAnalista = New-Object System.Windows.Forms.ComboBox
$cbAnalista.Location = New-Object System.Drawing.Point(10, 45); $cbAnalista.Size = New-Object System.Drawing.Size(280, 25)
$cbAnalista.DropDownStyle = "DropDownList" # Bloqueia digitação por padrão
$cbAnalista.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$grpSeguranca.Controls.Add($cbAnalista)

# --- CARREGAMENTO DA LISTA DE ANALISTAS (API) ---
# Executa a busca antes de liberar a interface para o usuário
try {
    $statusLabel.Text = "Buscando lista de analistas na nuvem..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    # URL com o novo parâmetro mode=get_analysts
    $urlAnalistas = $API_URL + "?mode=get_analysts"
    
    # Chama API com Retry
    $analistasRemotos = Invoke-Retry -ScriptBlock { 
        Invoke-RestMethod -Uri $urlAnalistas -Method Get -ErrorAction Stop 
    } -ErrorMessage "Falha ao buscar analistas"

    if ($analistasRemotos -is [System.Array] -and $analistasRemotos.Count -gt 0) {
        # Adiciona a opção TODOS no início da lista
        $cbAnalista.Items.Add("TODOS") | Out-Null
        $cbAnalista.Items.AddRange($analistasRemotos)
        Write-Host "Sucesso: $($analistasRemotos.Count) analistas carregados." -ForegroundColor Green
    }
    else {
        throw "Lista vazia ou formato inválido."
    }
}
catch {
    # Fallback: Se der erro, permite digitar o nome manualmente
    Write-Host "Erro ao carregar analistas: $_" -ForegroundColor Red
    $statusLabel.Text = "Erro ao carregar lista. Modo manual ativado."
    $cbAnalista.Items.Add("TODOS") | Out-Null
    $cbAnalista.Items.Add("--- DIGITE SEU NOME ---")
    $cbAnalista.DropDownStyle = "DropDown" # Libera digitação
}
finally {
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
}


# ==========================================================================
# SEÇÃO 2: BUSCA (TOPO MEIO)
# ==========================================================================
$grpBusca = New-Object System.Windows.Forms.GroupBox
$grpBusca.Location = New-Object System.Drawing.Point(340, 10)
$grpBusca.Size = New-Object System.Drawing.Size(350, 80)
$grpBusca.Text = "Busca de Demandas"
$mainForm.Controls.Add($grpBusca)

$lblFilial = New-Object System.Windows.Forms.Label
$lblFilial.Location = New-Object System.Drawing.Point(10, 20); $lblFilial.Size = New-Object System.Drawing.Size(200, 20)
$lblFilial.Text = "Filial (Nº) ou *:"
$grpBusca.Controls.Add($lblFilial)

$txtFilial = New-Object System.Windows.Forms.TextBox
$txtFilial.Location = New-Object System.Drawing.Point(10, 40); $txtFilial.Size = New-Object System.Drawing.Size(100, 25)
$txtFilial.Text = "*" # Default
$grpBusca.Controls.Add($txtFilial)

$btnBuscar = New-Object System.Windows.Forms.Button
$btnBuscar.Location = New-Object System.Drawing.Point(120, 38); $btnBuscar.Size = New-Object System.Drawing.Size(200, 28)
$btnBuscar.Text = "Carregar Demandas"
$btnBuscar.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $btnBuscar.ForeColor = [System.Drawing.Color]::White
$grpBusca.Controls.Add($btnBuscar)

# ==========================================================================
# SEÇÃO 3: BUSCA INDIVIDUAL
# ==========================================================================
$grpIndiv = New-Object System.Windows.Forms.GroupBox
$grpIndiv.Location = New-Object System.Drawing.Point(710, 10)
$grpIndiv.Size = New-Object System.Drawing.Size(300, 80)
$grpIndiv.Text = "Busca Direta (Emergência)"
$mainForm.Controls.Add($grpIndiv)

$lblUserUnico = New-Object System.Windows.Forms.Label
$lblUserUnico.Location = New-Object System.Drawing.Point(10, 20); $lblUserUnico.Size = New-Object System.Drawing.Size(150, 20)
$lblUserUnico.Text = "Usuário (Login):"
$grpIndiv.Controls.Add($lblUserUnico)

$txtUserUnico = New-Object System.Windows.Forms.TextBox
$txtUserUnico.Location = New-Object System.Drawing.Point(10, 40); $txtUserUnico.Size = New-Object System.Drawing.Size(120, 25)
$grpIndiv.Controls.Add($txtUserUnico)

$btnBuscarUnico = New-Object System.Windows.Forms.Button
$btnBuscarUnico.Location = New-Object System.Drawing.Point(140, 38); $btnBuscarUnico.Size = New-Object System.Drawing.Size(140, 28)
$btnBuscarUnico.Text = "Resetar Avulso"
$btnBuscarUnico.BackColor = [System.Drawing.Color]::Purple; $btnBuscarUnico.ForeColor = [System.Drawing.Color]::White
$grpIndiv.Controls.Add($btnBuscarUnico)


# ==========================================================================
# OPÇÕES DE CONFIGURAÇÃO (TAB)
# ==========================================================================
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 100)
$tabControl.Size = New-Object System.Drawing.Size(1140, 100)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$tabOpcoes = New-Object System.Windows.Forms.TabPage
$tabOpcoes.Text = "Opções de Conta"
$tabOpcoes.BackColor = [System.Drawing.Color]::WhiteSmoke
$tabControl.Controls.Add($tabOpcoes)

$chkEmail = New-Object System.Windows.Forms.CheckBox
$chkEmail.Location = New-Object System.Drawing.Point(20, 15); $chkEmail.Size = New-Object System.Drawing.Size(280, 20)
$chkEmail.Text = "Enviar E-mails (Gestor/Colaborador)"
$chkEmail.Checked = $true
$tabOpcoes.Controls.Add($chkEmail)

$chkUnlock = New-Object System.Windows.Forms.CheckBox
$chkUnlock.Location = New-Object System.Drawing.Point(20, 40); $chkUnlock.Size = New-Object System.Drawing.Size(280, 20)
$chkUnlock.Text = "Desbloquear Conta (Unlock)"
$chkUnlock.Checked = $true
$tabOpcoes.Controls.Add($chkUnlock)

$chkPwdPolicy = New-Object System.Windows.Forms.CheckBox
$chkPwdPolicy.Location = New-Object System.Drawing.Point(350, 15); $chkPwdPolicy.Size = New-Object System.Drawing.Size(250, 20)
$chkPwdPolicy.Text = "Forçar Expiração Mensal (GPO)"
$chkPwdPolicy.Checked = $true
$tabOpcoes.Controls.Add($chkPwdPolicy)

$chkAtivar = New-Object System.Windows.Forms.CheckBox
$chkAtivar.Location = New-Object System.Drawing.Point(350, 40); $chkAtivar.Size = New-Object System.Drawing.Size(250, 20)
$chkAtivar.Text = "Ativar conta (se desativada)"
$tabOpcoes.Controls.Add($chkAtivar)

$mainForm.Controls.Add($tabControl)

# ==========================================================================
# GRID E CONSOLE
# ==========================================================================
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 240)
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGridView.Size = New-Object System.Drawing.Size(1140, 310)
$dataGridView.AutoSizeColumnsMode = "Fill"; $dataGridView.AllowUserToAddRows = $false; $dataGridView.ReadOnly = $true
$dataGridView.BackgroundColor = [System.Drawing.Color]::White

$dataGridView.Columns.Clear()
$dataGridView.Columns.Add("Filial", "Filial / Origem") | Out-Null
$dataGridView.Columns.Add("User", "Usuário AD") | Out-Null
$dataGridView.Columns.Add("Nome", "Nome Funcionário") | Out-Null
$dataGridView.Columns.Add("Analista", "Atribuído Para") | Out-Null 
$dataGridView.Columns.Add("Email", "E-mail Colab") | Out-Null
$dataGridView.Columns.Add("EmailGestor", "E-mail Gestor") | Out-Null
$dataGridView.Columns.Add("CC", "Centro de Custo") | Out-Null 
$dataGridView.Columns.Add("SenhaNova", "Nova Senha") | Out-Null
$dataGridView.Columns.Add("Audit", "Já Resetado?") | Out-Null
$dataGridView.Columns.Add("Lideranca", "Liderança (Status Email)") | Out-Null
$dataGridView.Columns.Add("Status", "Status Reset") | Out-Null
$dataGridView.Columns["Email"].Visible = $false
$dataGridView.Columns["EmailGestor"].Visible = $false
$dataGridView.Columns["Audit"].Visible = $false
$dataGridView.Columns["CC"].Visible = $true 
$mainForm.Controls.Add($dataGridView)

$grpConsole = New-Object System.Windows.Forms.GroupBox
$grpConsole.Text = "Log de Execução"
$grpConsole.Location = New-Object System.Drawing.Point(20, 560)
$grpConsole.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpConsole.Size = New-Object System.Drawing.Size(1140, 120)
$mainForm.Controls.Add($grpConsole)

$txtConsole = New-Object System.Windows.Forms.TextBox
$txtConsole.Multiline = $true
$txtConsole.ScrollBars = "Vertical"
$txtConsole.ReadOnly = $true
$txtConsole.BackColor = [System.Drawing.Color]::Black
$txtConsole.ForeColor = [System.Drawing.Color]::Lime
$txtConsole.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtConsole.Dock = "Fill"
$grpConsole.Controls.Add($txtConsole)

$txtConsole.Dock = "Fill"
$grpConsole.Controls.Add($txtConsole)
$global:txtConsole = $txtConsole # Expose to function

# ==========================================================================
# FUNÇÃO DE ENVIO DE EMAIL VIA SMTP
# ==========================================================================
function Send-ResetEmail {
    param(
        [string]$Para,
        [string]$CC,
        [string]$Usuario,
        [string]$NomeColaborador,
        [string]$NovaSenha,
        [string]$Executor
    )
    
    $smtpServer = "smtpml.magazineluiza.intranet"
    $smtpPort = 25
    $remetente = "suporte-infra-cds@luizalabs.com"
    $assunto = "Senha Resetada - $Usuario"
    
    $primeiroNome = ($NomeColaborador -split ' ')[0]
    $dataHora = (Get-Date).ToString("dd/MM/yyyy HH:mm")
    
    # Template HTML com identidade visual Magalu
    $corpoHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body style="margin: 0; padding: 0; font-family: Arial, 'Helvetica Neue', sans-serif;">
    <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
        <!-- Header Magalu -->
        <div style="background: linear-gradient(135deg, #0086FF 0%, #0066CC 100%); padding: 25px; text-align: center;">
            <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: bold; letter-spacing: 1px;">Magalu</h1>
            <p style="color: rgba(255,255,255,0.9); margin: 5px 0 0 0; font-size: 12px;">Suporte Infra CDs</p>
        </div>
        
        <!-- Rainbow Strip -->
        <div style="height: 6px; background: linear-gradient(90deg, #0086FF 0%, #00C853 20%, #FFEB3B 40%, #FF9800 60%, #E91E63 80%, #9C27B0 100%);"></div>
        
        <!-- Content -->
        <div style="padding: 30px; background-color: #ffffff;">
            <h2 style="color: #0066CC; margin: 0 0 20px 0; font-size: 22px;">Senha Resetada com Sucesso</h2>
            
            <p style="color: #333; font-size: 15px; line-height: 1.6; margin-bottom: 20px;">
                Ola <strong>$primeiroNome</strong>,<br><br>
                A senha de acesso a rede foi resetada conforme solicitacao.
            </p>
            
            <div style="background-color: #f5f7fa; border-left: 4px solid #0086FF; padding: 20px; margin: 20px 0; border-radius: 0 8px 8px 0;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <tr>
                        <td style="padding: 8px 0; color: #666; width: 130px;"><strong>Usuario:</strong></td>
                        <td style="padding: 8px 0; color: #333; font-family: 'Courier New', monospace; background: #e8e8e8; padding: 8px 12px; border-radius: 4px;">$Usuario</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #666;"><strong>Nome:</strong></td>
                        <td style="padding: 8px 0; color: #333;">$NomeColaborador</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #666;"><strong>Nova Senha:</strong></td>
                        <td style="padding: 8px 0;">
                            <span style="background: linear-gradient(135deg, #0086FF, #0066CC); color: #fff; padding: 8px 16px; border-radius: 4px; font-family: 'Courier New', monospace; font-weight: bold; letter-spacing: 1px;">$NovaSenha</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #666;"><strong>Data/Hora:</strong></td>
                        <td style="padding: 8px 0; color: #333;">$dataHora</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #666;"><strong>Executado por:</strong></td>
                        <td style="padding: 8px 0; color: #333;">$Executor</td>
                    </tr>
                </table>
            </div>
            
            <div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 6px; padding: 15px; margin: 20px 0;">
                <p style="margin: 0; color: #856404; font-size: 13px;">
                    <strong>Importante:</strong> Voce sera solicitado a trocar a senha no proximo login. Crie uma senha segura e nao compartilhe com ninguem.
                </p>
            </div>
        </div>
        
        <!-- Footer -->
        <div style="background-color: #f5f5f5; padding: 20px; text-align: center; border-top: 1px solid #e0e0e0;">
            <p style="margin: 0; color: #888; font-size: 11px;">
                Este e um email automatico do sistema de Reset de Senhas.<br>
                Suporte Infra CDs - Leandro Araujo
            </p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $remetente
        $msg.To.Add($Para)
        if ($CC -and $CC -ne "") {
            $CC -split ";" | ForEach-Object { 
                if ($_.Trim() -ne "") { $msg.CC.Add($_.Trim()) }
            }
        }
        $msg.Subject = $assunto
        $msg.Body = $corpoHtml
        $msg.IsBodyHtml = $true
        $msg.BodyEncoding = [System.Text.Encoding]::UTF8
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtp.EnableSsl = $false
        $smtp.Timeout = 2000 # 2 segundos de timeout
        
        # Envia com Retry
        Invoke-Retry -ScriptBlock { 
            $smtp.Send($msg) 
        } -ErrorMessage "Falha no envio SMTP"
        
        Write-Console " > Email enviado para: $Para"
        return $true
    }
    catch {
        Write-Console " > ERRO ao enviar email: $_"
        return $false
    }
}

# ==========================================================================
# FUNÇÃO DE EMAIL PARA SOLICITAÇÃO DE CRIAÇÃO DE CONTA NO TURIA
# ==========================================================================
function Send-TuriaRequestEmail {
    param(
        [string]$Para,
        [string]$CC,
        [string]$NomeColaborador,
        [string]$Usuario,
        [string]$CentroCusto
    )
    
    $smtpServer = "smtpml.magazineluiza.intranet"
    $smtpPort = 25
    $remetente = "suporte-infra-cds@luizalabs.com"
    $assunto = "Acao Necessaria - Criar Conta de Rede no Turia para $NomeColaborador"
    
    $primeiroNome = ($NomeColaborador -split ' ')[0]
    
    # Template HTML com passo a passo didático
    $corpoHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body style="margin: 0; padding: 0; font-family: Arial, 'Helvetica Neue', sans-serif;">
    <div style="max-width: 650px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
        <!-- Header Magalu -->
        <div style="background: linear-gradient(135deg, #0086FF 0%, #0066CC 100%); padding: 25px; text-align: center;">
            <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: bold; letter-spacing: 1px;">Magalu</h1>
            <p style="color: rgba(255,255,255,0.9); margin: 5px 0 0 0; font-size: 12px;">Suporte Infra CDs</p>
        </div>
        
        <!-- Rainbow Strip -->
        <div style="height: 6px; background: linear-gradient(90deg, #0086FF 0%, #00C853 20%, #FFEB3B 40%, #FF9800 60%, #E91E63 80%, #9C27B0 100%);"></div>
        
        <!-- Content -->
        <div style="padding: 30px; background-color: #ffffff;">
            <h2 style="color: #dc3545; margin: 0 0 20px 0; font-size: 22px;">Conta de Rede Nao Encontrada</h2>
            
            <p style="color: #333; font-size: 15px; line-height: 1.6; margin-bottom: 20px;">
                Ola <strong>$primeiroNome</strong>,<br><br>
                Recebemos uma solicitacao de reset de senha para o colaborador abaixo, porem identificamos que <strong>ele ainda nao possui uma conta de rede (login de computador)</strong>.
            </p>
            
            <div style="background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 15px; margin: 20px 0; border-radius: 0 8px 8px 0;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <tr>
                        <td style="padding: 8px 0; color: #721c24; width: 130px;"><strong>Colaborador:</strong></td>
                        <td style="padding: 8px 0; color: #721c24;">$NomeColaborador</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #721c24;"><strong>Usuario Tentado:</strong></td>
                        <td style="padding: 8px 0; color: #721c24; font-family: 'Courier New', monospace;">$Usuario</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px 0; color: #721c24;"><strong>Centro de Custo:</strong></td>
                        <td style="padding: 8px 0; color: #721c24;">$CentroCusto</td>
                    </tr>
                </table>
            </div>
            
            <div style="background-color: #d4edda; border: 1px solid #28a745; border-radius: 8px; padding: 20px; margin: 25px 0;">
                <h3 style="color: #155724; margin: 0 0 15px 0; font-size: 18px;">Como Solicitar a Criacao da Conta</h3>
                <p style="color: #155724; font-size: 14px; margin-bottom: 15px;">
                    Siga o passo a passo abaixo para criar a conta de rede no sistema <strong>Turia</strong>:
                </p>
                
                <div style="background: #fff; border-radius: 6px; padding: 15px;">
                    <!-- Passo 1 -->
                    <div style="display: flex; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #ddd;">
                        <div style="background: #0086FF; color: #fff; width: 30px; height: 30px; border-radius: 50%; display: inline-block; text-align: center; line-height: 30px; font-weight: bold; margin-right: 15px; flex-shrink: 0;">1</div>
                        <div>
                            <strong style="color: #333;">Acesse o sistema Turia</strong><br>
                            <span style="color: #666; font-size: 13px;">Entre no site: <a href="https://iam.corp.luizalabs.com/" style="color: #0086FF;">https://iam.corp.luizalabs.com/</a></span><br>
                            <span style="color: #666; font-size: 13px;">Em "Minhas Requisicoes", clique no botao <strong style="color: #0086FF;">+ Criar</strong></span>
                        </div>
                    </div>
                    
                    <!-- Passo 2 -->
                    <div style="display: flex; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #ddd;">
                        <div style="background: #0086FF; color: #fff; width: 30px; height: 30px; border-radius: 50%; display: inline-block; text-align: center; line-height: 30px; font-weight: bold; margin-right: 15px; flex-shrink: 0;">2</div>
                        <div>
                            <strong style="color: #333;">Selecione "Acesso as aplicacoes"</strong><br>
                            <span style="color: #666; font-size: 13px;">Na tela "Voce deseja...", clique em <strong style="color: #0086FF;">Solicitar acessos</strong></span>
                        </div>
                    </div>
                    
                    <!-- Passo 3 -->
                    <div style="display: flex; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #ddd;">
                        <div style="background: #0086FF; color: #fff; width: 30px; height: 30px; border-radius: 50%; display: inline-block; text-align: center; line-height: 30px; font-weight: bold; margin-right: 15px; flex-shrink: 0;">3</div>
                        <div>
                            <strong style="color: #333;">Escolha para quem e o acesso</strong><br>
                            <span style="color: #666; font-size: 13px;">Selecione <strong style="color: #0086FF;">"Para outro colaborador"</strong> e clique em Proximo</span><br>
                            <span style="color: #666; font-size: 13px;">Busque pelo nome: <strong>$NomeColaborador</strong></span>
                        </div>
                    </div>
                    
                    <!-- Passo 4 -->
                    <div style="display: flex; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #ddd;">
                        <div style="background: #28a745; color: #fff; width: 30px; height: 30px; border-radius: 50%; display: inline-block; text-align: center; line-height: 30px; font-weight: bold; margin-right: 15px; flex-shrink: 0;">4</div>
                        <div>
                            <strong style="color: #333;">Use um colaborador como referencia (RECOMENDADO)</strong><br>
                            <span style="color: #666; font-size: 13px;">Selecione <strong style="color: #28a745;">"Usar colaborador como referencia"</strong></span><br>
                            <span style="color: #666; font-size: 13px;">Busque um colega que ja tenha os mesmos acessos necessarios</span><br>
                            <span style="color: #999; font-size: 12px; font-style: italic;">Isso copia automaticamente os acessos do colega</span>
                        </div>
                    </div>
                    
                    <!-- Passo 5 -->
                    <div style="display: flex;">
                        <div style="background: #0086FF; color: #fff; width: 30px; height: 30px; border-radius: 50%; display: inline-block; text-align: center; line-height: 30px; font-weight: bold; margin-right: 15px; flex-shrink: 0;">5</div>
                        <div>
                            <strong style="color: #333;">Selecione o sistema "AD Magalu"</strong><br>
                            <span style="color: #666; font-size: 13px;">Na lista de sistemas, marque a opcao <strong style="color: #0086FF;">"AD Magalu"</strong></span><br>
                            <span style="color: #666; font-size: 13px;">Este e o sistema responsavel pelo login no computador</span><br>
                            <span style="color: #666; font-size: 13px;">Finalize a solicitacao e aguarde a aprovacao</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 6px; padding: 15px; margin: 20px 0;">
                <p style="margin: 0; color: #856404; font-size: 13px;">
                    <strong>Importante:</strong> Apos a aprovacao da solicitacao, o colaborador recebera um e-mail com as credenciais de acesso. Esse processo pode levar ate 24 horas.
                </p>
            </div>
            
            <div style="text-align: center; margin-top: 25px;">
                <a href="https://iam.corp.luizalabs.com/" style="display: inline-block; background: linear-gradient(135deg, #0086FF, #0066CC); color: #fff; padding: 12px 30px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 14px;">Acessar o Turia Agora</a>
            </div>
        </div>
        
        <!-- Footer -->
        <div style="background-color: #f5f5f5; padding: 20px; text-align: center; border-top: 1px solid #e0e0e0;">
            <p style="margin: 0; color: #888; font-size: 11px;">
                Este e um email automatico do sistema de Reset de Senhas.<br>
                Suporte Infra CDs - Leandro Araujo
            </p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $remetente
        $msg.To.Add($Para)
        if ($CC -and $CC -ne "") {
            $CC -split ";" | ForEach-Object { 
                if ($_.Trim() -ne "") { $msg.CC.Add($_.Trim()) }
            }
        }
        $msg.Subject = $assunto
        $msg.Body = $corpoHtml
        $msg.IsBodyHtml = $true
        $msg.BodyEncoding = [System.Text.Encoding]::UTF8
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtp.EnableSsl = $false
        $smtp.Timeout = 2000 # 2 segundos de timeout
        
        # Envia com Retry
        Invoke-Retry -ScriptBlock { 
            $smtp.Send($msg) 
        } -ErrorMessage "Falha no envio SMTP (Turia)"
        
        Write-Console " > Email de solicitacao Turia enviado para: $Para"
        return $true
    }
    catch {
        Write-Console " > ERRO ao enviar email Turia: $_"
        return $false
    }
}

# ==========================================================================
# RODAPÉ
# ==========================================================================
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnReset.Location = New-Object System.Drawing.Point(20, 690); $btnReset.Size = New-Object System.Drawing.Size(220, 40)
$btnReset.Text = "EXECUTAR PROCESSO"; $btnReset.BackColor = [System.Drawing.Color]::ForestGreen; $btnReset.ForeColor = [System.Drawing.Color]::White; $btnReset.Enabled = $false
$btnReset.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($btnReset)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnExport.Location = New-Object System.Drawing.Point(260, 690); $btnExport.Size = New-Object System.Drawing.Size(160, 40)
$btnExport.Text = "Exportar Pendentes"; $btnExport.BackColor = [System.Drawing.Color]::Orange; $btnExport.Enabled = $false
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($btnExport)

$lblAssinatura = New-Object System.Windows.Forms.Label
$lblAssinatura.Text = "AD Reset Tool v1.0.0"
$lblAssinatura.AutoSize = $true
$lblAssinatura.ForeColor = [System.Drawing.Color]::Gray
$lblAssinatura.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$lblAssinatura.Location = New-Object System.Drawing.Point(1000, 700)
$mainForm.Controls.Add($lblAssinatura)

# ==========================================================================
# LÓGICA 1: CARGA (GET) + FILTRAGEM
# ==========================================================================
function CarregarDados($urlSufix) {
    if ($cbAnalista.Text.Length -lt 3) {
        [System.Windows.Forms.MessageBox]::Show("Selecione ou digite seu nome corretamente.", "Atenção", "OK", "Warning")
        return
    }
    
    $analistaSelecionado = $cbAnalista.Text
    
    # Parâmetro mode=api garante o retorno de fila de usuários
    $separador = if ($urlSufix -match "\?") { "&" } else { "?" }
    $urlFinal = $API_URL + $urlSufix + $separador + "mode=api"

    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $dataGridView.Rows.Clear()
    $statusLabel.Text = "Consultando API..."
    Write-Console "Iniciando consulta à API..."

    try {
        $apiData = Invoke-Retry -ScriptBlock { 
            Invoke-RestMethod -Uri $urlFinal -Method Get 
        } -ErrorMessage "Falha ao consultar API de Demandas"
        
        Write-Console "Resposta recebida com sucesso."
        
        $lista = @($apiData)

        # Checagem de Erro da API
        if ($lista.Count -eq 1 -and $null -ne $lista[0].error) {
            [System.Windows.Forms.MessageBox]::Show("API: $($lista[0].error)", "Informação", "OK", "Information")
            $statusLabel.Text = "API: $($lista[0].error)"
            Write-Console "MSG API: $($lista[0].error)"
            return
        }

        # Filtragem e População
        $contadorOw = 0     # Meus
        $contadorLivre = 0  # Livres
        $ignorados = 0      # Outros
        
        foreach ($user in $lista) {
            if ($null -eq $user.user_name) { continue }
            
            $atribuido = if ($user.analista) { $user.analista } else { "N/A" } 
            if ([string]::IsNullOrWhiteSpace($user.analista)) { $atribuido = "N/A" }
            
            # Filtro Lógico: Se busquei por filial e tem dono diferente, ignoro (exceto se selecionou TODOS)
            if ($urlSufix -match "filial=" -and $analistaSelecionado -ne "TODOS") {
                if ($atribuido -ne $analistaSelecionado -and $atribuido -ne "N/A") {
                    $ignorados++
                    continue 
                }
            }
            
            # Gera senha complexa que atende requisitos do domínio
            # Formato: ##Magazine## + 4 números aleatórios
            $rnd = Get-Random -Minimum 1000 -Maximum 9999
            $senha = "##Magazine##${rnd}" # Ex: ##Magazine##1234 (15 chars)
            $origem = if ($user.aba) { $user.aba } else { "N/A" }
            $email = if ($user.email_colaborador) { $user.email_colaborador } else { "" }
            $emailGestor = if ($user.email_gestor) { $user.email_gestor } else { "" }
            $jaResetado = if ($user.ja_resetado) { $true } else { $false }
            $cc = if ($user.centro_custo) { $user.centro_custo } else { "" }
            
            # Formata Nome (Primeira Letra Maiúscula)
            $nomeFormatado = $user.nome
            if ($nomeFormatado) {
                $nomeFormatado = (Get-Culture).TextInfo.ToTitleCase($nomeFormatado.ToLower())
            }
            
            # Prepara texto da coluna Liderança
            $liderancaTexto = if ($emailGestor) { $emailGestor } else { "Sem gestor" }

            $rowId = $dataGridView.Rows.Add($origem, $user.user_name, $nomeFormatado, $atribuido, $email, $emailGestor, $cc, $senha, $jaResetado, $liderancaTexto, "Pendente")
            
            if ($atribuido -eq "N/A") {
                $dataGridView.Rows[$rowId].Cells["Analista"].Value = "LIVRE (N/A)"
                $dataGridView.Rows[$rowId].Cells["Analista"].Style.ForeColor = "Blue"
                $contadorLivre++
            }
            else {
                $dataGridView.Rows[$rowId].Cells["Analista"].Style.ForeColor = "Green"
                $contadorOw++
            }
            
            if ($jaResetado) {
                $dataGridView.Rows[$rowId].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
                $dataGridView.Rows[$rowId].Cells["Status"].Value = "JÁ AUDITADO"
            }
        }

        $totalVisivel = $contadorOw + $contadorLivre

        if ($totalVisivel -gt 0) {
            $btnReset.Enabled = $true; $btnExport.Enabled = $true
            $statusLabel.Text = "$totalVisivel registros (Meus: $contadorOw | Livres: $contadorLivre)."
            Write-Console "Sucesso. Encontrados: Meus=$contadorOw, Livres=$contadorLivre."
            
            if ($analistaSelecionado -eq "TODOS") {
                [System.Windows.Forms.MessageBox]::Show("Busca Concluída!`n`nTotal de registros: $totalVisivel`n(Atribuídos: $contadorOw | Livres: $contadorLivre)", "Sucesso", "OK", "Information")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Busca Concluída!`n`nAtribuídos a VOCÊ: $contadorOw`nDisponíveis: $contadorLivre", "Sucesso", "OK", "Information")
            }
        }
        else {
            if ($ignorados -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Nenhum registro para VOCÊ ou LIVRE.`nMas existem $ignorados registros de OUTROS analistas.", "Aviso", "OK", "Warning")
                Write-Console "Aviso: Pendências de outros analistas ($ignorados)."
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Nenhum registro encontrado.", "Vazio", "OK", "Information")
                Write-Console "Nenhum resultado retornado."
            }
            $statusLabel.Text = "Nenhum resultado."
        }
    }
    catch { 
        [System.Windows.Forms.MessageBox]::Show("Erro Consultar API: $_", "Erro", "OK", "Error")
        $statusLabel.Text = "Erro conexão."
        Write-Console "ERRO FATAL: $_"
    }
    finally { $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default }
}

$btnBuscar.Add_Click({
        if ([string]::IsNullOrEmpty($txtFilial.Text)) { return }
        $filialEnc = if ($txtFilial.Text -eq "*") { "*" } else { [Uri]::EscapeDataString($txtFilial.Text) }
        CarregarDados "?filial=$filialEnc"
    })

$btnBuscarUnico.Add_Click({
        if ([string]::IsNullOrEmpty($txtUserUnico.Text)) { return }
        CarregarDados "?usuario=$($txtUserUnico.Text)"
    })

# ==========================================================================
# LÓGICA 2: RESET (POST)
# ==========================================================================
$btnReset.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("Confirma a execução dos resets listados?", "Confirmação", "YesNo", "Warning") -eq 'No') { return }
    
        $urlBase = $API_URL
        $executor = $env:USERNAME
    
        $enviarEmail = $chkEmail.Checked
        $ativarConta = $chkAtivar.Checked
        $desbloquearConta = $chkUnlock.Checked
        $forcarExpiracao = $chkPwdPolicy.Checked
    
        Write-Console "Iniciando processamento em lote..."

        foreach ($row in $dataGridView.Rows) {
            $userAD = $row.Cells["User"].Value
            if ($row.Cells["Audit"].Value -eq $true) {
                $row.Cells["Status"].Value = "IGNORADO (Já feito)"; continue
            }
            if ($row.Cells["Status"].Value -match "SUCESSO") { continue } 

            $senha = $row.Cells["SenhaNova"].Value; $origem = $row.Cells["Filial"].Value
            $nome = $row.Cells["Nome"].Value; $emailColab = $row.Cells["Email"].Value; $cc = $row.Cells["CC"].Value
            $emailGestor = $row.Cells["EmailGestor"].Value
        
            $statusLabel.Text = "Processando: $userAD..."
            Write-Console "Processando: $userAD..."
            [System.Windows.Forms.Application]::DoEvents()

            try {
                # Valida se o usuário existe no AD antes de tentar resetar
                $adUser = Get-ADUser -Identity $userAD -ErrorAction Stop
                
                if ($null -eq $adUser) {
                    throw "Usuário não encontrado no Active Directory."
                }
                
                $sec = ConvertTo-SecureString $senha -AsPlainText -Force
            
                # --- COMANDOS AD ---
                Set-ADAccountPassword -Identity $userAD -NewPassword $sec -Reset -ErrorAction Stop
                Set-ADUser -Identity $userAD -ChangePasswordAtLogon $true -ErrorAction Stop
            
                if ($ativarConta) { Enable-ADAccount -Identity $userAD -ErrorAction SilentlyContinue; Write-Console " > Conta ativada." }
                if ($desbloquearConta) { Unlock-ADAccount -Identity $userAD -ErrorAction SilentlyContinue; Write-Console " > Conta desbloqueada." }
                if ($forcarExpiracao) { Set-ADUser -Identity $userAD -PasswordNeverExpires $false -ErrorAction SilentlyContinue }
                # -------------------

                $row.Cells["Status"].Value = "SUCESSO"
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
                Write-Console " > SUCESSO! Senha resetada."
            
                $emailParaApi = if ($enviarEmail) { $emailColab } else { "" }
                $emailGestorParaApi = if ($enviarEmail) { $emailGestor } else { "" }
                $emailStatusLog = "Desabilitado"
                
                # ENVIO DE EMAIL VIA SMTP (PowerShell)
                if ($enviarEmail) {
                    $destinatarioEmail = if ($emailColab) { $emailColab } else { "" }
                    $ccEmail = if ($emailGestor) { $emailGestor } else { "" }
                    
                    if ($destinatarioEmail) {
                        $emailEnviado = Send-ResetEmail -Para $destinatarioEmail -CC $ccEmail -Usuario $userAD -NomeColaborador $nome -NovaSenha $senha -Executor $executor
                        
                        if ($emailEnviado) {
                            $emailStatusLog = "Enviado para $destinatarioEmail"
                            if ($ccEmail) {
                                $row.Cells["Lideranca"].Value = "$ccEmail [OK]"
                                $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Green
                            }
                            else {
                                $row.Cells["Lideranca"].Value = "Sem gestor [OK]"
                                $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Green
                            }
                        }
                        else {
                            $emailStatusLog = "Erro no envio"
                            $row.Cells["Lideranca"].Value = "Erro envio [X]"
                            $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Red
                        }
                    }
                    else {
                        $emailStatusLog = "Sem email cadastrado"
                        $row.Cells["Lideranca"].Value = "Sem email [X]"
                        $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Orange
                        Write-Console " > AVISO: Colaborador sem email cadastrado."
                    }
                }
                else {
                    $row.Cells["Lideranca"].Value = "Envio desabilitado"
                    $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Gray
                }
                
                # ENVIO DO LOG PARA API (APÓS email para registrar status correto)
                $p = @{
                    data_hora         = (Get-Date).ToString("dd/MM/yyyy HH:mm");
                    filial            = $origem;
                    user_name         = $userAD;
                    nome_colaborador  = $nome;
                    email_colaborador = $emailParaApi;
                    email_gestor      = $emailGestorParaApi;
                    centro_custo      = $cc; 
                    nova_senha        = $senha;
                    status            = "SUCESSO";
                    executor          = $executor;
                    email_status      = $emailStatusLog
                } | ConvertTo-Json
            
                try { 
                    Invoke-Retry -ScriptBlock {
                        Invoke-RestMethod -Uri $urlBase -Method Post -Body $p -ContentType "application/json" 
                    } -ErrorMessage "Falha ao enviar log API"
                    Write-Console " > Log enviado para API."
                }
                catch { 
                    Write-Console " > ERRO API (Log): $_" 
                }

            }
            catch {
                $err = $_.Exception.Message
                
                # Identifica o tipo de erro para melhor diagnóstico
                if ($err -match "não é possível localizar|cannot find|not found") {
                    Write-Console " > Usuario '$userAD' nao encontrado no AD."
                    
                    # Define destinatário: colaborador OU gestores
                    $destinatarioEmail = if ($emailColab) { $emailColab } else { $emailGestor }
                    $ccEmail = if ($emailColab) { $emailGestor } else { "" }
                    
                    # Envia email para criar conta no Turia
                    if ($enviarEmail -and $destinatarioEmail) {
                        $emailTuriaEnviado = Send-TuriaRequestEmail -Para $destinatarioEmail -CC $ccEmail -NomeColaborador $nome -Usuario $userAD -CentroCusto $cc
                        if ($emailTuriaEnviado) {
                            # SUCESSO - Email enviado, marcar como concluído
                            $row.Cells["Status"].Value = "SUCESSO"
                            $destTipo = if ($emailColab) { "Colab" } else { "Gestor" }
                            
                            # EXIBE O EMAIL DO DESTINATÁRIO NO GRID
                            $row.Cells["Lideranca"].Value = "Enviado: $destinatarioEmail"
                            $row.Cells["Lideranca"].ToolTipText = "Tipo: $destTipo | CC: $ccEmail"
                            
                            $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Blue
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
                            
                            # Envia log para API para remover da fila
                            $pTuria = @{
                                data_hora         = (Get-Date).ToString("dd/MM/yyyy HH:mm");
                                filial            = $origem;
                                user_name         = $userAD;
                                nome_colaborador  = $nome;
                                email_colaborador = $emailColab;
                                email_gestor      = $emailGestor;
                                centro_custo      = $cc;
                                nova_senha        = "N/A - Conta inexistente";
                                status            = "SUCESSO";
                                executor          = $executor
                            } | ConvertTo-Json
                            try { 
                                Invoke-Retry -ScriptBlock {
                                    Invoke-RestMethod -Uri $urlBase -Method Post -Body $pTuria -ContentType "application/json"
                                } -ErrorMessage "Falha ao atualizar fila (Turia)"
                                Write-Console " > Solicitacao removida da fila."
                            }
                            catch { Write-Console " > Erro ao atualizar fila: $_" }
                        }
                        else {
                            $row.Cells["Status"].Value = "ERRO EMAIL"
                            $row.Cells["Lideranca"].Value = "Erro email [X]"
                            $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Red
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                        }
                    }
                    else {
                        # Sem email de colaborador E sem email de gestor
                        $row.Cells["Status"].Value = "SEM EMAILS"
                        $row.Cells["Lideranca"].Value = "Sem destinatario [X]"
                        $row.Cells["Lideranca"].Style.ForeColor = [System.Drawing.Color]::Red
                        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                        Write-Console " > ERRO: Nenhum email disponivel para envio."
                    }
                }
                elseif ($err -match "senha|password|comprimento|complexidade|histórico") {
                    $row.Cells["Status"].Value = "ERRO POLÍTICA SENHA"
                    Write-Console " > ERRO: Senha não atende política do domínio - $err"
                }
                else {
                    $row.Cells["Status"].Value = "ERRO AD"
                    Write-Console " > ERRO AD: $err"
                }
                
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Processo Concluído.", "Fim", "OK", "Information")
        $statusLabel.Text = "Pronto."
        Write-Console "Processo finalizado."
    })

$btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV|*.csv"; $sfd.FileName = "Reset_Log_Atribuicao.csv"
        if ($sfd.ShowDialog() -eq "OK") {
            $data = foreach ($r in $dataGridView.Rows) { 
                [PSCustomObject]@{
                    Filial   = $r.Cells["Filial"].Value; 
                    User     = $r.Cells["User"].Value; 
                    Analista = $r.Cells["Analista"].Value;
                    Status   = $r.Cells["Status"].Value 
                } 
            }
            $data | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        }
    })

$statusLabel.Text = "Pronto."
$mainForm.ShowDialog()
