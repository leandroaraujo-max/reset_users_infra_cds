# ==============================================================================
# SCRIPT DE INSTALAÇÃO DA TAREFA AGENDADA (SYNC ESPELHO AD)
# ==============================================================================
# Execute este script COMO ADMINISTRADOR no servidor onde o script foi salvo.

$TaskName = "AD_Sync_Espelho_Agent"
$ScriptPath = "$PSScriptRoot\Sync_Espelho_AD.ps1"
$Description = "Agente de sincronização para espelhamento de usuários (Google Sheets <-> Active Directory Local)"

# Verifica se o script existe
if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERRO: Script não encontrado em: $ScriptPath" -ForegroundColor Red
    exit
}

# Caminho do PowerShell
$PowershellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

# Ação da Tarefa
$Action = New-ScheduledTaskAction -Execute $PowershellPath -Argument $Arguments

# Disparador (Trigger) - Iniciar diariamente e repetir a cada 1 minuto
$Trigger = New-ScheduledTaskTrigger -Daily -At (Get-Date).Date.AddHours(6)
# Configura a repetição manualmente acessando as propriedades do objeto COM interno se necessário, 
# mas via PowerShell puro, ajustamos assim:
$Trigger.Repetition.Interval = "PT1M" # 1 Minuto (Formato ISO8601 Duration)
$Trigger.Repetition.Duration = "P3650D" # Indeterminado/Longo (10 anos)

# Configurações da Tarefa
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)

# Principal (Usuário que executa - SYSTEM para garantir permissões locais)
# ATENÇÃO: Requer PowerShell como ADMINISTRADOR
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount

# Cria/Atualiza a Tarefa
try {
    # Tenta remover se já existir
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    
    # Registra a nova tarefa
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Description $Description -ErrorAction Stop
    
    Write-Host "SUCESSO: Tarefa '$TaskName' criada e agendada para rodar a cada 1 minuto." -ForegroundColor Green
    Write-Host "Script Alvo: $ScriptPath" -ForegroundColor Cyan
    Write-Host "NOTA: Verifique se o serviço 'Agendador de Tarefas' está rodando." -ForegroundColor Gray
} catch {
    Write-Host "FALHA CRÍTICA: Não foi possível criar a tarefa." -ForegroundColor Red
    Write-Host "Erro Detalhado: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "SOLUÇÃO: Execute o PowerShell como ADMINISTRADOR (Botão direito -> Executar como Administrador)." -ForegroundColor White
}
