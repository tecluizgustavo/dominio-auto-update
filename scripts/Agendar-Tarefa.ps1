<#
  Cria a tarefa agendada "Atualiza Dominio Auto" para rodar COM OU SEM usuario
  conectado ("executar estando o usuario conectado ou nao"). Rode ELEVADO no servidor.

  Pede a senha do Windows do usuario que vai executar a tarefa (digitada no prompt;
  o Agendador a guarda; este script nao a grava em lugar nenhum).
  Use o MESMO usuario que rodou Salvar-Credencial.ps1 (a credencial DPAPI so abre para ele).
  O usuario precisa ser administrador local e ter o direito "Fazer logon como tarefa em lote".
#>
param(
    [string]$Script = 'D:\DOMINIO\AtualizaDomSis\Atualiza-Dominio.ps1',
    [string]$Hora   = '05:00'   # depois do download (00:00-07:00 do Gerenciador) e antes do expediente
)

$cred = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message 'Senha do Windows do usuario que executara a tarefa'
if (-not $cred) { Write-Host 'Cancelado.'; exit 1 }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
          -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Script`""
$trig   = New-ScheduledTaskTrigger -Daily -At $Hora
$set    = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 3) -StartWhenAvailable -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName 'Atualiza Dominio Auto' -Action $action -Trigger $trig -Settings $set `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Highest -Force | Out-Null

Write-Host "Tarefa criada para $Hora como $($cred.UserName) (roda com ou sem usuario conectado)."
Write-Host "Teste: schtasks /run /tn `"Atualiza Dominio Auto`"   e depois leia D:\DOMINIO\AtualizaDomSis\logs-auto"
