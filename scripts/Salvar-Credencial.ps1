<#
  Guarda usuario/senha do login do Dominio CRIPTOGRAFADOS (DPAPI do Windows).
  Rode no servidor, logado com o MESMO usuario do Windows que vai executar a
  tarefa agendada: so ele, nessa maquina, consegue ler o arquivo.
  A senha e digitada no prompt do Windows; nao aparece na tela nem no script.
#>
param(
    [string]$Path     = 'D:\DOMINIO\AtualizaDomSis\cred-dominio.xml',
    [string]$Mensagem = 'Usuario e senha do login do Dominio'
)
# Para o e-mail:  .\Salvar-Credencial.ps1 -Path D:\DOMINIO\AtualizaDomSis\cred-email.xml -Mensagem 'Conta de e-mail que envia o relatorio'

$cred = Get-Credential -Message $Mensagem
if (-not $cred) { Write-Host 'Cancelado.'; exit 1 }
$cred | Export-Clixml -Path $Path

# so o usuario atual (e o sistema) acessa o arquivo
icacls $Path /inheritance:r /grant:r "${env:USERDOMAIN}\${env:USERNAME}:F" "SYSTEM:F" | Out-Null
Write-Host "Credencial de '$($cred.UserName)' salva em $Path"
