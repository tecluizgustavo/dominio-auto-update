<#
.SYNOPSIS
  Aplica automaticamente a atualizacao do Dominio Contabil (POP_00002 v1.1).

.DESCRIPTION
  1. Acha o pacote C*.exe mais novo em D:\DOMINIO\AtualizaDomSis\Contabil
  2. Compara com o ultimo pacote aplicado (arquivo de estado)
  3. Aborta se houver usuarios conectados (porta do banco)
  4. Para Agente_comunicacao e o servico ServicoDominioAtendimento
  5. Executa o instalador e clica nas telas (mensagens Win32, sem UI Automation)
  6. Abre o Dominio, faz login (credencial DPAPI) e clica Sim / ciencia / Iniciar Atualizacao / OK
  7. Sobe servico e Agente_comunicacao, fecha o Dominio Folha, grava log

  Usa so mensagens Win32 (EnumWindows, BM_CLICK, WM_SETTEXT), que funcionam
  numa tarefa agendada "executar estando o usuario conectado ou nao" (sessao
  nao interativa), sem ninguem logado. Precisa rodar ELEVADO.

.PARAMETER Seed
  Marca o pacote mais novo atual como "ja aplicado" e sai.

.PARAMETER DryRun
  Mostra o que faria, sem parar nada nem instalar.
#>
[CmdletBinding()]
param(
    [switch]$Seed,
    [switch]$DryRun,
    [switch]$TestEmail   # so envia um e-mail de teste e sai (nao atualiza nada)
)

# ===================== CONFIGURACAO (ajustar no servidor) =====================
$Cfg = @{
    PackageRoot   = 'D:\DOMINIO\AtualizaDomSis\Contabil'
    StateFile     = 'D:\DOMINIO\AtualizaDomSis\ultimo_pacote_aplicado.txt'
    LogDir        = 'D:\DOMINIO\AtualizaDomSis\logs-auto'

    # Servico do POP: "ServicoDominioAtendimento - 7080" (casa por nome ou nome de exibicao)
    ServiceMatch  = '*ServicoDominioAtendimento*'

    # Processo do agente (sem .exe) e caminho do .exe para reabrir
    AgentProcess  = 'Agente_comunicacao'
    # Vazio = usa o caminho do agente que esta rodando / procura sob D:\DOMINIO\Contabil
    AgentExe      = ''

    # Cliente Dominio que dispara a fase do banco
    ClientProcess = 'contabil'
    ClientExe     = 'D:\DOMINIO\Contabil\contabil.exe'

    # Porta TCP do banco (dbsrv17) para contar usuarios conectados (0 = nao checar)
    DbPort        = 2638

    # Regex do nome do processo (ou pasta) do Dominio Folha, fechado no fim
    FolhaMatch    = 'folha'

    # Credencial do login do Dominio, criptografada (DPAPI). Crie com Salvar-Credencial.ps1
    CredFile      = 'D:\DOMINIO\AtualizaDomSis\cred-dominio.xml'
    MaxLoginTries = 2   # evita bloquear a conta se a senha estiver errada
    # Maximo de MAQUINAS (IPs distintos) conectadas toleradas (as do proprio servidor nao contam)
    MaxConnections = 0

    # ---- E-mail de relatorio (desligado enquanto nao houver Graph ou SMTP configurado, e MailFrom/MailTo) ----
    # Via Microsoft Graph (recomendado): app registration com permissao de APLICATIVO Mail.Send.
    #   GraphTenantId / GraphClientId nao sao segredos. O client secret fica criptografado (DPAPI) em
    #   GraphCredFile, criado com Salvar-Credencial.ps1 (usuario = Client ID, senha = client secret).
    #   MailFrom = caixa que envia (precisa existir no tenant).
    GraphTenantId = ''
    GraphClientId = ''
    GraphCredFile = 'D:\DOMINIO\AtualizaDomSis\cred-graph.xml'

    # Alternativa SMTP (usada so se GraphTenantId estiver vazio):
    #   Microsoft 365 com login:  smtp.office365.com / 587 / SSL (precisa de SMTP AUTH liberado na caixa)
    #   Microsoft 365 por conector (sem senha, libera o IP do servidor): <dominio>.mail.protection.outlook.com / 25 / sem SSL
    SmtpServer   = ''
    SmtpPort     = 587
    SmtpUseSsl   = $true
    MailFrom     = ''
    MailTo       = @()   # varios: @('a@empresa.com.br','b@empresa.com.br')
    # Credencial do e-mail (DPAPI). Opcional: sem o arquivo, envia sem autenticar (modo conector).
    MailCredFile = 'D:\DOMINIO\AtualizaDomSis\cred-email.xml'

    InstallerTimeoutMin = 40
    DbPhaseTimeoutMin   = 20
}
# =============================================================================

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Cfg.LogDir | Out-Null
$LogFile = Join-Path $Cfg.LogDir ("atualiza-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))

function Log([string]$m) {
    $line = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $m
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

# Win32 puro: enumera janelas/controles e manda BM_CLICK / WM_SETTEXT / WM_CLOSE.
# (UI Automation nao e usado: nao funciona de forma confiavel sem sessao interativa.)
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class W32 {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
  [DllImport("user32.dll", EntryPoint="GetWindowLongW")] static extern int GetWindowLong(IntPtr h, int i);
  public static bool IsPassword(IntPtr h) { return (GetWindowLong(h, -16) & 0x20) != 0; }
  [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
  // Visivel RELATIVO a janela de topo: so conta o flag WS_VISIBLE do proprio controle e dos pais
  // ate (sem incluir) a janela de topo. Assim, numa sessao sem area de trabalho (topo invisivel),
  // ainda distingue a pagina atual do assistente das paginas ocultas.
  static bool VisRel(IntPtr h, IntPtr top) {
    IntPtr cur = h;
    while (cur != IntPtr.Zero && cur != top) {
      if ((GetWindowLong(cur, -16) & 0x10000000) == 0) return false;
      cur = GetParent(cur);
    }
    return true;
  }
  // controles filhos: { hwnd, classe, texto, habilitado, visivel }
  public static List<string[]> Children(IntPtr parent) {
    var r = new List<string[]>();
    EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) {
      var t = new StringBuilder(512); GetWindowText(h, t, 512);
      var c = new StringBuilder(256); GetClassName(h, c, 256);
      r.Add(new string[] { h.ToInt64().ToString(), c.ToString(), t.ToString(),
                           IsWindowEnabled(h) ? "1" : "0", VisRel(h, parent) ? "1" : "0" });
      return true;
    }, IntPtr.Zero);
    return r;
  }
  // janelas de topo: { hwnd, pid, titulo, visivel, classe }
  public static List<string[]> Tops() {
    var r = new List<string[]>();
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint pid; GetWindowThreadProcessId(h, out pid);
      var t = new StringBuilder(512); GetWindowText(h, t, 512);
      var c = new StringBuilder(256); GetClassName(h, c, 256);
      r.Add(new string[] { h.ToInt64().ToString(), pid.ToString(), t.ToString(), IsWindowVisible(h) ? "1" : "0", c.ToString() });
      return true;
    }, IntPtr.Zero);
    return r;
  }
}
"@

# ------------------------------ helpers de janela ------------------------------
function New-Win($Hwnd, $WinPid, $Title) {
    [pscustomobject]@{ WinPid = [int]$WinPid; Current = [pscustomobject]@{ NativeWindowHandle = [int64]$Hwnd; Name = $Title } }
}

# Sem sessao interativa (sessao 0) as janelas ficam com WS_VISIBLE=0 no topo; entao
# la aceitamos janelas invisiveis desde que tenham titulo.
$script:Interactive = ((Get-Process -Id $PID).SessionId -ne 0)

function Test-TopOk($t) { return ($t[3] -eq '1') -or (-not $script:Interactive) }

function Get-ProcWindows([int[]]$ProcIds) {
    $out = @()
    foreach ($t in [W32]::Tops()) {
        if ((Test-TopOk $t) -and ($ProcIds -contains [int]$t[1])) { $out += (New-Win $t[0] $t[1] $t[2]) }
    }
    return $out
}

function Get-TitleWindows([string]$Regex) {
    $out = @()
    foreach ($t in [W32]::Tops()) {
        if ((Test-TopOk $t) -and $t[2] -match $Regex) { $out += (New-Win $t[0] $t[1] $t[2]) }
    }
    return $out
}

# PIDs do processo e de todos os descendentes (instaladores Inno Setup se
# reexecutam em um processo filho que e o dono da janela).
function Get-DescendantPids([int]$RootPid) {
    # so processos da MESMA sessao (evita PID reaproveitado de processos de outras sessoes)
    $mySess = (Get-Process -Id $PID).SessionId
    $all = Get-CimInstance Win32_Process | Where-Object { $_.SessionId -eq $mySess } | Select-Object ProcessId, ParentProcessId
    $ids = New-Object System.Collections.Generic.List[int]
    $ids.Add($RootPid)
    for ($i = 0; $i -lt $ids.Count; $i++) {
        foreach ($c in ($all | Where-Object { $_.ParentProcessId -eq $ids[$i] })) {
            if (-not $ids.Contains([int]$c.ProcessId)) { $ids.Add([int]$c.ProcessId) }
        }
    }
    return $ids.ToArray()
}

function Clean([string]$s) { return ($s -replace '&', '').Trim() }

function Get-Win32Children($win) {
    $h = [int64]$win.Current.NativeWindowHandle
    if ($h -eq 0) { return @() }
    return [W32]::Children([IntPtr]$h)
}

# BM_CLICK (assincrono) no primeiro botao visivel/habilitado cujo texto casa. Retorna o texto.
function Click-Button($win, [string]$Regex) {
    foreach ($c in (Get-Win32Children $win)) {
        if ($c[1] -match 'Button|BitBtn' -and $c[1] -notmatch 'Check|Radio' -and $c[3] -eq '1' -and $c[4] -eq '1') {
            $n = Clean $c[2]
            if ($n -match $Regex) {
                [void][W32]::PostMessage([IntPtr][int64]$c[0], 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
                return $n
            }
        }
    }
    return $null
}

# Campos de edicao da janela: @{ User=<hwnd|0>; Pass=<hwnd|0> }
function Get-LoginFields($win) {
    $r = @{ User = [int64]0; Pass = [int64]0 }
    foreach ($c in (Get-Win32Children $win)) {
        if ($c[1] -match '^T?(New|Mask)?Edit$' -and $c[3] -eq '1' -and $c[4] -eq '1') {
            $h = [int64]$c[0]
            if ([W32]::IsPassword([IntPtr]$h)) { if (-not $r.Pass) { $r.Pass = $h } }
            elseif (-not $r.User) { $r.User = $h }
        }
    }
    return $r
}

# Preenche usuario/senha (WM_SETTEXT=0x0C). A senha so existe em memoria durante a chamada.
function Fill-Login($win, $cred) {
    $f = Get-LoginFields $win
    if (-not $f.Pass) { return $false }
    if ($f.User) { [void][W32]::SendMessage([IntPtr]$f.User, 0x000C, [IntPtr]::Zero, $cred.UserName) }
    $plain = $cred.GetNetworkCredential().Password
    [void][W32]::SendMessage([IntPtr]$f.Pass, 0x000C, [IntPtr]::Zero, $plain)
    $plain = $null
    return $true
}

# Marca checkbox/radio de aceite/ciencia (BM_GETCHECK=0xF0, BM_CLICK=0xF5). $true se marcou algo.
function Check-Accept($win) {
    $did = $false
    foreach ($c in (Get-Win32Children $win)) {
        if ($c[1] -match 'Check|Radio' -and $c[3] -eq '1' -and $c[4] -eq '1') {
            $n = Clean $c[2]
            if ($n -match 'aceito|concordo|de acordo|ci.ncia' -and $n -notmatch 'n.o aceito|n.o concordo') {
                $st = [W32]::SendMessage([IntPtr][int64]$c[0], 0x00F0, [IntPtr]::Zero, [IntPtr]::Zero)
                if ([int64]$st -ne 1) { [void][W32]::PostMessage([IntPtr][int64]$c[0], 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero); $did = $true }
            }
        }
    }
    return $did
}

function Get-WinText($win) {
    $txt = @()
    foreach ($c in (Get-Win32Children $win)) { if ($c[2]) { $txt += $c[2] } }
    return ($txt -join ' ')
}

# Fecha as janelas dos processos (WM_CLOSE) e, se sobrar processo, encerra.
function Close-Procs([int[]]$Ids, [string]$Label) {
    if (-not $Ids -or $Ids.Count -eq 0) { return }
    foreach ($w in (Get-ProcWindows $Ids)) {
        [void][W32]::PostMessage([IntPtr]$w.Current.NativeWindowHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    }
    Start-Sleep -Seconds 8
    foreach ($p in $Ids) {
        $left = Get-Process -Id $p -ErrorAction SilentlyContinue
        if ($left) { Log "$Label nao fechou sozinho; encerrando $($left.Name) (pid $p)."; Stop-Process -Id $p -Force }
    }
}

# Sobe um programa FORA da arvore da tarefa agendada (senao o Agendador mata o processo
# quando o script termina). Usa Win32_Process.Create; cai para Start-Process se falhar.
function Start-Detached([string]$Path) {
    $dir = Split-Path $Path
    try {
        $r = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = ('"{0}"' -f $Path); CurrentDirectory = $dir }
        if ($r.ReturnValue -eq 0) { return }
        Log "AVISO: Win32_Process.Create retornou $($r.ReturnValue)."
    } catch { Log "AVISO: Win32_Process.Create falhou: $($_.Exception.Message)" }
    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($admin) { Start-Process -FilePath $Path -WorkingDirectory $dir }
    else { Start-Process -FilePath $Path -WorkingDirectory $dir -Verb RunAs }
}

# A tela principal do Dominio ja carregou? (barra "Lanctos Contabeis" / "Barra Dominio Atendimento")
function Test-MainLoaded([int[]]$Ids) {
    foreach ($t in [W32]::Tops()) {
        if ($Ids -notcontains [int]$t[1]) { continue }
        foreach ($c in [W32]::Children([IntPtr][int64]$t[0])) {
            if ($c[2] -match 'Lanctos Cont|Barra Dom.nio') { return $true }
        }
    }
    return $false
}

# Falha na fase do banco: nao deixar o ambiente parado. Fecha o cliente e sobe o agente.
function Restore-AfterFailure {
    Close-Procs @(Get-Process -Name $Cfg.ClientProcess -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) 'Cliente Dominio'
    Start-Detached $agentPath
    Log "Ambiente restaurado (cliente fechado, agente reiniciado). Estado NAO gravado."
}

# ------------------------------ helpers gerais -------------------------------
function Get-LatestPackage {
    $pk = Get-ChildItem -Path $Cfg.PackageRoot -Recurse -File -Filter 'C*.exe' |
          Where-Object { $_.BaseName -match '^C\d{3}[A-Z]\d{4}$' } |
          Sort-Object BaseName -Descending
    if (-not $pk) { return $null }
    return $pk | Select-Object -First 1
}

function Get-UserConnections {
    if ($Cfg.DbPort -le 0) { return -1 }
    $local = @('127.0.0.1', '::1') + @(Get-NetIPAddress -ErrorAction SilentlyContinue | ForEach-Object { $_.IPAddress })
    $c = Get-NetTCPConnection -State Established -LocalPort $Cfg.DbPort -ErrorAction SilentlyContinue |
         Where-Object { $_.RemoteAddress -notin $local }
    # conta MAQUINAS distintas (cada estacao abre varias conexoes)
    $script:ConnHosts = (@($c | Select-Object -ExpandProperty RemoteAddress -Unique) -join ', ')
    return @($c | Select-Object -ExpandProperty RemoteAddress -Unique).Count
}

function Stop-Agent {
    Get-Process -Name $Cfg.AgentProcess -ErrorAction SilentlyContinue | Stop-Process -Force
}

# Envia o relatorio por e-mail (versao + log em anexo). Roda no fim, com sucesso ou falha.
# So envia se o script chegou a agir ($script:Acted) ou se houve ERRO; "Sem pacote novo" nao gera e-mail.
function Send-Report {
    if ($DryRun -or $Seed) { return }
    $useGraph = [bool]($Cfg.GraphTenantId -and $Cfg.GraphClientId)
    if (-not $useGraph -and -not $Cfg.SmtpServer) { return }
    if (-not $Cfg.MailTo -or $Cfg.MailTo.Count -eq 0 -or -not $Cfg.MailFrom) { return }
    try {
        $text = Get-Content -LiteralPath $LogFile -Raw -ErrorAction SilentlyContinue
        $ok  = $text -match 'SUCESSO:'
        $bad = $text -match 'ERRO:|FINALIZADO SEM CONFIRMAR|ABORTADO'
        if (-not ($script:Acted -or ($text -match 'ERRO:'))) { return }

        $pkgName = if ($pkg) { $pkg.BaseName } else { '?' }
        $verTxt = $pkgName
        if ($pkgName -match '^C(\d{2})(\d[A-Z])(\d{2})(\d{2})$') { $verTxt = "{0}.{1}-{2} (atualizacao {3}) [{4}]" -f $Matches[1], $Matches[2], $Matches[3], $Matches[4], $pkgName }
        $status = if ($ok) { 'SUCESSO' } elseif ($bad) { 'FALHA' } else { 'INCOMPLETO' }

        $tail = (($text -split "`r?`n" | Where-Object { $_ } | Select-Object -Last 15) -join "`r`n")
        $body = @"
Atualizacao automatica do Dominio Contabil

Resultado : $status
Servidor  : $env:COMPUTERNAME
Data/hora : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
Versao    : $verTxt
Anterior  : $applied

Ultimas linhas do log (log completo em anexo):

$tail
"@
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # anexa uma COPIA do log: se o envio falhar, o anexo fica travado e nao pode bloquear o log original
        $attach = Join-Path $env:TEMP ('envio-' + (Split-Path $LogFile -Leaf))
        Copy-Item -LiteralPath $LogFile -Destination $attach -Force
        $subject = "[Dominio] $status - $verTxt - $env:COMPUTERNAME"
        if ($useGraph) {
            # --- Microsoft Graph, fluxo client_credentials (permissao de aplicativo Mail.Send) ---
            $gc = Import-Clixml -Path $Cfg.GraphCredFile
            $secret = $gc.GetNetworkCredential().Password
            $tok = Invoke-RestMethod -Method Post -ErrorAction Stop `
                -Uri ("https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $Cfg.GraphTenantId) `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body @{ client_id = $Cfg.GraphClientId; client_secret = $secret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' }
            $secret = $null

            # anexo inline do Graph vai ate ~3 MB; se o log for maior, anexa so o final
            $bytes = [IO.File]::ReadAllBytes($attach)
            if ($bytes.Length -gt 2500000) { $bytes = $bytes[($bytes.Length - 2500000)..($bytes.Length - 1)] }
            $payload = @{
                message = @{
                    subject      = $subject
                    body         = @{ contentType = 'Text'; content = $body }
                    toRecipients = @($Cfg.MailTo | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
                    attachments  = @(@{
                        '@odata.type' = '#microsoft.graph.fileAttachment'
                        name          = (Split-Path $LogFile -Leaf)
                        contentType   = 'text/plain'
                        contentBytes  = [Convert]::ToBase64String($bytes)
                    })
                }
                saveToSentItems = $false
            }
            $json = $payload | ConvertTo-Json -Depth 8
            Invoke-RestMethod -Method Post -ErrorAction Stop `
                -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $Cfg.MailFrom) `
                -Headers @{ Authorization = "Bearer $($tok.access_token)" } `
                -ContentType 'application/json; charset=utf-8' `
                -Body ([Text.Encoding]::UTF8.GetBytes($json)) | Out-Null
        } else {
            $mail = @{
                SmtpServer = $Cfg.SmtpServer; Port = $Cfg.SmtpPort; From = $Cfg.MailFrom; To = $Cfg.MailTo
                Subject    = $subject
                Body       = $body; Attachments = $attach
                Encoding   = [Text.Encoding]::UTF8
            }
            if ($Cfg.SmtpUseSsl) { $mail.UseSsl = $true }
            if (Test-Path -LiteralPath $Cfg.MailCredFile) { $mail.Credential = Import-Clixml -Path $Cfg.MailCredFile }
            Send-MailMessage @mail -ErrorAction Stop
        }
        Log "E-mail enviado para $($Cfg.MailTo -join ', ')."
    } catch {
        Log "AVISO: falha ao enviar e-mail: $($_.Exception.Message)"
    } finally {
        if ($attach) { Remove-Item -LiteralPath $attach -Force -ErrorAction SilentlyContinue }
    }
}

# ================================== MAIN =====================================
try {
Log "Inicio. Log: $LogFile  DryRun=$DryRun  Usuario=$env:USERNAME  Sessao=$((Get-Process -Id $PID).SessionId)"

if ($TestEmail) {
    $script:Acted = $true
    Log "SUCESSO: teste de envio de e-mail (nenhuma atualizacao foi feita)."
    exit 0
}

$pkg = Get-LatestPackage
if (-not $pkg) { Log "Nenhum pacote C*.exe em $($Cfg.PackageRoot). Nada a fazer."; exit 0 }
Log "Pacote mais novo: $($pkg.BaseName)  ($($pkg.FullName))"

if ($Seed) {
    Set-Content -Path $Cfg.StateFile -Value $pkg.BaseName
    Log "Estado inicializado: $($pkg.BaseName) marcado como aplicado."
    exit 0
}

if (-not (Test-Path $Cfg.StateFile)) {
    Log "ERRO: sem arquivo de estado. Rode primeiro com -Seed (apos confirmar a versao atual)."
    exit 2
}
$applied = (Get-Content $Cfg.StateFile -Raw).Trim()
Log "Ultimo aplicado: $applied"

if ([string]::CompareOrdinal($pkg.BaseName, $applied) -le 0) {
    Log "Sem pacote novo. Fim."
    exit 0
}

# Pacote ainda sendo baixado? (modificado ha menos de 10 min)
if ((Get-Date) - $pkg.LastWriteTime -lt [TimeSpan]::FromMinutes(10)) {
    Log "Pacote alterado ha menos de 10 min; talvez ainda baixando. Aborta, tenta na proxima execucao."
    exit 0
}

# ---- valida caminhos ANTES de parar qualquer coisa
$agentPath = $Cfg.AgentExe
if (-not $agentPath) {
    $ap = Get-Process -Name $Cfg.AgentProcess -ErrorAction SilentlyContinue | Where-Object { $_.Path } | Select-Object -First 1
    if ($ap) { $agentPath = $ap.Path }
}
if (-not $agentPath) {
    # agente fora do ar: procura o .exe sob a pasta do Dominio
    $found = Get-ChildItem -Path (Split-Path $Cfg.PackageRoot -Parent | Split-Path -Parent | Join-Path -ChildPath 'Contabil') `
             -Recurse -Filter ($Cfg.AgentProcess + '.exe') -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { $agentPath = $found.FullName }
}
if (-not $agentPath -or -not (Test-Path -LiteralPath $agentPath)) {
    Log "ERRO: caminho do agente nao encontrado (agente fora do ar e AgentExe vazio). Preencha AgentExe."; exit 6
}
if (-not $Cfg.ClientExe -or -not (Test-Path -LiteralPath $Cfg.ClientExe)) {
    Log "ERRO: ClientExe invalido ('$($Cfg.ClientExe)'). Preencha o caminho do .exe do Dominio."; exit 6
}
Log "Agente: $agentPath"
Log "Cliente: $($Cfg.ClientExe)"

$users = Get-UserConnections
if ($users -lt 0) { Log "AVISO: DbPort=0, checagem de usuarios conectados DESLIGADA." }
elseif ($users -gt $Cfg.MaxConnections) {
    Log "ABORTADO: $users maquina(s) conectada(s) na porta $($Cfg.DbPort): $script:ConnHosts. Tenta na proxima execucao."
    exit 3
}

if ($DryRun) {
    Log "[DryRun] Pararia '$($Cfg.AgentProcess)' e servico '$($Cfg.ServiceMatch)'; executaria $($pkg.FullName); automatizaria telas; abriria $($Cfg.ClientExe); subiria agente e servico."
    exit 0
}

# ---- 4. parar agente e servico
$script:Acted = $true
Log "Parando $($Cfg.AgentProcess)..."
Stop-Agent
$svc = Get-Service | Where-Object { $_.Name -like $Cfg.ServiceMatch -or $_.DisplayName -like $Cfg.ServiceMatch }
if (-not $svc) { Log "AVISO: servico '$($Cfg.ServiceMatch)' nao encontrado." }
foreach ($s in $svc) {
    Log "Parando servico $($s.Name)..."
    if ($s.Status -ne 'Stopped') { Stop-Service -InputObject $s -Force; $s.WaitForStatus('Stopped', '00:02:00') }
}

# ---- 5. instalador
Log "Executando instalador $($pkg.Name)..."
$inst = Start-Process -FilePath $pkg.FullName -PassThru
$deadline = (Get-Date).AddMinutes($Cfg.InstallerTimeoutMin)
$tick = 0
# Ancorado: senao casa "Avancar pouco/muito" (botoes da barra de rolagem da tela de Informacao)
$order = '^(Concluir|Finalizar|Atualizar|Instalar|Avan.ar\s*>?|Pr.ximo\s*>?|Next\s*>?|Install|Finish|OK|Sim)$'
while (-not $inst.HasExited) {
    if ((Get-Date) -gt $deadline) { Log "ERRO: timeout do instalador."; Stop-Process -Id $inst.Id -Force; exit 4 }
    $wins = @(Get-ProcWindows (Get-DescendantPids $inst.Id))
    if ($wins.Count -eq 0) { $wins = @(Get-TitleWindows 'Assistente de Atualiza') }   # fallback por titulo
    $any = $false
    foreach ($w in $wins) {
        if (Check-Accept $w) { Log "Aceite marcado." }
        $clicked = Click-Button $w $order
        if ($clicked) { Log "Instalador: clicou '$clicked'"; $any = $true; Start-Sleep -Seconds 2; break }
    }
    if (-not $any) {
        # diagnostico a cada ~10s: o que o script enxerga
        $tick++
        if ($tick % 5 -eq 1) {
            $desc = @()
            foreach ($w in $wins) {
                $bn = (Get-Win32Children $w | Where-Object { $_[1] -match 'Button|BitBtn' } | ForEach-Object { Clean $_[2] }) -join '|'
                $desc += ("[{0}] botoes: {1}" -f $w.Current.Name, $bn)
            }
            Log ("diag: {0} janela(s) {1}" -f $wins.Count, ($desc -join ' ; '))
            # diagnostico extra: processos do instalador e TODAS as janelas de topo deles (inclusive invisiveis)
            $dp = Get-DescendantPids $inst.Id
            $pp = @(Get-Process -Id $dp -ErrorAction SilentlyContinue | ForEach-Object { '{0}(pid {1}, sessao {2})' -f $_.Name, $_.Id, $_.SessionId }) -join ', '
            $tt = @([W32]::Tops() | Where-Object { $dp -contains [int]$_[1] } | ForEach-Object { 'pid {0} classe={1} vis={2} titulo=''{3}''' -f $_[1], $_[4], $_[3], $_[2] }) -join ' ; '
            Log ("diag+: procs=[{0}] topos=[{1}]" -f $pp, $tt)
        }
    }
    Start-Sleep -Seconds 2
    $inst.Refresh()
}
Log "Instalador terminou (ExitCode=$($inst.ExitCode))."

# ---- servico de volta
foreach ($s in $svc) {
    Log "Iniciando servico $($s.Name)..."
    Start-Service -Name $s.Name
}

# Instalador com erro/cancelado: nao segue, devolve o ambiente ao estado anterior
if ($inst.ExitCode -ne 0) {
    Log "ERRO: instalador saiu com ExitCode=$($inst.ExitCode) (2 = cancelado/fechado antes de concluir). Abortando; estado NAO gravado."
    Start-Detached $agentPath
    Log "Agente reiniciado. Confira a versao do sistema antes de tentar de novo."
    exit 7
}

# ---- 6. fase do banco pelo cliente Dominio
Log "Abrindo cliente Dominio para atualizar o banco..."
$cli = Start-Process -FilePath $Cfg.ClientExe -PassThru
$deadline = (Get-Date).AddMinutes($Cfg.DbPhaseTimeoutMin)
$started = $false; $done = $false; $loginTries = 0; $warnedLogin = $false; $backupAsked = $false
$noUpdate = $false; $mainSeenAt = $null
$cred = $null
if (Test-Path -LiteralPath $Cfg.CredFile) {
    try { $cred = Import-Clixml -Path $Cfg.CredFile; Log "Credencial carregada para '$($cred.UserName)'." }
    catch { Log "AVISO: nao consegui abrir $($Cfg.CredFile) (outro usuario do Windows/maquina?). Login sera manual." }
} else { Log "AVISO: sem $($Cfg.CredFile); login sera manual." }
$tickDb = 0
while (-not $done) {
    if ((Get-Date) -gt $deadline) { Log "ERRO: timeout da fase do banco (login ou tela nao reconhecida)."; Restore-AfterFailure; exit 5 }
    $ids = @(Get-Process -Name $Cfg.ClientProcess -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $dbWins = @(Get-ProcWindows $ids)
    $acted = $false
    foreach ($w in $dbWins) {
        $txt = Get-WinText $w

        # "Confirma o backup COMPLETO agora?" : nao fazer backup agora. Cancelar e finalizar.
        if ($txt -match 'Confirma o backup') {
            $b = Click-Button $w '^Cancelar$'
            if ($b) { Log "Banco: apareceu confirmacao de backup; clicou '$b' e vai finalizar."; $backupAsked = $true; $done = $true; $acted = $true; break }
        }

        # tela de login (campo de senha presente)
        if ((Get-LoginFields $w).Pass) {
            if (-not $cred) {
                if (-not $warnedLogin) { Log "Banco: tela de login; sem credencial salva (rode Salvar-Credencial.ps1)."; $warnedLogin = $true }
            } elseif ($loginTries -ge $Cfg.MaxLoginTries) {
                Log "ERRO: login nao passou apos $loginTries tentativa(s). Parando para nao bloquear a conta."; Restore-AfterFailure; exit 8
            } else {
                $loginTries++
                [void](Fill-Login $w $cred)
                Log "Banco: usuario/senha preenchidos (tentativa $loginTries)."
                Start-Sleep -Seconds 1
                $b = Click-Button $w '^(OK|Entrar|Acessar|Confirmar|Login)$'
                if ($b) { Log "Banco: login, clicou '$b'" }
                Start-Sleep -Seconds 3
            }
            $acted = $true; break
        }

        # aviso "voce nao marcou a opcao..." : so fechar com OK (o checkbox e marcado antes do Iniciar)
        if ($txt -match 'n.o marcou a op') {
            $b = Click-Button $w '^OK$'
            if ($b) { Log "Banco: fechou aviso de checkbox ('$b')."; $acted = $true; break }
        }

        if ($started -and $txt -match 'conclu') {
            $b = Click-Button $w '^OK$'
            if ($b) { Log "Banco: atualizacao concluida, clicou OK."; $done = $true; $acted = $true; break }
        }

        # tela de backup: marcar a caixa de ciencia e SO ENTAO clicar Iniciar Atualizacao
        if (Check-Accept $w) { Log "Banco: caixa de ciencia marcada."; Start-Sleep -Seconds 2 }
        $b = Click-Button $w '^Iniciar Atualiza..o$'
        if ($b) { Log "Banco: clicou '$b'"; $started = $true; $acted = $true; break }
        if (-not $started) {
            $b = Click-Button $w '^Sim$'
            if ($b) { Log "Banco: clicou 'Sim'"; $acted = $true; break }
        }
    }
    # Sem pedido de atualizacao: se a tela principal carregou e nada apareceu em 45s, o banco ja esta atualizado.
    if ($acted) { $mainSeenAt = $null }
    elseif (-not $started -and -not $done -and (Test-MainLoaded $ids)) {
        if (-not $mainSeenAt) { $mainSeenAt = Get-Date; Log "Banco: tela principal carregada; aguardando eventual pedido de atualizacao (45s)." }
        elseif (((Get-Date) - $mainSeenAt).TotalSeconds -gt 45) {
            Log "Banco: nenhum pedido de atualizacao apos a tela principal; banco ja estava atualizado."
            $noUpdate = $true; $done = $true
        }
    }

    if (-not $acted) {
        # diagnostico a cada ~30s: janelas e botoes vistos (ajuda se alguma tela nao for reconhecida)
        $tickDb++
        if ($tickDb % 10 -eq 1) {
            $desc = @()
            # resumo compacto: so janelas com titulo, texto nos filhos ou botoes; o resto vira contagem por classe
            $anon = @{}
            foreach ($tp in [W32]::Tops()) {
                if ($ids -notcontains [int]$tp[1]) { continue }
                $ch = @([W32]::Children([IntPtr][int64]$tp[0]))
                $btn = ($ch | Where-Object { $_[1] -match 'Button|BitBtn' -and $_[2] } | ForEach-Object { Clean $_[2] }) -join '|'
                $kids = @()
                foreach ($c in ($ch | Where-Object { $_[2] } | Select-Object -First 8)) {
                    $tx = ($c[2] -replace '\s+', ' '); if ($tx.Length -gt 40) { $tx = $tx.Substring(0, 40) }
                    $kids += ('{0}:{1}' -f $c[1], $tx)
                }
                if ($tp[2] -or $kids.Count -gt 0 -or $btn) {
                    $desc += ("classe={0} vis={1} titulo='{2}' botoes=[{3}] filhos=[{4}]" -f $tp[4], $tp[3], $tp[2], $btn, ($kids -join ', '))
                } else { $anon[$tp[4]] = 1 + [int]$anon[$tp[4]] }
            }
            $anonTxt = ($anon.GetEnumerator() | Sort-Object Name | ForEach-Object { '{0} x{1}' -f $_.Name, $_.Value }) -join ', '
            Log ("diag banco: {0} | outras: {1}" -f ($desc -join ' ; '), $anonTxt)
        }
    }
    Start-Sleep -Seconds 3
}

# ---- fecha cliente e sobe agente
Close-Procs @(Get-Process -Name $Cfg.ClientProcess -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) 'Cliente Dominio'
Log "Iniciando $($Cfg.AgentProcess)..."
Start-Detached $agentPath

# fecha o Dominio Folha (se aberto)
Start-Sleep -Seconds 5
$folha = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $Cfg.FolhaMatch -or $_.Path -match ('\\' + $Cfg.FolhaMatch + '\\') })
foreach ($f in $folha) { Log "Fechando Folha: $($f.Name) (pid $($f.Id))" }
Close-Procs @($folha | ForEach-Object { $_.Id }) 'Folha'

if ($backupAsked) {
    Log "FINALIZADO SEM CONFIRMAR: o sistema pediu backup e foi cancelado. Estado NAO gravado; confira a versao e a atualizacao do banco."
    exit 9
}

if ($noUpdate) { Log "AVISO: o Dominio nao pediu atualizacao do banco (provavelmente ja atualizado). Se isto for uma atualizacao nova, confira a versao do banco." }
Set-Content -Path $Cfg.StateFile -Value $pkg.BaseName
Log "SUCESSO: $($pkg.BaseName) aplicado. Confira os 2 icones verdes do agente e a versao no canto superior direito."
exit 0
} finally {
    Send-Report
}
