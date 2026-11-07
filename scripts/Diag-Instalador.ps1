<#
  Diagnostico: com a janela "Assistente de Atualizacao" ABERTA na primeira tela
  (abra D:\...\C106A1004.exe na mao), rode este script como administrador.
  Ele lista o que o UI Automation enxerga e testa 3 formas de clicar em Avancar.
  Saida tambem salva em D:\DOMINIO\AtualizaDomSis\logs-auto\diag.txt
#>
param([switch]$Click)   # com -Click tenta clicar em "Avancar" (so avanca UMA tela)

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class W32 {
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, System.Text.StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
}
"@

$out = New-Object System.Collections.Generic.List[string]
function P([string]$s) { Write-Host $s; $out.Add($s) }

$root = [System.Windows.Automation.AutomationElement]::RootElement
$win = $null
# espera ate 90s a janela aparecer (o instalador se extrai antes de abrir)
for ($i = 0; $i -lt 45 -and -not $win; $i++) {
    $tops = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($t in $tops) { if ($t.Current.Name -match 'Assistente de Atualiza') { $win = $t; break } }
    if (-not $win) { Start-Sleep -Seconds 2 }
}
if (-not $win) { P "Janela 'Assistente de Atualizacao' NAO encontrada pelo UI Automation.";
    P "Janelas de topo vistas:"; foreach ($t in $tops) { P ("  '{0}' pid={1} class={2}" -f $t.Current.Name, $t.Current.ProcessId, $t.Current.ClassName) }
    $out | Set-Content D:\DOMINIO\AtualizaDomSis\logs-auto\diag.txt; exit 1 }

P ("Janela: '{0}' pid={1} class={2} hwnd={3}" -f $win.Current.Name, $win.Current.ProcessId, $win.Current.ClassName, $win.Current.NativeWindowHandle)
$p = Get-Process -Id $win.Current.ProcessId -ErrorAction SilentlyContinue
if ($p) { P ("Processo: {0}  {1}" -f $p.Name, $p.Path) }
P ("Eu estou elevado? " + ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))

P "---- Controles (UIA) ----"
$all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
P ("total: {0}" -f $all.Count)
$btn = $null
foreach ($e in $all) {
    $c = $e.Current
    $pats = ($e.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName -replace 'PatternIdentifiers.Pattern','' }) -join ','
    P ("{0,-12} name='{1}' class={2} enabled={3} offscreen={4} hwnd={5} patterns=[{6}]" -f `
        $c.ControlType.ProgrammaticName.Replace('ControlType.',''), $c.Name, $c.ClassName, $c.IsEnabled, $c.IsOffscreen, $c.NativeWindowHandle, $pats)
    if (-not $btn -and $c.Name -match 'Avan.ar') { $btn = $e }
}

P "---- Janelas filhas Win32 ----"
$h = [IntPtr]$win.Current.NativeWindowHandle
$cb = [W32+EnumProc]{ param($ch, $l)
    $s = New-Object System.Text.StringBuilder 256; [void][W32]::GetWindowText($ch, $s, 256)
    $k = New-Object System.Text.StringBuilder 256; [void][W32]::GetClassName($ch, $k, 256)
    P ("  hwnd={0} class={1} text='{2}'" -f $ch, $k, $s); $true }
[void][W32]::EnumChildWindows($h, $cb, [IntPtr]::Zero)

if ($Click) {
    if (-not $btn) { P "Botao Avancar nao encontrado via UIA." }
    else {
        P "Teste 1: InvokePattern..."
        try { $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); P "  Invoke OK (olhe se avancou)"; Start-Sleep 3 } catch { P ("  Invoke FALHOU: " + $_.Exception.Message) }
        $hb = [IntPtr]$btn.Current.NativeWindowHandle
        if ($hb -ne [IntPtr]::Zero) {
            P "Teste 2: BM_CLICK (SendMessage) no hwnd $hb..."
            [void][W32]::SendMessage($hb, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero); Start-Sleep 3; P "  enviado (olhe se avancou)"
        }
        P "Teste 3: foco + tecla Enter via SendKeys..."
        try { $btn.SetFocus(); Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); P "  enviado (olhe se avancou)" } catch { P ("  FALHOU: " + $_.Exception.Message) }
    }
}
$out | Set-Content D:\DOMINIO\AtualizaDomSis\logs-auto\diag.txt
Write-Host "`nSalvo em D:\DOMINIO\AtualizaDomSis\logs-auto\diag.txt"
