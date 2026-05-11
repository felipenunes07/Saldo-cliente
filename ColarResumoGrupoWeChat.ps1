param(
    [string]$QueuePath = "",
    [int]$OpenDelayMs = 1200,
    [int]$PasteDelayMs = 1000,
    [switch]$AutoStart,
    [switch]$NoFinalMessage
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Window {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@

function Select-QueuePath {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Escolha a fila do resumo grupo WeChat"
    $dialog.Filter = "Fila resumo WeChat (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
    $dialog.InitialDirectory = "C:\Users\felip\Dropbox\ETH\Saldo cliente\2026"
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    throw "Nenhuma fila foi selecionada."
}

function Copy-ImageToClipboard {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Imagem nao encontrada:`n$Path"
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $image = [System.Drawing.Image]::FromStream($stream)
        $bitmap = New-Object System.Drawing.Bitmap $image
        [System.Windows.Forms.Clipboard]::SetImage($bitmap)
    } finally {
        if ($image) { $image.Dispose() }
        $stream.Dispose()
    }
}

function Activate-WeChat {
    $process = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.MainWindowHandle -ne 0 -and (
                $_.ProcessName -like "*WeChat*" -or
                $_.ProcessName -like "*Weixin*" -or
                $_.MainWindowTitle -like "*WeChat*"
            )
        } |
        Sort-Object ProcessName, Id |
        Select-Object -First 1

    if ($null -eq $process) {
        throw "Nao encontrei o WeChat aberto. Abra o WeChat Desktop e deixe logado antes de rodar."
    }

    [void][Win32Window]::ShowWindowAsync($process.MainWindowHandle, 9)
    Start-Sleep -Milliseconds 250
    [void][Win32Window]::SetForegroundWindow($process.MainWindowHandle)
    Start-Sleep -Milliseconds 350
}

function Send-KeysSafe {
    param([string]$Keys, [int]$DelayMs = 250)
    [System.Windows.Forms.SendKeys]::SendWait($Keys)
    Start-Sleep -Milliseconds $DelayMs
}

if ([string]::IsNullOrWhiteSpace($QueuePath)) {
    $QueuePath = Select-QueuePath
}

$QueuePath = (Resolve-Path -LiteralPath $QueuePath).Path
$items = @(Import-Csv -LiteralPath $QueuePath -Encoding UTF8 | Sort-Object { [int]$_.Ordem })
if ($items.Count -eq 0) {
    throw "Nao encontrei itens na fila:`n$QueuePath"
}

$missing = @($items | Where-Object { -not (Test-Path -LiteralPath $_.Imagem) })
if ($missing.Count -gt 0) {
    $list = ($missing | ForEach-Object { $_.Imagem }) -join [Environment]::NewLine
    throw "Tem imagem faltando na fila:`n$list"
}

Write-Host ""
Write-Host "Modo rascunho automatico: vou abrir o grupo e colar as imagens do resumo."
Write-Host "NAO vou apertar Enter para enviar."
Write-Host ""
Write-Host "Importante: deixe o WeChat aberto, nao mexa no mouse/teclado durante a automacao."
Write-Host "Depois confira o chat antes de enviar."
Write-Host ""
if (-not $AutoStart) {
    Read-Host "Aperte ENTER para iniciar"
} else {
    Start-Sleep -Milliseconds 800
}

$groups = @($items | Group-Object GrupoWeChat)
$totalImages = $items.Count
$done = 0

foreach ($groupInfo in $groups) {
    $group = [string]$groupInfo.Name
    $groupItems = @($groupInfo.Group | Sort-Object { [int]$_.Ordem })

    Write-Host "Grupo: $group"
    Activate-WeChat
    Set-Clipboard -Value $group
    Send-KeysSafe "^f" 300
    Send-KeysSafe "^a" 120
    Send-KeysSafe "^v" 600
    Send-KeysSafe "{ENTER}" $OpenDelayMs

    foreach ($item in $groupItems) {
        $done++
        $tipo = [string]$item.Tipo
        $image = [string]$item.Imagem
        Write-Progress -Activity "Colando resumo no WeChat" -Status "$tipo -> $group" -PercentComplete ([int](($done / $totalImages) * 100))
        Write-Host ("[{0}/{1}] {2}" -f $done, $totalImages, $tipo)

        Activate-WeChat
        Copy-ImageToClipboard -Path $image
        Send-KeysSafe "^v" $PasteDelayMs
        Start-Sleep -Milliseconds 350
    }
}

Write-Progress -Activity "Colando resumo no WeChat" -Completed
Write-Host ""
Write-Host "Concluido. As imagens do resumo foram coladas como rascunho; confira antes de enviar."
if (-not $NoFinalMessage) {
    [System.Windows.Forms.MessageBox]::Show("Resumo colado no WeChat como rascunho. Confira antes de enviar.", "Resumo Grupo WeChat") | Out-Null
}
