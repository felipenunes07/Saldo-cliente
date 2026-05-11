param(
    [string]$QueuePath = "",
    [int]$StartAt = 1,
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
    $dialog.Title = "Escolha a fila de envio WeChat"
    $dialog.Filter = "Fila WeChat (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
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
                $_.MainWindowTitle -like "*WeChat*" -or
                $_.MainWindowTitle -like "*微信*"
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
$items = @(Import-Csv -LiteralPath $QueuePath -Encoding UTF8 | Where-Object { $_.Status -eq "OK" })
if ($items.Count -eq 0) {
    throw "Nao encontrei itens com Status OK na fila:`n$QueuePath"
}

if ($StartAt -lt 1) { $StartAt = 1 }

Write-Host ""
Write-Host "Modo rascunho automatico: vou abrir cada grupo e colar a imagem."
Write-Host "NAO vou apertar Enter para enviar."
Write-Host ""
Write-Host "Importante: deixe o WeChat aberto, nao mexa no mouse/teclado durante a automacao."
Write-Host "Depois confira os chats antes de enviar."
Write-Host ""
if (-not $AutoStart) {
    Read-Host "Aperte ENTER para iniciar"
} else {
    Start-Sleep -Milliseconds 800
}

$total = $items.Count
for ($i = $StartAt - 1; $i -lt $total; $i++) {
    $item = $items[$i]
    $position = $i + 1
    $group = [string]$item.GrupoWeChat
    $client = [string]$item.Cliente
    $image = [string]$item.Imagem
    $saldoImage = [string]$item.ImagemSaldo

    Write-Progress -Activity "Colando rascunhos no WeChat" -Status "Cliente $client -> $group" -PercentComplete ([int](($position / $total) * 100))
    Write-Host ("[{0}/{1}] {2} -> {3}" -f $position, $total, $client, $group)

    Activate-WeChat

    Set-Clipboard -Value $group
    Send-KeysSafe "^f" 300
    Send-KeysSafe "^a" 120
    Send-KeysSafe "^v" 600
    Send-KeysSafe "{ENTER}" $OpenDelayMs

    Activate-WeChat
    Copy-ImageToClipboard -Path $image
    Send-KeysSafe "^v" $PasteDelayMs

    if (-not [string]::IsNullOrWhiteSpace($saldoImage)) {
        Start-Sleep -Milliseconds 350
        Copy-ImageToClipboard -Path $saldoImage
        Send-KeysSafe "^v" $PasteDelayMs
    }

    Start-Sleep -Milliseconds 400
}

Write-Progress -Activity "Colando rascunhos no WeChat" -Completed
Write-Host ""
Write-Host "Concluido. As imagens foram coladas como rascunho; confira no WeChat antes de enviar."
if (-not $NoFinalMessage) {
    [System.Windows.Forms.MessageBox]::Show("Rascunhos colados no WeChat. Confira antes de enviar.", "WeChat Rascunhos") | Out-Null
}
