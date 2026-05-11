param(
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

function Select-WorkbookPath {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Escolha a planilha Saldo cliente do dia"
    $dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Todos os arquivos (*.*)|*.*"
    $dialog.InitialDirectory = "C:\Users\felip\Dropbox\ETH\Saldo cliente\2026"
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    throw "Nenhuma planilha foi selecionada."
}

function Run-Step {
    param(
        [string]$Title,
        [scriptblock]$Action
    )

    Write-Host ""
    Write-Host "============================================================"
    Write-Host $Title
    Write-Host "============================================================"
    & $Action
}

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Select-WorkbookPath
}

$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$root = $PSScriptRoot
$workbookDir = Split-Path -Parent $WorkbookPath

$prepararScript = Join-Path $root "PrepararResumoGrupoWeChat.ps1"
$colarScript = Join-Path $root "ColarResumoGrupoWeChat.ps1"

foreach ($script in @($prepararScript, $colarScript)) {
    if (-not (Test-Path -LiteralPath $script)) {
        throw "Nao encontrei o script necessario:`n$script"
    }
}

Write-Host "Planilha selecionada:"
Write-Host $WorkbookPath

$beforePrepare = Get-Date
Run-Step "1/2 - Preparar Resumo Grupo WeChat" {
    & $prepararScript -WorkbookPath $WorkbookPath -Silent
}

$queueDir = Get-ChildItem -LiteralPath $workbookDir -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "Resumo_Grupo_Envio_*" -and $_.LastWriteTime -ge $beforePrepare.AddMinutes(-1) } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $queueDir) {
    $queueDir = Get-ChildItem -LiteralPath $workbookDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "Resumo_Grupo_Envio_*" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

if ($null -eq $queueDir) {
    throw "Nao encontrei a pasta Resumo_Grupo_Envio criada pelo preparo."
}

$queuePath = Join-Path $queueDir.FullName "fila_resumo_grupo_wechat.csv"
if (-not (Test-Path -LiteralPath $queuePath)) {
    throw "Nao encontrei a fila de resumo:`n$queuePath"
}

Write-Host ""
Write-Host "Fila encontrada:"
Write-Host $queuePath

Run-Step "2/2 - Colar Resumo Grupo WeChat" {
    & $colarScript -QueuePath $queuePath -AutoStart -NoFinalMessage
}

Write-Host ""
Write-Host "Resumo grupo finalizado."
