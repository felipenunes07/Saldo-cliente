param(
    [string]$WorkbookPath = "",
    [string]$MapPath = "",
    [switch]$Silent,
    [switch]$NoFinalMessage
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
if (-not $Silent) {
}

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Show-Status {
    param([string]$Message, [int]$Current = 0, [int]$Total = 0)
    if ($Silent) { return }
    if ($Total -gt 0) {
        $percent = [Math]::Min(100, [Math]::Max(0, [int](($Current / $Total) * 100)))
        Write-Progress -Activity "Preparando envio WeChat" -Status $Message -PercentComplete $percent
        Write-Host ("[{0,3}%] {1}" -f $percent, $Message)
    } else {
        Write-Progress -Activity "Preparando envio WeChat" -Status $Message
        Write-Host "[...] $Message"
    }
}

function Write-PrepareLog {
    param([string]$Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss.fff} - {1}" -f (Get-Date), $Message
    if (-not [string]::IsNullOrWhiteSpace($script:PrepareLogPath)) {
        for ($try = 1; $try -le 5; $try++) {
            try {
                [System.IO.File]::AppendAllText($script:PrepareLogPath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                break
            } catch {
                if ($try -eq 5) { break }
                Start-Sleep -Milliseconds 80
            }
        }
    }
    if ($Silent) {
        Write-Host $line
    }
}

function Invoke-ExcelAction {
    param([scriptblock]$Action)
    $lastError = $null
    for ($try = 1; $try -le 20; $try++) {
        try {
            return & $Action
        } catch {
            $lastError = $_
            Start-Sleep -Milliseconds 250
        }
    }
    throw $lastError
}

function Select-WorkbookPath {
    if ($Silent) { throw "Informe o caminho da planilha com -WorkbookPath." }
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

function Get-WorksheetByName {
    param([object]$Workbook, [string]$Name)
    foreach ($ws in @($Workbook.Worksheets)) {
        if ((Normalize-Text $ws.Name).ToUpperInvariant() -eq $Name.ToUpperInvariant()) {
            return $ws
        }
    }
    $available = @($Workbook.Worksheets | ForEach-Object { $_.Name }) -join ", "
    throw "A planilha '$($Workbook.Name)' nao tem a aba '$Name'. Abas encontradas: $available"
}

function Load-ClientMap {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Join-Path $PSScriptRoot "WechatClienteMap.json"
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Nao encontrei o arquivo de mapeamento:`n$Path"
    }
    $jsonObject = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $map = @{}
    foreach ($property in $jsonObject.PSObject.Properties) {
        $map[$property.Name] = [string]$property.Value
    }
    return $map
}

function Find-DiarioBlocks {
    param([object]$Sheet)

    $blocks = New-Object System.Collections.ArrayList
    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 8).End(-4162).Row

    for ($row = 1; $row -le $lastRow; $row++) {
        $label = (Normalize-Text $Sheet.Cells($row, 8).Value2).ToUpperInvariant()
        if ($label -ne "TOTAL") { continue }

        $client = Normalize-Text $Sheet.Cells($row, 6).Value2
        if ([string]::IsNullOrWhiteSpace($client)) { continue }

        $previousTotalRow = 0
        for ($prev = $row - 1; $prev -ge 1; $prev--) {
            if ((Normalize-Text $Sheet.Cells($prev, 8).Value2).ToUpperInvariant() -eq "TOTAL") {
                $previousTotalRow = $prev
                break
            }
        }

        $headerRow = 0
        for ($scan = $row - 1; $scan -gt $previousTotalRow; $scan--) {
            $f = (Normalize-Text $Sheet.Cells($scan, 6).Value2).ToUpperInvariant()
            $g = (Normalize-Text $Sheet.Cells($scan, 7).Value2).ToUpperInvariant()
            $h = (Normalize-Text $Sheet.Cells($scan, 8).Value2).ToUpperInvariant()
            if ($f -eq "DATA" -and ($g -like "HORARIO*" -or $g -like "HORÃRIO*") -and $h -eq "BANCO") {
                $headerRow = $scan
                break
            }
            if (($f -like "HORARIO*" -or $f -like "HORÃRIO*") -and $g -eq "BANCO" -and $h -like "TRANSFER*") {
                $headerRow = $scan
                break
            }
        }
        if ($headerRow -eq 0) { continue }

        $detailCount = [Math]::Max(0, $row - $headerRow - 1)
        $filledCount = 0
        for ($detail = $headerRow + 1; $detail -le $row - 1; $detail++) {
            $amount = Normalize-Text $Sheet.Cells($detail, 9).Value2
            $bank = Normalize-Text $Sheet.Cells($detail, 8).Value2
            if (-not [string]::IsNullOrWhiteSpace($amount) -or -not [string]::IsNullOrWhiteSpace($bank)) {
                $filledCount++
            }
        }

        if ($filledCount -le 0) { continue }

        [void]$blocks.Add([pscustomobject]@{
            Cliente = $client
            HeaderRow = $headerRow
            TotalRow = $row
            DetailCount = $detailCount
            FilledCount = $filledCount
        })
    }

    return $blocks
}

function Find-SaldoBlock {
    param([object]$Sheet, [string]$Client)

    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 1).End(-4162).Row
    $lastCol = $Sheet.Cells(1, $Sheet.Columns.Count).End(-4159).Column
    if ($lastCol -lt 1) { $lastCol = 60 }

    for ($row = 1; $row -le $lastRow; $row++) {
        for ($col = 1; $col -le $lastCol; $col++) {
            $value = Normalize-Text $Sheet.Cells($row, $col).Value2
            if ($value -ne $Client) { continue }

            $label = (Normalize-Text $Sheet.Cells($row, $col + 1).Value2).ToUpperInvariant()
            if ($label -notlike "SALDO ANTIGO*") { continue }

            return [pscustomobject]@{
                StartRow = $row
                EndRow = $row + 9
                StartCol = $col
                EndCol = $col + 2
            }
        }
    }

    return $null
}

function Build-SaldoBlockMap {
    param([object]$Sheet)

    Write-PrepareLog "Indexando aba Saldo Clientes inicio"
    $map = @{}
    $usedRange = $Sheet.UsedRange
    $values = $usedRange.Value2
    $baseRow = $usedRange.Row
    $baseCol = $usedRange.Column

    if ($null -eq $values) {
        Write-PrepareLog "Indexando aba Saldo Clientes fim - vazio"
        return $map
    }

    $rowCount = $values.GetLength(0)
    $colCount = $values.GetLength(1)

    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -lt $colCount; $c++) {
            $client = Normalize-Text $values[$r, $c]
            if ([string]::IsNullOrWhiteSpace($client)) { continue }

            $nextCol = $c + 1
            $label = (Normalize-Text $values[$r, $nextCol]).ToUpperInvariant()
            if ($label -notlike "SALDO ANTIGO*") { continue }

            if (-not $map.ContainsKey($client)) {
                $map[$client] = [pscustomobject]@{
                    StartRow = $baseRow + $r - 1
                    EndRow = $baseRow + $r + 8
                    StartCol = $baseCol + $c - 1
                    EndCol = $baseCol + $c + 1
                }
            }
        }
    }

    Write-PrepareLog "Indexando aba Saldo Clientes fim - $($map.Count) blocos"
    return $map
}

function ColumnName {
    param([int]$Column)
    $name = ""
    while ($Column -gt 0) {
        $mod = ($Column - 1) % 26
        $name = [char](65 + $mod) + $name
        $Column = [math]::Floor(($Column - $mod) / 26)
    }
    return $name
}

function Export-RangeImage {
    param(
        [object]$Sheet,
        [string]$Address,
        [string]$OutputPath
    )

    Write-PrepareLog "Exportando imagem: $Address -> $OutputPath"
    $range = $Sheet.Range($Address)
    $Sheet.Activate() | Out-Null
    $window = $Sheet.Application.ActiveWindow
    if ($null -ne $window) {
        $window.ScrollRow = [Math]::Max(1, $range.Row - 2)
        $window.ScrollColumn = [Math]::Max(1, $range.Column - 1)
    }
    $range.Select() | Out-Null
    [System.Windows.Forms.Clipboard]::Clear()
    Start-Sleep -Milliseconds 300
    Write-PrepareLog "CopyPicture inicio: $Address"
    Invoke-ExcelAction { $range.CopyPicture(1, 2) | Out-Null }
    Write-PrepareLog "CopyPicture fim: $Address"
    Start-Sleep -Milliseconds 500

    $chartObject = $Sheet.ChartObjects().Add($range.Left, $range.Top, $range.Width + 8, $range.Height + 8)
    try {
        $chart = $chartObject.Chart
        $chartObject.Activate() | Out-Null
        Start-Sleep -Milliseconds 150
        Write-PrepareLog "Chart paste inicio: $Address"
        Invoke-ExcelAction { $chart.Paste() | Out-Null }
        Write-PrepareLog "Chart paste fim: $Address"
        Start-Sleep -Milliseconds 150
        Write-PrepareLog "Chart export inicio: $OutputPath"
        Invoke-ExcelAction { $chart.Export($OutputPath, "PNG") | Out-Null }
        Write-PrepareLog "Chart export fim: $OutputPath"
    } finally {
        $chartObject.Delete() | Out-Null
    }
}

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Select-WorkbookPath
}

$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$clientMap = Load-ClientMap -Path $MapPath
$saldoClientesOpcional = @("VALEG")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$workbook = $null
$script:PrepareLogPath = ""
$excelProcessId = $null

try {
    try {
        $excelHwnd = [IntPtr]$excel.Hwnd
        $excelProcess = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq $excelHwnd } | Select-Object -First 1
        if ($null -ne $excelProcess) { $excelProcessId = $excelProcess.Id }
    } catch {}

    Show-Status "Abrindo planilha..."
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $diario = Get-WorksheetByName -Workbook $workbook -Name "Diario"
    $saldoClientes = Get-WorksheetByName -Workbook $workbook -Name "Saldo Clientes"

    $baseDir = Split-Path -Parent $WorkbookPath
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputDir = Join-Path $baseDir "WeChat_Envio_$stamp"
    New-Item -ItemType Directory -Path $outputDir | Out-Null
    $script:PrepareLogPath = Join-Path $outputDir "preparar_wechat_log.txt"
    Write-PrepareLog "Inicio"
    Write-PrepareLog "Planilha: $WorkbookPath"
    Write-PrepareLog "Pasta saida: $outputDir"

    Show-Status "Localizando blocos preenchidos..."
    $blocks = @(Find-DiarioBlocks -Sheet $diario)
    if ($blocks.Count -eq 0) {
        throw "Nao encontrei blocos preenchidos na aba Diario."
    }
    Write-PrepareLog "Blocos preenchidos no Diario: $($blocks.Count)"
    $saldoBlockMap = Build-SaldoBlockMap -Sheet $saldoClientes

    $queue = New-Object System.Collections.ArrayList
    $review = New-Object System.Collections.ArrayList
    $saldoReview = New-Object System.Collections.ArrayList

    $index = 0
    foreach ($block in $blocks) {
        $index++
        $client = $block.Cliente
        Show-Status "Gerando imagem do cliente $client ($index de $($blocks.Count))..." $index $blocks.Count
        Write-PrepareLog "Cliente $client ($index/$($blocks.Count)) inicio - Diario linhas $($block.HeaderRow):$($block.TotalRow)"

        $groupName = ""
        if ($clientMap.ContainsKey($client)) {
            $groupName = [string]$clientMap[$client]
        }

        $safeClient = ($client -replace '[\\/:*?"<>|]', '_')
        $imagePath = Join-Path $outputDir ("{0}_comprovantes.png" -f $safeClient)
        $saldoImagePath = Join-Path $outputDir ("{0}_saldo.png" -f $safeClient)
        $address = "E$($block.HeaderRow):I$($block.TotalRow)"
        Write-PrepareLog "Cliente $client - comprovantes inicio"
        Export-RangeImage -Sheet $diario -Address $address -OutputPath $imagePath
        Write-PrepareLog "Cliente $client - comprovantes fim"

        $saldoBlock = $null
        if ($saldoBlockMap.ContainsKey($client)) {
            $saldoBlock = $saldoBlockMap[$client]
        }
        $saldoStatus = "OK"
        if ($null -eq $saldoBlock) {
            if ($saldoClientesOpcional -contains $client) {
                $saldoStatus = "NAO_PRECISA"
            } else {
                $saldoStatus = "REVISAR_SALDO_CLIENTES"
            }
            $saldoImagePath = ""
        } else {
            $startCol = ColumnName $saldoBlock.StartCol
            $endCol = ColumnName $saldoBlock.EndCol
            $saldoAddress = "$startCol$($saldoBlock.StartRow):$endCol$($saldoBlock.EndRow)"
            Write-PrepareLog "Cliente $client - saldo inicio - $saldoAddress"
            Export-RangeImage -Sheet $saldoClientes -Address $saldoAddress -OutputPath $saldoImagePath
            Write-PrepareLog "Cliente $client - saldo fim"
        }

        $status = if ([string]::IsNullOrWhiteSpace($groupName)) { "REVISAR_MAPEAMENTO" } else { "OK" }
        $item = [pscustomobject]@{
            Ordem = $index
            Cliente = $client
            GrupoWeChat = $groupName
            Status = $status
            StatusSaldo = $saldoStatus
            Comprovantes = $block.FilledCount
            Imagem = $imagePath
            ImagemSaldo = $saldoImagePath
        }

        [void]$queue.Add($item)
        if ($status -ne "OK") { [void]$review.Add($item) }
        if ($saldoStatus -ne "OK") { [void]$saldoReview.Add($item) }
        Write-PrepareLog "Cliente $client fim - status=$status statusSaldo=$saldoStatus"
    }

    $queuePath = Join-Path $outputDir "fila_envio_wechat.csv"
    $queue | Export-Csv -LiteralPath $queuePath -NoTypeInformation -Encoding UTF8

    $previewPath = Join-Path $outputDir "Conferir_Imagens.html"
    $htmlRows = foreach ($item in $queue) {
        $imageName = [System.IO.Path]::GetFileName($item.Imagem)
        $saldoImageName = if ([string]::IsNullOrWhiteSpace($item.ImagemSaldo)) { "" } else { [System.IO.Path]::GetFileName($item.ImagemSaldo) }
        $group = [System.Net.WebUtility]::HtmlEncode($item.GrupoWeChat)
        $client = [System.Net.WebUtility]::HtmlEncode($item.Cliente)
        $status = [System.Net.WebUtility]::HtmlEncode($item.Status)
        $statusSaldo = [System.Net.WebUtility]::HtmlEncode($item.StatusSaldo)
        $saldoHtml = if ([string]::IsNullOrWhiteSpace($saldoImageName)) { "<p><strong>Sem imagem de Saldo Clientes.</strong></p>" } else { "<img src=""$saldoImageName"" alt=""Saldo Cliente $client"">" }
        @"
<section class="item">
  <div class="meta">
    <strong>Cliente $client</strong>
    <span>Grupo: $group</span>
    <span>Status: $status</span>
    <span>Saldo Clientes: $statusSaldo</span>
  </div>
  <h2>Comprovantes Diario</h2>
  <img src="$imageName" alt="Cliente $client">
  <h2>Saldo Clientes</h2>
  $saldoHtml
</section>
"@
    }

    @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>Conferir Imagens WeChat</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; background: #f4f4f4; color: #111; }
    h1 { font-size: 22px; margin: 0 0 18px; }
    .item { background: white; border: 1px solid #ccc; margin: 0 0 24px; padding: 14px; }
    .meta { display: flex; gap: 18px; flex-wrap: wrap; margin-bottom: 12px; font-size: 15px; }
    img { max-width: 100%; height: auto; border: 1px solid #222; display: block; }
    h2 { font-size: 16px; margin: 14px 0 8px; }
  </style>
</head>
<body>
  <h1>Conferir Imagens WeChat</h1>
  $($htmlRows -join "`n")
</body>
</html>
"@ | Set-Content -LiteralPath $previewPath -Encoding UTF8

    $summaryPath = Join-Path $outputDir "LEIA_ME.txt"
    @(
        "Pacote de envio WeChat",
        "Planilha: $WorkbookPath",
        "Pasta: $outputDir",
        "Blocos encontrados: $($blocks.Count)",
        "Pendentes de mapeamento: $($review.Count)",
        "Sem bloco em Saldo Clientes: $($saldoReview.Count)",
        "",
        "Abra Conferir_Imagens.html para ver todas as imagens no navegador.",
        "Use o arquivo fila_envio_wechat.csv para conferir cliente -> grupo -> imagem.",
        "Depois rode Enviar WeChat Seguro.bat e selecione esse CSV."
    ) | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    Write-PrepareLog "Arquivos finais gravados"

    Show-Status "Concluido."
    if (-not $Silent) {
        Write-Progress -Activity "Preparando envio WeChat" -Completed
        if (-not $NoFinalMessage) {
            [System.Windows.Forms.MessageBox]::Show("Pacote WeChat criado em:`n$outputDir`n`nPendentes de mapeamento: $($review.Count)`nSem bloco em Saldo Clientes: $($saldoReview.Count)`n`nPara conferir, abra Conferir_Imagens.html.", "Preparar WeChat") | Out-Null
        }
    } else {
        Write-Host "Pacote WeChat criado em: $outputDir"
        Write-Host "Fila: $queuePath"
        Write-Host "Conferencia: $previewPath"
    }
    Write-PrepareLog "Concluido"
} catch {
    Write-PrepareLog "ERRO: $($_.Exception.Message)"
    if (-not $Silent -and -not $NoFinalMessage) {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro ao preparar WeChat") | Out-Null
    }
    throw
} finally {
    try { $excel.DisplayAlerts = $true } catch {}
    try {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
    } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $excelProcessId) {
        Start-Sleep -Milliseconds 800
        $stillRunning = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($null -ne $stillRunning) {
            Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
        }
    }
}

