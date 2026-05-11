param(
    [string]$WorkbookPath = "",
    [switch]$Silent
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

$Pepper = [char]::ConvertFromUtf32(0x1FAD1)
$GroupName = "2026 " + ($Pepper * 6)

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

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function To-Number {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return [double]$Value }
    $text = (Normalize-Text $Value).Replace(".", "").Replace(",", ".")
    $parsed = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Format-MoneyBr {
    param([double]$Value)
    return $Value.ToString("N2", [Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
}

function Format-PlainNumber {
    param([double]$Value)
    return ([int64][Math]::Round($Value, 0)).ToString([Globalization.CultureInfo]::InvariantCulture)
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

function Write-Log {
    param([string]$Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss.fff} - {1}" -f (Get-Date), $Message
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        [System.IO.File]::AppendAllText($script:LogPath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    if ($Silent) { Write-Host $line }
}

function Export-RangeImage {
    param(
        [object]$Sheet,
        [string]$Address,
        [string]$OutputPath
    )

    Write-Log "Exportando $Address -> $OutputPath"
    $range = $Sheet.Range($Address)
    $Sheet.Activate() | Out-Null
    $window = $Sheet.Application.ActiveWindow
    if ($null -ne $window) {
        $window.ScrollRow = [Math]::Max(1, $range.Row - 2)
        $window.ScrollColumn = [Math]::Max(1, $range.Column - 1)
    }
    $range.Select() | Out-Null
    [System.Windows.Forms.Clipboard]::Clear()
    Start-Sleep -Milliseconds 250
    Invoke-ExcelAction { $range.CopyPicture(1, 2) | Out-Null }
    Start-Sleep -Milliseconds 350

    $chartObject = $Sheet.ChartObjects().Add($range.Left, $range.Top, $range.Width + 8, $range.Height + 8)
    try {
        $chart = $chartObject.Chart
        $chartObject.Activate() | Out-Null
        Start-Sleep -Milliseconds 120
        Invoke-ExcelAction { $chart.Paste() | Out-Null }
        Start-Sleep -Milliseconds 120
        Invoke-ExcelAction { $chart.Export($OutputPath, "PNG") | Out-Null }
    } finally {
        $chartObject.Delete() | Out-Null
    }
}

function Safe-Name {
    param([string]$Text)
    return ($Text -replace '[\\/:*?"<>|\[\]]', '_')
}

function Get-UniqueWorksheetName {
    param([object]$Workbook, [string]$BaseName)

    $clean = Safe-Name $BaseName
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "RELATORIO" }

    for ($n = 0; $n -lt 1000; $n++) {
        $suffix = if ($n -eq 0) { "" } else { "_$n" }
        $maxBase = 31 - $suffix.Length
        $candidate = $clean.Substring(0, [Math]::Min($maxBase, $clean.Length)) + $suffix
        $exists = $false
        foreach ($ws in @($Workbook.Worksheets)) {
            if ($ws.Name -eq $candidate) {
                $exists = $true
                break
            }
        }
        if (-not $exists) { return $candidate }
    }

    throw "Nao consegui criar um nome unico de aba para '$BaseName'."
}

function Color-From-Hex {
    param([string]$Hex)
    $clean = $Hex.TrimStart("#")
    $r = [Convert]::ToInt32($clean.Substring(0,2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2,2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4,2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Add-TableImage {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [object[]]$Rows,
        [string]$OutputPath,
        [string]$BankColor = "#FFFFFF",
        [string]$BankTextColor = "#000000",
        [switch]$OutReport,
        [switch]$NoTotal,
        [object]$TotalOverride = $null
    )

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = Get-UniqueWorksheetName -Workbook $Workbook -BaseName $SheetName
    $sheet.Cells.Font.Name = "Arial"
    $sheet.Cells.HorizontalAlignment = -4108
    $sheet.Cells.VerticalAlignment = -4108
    $compact = (-not $OutReport -and $Rows.Count -gt 250)

    if ($OutReport) {
        $headers = @("Data", "CLIENT", "OUT", "STATUS", "MOTIVO")
    } else {
        $headers = @("CLIENTE", "DATA", "HORA", "BANCO", "VALOR")
    }

    for ($c = 1; $c -le $headers.Count; $c++) {
        $sheet.Cells(1, $c).Value = $headers[$c - 1]
    }

    $header = $sheet.Range("A1:E1")
    $header.Interior.Color = Color-From-Hex "#6A00FF"
    $header.Font.Color = Color-From-Hex "#FFFFFF"
    $header.Font.Bold = $true
    $header.Font.Size = if ($compact) { 12 } else { 18 }
    $header.RowHeight = if ($compact) { 16 } else { 28 }

    $total = 0.0
    $rowIndex = 2
    foreach ($item in $Rows) {
        if ($OutReport) {
            $sheet.Cells($rowIndex, 1).Value2 = $item.C1
            $sheet.Cells($rowIndex, 2).Value2 = $item.C2
            $sheet.Cells($rowIndex, 3).Value2 = Format-PlainNumber $item.C3
            $sheet.Cells($rowIndex, 4).Value2 = $item.C4
            $sheet.Cells($rowIndex, 5).Value2 = $item.C5
            $num = To-Number $item.C3
        } else {
            $sheet.Cells($rowIndex, 1).Value2 = $item.C1
            $sheet.Cells($rowIndex, 2).Value2 = $item.C2
            $sheet.Cells($rowIndex, 3).Value2 = $item.C3
            $sheet.Cells($rowIndex, 4).Value2 = $item.C4
            $sheet.Cells($rowIndex, 5).Value2 = Format-PlainNumber $item.C5
            $num = To-Number $item.C5
        }
        if ($null -ne $num) { $total += $num }
        $rowIndex++
    }

    $lastDataRow = [Math]::Max(1, $rowIndex - 1)
    $totalRow = if ($NoTotal) { $lastDataRow } else { $rowIndex }
    if (-not $NoTotal) {
        $displayTotal = if ($null -ne $TotalOverride) { [double]$TotalOverride } else { $total }
        if ($OutReport) {
            $sheet.Cells($totalRow, 3).Value2 = Format-PlainNumber $displayTotal
            $sheet.Cells($totalRow, 3).Font.Bold = $true
            $sheet.Cells($totalRow, 3).Interior.Color = Color-From-Hex "#D9D9D9"
        } else {
            $sheet.Cells($totalRow, 4).Value2 = "R$"
            $sheet.Cells($totalRow, 5).Value2 = Format-MoneyBr $displayTotal
            $sheet.Range("D${totalRow}:E${totalRow}").Interior.Color = Color-From-Hex "#D9D9D9"
            $sheet.Range("D${totalRow}:E${totalRow}").Font.Bold = $true
            $sheet.Range("D${totalRow}:E${totalRow}").Font.Size = if ($compact) { 12 } else { 20 }
            $sheet.Cells($totalRow, 5).NumberFormatLocal = "#.##0,00"
        }
    }

    if ($lastDataRow -ge 2) {
        $body = $sheet.Range("A2:E$lastDataRow")
        $body.Font.Size = if ($compact) { 8 } else { 16 }
        $body.Borders.LineStyle = 1
        $body.RowHeight = if ($compact) { 11 } else { 26 }
        if ($OutReport) {
            $sheet.Range("C2:C$lastDataRow").Font.Color = Color-From-Hex "#FF0000"
            $sheet.Range("D2:D$lastDataRow").Interior.Color = Color-From-Hex "#00B050"
            $sheet.Range("D2:D$lastDataRow").Font.Color = Color-From-Hex "#FFFFFF"
            $sheet.Range("D2:D$lastDataRow").Font.Bold = $true
        } else {
            $sheet.Range("D2:D$lastDataRow").Interior.Color = Color-From-Hex $BankColor
            $sheet.Range("D2:D$lastDataRow").Font.Color = Color-From-Hex $BankTextColor
            $sheet.Range("D2:D$lastDataRow").Font.Bold = $true
        }
    }

    $sheet.Range("A1:E$totalRow").Borders.LineStyle = 1
    if ($OutReport) {
        $sheet.Columns.Item(1).ColumnWidth = 15
        $sheet.Columns.Item(2).ColumnWidth = 12
        $sheet.Columns.Item(3).ColumnWidth = 14
        $sheet.Columns.Item(4).ColumnWidth = 14
        $sheet.Columns.Item(5).ColumnWidth = 18
    } else {
        $sheet.Columns.Item(1).ColumnWidth = 14
        $sheet.Columns.Item(2).ColumnWidth = 16
        $sheet.Columns.Item(3).ColumnWidth = 13
        $sheet.Columns.Item(4).ColumnWidth = 20
        $sheet.Columns.Item(5).ColumnWidth = 18
    }
    if (-not $NoTotal) {
        $sheet.Cells($totalRow, 1).RowHeight = if ($compact) { 18 } else { 32 }
    }

    Export-RangeImage -Sheet $sheet -Address "A1:E$totalRow" -OutputPath $OutputPath
}

function Read-InRows {
    param([object]$Sheet)
    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 1).End(-4162).Row
    $rows = New-Object System.Collections.ArrayList
    for ($r = 2; $r -le $lastRow; $r++) {
        $client = Normalize-Text $Sheet.Cells($r, 1).Text
        $date = Normalize-Text $Sheet.Cells($r, 2).Text
        $hour = Normalize-Text $Sheet.Cells($r, 3).Text
        $bank = Normalize-Text $Sheet.Cells($r, 4).Text
        $value = To-Number $Sheet.Cells($r, 5).Value2
        if ([string]::IsNullOrWhiteSpace($client) -or [string]::IsNullOrWhiteSpace($bank) -or $null -eq $value) { continue }
        [void]$rows.Add([pscustomobject]@{ C1=$client; C2=$date; C3=$hour; C4=$bank; C5=$value })
    }
    return $rows
}

function Read-OutRows {
    param([object]$Sheet)
    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 1).End(-4162).Row
    $rows = New-Object System.Collections.ArrayList
    for ($r = 2; $r -le $lastRow; $r++) {
        $date = Normalize-Text $Sheet.Cells($r, 1).Text
        $client = Normalize-Text $Sheet.Cells($r, 2).Text
        $out = To-Number $Sheet.Cells($r, 3).Value2
        $status = Normalize-Text $Sheet.Cells($r, 4).Text
        $reason = Normalize-Text $Sheet.Cells($r, 5).Text
        if ([string]::IsNullOrWhiteSpace($client) -or $null -eq $out) { continue }
        [void]$rows.Add([pscustomobject]@{ C1=$date; C2=$client; C3=$out; C4=$status; C5=$reason })
    }
    return $rows
}

function Prepare-OutSheetForExport {
    param([object]$Sheet)

    $lastByDate = $Sheet.Cells($Sheet.Rows.Count, 1).End(-4162).Row
    $lastByValue = $Sheet.Cells($Sheet.Rows.Count, 3).End(-4162).Row
    $lastRow = [Math]::Max($lastByDate, $lastByValue)
    if ($lastRow -lt 1) { $lastRow = 1 }

    $dataLastRow = 1
    $total = 0.0
    for ($r = 2; $r -le $lastRow; $r++) {
        $client = Normalize-Text $Sheet.Cells($r, 2).Text
        $value = To-Number $Sheet.Cells($r, 3).Value2
        if (-not [string]::IsNullOrWhiteSpace($client) -and $null -ne $value) {
            $total += $value
            $dataLastRow = $r
        }
    }

    $totalRow = $dataLastRow + 1
    $Sheet.Range("A${totalRow}:E${totalRow}").ClearContents() | Out-Null
    $Sheet.Cells($totalRow, 3).Value2 = Format-PlainNumber $total

    $totalCell = $Sheet.Cells($totalRow, 3)
    $totalCell.Interior.Color = Color-From-Hex "#D9D9D9"
    $totalCell.Font.Bold = $true
    $totalCell.Font.Color = Color-From-Hex "#000000"
    $totalCell.HorizontalAlignment = -4108
    $totalCell.VerticalAlignment = -4108

    $totalRange = $Sheet.Range("A${totalRow}:E${totalRow}")
    $totalRange.Borders.LineStyle = -4142

    if ($dataLastRow -ge 2) {
        $dataRange = $Sheet.Range("A1:E$totalRow")
        $dataRange.Borders.LineStyle = -4142
    }

    return $totalRow
}

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Select-WorkbookPath
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$baseDir = Split-Path -Parent $WorkbookPath
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Join-Path $baseDir "Resumo_Grupo_Envio_$stamp"
New-Item -ItemType Directory -Path $outputDir | Out-Null
$script:LogPath = Join-Path $outputDir "preparar_resumo_grupo_log.txt"
Write-Log "Inicio"
Write-Log "Planilha: $WorkbookPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$sourceWorkbook = $null
$reportWorkbook = $null

try {
    $sourceWorkbook = $excel.Workbooks.Open($WorkbookPath)
    $inSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "IN"
    $outSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "OUT"
    $resumoSaldoSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "RESUMO SALDO"

    $inRows = Read-InRows -Sheet $inSheet
    $outRows = Read-OutRows -Sheet $outSheet

    $reportWorkbook = $excel.Workbooks.Add()
    while ($reportWorkbook.Worksheets.Count -gt 1) {
        $reportWorkbook.Worksheets.Item(1).Delete()
    }
    $reportWorkbook.Worksheets.Item(1).Name = "TEMP"

    $queue = New-Object System.Collections.ArrayList
    $order = 1

    $bankColors = @{
        "DIAMOND" = "#12AEE3"
        "CLEEND" = "#50E87A"
        "AMD" = "#FFFFFF"
    }
    $bankTextColors = @{
        "DIAMOND" = "#FFFFFF"
        "CLEEND" = "#000000"
        "AMD" = "#000000"
    }
    foreach ($bank in @("DIAMOND", "CLEEND", "AMD")) {
        Write-Log "Gerando banco $bank"
        $rows = @($inRows | Where-Object { (Normalize-Text $_.C4).ToUpperInvariant() -eq $bank })
        $path = Join-Path $outputDir ("{0:D2}_{1}.png" -f $order, $bank)
        Add-TableImage -Workbook $reportWorkbook -SheetName $bank -Rows $rows -OutputPath $path -BankColor $bankColors[$bank] -BankTextColor $bankTextColors[$bank]
        [void]$queue.Add([pscustomobject]@{ Ordem=$order; GrupoWeChat=$GroupName; Tipo=$bank; Imagem=$path })
        $order++
    }

    $moneyRows = @($inRows | Where-Object { (Normalize-Text $_.C4).ToUpperInvariant() -in @("CHEQUE", "DINHEIRO") })
    if ($moneyRows.Count -gt 0) {
        Write-Log "Gerando CHEQUE_DINHEIRO"
        $moneyPath = Join-Path $outputDir ("{0:D2}_CHEQUE_DINHEIRO.png" -f $order)
        Add-TableImage -Workbook $reportWorkbook -SheetName "CHEQUE_DINHEIRO" -Rows $moneyRows -OutputPath $moneyPath
        [void]$queue.Add([pscustomobject]@{ Ordem=$order; GrupoWeChat=$GroupName; Tipo="CHEQUE_DINHEIRO"; Imagem=$moneyPath })
        $order++
    } else {
        Write-Log "Pulando CHEQUE_DINHEIRO: sem linhas no dia"
    }

    Write-Log "Gerando OUT"
    $outPath = Join-Path $outputDir ("{0:D2}_OUT.png" -f $order)
    $outLastRow = Prepare-OutSheetForExport -Sheet $outSheet
    if ($outLastRow -lt 1) { $outLastRow = 1 }
    if ($outSheet.Columns.Item(1).ColumnWidth -lt 12) { $outSheet.Columns.Item(1).ColumnWidth = 12 }
    if ($outSheet.Columns.Item(2).ColumnWidth -lt 10) { $outSheet.Columns.Item(2).ColumnWidth = 10 }
    if ($outSheet.Columns.Item(3).ColumnWidth -lt 12) { $outSheet.Columns.Item(3).ColumnWidth = 12 }
    if ($outSheet.Columns.Item(4).ColumnWidth -lt 10) { $outSheet.Columns.Item(4).ColumnWidth = 10 }
    if ($outSheet.Columns.Item(5).ColumnWidth -lt 16) { $outSheet.Columns.Item(5).ColumnWidth = 16 }
    Export-RangeImage -Sheet $outSheet -Address "A1:E$outLastRow" -OutputPath $outPath
    [void]$queue.Add([pscustomobject]@{ Ordem=$order; GrupoWeChat=$GroupName; Tipo="OUT"; Imagem=$outPath })
    $order++

    Write-Log "Gerando RESUMO_IN"
    $summaryPath = Join-Path $outputDir ("{0:D2}_RESUMO_IN.png" -f $order)
    Export-RangeImage -Sheet $inSheet -Address "H3:L30" -OutputPath $summaryPath
    [void]$queue.Add([pscustomobject]@{ Ordem=$order; GrupoWeChat=$GroupName; Tipo="RESUMO_IN"; Imagem=$summaryPath })
    $order++

    Write-Log "Gerando CLIENT_TOTAL"
    $clientTotalPath = Join-Path $outputDir ("{0:D2}_CLIENT_TOTAL.png" -f $order)
    Export-RangeImage -Sheet $resumoSaldoSheet -Address "A1:B32" -OutputPath $clientTotalPath
    [void]$queue.Add([pscustomobject]@{ Ordem=$order; GrupoWeChat=$GroupName; Tipo="CLIENT_TOTAL"; Imagem=$clientTotalPath })

    $queuePath = Join-Path $outputDir "fila_resumo_grupo_wechat.csv"
    $queue | Export-Csv -LiteralPath $queuePath -NoTypeInformation -Encoding UTF8

    $htmlRows = foreach ($item in $queue) {
        $imageName = [System.IO.Path]::GetFileName($item.Imagem)
        $tipo = [System.Net.WebUtility]::HtmlEncode($item.Tipo)
        @"
<section class="item">
  <h2>$($item.Ordem) - $tipo</h2>
  <img src="$imageName" alt="$tipo">
</section>
"@
    }
    $previewPath = Join-Path $outputDir "Conferir_Resumo_Grupo.html"
    @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>Conferir Resumo Grupo</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; background: #f4f4f4; color: #111; }
    .item { background: white; border: 1px solid #ccc; margin: 0 0 24px; padding: 14px; }
    h1 { font-size: 22px; }
    h2 { font-size: 16px; margin: 0 0 10px; }
    img { max-width: 100%; height: auto; border: 1px solid #222; display: block; }
  </style>
</head>
<body>
  <h1>Resumo para $GroupName</h1>
  $($htmlRows -join "`n")
</body>
</html>
"@ | Set-Content -LiteralPath $previewPath -Encoding UTF8

    @(
        "Resumo Grupo WeChat",
        "Planilha: $WorkbookPath",
        "Grupo: $GroupName",
        "Pasta: $outputDir",
        "Fila: $queuePath",
        "Conferencia: $previewPath"
    ) | Set-Content -LiteralPath (Join-Path $outputDir "LEIA_ME.txt") -Encoding UTF8

    Write-Log "Concluido"
    if ($Silent) {
        Write-Host "Resumo criado em: $outputDir"
        Write-Host "Fila: $queuePath"
        Write-Host "Conferencia: $previewPath"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Resumo criado em:`n$outputDir`n`nAbra Conferir_Resumo_Grupo.html para conferir.", "Resumo Grupo WeChat") | Out-Null
    }
} catch {
    Write-Log "ERRO: $($_.Exception.Message)"
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro resumo grupo") | Out-Null
    }
    throw
} finally {
    try { if ($null -ne $reportWorkbook) { $reportWorkbook.Close($false) | Out-Null } } catch {}
    try { if ($null -ne $sourceWorkbook) { $sourceWorkbook.Close($false) | Out-Null } } catch {}
    try { $excel.DisplayAlerts = $true } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
