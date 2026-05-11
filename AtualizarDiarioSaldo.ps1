param(
    [string]$WorkbookPath = "",
    [switch]$NoSave,
    [switch]$Silent,
    [switch]$PickFile,
    [switch]$CloseAfterSave,
    [switch]$NoFinalMessage
)

$ErrorActionPreference = "Stop"
if (-not $Silent) {
    Add-Type -AssemblyName System.Windows.Forms
}

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Is-Blank {
    param([object]$Value)
    return [string]::IsNullOrWhiteSpace((Normalize-Text $Value))
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

function To-ExcelValue {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return [double]$Value }
    return [string](Normalize-Text $Value)
}

function To-Scalar {
    param([object]$Value)
    if ($Value -is [array]) {
        if ($Value.Count -eq 0) { return "" }
        return $Value[$Value.Count - 1]
    }
    return $Value
}

function Invoke-ExcelAction {
    param([scriptblock]$Action)
    $lastError = $null
    for ($try = 1; $try -le 20; $try++) {
        try {
            return & $Action
        } catch {
            $lastError = $_
            if ($_.Exception.Message -notlike "*0x800AC472*") {
                throw
            }
            Start-Sleep -Milliseconds 250
        }
    }
    throw $lastError
}

function Show-Status {
    param(
        [string]$Message,
        [int]$Current = 0,
        [int]$Total = 0
    )

    if ($Silent) { return }

    if ($Total -gt 0) {
        $percent = [Math]::Min(100, [Math]::Max(0, [int](($Current / $Total) * 100)))
        Write-Progress -Activity "Atualizando Diario" -Status $Message -PercentComplete $percent
        Write-Host ("[{0,3}%] {1}" -f $percent, $Message)
    } else {
        Write-Progress -Activity "Atualizando Diario" -Status $Message
        Write-Host "[...] $Message"
    }
}

function Add-Log {
    param([string]$Message)
    $script:AutoDiarioLog.Add(("{0:HH:mm:ss} - {1}" -f (Get-Date), $Message)) | Out-Null
}

function Get-RunningExcel {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        [void]$app.Workbooks.Count
        return $app
    } catch {
        return $null
    }
}

function Select-WorkbookPath {
    if ($Silent) {
        throw "Abra a planilha 'Saldo cliente' no Excel ou informe o caminho com -WorkbookPath."
    }

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

function Get-Workbook {
    param([object]$Excel, [string]$Path)

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        $resolved = (Resolve-Path -LiteralPath $Path).Path
        foreach ($wb in @($Excel.Workbooks)) {
            if ($wb.FullName -eq $resolved) {
                $script:WorkbookOpenedByScript = $false
                return $wb
            }
        }
        $workbooks = $Excel.Workbooks
        $opened = $workbooks.Open($resolved)
        if ($null -ne $opened) {
            $script:WorkbookOpenedByScript = $true
            return $opened
        }

        foreach ($wb in @($Excel.Workbooks)) {
            if ($wb.FullName -eq $resolved -or $wb.Name -eq [IO.Path]::GetFileName($resolved)) {
                $script:WorkbookOpenedByScript = $false
                return $wb
            }
        }

        throw "Nao consegui abrir a planilha:`n$resolved"
    }

    if ($Excel.Workbooks.Count -gt 0) {
        $active = $Excel.ActiveWorkbook
        if ($null -ne $active -and $active.Name -like "*Saldo*cliente*") {
            $script:WorkbookOpenedByScript = $false
            return $active
        }

        foreach ($wb in @($Excel.Workbooks)) {
            if ($wb.Name -like "*Saldo*cliente*") {
                $script:WorkbookOpenedByScript = $false
                return $wb
            }
        }
    }

    $selectedPath = Select-WorkbookPath
    $resolved = (Resolve-Path -LiteralPath $selectedPath).Path
    $workbooks = $Excel.Workbooks
    $opened = $workbooks.Open($resolved)
    if ($null -ne $opened) {
        $script:WorkbookOpenedByScript = $true
        return $opened
    }

    foreach ($wb in @($Excel.Workbooks)) {
        if ($wb.FullName -eq $resolved -or $wb.Name -eq [IO.Path]::GetFileName($resolved)) {
            $script:WorkbookOpenedByScript = $false
            return $wb
        }
    }

    throw "Nao consegui abrir a planilha:`n$resolved"
}

function Get-WorksheetByName {
    param([object]$Workbook, [string]$Name)

    if ($null -eq $Workbook) {
        throw "Nao consegui identificar a planilha do Excel. Tente arrastar o arquivo do dia em cima do .bat."
    }

    foreach ($ws in @($Workbook.Worksheets)) {
        if ((Normalize-Text $ws.Name).ToUpperInvariant() -eq $Name.ToUpperInvariant()) {
            return $ws
        }
    }

    $available = @($Workbook.Worksheets | ForEach-Object { $_.Name }) -join ", "
    throw "A planilha '$($Workbook.Name)' nao tem a aba '$Name'. Abas encontradas: $available"
}

function Find-Block {
    param([object]$Sheet, [string]$Client)

    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 8).End(-4162).Row
    for ($row = 1; $row -le $lastRow; $row++) {
        $totalLabel = (Normalize-Text $Sheet.Cells($row, 8).Value2).ToUpperInvariant()
        $clientAtTotal = Normalize-Text $Sheet.Cells($row, 6).Value2
        if ($totalLabel -eq "TOTAL" -and $clientAtTotal -eq $Client) {
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

            if ($headerRow -eq 0) {
                throw "Achei o TOTAL do cliente '$Client' na linha $row, mas nao achei o cabecalho acima dele."
            }

            return [pscustomobject]@{
                HeaderRow = $headerRow
                TotalRow = $row
            }
        }
    }

    return $null
}

function Append-Block {
    param([object]$Sheet, [string]$Client)

    $lastRow = $Sheet.Cells($Sheet.Rows.Count, 8).End(-4162).Row
    $headerRow = $lastRow + 3
    $totalRow = $headerRow + 1

    $Sheet.Cells($headerRow, 6).Value2 = "DATA"
    $Sheet.Cells($headerRow, 7).Value2 = "HORARIO"
    $Sheet.Cells($headerRow, 8).Value2 = "BANCO"
    $Sheet.Cells($headerRow, 9).Value2 = "TRANSFERENCIA"
    $Sheet.Cells($totalRow, 6).Value2 = $Client
    $Sheet.Cells($totalRow, 8).Value2 = "TOTAL"
    $Sheet.Cells($totalRow, 9).Formula = "=SUM(I$($headerRow + 1):I$($totalRow - 1))"

    return [pscustomobject]@{
        HeaderRow = $headerRow
        TotalRow = $totalRow
    }
}

function Resize-Block {
    param([object]$Sheet, [pscustomobject]$Block, [int]$DesiredRows)

    $currentRows = [Math]::Max(0, $Block.TotalRow - $Block.HeaderRow - 1)
    if ($currentRows -gt 0) {
        $allDetailRows = "$(($Block.HeaderRow + 1)):$($Block.TotalRow - 1)"
        Invoke-ExcelAction { $Sheet.Rows($allDetailRows).Hidden = $false }
    }

    if ($DesiredRows -gt $currentRows) {
        $toInsert = $DesiredRows - $currentRows
        $insertAt = $Block.TotalRow
        $rangeAddress = "${insertAt}:$($insertAt + $toInsert - 1)"
        $range = $Sheet.Rows($rangeAddress)
        Invoke-ExcelAction { $range.Insert() | Out-Null }

        $copyFrom = if ($currentRows -gt 0) { $insertAt - 1 } else { $Block.HeaderRow }
        Invoke-ExcelAction { $Sheet.Rows($copyFrom).Copy() }
        Invoke-ExcelAction { $Sheet.Rows($rangeAddress).PasteSpecial(-4122) }
        $Sheet.Application.CutCopyMode = $false

        $Block.TotalRow += $toInsert
        $currentRows = [Math]::Max(0, $Block.TotalRow - $Block.HeaderRow - 1)
    }

    if ($currentRows -gt $DesiredRows) {
        $hideFrom = $Block.HeaderRow + $DesiredRows + 1
        $hideTo = $Block.TotalRow - 1
        if ($hideFrom -le $hideTo) {
            Invoke-ExcelAction { $Sheet.Rows("${hideFrom}:${hideTo}").Hidden = $true }
        }
    }

    return $Block
}

function Update-Block {
    param([object]$Sheet, [string]$Client, [System.Collections.IList]$Rows)

    $script:AutoDiarioStage = "procurando bloco"
    $block = Find-Block -Sheet $Sheet -Client $Client
    if ($null -eq $block) {
        $script:AutoDiarioStage = "criando bloco"
        $block = Append-Block -Sheet $Sheet -Client $Client
    }

    $script:AutoDiarioStage = "ajustando quantidade de linhas"
    $block = Resize-Block -Sheet $Sheet -Block $block -DesiredRows $Rows.Count
    $startRow = $block.HeaderRow + 1
    $endRow = $block.TotalRow - 1

    if ($endRow -ge $startRow) {
        $script:AutoDiarioStage = "limpando linhas antigas"
        $detailAddress = "E${startRow}:I${endRow}"
        Invoke-ExcelAction { $Sheet.Range($detailAddress).ClearContents() | Out-Null }
    }

    if ($Rows.Count -gt 0) {
        $script:AutoDiarioStage = "montando dados do cliente $Client"
        $values = [Array]::CreateInstance([object], @($Rows.Count, 5), @(1, 1))
        for ($i = 0; $i -lt $Rows.Count; $i++) {
            $item = $Rows[$i]
            $targetIndex = $i + 1
            $values[$targetIndex, 1] = [double]($i + 1)
            $values[$targetIndex, 2] = To-Scalar (To-ExcelValue $item[0])
            $values[$targetIndex, 3] = To-Scalar (To-ExcelValue $item[1])
            $values[$targetIndex, 4] = [string](To-Scalar $item[2])
            $values[$targetIndex, 5] = To-Scalar (To-ExcelValue $item[3])
        }
        $script:AutoDiarioStage = "gravando comprovantes do cliente $Client"
        $writeEndRow = $startRow + $Rows.Count - 1
        $writeAddress = "E${startRow}:I${writeEndRow}"
        if ($Rows.Count -eq 1) {
            $item = $Rows[0]
            $singleData = To-Scalar (To-ExcelValue $item[0])
            $singleHora = To-Scalar (To-ExcelValue $item[1])
            $singleBanco = [string](To-Scalar $item[2])
            $singleValor = To-Scalar (To-ExcelValue $item[3])
            Invoke-ExcelAction { $Sheet.Cells($startRow, 5).Value = 1 }
            Invoke-ExcelAction { $Sheet.Cells($startRow, 6).Value = $singleData }
            Invoke-ExcelAction { $Sheet.Cells($startRow, 7).Value = $singleHora }
            Invoke-ExcelAction { $Sheet.Cells($startRow, 8).Value = $singleBanco }
            Invoke-ExcelAction { $Sheet.Cells($startRow, 9).Value = $singleValor }
        } else {
            Invoke-ExcelAction { $Sheet.Range($writeAddress).Value = $values }
        }
    }

    $script:AutoDiarioStage = "gravando total"
    $totalFormula = if ($endRow -ge $startRow) { "=SUM(I${startRow}:I${endRow})" } else { "=0" }
    Invoke-ExcelAction { $Sheet.Cells($block.TotalRow, 6).Value = $Client }
    Invoke-ExcelAction { $Sheet.Cells($block.TotalRow, 8).Value = "TOTAL" }
    Invoke-ExcelAction { $Sheet.Cells($block.TotalRow, 9).Formula = $totalFormula }
}

$script:AutoDiarioLog = New-Object System.Collections.ArrayList

if ($PickFile -and [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Select-WorkbookPath
}

$excel = $null
$createdExcel = $false
$script:WorkbookOpenedByScript = $false
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $excel = Get-RunningExcel
} else {
    $runningExcel = Get-RunningExcel
    if ($null -ne $runningExcel) {
        try {
            $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
            foreach ($wb in @($runningExcel.Workbooks)) {
                if ($wb.FullName -eq $resolvedWorkbookPath) {
                    $excel = $runningExcel
                    break
                }
            }
        } catch {}
    }
}
if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $createdExcel = $true
}

try {
    Show-Status "Abrindo planilha..."
    $workbook = Get-Workbook -Excel $excel -Path $WorkbookPath
    Add-Log "Arquivo: $($workbook.FullName)"
    Show-Status "Lendo abas IN e Diario..."
    $sheetIn = Get-WorksheetByName -Workbook $workbook -Name "IN"
    $sheetDiario = Get-WorksheetByName -Workbook $workbook -Name "Diario"

    if (-not [string]::IsNullOrWhiteSpace($workbook.Path)) {
        Show-Status "Criando backup antes de alterar..."
        $backupDir = Join-Path $workbook.Path "Backup_AutoDiario"
        if (-not (Test-Path -LiteralPath $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir | Out-Null
        }
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupPath = Join-Path $backupDir ("{0}_backup_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($workbook.Name), $stamp, [IO.Path]::GetExtension($workbook.Name))
        [void]$workbook.SaveCopyAs($backupPath)
    }

    $lastInRow = $sheetIn.Cells($sheetIn.Rows.Count, 1).End(-4162).Row
    $groups = [ordered]@{}
    $totalRead = 0
    $skippedBank = 0
    $skippedInvalid = 0

    Show-Status "Lendo comprovantes da aba IN..."
    for ($row = 2; $row -le $lastInRow; $row++) {
        $client = Normalize-Text $sheetIn.Cells($row, 1).Value2
        $value = To-Number $sheetIn.Cells($row, 5).Value2
        $bank = Normalize-Text $sheetIn.Cells($row, 4).Value2

        if ([string]::IsNullOrWhiteSpace($client) -or $null -eq $value) {
            $skippedInvalid++
            continue
        }

        $totalRead++

        if ($bank.ToUpperInvariant() -in @("CHEQUE", "DINHEIRO")) {
            $skippedBank++
            continue
        }

        if (-not $groups.Contains($client)) {
            $groups[$client] = New-Object System.Collections.ArrayList
        }

        [void]$groups[$client].Add([object[]]@(
            $sheetIn.Cells($row, 2).Value2,
            $sheetIn.Cells($row, 3).Value2,
            $bank,
            $value
        ))
    }

    if ($groups.Count -eq 0) {
        throw "Nao encontrei comprovantes validos na aba IN."
    }

    Add-Log "Comprovantes validos lidos: $totalRead"
    Add-Log "Ignorados por BANCO CHEQUE/DINHEIRO: $skippedBank"
    Add-Log "Clientes com comprovantes para Diario: $($groups.Count)"

    $previousCalculation = $excel.Calculation
    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $excel.Calculation = -4135

    $processed = 0
    $totalClients = $groups.Count
    foreach ($client in @($groups.Keys)) {
        try {
            $processed++
            Show-Status "Processando cliente $client ($processed de $totalClients)..." $processed $totalClients
            Add-Log "Cliente ${client}: $($groups[$client].Count) comprovante(s)"
            Update-Block -Sheet $sheetDiario -Client $client -Rows $groups[$client]
        } catch {
            throw "Erro ao atualizar cliente '$client' ($script:AutoDiarioStage): $($_.Exception.Message)"
        }
    }

    Show-Status "Recalculando totais..."
    $excel.Calculation = $previousCalculation
    $sheetDiario.Calculate()

    if (-not $NoSave) {
        Show-Status "Salvando planilha..."
        Invoke-ExcelAction { $workbook.Save() }
    }

    if (-not [string]::IsNullOrWhiteSpace($workbook.Path)) {
        $logPath = Join-Path $workbook.Path "AtualizarDiarioSaldo_ultimo_log.txt"
        $script:AutoDiarioLog | Set-Content -LiteralPath $logPath -Encoding UTF8
    }

    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true

    $message = "Diario atualizado.`nArquivo:`n$($workbook.FullName)`n`nClientes processados: $($groups.Count).`nIgnorados CHEQUE/DINHEIRO: $skippedBank."
    if ($backupPath) { $message += "`nBackup criado em:`n$backupPath" }
    if ($Silent) {
        Write-Host $message
    } else {
        Write-Progress -Activity "Atualizando Diario" -Completed
        Write-Host ""
        Write-Host "Concluido."
        if (-not $NoFinalMessage) {
            [System.Windows.Forms.MessageBox]::Show($message, "Atualizar Diario Saldo") | Out-Null
        }
    }
} catch {
    try {
        $excel.ScreenUpdating = $true
        $excel.DisplayAlerts = $true
        if ($null -ne $previousCalculation) { $excel.Calculation = $previousCalculation }
    } catch {}
    $errorMessage = "$($_.Exception.Message)"
    if ($_.InvocationInfo) {
        $errorMessage += "`nLinha: $($_.InvocationInfo.ScriptLineNumber)`nComando: $($_.InvocationInfo.Line)"
    }
    if ($Silent -or $NoFinalMessage) {
        Write-Error $errorMessage
    } else {
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Erro ao atualizar Diario") | Out-Null
    }
    throw
} finally {
    if ($CloseAfterSave -and $script:WorkbookOpenedByScript -and $null -ne $workbook) {
        try {
            $workbook.Close($false) | Out-Null
        } catch {}
    }
    try {
        if ($createdExcel -and $null -ne $excel -and $excel.Workbooks.Count -eq 0) {
            $excel.Quit()
        }
    } catch {}
}




