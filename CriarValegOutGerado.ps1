param(
    [string]$WorkbookPath = "",
    [string]$TemplatePath = "",
    [string]$TargetPath = "",
    [switch]$Overwrite,
    [switch]$SkipWeChat,
    [switch]$Silent
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ValegOutWin32Window {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@

function Trace-Step {
    param([string]$Message)
    if ($env:VALEG_DEBUG -eq "1") {
        Write-Host ("[VALEG+OUT] " + $Message)
        try { [Console]::Out.Flush() } catch {}
    }
}

$Pepper = [char]::ConvertFromUtf32(0x1FAD1)
$ResumoGrupoWeChat = "2026 " + ($Pepper * 6)

function Resolve-2026FromPath {
    param([string]$PathHint)

    if ([string]::IsNullOrWhiteSpace($PathHint)) { return "" }
    if (-not (Test-Path -LiteralPath $PathHint)) { return "" }

    $item = Get-Item -LiteralPath $PathHint -ErrorAction SilentlyContinue
    if ($null -eq $item) { return "" }

    $dir = if ($item.PSIsContainer) { $item } else { $item.Directory }
    while ($null -ne $dir) {
        if ($dir.Name -eq "2026") { return $dir.FullName }

        $child2026 = Join-Path $dir.FullName "2026"
        if (Test-Path -LiteralPath $child2026) { return $child2026 }

        $dir = $dir.Parent
    }

    return ""
}

function Get-SaldoCliente2026Root {
    param([string]$PathHint = "")

    foreach ($candidate in @(
        (Resolve-2026FromPath -PathHint $PathHint),
        (Resolve-2026FromPath -PathHint $PSScriptRoot),
        (Resolve-2026FromPath -PathHint (Get-Location).Path),
        (Join-Path $env:USERPROFILE "Dropbox\ETH\Saldo cliente\2026")
    )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $userRoots = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName "Dropbox\ETH\Saldo cliente\2026" } |
        Where-Object { Test-Path -LiteralPath $_ }

    $resolvedUserRoot = @($userRoots | Select-Object -First 1)
    if ($resolvedUserRoot.Count -gt 0) { return $resolvedUserRoot[0] }

    throw "Nao encontrei a pasta base '2026' do projeto Saldo cliente neste computador."
}

function Select-WorkbookPath {
    if ($Silent) { throw "Informe o caminho da planilha com -WorkbookPath." }
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Escolha a planilha Saldo cliente do dia"
    $dialog.Filter = "Planilhas Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Todos os arquivos (*.*)|*.*"
    $dialog.InitialDirectory = Get-SaldoCliente2026Root -PathHint $PSScriptRoot
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
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
    $text = (Normalize-Text $Value).Replace("R$", "").Replace(".", "").Replace(",", ".")
    $parsed = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Format-PlainNumber {
    param([double]$Value)
    return ([int64][Math]::Round($Value, 0)).ToString([Globalization.CultureInfo]::InvariantCulture)
}

function Format-MoneyBr {
    param([double]$Value)
    return " R$ " + $Value.ToString("N2", [Globalization.CultureInfo]::GetCultureInfo("pt-BR")) + " "
}

function Get-WorksheetByName {
    param([object]$Workbook, [string]$Name)
    foreach ($ws in @($Workbook.Worksheets)) {
        if ((Normalize-Text $ws.Name).ToUpperInvariant() -eq $Name.ToUpperInvariant()) { return $ws }
    }
    $available = @($Workbook.Worksheets | ForEach-Object { $_.Name }) -join ", "
    throw "A planilha '$($Workbook.Name)' nao tem a aba '$Name'. Abas encontradas: $available"
}

function Get-DateText {
    param([object]$Workbook, [string]$Path)
    try {
        $inSheet = Get-WorksheetByName -Workbook $Workbook -Name "IN"
        $text = Normalize-Text $inSheet.Range("H3").Text
        if ($text -match '^\d{2}/\d{2}/\d{4}$') { return $text }
    } catch {}
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ($name -match '(\d{2})\.(\d{2})\.(\d{2})') { return "$($matches[1])/$($matches[2])/20$($matches[3])" }
    return (Get-Date).ToString("dd/MM/yyyy")
}

function Find-TemplatePath {
    param([string]$BaseDir, [string]$DateFileText)
    $root = Get-SaldoCliente2026Root -PathHint $BaseDir
    $candidate = Get-ChildItem -LiteralPath $root -Recurse -File -Filter "VALEG + OUT *.xlsx" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DirectoryName -ne $BaseDir -and
            $_.Name -notmatch "TESTE|LIMPO|CORRIGIDO|COMPLETO|GERADO|FINAL"
        } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -ne $candidate) { return $candidate.FullName }
    throw "Nao encontrei nenhuma planilha VALEG + OUT para usar como modelo."
}

function Resolve-OpenableTemplatePath {
    param(
        [object]$Excel,
        [string]$PreferredPath,
        [string]$BaseDir
    )

    $candidates = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($PreferredPath) -and (Test-Path -LiteralPath $PreferredPath)) {
        [void]$candidates.Add((Resolve-Path -LiteralPath $PreferredPath).Path)
    }

    $root = Get-SaldoCliente2026Root -PathHint $BaseDir
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "VALEG + OUT *.xlsx" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DirectoryName -ne $BaseDir -and
            $_.Name -notmatch "TESTE|LIMPO|CORRIGIDO|COMPLETO|GERADO|FINAL"
        } |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            if (-not $candidates.Contains($_.FullName)) { [void]$candidates.Add($_.FullName) }
        }

    foreach ($candidate in @($candidates)) {
        $testWorkbook = $null
        try {
            $testWorkbook = $Excel.Workbooks.Open($candidate, 0, $true)
            [void]$testWorkbook.Worksheets.Item("PAGAMENTOS")
            $testWorkbook.Close($false) | Out-Null
            return $candidate
        } catch {
            try { if ($null -ne $testWorkbook) { $testWorkbook.Close($false) | Out-Null } } catch {}
        }
    }

    throw "Nao encontrei nenhum modelo VALEG + OUT que o Excel consiga abrir."
}

function Resolve-ValegFormatTemplatePath {
    param(
        [object]$Excel,
        [string]$BaseDir
    )

    $root = Get-SaldoCliente2026Root -PathHint $BaseDir
    $candidates = Get-ChildItem -LiteralPath $root -Recurse -File -Filter "VALEG + OUT *.xlsx" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DirectoryName -ne $BaseDir -and
            $_.Name -notmatch "TESTE|LIMPO|CORRIGIDO|COMPLETO|GERADO|FINAL"
        } |
        Sort-Object LastWriteTime -Descending

    foreach ($candidate in @($candidates)) {
        $testWorkbook = $null
        try {
            $testWorkbook = $Excel.Workbooks.Open($candidate.FullName, 0, $true)
            $testSheet = $testWorkbook.Worksheets.Item("PAGAMENTOS")
            if (Test-ValegTemplateUsable -Sheet $testSheet) {
                $testWorkbook.Close($false) | Out-Null
                return $candidate.FullName
            }
            $testWorkbook.Close($false) | Out-Null
        } catch {
            try { if ($null -ne $testWorkbook) { $testWorkbook.Close($false) | Out-Null } } catch {}
        }
    }

    return ""
}

function Copy-Format {
    param(
        [object]$SourceSheet,
        [string]$SourceAddress,
        [object]$DestSheet,
        [string]$DestAddress
    )
    $sourceRange = $SourceSheet.Range($SourceAddress)
    $destRange = $DestSheet.Range($DestAddress)
    $sourceRange.Copy() | Out-Null
    $destRange.PasteSpecial(-4122) | Out-Null # xlPasteFormats
    $DestSheet.Application.CutCopyMode = $false
    
    # Sincroniza a altura da linha para que o tamanho fique igual ao modelo
    try {
        $destRange.RowHeight = $sourceRange.RowHeight
    } catch {}
}

function Get-ValegTemplateRows {
    param([object]$Sheet)

    $totalRow = $null
    for ($r = 2; $r -le 30; $r++) {
        if ((Normalize-Text $Sheet.Cells($r, 1).Text).ToUpperInvariant() -eq "VALEG" -and
            (Normalize-Text $Sheet.Cells($r, 3).Text).ToUpperInvariant() -eq "TOTAL") {
            $totalRow = $r
            break
        }
    }

    $dataRow = 2
    if ($null -ne $totalRow) {
        for ($r = 2; $r -lt $totalRow; $r++) {
            if (-not [string]::IsNullOrWhiteSpace((Normalize-Text $Sheet.Cells($r, 4).Text))) {
                $dataRow = $r
                break
            }
        }
    }

    [pscustomobject]@{
        HeaderRow = 1
        DataRow = $dataRow
        TotalRow = $totalRow
        HasDataBeforeTotal = ($null -ne $totalRow -and $totalRow -gt 2)
    }
}

function Get-OutTemplateRows {
    param([object]$Sheet)

    $totalRow = $null
    for ($r = 2; $r -le 40; $r++) {
        $outValue = Normalize-Text $Sheet.Cells($r, 8).Text
        $status = Normalize-Text $Sheet.Cells($r, 9).Text
        $motivo = Normalize-Text $Sheet.Cells($r, 10).Text
        if (-not [string]::IsNullOrWhiteSpace($outValue) -and
            [string]::IsNullOrWhiteSpace($status) -and
            [string]::IsNullOrWhiteSpace($motivo)) {
            $totalRow = $r
            break
        }
    }

    if ($null -eq $totalRow) { $totalRow = 5 }

    $moneyHeaderRow = $null
    for ($r = 2; $r -le 40; $r++) {
        if ((Normalize-Text $Sheet.Cells($r, 6).Text).ToUpperInvariant() -eq "CLIENTE" -and
            (Normalize-Text $Sheet.Cells($r, 7).Text).ToUpperInvariant() -eq "DATA" -and
            (Normalize-Text $Sheet.Cells($r, 8).Text).ToUpperInvariant() -eq "HORA" -and
            (Normalize-Text $Sheet.Cells($r, 9).Text).ToUpperInvariant() -eq "BANCO") {
            $moneyHeaderRow = $r
            break
        }
    }

    if ($null -eq $moneyHeaderRow) { $moneyHeaderRow = $totalRow + 1 }

    [pscustomobject]@{
        HeaderRow = 1
        DataRow = 2
        TotalRow = $totalRow
        MoneyHeaderRow = $moneyHeaderRow
        MoneyDataRow = $moneyHeaderRow + 1
    }
}

function Test-ValegTemplateUsable {
    param([object]$Sheet)
    $rows = Get-ValegTemplateRows -Sheet $Sheet
    return [bool]$rows.HasDataBeforeTotal
}

function Replace-ModelPictures {
    param(
        [object]$SourceWorkbook,
        [object]$TargetSheet,
        [int]$ValegEndRow = 6
    )

    $pictures = @($TargetSheet.Shapes | Sort-Object Left)
    if ($pictures.Count -lt 3) { return }

    $slots = @(
        @{ Shape=$pictures[0]; Sheet="RESUMO SALDO"; Address="A1:B32" },
        @{ Shape=$pictures[1]; Sheet="IN"; Address="H3:I17" },
        @{ Shape=$pictures[2]; Sheet="IN"; Address="K3:L30" }
    )

    foreach ($slot in $slots) {
        $oldShape = $slot.Shape
        $left = $oldShape.Left
        $top = $oldShape.Top
        $width = $oldShape.Width
        $height = $oldShape.Height

        $sourceSheet = Get-WorksheetByName -Workbook $SourceWorkbook -Name $slot.Sheet
        $sourceRange = $sourceSheet.Range($slot.Address)
        $sourceRange.CopyPicture(1, 2) | Out-Null
        Start-Sleep -Milliseconds 250

        $TargetSheet.Paste() | Out-Null
        $newShape = $TargetSheet.Shapes.Item($TargetSheet.Shapes.Count)
        $newShape.Left = $left
        $newShape.Top = $top
        $newShape.Width = $width
        $newShape.Height = $height
        if ($slot.Sheet -eq "RESUMO SALDO") {
            # Posiciona o resumo saldo abaixo do bloco VALEG para nao ficar em cima
            $targetRow = [Math]::Max(6, $ValegEndRow + 2)
            $newShape.Top = [double]$TargetSheet.Cells.Item($targetRow, 1).Top
            # Alinha na Coluna C (3) para ficar mais centralizado e nao na pontinha
            $newShape.Left = [double]$TargetSheet.Cells.Item($targetRow, 3).Left
        }
        $oldShape.Delete() | Out-Null
        $TargetSheet.Application.CutCopyMode = $false
    }
}

function Read-ValegRows {
    param([object]$DiarioSheet)
    $lastRow = $DiarioSheet.Cells($DiarioSheet.Rows.Count, 6).End(-4162).Row
    $totalRow = $null
    for ($r = 1; $r -le $lastRow; $r++) {
        if ((Normalize-Text $DiarioSheet.Cells($r, 6).Text).ToUpperInvariant() -eq "VALEG" -and
            (Normalize-Text $DiarioSheet.Cells($r, 8).Text).ToUpperInvariant() -eq "TOTAL") {
            $totalRow = $r
            break
        }
    }
    if ($null -eq $totalRow) { return @{ Rows = @(); Total = 0.0 } }

    $headerRow = $null
    for ($r = $totalRow - 1; $r -ge 1; $r--) {
        if ((Normalize-Text $DiarioSheet.Cells($r, 6).Text).ToUpperInvariant() -eq "DATA" -and
            (Normalize-Text $DiarioSheet.Cells($r, 8).Text).ToUpperInvariant() -eq "BANCO") {
            $headerRow = $r
            break
        }
    }
    if ($null -eq $headerRow) { return @{ Rows = @(); Total = 0.0 } }

    $rows = New-Object System.Collections.ArrayList
    $total = 0.0
    for ($r = $headerRow + 1; $r -lt $totalRow; $r++) {
        $date = Normalize-Text $DiarioSheet.Cells($r, 6).Text
        $hour = Normalize-Text $DiarioSheet.Cells($r, 7).Text
        $bank = Normalize-Text $DiarioSheet.Cells($r, 8).Text
        $value = To-Number $DiarioSheet.Cells($r, 9).Value2
        if ([string]::IsNullOrWhiteSpace($date) -or [string]::IsNullOrWhiteSpace($bank) -or $null -eq $value) { continue }
        $total += $value
        [void]$rows.Add([pscustomobject]@{ Date=$date; Hour=$hour; Bank=$bank; Value=$value })
    }
    return @{ Rows = @($rows); Total = $total }
}

function Read-OutRows {
    param([object]$OutSheet)
    $lastRow = $OutSheet.Cells($OutSheet.Rows.Count, 1).End(-4162).Row
    $rows = New-Object System.Collections.ArrayList
    $total = 0.0
    for ($r = 2; $r -le $lastRow; $r++) {
        $date = Normalize-Text $OutSheet.Cells($r, 1).Text
        $client = Normalize-Text $OutSheet.Cells($r, 2).Text
        $value = To-Number $OutSheet.Cells($r, 3).Value2
        $status = Normalize-Text $OutSheet.Cells($r, 4).Text
        $reason = Normalize-Text $OutSheet.Cells($r, 5).Text
        if ([string]::IsNullOrWhiteSpace($client) -or $null -eq $value) { continue }
        $total += $value
        [void]$rows.Add([pscustomobject]@{ Date=$date; Client=$client; Value=$value; Status=$status; Reason=$reason })
    }
    return @{ Rows = @($rows); Total = $total }
}

function Read-ChequeDinheiroRows {
    param([object]$InSheet)
    $lastRow = $InSheet.Cells($InSheet.Rows.Count, 1).End(-4162).Row
    $rows = New-Object System.Collections.ArrayList
    for ($r = 2; $r -le $lastRow; $r++) {
        $bank = (Normalize-Text $InSheet.Cells($r, 4).Text).ToUpperInvariant()
        if ($bank -notin @("CHEQUE", "DINHEIRO")) { continue }
        $client = Normalize-Text $InSheet.Cells($r, 1).Text
        $date = Normalize-Text $InSheet.Cells($r, 2).Text
        $hour = Normalize-Text $InSheet.Cells($r, 3).Text
        $value = To-Number $InSheet.Cells($r, 5).Value2
        if ([string]::IsNullOrWhiteSpace($client) -or $null -eq $value) { continue }
        [void]$rows.Add([pscustomobject]@{ Client=$client; Date=$date; Hour=$hour; Bank=$bank; Value=$value })
    }
    return @($rows)
}

function Ensure-ClearArea {
    param([object]$Sheet)
    # Limpa a area editavel. As 3 tabelas grandes sao Shapes do Excel e
    # precisam continuar exatamente no lugar.
    $Sheet.Range("A1:J200").Clear() | Out-Null
}

function Prepare-WhiteCanvas {
    param([object]$Sheet, [int]$LastRow)
    $Sheet.Range("A1:J$LastRow").Interior.Color = 16777215
}

function Write-ValegBlock {
    param([object]$Sheet, [object]$TemplateSheet, [object[]]$Rows, [double]$Total)
    $templateRows = Get-ValegTemplateRows -Sheet $TemplateSheet
    $dataFormatSheet = if ($null -ne $script:ValegFormatTemplateSheet) { $script:ValegFormatTemplateSheet } else { $TemplateSheet }
    $dataFormatRows = Get-ValegTemplateRows -Sheet $dataFormatSheet

    Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("A{0}:D{0}" -f $templateRows.HeaderRow) -DestSheet $Sheet -DestAddress "A1:D1"
    $Sheet.Cells(1, 1).Value2 = "DATA"
    $Sheet.Cells(1, 2).Value2 = "HORARIO"
    $Sheet.Cells(1, 3).Value2 = "BANCO"
    $Sheet.Cells(1, 4).Value2 = "TRANSFER" + [char]0x00CA + "NCIA"

    $r = 2
    foreach ($item in $Rows) {
        Copy-Format -SourceSheet $dataFormatSheet -SourceAddress ("A{0}:D{0}" -f $dataFormatRows.DataRow) -DestSheet $Sheet -DestAddress "A${r}:D${r}"
        $Sheet.Cells($r, 1).NumberFormat = "@"
        $Sheet.Cells($r, 2).NumberFormat = "@"
        $Sheet.Cells($r, 4).NumberFormat = "0"
        $Sheet.Cells($r, 1).Value2 = $item.Date
        $Sheet.Cells($r, 2).Value2 = $item.Hour
        $Sheet.Cells($r, 3).Value2 = $item.Bank
        $Sheet.Cells($r, 4).Value2 = [double]$item.Value
        $r++
    }

    $totalFormatSheet = $TemplateSheet
    $totalTemplateRow = $templateRows.TotalRow
    if ($null -eq $totalTemplateRow) {
        $totalFormatSheet = $dataFormatSheet
        $totalTemplateRow = if ($null -ne $dataFormatRows.TotalRow) { $dataFormatRows.TotalRow } else { 3 }
    }
    Copy-Format -SourceSheet $totalFormatSheet -SourceAddress ("A{0}:D{0}" -f $totalTemplateRow) -DestSheet $Sheet -DestAddress "A${r}:D${r}"
    $Sheet.Cells($r, 1).Value2 = "VALEG"
    $Sheet.Cells($r, 3).Value2 = "TOTAL"
    $Sheet.Cells($r, 4).NumberFormatLocal = """R$"" #.##0,00"
    $Sheet.Cells($r, 4).Value2 = [double]$Total
    return $r
}

function Write-OutBlock {
    param([object]$Sheet, [object]$TemplateSheet, [object[]]$Rows, [double]$Total)
    $templateRows = Get-OutTemplateRows -Sheet $TemplateSheet
    Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("F{0}:J{0}" -f $templateRows.HeaderRow) -DestSheet $Sheet -DestAddress "F1:J1"
    $Sheet.Cells(1, 6).Value2 = "Data"
    $Sheet.Cells(1, 7).Value2 = "CLIENT"
    $Sheet.Cells(1, 8).Value2 = "OUT"
    $Sheet.Cells(1, 9).Value2 = "STATUS"
    $Sheet.Cells(1, 10).Value2 = "MOTIVO"

    $r = 2
    foreach ($item in $Rows) {
        Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("F{0}:J{0}" -f $templateRows.DataRow) -DestSheet $Sheet -DestAddress "F${r}:J${r}"
        $Sheet.Cells($r, 6).NumberFormat = "@"
        $Sheet.Cells($r, 6).Value2 = $item.Date
        $Sheet.Cells($r, 7).Value2 = $item.Client
        $Sheet.Cells($r, 8).Value2 = Format-PlainNumber $item.Value
        $Sheet.Cells($r, 9).Value2 = $item.Status
        $Sheet.Cells($r, 10).Value2 = $item.Reason
        $r++
    }

    Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("F{0}:J{0}" -f $templateRows.TotalRow) -DestSheet $Sheet -DestAddress "F${r}:J${r}"
    $Sheet.Range("F${r}:J${r}").ClearContents() | Out-Null
    $Sheet.Cells($r, 8).Value2 = Format-PlainNumber $Total
    return $r
}

function Write-MoneyBlock {
    param([object]$Sheet, [object]$TemplateSheet, [int]$StartRow, [object[]]$Rows)
    $templateRows = Get-OutTemplateRows -Sheet $TemplateSheet
    
    # Usa o formato do Header principal (igual ao do bloco OUT) para ficar no mesmo tamanho
    Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("F{0}:J{0}" -f $templateRows.HeaderRow) -DestSheet $Sheet -DestAddress "F${StartRow}:J${StartRow}"
    $Sheet.Cells($StartRow, 6).Value2 = "CLIENTE"
    $Sheet.Cells($StartRow, 7).Value2 = "DATA"
    $Sheet.Cells($StartRow, 8).Value2 = "HORA"
    $Sheet.Cells($StartRow, 9).Value2 = "BANCO"
    $Sheet.Cells($StartRow, 10).Value2 = "VALOR"
    # Centraliza o cabecalho
    $Sheet.Range("F${StartRow}:J${StartRow}").HorizontalAlignment = -4108 # xlCenter

    $r = $StartRow + 1
    foreach ($item in $Rows) {
        # Usa o formato do DataRow (igual ao do bloco OUT) para que o tamanho aumente conforme os demais
        Copy-Format -SourceSheet $TemplateSheet -SourceAddress ("F{0}:J{0}" -f $templateRows.DataRow) -DestSheet $Sheet -DestAddress "F${r}:J${r}"
        $Sheet.Cells($r, 7).NumberFormat = "@"
        $Sheet.Cells($r, 8).NumberFormat = "@"
        $Sheet.Cells($r, 6).Value2 = $item.Client
        $Sheet.Cells($r, 7).Value2 = $item.Date
        $Sheet.Cells($r, 8).Value2 = $item.Hour
        $Sheet.Cells($r, 9).Value2 = $item.Bank
        $Sheet.Cells($r, 10).Value2 = Format-PlainNumber $item.Value
        # Centraliza os valores para ficarem organizados
        $Sheet.Range("F${r}:J${r}").HorizontalAlignment = -4108 # xlCenter
        $r++
    }
    return [Math]::Max($StartRow, $r - 1)
}

function Configure-PdfPage {
    param([object]$Sheet)

    $maxBottom = 0.0
    foreach ($shape in @($Sheet.Shapes)) {
        $bottom = [double]$shape.Top + [double]$shape.Height
        if ($bottom -gt $maxBottom) { $maxBottom = $bottom }
    }

    $lastRow = 58
    for ($r = 1; $r -le 200; $r++) {
        $rowBottom = [double]$Sheet.Cells.Item($r, 1).Top + [double]$Sheet.Rows.Item($r).Height
        if ($rowBottom -ge ($maxBottom + 8)) {
            $lastRow = [Math]::Max($r, 58)
            break
        }
    }

    $Sheet.PageSetup.PrintArea = "`$A`$1:`$AA`$$lastRow"
    $Sheet.PageSetup.Orientation = 2 # xlLandscape
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = 1
    $Sheet.PageSetup.LeftMargin = 18
    $Sheet.PageSetup.RightMargin = 18
    $Sheet.PageSetup.TopMargin = 18
    $Sheet.PageSetup.BottomMargin = 18
    $Sheet.PageSetup.CenterHorizontally = $true
    $Sheet.PageSetup.CenterVertically = $true
}

function Activate-WeChat {
    $wshell = New-Object -ComObject WScript.Shell
    Trace-Step "Tentando ativar janela do WeChat..."

    # Tentativa 1: AppActivate (Metodo mais direto do Windows)
    if ($wshell.AppActivate("WeChat")) {
        Start-Sleep -Milliseconds 800
        return
    }

    # Tentativa 2: Busca manual por processo com janela visivel
    $process = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.MainWindowHandle -ne 0 -and 
            $_.MainWindowTitle -ne "" -and (
                $_.ProcessName -like "*WeChat*" -or
                $_.ProcessName -like "*Weixin*" -or
                $_.MainWindowTitle -like "*WeChat*"
            )
        } |
        Sort-Object ProcessName, Id |
        Select-Object -First 1

    if ($null -eq $process) {
        throw "Nao encontrei o WeChat aberto. Por favor, abra o WeChat Desktop e deixe logado."
    }

    Trace-Step "Processo WeChat encontrado: $($process.ProcessName). Forcando foco..."
    [void][ValegOutWin32Window]::ShowWindowAsync($process.MainWindowHandle, 9) # 9 = SW_RESTORE
    Start-Sleep -Milliseconds 400
    [void][ValegOutWin32Window]::SetForegroundWindow($process.MainWindowHandle)
    Start-Sleep -Milliseconds 800
}

function Send-KeysSafe {
    param([string]$Keys, [int]$DelayMs = 250)
    [System.Windows.Forms.SendKeys]::SendWait($Keys)
    Start-Sleep -Milliseconds $DelayMs
}

function Copy-FileToClipboard {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Arquivo nao encontrado para colar no WeChat:`n$Path"
    }
    $files = New-Object System.Collections.Specialized.StringCollection
    [void]$files.Add($Path)
    [System.Windows.Forms.Clipboard]::SetFileDropList($files)
}

function Paste-PdfToWeChatDraft {
    param([string]$PdfPath, [string]$GroupName)

    Write-Host ""
    Write-Host "Colando PDF no WeChat como rascunho:"
    Write-Host $GroupName
    Write-Host "NAO vou apertar Enter para enviar."

    Activate-WeChat
    # Limpa o clipboard antes de definir o novo valor para evitar conflitos
    [System.Windows.Forms.Clipboard]::Clear()
    Start-Sleep -Milliseconds 100
    Set-Clipboard -Value $GroupName
    Send-KeysSafe "^f" 500
    Send-KeysSafe "^a" 200
    Send-KeysSafe "^v" 800
    Send-KeysSafe "{ENTER}" 2000 # Aumentado para 2s para garantir que o grupo carregue

    Activate-WeChat
    Trace-Step "Copiando PDF para o clipboard: $PdfPath"
    Copy-FileToClipboard -Path $PdfPath
    Start-Sleep -Milliseconds 500
    Send-KeysSafe "^v" 1500
    Trace-Step "PDF colado no WeChat"
}

$targetWasDefault = [string]::IsNullOrWhiteSpace($TargetPath)
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Select-WorkbookPath }
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$baseDir = Split-Path -Parent $WorkbookPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$sourceWorkbook = $null
$templateWorkbook = $null
$valegFormatWorkbook = $null
$targetWorkbook = $null

try {
    Trace-Step "abrindo saldo cliente"
    $sourceWorkbook = $excel.Workbooks.Open($WorkbookPath)
    Trace-Step "saldo cliente aberto"
    $diarioSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "Diario"
    $outSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "OUT"
    $inSheet = Get-WorksheetByName -Workbook $sourceWorkbook -Name "IN"

    $dateText = Get-DateText -Workbook $sourceWorkbook -Path $WorkbookPath
    $dateFileText = $dateText.Replace("/", ".")
    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Find-TemplatePath -BaseDir $baseDir -DateFileText $dateFileText
    }
    $TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
    Trace-Step "resolvendo modelo"
    $TemplatePath = Resolve-OpenableTemplatePath -Excel $excel -PreferredPath $TemplatePath -BaseDir $baseDir
    Trace-Step "modelo: $TemplatePath"

    $valegFormatTemplatePath = Resolve-ValegFormatTemplatePath -Excel $excel -BaseDir $baseDir
    Trace-Step "modelo formato VALEG: $valegFormatTemplatePath"

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        $TargetPath = Join-Path $baseDir "VALEG + OUT $dateFileText.xlsx"
    }
    $pdfPath = [System.IO.Path]::ChangeExtension($TargetPath, ".pdf")

    $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $templateFullPath = [System.IO.Path]::GetFullPath($TemplatePath)
    $targetIsTemplate = [string]::Equals($targetFullPath, $templateFullPath, [System.StringComparison]::OrdinalIgnoreCase)

    if ((Test-Path -LiteralPath $TargetPath) -and $Overwrite -and -not $targetIsTemplate) { Remove-Item -LiteralPath $TargetPath -Force }
    if ((Test-Path -LiteralPath $pdfPath) -and $Overwrite) { Remove-Item -LiteralPath $pdfPath -Force }

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        Trace-Step "copiando modelo para destino"
        Copy-Item -LiteralPath $TemplatePath -Destination $TargetPath
        Trace-Step "destino copiado"
    }

    if ($targetIsTemplate) {
        Trace-Step "abrindo destino como modelo"
        $targetWorkbook = $excel.Workbooks.Open($TargetPath, 0, $false)
        $templateWorkbook = $targetWorkbook
        $sheet = $targetWorkbook.Worksheets.Item("PAGAMENTOS")
        $templateSheet = $sheet
    } else {
        Trace-Step "abrindo modelo somente leitura"
        $templateWorkbook = $excel.Workbooks.Open($TemplatePath, 0, $true)
        $templateSheet = $templateWorkbook.Worksheets.Item("PAGAMENTOS")
        $script:ValegFormatTemplateSheet = $templateSheet
        if (-not [string]::IsNullOrWhiteSpace($valegFormatTemplatePath)) {
            $valegFormatWorkbook = $excel.Workbooks.Open($valegFormatTemplatePath, 0, $true)
            $script:ValegFormatTemplateSheet = $valegFormatWorkbook.Worksheets.Item("PAGAMENTOS")
        }
        Trace-Step "abrindo destino"
        $targetWorkbook = $excel.Workbooks.Open($TargetPath, 0, $false)
        $sheet = $targetWorkbook.Worksheets.Item("PAGAMENTOS")
        Trace-Step "destino aberto"
    }

    Trace-Step "lendo dados"
    $valeg = Read-ValegRows -DiarioSheet $diarioSheet
    $out = Read-OutRows -OutSheet $outSheet
    $moneyRows = Read-ChequeDinheiroRows -InSheet $inSheet
    Trace-Step ("dados lidos: VALEG={0} OUT={1} cheque/dinheiro={2}" -f @($valeg.Rows).Count, @($out.Rows).Count, @($moneyRows).Count)

    Ensure-ClearArea -Sheet $sheet

    Trace-Step "preenchendo blocos"
    $valegEnd = Write-ValegBlock -Sheet $sheet -TemplateSheet $templateSheet -Rows @($valeg.Rows) -Total ([double]$valeg.Total)
    $outEnd = Write-OutBlock -Sheet $sheet -TemplateSheet $templateSheet -Rows @($out.Rows) -Total ([double]$out.Total)
    $moneyStart = $outEnd + 1
    $moneyEnd = Write-MoneyBlock -Sheet $sheet -TemplateSheet $templateSheet -StartRow $moneyStart -Rows @($moneyRows)
    Trace-Step "blocos preenchidos"

    Trace-Step "atualizando imagens grandes"
    Replace-ModelPictures -SourceWorkbook $sourceWorkbook -TargetSheet $sheet -ValegEndRow $valegEnd
    Trace-Step "imagens grandes atualizadas"

    for ($c = 1; $c -le 10; $c++) {
        $sheet.Columns.Item($c).ColumnWidth = $templateSheet.Columns.Item($c).ColumnWidth
    }

    Configure-PdfPage -Sheet $sheet

    Trace-Step "finalizando processo excel"
    $targetWorkbook.Save() | Out-Null
    $sheet.ExportAsFixedFormat(0, $pdfPath) | Out-Null
    Trace-Step "pdf exportado em $pdfPath"

    if (-not $SkipWeChat) {
        Trace-Step "iniciando colagem no WeChat em 2 segundos..."
        Start-Sleep -Seconds 2
        Paste-PdfToWeChatDraft -PdfPath $pdfPath -GroupName $ResumoGrupoWeChat
    }

    if ($Silent) {
        Write-Host "VALEG + OUT gerado: $TargetPath"
        Write-Host "PDF gerado: $pdfPath"
        Write-Host "Modelo usado: $TemplatePath"
        if (-not $SkipWeChat) { Write-Host "PDF colado no WeChat como rascunho: $ResumoGrupoWeChat" }
    } else {
        $msg = "VALEG + OUT gerado:`n$TargetPath`n`nPDF:`n$pdfPath`n`nModelo usado:`n$TemplatePath"
        if (-not $SkipWeChat) { $msg += "`n`nPDF colado no WeChat como rascunho:`n$ResumoGrupoWeChat" }
        [System.Windows.Forms.MessageBox]::Show($msg, "Criar VALEG + OUT") | Out-Null
    }
} catch {
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro ao criar VALEG + OUT") | Out-Null
    }
    throw
} finally {
    try { if ($null -ne $targetWorkbook) { $targetWorkbook.Close($true) | Out-Null } } catch {}
    try { if ($null -ne $valegFormatWorkbook) { $valegFormatWorkbook.Close($false) | Out-Null } } catch {}
    try { if ($null -ne $templateWorkbook -and -not [object]::ReferenceEquals($templateWorkbook, $targetWorkbook)) { $templateWorkbook.Close($false) | Out-Null } } catch {}
    try { if ($null -ne $sourceWorkbook) { $sourceWorkbook.Close($false) | Out-Null } } catch {}
    try { $excel.DisplayAlerts = $true } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
