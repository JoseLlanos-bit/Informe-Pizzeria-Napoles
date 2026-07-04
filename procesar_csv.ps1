param(
    [Parameter(Mandatory=$true)]
    [string]$Turno,
    [string]$CsvFolder = "C:\Users\jmlla\Desktop\Pizzeria Napoles\Ventas para dashboard",
    [string]$IndexHtml = "C:\Users\jmlla\Desktop\Dashboard PzNp\index.html",
    [string]$MapFile = "C:\Users\jmlla\AppData\Local\Temp\opencode\excel_map.json",
    [int]$CostOverride = 0,
    [int]$PropinaOverride = 0,
    [int]$DescuentoOverride = 0
)

$ErrorActionPreference = "Stop"

Write-Host "=== PROCESAR CSV ===" -ForegroundColor Cyan
Write-Host "Turno: $Turno"
Write-Host ""

# ── Load category map ──
$map = @{}
if (Test-Path $MapFile) {
    $raw = Get-Content $MapFile -Raw -Encoding UTF8
    $map = $raw | ConvertFrom-Json
}
Write-Host ("Map loaded: $($map.Count) items") -ForegroundColor Gray

# ── Build CSV path ──
$csvPath = Join-Path $CsvFolder "${Turno} ventas_totales.csv"
if (-not (Test-Path $csvPath)) {
    Write-Host "ERROR: CSV no encontrado: $csvPath" -ForegroundColor Red
    exit 1
}
Write-Host ("CSV: $csvPath") -ForegroundColor Gray

# ── Parse CSV ──
$rows = Import-Csv -LiteralPath $csvPath -Delimiter "`t"
Write-Host ("Filas leidas: $($rows.Count)") -ForegroundColor Gray

# ── Generate DATA entries ──
$dataLines = @()
$vSum = 0; $cSum = 0; $qSum = 0; $dSum = 0
$orderIds = @{}
$propinaByOrder = @{}
$hourly = @{}
$orderHours = @{}

foreach ($r in $rows) {
    # Clean ID: CSV has "'ID" format (double-quoted with leading ')
    $oid = $r.'ID de orden'.Trim('"', "'")
    
    $nom = $r.Nombre
    $q = [int]$r.Cantidad
    $pr = [int]$r.'Precio a Pagar'
    $co = [int]$r.Costo
    $desc = [int]$r.Descuento
    
    if ($pr -eq 0 -and $desc -eq 0) { continue }
    
    # Category lookup
    $cat = "Otros"; $subcat = "Otros"
    $match = $map.$nom
    if ($match) { $cat = $match.cat; $subcat = $match.subcat }
    
    $u = $pr - $co
    if ($pr -gt 0) { $m = ([Math]::Round($u / $pr * 100, 1)).ToString([System.Globalization.CultureInfo]::InvariantCulture) } else { $m = "0" }
    
    # Proper JSON escaping for product name
    $nomEsc = $nom -replace '"', '\"' -replace "`n", "" -replace "`r", ""
    
    $json = ('{{"p":"{0}","q":{1},"pr":{2},"co":{3},"u":{4},"m":{5},' +
             '"cat":"{6}","subcat":"{7}","g":"","f":"2026-{8}","fp":"",' +
             '"id":"{9}","tn":"{10}","d":{11}}}') -f
        $nomEsc, $q, $pr, $co, $u, $m,
        $cat, $subcat,
        ($Turno -replace '-', '-'),  # keep the same format
        $oid, $Turno, $desc
    
    $dataLines += $json
    $vSum += $pr; $cSum += $co; $qSum += $q; $dSum += $desc
    $orderIds[$oid] = $true
    
    if (-not $propinaByOrder.ContainsKey($oid)) {
        $propinaByOrder[$oid] = [int]$r.Propina
    }
    
    # Hourly aggregation
    $hora = $r.'Hora de creacion'
    $hr = [int]($hora -split ":")[0]
    if (-not $hourly.ContainsKey($hr)) { $hourly[$hr] = 0 }
    $hourly[$hr] += $pr
    
    # Order-hour mapping
    if (-not $orderHours.ContainsKey($oid)) { $orderHours[$oid] = $hr }
}

$totalOrd = $orderIds.Count
$totalProp = ($propinaByOrder.Values | Measure-Object -Sum).Sum
if ($PropinaOverride -gt 0) { $totalProp = $PropinaOverride }
$uSum = $vSum - $cSum
if ($vSum -gt 0) { $mg = ([Math]::Round($uSum / $vSum * 100, 1)).ToString([System.Globalization.CultureInfo]::InvariantCulture) } else { $mg = "0" }

$finalCost = if ($CostOverride -gt 0) { $CostOverride } else { $cSum }
$finalDesc = if ($DescuentoOverride -ne 0) { $DescuentoOverride } else { $dSum }

Write-Host ""
Write-Host ("Items: {0,6}" -f $dataLines.Count) -ForegroundColor Yellow
Write-Host ("Ordenes: {0,5}" -f $totalOrd) -ForegroundColor Yellow
Write-Host ("Venta: {0,8}" -f $vSum) -ForegroundColor Yellow
Write-Host ("Costo: {0,8} (CSV: {1,8})" -f $finalCost, $cSum) -ForegroundColor Yellow
Write-Host ("Propina: {0,7}" -f $totalProp) -ForegroundColor Yellow
Write-Host ("Margen: {0,5}%" -f $mg) -ForegroundColor Yellow

# ── Build insertion blocks ──

# DATA block
$dataBlock = "`r`n" + ($dataLines -join ",`r`n")

# TURNOS entry
$turnoJson = ('{{"tn":"{0}","d":{1},"u":{2},"p":{3},"mg":{4},"t":{5},"v":{6},"c":{7},"mk":0,"n":{8}}}') -f
    $Turno, $finalDesc, $uSum, $totalProp, $mg, $totalOrd, $vSum, $finalCost, $qSum

# HOURLY_DATA entry
$hourKeys = $hourly.Keys | Sort-Object
$hourInner = ($hourKeys | ForEach-Object { "`"$_`":$($hourly[$_])" }) -join ","
$hourlyLine = "  `"$Turno`": { $hourInner }"

# ORDER_HOURS entries (FIXED: use "ID":hr not "'ID'":hr)
$ohLines = @()
foreach ($oid in ($orderHours.Keys | Sort-Object)) {
    $ohLines += "  `"$oid`":$($orderHours[$oid])"
}
$ohBlock = "`r`n" + ($ohLines -join ",`r`n")

# ── Read index.html ──
$c = [System.IO.File]::ReadAllText($IndexHtml, [System.Text.Encoding]::UTF8)

# ===== 1. HOURLY_DATA — insert before `};` =====
$hdIdx = $c.IndexOf("var HOURLY_DATA = {")
$hdEndKnown = $c.IndexOf("var SORTED_TURNOS", $hdIdx)

# Find the closing `};` of HOURLY_DATA by scanning between hdIdx and SORTED_TURNOS
$searchStart = $hdIdx
$braceDepth = 0
$inStr = $null
$hdCloseBrace = -1
for ($i = $searchStart; $i -lt $hdEndKnown; $i++) {
    $ch = $c[$i]; $nn = if ($i + 1 -lt $c.Length) { $c[$i+1] } else { $null }
    if ($inStr) {
        if ($ch -eq '\' -and $nn -eq $inStr) { $i++ }
        elseif ($ch -eq $inStr) { $inStr = $null }
        continue
    }
    if ($ch -eq "'" -or $ch -eq '"') { $inStr = $ch }
    if ($ch -eq '{') { $braceDepth++ }
    elseif ($ch -eq '}') {
        $braceDepth--
        if ($braceDepth -eq 0) {
            # Found closing brace of HOURLY_DATA, check if followed by ;
            if ($nn -eq ';') { $hdCloseBrace = $i; break }
        }
    }
}
if ($hdCloseBrace -lt 0) { Write-Host "ERROR: No se pudo encontrar el cierre de HOURLY_DATA" -ForegroundColor Red; exit 1 }
$c = $c.Substring(0, $hdCloseBrace) + ",`r`n" + $hourlyLine + $c.Substring($hdCloseBrace)
Write-Host "HOURLY_DATA insertado" -ForegroundColor Gray

# ===== 2. ORDER_HOURS — insert before `};` =====
$ohIdx = $c.IndexOf("var ORDER_HOURS = {")
$ohEndSearch = $c.IndexOf("var HOURLY_DATA", $ohIdx)
$braceDepth = 0; $inStr = $null; $ohCloseBrace = -1
for ($i = $ohIdx; $i -lt $ohEndSearch; $i++) {
    $ch = $c[$i]; $nn = if ($i + 1 -lt $c.Length) { $c[$i+1] } else { $null }
    if ($inStr) {
        if ($ch -eq '\' -and $nn -eq $inStr) { $i++ }
        elseif ($ch -eq $inStr) { $inStr = $null }
        continue
    }
    if ($ch -eq "'" -or $ch -eq '"') { $inStr = $ch }
    if ($ch -eq '{') { $braceDepth++ }
    elseif ($ch -eq '}') {
        $braceDepth--
        if ($braceDepth -eq 0 -and $nn -eq ';') { $ohCloseBrace = $i; break }
    }
}
if ($ohCloseBrace -lt 0) { Write-Host "ERROR: No se pudo encontrar el cierre de ORDER_HOURS" -ForegroundColor Red; exit 1 }
$c = $c.Substring(0, $ohCloseBrace) + "," + $ohBlock + "`r`n" + $c.Substring($ohCloseBrace)
Write-Host "ORDER_HOURS insertado" -ForegroundColor Gray

# ===== 3. DATA — insert at end of array =====
$dIdx = $c.IndexOf("var DATA = [")
$dClose = $c.IndexOf("];", $dIdx)
$c = $c.Substring(0, $dClose) + "," + $dataBlock + "`r`n" + $c.Substring($dClose)
Write-Host "DATA insertado ($($dataLines.Count) entries)" -ForegroundColor Gray

# ===== 4. TURNOS — insert at end of array =====
$tnIdx = $c.IndexOf("var TURNOS = [")
$tnClose = $c.IndexOf("];", $tnIdx)
$c = $c.Substring(0, $tnClose) + ",`r`n" + $turnoJson + "`r`n" + $c.Substring($tnClose)
Write-Host "TURNOS insertado" -ForegroundColor Gray

# ===== 5. SORTED_TURNOS — add if not present =====
$stIdx = $c.IndexOf("var SORTED_TURNOS = [")
$stClose = $c.IndexOf("];", $stIdx)
$stSection = $c.Substring($stIdx, $stClose - $stIdx)
if (-not $stSection.Contains("`"$Turno`"")) {
    # Insert before closing ]
    $c = $c.Substring(0, $stClose) + ",`"$Turno`"" + $c.Substring($stClose)
    Write-Host "SORTED_TURNOS actualizado" -ForegroundColor Gray
} else {
    Write-Host "SORTED_TURNOS ya contiene $Turno" -ForegroundColor Gray
}

# ── Post-insertion validation ──
Write-Host ""
Write-Host "=== VALIDACION POST-INSERCION ===" -ForegroundColor Cyan

$scStart = $c.IndexOf('<script>', 148500)
$scEnd = $c.IndexOf('</script>', $scStart)
$scContent = $c.Substring($scStart + 8, $scEnd - $scStart - 8)

$depth = 0; $inStr = $null; $negCount = 0
for ($i = 0; $i -lt $scContent.Length; $i++) {
    $ch = $scContent[$i]; $nn = if ($i + 1 -lt $scContent.Length) { $scContent[$i+1] } else { $null }
    if ($inStr) {
        if ($ch -eq '\' -and $nn -eq $inStr) { $i++ }
        elseif ($ch -eq $inStr) { $inStr = $null }
        continue
    }
    if ($ch -eq "'" -or $ch -eq '"') { $inStr = $ch }
    if ($ch -eq '{') { $depth++ }
    elseif ($ch -eq '}') { $depth--; if ($depth -lt 0) { $negCount++; $depth = 0 } }
}

if ($depth -eq 0 -and $negCount -eq 0) {
    Write-Host "VALIDACION OK (depth=$depth)" -ForegroundColor Green
} else {
    Write-Host "ERROR: depth=$depth, negatives=$negCount" -ForegroundColor Red
    Write-Host "Revise el archivo antes de usar" -ForegroundColor Red
    exit 1
}

# ── Write file ──
[System.IO.File]::WriteAllText($IndexHtml, $c, [System.Text.Encoding]::UTF8)
Write-Host ""
Write-Host "ARCHIVO GUARDADO" -ForegroundColor Green
Write-Host ("Tamano: {0:F2} MB" -f ((Get-Item $IndexHtml).Length / 1MB)) -ForegroundColor Gray
Write-Host ""
Write-Host "Para validar despues de cualquier cambio manual, ejecute:" -ForegroundColor Cyan
Write-Host "  powershell -File `"$(Join-Path (Split-Path $IndexHtml -Parent) validar.ps1)`"" -ForegroundColor Cyan
