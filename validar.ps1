param(
    [switch]$Fix,
    [string]$Path = "C:\Users\jmlla\Desktop\Dashboard PzNp\index.html"
)

$text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)

function Scan-Braces {
    param($Content, $Name, $Start, $End)
    $depth = 0; $inStr = $null; $negCount = 0
    for ($i = $Start; $i -lt $End; $i++) {
        $c = $Content[$i]; $nc = if ($i + 1 -lt $End) { $Content[$i+1] } else { $null }
        if ($inStr) {
            if ($c -eq '\' -and $nc -eq $inStr) { $i++ }
            elseif ($c -eq $inStr) { $inStr = $null }
            continue
        }
        if ($c -eq "'" -or $c -eq '"') { $inStr = $c }
        if ($c -eq '{') { $depth++ }
        elseif ($c -eq '}') { $depth--; if ($depth -lt 0) { $negCount++; $depth = 0 } }
    }
    return @{name=$Name; depth=$depth; negative=$negCount}
}

# Find the main script section
$tagStart = $text.IndexOf('<script>', 148500)
$tagEnd = $text.IndexOf('</script>', $tagStart)
$sc = $text.Substring($tagStart + 8, $tagEnd - $tagStart - 8)

# Map each section
$sections = @(
    @{name="ORDER_HOURS"; pat="var ORDER_HOURS"; next="var HOURLY_DATA"}
    @{name="HOURLY_DATA"; pat="var HOURLY_DATA"; next="var SORTED_TURNOS"}
    @{name="SORTED_TURNOS"; pat="var SORTED_TURNOS"; next="var DATA = ["}
    @{name="DATA"; pat="var DATA = ["; next="var TURNOS = ["}
    @{name="TURNOS"; pat="var TURNOS = ["; next="function UPDATE"}
)

$allOk = $true
$prevEnd = 0
$scriptEnd = $sc.Length

Write-Host "=== VALIDACION DE ESTRUCTURA ===" -ForegroundColor Cyan
Write-Host "Archivo: $Path" -ForegroundColor Cyan
Write-Host ""

foreach ($sec in $sections) {
    $startIdx = $sc.IndexOf($sec.pat)
    if ($startIdx -lt 0) { Write-Host "NO ENCONTRADO: $($sec.name)" -ForegroundColor Red; $allOk = $false; continue }
    
    $nextIdx = if ($sec.next -and $sec.next.StartsWith("function")) { $scriptEnd } else { $sc.IndexOf($sec.next, $startIdx + 1) }
    if ($nextIdx -lt 0) { $nextIdx = $scriptEnd }
    
    # Find actual end of the declaration (after ; or ] or })
    $endIdx = $nextIdx
    
    $result = Scan-Braces $sc $sec.name $startIdx ($endIdx)
    $status = if ($result.depth -eq 0 -and $result.negative -eq 0) { "OK" } else { "ERROR" }
    if ($status -eq "OK") {
        Write-Host ("  {0,-20} depth={1,3} negatives={2,3}  OK" -f $sec.name, $result.depth, $result.negative) -ForegroundColor Green
    } else {
        Write-Host ("  {0,-20} depth={1,3} negatives={2,3}  ERROR" -f $sec.name, $result.depth, $result.negative) -ForegroundColor Red
        $allOk = $false
    }
}

# Full script scan
$fullResult = Scan-Braces $sc "SCRIPT COMPLETO" 0 $scriptEnd
Write-Host ""
if ($fullResult.depth -eq 0 -and $fullResult.negative -eq 0) {
    Write-Host ("SCRIPT COMPLETO: depth={0,3} negatives={1,3}  OK" -f $fullResult.depth, $fullResult.negative) -ForegroundColor Green
} else {
    Write-Host ("SCRIPT COMPLETO: depth={0,3} negatives={1,3}  ERROR" -f $fullResult.depth, $fullResult.negative) -ForegroundColor Red
    $allOk = $false
}

# Count DATA entries
$dataStart = $sc.IndexOf('var DATA = [')
$dataEnd = $sc.IndexOf('];', $dataStart)
$dataSection = $sc.Substring($dataStart, $dataEnd - $dataStart + 2)
$dataCount = [regex]::Matches($dataSection, '\{"p":"').Count
$braceCount = ($dataSection.ToCharArray() | Where-Object { $_ -eq '{' }).Count

# Count TURNOS entries
$tStart = $sc.IndexOf('var TURNOS = [')
$tEnd = $sc.IndexOf('];', $tStart)
$tSec = $sc.Substring($tStart, $tEnd - $tStart + 2)
$tCount = [regex]::Matches($tSec, '\{').Count

Write-Host ""
Write-Host ("DATA entries: {0} (via {{p), {1} (via {{{{ )" -f $dataCount, $braceCount) -ForegroundColor Cyan
Write-Host ("TURNOS entries: ${tCount}") -ForegroundColor Cyan

# Check for dangling fragments
$dangling = [regex]::Matches($sc, '"d":-?\d+\}\r?\n"d":-?\d+\}')
Write-Host ("Dangling fragments: $($dangling.Count)") -ForegroundColor $(if ($dangling.Count -eq 0) { "Green" } else { "Red" })

if ($allOk) {
    Write-Host ""
    Write-Host "TODO OK" -ForegroundColor Green
    exit 0
} else {
    Write-Host ""
    Write-Host "HAY ERRORES - revisar arriba" -ForegroundColor Red
    exit 1
}
