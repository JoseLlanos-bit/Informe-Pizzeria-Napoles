param(
    [Parameter(Mandatory=$true)]
    [string]$Mes,                          # ej: "07"
    [string]$MesNombre,                    # ej: "Julio"
    [string]$CsvFolder = "C:\Users\jmlla\Desktop\Pizzeria Napoles\Ventas para dashboard"
)

# --- Mapa de abreviaciones de mes en espanol (minusculas) ---
$mesAbr = @{
    "01"="ene"; "02"="feb"; "03"="mar"; "04"="abr"; "05"="may"; "06"="jun";
    "07"="jul"; "08"="ago"; "09"="sep"; "10"="oct"; "11"="nov"; "12"="dic"
}

$files = Get-ChildItem -Path $CsvFolder -Filter "*-$Mes ventas_totales.csv" | Sort-Object Name
if ($files.Count -eq 0) {
    Write-Host "ERROR: No se encontraron CSV diarios para el mes $Mes"
    exit 1
}
Write-Host "Consolidando $($files.Count) archivos diarios del mes $Mes"

$sb = New-Object System.Text.StringBuilder
$headerDone = $false
$totalRecords = 0

foreach ($f in $files) {
    $day = $f.Name.Substring(0, 2)                     # "01"
    $turnoLabel = "$day-$($mesAbr[$Mes])"              # "01-jul"
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    $lines = $text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($i -eq 0) {
            # Header: solo primera vez
            if (-not $headerDone) {
                $fields = $line -split "`t"
                $clean = @()
                foreach ($fld in $fields) {
                    if ($fld.Length -ge 2 -and $fld[0] -eq '"' -and $fld[$fld.Length-1] -eq '"') {
                        $clean += $fld.Substring(1, $fld.Length - 2)
                    } else {
                        $clean += $fld
                    }
                }
                [void]$sb.AppendLine(($clean -join ';') + ';Turno')
                $headerDone = $true
            }
            continue
        }
        $fields = $line -split "`t"
        $clean = @()
        for ($c = 0; $c -lt $fields.Count; $c++) {
            $fld = $fields[$c]
            if ($fld.Length -ge 2 -and $fld[0] -eq '"' -and $fld[$fld.Length-1] -eq '"') {
                $fld = $fld.Substring(1, $fld.Length - 2)
            }
            if ($c -eq 1 -or $c -eq 3 -or $c -eq 42) {
                # Transformar fecha YYYY-MM-DD -> DD-MM-YYYY
                $m = [regex]::Match($fld, '^(\d{4})-(\d{2})-(\d{2})$')
                if ($m.Success) {
                    $fld = $m.Groups[3].Value + '-' + $m.Groups[2].Value + '-' + $m.Groups[1].Value
                }
            }
            $clean += $fld
        }
        [void]$sb.AppendLine(($clean -join ';') + ';' + $turnoLabel + ';')
        $totalRecords++
    }
}

$outPath = Join-Path $CsvFolder "Consolidado $MesNombre $((Get-Date).Year).csv"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($outPath, $sb.ToString(), $utf8Bom)

Write-Host "Registros: $totalRecords"
Write-Host "Guardado: $outPath"
