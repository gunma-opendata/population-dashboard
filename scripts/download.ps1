# 群馬県移動人口調査 データダウンロード・集計スクリプト
# 出典: https://toukei.pref.gunma.jp/idj/
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Continue'

$baseUrl  = 'https://toukei.pref.gunma.jp/idj/data'
$cacheDir = 'C:\Users\tkudo\gunma_pop\data\csv'
$outFile  = 'C:\Users\tkudo\gunma_pop\data\population.js'

# 対象年月 (2016/01 - 2026/05)
$months = @()
for ($y = 2016; $y -le 2026; $y++) {
    $maxM = if ($y -eq 2026) { 5 } else { 12 }
    for ($m = 1; $m -le $maxM; $m++) { $months += '{0}{1:D2}' -f $y, $m }
}

# 対象市区町村 (GeoJSONのnameプロパティと一致)
$validCities = @(
    '前橋市','高崎市','桐生市','伊勢崎市','太田市','沼田市','館林市','渋川市','藤岡市','富岡市','安中市','みどり市',
    '榛東村','吉岡町','上野村','神流町','下仁田町','南牧村','甘楽町','中之条町','長野原町','嬬恋村','草津町','高山村','東吾妻町',
    '片品村','川場村','昭和村','みなかみ町','玉村町','板倉町','明和町','千代田町','大泉町','邑楽町'
)

# =====================================================================
#  CSV パーサ (2024/06 以降)
# =====================================================================
function Parse-Num($s) {
    $c = $s.Trim() -replace '[\s,"]', '' -replace '^[^\d-]+', ''
    if ($c -eq '' -or $c -eq '-') { return 0 }
    try { return [int]$c } catch { return 0 }
}
function Parse-CsvLine($line) {
    $fields = [System.Collections.Generic.List[string]]::new()
    $current = [System.Text.StringBuilder]::new()
    $inQuote = $false
    foreach ($ch in $line.ToCharArray()) {
        if ($ch -eq '"') { $inQuote = -not $inQuote }
        elseif ($ch -eq ',' -and -not $inQuote) { $fields.Add($current.ToString()); $current.Clear() | Out-Null }
        else { $current.Append($ch) | Out-Null }
    }
    $fields.Add($current.ToString())
    return ,$fields.ToArray()
}

# =====================================================================
#  Excel COM (XLS/XLSX: 2016/01 - 2024/05)
# =====================================================================
$xlApp = $null
function Get-XlApp {
    if ($null -eq $script:xlApp) {
        $script:xlApp = New-Object -ComObject Excel.Application
        $script:xlApp.Visible = $false
        $script:xlApp.DisplayAlerts = $false
    }
    return $script:xlApp
}

function Read-ExcelData($path) {
    $xl = Get-XlApp
    $wb = $xl.Workbooks.Open($path, $false, $true)
    try {
        $ws = $wb.Sheets.Item(1)
        $maxRow = $ws.UsedRange.Rows.Count
        # 名前列を動的検出 (col 1-5 のいずれかに "県計" がある)
        $nameCol = 2
        for ($c = 1; $c -le 5; $c++) {
            $v = $ws.Cells.Item(7, $c).Text.Trim() -replace '\s+', ''
            if ($v -match '^県+計') { $nameCol = $c; break }
        }
        $result = @{ prefecture = $null; cities = [ordered]@{} }
        for ($r = 7; $r -le $maxRow; $r++) {
            $nameRaw = $ws.Cells.Item($r, $nameCol).Text.Trim() -replace '\s+', ''
            if ($nameRaw -eq '') { continue }
            $gv = { param($off) $v = $ws.Cells.Item($r, $nameCol+$off).Value2; if ($null -eq $v) { 0 } else { [int][Math]::Round($v) } }
            if ($nameRaw -match '^県+計$') {
                $result.prefecture = @{
                    pop=(& $gv 1);male=(& $gv 2);female=(& $gv 3);change=(& $gv 4)
                    births=(& $gv 6);deaths=(& $gv 7);inflow=(& $gv 9);outflow=(& $gv 10)
                    households=(& $gv 11);hhChange=(& $gv 12)
                }
            } elseif ($validCities -contains $nameRaw) {
                $result.cities[$nameRaw] = @{
                    pop=(& $gv 1);male=(& $gv 2);female=(& $gv 3);change=(& $gv 4)
                    births=(& $gv 6);deaths=(& $gv 7);inflow=(& $gv 9);outflow=(& $gv 10)
                    households=(& $gv 11);hhChange=(& $gv 12)
                }
            }
        }
        return $result
    } finally {
        $wb.Close($false)
    }
}

# =====================================================================
#  メインループ
# =====================================================================
$prefData   = [ordered]@{}
$citiesData = [ordered]@{}
foreach ($city in $validCities) { $citiesData[$city] = [ordered]@{} }
$available  = @()

foreach ($ym in $months) {
    # --- CSV (2024/06+) ---
    $csvPath = "$cacheDir\idj$ym.csv"
    if (-not (Test-Path $csvPath)) {
        try { Invoke-WebRequest -Uri "$baseUrl/idj$ym.csv" -OutFile $csvPath -UseBasicParsing -ErrorAction Stop } catch {}
    }
    if (Test-Path $csvPath) {
        $raw    = [IO.File]::ReadAllBytes($csvPath)
        $hasBom = $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF
        $enc    = if ($hasBom) { [Text.Encoding]::UTF8 } else { [Text.Encoding]::GetEncoding('shift_jis') }
        $lines  = $enc.GetString($raw).Replace("`r`n","`n").Replace("`r","`n").Split("`n")
        $prefFound = $false
        foreach ($line in $lines) {
            if (-not $line.Trim()) { continue }
            $cols = Parse-CsvLine $line
            if ($cols.Count -lt 12) { continue }
            $name = $cols[0].Trim() -replace '\s+', ''
            if (-not $prefFound -and ($name -match '^県+計$')) {
                $pop = Parse-Num $cols[1]
                if ($pop -gt 0) {
                    $prefData[$ym] = [ordered]@{
                        pop=$pop;male=(Parse-Num $cols[2]);female=(Parse-Num $cols[3])
                        change=(Parse-Num $cols[4]);births=(Parse-Num $cols[6]);deaths=(Parse-Num $cols[7])
                        inflow=(Parse-Num $cols[9]);outflow=(Parse-Num $cols[10])
                        households=(Parse-Num $cols[11]);hhChange=(Parse-Num $cols[12])
                    }
                    $prefFound = $true
                }
            }
            if ($validCities -contains $name) {
                $pop = Parse-Num $cols[1]
                if ($pop -gt 0) {
                    $citiesData[$name][$ym] = [ordered]@{
                        pop=$pop;male=(Parse-Num $cols[2]);female=(Parse-Num $cols[3])
                        change=(Parse-Num $cols[4]);births=(Parse-Num $cols[6]);deaths=(Parse-Num $cols[7])
                        inflow=(Parse-Num $cols[9]);outflow=(Parse-Num $cols[10])
                        households=(Parse-Num $cols[11]);hhChange=(Parse-Num $cols[12])
                    }
                }
            }
        }
        if ($prefFound) { $available += $ym; Write-Host -NoNewline '.' } else { Write-Host -NoNewline '?' }
        continue
    }

    # --- XLS/XLSX (〜2024/05) ---
    $xlFile = $null
    foreach ($ext in @('xlsx','xls')) {
        $p = "$cacheDir\idj$ym.$ext"
        if (-not (Test-Path $p)) {
            try { Invoke-WebRequest -Uri "$baseUrl/idj$ym.$ext" -OutFile $p -UseBasicParsing -ErrorAction Stop } catch {}
        }
        if (Test-Path $p) { $xlFile = $p; break }
    }
    if ($null -eq $xlFile) { Write-Host -NoNewline 'x'; continue }

    try {
        $data = Read-ExcelData $xlFile
        if ($null -ne $data.prefecture) {
            $prefData[$ym] = [ordered]@{
                pop=$data.prefecture.pop;male=$data.prefecture.male;female=$data.prefecture.female
                change=$data.prefecture.change;births=$data.prefecture.births;deaths=$data.prefecture.deaths
                inflow=$data.prefecture.inflow;outflow=$data.prefecture.outflow
                households=$data.prefecture.households;hhChange=$data.prefecture.hhChange
            }
            foreach ($city in $validCities) {
                if ($data.cities.Contains($city)) {
                    $citiesData[$city][$ym] = [ordered]@{
                        pop=$data.cities[$city].pop;male=$data.cities[$city].male;female=$data.cities[$city].female
                        change=$data.cities[$city].change;births=$data.cities[$city].births;deaths=$data.cities[$city].deaths
                        inflow=$data.cities[$city].inflow;outflow=$data.cities[$city].outflow
                        households=$data.cities[$city].households;hhChange=$data.cities[$city].hhChange
                    }
                }
            }
            $available += $ym; Write-Host -NoNewline 'e'
        } else { Write-Host -NoNewline '?' }
    } catch { Write-Host -NoNewline '!'; Write-Error $_ }
}

# Excel 終了
if ($null -ne $script:xlApp) {
    $script:xlApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:xlApp) | Out-Null
    $script:xlApp = $null
}
Write-Host ""
Write-Host "Months: $($available.Count)"

# =====================================================================
#  JSON 出力
# =====================================================================
function NullOrVal($v) { if ($null -eq $v) { 'null' } else { [string]$v } }
function Obj2Json($d) {
    '{"pop":'+$d.pop+',"male":'+$d.male+',"female":'+$d.female+
    ',"change":'+$d.change+',"births":'+(NullOrVal $d.births)+',"deaths":'+(NullOrVal $d.deaths)+
    ',"inflow":'+(NullOrVal $d.inflow)+',"outflow":'+(NullOrVal $d.outflow)+
    ',"households":'+$d.households+',"hhChange":'+(NullOrVal $d.hhChange)+'}'
}

$prefJson   = ($available | Where-Object { $prefData.Contains($_) } |
               ForEach-Object { '"' + $_ + '":' + (Obj2Json $prefData[$_]) }) -join ','
$citiesJson = ($validCities |
               ForEach-Object {
                   $city = $_
                   $inner = ($citiesData[$city].Keys | ForEach-Object { '"' + $_ + '":' + (Obj2Json $citiesData[$city][$_]) }) -join ','
                   '"' + $city + '":{' + $inner + '}'
               }) -join ','
$monthsJson = '"' + ($available -join '","') + '"'
$latest     = $available[-1]
$generated  = Get-Date -Format 'yyyy-MM-dd'

$js = @"
// 群馬県移動人口調査
// 生成日: $generated
// 出典: https://toukei.pref.gunma.jp/idj/
window.POP_DATA = {"meta":{"generated":"$generated","source":"https://toukei.pref.gunma.jp/idj/","latest":"$latest"},"months":[$monthsJson],"prefecture":{$prefJson},"cities":{$citiesJson}};
"@

[IO.File]::WriteAllText($outFile, $js, [Text.Encoding]::UTF8)
Write-Host "Saved: $outFile ($([Math]::Round((Get-Item $outFile).Length/1024))KB)"