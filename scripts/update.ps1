# 群馬県移動人口調査 差分更新スクリプト (GitHub Actions用)
# 既存の population.js を読み込み、新しい月のCSVのみ追加する
$ErrorActionPreference = 'Continue'

$baseUrl = 'https://toukei.pref.gunma.jp/idj/data'
$root    = Split-Path $PSScriptRoot -Parent
$outFile = "$root/data/population.js"

# =====================================================================
#  既存データ読み込み
# =====================================================================
$existing = [IO.File]::ReadAllText($outFile, [Text.Encoding]::UTF8)
# window.POP_DATA = {...}; から JSON 部分を抽出
$jsonStr = $existing -replace '^[\s\S]*?window\.POP_DATA\s*=\s*', '' -replace ';\s*$', ''
$data = $jsonStr | ConvertFrom-Json

$latestYm  = $data.meta.latest   # 例: "202605"
$latestY   = [int]$latestYm.Substring(0,4)
$latestM   = [int]$latestYm.Substring(4,2)

# 次月から今月までを対象に
$today  = Get-Date
$endY   = $today.Year
$endM   = $today.Month
$newMonths = @()
$y = $latestY; $m = $latestM + 1
if ($m -gt 12) { $m = 1; $y++ }
while ($y -lt $endY -or ($y -eq $endY -and $m -le $endM)) {
    $newMonths += '{0}{1:D2}' -f $y, $m
    $m++
    if ($m -gt 12) { $m = 1; $y++ }
}

if ($newMonths.Count -eq 0) {
    Write-Host "No new months to update."
    exit 0
}
Write-Host "Checking: $($newMonths -join ', ')"

# =====================================================================
#  CSV パーサ
# =====================================================================
$validCities = @(
    '前橋市','高崎市','桐生市','伊勢崎市','太田市','沼田市','館林市','渋川市','藤岡市','富岡市','安中市','みどり市',
    '榛東村','吉岡町','上野村','神流町','下仁田町','南牧村','甘楽町','中之条町','長野原町','嬬恋村','草津町','高山村','東吾妻町',
    '片品村','川場村','昭和村','みなかみ町','玉村町','板倉町','明和町','千代田町','大泉町','邑楽町'
)
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
#  新月データ取得
# =====================================================================
$added = @()
foreach ($ym in $newMonths) {
    $tmpFile = [IO.Path]::GetTempFileName() + '.csv'
    try {
        Invoke-WebRequest -Uri "$baseUrl/idj$ym.csv" -OutFile $tmpFile -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "$ym : not available yet"
        continue
    }
    $raw    = [IO.File]::ReadAllBytes($tmpFile)
    $hasBom = $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF
    $enc    = if ($hasBom) { [Text.Encoding]::UTF8 } else { [Text.Encoding]::GetEncoding('shift_jis') }
    $lines  = $enc.GetString($raw).Replace("`r`n","`n").Replace("`r","`n").Split("`n")
    Remove-Item $tmpFile -ErrorAction SilentlyContinue

    $prefEntry = $null
    $cityEntries = @{}
    foreach ($line in $lines) {
        if (-not $line.Trim()) { continue }
        $cols = Parse-CsvLine $line
        if ($cols.Count -lt 12) { continue }
        $name = $cols[0].Trim() -replace '\s+', ''
        if ($null -eq $prefEntry -and $name -match '^県+計$') {
            $pop = Parse-Num $cols[1]
            if ($pop -gt 0) {
                $prefEntry = [PSCustomObject]@{
                    pop=(Parse-Num $cols[1]);male=(Parse-Num $cols[2]);female=(Parse-Num $cols[3])
                    change=(Parse-Num $cols[4]);births=(Parse-Num $cols[6]);deaths=(Parse-Num $cols[7])
                    inflow=(Parse-Num $cols[9]);outflow=(Parse-Num $cols[10])
                    households=(Parse-Num $cols[11]);hhChange=(Parse-Num $cols[12])
                }
            }
        }
        if ($validCities -contains $name) {
            $pop = Parse-Num $cols[1]
            if ($pop -gt 0) {
                $cityEntries[$name] = [PSCustomObject]@{
                    pop=$pop;male=(Parse-Num $cols[2]);female=(Parse-Num $cols[3])
                    change=(Parse-Num $cols[4]);births=(Parse-Num $cols[6]);deaths=(Parse-Num $cols[7])
                    inflow=(Parse-Num $cols[9]);outflow=(Parse-Num $cols[10])
                    households=(Parse-Num $cols[11]);hhChange=(Parse-Num $cols[12])
                }
            }
        }
    }

    if ($null -ne $prefEntry) {
        # データをオブジェクトに追記
        $data.prefecture | Add-Member -MemberType NoteProperty -Name $ym -Value $prefEntry -Force
        foreach ($city in $validCities) {
            if ($cityEntries.ContainsKey($city)) {
                $data.cities.$city | Add-Member -MemberType NoteProperty -Name $ym -Value $cityEntries[$city] -Force
            }
        }
        $added += $ym
        Write-Host "$ym : OK"
    } else {
        Write-Host "$ym : parse failed"
    }
}

if ($added.Count -eq 0) {
    Write-Host "Nothing added."
    exit 0
}

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

$allMonths  = ($data.months + $added)
$latest     = $allMonths[-1]
$generated  = Get-Date -Format 'yyyy-MM-dd'

$prefJson = ($allMonths | Where-Object { $null -ne $data.prefecture.$_ } |
    ForEach-Object { '"' + $_ + '":' + (Obj2Json $data.prefecture.$_) }) -join ','

$citiesJson = ($validCities | ForEach-Object {
    $city  = $_
    $inner = ($allMonths | Where-Object { $null -ne $data.cities.$city.$_ } |
        ForEach-Object { '"' + $_ + '":' + (Obj2Json $data.cities.$city.$_) }) -join ','
    '"' + $city + '":{' + $inner + '}'
}) -join ','

$monthsJson = '"' + ($allMonths -join '","') + '"'

$js = @"
// 群馬県移動人口調査
// 生成日: $generated
// 出典: https://toukei.pref.gunma.jp/idj/
window.POP_DATA = {"meta":{"generated":"$generated","source":"https://toukei.pref.gunma.jp/idj/","latest":"$latest"},"months":[$monthsJson],"prefecture":{$prefJson},"cities":{$citiesJson}};
"@

[IO.File]::WriteAllText($outFile, $js, [Text.Encoding]::UTF8)
Write-Host "Updated: $outFile (added: $($added -join ', '))"
