# 群馬県移動人口調査 データダウンロード・集計スクリプト
# 出典: https://toukei.pref.gunma.jp/idj/
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Continue'

$baseUrl  = 'https://toukei.pref.gunma.jp/idj/data'
$cacheDir = 'C:\Users\tkudo\gunma_pop\data\csv'
$outFile  = 'C:\Users\tkudo\gunma_pop\data\population.js'

# 対象年月 (2021/01 - 2026/05)
$months = @()
for ($y = 2021; $y -le 2026; $y++) {
    $maxM = if ($y -eq 2026) { 5 } else { 12 }
    for ($m = 1; $m -le $maxM; $m++) { $months += '{0}{1:D2}' -f $y, $m }
}

# 対象市区町村 (GeoJSONのnameプロパティと一致)
$validCities = @(
    '前橋市','高崎市','桐生市','伊勢崎市','太田市','沼田市','館林市','渋川市','藤岡市','富岡市','安中市','みどり市',
    '榛東村','吉岡町','上野村','神流町','下仁田町','南牧村','甘楽町','中之条町','長野原町','嬬恋村','草津町','高山村','東吾妻町',
    '片品村','川場村','昭和村','みなかみ町','玉村町','板倉町','明和町','千代田町','大泉町','邑楽町'
)

function Parse-Num($s) {
    $c = $s.Trim() -replace '[\s,"]', ''
    if ($c -eq '' -or $c -eq '-') { return 0 }
    try { return [int]$c } catch { return 0 }
}

$prefData   = [ordered]@{}
$citiesData = [ordered]@{}
foreach ($city in $validCities) { $citiesData[$city] = [ordered]@{} }
$available  = @()

foreach ($ym in $months) {
    $csvPath = "$cacheDir\idj$ym.csv"
    if (-not (Test-Path $csvPath)) {
        try {
            Invoke-WebRequest -Uri "$baseUrl/idj$ym.csv" -OutFile $csvPath -UseBasicParsing
            Write-Host "Downloaded: $ym"
        } catch {
            Write-Host "Skip: $ym"
            continue
        }
    }

    $raw   = [IO.File]::ReadAllBytes($csvPath)
    $txt   = [Text.Encoding]::GetEncoding('shift_jis').GetString($raw)
    $lines = $txt.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n")

    $available  += $ym
    $prefFound   = $false

    foreach ($line in $lines) {
        if (-not $line.Trim()) { continue }
        $cols = $line.Split(',')
        if ($cols.Count -lt 12) { continue }
        $name = $cols[0].Trim() -replace '\s+', ''

        # 県計
        if (-not $prefFound -and ($name -eq '県計' -or $name -match '^県+計$')) {
            $pop = Parse-Num $cols[1]
            if ($pop -gt 0) {
                $prefData[$ym] = [ordered]@{
                    pop=($pop); male=(Parse-Num $cols[2]); female=(Parse-Num $cols[3])
                    change=(Parse-Num $cols[4]); births=(Parse-Num $cols[6]); deaths=(Parse-Num $cols[7])
                    inflow=(Parse-Num $cols[9]); outflow=(Parse-Num $cols[10])
                    households=(Parse-Num $cols[11]); hhChange=(Parse-Num $cols[12])
                }
                $prefFound = $true
            }
        }

        # 市区町村
        if ($validCities -contains $name) {
            $pop = Parse-Num $cols[1]
            if ($pop -gt 0) {
                $citiesData[$name][$ym] = [ordered]@{
                    pop=($pop); male=(Parse-Num $cols[2]); female=(Parse-Num $cols[3])
                    change=(Parse-Num $cols[4]); births=(Parse-Num $cols[6]); deaths=(Parse-Num $cols[7])
                    inflow=(Parse-Num $cols[9]); outflow=(Parse-Num $cols[10])
                    households=(Parse-Num $cols[11]); hhChange=(Parse-Num $cols[12])
                }
            }
        }
    }
    Write-Host -NoNewline "."
}
Write-Host ""
Write-Host "Months: $($available.Count)"

# JSON直列化（手動）
function Obj2Json($d) {
    '{"pop":'+$d.pop+',"male":'+$d.male+',"female":'+$d.female+
    ',"change":'+$d.change+',"births":'+$d.births+',"deaths":'+$d.deaths+
    ',"inflow":'+$d.inflow+',"outflow":'+$d.outflow+
    ',"households":'+$d.households+',"hhChange":'+$d.hhChange+'}'
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
