# ============================================================
#  周辺地価公示等抽出ツール Ver.12（Excel連携対応版）
#
#  【貼り付け実行での使い方】
#  設定ブロックの値を書き換えてPowerShellに貼り付けて実行。
#
#  【Excelランチャーからの使い方】
#  Excelが $Config_* 変数を設定してからこのスクリプトを読み込みます。
#  Excelで設定済みの変数は上書きしません。
# ============================================================

# ============================================================
#  設定ブロック
#  ・Excelランチャー使用時: Excelが設定した値をそのまま使用（上書きしない）
#  ・貼り付け実行時: ここの値を使用（必要な箇所だけ書き換えてください）
# ============================================================

# 各変数が未定義の場合のみ初期値を設定する（Excelランチャーとの連携用）
if (-not (Get-Variable 'Config_LandPriceGeoJson' -ErrorAction SilentlyContinue)) {
    # 【必須】地価公示GeoJSONのパス
    $Config_LandPriceGeoJson = "C:\Work\L01.geojson"
}
if (-not (Get-Variable 'Config_ChibanGeoJson' -ErrorAction SilentlyContinue)) {
    # 地番GeoJSONのパス（使わない場合は "" のまま）
    $Config_ChibanGeoJson = ""
}
if (-not (Get-Variable 'Config_Chiban' -ErrorAction SilentlyContinue)) {
    # 検索する地番（例: "123-4"）
    $Config_Chiban = ""
}
if (-not (Get-Variable 'Config_Latitude' -ErrorAction SilentlyContinue)) {
    $Config_Latitude = 0.0
}
if (-not (Get-Variable 'Config_Longitude' -ErrorAction SilentlyContinue)) {
    $Config_Longitude = 0.0
}
if (-not (Get-Variable 'Config_Prefecture' -ErrorAction SilentlyContinue)) {
    # 都道府県名またはコード（例: "青森" / "02"）
    $Config_Prefecture = ""
}
if (-not (Get-Variable 'Config_RadiusMeters' -ErrorAction SilentlyContinue)) {
    $Config_RadiusMeters = 1000
}
if (-not (Get-Variable 'Config_TargetAreaM2' -ErrorAction SilentlyContinue)) {
    $Config_TargetAreaM2 = 0.0
}
if (-not (Get-Variable 'Config_TargetZoningCode' -ErrorAction SilentlyContinue)) {
    $Config_TargetZoningCode = ""
}
if (-not (Get-Variable 'Config_AddressKeyword' -ErrorAction SilentlyContinue)) {
    $Config_AddressKeyword = ""
}
if (-not (Get-Variable 'Config_OutputCsv' -ErrorAction SilentlyContinue)) {
    # 【必須】出力CSVパス
    $Config_OutputCsv = ".\周辺地価公示参考一覧.csv"
}
if (-not (Get-Variable 'Config_OutputEncoding' -ErrorAction SilentlyContinue)) {
    $Config_OutputEncoding = "UTF8BOM"
}
if (-not (Get-Variable 'Config_OutputExcel' -ErrorAction SilentlyContinue)) {
    $Config_OutputExcel = ""
}
if (-not (Get-Variable 'Config_ChibanSelectMode' -ErrorAction SilentlyContinue)) {
    $Config_ChibanSelectMode = ""
}
if (-not (Get-Variable 'Config_SelectedIndex' -ErrorAction SilentlyContinue)) {
    $Config_SelectedIndex = 0
}
if (-not (Get-Variable 'Config_NoPrompt' -ErrorAction SilentlyContinue)) {
    $Config_NoPrompt = $false
}
if (-not (Get-Variable 'Config_OutputAllProperties' -ErrorAction SilentlyContinue)) {
    $Config_OutputAllProperties = $false
}

# ============================================================
#  ここから下は変更不要
# ============================================================
$ErrorActionPreference = "Stop"

#region コードテーブル

$ZONING_MAP = @{
    "1"="第一種低層住居専用地域";"2"="第二種低層住居専用地域"
    "3"="第一種中高層住居専用地域";"4"="第二種中高層住居専用地域"
    "5"="第一種住居地域";"6"="第二種住居地域";"7"="準住居地域";"8"="田園住居地域"
    "9"="近隣商業地域";"10"="商業地域"
    "11"="準工業地域";"12"="工業地域";"13"="工業専用地域"
    "21"="第一種低層住居専用地域";"22"="第二種低層住居専用地域"
    "23"="第一種中高層住居専用地域";"24"="第二種中高層住居専用地域"
    "25"="第一種住居地域";"26"="第二種住居地域";"27"="準住居地域"
    "31"="近隣商業地域";"32"="商業地域"
    "41"="準工業地域";"42"="工業地域";"43"="工業専用地域"
    "51"="市街化調整区域（用途地域なし）";"99"="不明・その他"
}

$ZONING_CATEGORY = @{
    "1"="住宅系";"2"="住宅系";"3"="住宅系";"4"="住宅系"
    "5"="住宅系";"6"="住宅系";"7"="住宅系";"8"="住宅系"
    "9"="商業系";"10"="商業系"
    "11"="工業系";"12"="工業系";"13"="工業系"
    "21"="住宅系";"22"="住宅系";"23"="住宅系";"24"="住宅系"
    "25"="住宅系";"26"="住宅系";"27"="住宅系"
    "31"="商業系";"32"="商業系"
    "41"="工業系";"42"="工業系";"43"="工業系"
}

$USE_MAP = @{
    "00"="未分類";"01"="住宅";"02"="共同住宅";"03"="店舗兼住宅"
    "04"="店舗";"05"="事務所";"06"="店舗・事務所兼住宅";"07"="旅館・ホテル"
    "08"="給油所";"09"="工場・倉庫";"10"="農地";"11"="山林"
    "12"="原野";"13"="雑種地";"14"="空地";"15"="作業場"
    "20"="宅地見込地";"99"="その他"
    "1"="住宅";"2"="共同住宅";"3"="店舗兼住宅";"4"="店舗"
    "5"="事務所";"6"="店舗・事務所兼住宅";"7"="旅館・ホテル"
    "8"="給油所";"9"="工場・倉庫"
}

$PREF_BOUNDS = @{
    "01"=@(41.3,45.6,139.3,145.9);"02"=@(40.2,41.6,139.6,141.7)
    "03"=@(38.9,40.5,140.6,142.1);"04"=@(37.8,39.0,140.2,141.7)
    "05"=@(38.9,40.5,139.7,141.1);"06"=@(37.7,38.9,139.5,140.9)
    "07"=@(36.8,37.9,139.1,141.1);"08"=@(35.7,36.9,139.7,140.9)
    "09"=@(36.2,37.2,139.3,140.3);"10"=@(35.9,36.9,138.4,139.7)
    "11"=@(35.7,36.3,138.7,139.9);"12"=@(34.9,36.1,139.7,140.9)
    "13"=@(24.0,35.9,136.0,139.9);"14"=@(35.1,35.7,138.9,139.8)
    "15"=@(36.8,38.6,137.6,139.6);"16"=@(36.4,37.0,136.7,137.7)
    "17"=@(36.1,37.1,136.1,137.4);"18"=@(35.3,36.3,135.5,136.9)
    "19"=@(35.2,36.0,138.3,139.3);"20"=@(35.2,37.0,136.9,138.7)
    "21"=@(35.1,36.4,136.2,137.7);"22"=@(34.6,35.7,137.5,139.2)
    "23"=@(34.6,35.5,136.7,137.9);"24"=@(33.7,35.3,135.8,136.9)
    "25"=@(34.8,35.7,135.7,136.6);"26"=@(34.7,35.8,135.0,136.2)
    "27"=@(34.3,35.0,135.1,135.9);"28"=@(33.9,35.8,134.2,135.5)
    "29"=@(33.8,35.0,135.6,136.3);"30"=@(33.4,34.5,135.0,136.0)
    "31"=@(35.0,35.8,133.2,134.5);"32"=@(34.3,35.8,131.7,133.5)
    "33"=@(34.2,35.3,133.3,134.6);"34"=@(34.0,35.1,132.1,133.4)
    "35"=@(33.7,35.1,130.6,132.3);"36"=@(33.5,34.5,133.6,134.9)
    "37"=@(33.9,34.7,133.5,134.6);"38"=@(32.9,33.9,132.1,133.7)
    "39"=@(32.7,33.7,132.6,134.3);"40"=@(32.7,34.3,129.5,131.7)
    "41"=@(32.8,34.0,129.3,130.6);"42"=@(31.9,34.8,128.6,130.4)
    "43"=@(32.0,33.2,130.0,131.3);"44"=@(32.7,33.8,130.9,132.3)
    "45"=@(31.3,32.9,130.5,132.1);"46"=@(27.0,32.3,128.3,131.4)
    "47"=@(24.0,27.1,122.9,131.3)
}

$PREF_NAME_TO_CODE = @{
    "北海道"="01"
    "青森"="02";"青森県"="02";"岩手"="03";"岩手県"="03"
    "宮城"="04";"宮城県"="04";"秋田"="05";"秋田県"="05"
    "山形"="06";"山形県"="06";"福島"="07";"福島県"="07"
    "茨城"="08";"茨城県"="08";"栃木"="09";"栃木県"="09"
    "群馬"="10";"群馬県"="10";"埼玉"="11";"埼玉県"="11"
    "千葉"="12";"千葉県"="12";"東京"="13";"東京都"="13"
    "神奈川"="14";"神奈川県"="14";"新潟"="15";"新潟県"="15"
    "富山"="16";"富山県"="16";"石川"="17";"石川県"="17"
    "福井"="18";"福井県"="18";"山梨"="19";"山梨県"="19"
    "長野"="20";"長野県"="20";"岐阜"="21";"岐阜県"="21"
    "静岡"="22";"静岡県"="22";"愛知"="23";"愛知県"="23"
    "三重"="24";"三重県"="24";"滋賀"="25";"滋賀県"="25"
    "京都"="26";"京都府"="26";"大阪"="27";"大阪府"="27"
    "兵庫"="28";"兵庫県"="28";"奈良"="29";"奈良県"="29"
    "和歌山"="30";"和歌山県"="30";"鳥取"="31";"鳥取県"="31"
    "島根"="32";"島根県"="32";"岡山"="33";"岡山県"="33"
    "広島"="34";"広島県"="34";"山口"="35";"山口県"="35"
    "徳島"="36";"徳島県"="36";"香川"="37";"香川県"="37"
    "愛媛"="38";"愛媛県"="38";"高知"="39";"高知県"="39"
    "福岡"="40";"福岡県"="40";"佐賀"="41";"佐賀県"="41"
    "長崎"="42";"長崎県"="42";"熊本"="43";"熊本県"="43"
    "大分"="44";"大分県"="44";"宮崎"="45";"宮崎県"="45"
    "鹿児島"="46";"鹿児島県"="46";"沖縄"="47";"沖縄県"="47"
}
#endregion

#region ログ管理

$Script:LogEntries = [System.Collections.Generic.List[string]]::new()
$Script:Warnings   = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $Script:LogEntries.Add("[$ts][$Level] $Message") | Out-Null
    switch ($Level) {
        "WARN"  {
            Write-Warning $Message
            $Script:Warnings.Add($Message) | Out-Null
        }
        "OK"    { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}

function Export-ProcessLog {
    param([string]$BaseCsvPath, [hashtable]$Summary, [string]$Encoding = "UTF8BOM")
    $logPath = ($BaseCsvPath -replace '\.csv$','') + "_処理ログ.csv"
    # Ver.12 fix: Config_OutputEncoding に追従（結果CSVと文字コードを揃える）
    $enc     = if ($Encoding -eq "SJIS") { [System.Text.Encoding]::GetEncoding(932) }
               else { New-Object System.Text.UTF8Encoding $true }
    $rows    = [System.Collections.Generic.List[object]]::new()
    foreach ($kv in $Summary.GetEnumerator() | Sort-Object Key) {
        $rows.Add([PSCustomObject]@{ 区分="サマリー"; 項目=$kv.Key; 内容=[string]$kv.Value }) | Out-Null
    }
    foreach ($w in $Script:Warnings) {
        $rows.Add([PSCustomObject]@{ 区分="警告"; 項目=""; 内容=$w }) | Out-Null
    }
    foreach ($l in $Script:LogEntries) {
        $rows.Add([PSCustomObject]@{ 区分="ログ"; 項目=""; 内容=$l }) | Out-Null
    }
    $props = @("区分","項目","内容")
    $sb    = New-Object System.Text.StringBuilder
    $null  = $sb.AppendLine(($props | ForEach-Object { "`"$_`"" }) -join ",")
    foreach ($row in $rows) {
        $cells = foreach ($p in $props) {
            $v = $row.$p
            if ($null -eq $v) { '""' } else { '"{0}"' -f ([string]$v).Replace('"','""') }
        }
        $null = $sb.AppendLine($cells -join ",")
    }
    $dir = Split-Path -Parent $logPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($logPath, $sb.ToString(), $enc)
    return $logPath
}
#endregion

#region ユーティリティ関数

function Decode-ZoningCode {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return "" }
    $raw = $Code.Trim(); $s = $raw.TrimStart("0"); if ($s -eq "") { $s = "0" }
    if ($ZONING_MAP.ContainsKey($raw)) { return $ZONING_MAP[$raw] }
    if ($ZONING_MAP.ContainsKey($s))   { return $ZONING_MAP[$s] }
    return ""
}

function Decode-UseCode {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return "" }
    $raw = $Code.Trim(); $s = $raw.TrimStart("0"); if ($s -eq "") { $s = "0" }
    if ($USE_MAP.ContainsKey($raw)) { return $USE_MAP[$raw] }
    if ($USE_MAP.ContainsKey($s))   { return $USE_MAP[$s] }
    return "コード:$raw"
}

function Get-ZoningNormCode {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return "" }
    $s = $Code.Trim().TrimStart("0")
    return if ($s -eq "") { "0" } else { $s }
}

function Convert-ToDoubleOrNull {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $s = ([string]$Value).Trim() -replace ",", ""
    if ($s -eq "") { return $null }
    $n = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return $null
}

function Get-PropValueWithKey {
    param([object]$Properties, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Properties.PSObject.Properties.Name -contains $name) {
            $v = $Properties.$name
            if ($null -ne $v -and ([string]$v).Trim() -ne "") {
                return @{ Value=$v; Key=$name }
            }
        }
    }
    return @{ Value=$null; Key="" }
}

function Get-PropValue {
    param([object]$Properties, [string[]]$Names)
    return (Get-PropValueWithKey -Properties $Properties -Names $Names).Value
}

function Get-HaversineMeters {
    param([double]$Lat1, [double]$Lon1, [double]$Lat2, [double]$Lon2)
    $R = 6371000.0; $rad = [Math]::PI / 180.0
    $dL = ($Lat2-$Lat1)*$rad; $dO = ($Lon2-$Lon1)*$rad
    $a  = [Math]::Sin($dL/2)*[Math]::Sin($dL/2) +
          [Math]::Cos($Lat1*$rad)*[Math]::Cos($Lat2*$rad)*[Math]::Sin($dO/2)*[Math]::Sin($dO/2)
    return $R * 2 * [Math]::Atan2([Math]::Sqrt($a), [Math]::Sqrt(1-$a))
}

function Get-BoundingBox {
    param([double]$Lat, [double]$Lon, [int]$Meters)
    $dLat = $Meters / 111320.0
    $dLon = $Meters / (111320.0 * [Math]::Cos($Lat * [Math]::PI / 180.0))
    return @{ MinLat=$Lat-$dLat; MaxLat=$Lat+$dLat; MinLon=$Lon-$dLon; MaxLon=$Lon+$dLon }
}

function Load-GeoJson {
    param([string]$Path)
    $mb = [Math]::Round((Get-Item $Path).Length / 1MB, 1)
    Write-Log "読込中: $(Split-Path -Leaf $Path) ($mb MB)"
    $text = Get-Content $Path -Raw -Encoding UTF8
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        return $text | ConvertFrom-Json -Depth 100
    } else {
        return $text | ConvertFrom-Json
    }
}

function Get-Median {
    param([double[]]$Values)
    if ($Values.Count -eq 0) { return $null }
    $s = $Values | Sort-Object; $n = $s.Count
    if ($n % 2 -eq 1) { return $s[($n-1)/2] }
    return ($s[$n/2-1] + $s[$n/2]) / 2.0
}
#endregion

#region 実行前チェック

function Invoke-PreflightCheck {
    $errors = [System.Collections.Generic.List[string]]::new()

    if (-not (Test-Path $Config_LandPriceGeoJson)) {
        $errors.Add("地価GeoJSONが見つかりません: $Config_LandPriceGeoJson") | Out-Null
    }
    if ($Config_ChibanGeoJson -ne "" -and -not (Test-Path $Config_ChibanGeoJson)) {
        $errors.Add("地番GeoJSONが見つかりません: $Config_ChibanGeoJson") | Out-Null
    }

    $hasChibanMode = ($Config_ChibanGeoJson -ne "" -and $Config_Chiban -ne "")
    $hasLatLonMode = ($Config_Latitude -ne 0 -or $Config_Longitude -ne 0)
    if (-not $hasChibanMode -and -not $hasLatLonMode) {
        $errors.Add("地番（ChibanGeoJson + Chiban）または緯度経度（Latitude/Longitude）のどちらかを設定してください。") | Out-Null
    }
    if ($Config_ChibanGeoJson -ne "" -and $Config_Chiban -eq "") {
        $errors.Add("Config_ChibanGeoJson を設定した場合は Config_Chiban も必要です。") | Out-Null
    }
    if ($Config_OutputEncoding -notin @("UTF8BOM","SJIS")) {
        $errors.Add("Config_OutputEncoding は UTF8BOM または SJIS を指定してください。") | Out-Null
    }

    # 出力先ディレクトリ作成
    foreach ($outPath in @($Config_OutputCsv, $Config_OutputExcel) | Where-Object { $_ -ne "" }) {
        $dir = Split-Path -Parent $outPath
        if ($dir -ne "" -and -not (Test-Path $dir)) {
            try {
                New-Item -ItemType Directory -Force -Path $dir | Out-Null
                Write-Log "出力先ディレクトリを作成しました: $dir"
            } catch {
                $errors.Add("出力先ディレクトリを作成できません: $dir") | Out-Null
            }
        }
    }

    # Excel出力の依存チェック（警告のみ、エラーにしない）
    if ($Config_OutputExcel -ne "") {
        if (-not (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue)) {
            Write-Log "ImportExcel モジュールが見つかりません。Excel出力はスキップします。" "WARN"
            Write-Log "インストール: Install-Module ImportExcel -Scope CurrentUser" "WARN"
        }
    }

    if ($errors.Count -gt 0) {
        Write-Host ""
        Write-Host "════════════════════════════════════════════" -ForegroundColor Red
        Write-Host "【実行前チェックエラー】" -ForegroundColor Red
        foreach ($e in $errors) {
            Write-Host "  × $e" -ForegroundColor Red
            $Script:LogEntries.Add("[ERROR] $e") | Out-Null
        }
        Write-Host "════════════════════════════════════════════" -ForegroundColor Red
        throw "実行前チェックで $($errors.Count) 件のエラーがありました。設定ブロックを確認してください。"
    }

    Write-Log "実行前チェック: 全項目OK" "OK"
}
#endregion

#region 座標系バリデーション

function Resolve-PrefectureCode {
    param([string]$Input)
    if ([string]::IsNullOrWhiteSpace($Input)) { return $null }
    $s = $Input.Trim()
    $n = 0
    if ([int]::TryParse($s, [ref]$n)) { return $n.ToString("D2") }
    if ($PREF_NAME_TO_CODE.ContainsKey($s)) { return $PREF_NAME_TO_CODE[$s] }
    return $null
}

function Assert-CoordinateValid {
    param([double]$Lat, [double]$Lon, [string]$Source, [string]$PrefCode = "")

    if ($Lat -lt -90 -or $Lat -gt 90 -or $Lon -lt -180 -or $Lon -gt 180) {
        $msg = "【座標系エラー】$Source : 緯度=$Lat 経度=$Lon が WGS84 範囲外。" +
               "公共座標系（平面直角XY）のGeoJSONの可能性があります。" +
               "G空間情報センター版を使用するか、QGISで変換してください。" +
               " → https://front.geospatial.jp/moj-chizu-shp-download/"
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
        Write-Host $msg -ForegroundColor Red
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
        $Script:LogEntries.Add("[ERROR] $msg") | Out-Null
        throw "座標系エラー: 処理を中止します。"
    }

    if ($Lat -lt 20 -or $Lat -gt 50 -or $Lon -lt 122 -or $Lon -gt 154) {
        Write-Log "座標 ($Lat, $Lon) が日本の概略範囲外です。($Source)" "WARN"
        return
    }

    if ($PrefCode -ne "" -and $PREF_BOUNDS.ContainsKey($PrefCode)) {
        $b = $PREF_BOUNDS[$PrefCode]
        if ($Lat -lt $b[0] -or $Lat -gt $b[1] -or $Lon -lt $b[2] -or $Lon -gt $b[3]) {
            $prefName = ($PREF_NAME_TO_CODE.GetEnumerator() |
                         Where-Object { $_.Value -eq $PrefCode } |
                         Select-Object -First 1).Key
            $warnMsg = ("座標 ($Lat, $Lon) が {0} の概略範囲外です。($Source)  " +
                        "県概略範囲: 緯度{1}〜{2} 経度{3}〜{4}  " +
                        "別の県のデータを指定していないか確認してください。") `
                       -f $prefName, $b[0], $b[1], $b[2], $b[3]
            Write-Log $warnMsg "WARN"
        }
    }
}

function Get-FirstValidCoordinate {
    param([object]$GeoJson)
    if ($null -eq $GeoJson.features -or $GeoJson.features.Count -eq 0) { return $null }
    foreach ($f in $GeoJson.features | Select-Object -First 20) {
        $g = $f.geometry; if ($null -eq $g) { continue }
        $coords = $null
        if     ($g.type -eq "Point")        { $coords = $g.coordinates }
        elseif ($g.type -eq "Polygon")      { $coords = @($g.coordinates[0])[0] }
        elseif ($g.type -eq "MultiPolygon") { $coords = @(@($g.coordinates[0])[0])[0] }
        if ($null -ne $coords -and $coords.Count -ge 2) {
            return @{ Lon=[double]$coords[0]; Lat=[double]$coords[1] }
        }
    }
    return $null
}
#endregion

#region Polygon面積重心計算
# Shoelace + Centroid公式。穴あきポリゴン（内側リング）は考慮しません。

function Get-PolygonCentroid {
    param([double[]]$Lons, [double[]]$Lats)
    $n = $Lons.Count; if ($n -lt 3) { return $null }
    $area = 0.0; $cx = 0.0; $cy = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $j = ($i+1) % $n
        $cross = $Lons[$i]*$Lats[$j] - $Lons[$j]*$Lats[$i]
        $area += $cross
        $cx   += ($Lons[$i]+$Lons[$j]) * $cross
        $cy   += ($Lats[$i]+$Lats[$j]) * $cross
    }
    $area *= 0.5
    if ([Math]::Abs($area) -lt 1e-12) { return $null }
    return @{ Lon=($cx/(6.0*$area)); Lat=($cy/(6.0*$area)) }
}

function Get-RingCentroidAndArea {
    param([object]$Ring)
    $lons = [System.Collections.Generic.List[double]]::new()
    $lats = [System.Collections.Generic.List[double]]::new()
    foreach ($pt in @($Ring)) {
        $arr = @($pt)
        if ($arr.Count -ge 2) {
            $lons.Add([double]$arr[0]) | Out-Null
            $lats.Add([double]$arr[1]) | Out-Null
        }
    }
    if ($lons.Count -lt 3) { return $null }
    $lonArr = $lons.ToArray(); $latArr = $lats.ToArray()
    $area = 0.0
    for ($i = 0; $i -lt $lonArr.Count; $i++) {
        $j = ($i+1) % $lonArr.Count
        $area += $lonArr[$i]*$latArr[$j] - $lonArr[$j]*$latArr[$i]
    }
    $c = Get-PolygonCentroid -Lons $lonArr -Lats $latArr
    return @{ Centroid=$c; AbsArea=[Math]::Abs($area*0.5) }
}

function Get-PointCoordinatesWithMethod {
    param([object]$Geometry)
    if ($null -eq $Geometry) { return $null }

    if ($Geometry.type -eq "Point") {
        $c = $Geometry.coordinates
        if ($null -eq $c -or $c.Count -lt 2) { return $null }
        return @{ Lon=[double]$c[0]; Lat=[double]$c[1]; Method="Point" }
    }

    if ($Geometry.type -eq "Polygon") {
        $r = Get-RingCentroidAndArea -Ring $Geometry.coordinates[0]
        if ($null -eq $r -or $null -eq $r.Centroid) { return $null }
        return @{ Lon=$r.Centroid.Lon; Lat=$r.Centroid.Lat; Method="Polygon面積重心" }
    }

    if ($Geometry.type -eq "MultiPolygon") {
        $best = $null; $bestArea = 0.0; $cnt = 0
        foreach ($poly in $Geometry.coordinates) {
            $cnt++
            $r = Get-RingCentroidAndArea -Ring $poly[0]
            if ($null -ne $r -and $r.AbsArea -gt $bestArea) {
                $bestArea = $r.AbsArea; $best = $r
            }
        }
        if ($null -eq $best -or $null -eq $best.Centroid) { return $null }
        return @{ Lon=$best.Centroid.Lon; Lat=$best.Centroid.Lat
                  Method="MultiPolygon最大面積採用(全${cnt}筆中)" }
    }

    return $null
}
#endregion

#region 地番正規化・検索・候補CSV出力
# 地番フィールド専用の正規化。住所フィールドには適用しない。

function Normalize-Chiban {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return "" }
    # 全角数字→半角
    $s = $Raw `
        -replace "０","0" -replace "１","1" -replace "２","2" -replace "３","3" -replace "４","4" `
        -replace "５","5" -replace "６","6" -replace "７","7" -replace "８","8" -replace "９","9"
    # 地番の区切り記号を正規化（地番フィールド専用変換）
    $s = $s -replace "番地の?", "-" -replace "番の?", "-"
    # 「数字の数字」パターンのみ「の」を「-」に変換（住所の「みどりの一丁目」を壊さないための限定変換）
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, '(\d)の(\d)', '$1-$2')
    # 空白・末尾「地」除去
    $s = $s -replace "\s", "" -replace "地$", ""
    return $s.Trim("-")
}

$CHIBAN_FIELDS  = @("地番","chiban","CHIBAN","筆地番","代表地番","fude_chiban","lot_number")
$ADDRESS_FIELDS = @("所在","住所","所在地","location","ADDRESS","大字","丁目","町名")

function Find-TargetLand {
    param([object]$GeoJson, [string]$ChibanQuery, [string]$AddrKeyword)
    if ($null -eq $GeoJson.features) { throw "地番GeoJSONに 'features' が見つかりません。" }

    $nq         = Normalize-Chiban $ChibanQuery
    $exactList  = [System.Collections.Generic.List[object]]::new()
    $prefixList = [System.Collections.Generic.List[object]]::new()

    foreach ($f in $GeoJson.features) {
        $cv = Get-PropValue $f.properties $CHIBAN_FIELDS
        if ($null -eq $cv) { continue }
        $cn     = Normalize-Chiban ([string]$cv)
        $isE    = ($cn -eq $nq)
        $isP    = (-not $isE -and $cn.StartsWith($nq + "-"))
        if (-not $isE -and -not $isP) { continue }

        # 住所キーワードフィルタ（住所フィールドは変換なしでそのまま検索）
        if ($AddrKeyword -ne "") {
            $av = Get-PropValue $f.properties $ADDRESS_FIELDS
            $as = if ($null -ne $av) { [string]$av } else { "" }
            if ($as -notlike "*$AddrKeyword*") { continue }
        }

        if ($isE) { $exactList.Add($f)  | Out-Null }
        if ($isP) { $prefixList.Add($f) | Out-Null }
    }

    if ($exactList.Count  -gt 0) { return @{ Matches=$exactList;  MatchType="完全一致";             IsPrefix=$false } }
    if ($prefixList.Count -gt 0) { return @{ Matches=$prefixList; MatchType="枝番候補（前方一致）"; IsPrefix=$true  } }
    return @{ Matches=[System.Collections.Generic.List[object]]::new(); MatchType="なし"; IsPrefix=$false }
}

function Get-LandCenterAndArea {
    param([object]$Feature)
    $p = $Feature.properties
    $repLat = Convert-ToDoubleOrNull (Get-PropValue $p @("代表点緯度","rep_lat","latitude","LAT"))
    $repLon = Convert-ToDoubleOrNull (Get-PropValue $p @("代表点経度","rep_lon","longitude","LON"))
    $ptm = $null
    if ($null -ne $repLat -and $null -ne $repLon -and $repLat -ne 0 -and $repLon -ne 0) {
        $ptm = @{ Lat=$repLat; Lon=$repLon; Method="代表点（properties）" }
    } else {
        $ptm = Get-PointCoordinatesWithMethod $Feature.geometry
    }
    $area = Convert-ToDoubleOrNull (Get-PropValue $p @("地積","面積","AREA","area","chiseki","地積_m2","登記地積","現況地積"))
    if ($null -ne $ptm) { return @{ Lat=$ptm.Lat; Lon=$ptm.Lon; AreaM2=$area; Method=$ptm.Method } }
    return $null
}

function Export-ChibanCandidates {
    param([object[]]$Candidates, [string]$Path, [string]$MatchType, [string]$Encoding = "UTF8BOM")
    $rows = for ($i = 0; $i -lt $Candidates.Count; $i++) {
        $mp = $Candidates[$i].properties
        [PSCustomObject]@{
            選択番号 = $i + 1
            一致種別 = $MatchType
            地番     = [string](Get-PropValue $mp $CHIBAN_FIELDS)
            所在     = [string](Get-PropValue $mp $ADDRESS_FIELDS)
            地積_m2  = Convert-ToDoubleOrNull (Get-PropValue $mp @("地積","面積","AREA","area"))
        }
    }
    # Ver.12 fix: Config_OutputEncoding に追従（結果CSVと文字コードを揃える）
    $enc   = if ($Encoding -eq "SJIS") { [System.Text.Encoding]::GetEncoding(932) }
             else { New-Object System.Text.UTF8Encoding $true }
    $props = $rows[0].PSObject.Properties.Name
    $sb    = New-Object System.Text.StringBuilder
    $null  = $sb.AppendLine(($props | ForEach-Object { "`"$_`"" }) -join ",")
    foreach ($row in $rows) {
        $cells = foreach ($p in $props) {
            $v = $row.$p
            if ($null -eq $v) { '""' } else { '"{0}"' -f ([string]$v).Replace('"','""') }
        }
        $null = $sb.AppendLine($cells -join ",")
    }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), $enc)
}
#endregion

#region 条件メモ・差分メモ生成

function Get-DistanceRank {
    param([double]$Dist, [int]$Radius)
    $p = $Dist / $Radius
    if    ($p -le 0.25) { return "近接" }
    elseif($p -le 0.50) { return "近傍" }
    elseif($p -le 0.75) { return "中距離" }
    else                { return "遠距離" }
}

function Build-ConditionMemo {
    param([double]$Dist, [int]$Radius, [string]$ZoningCode, [string]$TargetZoning,
          [double]$AreaM2, [double]$TargetArea)
    $parts = @("距離:$(Get-DistanceRank -Dist $Dist -Radius $Radius)($([Math]::Round($Dist,0))m)")
    if ($TargetZoning -ne "" -and $ZoningCode -ne "") {
        $zc = Get-ZoningNormCode $ZoningCode; $tz = Get-ZoningNormCode $TargetZoning
        if ($zc -eq $tz) { $parts += "用途:同一地域" }
        else {
            $zcat = if ($ZONING_CATEGORY.ContainsKey($zc)) { $ZONING_CATEGORY[$zc] } else { "不明" }
            $tcat = if ($ZONING_CATEGORY.ContainsKey($tz)) { $ZONING_CATEGORY[$tz] } else { "不明" }
            if ($zcat -eq $tcat) { $parts += "用途:系統内別地域" }
            else                 { $parts += "用途:系統相違($zcat)" }
        }
    } elseif ($TargetZoning -ne "") { $parts += "用途:（データなし）" }
    if ($TargetArea -gt 0 -and $AreaM2 -gt 0) {
        $parts += "地積:差$('{0:P0}' -f ([Math]::Abs($AreaM2-$TargetArea)/$TargetArea))"
    } elseif ($TargetArea -gt 0) { $parts += "地積:（データなし）" }
    return ($parts -join " / ")
}

function Build-DiffMemo {
    param([string]$ZoningCode, [string]$TargetZoning, [double]$AreaM2, [double]$TargetArea)
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($TargetZoning -ne "" -and $ZoningCode -ne "") {
        $zc   = Get-ZoningNormCode $ZoningCode; $tz = Get-ZoningNormCode $TargetZoning
        $zcLbl = Decode-ZoningCode $ZoningCode;  $tzLbl = Decode-ZoningCode $TargetZoning
        if ($zcLbl -eq "") { $zcLbl = "コード:$ZoningCode" }
        if ($tzLbl -eq "") { $tzLbl = "コード:$TargetZoning" }
        if ($zc -eq $tz)  { $parts.Add("用途:対象=$tzLbl / 参考=同一") | Out-Null }
        else               { $parts.Add("用途:対象=$tzLbl → 参考=$zcLbl") | Out-Null }
    }
    if ($TargetArea -gt 0 -and $AreaM2 -gt 0) {
        $diff  = $AreaM2 - $TargetArea
        $ratio = $diff / $TargetArea
        $sign  = if ($diff -ge 0) { "+" } else { "" }
        $parts.Add("地積:対象${TargetArea}㎡ → 参考${AreaM2}㎡(${sign}$('{0:P0}' -f $ratio))") | Out-Null
    } elseif ($TargetArea -gt 0) {
        $parts.Add("地積:対象${TargetArea}㎡ → 参考データなし") | Out-Null
    }
    if ($parts.Count -eq 0) { return "" }
    return ($parts -join " / ")
}
#endregion

#region CSV出力・サマリー表示

function Write-CsvEncoded {
    param([object[]]$Data, [string]$Path, [string]$Encoding)
    $enc = if ($Encoding -eq "SJIS") { [System.Text.Encoding]::GetEncoding(932) }
           else { New-Object System.Text.UTF8Encoding $true }
    if ($null -eq $Data -or $Data.Count -eq 0) {
        [System.IO.File]::WriteAllText($Path, "", $enc); return
    }
    $props = $Data[0].PSObject.Properties.Name
    $sb    = New-Object System.Text.StringBuilder
    $null  = $sb.AppendLine(($props | ForEach-Object { "`"$_`"" }) -join ",")
    foreach ($row in $Data) {
        $cells = foreach ($prop in $props) {
            $v = $row.$prop
            if ($null -eq $v) { '""' } else { '"{0}"' -f ([string]$v).Replace('"','""') }
        }
        $null = $sb.AppendLine($cells -join ",")
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), $enc)
}

function Show-PriceSummary {
    param([string]$Label, [double[]]$Prices)
    if ($null -eq $Prices -or $Prices.Count -eq 0) {
        Write-Host ("  {0,-14}: データなし" -f $Label) -ForegroundColor DarkYellow; return
    }
    $avg = ($Prices | Measure-Object -Average).Average
    $min = ($Prices | Measure-Object -Minimum).Minimum
    $max = ($Prices | Measure-Object -Maximum).Maximum
    $med = Get-Median -Values $Prices
    Write-Host ("  {0,-14}: {1,3}件  最小{2,9:N0}  中央値{3,9:N0}  平均{4,9:N0}  最大{5,9:N0}  円/㎡" -f `
        $Label, $Prices.Count, $min, $med, $avg, $max) -ForegroundColor Yellow
}
#endregion

#region メイン処理

$sw = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "===== 周辺地価公示抽出ツール Ver.12 開始 ====="
Write-Log "実行日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# 実行前一括チェック
Invoke-PreflightCheck

# 都道府県コード解決
$prefCode = ""
if ($Config_Prefecture -ne "") {
    $prefCode = Resolve-PrefectureCode -Input $Config_Prefecture
    if ($null -eq $prefCode) {
        Write-Log "都道府県 '$Config_Prefecture' を認識できません。県別範囲チェックをスキップします。" "WARN"
        $prefCode = ""
    } else {
        $prefName = ($PREF_NAME_TO_CODE.GetEnumerator() |
                     Where-Object { $_.Value -eq $prefCode } |
                     Select-Object -First 1).Key
        Write-Log "都道府県: $prefName（コード: $prefCode）"
    }
}

$centerLat     = $Config_Latitude
$centerLon     = $Config_Longitude
$targetArea    = $Config_TargetAreaM2
$chibanInfo    = ""
$centerMethod  = if ($Config_Latitude -ne 0 -or $Config_Longitude -ne 0) { "手動指定" } else { "" }
$selectedChiban = ""

# ── 地番GeoJSONモード
if ($Config_ChibanGeoJson -ne "") {

    $chibanJson  = Load-GeoJson -Path $Config_ChibanGeoJson
    $sampleCoord = Get-FirstValidCoordinate -GeoJson $chibanJson
    if ($null -ne $sampleCoord) {
        Write-Log ("地番GeoJSONサンプル座標: Lon={0:F6}, Lat={1:F6}" -f $sampleCoord.Lon, $sampleCoord.Lat)
        Assert-CoordinateValid -Lat $sampleCoord.Lat -Lon $sampleCoord.Lon `
            -Source "地番GeoJSON" -PrefCode $prefCode
    }

    $addrMsg = if ($Config_AddressKeyword -ne "") { "（所在キーワード:'$Config_AddressKeyword'）" } else { "" }
    Write-Log "地番検索: '$Config_Chiban'（正規化:'$(Normalize-Chiban $Config_Chiban)'）$addrMsg"

    $result  = Find-TargetLand -GeoJson $chibanJson `
        -ChibanQuery $Config_Chiban -AddrKeyword $Config_AddressKeyword
    $matched = $result.Matches

    if ($matched.Count -eq 0) {
        $hint = if ($Config_AddressKeyword -ne "") {
            " Config_AddressKeyword を外すか別のキーワードを試してください。"
        } else { "" }
        throw "地番 '$Config_Chiban' が見つかりませんでした。表記（例: 123-4, 123番4）を確認してください。$hint"
    }

    if ($result.IsPrefix) {
        Write-Log "完全一致なし。枝番候補（前方一致）: $($matched.Count) 件" "WARN"
    } else {
        Write-Log "完全一致: $($matched.Count) 件"
    }

    # 候補CSV出力モード
    if ($Config_ChibanSelectMode -ne "") {
        Export-ChibanCandidates -Candidates $matched -Path $Config_ChibanSelectMode `
            -MatchType $result.MatchType -Encoding $Config_OutputEncoding
        Write-Log "候補CSV出力: $Config_ChibanSelectMode  ($($matched.Count) 件)" "OK"
        Write-Host ""
        Write-Host "候補CSVを確認後、Config_SelectedIndex に番号を設定して再実行してください。" -ForegroundColor Yellow
        Write-Host "（Config_ChibanSelectMode は空文字に戻してください）" -ForegroundColor Yellow
        return
    }

    # 選択処理
    $selectedFeature = $null
    $needSelection   = ($matched.Count -gt 1) -or $result.IsPrefix

    if ($Config_SelectedIndex -gt 0) {
        if ($Config_SelectedIndex -gt $matched.Count) {
            throw "Config_SelectedIndex $Config_SelectedIndex は範囲外です。（候補数: $($matched.Count)）"
        }
        $selectedFeature = $matched[$Config_SelectedIndex - 1]
        Write-Log "選択番号 $Config_SelectedIndex を使用"

    } elseif (-not $needSelection) {
        $selectedFeature = $matched[0]

    } else {
        # Config_NoPrompt = $true の場合は対話に入らず候補CSVを自動出力して終了
        if ($Config_NoPrompt) {
            $autoCandPath = ($Config_OutputCsv -replace '\.csv$','') + "_地番候補.csv"
            Export-ChibanCandidates -Candidates $matched -Path $autoCandPath `
                -MatchType $result.MatchType -Encoding $Config_OutputEncoding
            Write-Log "NoPromptモード: 複数候補のため候補CSVを出力して終了 → $autoCandPath" "WARN"
            Write-Host "候補CSV: $autoCandPath" -ForegroundColor Yellow
            Write-Host "Config_SelectedIndex に番号を設定して再実行してください。" -ForegroundColor Yellow
            return
        }

        # 対話選択
        $maxShow = [Math]::Min($matched.Count, 30)
        Write-Host "`n番号を入力して選択してください:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $maxShow; $i++) {
            $mp    = $matched[$i].properties
            $cv    = Get-PropValue $mp $CHIBAN_FIELDS
            $sz    = Get-PropValue $mp $ADDRESS_FIELDS
            $ar    = Convert-ToDoubleOrNull (Get-PropValue $mp @("地積","面積","AREA","area"))
            $arStr = if ($null -ne $ar) { "地積:${ar}㎡" } else { "" }
            Write-Host ("  [{0,2}] 地番:{1,-12}  所在:{2,-30}  {3}" -f ($i+1), $cv, $sz, $arStr) `
                -ForegroundColor Yellow
        }
        if ($matched.Count -gt 30) {
            $moreMsg = ("  ... 他 {0} 件。Config_AddressKeyword で絞り込むか " +
                        "Config_ChibanSelectMode を使用してください。") -f ($matched.Count - 30)
            Write-Host $moreMsg -ForegroundColor DarkYellow
        }
        Write-Host ""
        $inputStr = Read-Host "番号を入力（1〜$maxShow、Enterでキャンセル）"
        if ([string]::IsNullOrWhiteSpace($inputStr)) { throw "選択がキャンセルされました。" }
        $idxSel = 0
        if (-not [int]::TryParse($inputStr, [ref]$idxSel) -or
            $idxSel -lt 1 -or $idxSel -gt $maxShow) {
            throw "無効な番号です: $inputStr"
        }
        $selectedFeature = $matched[$idxSel - 1]
        Write-Log "対話選択: 番号 $idxSel"
    }

    $landInfo = Get-LandCenterAndArea -Feature $selectedFeature
    if ($null -eq $landInfo) { throw "対象地の座標を取得できませんでした。" }

    Assert-CoordinateValid -Lat $landInfo.Lat -Lon $landInfo.Lon `
        -Source "対象地座標" -PrefCode $prefCode

    $centerLat    = $landInfo.Lat
    $centerLon    = $landInfo.Lon
    $centerMethod = $landInfo.Method

    if ($targetArea -eq 0 -and $null -ne $landInfo.AreaM2 -and $landInfo.AreaM2 -gt 0) {
        $targetArea = $landInfo.AreaM2
        Write-Log ("地積を地番GeoJSONから取得: {0:N1} ㎡" -f $targetArea)
    }

    $cp             = $selectedFeature.properties
    $selectedChiban = [string](Get-PropValue $cp $CHIBAN_FIELDS)
    $chibanInfo     = "地番:$selectedChiban  所在:$(Get-PropValue $cp $ADDRESS_FIELDS)  座標取得:$($landInfo.Method)"
    Write-Log "対象地: $chibanInfo" "OK"
    Write-Log "中心座標: ($centerLat, $centerLon)" "OK"

} else {
    Assert-CoordinateValid -Lat $centerLat -Lon $centerLon `
        -Source "手動指定座標" -PrefCode $prefCode
    Write-Log "中心座標（手動）: ($centerLat, $centerLon)"
}

# ── 地価GeoJSON読み込み・座標系チェック
$priceJson = Load-GeoJson -Path $Config_LandPriceGeoJson
if ($null -eq $priceJson.features) {
    throw ("地価GeoJSONに 'features' が見つかりません。" +
           "GeoJSON形式かどうか確認してください（対応形式: FeatureCollection）。")
}

# 地価公示GeoJSONの形式簡易チェック（最初の5件でL01系属性の存在を確認）
$l01Found = $false
$sampleProps = @()
foreach ($f in $priceJson.features | Select-Object -First 5) {
    if ($null -ne $f.properties) { $sampleProps += $f.properties }
}
foreach ($sp in $sampleProps) {
    $keys = @($sp.PSObject.Properties.Name)
    if ($keys | Where-Object { $_ -like "L01_*" -or $_ -eq "価格" -or $_ -eq "公示価格" }) {
        $l01Found = $true; break
    }
}
if (-not $l01Found -and $sampleProps.Count -gt 0) {
    $sampleKeys = ($sampleProps[0].PSObject.Properties.Name | Select-Object -First 8) -join ", "
    Write-Log ("地価公示GeoJSONの標準属性（L01_*）が見つかりません。" +
               "地価公示以外のGeoJSONを指定している可能性があります。" +
               "先頭フィーチャの属性: $sampleKeys") "WARN"
    Write-Log "処理を続行しますが、価格・用途地域等が正しく取得できない可能性があります。" "WARN"
}

$priceSample = Get-FirstValidCoordinate -GeoJson $priceJson
if ($null -ne $priceSample) {
    Assert-CoordinateValid -Lat $priceSample.Lat -Lon $priceSample.Lon `
        -Source "地価GeoJSON" -PrefCode $prefCode
}

$total  = $priceJson.features.Count
Write-Log ("地価データ: {0:N0} 件読込完了 ({1:F1} 秒)" -f $total, $sw.Elapsed.TotalSeconds)

$bb          = Get-BoundingBox -Lat $centerLat -Lon $centerLon -Meters $Config_RadiusMeters
$results     = [System.Collections.Generic.List[object]]::new()
$prgInt      = [Math]::Max(1, [int]($total / 50))
$idx         = 0
$hasZonCond  = ($Config_TargetZoningCode -ne "")
$hasAreaCond = ($targetArea -gt 0)

foreach ($feature in $priceJson.features) {
    $idx++
    if ($idx % $prgInt -eq 0 -or $idx -eq $total) {
        Write-Progress -Activity "地価データ処理中" `
            -Status ("{0:N0}/{1:N0} | 抽出:{2}件" -f $idx, $total, $results.Count) `
            -PercentComplete ([int]($idx * 100 / $total))
    }

    $ptm = Get-PointCoordinatesWithMethod $feature.geometry
    if ($null -eq $ptm) { continue }
    $ptLat = $ptm.Lat; $ptLon = $ptm.Lon

    if ($ptLat -lt $bb.MinLat -or $ptLat -gt $bb.MaxLat -or
        $ptLon -lt $bb.MinLon -or $ptLon -gt $bb.MaxLon) { continue }

    $dist = Get-HaversineMeters -Lat1 $centerLat -Lon1 $centerLon -Lat2 $ptLat -Lon2 $ptLon
    if ($dist -gt $Config_RadiusMeters) { continue }

    $p      = $feature.properties
    $priceR = Get-PropValueWithKey $p @("L01_008","価格","公示価格","price")
    $zoneR  = Get-PropValueWithKey $p @("L01_027","L01_020","用途地域","zoning")
    $useR   = Get-PropValueWithKey $p @("L01_025","L01_009","利用現況","用途","use")
    $areaR  = Get-PropValueWithKey $p @("L01_011","地積","面積","area")

    $priceNum = Convert-ToDoubleOrNull $priceR.Value
    $areaNum  = Convert-ToDoubleOrNull $areaR.Value
    $zoneCode = if ($null -ne $zoneR.Value) { [string]$zoneR.Value } else { "" }
    $useCode  = if ($null -ne $useR.Value)  { [string]$useR.Value }  else { "" }

    $areaForCalc = if ($null -ne $areaNum) { $areaNum } else { 0.0 }

    $condMemo = Build-ConditionMemo `
        -Dist $dist -Radius $Config_RadiusMeters `
        -ZoningCode $zoneCode -TargetZoning $Config_TargetZoningCode `
        -AreaM2 $areaForCalc -TargetArea $targetArea

    $diffMemo = Build-DiffMemo `
        -ZoningCode $zoneCode -TargetZoning $Config_TargetZoningCode `
        -AreaM2 $areaForCalc -TargetArea $targetArea

    # ── 出力列の構成
    # 【基本列】まず見る列（左側）
    # 距離m / 条件メモ / 差分メモ / 所在地 / 価格_円m2 / 地積_m2 / 用途地域 / 利用現況 / 調査年次
    # 【検証列】データ確認用（右側）
    # 標準地番号 / 調査基準日 / 用途地域_元値 / 価格_取得列 / 用途地域_取得列 / 利用現況_取得列 / 座標取得方法 / 緯度 / 経度
    # 【任意列】Config_OutputAllProperties = $true のときのみ出力
    # 全属性JSON

    $row = [ordered]@{
        # ── 基本列（毎回確認する列）
        距離m           = [Math]::Round($dist, 1)
        条件メモ        = $condMemo
        差分メモ        = $diffMemo
        所在地          = [string](Get-PropValue $p @("L01_001","L01_022","所在地","住所","address"))
        価格_円m2       = $priceNum
        地積_m2         = $areaNum
        用途地域        = Decode-ZoningCode -Code $zoneCode
        利用現況        = Decode-UseCode    -Code $useCode
        調査年次        = [string](Get-PropValue $p @("L01_005","調査年次","年次","year"))
        # ── 検証列（必要に応じて確認する列）
        標準地番号      = [string](Get-PropValue $p @("L01_002","標準地番号","id","ID"))
        調査基準日      = [string](Get-PropValue $p @("L01_006","調査基準日","基準日"))
        用途地域_元値   = $zoneCode
        価格_取得列     = $priceR.Key
        用途地域_取得列 = $zoneR.Key
        利用現況_取得列 = $useR.Key
        座標取得方法    = $ptm.Method
        緯度            = $ptLat
        経度            = $ptLon
    }
    if ($Config_OutputAllProperties) {
        $row["全属性JSON"] = ($p | ConvertTo-Json -Compress -Depth 10)
    }
    $results.Add([PSCustomObject]$row) | Out-Null
}

Write-Progress -Activity "地価データ処理中" -Completed

if ($results.Count -eq 0) {
    Write-Log ("抽出件数0件。" +
        "半径不足（現在:$($Config_RadiusMeters)m）、座標ズレ（中心:$centerLat,$centerLon）、" +
        "エリア不一致の可能性があります。") "WARN"
}

# 並び順: 用途地域一致優先 → 距離昇順
$sorted = $results | Sort-Object `
    @{
        Expression = {
            $zc = Get-ZoningNormCode $_.用途地域_元値
            $tz = Get-ZoningNormCode $Config_TargetZoningCode
            if ($Config_TargetZoningCode -eq "") { return 1 }
            if ($zc -eq $tz) { return 0 }
            $zcat = if ($ZONING_CATEGORY.ContainsKey($zc)) { $ZONING_CATEGORY[$zc] } else { "" }
            $tcat = if ($ZONING_CATEGORY.ContainsKey($tz)) { $ZONING_CATEGORY[$tz] } else { "" }
            if ($zcat -ne "" -and $zcat -eq $tcat) { return 1 }
            return 2
        }
        Ascending = $true
    },
    @{ Expression = { [double]$_.距離m }; Ascending = $true }

# CSV 出力
Write-CsvEncoded -Data $sorted -Path $Config_OutputCsv -Encoding $Config_OutputEncoding
$encLabel = if ($Config_OutputEncoding -eq "SJIS") { "Shift_JIS" } else { "UTF-8 BOM付き" }

# Excel 出力（ImportExcel があれば）
$excelNote = ""
if ($Config_OutputExcel -ne "") {
    if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $excelCols = @("距離m","条件メモ","差分メモ","所在地","価格_円m2","地積_m2",
                       "用途地域","利用現況","調査年次",
                       "標準地番号","調査基準日","用途地域_元値",
                       "価格_取得列","用途地域_取得列","利用現況_取得列","座標取得方法","緯度","経度")
        if ($Config_OutputAllProperties) { $excelCols += "全属性JSON" }
        $sorted | Select-Object $excelCols |
            Export-Excel $Config_OutputExcel -AutoSize -FreezeTopRow -BoldTopRow `
                -WorksheetName "地価公示参考一覧" -OverwriteExistingFile
        $excelNote = "`n出力(Excel): $Config_OutputExcel"
    }
}

$elapsed = $sw.Elapsed.TotalSeconds
Write-Log ("処理完了: 全{0:N0}件中 {1}件抽出 | {2:F1}秒" -f $total, $sorted.Count, $elapsed) "OK"

# 処理ログCSV出力
$logSummary = [ordered]@{
    "処理日時"        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    "入力地番"        = $Config_Chiban
    "選択地番"        = $selectedChiban
    "対象地情報"      = if ($chibanInfo -ne "") { $chibanInfo } else { "（緯度経度指定）" }
    "中心緯度"        = $centerLat
    "中心経度"        = $centerLon
    "座標取得方法"    = $centerMethod
    "対象地積_m2"     = $targetArea
    "対象用途地域コード" = $Config_TargetZoningCode
    "都道府県"        = $Config_Prefecture
    "半径_m"          = $Config_RadiusMeters
    "地価データ件数"  = $total
    "抽出件数"        = $sorted.Count
    "警告件数"        = $Script:Warnings.Count
    "出力CSV"         = $Config_OutputCsv
    "処理時間_秒"     = [Math]::Round($elapsed, 2)
}
$logPath = Export-ProcessLog -BaseCsvPath $Config_OutputCsv -Summary $logSummary -Encoding $Config_OutputEncoding

# ── 結果表示
Write-Host "`n══════════════════════════════════════════════════════════" -ForegroundColor Green
if ($chibanInfo -ne "") { Write-Host "対象地: $chibanInfo" -ForegroundColor Green }
Write-Host ("完了: 全{0:N0}件中 {1}件抽出 | {2:F1}秒" -f $total, $sorted.Count, $elapsed) -ForegroundColor Green
Write-Host "出力(CSV/$encLabel): $Config_OutputCsv$excelNote" -ForegroundColor Green
Write-Host "処理ログ: $logPath" -ForegroundColor Green
if ($Script:Warnings.Count -gt 0) {
    Write-Host "警告: $($Script:Warnings.Count) 件（処理ログ参照）" -ForegroundColor Yellow
}
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green

# ── サマリー表示
if ($sorted.Count -gt 0) {
    Write-Host "`n【参考価格サマリー】半径$($Config_RadiusMeters)m  中心:($centerLat, $centerLon)" -ForegroundColor Yellow
    if ($hasZonCond -or $hasAreaCond) {
        $condParts = @()
        if ($hasZonCond)  { $condParts += "用途:$Config_TargetZoningCode（$(Decode-ZoningCode $Config_TargetZoningCode)）" }
        if ($hasAreaCond) { $condParts += "地積:${targetArea}㎡" }
        Write-Host ("  対象地条件: " + ($condParts -join "  ")) -ForegroundColor Yellow
    }

    $allP   = @($sorted | Where-Object { $null -ne $_.価格_円m2 } | ForEach-Object { [double]$_.価格_円m2 })
    $getP   = {
        param([string]$cat)
        @($sorted | Where-Object {
            $null -ne $_.価格_円m2 -and
            $ZONING_CATEGORY[(Get-ZoningNormCode $_.用途地域_元値)] -eq $cat
        } | ForEach-Object { [double]$_.価格_円m2 })
    }
    $matchP = @($sorted | Where-Object {
        $null -ne $_.価格_円m2 -and $hasZonCond -and
        (Get-ZoningNormCode $_.用途地域_元値) -eq (Get-ZoningNormCode $Config_TargetZoningCode)
    } | ForEach-Object { [double]$_.価格_円m2 })

    Show-PriceSummary "全体"      $allP
    Show-PriceSummary "住宅系"    (& $getP "住宅系")
    Show-PriceSummary "商業系"    (& $getP "商業系")
    Show-PriceSummary "工業系"    (& $getP "工業系")
    if ($hasZonCond) { Write-Host ""; Show-PriceSummary "用途一致のみ" $matchP }

    Write-Host "`n【上位10件（基本列）】" -ForegroundColor Yellow
    $sorted | Select-Object -First 10 距離m, 条件メモ, 差分メモ, 所在地, 価格_円m2, 地積_m2, 用途地域 |
        Format-Table -AutoSize
}

Write-Host @"

【本ツールについて】周辺参考情報の整理ツールです。
─────────────────────────────────────────────────────────────────
・出力は参考資料です。価格妥当性の判定には使用できません。
・接道/前面道路幅員/形状/高低差/建蔽率/容積率/商業繁華性等は
  判定できません。出力後に地図・現地・専門家で確認してください。
・最終判断は不動産鑑定士等の専門家にご確認ください。
─────────────────────────────────────────────────────────────────
"@ -ForegroundColor DarkGray
#endregion
