<#
.SYNOPSIS
    performances.json을 읽어 지도용 데이터 파일(map\performances-data.js)을 만든다.

.DESCRIPTION
    키를 짧게 줄여 파일 크기를 낮춘다.
    좌표가 없는 공연도 포함하되 la/ln을 null로 둔다. 지도에는 안 올라가고 목록에만 나온다.
    줄거리(sty)는 길어서 넣지 않는다. 원본이 필요하면 performances.json을 본다.
    map\index.html이 <script src>로 읽으므로 file://로 열어도 CORS 문제가 없다.

.NOTES
    종료 코드: 0=성공, 1=실패
#>
[CmdletBinding()]
param(
    [string]$InFile,
    [string]$OutFile,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

Import-Module (Join-Path $PSScriptRoot 'lib\KopisCommon.psm1') -Force

if (-not $InFile)  { $InFile  = Join-Path $PSScriptRoot 'performances.json' }
if (-not $OutFile) { $OutFile = Join-Path $PSScriptRoot 'map\performances-data.js' }

function Write-Log { param([string]$m) if (-not $Quiet) { Write-Host $m } }

try {
    if (-not (Test-Path -LiteralPath $InFile)) {
        throw "입력 파일이 없다: $InFile  (먼저 update-performances.ps1을 실행할 것)"
    }

    $doc = Get-Content -LiteralPath $InFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $all = @($doc.items)
    Write-Log "입력: $InFile ($($all.Count)건, savedAt=$($doc._meta.savedAt))"

    $numeric = '^-?\d+(\.\d+)?$'
    $out = [System.Collections.Generic.List[object]]::new()
    $mapped = 0
    $noCoord = 0
    $outOfRange = 0

    # --- 용량 절감 (전부 무손실) -------------------------------------------
    # 1) 포스터 URL은 3,019건이 전부 같은 접두사를 쓴다. 접두사를 한 번만 두고 파일명만 담는다.
    # 2) 예매처 이름은 소수의 값이 반복된다(놀유니버스 1596회 등). 사전으로 빼고 색인만 담는다.
    # 3) vid/or/dh는 index.html이 한 번도 읽지 않는다. 빼도 화면이 달라지지 않는다.
    # 4) 빈 문자열 필드는 키까지 통째로 생략한다. JS는 undefined를 빈 값과 같게 다룬다.
    # 5) 공연장명·주소·좌표는 공연이 아니라 공연장의 속성이다. 공연 3,019건이 공연장 765곳을
    #    가리키므로 같은 문자열이 평균 4번씩 중복된다. 공연장 사전으로 빼고 색인만 담는다.
    $PO_PREFIX = 'https://www.kopis.or.kr/upload/pfmPoster/'
    $bkNames = [System.Collections.Generic.List[string]]::new()
    $bkIndex = @{}
    $venues  = [System.Collections.Generic.List[object]]::new()
    $venueIx = @{}

    foreach ($p in $all) {
        $la = $null
        $ln = $null
        if ($p.latitude -match $numeric -and $p.longitude -match $numeric) {
            $a = [double]$p.latitude
            $o = [double]$p.longitude
            # 한반도 범위를 벗어난 좌표는 잘못된 값으로 본다.
            if ($a -ge 33 -and $a -le 39 -and $o -ge 124 -and $o -le 132) {
                # 소수점 6자리면 약 0.1m 정밀도다. 그 이상은 용량만 먹는다.
                $la = [math]::Round($a, 6)
                $ln = [math]::Round($o, 6)
                $mapped++
            } else {
                $outOfRange++
            }
        } else {
            $noCoord++
        }

        # 포스터: 공통 접두사를 떼고 파일명만 남긴다.
        # 접두사가 다른 예외가 나오면 전체 URL을 그대로 두고, JS가 'http' 시작 여부로 구분한다.
        $po = [string]$p.poster
        if ($po.StartsWith($PO_PREFIX)) { $po = $po.Substring($PO_PREFIX.Length) }

        # 예매처: 이름을 사전으로 빼고 색인(i)과 URL(u)만 담는다.
        #
        # KOPIS 데이터에는 스킴이 없는 URL이 섞여 있다(6,074건 중 49건).
        # 'mangoticket.co.kr/...' 처럼 그대로 href에 넣으면 브라우저가 상대경로로 읽어
        # 'https://<사이트>/mangoticket.co.kr/...' 로 가서 404가 난다. 반드시 스킴을 붙인다.
        # https가 아니라 http를 붙이는 이유: 그 사이트가 TLS를 지원하면 대부분 리다이렉트하지만,
        # 지원하지 않는데 https를 붙이면 연결 자체가 실패한다.
        # (링크 이동은 혼합 콘텐츠 차단 대상이 아니라서 https 페이지에서도 http 링크는 동작한다.)
        $bk = [System.Collections.Generic.List[object]]::new()
        foreach ($b in @($p.booking)) {
            $url = ([string]$b.url).Trim()
            # '-' 같은 자리표시자나 도메인조차 없는 값은 링크가 될 수 없으니 버린다.
            if ([string]::IsNullOrWhiteSpace($url) -or $url -notmatch '\.') { continue }
            if ($url -notmatch '^https?://') { $url = 'http://' + $url }

            $nm = [string]$b.nm
            if (-not $bkIndex.ContainsKey($nm)) {
                $bkIndex[$nm] = $bkNames.Count
                $bkNames.Add($nm)
            }
            $bk.Add([ordered]@{ i = $bkIndex[$nm]; u = $url })
        }

        # 공연장: 이름+주소+좌표를 묶어 사전에 넣고 색인만 담는다.
        # 키는 mt10id를 쓰되, 없으면 이름+주소로 만든다(같은 공연장이 갈라지지 않게).
        $vkey = if ($p.mt10id) { [string]$p.mt10id } else { "$($p.fcltynm)|$($p.venueAdres)" }
        if (-not $venueIx.ContainsKey($vkey)) {
            $ve = [ordered]@{ n = [string]$p.fcltynm }
            if (-not [string]::IsNullOrEmpty([string]$p.venueAdres)) { $ve['a'] = [string]$p.venueAdres }
            if ($null -ne $la) { $ve['la'] = $la; $ve['ln'] = $ln }
            $venueIx[$vkey] = $venues.Count
            $venues.Add($ve)
        }

        # 빈 값은 키까지 생략한다. 필드가 여럿이라 키 이름만으로도 무게가 상당하다.
        $row = [ordered]@{}
        $row['id'] = [string]$p.mt20id
        $row['vi'] = $venueIx[$vkey]
        foreach ($pair in @(
            @('n',   $p.prfnm),        @('g',   $p.genrenm),
            @('s',   $p.prfpdfrom),    @('e',   $p.prfpdto),
            @('st',  $p.prfstate),     @('ar',  $p.area),
            @('c',   $p.prfcast),      @('cr',  $p.prfcrew),
            @('ent', $p.entrpsnm),     @('rt',  $p.prfruntime),
            @('ag',  $p.prfage),       @('pr',  $p.pcseguidance),
            @('dt',  $p.dtguidance),   @('po',  $po)
        )) {
            $val = [string]$pair[1]
            if (-not [string]::IsNullOrEmpty($val)) { $row[$pair[0]] = $val }
        }
        if ($bk.Count -gt 0) { $row['bk'] = $bk }

        $out.Add($row)
    }

    Write-Log "  지도 표시 대상 : $mapped 건"
    Write-Log "  좌표 없음      : $noCoord 건"
    Write-Log "  범위 벗어남    : $outOfRange 건"

    if ($mapped -eq 0) { throw "지도에 올릴 좌표가 하나도 없다. 생성을 중단한다." }

    $meta = [ordered]@{
        savedAt      = $doc._meta.savedAt
        builtAt      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        totalRecords = $all.Count
        mappedCount  = $mapped
        noCoordCount = $noCoord + $outOfRange
        rangeFrom    = $doc._meta.rangeFrom
        rangeTo      = $doc._meta.rangeTo
        # 서버(Actions)에서 돌면 로그를 열어 봐야 알 수 있는 값들이다.
        # 배포된 파일에 넣어 두면 공개 사이트에서 바로 확인할 수 있다.
        detailFail   = $doc._meta.detailFailCount
        venueCount   = $doc._meta.venueCount
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('// build-map.ps1이 생성한 파일이다. 직접 수정하지 말 것.')
    [void]$sb.AppendLine("// 생성 시각: $($meta.builtAt) / 원본 수집 시각: $($meta.savedAt)")
    # -AsArray: 항목이 1건일 때 배열이 아니라 객체로 나와 SHOWS.filter가 터지는 걸 막는다.
    # -Depth 6: SHOWS(1) -> 공연(2) -> bk 배열(3) -> 예매처 객체(4) -> 그 속성(5)까지 여유를 둔다.
    #           깊이가 모자라면 값이 "System.Collections.Hashtable" 문자열로 뭉개진다.
    [void]$sb.AppendLine("const SHOW_META = $($meta | ConvertTo-Json -Depth 3 -Compress);")
    [void]$sb.AppendLine("const PO_PREFIX = '$PO_PREFIX';")
    [void]$sb.AppendLine("const BK_NM = $($bkNames | ConvertTo-Json -Compress -AsArray);")
    [void]$sb.AppendLine("const VENUES = $($venues | ConvertTo-Json -Depth 3 -Compress -AsArray);")
    [void]$sb.AppendLine("const SHOWS = $($out | ConvertTo-Json -Depth 6 -Compress -AsArray);")

    Write-Utf8NoBom -Path $OutFile -Content $sb.ToString()

    $kb = [math]::Round((Get-Item -LiteralPath $OutFile).Length / 1KB)
    Write-Log ""
    Write-Log "생성 완료: $OutFile ($kb KB)"
    if ($kb -gt 2048) {
        # 경고로 흘리지 않고 실패로 끝낸다. loop가 "통과"로 오인하고 배포까지 가면 안 된다.
        throw "생성 파일이 2MB를 넘었다 ($kb KB). 기획서 9.5에 따라 필드 축소 판단이 필요하다."
    }
    exit 0
}
catch {
    Write-Error "지도 데이터 생성 실패: $($_.Exception.Message)"
    exit 1
}
