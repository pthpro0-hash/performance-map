<#
.SYNOPSIS
    KOPIS에서 연극·뮤지컬 공연을 수집해 performances.json과 venues.json으로 저장한다.

.DESCRIPTION
    3단 수집이다.
      1) 공연목록 — 장르 2종 × 31일 이하 구간 × 페이지. mt20id를 모은다.
      2) 공연상세 — mt20id마다 1회. 공연시설 ID와 상세 정보를 얻는다. cache\details\ 에 캐시한다.
      3) 공연시설 상세 — mt10id마다 1회. 좌표를 얻는다. venues.json에 캐시한다.

    공연목록에는 공연시설 ID가 없어서 2단계를 거치지 않으면 좌표를 얻을 수 없다.
    첫 실행은 공연 건수만큼 상세 조회가 발생해 오래 걸린다. 두 번째부터는 캐시로 건너뛴다.

.PARAMETER MonthsBack
    수집 시작을 오늘 기준 몇 개월 전으로 할지.
    Phase 0 실측 결과 KOPIS는 "기간 겹침" 방식이므로 기본값 3으로 충분하다.
    (docs\specs\2026-08-14-API실측결과.md 참고)

.PARAMETER MonthsAhead
    수집 끝을 오늘 기준 몇 개월 후로 할지. 기본 6.

.PARAMETER Genre
    수집할 장르 코드. 기본은 연극(AAAA)과 뮤지컬(GGGA).

.PARAMETER RefreshDetails
    캐시를 무시하고 공연상세를 전부 다시 받는다.

.EXAMPLE
    .\update-performances.ps1
.EXAMPLE
    .\update-performances.ps1 -MonthsBack 0 -MonthsAhead 0 -Genre AAAA
    작은 범위 스모크 테스트. 스크립트를 편집하지 않고 동작을 확인할 수 있다.
.NOTES
    종료 코드: 0=성공, 1=실패
#>
[CmdletBinding()]
param(
    [string]$ServiceKey,
    [int]$MonthsBack = 3,
    [int]$MonthsAhead = 6,
    [string[]]$Genre = @('AAAA', 'GGGA'),
    [int]$DelayMs = 200,
    [int]$KeepBackups = 12,
    [string]$OutName = 'performances.json',
    [switch]$RefreshDetails,
    [switch]$NoBackup,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

Import-Module (Join-Path $PSScriptRoot 'lib\KopisCommon.psm1') -Force

# 인증키는 스크립트에 적지 않는다. -ServiceKey / 환경변수 / .kopis-key 순으로 찾는다.
$ServiceKey = Get-KopisServiceKey -Explicit $ServiceKey -Root $PSScriptRoot

$OutFile   = Join-Path $PSScriptRoot $OutName
$VenueFile = Join-Path $PSScriptRoot 'venues.json'
$CacheDir  = Join-Path $PSScriptRoot 'cache\details'

$GENRE_NAME = @{ 'AAAA' = '연극'; 'GGGA' = '뮤지컬' }

function Write-Log { param([string]$m) if (-not $Quiet) { Write-Host $m } }

# ---------------------------------------------------------------------------
# 1단계: 공연목록
# ---------------------------------------------------------------------------
function Get-PerformanceList {
    param([string]$Key, [datetime]$From, [datetime]$To, [string[]]$Genres)

    $byId = @{}
    $ranges = @(Get-KopisDateRange -From $From -To $To)
    Write-Log "1단계: 공연목록 조회 (구간 $($ranges.Count)개 x 장르 $($Genres.Count)종)"

    foreach ($r in $ranges) {
        foreach ($g in $Genres) {
            $page = 1
            while ($true) {
                $xml = Invoke-KopisApi -Path 'pblprfr' -ServiceKey $Key -DelayMs $DelayMs -Query @{
                    stdate = $r.Start; eddate = $r.End
                    cpage  = $page;    rows   = 100
                    shcate = $g
                }
                $batch = @($xml.dbs.db)
                foreach ($item in $batch) {
                    if (-not $byId.ContainsKey($item.mt20id)) {
                        $byId[$item.mt20id] = [ordered]@{
                            mt20id    = [string]$item.mt20id
                            prfnm     = [string]$item.prfnm
                            prfpdfrom = [string]$item.prfpdfrom
                            prfpdto   = [string]$item.prfpdto
                            fcltynm   = [string]$item.fcltynm
                            poster    = [string]$item.poster
                            area      = [string]$item.area
                            genrenm   = [string]$item.genrenm
                            openrun   = [string]$item.openrun
                            prfstate  = [string]$item.prfstate
                        }
                    }
                }
                # 100건 미만이면 마지막 페이지다.
                if ($batch.Count -lt 100) { break }
                $page++
                if ($page -gt 50) { throw "페이지가 50을 넘었다. 무한루프로 보고 중단한다. ($($r.Start)~$($r.End) / $g)" }
            }
            Write-Log ("  {0}~{1} {2} 누적 {3}건" -f $r.Start, $r.End, $GENRE_NAME[$g], $byId.Count)
        }
    }

    if ($byId.Count -eq 0) { throw "공연목록이 0건이다. 기존 파일을 덮어쓰지 않는다." }
    return $byId
}

# ---------------------------------------------------------------------------
# 2단계: 공연상세
# 호출량의 대부분이 여기서 발생한다. 한 건 받을 때마다 즉시 캐시에 써서,
# 중간에 끊겨도 다시 실행하면 이어서 진행되게 한다.
# ---------------------------------------------------------------------------
function Get-PerformanceDetail {
    param([string]$Key, [hashtable]$List, [switch]$Refresh)

    if (-not (Test-Path -LiteralPath $CacheDir)) {
        New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
    }

    $details = @{}
    $failed  = @()
    $fromCache = 0
    $fetched   = 0
    $total = $List.Count
    $i = 0

    Write-Log "2단계: 공연상세 조회 ($total 건)"

    foreach ($id in $List.Keys) {
        $i++
        $cacheFile = Join-Path $CacheDir "$id.json"

        if (-not $Refresh -and (Test-Path -LiteralPath $cacheFile)) {
            try {
                $details[$id] = Get-Content -LiteralPath $cacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
                $fromCache++
                continue
            }
            catch {
                # 캐시가 깨졌으면 지우고 다시 받는다.
                Remove-Item -LiteralPath $cacheFile -Force -ErrorAction SilentlyContinue
            }
        }

        try {
            $xml = Invoke-KopisApi -Path "pblprfr/$id" -ServiceKey $Key -DelayMs $DelayMs
            $db = $xml.dbs.db

            # 예매처: <relates><relate><relatenm>/<relateurl></relate></relates>
            # Phase 0 실측 3번에서 이 구조가 그대로임을 확인했다.
            $booking = @()
            if ($db.relates -and $db.relates.relate) {
                foreach ($rel in @($db.relates.relate)) {
                    if ($rel.relateurl) {
                        $booking += [ordered]@{ nm = [string]$rel.relatenm; url = [string]$rel.relateurl }
                    }
                }
            }

            $d = [ordered]@{
                mt20id       = [string]$db.mt20id
                mt10id       = [string]$db.mt10id
                prfcast      = [string]$db.prfcast
                prfcrew      = [string]$db.prfcrew
                prfruntime   = [string]$db.prfruntime
                prfage       = [string]$db.prfage
                entrpsnm     = [string]$db.entrpsnm
                pcseguidance = [string]$db.pcseguidance
                dtguidance   = [string]$db.dtguidance
                sty          = [string]$db.sty
                daehakro     = [string]$db.daehakro
                visit        = [string]$db.visit
                child        = [string]$db.child
                updatedate   = [string]$db.updatedate
                booking      = $booking
            }
            Write-Utf8NoBom -Path $cacheFile -Content ($d | ConvertTo-Json -Depth 5)
            $details[$id] = $d | ConvertTo-Json -Depth 5 | ConvertFrom-Json
            $fetched++
        }
        catch {
            # 한 건 실패로 전체를 멈추지 않는다.
            $failed += [ordered]@{ mt20id = $id; error = $_.Exception.Message }
        }

        if ($i % 100 -eq 0) {
            Write-Log ("  {0}/{1} (캐시 {2} / 신규 {3} / 실패 {4})" -f $i, $total, $fromCache, $fetched, $failed.Count)
        }
    }

    Write-Log ("  완료: 캐시 {0} / 신규 {1} / 실패 {2}" -f $fromCache, $fetched, $failed.Count)
    return [pscustomobject]@{ Details = $details; Failed = $failed }
}

# ---------------------------------------------------------------------------
# 3단계: 공연시설 상세 (좌표)
# 시설 수는 공연 수보다 훨씬 적고 잘 변하지 않으므로 venues.json에 모아 재사용한다.
# ---------------------------------------------------------------------------
function Get-Venue {
    param([string]$Key, [string[]]$VenueIds)

    $venues = @{}
    if (Test-Path -LiteralPath $VenueFile) {
        try {
            $existing = Get-Content -LiteralPath $VenueFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $existing.PSObject.Properties) { $venues[$p.Name] = $p.Value }
        }
        catch {
            # 깨진 캐시로 매 실행이 같은 자리에서 죽는 걸 막는다. 비우고 다시 받는다.
            Write-Log "  venues.json을 읽을 수 없다. 처음부터 다시 받는다. ($($_.Exception.Message))"
            $venues = @{}
        }
    }

    $need = @($VenueIds | Where-Object { $_ -and -not $venues.ContainsKey($_) })
    Write-Log "3단계: 공연시설 조회 (전체 $($VenueIds.Count)곳 중 신규 $($need.Count)곳)"

    $fetched = 0
    $done    = 0
    foreach ($vid in $need) {
        try {
            $xml = Invoke-KopisApi -Path "prfplc/$vid" -ServiceKey $Key -DelayMs $DelayMs
            $v = $xml.dbs.db
            $venues[$vid] = [ordered]@{
                mt10id     = [string]$v.mt10id
                fcltynm    = [string]$v.fcltynm
                adres      = [string]$v.adres
                la         = [string]$v.la
                lo         = [string]$v.lo
                telno      = [string]$v.telno
                relateurl  = [string]$v.relateurl
                seatscale  = [string]$v.seatscale
                sidonm     = [string]$v.sidonm
                gugunnm    = [string]$v.gugunnm
            }
            $fetched++
        }
        catch {
            # 시설 조회 실패는 좌표 없음으로 처리한다. 공연 자체는 목록에 남는다.
            $venues[$vid] = [ordered]@{ mt10id = $vid; fcltynm = ''; adres = ''; la = ''; lo = ''; error = $_.Exception.Message }
        }
        # 진행 표시는 $fetched가 아니라 처리 건수로 센다.
        # 실패가 이어지면 $fetched가 멈춰 같은 줄이 반복 출력된다.
        $done++
        if ($done % 50 -eq 0) { Write-Log "  $done / $($need.Count) (성공 $fetched)" }
    }

    Write-Utf8NoBom -Path $VenueFile -Content ($venues | ConvertTo-Json -Depth 5)

    $withCoord = @($venues.Values | Where-Object { $_.la -match '^-?\d+(\.\d+)?$' }).Count
    Write-Log "  공연장 $($venues.Count)곳 / 좌표 있음 $withCoord 곳"
    return $venues
}

# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------
try {
    $started = Get-Date
    $from = $started.AddMonths(-$MonthsBack)
    $to   = $started.AddMonths($MonthsAhead)
    Write-Log "공연 데이터 수집 시작 ($($started.ToString('yyyy-MM-dd HH:mm:ss')))"
    Write-Log "범위: $($from.ToString('yyyy-MM-dd')) ~ $($to.ToString('yyyy-MM-dd')) / 장르: $(($Genre | ForEach-Object { $GENRE_NAME[$_] }) -join ', ')"

    $list = Get-PerformanceList -Key $ServiceKey -From $from -To $to -Genres $Genre
    $det  = Get-PerformanceDetail -Key $ServiceKey -List $list -Refresh:$RefreshDetails
    $vids = @($det.Details.Values | ForEach-Object { $_.mt10id } | Where-Object { $_ } | Select-Object -Unique)
    $venues = Get-Venue -Key $ServiceKey -VenueIds $vids

    # 세 갈래를 하나로 합친다.
    $items = [System.Collections.Generic.List[object]]::new()
    $noCoord = 0
    foreach ($id in $list.Keys) {
        $base = $list[$id]
        $d = $det.Details[$id]
        $v = if ($d -and $d.mt10id -and $venues.ContainsKey($d.mt10id)) { $venues[$d.mt10id] } else { $null }

        $hasCoord = $v -and $v.la -match '^-?\d+(\.\d+)?$' -and $v.lo -match '^-?\d+(\.\d+)?$'
        if (-not $hasCoord) { $noCoord++ }

        $items.Add([ordered]@{
            mt20id       = $base.mt20id
            prfnm        = $base.prfnm
            genrenm      = $base.genrenm
            prfpdfrom    = ConvertTo-KopisIsoDate $base.prfpdfrom
            prfpdto      = ConvertTo-KopisIsoDate $base.prfpdto
            prfstate     = $base.prfstate
            fcltynm      = $base.fcltynm
            area         = $base.area
            openrun      = $base.openrun
            # 혼합 콘텐츠 차단을 피하려고 https로 바꾼다. Phase 0의 5번에서 200 OK를 확인했다.
            poster       = ($base.poster -replace '^http://', 'https://')
            mt10id       = if ($d) { $d.mt10id } else { '' }
            prfcast      = if ($d) { $d.prfcast } else { '' }
            prfcrew      = if ($d) { $d.prfcrew } else { '' }
            entrpsnm     = if ($d) { $d.entrpsnm } else { '' }
            prfruntime   = if ($d) { $d.prfruntime } else { '' }
            prfage       = if ($d) { $d.prfage } else { '' }
            pcseguidance = if ($d) { $d.pcseguidance } else { '' }
            dtguidance   = if ($d) { $d.dtguidance } else { '' }
            sty          = if ($d) { $d.sty } else { '' }
            daehakro     = if ($d) { $d.daehakro } else { '' }
            booking      = if ($d) { $d.booking } else { @() }
            venueAdres   = if ($v) { $v.adres } else { '' }
            venueTel     = if ($v) { $v.telno } else { '' }
            venueUrl     = if ($v) { $v.relateurl } else { '' }
            latitude     = if ($hasCoord) { $v.la } else { '' }
            longitude    = if ($hasCoord) { $v.lo } else { '' }
        })
    }

    $now = Get-Date
    $doc = [ordered]@{
        _meta = [ordered]@{
            savedAt         = $now.ToString('yyyy-MM-dd HH:mm:ss')
            savedAtISO      = $now.ToString('yyyy-MM-ddTHH:mm:sszzz')
            recordCount     = $items.Count
            rangeFrom       = $from.ToString('yyyyMMdd')
            rangeTo         = $to.ToString('yyyyMMdd')
            genres          = $Genre
            venueCount      = $venues.Count
            noCoordCount    = $noCoord
            detailFailCount = $det.Failed.Count
            source          = '공연예술통합전산망(KOPIS) 오픈API'
            endpoint        = 'http://www.kopis.or.kr/openApi/restful/pblprfr'
            generatedBy     = 'update-performances.ps1'
            elapsedMinutes  = [math]::Round(($now - $started).TotalMinutes, 1)
            notes           = @(
                'stdate/eddate는 최대 31일이라 기간을 구간으로 잘라 조회한다.'
                'KOPIS는 기간 겹침 방식이라 과거 3개월이면 진행 중인 장기 공연도 포함된다.'
                '공연목록에는 mt10id가 없다. 좌표를 얻으려면 공연상세를 거쳐 공연시설 상세까지 가야 한다.'
                'poster는 원본이 http다. 혼합 콘텐츠 차단을 피하려고 https로 바꿔 저장한다.'
                'prfstate는 코드값이 아니라 한글 문자열(공연중/공연예정/공연완료)이다.'
                'KOPIS는 예매 등록 기반이라 예매처를 쓰지 않는 소극장 공연은 누락될 수 있다.'
                'latitude/longitude가 빈 값이면 공연장에 좌표가 없다는 뜻이다. 지도에 올리지 않고 목록에만 노출한다.'
            )
        }
        items = $items
    }

    # 기존 파일 백업
    # venues.json도 함께 남긴다. 매 실행 통째로 덮어쓰는 파일이라 깨지면 시설 좌표를 전부 다시 받아야 한다.
    if ((Test-Path -LiteralPath $OutFile) -and -not $NoBackup) {
        $backupDir = Join-Path $PSScriptRoot 'backup'
        if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
        $stamp = (Get-Item -LiteralPath $OutFile).LastWriteTime.ToString('yyyyMMdd_HHmmss')
        Copy-Item -LiteralPath $OutFile -Destination (Join-Path $backupDir "performances_$stamp.json") -Force
        if (Test-Path -LiteralPath $VenueFile) {
            Copy-Item -LiteralPath $VenueFile -Destination (Join-Path $backupDir "venues_$stamp.json") -Force
        }
        foreach ($prefix in 'performances', 'venues') {
            $old = @(Get-ChildItem -LiteralPath $backupDir -Filter "$prefix`_*.json" |
                     Sort-Object LastWriteTime -Descending | Select-Object -Skip $KeepBackups)
            if ($old.Count) { $old | Remove-Item -Force -ErrorAction SilentlyContinue }
        }
    }

    Write-Utf8NoBom -Path $OutFile -Content ($doc | ConvertTo-Json -Depth 8)

    # 상세 조회 실패 내역을 남긴다. _meta에는 건수만 들어가서
    # 왜 실패했는지 확인할 방법이 없으면 다음 실행에서 판단할 근거가 사라진다.
    # (실패한 건은 캐시에 쓰이지 않으므로 다음 실행에서 자동으로 재시도된다.)
    if ($det.Failed.Count -gt 0) {
        $failLog = Join-Path $PSScriptRoot 'logs\detail-failures.json'
        Write-Utf8NoBom -Path $failLog -Content ($det.Failed | ConvertTo-Json -Depth 3 -AsArray)
        Write-Log "  실패 내역: $failLog"
    }

    $sizeMb = [math]::Round((Get-Item -LiteralPath $OutFile).Length / 1MB, 2)
    Write-Log ""
    Write-Log "저장 완료: $OutFile ($sizeMb MB, $($items.Count)건)"
    Write-Log "  좌표 없음 : $noCoord 건"
    Write-Log "  상세 실패 : $($det.Failed.Count) 건"
    Write-Log "  소요 시간 : $($doc._meta.elapsedMinutes) 분"
    exit 0
}
catch {
    Write-Error "수집 실패: $($_.Exception.Message)"
    exit 1
}
