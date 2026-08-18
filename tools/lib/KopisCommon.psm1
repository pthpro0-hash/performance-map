<#
.SYNOPSIS
    KOPIS 수집 스크립트들이 공유하는 함수 모음.
.NOTES
    네트워크를 타지 않는 순수 함수는 tests\ 에서 검증한다.
#>

function Get-KopisDateRange {
    <#
    .SYNOPSIS
        기간을 KOPIS의 31일 제한에 맞는 구간 목록으로 자른다.
    .OUTPUTS
        @{ Start='yyyyMMdd'; End='yyyyMMdd' } 배열
    #>
    param(
        [Parameter(Mandatory)][datetime]$From,
        [Parameter(Mandatory)][datetime]$To
    )
    $ranges = @()
    $cur = $From.Date
    $end = $To.Date
    while ($cur -le $end) {
        # 시작일 포함 31일이 되도록 30을 더한다. 31을 더하면 32일이 되어 API가 거부한다.
        $chunkEnd = $cur.AddDays(30)
        if ($chunkEnd -gt $end) { $chunkEnd = $end }
        $ranges += [pscustomobject]@{
            Start = $cur.ToString('yyyyMMdd')
            End   = $chunkEnd.ToString('yyyyMMdd')
        }
        $cur = $chunkEnd.AddDays(1)
    }
    return $ranges
}

function Get-KopisServiceKey {
    <#
    .SYNOPSIS
        KOPIS 인증키를 찾아서 돌려준다.
    .DESCRIPTION
        스크립트에 키를 직접 적어 두면 공개 저장소에 그대로 올라간다.
        아래 순서로 찾고, 없으면 오류로 멈춘다.

          1) -ServiceKey 인자
          2) 환경변수 KOPIS_API_KEY      (GitHub Actions는 이 경로를 쓴다)
          3) 프로젝트 루트의 .kopis-key  (로컬 전용. .gitignore로 빠져 있다)
    .PARAMETER Explicit
        스크립트가 -ServiceKey로 받은 값. 비어 있으면 다음 순서로 넘어간다.
    .PARAMETER Root
        .kopis-key를 찾을 폴더.
    #>
    param(
        [string]$Explicit,
        [string]$Root
    )

    if (-not [string]::IsNullOrWhiteSpace($Explicit)) { return $Explicit.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($env:KOPIS_API_KEY)) { return $env:KOPIS_API_KEY.Trim() }

    if ($Root) {
        $keyFile = Join-Path $Root '.kopis-key'
        if (Test-Path -LiteralPath $keyFile) {
            $k = (Get-Content -LiteralPath $keyFile -Raw -ErrorAction SilentlyContinue)
            if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
        }
    }

    throw @"
KOPIS 인증키를 찾을 수 없다. 아래 중 하나로 넣을 것.

  1) 이 실행에만          : -ServiceKey <키>
  2) 환경변수             : `$env:KOPIS_API_KEY = '<키>'
  3) 로컬 파일(권장)      : 프로젝트 루트에 .kopis-key 파일을 만들고 키만 한 줄 적기
                            (.gitignore에 있어 저장소로 올라가지 않는다)

  GitHub Actions에서는 저장소 Settings > Secrets and variables > Actions 에
  KOPIS_API_KEY 를 등록하면 워크플로가 환경변수로 넘겨준다.
"@
}

function ConvertTo-KopisIsoDate {
    <#
    .SYNOPSIS
        KOPIS의 '2016.05.12' 형식을 '2016-05-12'로 바꾼다.
        날짜로 보이지 않으면 원본을 그대로 돌려준다.
    #>
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return '' }
    $t = $Raw.Trim()
    if ($t -match '^(\d{4})[.\-/](\d{2})[.\-/](\d{2})$') {
        return "$($Matches[1])-$($Matches[2])-$($Matches[3])"
    }
    return $t
}

$script:KopisBase = 'http://www.kopis.or.kr/openApi/restful'

function Invoke-KopisApi {
    <#
    .SYNOPSIS
        KOPIS API를 호출하고 XmlDocument를 돌려준다.
    .DESCRIPTION
        - 응답 인코딩을 바이트에서 직접 UTF-8로 읽는다. Invoke-WebRequest의 자동 판별이 한글을 깨뜨리는 경우가 있다.
        - 실패 시 지수 백오프로 재시도한다.
        - 응답 루트가 기대한 노드가 아니면 오류로 본다(API가 오류를 XML로 돌려주기 때문).
    .PARAMETER Path
        'pblprfr' 또는 'pblprfr/PF132236' 같은 base 이후 경로
    .PARAMETER Query
        쿼리 파라미터 해시테이블. service는 자동으로 붙는다.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$Query = @{},
        [Parameter(Mandatory)][string]$ServiceKey,
        [int]$MaxRetry = 3,
        [int]$DelayMs = 200
    )

    $pairs = @("service=$ServiceKey")
    foreach ($k in $Query.Keys) {
        if ($null -ne $Query[$k] -and $Query[$k] -ne '') {
            $pairs += "$k=$([uri]::EscapeDataString([string]$Query[$k]))"
        }
    }
    $uri = "$script:KopisBase/$Path`?" + ($pairs -join '&')

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Start-Sleep -Milliseconds $DelayMs
            $resp  = Invoke-WebRequest -Uri $uri -Method Get -TimeoutSec 60
            $bytes = $resp.RawContentStream.ToArray()
            $text  = [System.Text.Encoding]::UTF8.GetString($bytes)
            if ([string]::IsNullOrWhiteSpace($text)) { throw "빈 응답" }
            $xml = [xml]$text
            if ($xml.DocumentElement.Name -notin 'dbs', 'boxofs', 'prfsts') {
                throw "예상 밖 응답 루트: $($xml.DocumentElement.Name) / 내용: $($text.Substring(0, [Math]::Min(200, $text.Length)))"
            }
            return $xml
        }
        catch {
            # 400/404는 다시 불러도 같은 결과다. 재시도하면 건당 7초씩 낭비되고
            # 상세 조회가 수천 건이라 전체 소요 시간이 크게 늘어난다.
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -in 400, 401, 403, 404) {
                throw "KOPIS 호출 실패 (HTTP $status, 재시도 안 함): $Path"
            }
            if ($attempt -ge $MaxRetry) {
                throw "KOPIS 호출 실패 ($attempt회 시도): $Path / $($_.Exception.Message)"
            }
            # 1초, 2초, 4초로 늘려가며 기다린다.
            Start-Sleep -Seconds ([math]::Pow(2, $attempt - 1))
        }
    }
}

function Write-Utf8NoBom {
    <#
    .SYNOPSIS
        BOM 없는 UTF-8로 파일을 쓴다. 임시 파일에 먼저 쓰고 교체하므로 중간에 실패해도 기존 파일이 남는다.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $tmp = "$Path.tmp"
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, $Content, $utf8NoBom)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

Export-ModuleMember -Function Get-KopisServiceKey, Get-KopisDateRange, ConvertTo-KopisIsoDate, Invoke-KopisApi, Write-Utf8NoBom
