# FKRAIN
# ius

[CmdletBinding()]
param(
    [string]$Cookie,
    [Nullable[int]]$UserId,
    [string]$CourseId,
    [string]$ClassroomId,
    [string]$LessonId,
    [string]$VideoId,
    [Nullable[double]]$Duration,
    [int]$Interval = 5,
    [int]$BatchSize = 120,
    [int]$Timeout = 20,
    [int]$Retries = 2,
    [int]$SleepMsBetweenBatch = 80,
    [switch]$DryRun,
    [switch]$AutoFixMs,
    [switch]$NoAutoFixMs,
    [switch]$AutoFill,
    [switch]$NoAutoFill,
    [switch]$Menu,
    [switch]$NoMenu
)

# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$global:YKT_BASE_URL = "https://www.yuketang.cn"

function Write-Info($Message)  { Write-Host $Message }
function Write-Warn($Message)  { Write-Host $Message }
function Write-ErrorLine($Message) { Write-Host $Message }

function ConvertTo-Hashtable {
    param([System.Collections.Specialized.NameValueCollection]$Collection)
    $hash = @{}
    foreach ($key in $Collection.Keys) {
        $hash[$key] = $Collection[$key]
    }
    return $hash
}

function Parse-Cookie {
    param([string]$CookieText)
    $result = @{}
    if ([string]::IsNullOrWhiteSpace($CookieText)) {
        return $result
    }
    foreach ($part in $CookieText.Split(';')) {
        $trimmed = $part.Trim()
        if ($trimmed -and $trimmed.Contains('=')) {
            $kv = $trimmed.Split('=', 2)
            $result[$kv[0].Trim()] = $kv[1].Trim()
        }
    }
    return $result
}

function New-YKTWebSession {
    param(
        [hashtable]$CookieMap,
        [string]$DefaultDomain = ".yuketang.cn"
    )
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    foreach ($entry in $CookieMap.GetEnumerator()) {
        if ([string]::IsNullOrWhiteSpace($entry.Key)) {
            continue
        }
        $cookieValue = if ($entry.Value) { $entry.Value } else { "" }
        $cookie = New-Object System.Net.Cookie($entry.Key, $cookieValue, "/", $DefaultDomain)
        $session.Cookies.Add($cookie)
        try {
            $preferredHost = try { ([Uri]$global:YKT_BASE_URL).Host } catch { "www.yuketang.cn" }
            $altCookie = New-Object System.Net.Cookie($entry.Key, $cookieValue, "/", $preferredHost)
            $session.Cookies.Add($altCookie)
        } catch {
            # ignore duplicate
        }
    }
    return $session
}

function Get-EffectiveCookie {
    param([string]$CookieParam)
    if (-not [string]::IsNullOrWhiteSpace($CookieParam)) {
        return $CookieParam.Trim()
    }
    $envCookie = $env:YKT_COOKIE
    if (-not [string]::IsNullOrWhiteSpace($envCookie)) {
        return $envCookie.Trim()
    }
    Write-Info "Please paste the Yuketang cookie (e.g. csrftoken=...; sessionid=...)"
    $inputCookie = Read-Host "Cookie"
    if ([string]::IsNullOrWhiteSpace($inputCookie)) {
        throw "Cookie is required."
    }
    return $inputCookie.Trim()
}

function New-YKTHeaders {
    param(
        [string]$Cookie,
        [string]$ClassroomId,
        [string]$UniversityId,
        [string]$Referer,
        [string]$CsrfToken
    )

    $headers = @{
        "User-Agent"       = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36"
        "Accept"           = "application/json, text/plain, */*"
        "origin"           = $global:YKT_BASE_URL
        "x-client"         = "web"
        "xt-agent"         = "web"
        "xtbz"             = "ykt"
        "accept-language"  = "zh-CN,zh;q=0.9"
        "accept-encoding"  = "gzip, deflate"
        "cache-control"    = "no-cache"
        "pragma"           = "no-cache"
        "dnt"              = "1"
        "sec-fetch-dest"   = "empty"
        "sec-fetch-mode"   = "cors"
        "sec-fetch-site"   = "same-origin"
        "sec-ch-ua"         = "`"Chromium`";v=`"122`", `"Not(A:Brand`";v=`"24`", `"Google Chrome`";v=`"122`""
        "sec-ch-ua-mobile"  = "?0"
        "sec-ch-ua-platform"= "`"Windows`""
        "x-requested-with"  = "XMLHttpRequest"
    }
    if ($Cookie) {
        $headers["Cookie"] = $Cookie
    }
    if ($ClassroomId) {
        $headers["classroom-id"] = "$ClassroomId"
        $headers["Classroom-Id"] = "$ClassroomId"
    }
    if ($UniversityId) {
        $headers["university-id"] = "$UniversityId"
        $headers["uv-id"]         = "$UniversityId"
    }
    if ($Referer) {
        $headers["referer"] = $Referer
    }
    if ($CsrfToken) {
        $headers["x-csrftoken"] = $CsrfToken
    }
    return $headers
}

function Invoke-YKTRequest {
    param(
        [string]$Method,
        [string]$Url,
        [hashtable]$Headers,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [int]$Timeout = 20,
        [string]$Body
    )
    $invokeParams = @{
        Uri        = $Url
        Method     = $Method
        Headers    = $Headers
        TimeoutSec = $Timeout
        ErrorAction = "Stop"
        WebSession = $WebSession
        UseBasicParsing = $true
    }
    if ($Body) {
        $invokeParams["Body"] = $Body
    }
    if ($Method -ne "GET") {
        $invokeParams["ContentType"] = "application/json;charset=UTF-8"
    }
    return Invoke-WebRequest @invokeParams
}

function Get-LessonSummary {
    param(
        [string]$LessonId,
        [string]$ClassroomId,
        [string]$Cookie,
        [string]$UniversityId,
        [string]$CsrfToken,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [int]$Timeout = 20
    )
    $referer = "$global:YKT_BASE_URL/v2/web/v3/playback/$LessonId/slide/2/0"
    $headers = New-YKTHeaders -Cookie $Cookie -ClassroomId $ClassroomId -UniversityId $UniversityId -Referer $referer -CsrfToken $CsrfToken
    $url = "$global:YKT_BASE_URL/api/v3/lesson-summary/replay?lesson_id=$LessonId"
    try {
        $rawResponse = Invoke-YKTRequest -Method "GET" -Url $url -Headers $headers -WebSession $WebSession -Timeout $Timeout
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }

    $statusCode = $rawResponse.StatusCode
    $content    = $rawResponse.Content
    try {
        $response = $content | ConvertFrom-Json
    }
    catch {
        return @{ success = $false; error = "Non-JSON response"; raw = $content; status = $statusCode }
    }

    if ($statusCode -ne 200) {
        return @{ success = $false; error = "HTTP $statusCode"; raw = $content; status = $statusCode }
    }

    if (-not $response -or $response.code -ne 0) {
        return @{ success = $false; error = ($response.msg | Out-String); raw = $content; status = $statusCode }
    }

    $data = $response.data
    $lesson = $data.lesson
    $lives = @()
    foreach ($item in ($data.live | Where-Object { $_ })) {
        $lives += [pscustomobject]@{
            id            = "$($item.id)"
            source        = $item.source
            url           = $item.url
            start         = $item.start
            end           = $item.end
            duration_sec  = if ($item.duration) { [math]::Round(([double]$item.duration) / 1000, 3) } else { 0.0 }
            order         = $item.order
        }
    }

    $durationRaw = $data.lessonDuration
    $durationSec = if ($durationRaw) { [math]::Round(([double]$durationRaw) / 1000, 3) } else { 0.0 }
    $best = Select-BestVideo -LiveList $lives -DefaultId $null -DefaultDuration $durationSec

    return @{
        success            = $true
        duration_sec       = $durationSec
        video_duration_sec = $best.duration
        video_id           = $best.id
        live_list          = $lives
        lesson             = $lesson
        course             = $lesson.course
        classroom          = $lesson.classroom
        user_id            = $data.userId
    }
}

function Get-CourseList {
    param(
        [string]$Cookie,
        [string]$UniversityId,
        [string]$CsrfToken,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [int]$Timeout = 20
    )
    $headers = New-YKTHeaders -Cookie $Cookie -ClassroomId "0" -UniversityId $UniversityId -Referer "$global:YKT_BASE_URL/v2/web/index" -CsrfToken $CsrfToken
    $url = "$global:YKT_BASE_URL/v2/api/web/courses/list?identity=2"
    try {
        $rawResponse = Invoke-YKTRequest -Method "GET" -Url $url -Headers $headers -WebSession $WebSession -Timeout $Timeout
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }

    $statusCode = $rawResponse.StatusCode
    $content = $rawResponse.Content
    try {
        $response = $content | ConvertFrom-Json
    }
    catch {
        return @{ success = $false; error = "Non-JSON response"; raw = $content; status = $statusCode }
    }

    if ($statusCode -ne 200) {
        return @{ success = $false; error = "HTTP $statusCode"; raw = $content; status = $statusCode }
    }

    if (-not $response -or $response.errcode -ne 0) {
        return @{ success = $false; error = ($response.errmsg | Out-String); raw = $content; status = $statusCode }
    }

    $classes = @()
    foreach ($item in ($response.data.list | Where-Object { $_ })) {
        $course = $item.course
        $classes += [pscustomobject]@{
            course_id    = "$($course.id)"
            course_name  = $course.name
            classroom_id = "$($item.classroom_id)"
            class_name   = $item.name
            term         = $item.term
        }
    }
    $userId = $null
    if ($response.data -and $response.data.user_id) {
        try { $userId = [int]$response.data.user_id } catch { $userId = $response.data.user_id }
    }
    return @{ success = $true; data = $classes; user_id = $userId }
}

function Get-LessonList {
    param(
        [string]$Cookie,
        [string]$ClassroomId,
        [string]$UniversityId,
        [string]$CsrfToken,
        [Nullable[int]]$UserId,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [int]$Timeout = 20
    )
    $referer = $null
    if ($ClassroomId -and $UniversityId) {
        $referer = "$global:YKT_BASE_URL/v2/web/studentLog/$ClassroomId?university_id=$UniversityId&platform_id=3&classroom_id=$ClassroomId&content_url="
    }
    $headers = New-YKTHeaders -Cookie $Cookie -ClassroomId $ClassroomId -UniversityId $UniversityId -Referer $referer -CsrfToken $CsrfToken
    $url = "{0}/v2/api/web/logs/learn/{1}?actype=-1&page=0&offset=20&sort=-1" -f $global:YKT_BASE_URL, $ClassroomId
    try {
        $rawResponse = Invoke-YKTRequest -Method "GET" -Url $url -Headers $headers -WebSession $WebSession -Timeout $Timeout
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }

    $statusCode = $rawResponse.StatusCode
    $content = $rawResponse.Content
    try {
        $response = $content | ConvertFrom-Json
    }
    catch {
        return @{ success = $false; error = "Non-JSON response"; raw = $content; status = $statusCode }
    }

    if ($statusCode -ne 200) {
        return @{ success = $false; error = "HTTP $statusCode"; raw = $content; status = $statusCode }
    }

    if (-not $response -or $response.errcode -ne 0) {
        return @{ success = $false; error = ($response.errmsg | Out-String); raw = $content; status = $statusCode }
    }

    $activities = $response.data.activities | Where-Object { $_.type -eq 14 }
    return @{ success = $true; data = $activities }
}

function Select-BestVideo {
    param(
        [System.Collections.IEnumerable]$LiveList,
        [string]$DefaultId,
        [Nullable[double]]$DefaultDuration
    )
    $best = $null
    $bestScore = [double]::MinValue
    foreach ($item in $LiveList) {
        $duration = 0.0
        if ($item.duration_sec -ne $null -and "$($item.duration_sec)" -ne "") {
            $duration = [double]$item.duration_sec
        }
        $order = if ($item.order -ne $null -and "$($item.order)" -ne "") { -1 * [double]$item.order } else { 0 }
        $score = $duration * 1000 + $order
        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $item
        }
    }

    if ($best) {
        $bestDuration = $DefaultDuration
        if ($best.duration_sec -ne $null -and "$($best.duration_sec)" -ne "") {
            $bestDuration = [double]$best.duration_sec
        }
        return @{
            id       = if ($best.id) { "$($best.id)" } else { $DefaultId }
            duration = $bestDuration
        }
    }
    return @{
        id       = $DefaultId
        duration = $DefaultDuration
    }
}

function Get-DisplaySlice {
    param(
        [string]$Text,
        [int]$MaxColumns
    )
    if (-not $Text) { return '' }
    $out = ''
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        $dw = if ($code -lt 128) { 1 } else { 2 }
        if ($width + $dw -gt $MaxColumns) { break }
        $out += $ch
        $width += $dw
    }
    return $out
}

if (-not $script:__progressRows) { $script:__progressRows=@{} }

function Show-AsciiProgress {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][int]$Current,
        [Parameter(Mandatory=$true)][int]$Total,
        [int]$Id = 0
    )
    if ($Total -le 0) { $Total = 1 }
    if ($Current -lt 0) { $Current = 0 }
    if ($Current -gt $Total) { $Current = $Total }

    # 分配并记住该进度条所在行
    if (-not $script:__progressRows.ContainsKey($Id)) {
        try {
            Write-Host ""  # 预留一行
            $script:__progressRows[$Id] = [Console]::CursorTop - 1
        } catch {
            $script:__progressRows[$Id] = 0
        }
    }
    $row = $script:__progressRows[$Id]

    # 计算进度条文本
    $percent = [Math]::Floor(($Current / [double]$Total) * 100)
    $win = $host.UI.RawUI.WindowSize
    $maxWidth = if ($win.Width -gt 20) { $win.Width } else { 120 }
    $barWidth = [Math]::Max(10, [Math]::Min(40, $maxWidth - 30))
    $filled = [Math]::Min($barWidth, [Math]::Floor($barWidth * $percent / 100.0))
    $bar = ('▓' * $filled) + ('░' * ($barWidth - $filled))
    $content = ("[{0}] {1,3}% {2} ({3}/{4})" -f $bar, $percent, $Title, $Current, $Total)
    $line = $content
    if ($line.Length -lt ($maxWidth - 1)) {
        $line = $line + (' ' * (($maxWidth - 1) - $line.Length))
    }

    # 将光标定位到该进度条行首并覆盖
    try { [Console]::SetCursorPosition(0, [Math]::Max(0,$row)) } catch {}
    try { Write-Host $line -NoNewline -ForegroundColor DarkCyan } catch { Write-Host $line -NoNewline }
}

function End-AsciiProgress {
    param([int]$Id = 0)
    try {
        if ($script:__progressRows.ContainsKey($Id)) {
            $row = $script:__progressRows[$Id]
            $win = $host.UI.RawUI.WindowSize
            $maxWidth = if ($win.Width -gt 20) { $win.Width } else { 120 }
            $clear = ' ' * ($maxWidth - 1)
            [Console]::SetCursorPosition(0, [Math]::Max(0,$row))
            Write-Host $clear -NoNewline
            Write-Host ""
            $script:__progressRows.Remove($Id) | Out-Null
        }
    } catch {}
}

function Invoke-ScrollableMenu {
    param(
        [Parameter(Mandatory=$true)][object[]]$Items,
        [Parameter(Mandatory=$true)][scriptblock]$Format,
        [int]$StartIndex = 0,
        [switch]$AllowBack,
        [switch]$AllowQuit,
        [string]$Title = 'Select'
    )

    if (-not $Items -or $Items.Count -eq 0) {
        return @{ Kind = 'Empty'; Value = -1 }
    }

    $idx = [math]::Max(0, [math]::Min($StartIndex, $Items.Count - 1))
    $offset = 0
    [System.Console]::CursorVisible = $false
    try {
        while ($true) {
            Clear-Host
            $w = $host.UI.RawUI.WindowSize.Width
            $h = $host.UI.RawUI.WindowSize.Height
            $visible = $h - 4
            if ($visible -lt 3) { $visible = 3 }

            if ($idx -lt $offset) { $offset = $idx }
            if ($idx -ge $offset + $visible) { $offset = $idx - $visible + 1 }
            if ($offset -lt 0) { $offset = 0 }

            Write-Host ($Title + ' (' + $Items.Count + ' items)') -ForegroundColor Cyan
            Write-Host '使用方向键/翻页键移动，回车确认，q 返回，Esc 退出' -ForegroundColor DarkGray

            $maxCols = $w - 4
            if ($maxCols -lt 20) { $maxCols = 20 }
            $end = [Math]::Min($Items.Count, $offset + $visible)
            for ($i=$offset; $i -lt $end; $i++) {
                $text = & $Format $Items[$i]
                if ($null -eq $text) { $text = '' } else { $text = [string]$text }
                $text = Get-DisplaySlice -Text $text -MaxColumns $maxCols
                $prefix = if ($i -eq $idx) { '> ' } else { '  ' }
                if ($i -eq $idx) {
                    Write-Host ($prefix + $text) -ForegroundColor Yellow
                } else {
                    Write-Host ($prefix + $text) -ForegroundColor DarkGray
                }
            }

            $key = [System.Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow' { if ($idx -gt 0) { $idx-- }; continue }
                'LeftArrow' { if ($idx -gt 0) { $idx-- }; continue }
                'DownArrow' { if ($idx -lt $Items.Count-1) { $idx++ }; continue }
                'RightArrow' { if ($idx -lt $Items.Count-1) { $idx++ }; continue }
                'PageUp' { $idx = [Math]::Max(0, $idx - $visible); continue }
                'PageDown' { $idx = [Math]::Min($Items.Count-1, $idx + $visible); continue }
                'Home' { $idx = 0; continue }
                'End' { $idx = $Items.Count-1; continue }
                'Enter' { return @{ Kind='Index'; Value=$idx } }
                'Escape' { if ($AllowQuit) { return @{ Kind='Quit' } } else { continue } }
                Default {
                    $ch = $key.KeyChar
                    if ($AllowBack -and ($ch -eq 'q' -or $ch -eq 'Q')) { return @{ Kind='Back' } }
                }
            }
        }
    }
    finally {
        [System.Console]::CursorVisible = $true
    }
}

function Invoke-MenuSelection {
    param(
        [Parameter(Mandatory=$true)][object[]]$Items,
        [Parameter(Mandatory=$true)][scriptblock]$Format,
        [int]$StartIndex = 0,
        [switch]$AllowBack,
        [switch]$AllowQuit,
        [string]$Title = 'Select'
    )

    if (-not $Items -or $Items.Count -eq 0) {
        return @{ Kind = 'Empty'; Value = -1 }
    }

    $idx = [math]::Max(0, [math]::Min($StartIndex, $Items.Count - 1))
    [System.Console]::CursorVisible = $false
    try {
        while ($true) {
            Clear-Host
            Write-Host ($Title + ' (' + $Items.Count + ' items)')
            Write-Host '使用方向键移动，回车确认，q 返回，Esc 退出'
            $maxLen = ($host.UI.RawUI.WindowSize.Width - 4)
            if ($maxLen -lt 20) { $maxLen = 20 }
            for ($i=0; $i -lt $Items.Count; $i++) {
                $text = & $Format $Items[$i]
                if ($null -eq $text) { $text = '' } else { $text = [string]$text }
                $text = Get-DisplaySlice -Text $text -MaxColumns $maxLen
                $prefix = if ($i -eq $idx) { '> ' } else { '  ' }
                if ($i -eq $idx) {
                    Write-Host ($prefix + $text) -ForegroundColor Yellow
                } else {
                    Write-Host ($prefix + $text) -ForegroundColor DarkGray
                }
            }

            $key = [System.Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow' { if ($idx -gt 0) { $idx-- } else { $idx = $Items.Count-1 }; continue }
                'LeftArrow' { if ($idx -gt 0) { $idx-- } else { $idx = $Items.Count-1 }; continue }
                'DownArrow' { if ($idx -lt $Items.Count-1) { $idx++ } else { $idx = 0 }; continue }
                'RightArrow' { if ($idx -lt $Items.Count-1) { $idx++ } else { $idx = 0 }; continue }
                'Enter' { return @{ Kind='Index'; Value=$idx } }
                'Escape' { if ($AllowQuit) { return @{ Kind='Quit' } } else { continue } }
                Default {
                    $ch = $key.KeyChar
                    if ($AllowBack -and ($ch -eq 'q' -or $ch -eq 'Q')) { return @{ Kind='Back' } }
                }
            }
        }
    }
    finally {
        [System.Console]::CursorVisible = $true
    }
}

function Invoke-IndexedSelection {
    param(
        [int]$Count,
        [string]$Prompt,
        [switch]$AllowBack,
        [switch]$AllowQuit
    )

    if ($Count -le 0) {
        return @{ Kind = 'Empty'; Value = -1 }
    }

    if ($Prompt) {
        Write-Host $Prompt
    }

    $index = 0
    $buffer = ""

    while ($true) {
        $instructions = "Use Left/Right to move, Enter to confirm"
        if ($AllowBack) { $instructions += ", press q to go back" }
        if ($AllowQuit) { $instructions += ", Esc to quit" }
        $display = [string]::Concat($instructions, ' | ', ($index + 1), '/', $Count)
        if ($buffer.Length -gt 0) {
            $display += " | typing: $buffer"
        }
        Write-Host ("`r{0}{1}" -f $display, (' ' * 60)) -NoNewline

        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'LeftArrow' {
                $index = ($index - 1 + $Count) % $Count
                $buffer = ""
                continue
            }
            'UpArrow' {
                $index = ($index - 1 + $Count) % $Count
                $buffer = ""
                continue
            }
            'RightArrow' {
                $index = ($index + 1) % $Count
                $buffer = ""
                continue
            }
            'DownArrow' {
                $index = ($index + 1) % $Count
                $buffer = ""
                continue
            }
            'Backspace' {
                if ($buffer.Length -gt 0) {
                    $buffer = $buffer.Substring(0, $buffer.Length - 1)
                }
                continue
            }
            'Enter' {
                Write-Host ""
                if ($buffer.Length -gt 0) {
                    try {
                        $parsed = [int]$buffer
                        if ($parsed -ge 1 -and $parsed -le $Count) {
                            return @{ Kind = 'Index'; Value = $parsed - 1 }
                        }
                    }
                    catch {}
                    $buffer = ""
                    continue
                }
                return @{ Kind = 'Index'; Value = $index }
            }
            'Escape' {
                Write-Host ""
                if ($AllowQuit) {
                    return @{ Kind = 'Quit' }
                }
                else {
                    continue
                }
            }
            default {
                $char = $key.KeyChar
                if ($AllowBack -and ($char -eq 'q' -or $char -eq 'Q')) {
                    Write-Host ""
                    return @{ Kind = 'Back' }
                }
                if ([char]::IsDigit($char)) {
                    $buffer += $char
                    try {
                        $parsed = [int]$buffer
                        if ($parsed -ge 1 -and $parsed -le $Count) {
                            $index = $parsed - 1
                        }
                    }
                    catch {}
                    continue
                }
            }
        }
    }
}

function Normalize-Duration {
    param(
        [Nullable[double]]$Duration,
        [bool]$AutoFix = $true
    )
    if (-not $Duration.HasValue) {
        throw "Video duration is required."
    }
    $value = [double]$Duration.Value
    if ($AutoFix -and $value -gt 86400) {
        $candidate = $value / 1000.0
        if ($candidate -le 43200) {
            Write-Info "Detected millisecond duration. Auto adjusted to $candidate seconds."
            return $candidate
        }
    }
    return $value
}

function New-HeartbeatEvents {
    param(
        [int]$UserId,
        [string]$CourseId,
        [string]$ClassroomId,
        [string]$LessonId,
        [string]$VideoId,
        [double]$Duration,
        [int]$Interval
    )
    $start = [long][DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    $events = New-Object System.Collections.Generic.List[object]
    $seq = 1
    $pg = "${VideoId}_${start}"

    $addEvent = {
        param([string]$Type, [double]$CP, [int]$Offset)
        $timestamp = $start + [long]$Offset
        $event = [ordered]@{
            ts          = "$timestamp"
            i           = 20
            et          = $Type
            p           = "web"
            t           = "ykt_playback"
            u           = $UserId
            c           = $CourseId
            classroomid = $ClassroomId
            lob         = "ykt"
            v           = $VideoId
            fp          = 0
            tp          = 0
            d           = 0
            pg          = $pg
            n           = "ali-cdn.xuetangx.com"
            lesson_id   = $LessonId
            source      = "ks"
            sp          = 1
            sq          = $seq
            cp          = [math]::Round($CP, 3)
        }
        $seq++
        $events.Add($event) | Out-Null
    }

    & $addEvent "loadstart" 0 0
    & $addEvent "loadeddata" 0 600
    & $addEvent "play" 0 900
    & $addEvent "playing" 0 920
    & $addEvent "waiting" 0 1100
    & $addEvent "playing" 0 1400

    $current = [double]$Interval
    $offset = 2000
    while ($current -le $Duration) {
        & $addEvent "heartbeat" ([math]::Min($current, $Duration)) $offset
        $current += $Interval
        $offset += $Interval * 1000
    }

    $last = $events[$events.Count - 1]
    if ($last.et -eq "heartbeat" -and $last.cp -lt $Duration) {
        & $addEvent "heartbeat" $Duration $offset
    }

    return $events
}

function Send-Heartbeats {
    param(
        [string]$Cookie,
        [string]$ClassroomId,
        [string]$UniversityId,
        [string]$CsrfToken,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [System.Collections.Generic.List[object]]$Heartbeats,
        [int]$BatchSize,
        [int]$Timeout,
        [int]$Retries,
        [int]$SleepMs,
        [string]$LessonId
    )
    $referer = "$global:YKT_BASE_URL/v2/web/v3/playback/$LessonId/slide/2/0"
    $headers = New-YKTHeaders -Cookie $Cookie -ClassroomId $ClassroomId -UniversityId $UniversityId -Referer $referer -CsrfToken $CsrfToken
    $url = "$global:YKT_BASE_URL/video-log/heartbeat/"

    if ($BatchSize -le 0) {
        $BatchSize = $Heartbeats.Count
    }

    $total = $Heartbeats.Count
    $sent = 0
    $lastStatus = $null
    $lastResponse = $null

    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $chunk = $Heartbeats[$i..([math]::Min($i + $BatchSize - 1, $total - 1))]
        $payload = @{ heart_data = $chunk } | ConvertTo-Json -Depth 6
        $attempt = 0
        $ok = $false
        $errorMessage = $null

        while ($attempt -le $Retries) {
            try {
                $response = Invoke-YKTRequest -Method "POST" -Url $url -Headers $headers -WebSession $WebSession -Timeout $Timeout -Body $payload
                $lastStatus = $response.StatusCode
                $lastResponse = $response.Content
                $done = [Math]::Min($i + $BatchSize, $total)
                Show-AsciiProgress -Title "发送心跳" -Current $done -Total $total -Id 2
                if ($response.StatusCode -eq 200) {
                    $ok = $true
                    break
                }
                else {
                    $errorMessage = $response.Content
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
            }
            $attempt++
            if ($attempt -le $Retries) {
                Write-Warn ("重试第 {0} 次: {1}" -f $attempt, $errorMessage)
                Start-Sleep -Milliseconds 300
            }
        }

        if (-not $ok) {
            return @{
                success     = $false
                count       = $sent
                status_code = $lastStatus
                response    = $errorMessage
            }
        }

        $sent += $chunk.Count
        if ($sent -lt $total -and $SleepMs -gt 0) {
            Start-Sleep -Milliseconds $SleepMs
        }
    }

    End-AsciiProgress -Id 2
    return @{
        success     = $true
        count       = $sent
        status_code = $lastStatus
        response    = $lastResponse
    }
}

function Run-InteractiveMenu {
    param(
        [hashtable]$Context
    )
    while ($true) {
        $courseRes = Get-CourseList -Cookie $Context.Cookie -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -WebSession $Context.Session -Timeout $Context.Timeout
        if (-not $courseRes.success) {
            Write-ErrorLine ("获取课程列表失败: {0}" -f $courseRes.error)
            if ($courseRes.status) {
                Write-ErrorLine "Status: $($courseRes.status)"
            }
            if ($courseRes.raw) {
                Write-Host "原始响应:" -ForegroundColor Yellow
                Write-Host $courseRes.raw
            }
            return
        }
        if (-not $Context.UserId -and $courseRes.user_id) {
            try {
                $Context["UserId"] = [int]$courseRes.user_id
            } catch {
                $Context["UserId"] = $courseRes.user_id
            }
        }
        $courses = $courseRes.data
        if (-not $courses -or $courses.Count -eq 0) {
            Write-Warn "未找到课程"
            return
        }

        Write-Host "`n=== 课程/班级列表 ==="
        $fmtCourse = {
            param($x)
            "[{0}] {1} - {2} (course_id={3})" -f $x.classroom_id, $x.course_name, $x.class_name, $x.course_id
        }
        $selResult = Invoke-ScrollableMenu -Items $courses -Format $fmtCourse -Title '课程/班级' -AllowQuit
        switch ($selResult.Kind) {
            'Quit' {
                Write-Info "已退出"
                return
            }
            'Back' {
                Write-Info "已退出"
                return
            }
            'Index' {
                $selected = $courses[$selResult.Value]
            }
            default {
                continue
            }
        }
        Write-Info "`n已选择: [$($selected.classroom_id)] $($selected.course_name) - $($selected.class_name)"
        Run-LessonMenu -Context $Context -CourseId $selected.course_id -ClassroomId $selected.classroom_id -CourseName $selected.course_name -ClassName $selected.class_name
    }
}

function Run-LessonMenu {
    param(
        [hashtable]$Context,
        [string]$CourseId,
        [string]$ClassroomId,
        [string]$CourseName,
        [string]$ClassName
    )
    while ($true) {
        $lessonRes = Get-LessonList -Cookie $Context.Cookie -ClassroomId $ClassroomId -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -UserId $Context.UserId -WebSession $Context.Session -Timeout $Context.Timeout
        if (-not $lessonRes.success) {
            Write-ErrorLine ("获取回放列表失败: {0}" -f $lessonRes.error)
            if ($lessonRes.raw) {
                Write-ErrorLine ("原始响应: {0}" -f $lessonRes.raw)
            }
            return
        }
        $lessons = $lessonRes.data
        Write-Host "`n当前: $CourseName - $ClassName (classroom_id=$ClassroomId)"
        if (-not $lessons -or $lessons.Count -eq 0) {
            Write-Warn "该班级暂无回放"
        }

        $actions = @(
            @{ label = "自动刷本班全部回放"; code = "auto" }
            @{ label = "手动选择单个回放"; code = "manual" }
            @{ label = "返回上一层"; code = "back" }
            @{ label = "退出程序"; code = "quit" }
        )
        Write-Host ""
        for ($a = 0; $a -lt $actions.Count; $a++) {
            Write-Host (" {0}. {1}" -f ($a + 1), $actions[$a].label)
        }
        $fmtAction = { param($x) $x.label }
        $actionSel = Invoke-MenuSelection -Items $actions -Format $fmtAction -Title '选择操作' -AllowBack -AllowQuit
        switch ($actionSel.Kind) {
            'Quit' {
                Write-Info "已退出"
                exit
            }
            'Back' {
                return
            }
            'Index' {
                $selectedAction = $actions[$actionSel.Value].code
                switch ($selectedAction) {
                    "auto" {
                        Clear-Host
                        # 取消“仅刷未观看”选项，默认刷全部
                        Invoke-BatchBrush -Context $Context -CourseId $CourseId -ClassroomId $ClassroomId -Lessons $lessons -OnlyUnviewed $false
                    }
                    "manual" {
                        if (-not $lessons -or $lessons.Count -eq 0) {
                            Write-Warn "当前无可选回放"
                            continue
                        }
                        $fmtLesson = {
                            param($x)
                            $t = if ($x.create_time) { (Get-Date -Date ([DateTimeOffset]::FromUnixTimeMilliseconds($x.create_time).LocalDateTime) -Format "yyyy-MM-dd HH:mm") } else { "-" }
                            "[{0}] {1} 时间:{2}" -f $x.courseware_id, $x.title, $t
                        }
                        $lessonSel = Invoke-ScrollableMenu -Items $lessons -Format $fmtLesson -Title '选择回放' -AllowBack -AllowQuit
                        switch ($lessonSel.Kind) {
                            'Quit' {
                                Write-Info "已退出"
                                exit
                            }
                            'Back' {
                                continue
                            }
                            'Index' {
                                Clear-Host
                                Invoke-SingleLesson -Context $Context -CourseId $CourseId -ClassroomId $ClassroomId -LessonItem $lessons[$lessonSel.Value]
                            }
                        }
                    }
                    "back" {
                        return
                    }
                    "quit" {
                        Write-Info "已退出"
                        exit
                    }
                }
            }
        }
    }
}

function Invoke-BatchBrush {
    param(
        [hashtable]$Context,
        [string]$CourseId,
        [string]$ClassroomId,
        [System.Collections.IEnumerable]$Lessons,
        [bool]$OnlyUnviewed
    )
    $filtered = @()
    foreach ($item in $Lessons) {
        if (-not $OnlyUnviewed -or -not $item.attend_status) {
            $filtered += $item
        }
    }
    if (-not $filtered -or $filtered.Count -eq 0) {
        Write-Warn "Nothing to do"
        return
    }
    $ordered = $filtered | Sort-Object -Property create_time
    $total = $ordered.Count
    Write-Host ("开始批量刷课: 共 {0} 个" -f $total) -ForegroundColor Cyan
    $ok = 0
    for ($i = 0; $i -lt $total; $i++) {
        $item = $ordered[$i]
        Show-AsciiProgress -Title "批量刷课" -Current $i -Total $total -Id 1
        # 将光标移到批量进度条下一行行首，避免覆盖
        try {
            if ($script:__progressRows.ContainsKey(1)) {
                [Console]::SetCursorPosition(0, [Math]::Max(0, $script:__progressRows[1] + 1))
            }
        } catch {}
        Write-Host ""
        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $total, $item.title) -ForegroundColor Yellow
        Write-Host ""  # 预留一行给心跳进度条
        if (Invoke-SingleLesson -Context $Context -CourseId $CourseId -ClassroomId $ClassroomId -LessonItem $item) {
            $ok++
        }
        # 每节完成后暂停 10 秒（最后一节可跳过）
        if ($i -lt ($total - 1)) {
            try { Start-Sleep -Seconds 10 } catch {}
        }
    }
    End-AsciiProgress -Id 1
    Write-Info ("`n批量刷课完成: 成功 {0}/{1}" -f $ok, $total)
}

function Invoke-SingleLesson {
    param(
        [hashtable]$Context,
        [string]$CourseId,
        [string]$ClassroomId,
        [pscustomobject]$LessonItem
    )
    $lessonId = "$($LessonItem.courseware_id)"
    $summary = Get-LessonSummary -LessonId $lessonId -ClassroomId $ClassroomId -Cookie $Context.Cookie -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -WebSession $Context.Session -Timeout $Context.Timeout
    if (-not $summary.success) {
        Write-ErrorLine ("获取回放信息失败: {0}" -f $summary.error)
        return $false
    }
    $durationCandidate = if ($summary.video_duration_sec) { [double]$summary.video_duration_sec } elseif ($summary.duration_sec) { [double]$summary.duration_sec } else { $null }
    $best = Select-BestVideo -LiveList $summary.live_list -DefaultId $summary.video_id -DefaultDuration $durationCandidate
    if (-not $best.id -or $best.duration -eq $null -or $best.duration -le 0) {
        Write-ErrorLine "Cannot resolve video id or duration"
        return $false
    }
    $userId = if ($summary.user_id) { [int]$summary.user_id } else { $Context.UserId }
    $course = if ($CourseId) { $CourseId } elseif ($summary.course.id) { "$($summary.course.id)" } else { $null }
    if (-not $userId -or -not $course) {
        Write-ErrorLine "Missing user_id or course_id"
        return $false
    }

    Write-Host ""
    Write-Host ("正在刷课: {0} (lesson_id={1}, video_id={2}, 时长={3}s)" -f $LessonItem.title, $lessonId, $best.id, $best.duration) -ForegroundColor Yellow
    Write-Host ""  # 预留进度条行
    $hb = New-HeartbeatEvents -UserId $userId -CourseId $course -ClassroomId $ClassroomId -LessonId $lessonId -VideoId $best.id -Duration $best.duration -Interval $Context.Interval
    if ($Context.DryRun) {
        Write-Info "Dry run: generated $($hb.Count) heartbeat events"
        $maxIndex = [math]::Min(4, $hb.Count - 1)
        for ($i = 0; $i -le $maxIndex; $i++) {
            $item = $hb[$i]
            Write-Host ("{0} et={1} cp={2} ts={3}" -f $item.sq, $item.et, $item.cp, $item.ts)
        }
        return $true
    }
    $res = Send-Heartbeats -Cookie $Context.Cookie -ClassroomId $ClassroomId -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -WebSession $Context.Session -Heartbeats $hb -BatchSize $Context.BatchSize -Timeout $Context.Timeout -Retries $Context.Retries -SleepMs $Context.SleepMs -LessonId $lessonId
    if ($res.success) {
        Write-Host "   Success" -ForegroundColor Green
        return $true
    }
    else {
        Write-ErrorLine "   Failed: $($res.response)"
        return $false
    }
}

function Build-Context {
    param(
        [string]$Cookie,
        [Nullable[int]]$UserId,
        [string]$CourseId,
        [string]$ClassroomId,
        [string]$LessonId,
        [string]$VideoId,
        [Nullable[double]]$Duration,
        [int]$Interval,
        [int]$BatchSize,
        [int]$Timeout,
        [int]$Retries,
        [int]$SleepMs,
        [bool]$DryRun,
        [bool]$AutoFix,
        [bool]$AutoFill,
        [bool]$Menu
    )
    $cookieMap = Parse-Cookie -CookieText $Cookie
    $csrf = $cookieMap["csrftoken"]
    $university = $cookieMap["uv_id"]
    if (-not $university) { $university = $cookieMap["university_id"] }

    $session = New-YKTWebSession -CookieMap $cookieMap
    $session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36'

    return @{
        Cookie      = $Cookie
        UserId      = if ($UserId) { [int]$UserId } else { $UserId }
        CourseId    = $CourseId
        ClassroomId = $ClassroomId
        LessonId    = $LessonId
        VideoId     = $VideoId
        Duration    = $Duration
        Interval    = $Interval
        BatchSize   = $BatchSize
        Timeout     = $Timeout
        Retries     = $Retries
        SleepMs     = $SleepMs
        DryRun      = $DryRun
        AutoFix     = $AutoFix
        AutoFill    = $AutoFill
        Menu        = $Menu
        CsrfToken   = $csrf
        UniversityId= $university
        Session     = $session
    }
}

function Invoke-DirectMode {
    param([hashtable]$Context)
    if (-not $Context.ClassroomId) { throw "ClassroomId is required" }
    if (-not $Context.LessonId) { throw "LessonId is required" }

    $summary = $null
    if ($Context.AutoFill -or -not $Context.VideoId -or -not $Context.Duration -or -not $Context.CourseId -or -not $Context.UserId) {
        Write-Info "Fetching playback summary from API..."
        $summary = Get-LessonSummary -LessonId $Context.LessonId -ClassroomId $Context.ClassroomId -Cookie $Context.Cookie -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -WebSession $Context.Session -Timeout $Context.Timeout
        if (-not $summary.success) {
            throw "Failed to fetch playback summary: $($summary.error)"
        }
        if (-not $Context.VideoId -and $summary.video_id) { $Context.VideoId = "$($summary.video_id)" }
        if (-not $Context.Duration -and ($summary.video_duration_sec -or $summary.duration_sec)) {
            if ($summary.video_duration_sec) {
                $Context.Duration = [double]$summary.video_duration_sec
            }
            elseif ($summary.duration_sec) {
                $Context.Duration = [double]$summary.duration_sec
            }
        }
        if (-not $Context.CourseId -and $summary.course.id) { $Context.CourseId = "$($summary.course.id)" }
        if (-not $Context.UserId -and $summary.user_id) { $Context.UserId = [int]$summary.user_id }
    }

    if (-not $Context.VideoId) { throw "Missing video_id and auto fill failed" }
    if (-not $Context.CourseId) { throw "Missing course_id and auto fill failed" }
    if (-not $Context.UserId) { throw "Missing user_id and auto fill failed" }
    if (-not $Context.Duration) { throw "Missing duration and auto fill failed" }

    $duration = Normalize-Duration -Duration $Context.Duration -AutoFix $Context.AutoFix
    $events = New-HeartbeatEvents -UserId $Context.UserId -CourseId $Context.CourseId -ClassroomId $Context.ClassroomId -LessonId $Context.LessonId -VideoId $Context.VideoId -Duration $duration -Interval $Context.Interval

    if ($Context.DryRun) {
        Write-Info "Dry run: generated $($events.Count) heartbeat events"
        $maxIndex = [math]::Min(4, $events.Count - 1)
        for ($i = 0; $i -le $maxIndex; $i++) {
            $item = $events[$i]
            Write-Host ("{0} et={1} cp={2} ts={3}" -f $item.sq, $item.et, $item.cp, $item.ts)
        }
        return
    }

    Write-Info "Submitting heartbeat payload"
    Write-Info ("  user: {0}  course: {1}  classroom: {2}" -f $Context.UserId, $Context.CourseId, $Context.ClassroomId)
    Write-Info ("  lesson_id: {0}  video_id: {1}" -f $Context.LessonId, $Context.VideoId)
    Write-Info ("  duration: {0}s  interval: {1}s  count: {2}  batch: {3}" -f $duration, $Context.Interval, $events.Count, $Context.BatchSize)

    $res = Send-Heartbeats -Cookie $Context.Cookie -ClassroomId $Context.ClassroomId -UniversityId $Context.UniversityId -CsrfToken $Context.CsrfToken -WebSession $Context.Session -Heartbeats $events -BatchSize $Context.BatchSize -Timeout $Context.Timeout -Retries $Context.Retries -SleepMs $Context.SleepMs -LessonId $Context.LessonId
    if ($res.success) {
        Write-Info ("`n完成 - 条数: {0}" -f $res.count)
    }
    else {
        Write-ErrorLine ("`n失败 - 已发送 {0} 条，状态码 {1}，响应: {2}" -f $res.count, $res.status_code, $res.response)
    }
} 

try {
    function Show-Home {
        Clear-Host
        $artText = @"
  _____   _  __  ____       _      ___   _   _ 
 |  ___| | |/ / |  _ \     / \    |_ _| | \ | |
 | |_    | ' /  | |_) |   / _ \    | |  |  \| |
 |  _|   | . \  |  _ <   / ___ \   | |  | |\  |
 |_|     |_|\_\ |_| \_\ /_/   \_\ |___| |_| \_|
                                               
  _                                            
 (_)  _   _   ___                              
 | | | | | | / __|                             
 | | | |_| | \__ \                             
 |_|  \__,_| |___/                                                
"@ -split "`r?`n"
        $palette = @('Cyan','Magenta','Yellow','Green','Blue','DarkCyan','DarkMagenta','DarkYellow')
        $frames = 10
        try {
            for ($f=0; $f -lt $frames; $f++) {
                Clear-Host
                for ($i=0; $i -lt $artText.Count; $i++) {
                    $col = $palette[($i + $f) % $palette.Count]
                    Write-Host $artText[$i] -ForegroundColor $col
                }
                Write-Host ''
                Write-Host "免责声明" -ForegroundColor Cyan
                Write-Host "本工具仅供学习与研究使用，请严格遵守法律法规与平台条款。" -ForegroundColor DarkGray
                Write-Host "开发者不对任何不当使用、封号、数据损失或其他后果承担责任。" -ForegroundColor DarkGray
                Write-Host "运行本程序即视为您已阅读并同意本免责声明。" -ForegroundColor DarkGray
                Write-Host ''
                Write-Host "按 Enter 继续，Esc 退出..." -ForegroundColor DarkGray
                try {
                    if ([Console]::KeyAvailable) {
                        $k=[Console]::ReadKey($true)
                        if ($k.Key -eq 'Enter') { break }
                        if ($k.Key -eq 'Escape') { exit 0 }
                    }
                } catch { }
                Start-Sleep -Milliseconds 120
            }
        } catch { }
        try {
            while ($true) {
                $k=[Console]::ReadKey($true)
                if ($k.Key -eq 'Enter') { break }
                if ($k.Key -eq 'Escape') { exit 0 }
            }
        } catch { }
    }

    function Show-Disclaimer {
        Clear-Host
        Write-Host "免责声明" -ForegroundColor Cyan
        Write-Host "本工具仅供学习与研究使用，请严格遵守法律法规与平台条款。" -ForegroundColor DarkGray
        Write-Host "开发者不对任何不当使用、封号、数据损失或其他后果承担责任。" -ForegroundColor DarkGray
        Write-Host "运行本程序即视为您已阅读并同意本免责声明。" -ForegroundColor DarkGray
        Write-Host ''
        Write-Host "按 Enter 继续，Esc 退出..." -ForegroundColor DarkGray
        try {
            while ($true) {
                $k=[Console]::ReadKey($true)
                if ($k.Key -eq 'Enter') { break }
                if ($k.Key -eq 'Escape') { exit 0 }
            }
        } catch { }
    }

    if (-not $Cookie) {
        Show-Home
        Clear-Host
        # 平台选择（方向键上下选择，回车确认）
        $platforms = @(
            @{ label = '雨课堂'; sub = 'www' }
            @{ label = '荷塘雨课堂'; sub = 'pro' }
            @{ label = '长江雨课堂'; sub = 'changjiang' }
            @{ label = '黄河雨课堂'; sub = 'huanghe' }
        )
        $fmtPlatform = { param($x) $x.label }
        $sel = Invoke-ScrollableMenu -Items $platforms -Format $fmtPlatform -Title '选择平台' -AllowQuit
        switch ($sel.Kind) {
            'Quit' { Write-Info '已退出'; exit 0 }
            'Index' { $selected = $platforms[$sel.Value] }
            default { $selected = $platforms[0] }
        }
        $global:YKT_BASE_URL = "https://$($selected.sub).yuketang.cn"

        Clear-Host
        Write-Host "请在浏览器打开获取 Cookie:" -ForegroundColor Cyan
        Write-Host ("  {0}/v2/web/index" -f $global:YKT_BASE_URL)
	Write-Host "  只需要填入csrftoken=xxxxxxxxxxxxxxxxxxxx; sessionid=xxxxxxxxxxxxxxxxx;"
        Write-Host ''
        $Cookie = Read-Host "请粘贴 Cookie"
    }
    $cookie = Get-EffectiveCookie -CookieParam $Cookie
    $autoFix = if ($NoAutoFixMs) { $false } elseif ($AutoFixMs) { $true } else { $true }
    $autoFill = if ($NoAutoFill) { $false } elseif ($AutoFill) { $true } else { $true }
    $menuMode = if ($NoMenu) { $false } elseif ($Menu) { $true } else { $true }

    $context = Build-Context -Cookie $cookie -UserId $UserId -CourseId $CourseId -ClassroomId $ClassroomId -LessonId $LessonId -VideoId $VideoId -Duration $Duration -Interval $Interval -BatchSize $BatchSize -Timeout $Timeout -Retries $Retries -SleepMs $SleepMsBetweenBatch -DryRun $DryRun.IsPresent -AutoFix $autoFix -AutoFill $autoFill -Menu $menuMode

    if ($context.Menu) {
        Run-InteractiveMenu -Context $context
    }
    else {
        Invoke-DirectMode -Context $context
    }
}
catch {
    $err = $_.Exception.Message
    Write-Host Error:
    Write-Host $err
    exit 1
}
