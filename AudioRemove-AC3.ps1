# ============================================================
# Script     : AudioRemove-AC3.ps1 — (Version 3.7)
# Overview   : Pure remux (no re-encode).
#
# Purpose    : Remove all AC3 and E-AC3 audio streams (typically low-bitrate).
# Result     : "Lean Master" containing only Video, high-fidelity audio
#              (TrueHD / AAC 7.1 / DTS-HD MA).
# Use        : Prepares file for Conversion Engines (DDP51 / Keep71) to generate
#              high-bitrate (1024k) E-AC3 5.1 compatibility tracks.
#
# Retains    : English + untagged subtitles,
#              all attachment streams (fonts, cover art).
# Drop/Keep  : (False: drops non-English subtitles / True: Keeps all subtitles)
#
# Usage (Windows)     : pwsh -ExecutionPolicy Bypass -File .\AudioRemove-AC3.ps1 ".\YourMovie.mkv"
# Usage (macOS/Linux) : pwsh -File ./AudioRemove-AC3.ps1 "./YourMovie.mkv"
#
# Output     : Creates "YourMovie_remux.mkv" in the same folder.
#              Will NOT overwrite existing files.
#
# Utility for: Conversion Engines (DDP51.ps1 / Keep71.ps1)
# Compatible : PS 7.6.1 | FFmpeg 8.1
# ============================================================
#Requires -Version 7.6

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputFile,

    [ValidateSet('Enable','Disable')]
    [string]$SubtitleInfo = 'Enable', # Show subtitle stream info in console output

    [switch]$KeepAllSubs = $false     # ($False=Drops non-Eng subtitle, $True=Keeps all subtitles)
)

# Enable strict mode; halt on any error.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================
# ENGINE: Banner + Stopwatch
# ==============================
$banner = @"

────────────────────────────────────────────────────────────
 AudioRemove-AC3 v3.7
 Pure remux: strips AC3/E-AC3 audio via FFmpeg
────────────────────────────────────────────────────────────
"@
Write-Host $banner -ForegroundColor Cyan
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ==============================
# ENGINE: ffmpeg/ffprobe setup
# ==============================
$ext     = $IsWindows ? '.exe' : ''
$ffprobe = Join-Path $PSScriptRoot "ffprobe$ext"
$ffmpeg  = Join-Path $PSScriptRoot "ffmpeg$ext"

foreach ($bin in $ffprobe, $ffmpeg) {
    if (-not (Test-Path -LiteralPath $bin)) {
        throw "Missing required binary: $bin"
    }
    if (-not $IsWindows -and -not ((Get-Item -LiteralPath $bin).UnixFileMode -band [System.IO.UnixFileMode]::UserExecute)) {
        throw "Binary not executable: $bin — run: chmod +x `"$bin`""
    }
}

if (-not (Test-Path -LiteralPath $InputFile)) {
    throw "Input file not found: $InputFile"
}

# ----------------------
# Input / output paths
# ----------------------

$fullInput = [System.IO.Path]::GetFullPath($InputFile)
Write-Host "Input: $fullInput" -ForegroundColor DarkGray

$outDir  = [System.IO.Path]::GetDirectoryName($fullInput)
$base    = [System.IO.Path]::GetFileNameWithoutExtension($fullInput)

# ----------------------------
# Output auto-increment
# ----------------------------
$baseOut = $base + "_remux"
$OutputFile = Join-Path $outDir ($baseOut + ".mkv")

$counter = 1
while (Test-Path -LiteralPath $OutputFile) {
    $OutputFile = Join-Path $outDir ("{0}_{1}.mkv" -f $baseOut, $counter)
    $counter++
}

# =======================
# ENGINE: Probe function
# =======================
function Get-StreamInfo {
    param(
        [string]$File,
        [string]$ffprobePath
    )

    $probeArgs = @(
        '-v',            'error'
        '-probesize',       '200M'
        '-analyzeduration', '200M'
        '-show_entries', 'stream=index,codec_type,codec_name:stream_tags=language,title'
        '-of',           'json'
        $File
    )

    $probeJson = & $ffprobePath @probeArgs

    if ($LASTEXITCODE -ne 0) {
        throw "ffprobe failed. Aborting."
    }

    $streams = ($probeJson | ConvertFrom-Json).streams

    if (-not $streams) {
        throw "ffprobe returned no streams. File may be corrupt or empty."
    }

    return $streams
}

# =========================================
# ENGINE: Duration probe (for progress bar)
# =========================================
# Returns total container duration in seconds (double), or 0 if unavailable.
# Drives the duration-based percent in the progress bar. A 0 result triggers
# the indeterminate fallback (media-time readout instead of a percent bar).
function Get-DurationSeconds {
    param(
        [string]$File,
        [string]$ffprobePath
    )

    $durArgs = @(
        '-v',            'error'
        '-show_entries', 'format=duration'
        '-of',           'default=noprint_wrappers=1:nokey=1'
        $File
    )

    $raw = & $ffprobePath @durArgs

    if ($LASTEXITCODE -ne 0) { return [double]0 }

    $val = [double]0
    if ([double]::TryParse(
            ("$raw").Trim(),
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$val)) {
        return $val
    }
    return [double]0
}

$streams = Get-StreamInfo -File $fullInput -ffprobePath $ffprobe

# ---------------------------------
# Strict-mode safe property helper
# ---------------------------------
# Under Set-StrictMode, accessing a missing property throws.
# This helper avoids that by:
#   1. Checking for a null object.
#   2. Checking for a blank property name.
#   3. Using PSObject.Properties[] instead of direct access.
# Use for any ffprobe field that may be missing.
function Get-PropSafe {
    param(
        [object]$Obj,
        [string]$Prop
    )
    if ($null -eq $Obj)                 { return $null }
    if ([string]::IsNullOrEmpty($Prop)) { return $null }
    $p = $Obj.PSObject.Properties[$Prop]
    if ($null -eq $p)                   { return $null }
    return $p.Value
}

# -----------------------
# Single Line Progress bar
# -----------------------
function Show-ProgressBar {
    param(
        [int]$Percent,
        [int]$Width = 40
    )
    $Percent = [Math]::Max(0, [Math]::Min(100, $Percent))
    $filled  = [Math]::Floor($Width * $Percent / 100)
    $empty   = $Width - $filled
    $bar     = ('█' * $filled) + ('░' * $empty)

    if ($Percent -eq 100) {
        Write-Host -NoNewline ("`r  [{0}] {1,3}%" -f $bar, $Percent) -ForegroundColor Green
    }
    else {
        Write-Host -NoNewline ("`r  [{0}] {1,3}%" -f $bar, $Percent) -ForegroundColor Blue
    }
}

# -------------------------------------------
# Track object model (audio/subs/attachments)
# -------------------------------------------

# AC3 family definition
$AC3Codecs = @('ac3', 'eac3')

# Clean up all streams and turn them into track objects we can work with easily.
#
# Language handling has 3 layers of safety:
#   1. Get-PropSafe grabs the 'tags' object safely (no null errors).
#   2. We read the 'language' field safely, even if 'tags' isn't there.
#   3. Language values are cleaned up (trim + lowercase) for consistent comparison.
$Tracks = @(
    $streams | ForEach-Object {
        $tagsObj = Get-PropSafe $_ 'tags'
        $rawLang = Get-PropSafe $tagsObj 'language'
        $lang    = ($rawLang -is [string] -and $rawLang.Trim().Length -gt 0) ? $rawLang.Trim().ToLower() : $null

        [pscustomobject]@{
            Index      = $_.index
            Type       = $_.codec_type
            Codec      = $_.codec_name
            Language   = $lang
            Title      = (Get-PropSafe $tagsObj 'title')
            IsAC3      = ($_.codec_type -eq 'audio' -and $_.codec_name -in $AC3Codecs)
            IsAudio    = ($_.codec_type -eq 'audio')
            IsSubtitle = ($_.codec_type -eq 'subtitle')
            IsAttach   = ($_.codec_type -eq 'attachment')
            IsData     = ($_.codec_type -eq 'data')
            IsEnglish  = ($_.codec_type -eq 'subtitle' -and $lang -eq 'eng')
            IsUntagged = ($_.codec_type -eq 'subtitle' -and ([string]::IsNullOrEmpty($lang) -or $lang -eq 'und'))
        }
    }
)

# -----------------------------------
# Audio/subtitle/attachment selection
# -----------------------------------

# Audio tracks
$AudioTracks = @($Tracks | Where-Object { $_.IsAudio })
$keepAudio   = @($AudioTracks | Where-Object { -not $_.IsAC3 })
$dropAudio   = @($AudioTracks | Where-Object { $_.IsAC3 })

if ($dropAudio.Count -eq 0) {
    $stopwatch.Stop()
    Write-Host "No AC3/E-AC3 audio streams found -- nothing to remove. Exiting."
    exit 0
}

if ($keepAudio.Count -eq 0) {
    $stopwatch.Stop()
    Write-Warning "All audio tracks are AC3/E-AC3 -- nothing to keep. Exiting."
    exit 2   # exit 2 = aborted (destructive action prevented); exit 0 = no AC3 found (clean no-op)
}

# Logging
Write-Host "Keeping audio  (global stream index): $($keepAudio.Index -join ', ')" -ForegroundColor DarkCyan
Write-Host "Dropping AC3   (global stream index): $($dropAudio.Index -join ', ')" -ForegroundColor DarkCyan

# Subtitles: English + untagged
$SubTracks    = @($Tracks | Where-Object { $_.IsSubtitle })
$engSubs      = @($SubTracks | Where-Object { $_.IsEnglish })
$untaggedSubs = @($SubTracks | Where-Object { $_.IsUntagged })

if ($untaggedSubs.Count -gt 0) {
    if ($SubtitleInfo -eq 'Enable') {
        Write-Warning "Subtitle stream(s) with no language tag or 'und' found (indices: $($untaggedSubs.Index -join ', ')) -- keeping as language is unknown."
    }
    $engSubs = @(($engSubs + $untaggedSubs) | Sort-Object Index)
}

if (-not $KeepAllSubs -and $engSubs.Count -eq 0) {
    if ($SubtitleInfo -eq 'Enable') {
        Write-Warning "No English subtitle streams found -- output will have no subtitles."
    }
}

$selectedSubs = $KeepAllSubs ? $SubTracks : $engSubs

$AttachTracks = @($Tracks | Where-Object { $_.IsAttach })
if ($AttachTracks.Count -gt 0) {
    Write-Host "Keeping attachment stream(s): $($AttachTracks.Index -join ', ')"
}

$DataTracks = @($Tracks | Where-Object { $_.IsData })
if ($DataTracks.Count -gt 0) {
    Write-Warning "Data stream(s) detected at index/indices ($($DataTracks.Index -join ', ')) -- these will be DROPPED from the output."
}

# -----------------------
# Engine Helper Functions
# -----------------------

# ---------------------------
# ENGINE: Unified map builder
# ---------------------------
function New-MapArgs {
    param([array]$Tracks)

    $mapArgs = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $Tracks) {
        $mapArgs.Add('-map')
        $mapArgs.Add("0:$($t.Index)")
    }
    return $mapArgs.ToArray()
}

# -----------------------------------
# ENGINE: Audio-track settings helper
# -----------------------------------
# Replaces the two flags "-disposition:a none" and "-disposition:a:0 default".
# The two flags cause FFmpeg to warn "Multiple -disposition options specified for
# stream N" because both flags target a:0 at the same time.
#
# This helper produces one instruction per kept audio stream — no overlaps.
#   a:0  >> default   (first kept audio is the default playback track)
#   a:1+ >> 0         (clears disposition flags; '0' is the documented clear value)
function New-AudioDispositionArgs {
    param([int]$Count)

    $dispArgs = [System.Collections.Generic.List[string]]::new()
    for ($j = 0; $j -lt $Count; $j++) {
        $dispArgs.Add("-disposition:a:$j")
        $dispArgs.Add(($j -eq 0) ? 'default' : '0')
    }
    return $dispArgs.ToArray()
}

# ---------------------
# Main Execution Block
# ---------------------

# ===================================
# ENGINE: Build FFmpeg argument array
# ===================================

# Audio map args
$audioMapArgs = New-MapArgs $keepAudio

# Subtitle map args
if ($selectedSubs.Count -eq 0) {
    $subMapArgs = @()
} else {
    $subMapArgs = New-MapArgs $selectedSubs
    if ($SubtitleInfo -eq 'Enable') {
        Write-Host "Keeping subtitle stream(s): $($selectedSubs.Index -join ', ')" -ForegroundColor DarkCyan
    }
}

$attachMapArgs = @()
if ($AttachTracks.Count -gt 0) {
    $attachMapArgs = New-MapArgs $AttachTracks
}

#    -probesize        200M  = gives FFmpeg a larger initial scan of the file
#    -analyzeduration  200M  = about 200 seconds
#
#    -n : don't overwrite the output file if it already exists.
#
#    -progress pipe:1 : emit machine-readable progress blocks to stdout
#                       (parsed for out_time → percent). Pairs with -nostats
#                       so the human-oriented stats line is suppressed.
#
#    0:V? : maps all video streams except attached pictures.
#           Safe for audio-only MKVs. Use 0:v? if you also want
#           to keep cover art stored as an ATTACHED_PIC video stream.
$ffArgs = @(
    '-n'
    '-hide_banner'
    '-nostdin'                              # don't consume console stdin while we read stdout
    '-loglevel',         'error'
    '-nostats'                              # suppress default stats; we render our own bar
    '-progress',         'pipe:1'           # machine-readable progress → stdout
    '-stats_period',     '0.2'              # progress update cadence (default 0.5s) → smoother bar
    '-probesize',        '200M'
    '-analyzeduration',  '200M'
    '-fflags',           '+discardcorrupt'  # discard corrupt packets instead of aborting
    '-i',                $fullInput
    '-avoid_negative_ts','make_zero'        # clamp negative PTS from source container
    '-map_metadata',     '0'                # directly carry over container metadata
    '-map_chapters',     '0'                # directly carry over chapters
    '-map',              '0:V?'
)

$ffArgs += $audioMapArgs
$ffArgs += $subMapArgs
$ffArgs += $attachMapArgs

# Give each audio track a single disposition setting.
# This avoids FFmpeg warnings about overlapping dispositions.
$ffArgs += New-AudioDispositionArgs $keepAudio.Count

# Stamp MKV native TrackName element for each kept audio track that has a title.
# Ensures VLC and other players display the track name correctly regardless of
# whether the source stored it in the Tags block or the TrackName element.
$ffArgs += @(for ($j = 0; $j -lt $keepAudio.Count; $j++) {
    if ($keepAudio[$j].Title) { "-metadata:s:a:$j"; "title=$($keepAudio[$j].Title)" }
})

$ffArgs += @(
    '-max_muxing_queue_size',  '14000'      # prevents video flooding mux queue on long files
    '-c',                      'copy'       # remux - no re-encode
    $OutputFile
)

# ======================
# ENGINE: Execute FFmpeg
# ======================
# System.Diagnostics.Process is used (not the & operator) because:
#   1. RedirectStandardOutput lets us read -progress blocks line by line
#   2. RedirectStandardError captures real errors without polluting stdout
#   3. stderr is drained asynchronously to avoid buffer-fill deadlock
$displayArgs = $ffArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }
Write-Verbose "FFmpeg command (reference): ffmpeg $($displayArgs -join ' ')"

$durationSec = Get-DurationSeconds -File $fullInput -ffprobePath $ffprobe
$inputBytes  = (Get-Item -LiteralPath $fullInput).Length

$psi = [System.Diagnostics.ProcessStartInfo]::new()
$psi.FileName               = $ffmpeg
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute        = $false
$psi.CreateNoWindow         = $true
$psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
$psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
foreach ($a in $ffArgs) { $psi.ArgumentList.Add($a) }

$proc = [System.Diagnostics.Process]::new()
$proc.StartInfo = $psi
[void]$proc.Start()

# Drain stderr asynchronously (Task<string>) to prevent deadlock if the
# stderr buffer fills while we read the stdout progress stream.
$stderrTask = $proc.StandardError.ReadToEndAsync()

Write-Host ""
Write-Host "Remuxing..." -ForegroundColor White

$lastPct = -1
$hasDur  = $durationSec -gt 0

# FFmpeg -progress emits repeating key=value blocks. The fields we care about:
#   out_time_us=<microseconds>   (preferred; FFmpeg 8.1)
#   out_time_ms=<microseconds>   (legacy alias, also microseconds despite the name)
#   progress=continue | end
while (-not $proc.StandardOutput.EndOfStream) {
    $line = $proc.StandardOutput.ReadLine()
    if ($null -eq $line) { continue }

    if ($hasDur -and ($line -match '^out_time_us=(\d+)' -or $line -match '^out_time_ms=(\d+)')) {
        $us  = [double]$Matches[1]
        $pct = [int][Math]::Floor(($us / 1e6) / $durationSec * 100)
        if ($pct -ne $lastPct) {
            Show-ProgressBar -Percent $pct
            $lastPct = $pct
        }
    }
    elseif (-not $hasDur -and $line -match '^out_time=(\d{2}:\d{2}:\d{2})') {
        # Indeterminate fallback: no duration → show media-time readout instead of %.
        Write-Host -NoNewline ("`r  Processing... {0}" -f $Matches[1]) -ForegroundColor Blue
    }

    if ($line -match '^progress=end') {
        if ($hasDur) { Show-ProgressBar -Percent 100 }
    }
}

$proc.WaitForExit()
$exitCode   = $proc.ExitCode
$stderrText = ($stderrTask.Result ?? '').Trim()
Write-Host ""  # newline after progress bar

$stopwatch.Stop()
$elapsed    = $stopwatch.Elapsed
$elapsedStr = "{0:D2}:{1:D2}:{2:D2}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

if ($exitCode -ne 0) {
    if (Test-Path -LiteralPath $OutputFile) {
        # Remove partial output on error
        Remove-Item -LiteralPath $OutputFile -Force -ErrorAction SilentlyContinue
    }
    throw "FFmpeg exited with code $exitCode.`n$stderrText"
}

if ($stderrText) {
    Write-Warning "FFmpeg emitted messages on stderr:"
    Write-Host $stderrText -ForegroundColor DarkYellow
}

$outputBytes = (Get-Item -LiteralPath $OutputFile).Length
$savedMiB    = ($inputBytes - $outputBytes) / 1MB
Write-Host "`nDone. Output: $OutputFile" -ForegroundColor Green
$absMiB      = [Math]::Abs($savedMiB)
$savedLabel  = $savedMiB -ge 0 ? "saved $("{0:N0}" -f $absMiB) MiB" : "overhead +$("{0:N0}" -f $absMiB) MiB"
Write-Host ("Size : {0:N0} MiB → {1:N0} MiB  ($savedLabel)" -f ($inputBytes/1MB), ($outputBytes/1MB)) -ForegroundColor DarkCyan
Write-Host ("Time : $elapsedStr") -ForegroundColor DarkCyan

# ======================
# ENGINE: Summary Engine
# ======================
Write-Host ""

$header = "{0,-4} {1,-20} {2,-10} {3}" -f "Idx", "Codec", "Action", "Rule"
Write-Host $header -ForegroundColor White
Write-Host ("-" * 52) -ForegroundColor Yellow
$VideoTracks = @($Tracks | Where-Object { $_.Type -eq 'video' })
foreach ($t in $VideoTracks) {
    Write-Host ("{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, 'Kept', 'Video passthrough') -ForegroundColor White
}
foreach ($t in $AudioTracks) {

    $Action = $t.IsAC3 ? "Removed"      : "Kept"
    $Rule   = $t.IsAC3 ? "AC3 removed"  : "Non-AC3 kept"
    $Color  = $t.IsAC3 ? "Red"          : "Green"

    $line = "{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, $Action, $Rule
    Write-Host $line -ForegroundColor $Color
}

# ---- SUBTITLES (English + untagged only) ----
if ($SubtitleInfo -eq 'Enable') {
    foreach ($t in $engSubs) {

        $Action = "Kept"
        $Rule   = $t.IsEnglish ? "English subtitle" : "Untagged subtitle"
        $Color  = "Cyan"

        $line = "{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, $Action, $Rule
        Write-Host $line -ForegroundColor $Color
    }

    $nonEngCount = $SubTracks.Count - $engSubs.Count
    if ($nonEngCount -gt 0) {
        $Action = $KeepAllSubs ? 'Kept'   : 'Dropped'
        $Color  = $KeepAllSubs ? 'Cyan'   : 'Yellow'
        $line   = "{0,-4} {1,-20} {2,-10} {3}" -f '—', 'subtitle(s)', $Action, "$nonEngCount Non-English sub(s)"
        Write-Host $line -ForegroundColor $Color
    }
}

# ----- ATTACHMENTS -----
foreach ($t in $AttachTracks) {

    $Action = "Kept"
    $Rule   = "Attachment passthrough"
    $Color  = "Magenta"

    $line = "{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, $Action, $Rule
    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * 52) -ForegroundColor Yellow
Write-Host ""