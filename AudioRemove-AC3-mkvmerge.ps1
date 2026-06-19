# ============================================================
# Script     : AudioRemove-AC3-mkvmerge.ps1 — (Version 1.0.7)
# Compatible : PS 7.6.1 | MKVToolNix 98.0+, uses only mkvmerge.exe
# Overview   : Pure remux via mkvmerge (no re-encode, byte-identical audio).
#
# Purpose    : Remove all AC3 and E-AC3 audio streams (typically low-bitrate).
# Result     : "Lean Master" containing only Video, high-fidelity audio
#              (TrueHD / AAC 7.1 / DTS-HD MA).
# Use        : Prepares source for DDP51/Keep71 to generate 1024k E-AC3 5.1 tracks.
#
# Retains    : Eng/untagged subs + attachments.  -KeepAllSubs keeps all subs.
#
# Drop/Keep  : (False: drops non-English subtitles / True: Keeps all subtitles)
#
# Usage (Windows)     : pwsh -ExecutionPolicy Bypass -File .\AudioRemove-AC3-mkvmerge.ps1 ".\YourMovie.mkv"
# Usage (macOS/Linux) : pwsh -File ./AudioRemove-AC3-mkvmerge.ps1 "./YourMovie.mkv"
#
# Output     : Creates "YourMovie_remux.mkv" in the same folder.
#              Will NOT overwrite existing files (auto-increments).
#
# Utility for: Conversion Engines (DDP51.ps1 / Keep71.ps1)
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
Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ==============================
# ENGINE: Banner + Stopwatch
# ==============================
$banner = @"

────────────────────────────────────────────────────────────
 AudioRemove-AC3-mkvmerge v1.0.7
 Pure remux: strips AC3/E-AC3 audio via mkvmerge
────────────────────────────────────────────────────────────
"@
Write-Host $banner -ForegroundColor Cyan
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ==============================
# ENGINE: mkvmerge binary setup
# ==============================
# Discovery order:
#   1. PATH (Get-Command)
#   2. $PSScriptRoot fallback
$ext = $IsWindows ? '.exe' : ''
$mkvmerge = $null

$cmd = Get-Command -Name "mkvmerge$ext" -ErrorAction SilentlyContinue
if ($cmd) {
    $mkvmerge = $cmd.Source
} else {
    $candidate = Join-Path $PSScriptRoot "mkvmerge$ext"
    if (Test-Path -LiteralPath $candidate) {
        $mkvmerge = $candidate
    }
}

if (-not $mkvmerge) {
    throw "mkvmerge not found in PATH or $PSScriptRoot. Install MKVToolNix 98.0+ (https://mkvtoolnix.download)."
}

if (-not $IsWindows -and -not ((Get-Item -LiteralPath $mkvmerge).UnixFileMode -band [System.IO.UnixFileMode]::UserExecute)) {
    throw "Binary not executable: $mkvmerge — run: chmod +x `"$mkvmerge`""
}

if (-not (Test-Path -LiteralPath $InputFile)) {
    throw "Input file not found: $InputFile"
}

# ----------------------
# Input / output paths
# ----------------------

$fullInput = [System.IO.Path]::GetFullPath($InputFile)
Write-Host "Input: $fullInput" -ForegroundColor DarkGray

$inputExt  = [System.IO.Path]::GetExtension($fullInput).ToLower()

if ($inputExt -ne '.mkv') {
    Write-Warning "Input is not an .mkv file (got '$inputExt'). mkvmerge will attempt to remux but compatibility is not guaranteed."
}

$outDir = [System.IO.Path]::GetDirectoryName($fullInput)
$base   = [System.IO.Path]::GetFileNameWithoutExtension($fullInput)

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

# ---------------------------------
# Strict-mode safe property helper
# ---------------------------------
# Under Set-StrictMode, accessing a missing property throws.
# This helper avoids that by:
#   1. Checking for a null object.
#   2. Checking for a blank property name.
#   3. Using PSObject.Properties[] instead of direct access.
# Use for any mkvmerge JSON field that may be missing.
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
# ENGINE: Codec label fmt
# -----------------------
# Truncates/pads codec strings to fixed-width column (26 chars).
# Used by summary engine for clean column alignment regardless of
# whether the codec name is "AC-3" (4 chars) or "TrueHD Atmos / AC-3" (long).
function Format-CodecLabel {
    param(
        [string]$Codec,
        [int]$Width = 26
    )
    if ([string]::IsNullOrWhiteSpace($Codec)) { return ('?').PadRight($Width) }
    if ($Codec.Length -gt $Width) {
        return $Codec.Substring(0, $Width - 1) + '…'
    }
    return $Codec.PadRight($Width)
}

# -----------------------
# ENGINE: Progress bar
# -----------------------
# Renders a single-line block progress bar driven by mkvmerge --gui-mode.
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

# =======================
# ENGINE: mkvmerge probe
# =======================
# Uses mkvmerge -J for JSON track identification.
# System.Diagnostics.Process is used (not & operator) so stdout and stderr
# can be captured independently — preserving JSON integrity even if mkvmerge
# emits warnings to stderr.
function Get-MkvJson {
    param(
        [string]$File,
        [string]$MkvMergePath
    )

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName               = $MkvMergePath
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
    $psi.ArgumentList.Add('-J')
    $psi.ArgumentList.Add($File)

    $proc = [System.Diagnostics.Process]::new()
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($proc.ExitCode -ne 0) {
        throw "mkvmerge -J failed (exit $($proc.ExitCode)):`n$stderr"
    }

    try {
        $parsed = $stdout | ConvertFrom-Json -Depth 64
    } catch {
        throw "Failed to parse mkvmerge JSON output: $($_.Exception.Message)"
    }

    if ($null -eq $parsed -or $null -eq (Get-PropSafe $parsed 'tracks')) {
        throw "mkvmerge returned no tracks. File may be corrupt or unsupported."
    }

    return $parsed
}

$probe = Get-MkvJson -File $fullInput -MkvMergePath $mkvmerge

# -------------------------------------------
# Track object model (video/audio/subs)
# -------------------------------------------
# mkvmerge JSON schema (key differences from ffprobe):
#   - track.id        (mkvmerge TID, used directly in --audio-tracks/--subtitle-tracks)
#   - track.type      ('video' | 'audio' | 'subtitles')   <- note 'subtitles' plural
#   - track.codec     (human-readable, e.g. "TrueHD Atmos")
#   - track.properties.codec_id    ("A_AC3", "A_EAC3", "A_TRUEHD", ...)
#   - track.properties.language    (ISO 639-2 code)
#   - track.properties.track_name  (track title)
#
# AC3 family is detected via codec_id (the canonical Matroska CodecID),
# not the human-readable codec field — codec_id is stable across mkvmerge
# versions and language packs.

# AC3 family definition (Matroska CodecIDs)
# Matched by prefix so legacy/variant IDs are also caught:
#   A_AC3, A_AC3/BSID9, A_AC3/BSID10  -> Dolby Digital
#   A_EAC3                            -> Dolby Digital Plus
$AC3Pattern = '^A_E?AC3'

$Tracks = @(
    (Get-PropSafe $probe 'tracks') | ForEach-Object {
        $props    = Get-PropSafe $_ 'properties'
        $rawLang  = Get-PropSafe $props 'language'
        $lang     = ($rawLang -is [string] -and $rawLang.Trim().Length -gt 0) ? $rawLang.Trim().ToLower() : $null
        $codecId  = Get-PropSafe $props 'codec_id'
        $title    = Get-PropSafe $props 'track_name'
        $type     = Get-PropSafe $_ 'type'
        $codec    = Get-PropSafe $_ 'codec'

        [pscustomobject]@{
            Id         = [int](Get-PropSafe $_ 'id')
            Type       = $type
            Codec      = $codec
            CodecId    = $codecId
            Language   = $lang
            Title      = $title
            IsAC3      = ($type -eq 'audio'     -and $codecId -match $AC3Pattern)
            IsAudio    = ($type -eq 'audio')
            IsVideo    = ($type -eq 'video')
            IsSubtitle = ($type -eq 'subtitles')
            IsEnglish  = ($type -eq 'subtitles' -and $lang -eq 'eng')
            IsUntagged = ($type -eq 'subtitles' -and ([string]::IsNullOrEmpty($lang) -or $lang -eq 'und'))
        }
    }
)

# -----------------------------------
# Audio / subtitle selection
# -----------------------------------

$VideoTracks = @($Tracks | Where-Object { $_.IsVideo })
$AudioTracks = @($Tracks | Where-Object { $_.IsAudio })
$keepAudio   = @($AudioTracks | Where-Object { -not $_.IsAC3 })
$dropAudio   = @($AudioTracks | Where-Object {       $_.IsAC3 })

if ($dropAudio.Count -eq 0) {
    $stopwatch.Stop()
    Write-Host "No AC3/E-AC3 audio streams found -- nothing to remove. Exiting." -ForegroundColor Yellow
    exit 0
}

if ($keepAudio.Count -eq 0) {
    $stopwatch.Stop()
    Write-Warning "All audio tracks are AC3/E-AC3 -- nothing to keep. Exiting."
    exit 2   # exit 2 = aborted (destructive action prevented); exit 0 = no AC3 found (clean no-op)
}

# Logging
Write-Host "Keeping audio  (mkvmerge TID): $($keepAudio.Id -join ', ')" -ForegroundColor DarkCyan
Write-Host "Dropping AC3   (mkvmerge TID): $($dropAudio.Id -join ', ')" -ForegroundColor DarkCyan

# Subtitles: English + untagged (default), or all (-KeepAllSubs)
$SubTracks    = @($Tracks | Where-Object { $_.IsSubtitle })
$engSubs      = @($SubTracks | Where-Object { $_.IsEnglish })
$untaggedSubs = @($SubTracks | Where-Object { $_.IsUntagged })

if ($untaggedSubs.Count -gt 0) {
    if ($SubtitleInfo -eq 'Enable') {
        Write-Warning "Subtitle stream(s) with no language tag or 'und' found (TIDs: $($untaggedSubs.Id -join ', ')) -- keeping as language is unknown."
    }
    $engSubs = @(($engSubs + $untaggedSubs) | Sort-Object Id)
}

if (-not $KeepAllSubs -and $engSubs.Count -eq 0) {
    if ($SubtitleInfo -eq 'Enable') {
        Write-Warning "No English subtitle streams found -- output will have no subtitles."
    }
}

$selectedSubs = $KeepAllSubs ? $SubTracks : $engSubs

# Attachments (mkvmerge carries them automatically; just count for the summary)
$attachments = Get-PropSafe $probe 'attachments'
$attachCount = if ($attachments) { @($attachments).Count } else { 0 }
if ($attachCount -gt 0) {
    Write-Host "Keeping $attachCount attachment(s) (carried by mkvmerge automatically)" -ForegroundColor DarkGray
}

if ($selectedSubs.Count -gt 0 -and $SubtitleInfo -eq 'Enable') {
    Write-Host "Keeping subtitle(s) (mkvmerge TID): $($selectedSubs.Id -join ', ')" -ForegroundColor DarkCyan
}

# -----------------------
# ENGINE: Build mkvmerge args
# -----------------------
# Notes:
#   --no-date        : omits the segment date field only. NOT full determinism:
#                      audio/video payload is copied verbatim, but mkvmerge still
#                      rewrites track-statistics tags. Use --deterministic <seed>
#                      if byte-identical reruns are actually required.
#   --audio-tracks   : explicit kept TIDs (per-file option; precedes input file)
#   --subtitle-tracks: explicit kept TIDs; falls back to -S when none selected
#   --track-order    : preserves source order minus removed AC3
#                      (default re-sorts by type: video → audio → subs → other)
#   Attachments      : carried by default; no flag needed
#   Chapters/metadata: carried automatically by mkvmerge

# Compose --track-order to preserve source order minus AC3 (file ID is always 0)
$orderedKept = @(
    @($VideoTracks) + @($keepAudio) + @($selectedSubs)
) | Sort-Object Id
$trackOrder  = ($orderedKept | ForEach-Object { "0:$($_.Id)" }) -join ','

$mkvArgs = [System.Collections.Generic.List[string]]::new()
$mkvArgs.Add('--gui-mode')
$mkvArgs.Add('--no-date')
$mkvArgs.Add('-o')
$mkvArgs.Add($OutputFile)

# Audio: explicit kept TIDs (must have at least one — guarded above)
$mkvArgs.Add('--audio-tracks')
$mkvArgs.Add(($keepAudio.Id -join ','))

# Subtitles: explicit kept TIDs, or -S if none
if ($selectedSubs.Count -gt 0) {
    $mkvArgs.Add('--subtitle-tracks')
    $mkvArgs.Add(($selectedSubs.Id -join ','))
} else {
    $mkvArgs.Add('-S')
}

# Track order preservation
if ($orderedKept.Count -gt 0) {
    $mkvArgs.Add('--track-order')
    $mkvArgs.Add($trackOrder)
}

$mkvArgs.Add($fullInput)

# ======================
# ENGINE: Execute mkvmerge
# ======================
# System.Diagnostics.Process is used (not & operator) because:
#   1. ArgumentList provides automatic argument escaping (no quoting bugs)
#   2. RedirectStandardOutput allows line by line reading for -gui mode parsing
#   3. RedirectStandardError captures errors without polluting stdout/JSON
$displayArgs = $mkvArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }
Write-Host "mkvmerge command: mkvmerge $($displayArgs -join ' ')" -ForegroundColor DarkGray

$psi = [System.Diagnostics.ProcessStartInfo]::new()
$psi.FileName               = $mkvmerge
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute        = $false
$psi.CreateNoWindow         = $true
$psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
$psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
foreach ($a in $mkvArgs) { $psi.ArgumentList.Add($a) }

$proc = [System.Diagnostics.Process]::new()
$proc.StartInfo = $psi

[void]$proc.Start()

# stderr is drained asynchronously via Task<string> to:
#   1. Avoid deadlock if stderr buffer fills while we read stdout
#   2. Avoid thread-safety issues with event-handler-based collection
#      (event fires on thread pool; List<string> is not thread-safe)
$stderrTask = $proc.StandardError.ReadToEndAsync()

Write-Host ""
Write-Host "Remuxing..." -ForegroundColor White

$inputBytes = (Get-Item -LiteralPath $fullInput).Length
$lastPct    = -1

while (-not $proc.StandardOutput.EndOfStream) {
    $line = $proc.StandardOutput.ReadLine()
    if ($null -eq $line) { continue }

    if ($line -match '^#GUI#progress\s+(\d+)%') {
        $pct = [int]$Matches[1]
        if ($pct -ne $lastPct) {
            Show-ProgressBar -Percent $pct
            $lastPct = $pct
        }
    }
}

$proc.WaitForExit()
$exitCode   = $proc.ExitCode
$stderrText = ($stderrTask.Result ?? '').Trim()
Write-Host ""  # newline after progress bar

$stopwatch.Stop()
$elapsed    = $stopwatch.Elapsed
$elapsedStr = "{0:D2}:{1:D2}:{2:D2}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

# mkvmerge exit codes: 0 = success, 1 = warnings, 2 = errors
if ($exitCode -ge 2) {
    if (Test-Path -LiteralPath $OutputFile) {
        # Remove partial output on error
        Remove-Item -LiteralPath $OutputFile -Force -ErrorAction SilentlyContinue
    }
    throw "mkvmerge exited with code $exitCode.`n$stderrText"
}

if ($exitCode -eq 1) {
    Write-Warning "mkvmerge completed with warnings (exit 1)."
    if ($stderrText) { Write-Host $stderrText -ForegroundColor DarkYellow }
}

$outputBytes = (Get-Item -LiteralPath $OutputFile).Length
$savedMiB    = ($inputBytes - $outputBytes) / 1MB
$absMiB      = [Math]::Abs($savedMiB)
$savedLabel  = $savedMiB -ge 0 ? "saved $("{0:N0}" -f $absMiB) MiB" : "overhead +$("{0:N0}" -f $absMiB) MiB"

Write-Host "`nDone. Output: $OutputFile" -ForegroundColor Green
Write-Host ("Size : {0:N0} MiB → {1:N0} MiB  ($savedLabel)" -f ($inputBytes/1MB), ($outputBytes/1MB)) -ForegroundColor DarkCyan
Write-Host ("Time : $elapsedStr") -ForegroundColor DarkCyan

# ======================
# ENGINE: Summary Engine
# ======================
# Column widths: TID(4) Codec(26) Action(10) Rule(rest)
Write-Host ""
$header = "{0,-4} {1} {2,-10} {3}" -f "TID", (Format-CodecLabel "Codec"), "Action", "Rule"
Write-Host $header -ForegroundColor White
Write-Host ("-" * 58) -ForegroundColor Yellow

foreach ($t in $VideoTracks) {
    $row = "{0,-4} {1} {2,-10} {3}" -f $t.Id, (Format-CodecLabel $t.Codec), 'Kept', 'Video passthrough'
    Write-Host $row -ForegroundColor White
}

foreach ($t in $AudioTracks) {
    $Action = $t.IsAC3 ? "Removed"     : "Kept"
    $Rule   = $t.IsAC3 ? "AC3 removed" : "Non-AC3 kept"
    $Color  = $t.IsAC3 ? "Red"         : "Green"

    $row = "{0,-4} {1} {2,-10} {3}" -f $t.Id, (Format-CodecLabel $t.Codec), $Action, $Rule
    Write-Host $row -ForegroundColor $Color
}

# ---- SUBTITLES (English + untagged by default; all if -KeepAllSubs) ----
if ($SubtitleInfo -eq 'Enable') {
    foreach ($t in $engSubs) {

        $Rule = $t.IsEnglish ? "English subtitle" : "Untagged subtitle"

        $row = "{0,-4} {1} {2,-10} {3}" -f $t.Id, (Format-CodecLabel $t.Codec), 'Kept', $Rule
        Write-Host $row -ForegroundColor Cyan
    }

    $nonEngCount = $SubTracks.Count - $engSubs.Count
    if ($nonEngCount -gt 0) {
        $Action = $KeepAllSubs ? 'Kept'    : 'Dropped'
        $Color  = $KeepAllSubs ? 'Cyan'    : 'Yellow'
        $row    = "{0,-4} {1} {2,-10} {3}" -f '—', (Format-CodecLabel 'subtitle(s)'), $Action, "$nonEngCount Non-English sub(s)"
        Write-Host $row -ForegroundColor $Color
    }
}

# ----- ATTACHMENTS -----
if ($attachCount -gt 0) {
    $row = "{0,-4} {1} {2,-10} {3}" -f '—', (Format-CodecLabel 'attachment(s)'), 'Kept', "$attachCount attachment(s) passthrough"
    Write-Host $row -ForegroundColor Magenta
}

Write-Host ("-" * 58) -ForegroundColor Yellow
Write-Host ""