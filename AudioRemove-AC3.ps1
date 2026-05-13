# ============================================================
# Script     : AudioRemove-AC3.ps1 — (Version 3.2)
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
# Drops      : non-English subtitles.
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
    [string]$SubtitleInfo = 'Enable'  # Show subtitle stream info in console output
)

# Enable strict mode; halt on any error.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

# Resolve full absolute path to input file
$fullInput = [System.IO.Path]::GetFullPath($InputFile)

# Extract folder + base name
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
        '-show_entries', 'stream=index,codec_type,codec_name:stream_tags=language'
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
        $tagsObj = Get-PropSafe $_ 'tags'                        # Layer 1: safe tags access
        $rawLang = Get-PropSafe $tagsObj 'language'              # Layer 2: safe language access
        $lang    = ($rawLang -is [string] -and $rawLang.Trim().Length -gt 0) ? $rawLang.Trim().ToLower() : $null   # Layer 3: clean up

        [pscustomobject]@{
            Index      = $_.index
            Type       = $_.codec_type
            Codec      = $_.codec_name
            Language   = $lang
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

# Early exits
if ($dropAudio.Count -eq 0) {
    Write-Host "No AC3/E-AC3 audio streams found -- nothing to remove. Exiting."
    exit 0
}

if ($keepAudio.Count -eq 0) {
    Write-Warning "All audio tracks are AC3/E-AC3 -- nothing to keep. Exiting."
    exit 2   # exit 2 = aborted (destructive action prevented); exit 0 = no AC3 found (clean no-op)
}

# Logging
Write-Host "Keeping audio  (global stream index): $($keepAudio.Index -join ', ')" -ForegroundColor Green
Write-Host "Dropping AC3   (global stream index): $($dropAudio.Index -join ', ')" -ForegroundColor Red

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

if ($engSubs.Count -eq 0) {
    if ($SubtitleInfo -eq 'Enable') {
        Write-Warning "No English subtitle streams found -- output will have no subtitles."
    }
}

# Attachments
$AttachTracks = @($Tracks | Where-Object { $_.IsAttach })
if ($AttachTracks.Count -gt 0) {
    Write-Host "Keeping attachment stream(s): $($AttachTracks.Index -join ', ')"
}

# Data streams (warning only-dropped)
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
#   a:1+ >> none      (all other kept audio tracks get no track flags)
function New-AudioDispositionArgs {
    param([int]$Count)

    $dispArgs = [System.Collections.Generic.List[string]]::new()
    for ($j = 0; $j -lt $Count; $j++) {
        $dispArgs.Add("-disposition:a:$j")
        $dispArgs.Add(($j -eq 0) ? 'default' : 'none')
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
if ($engSubs.Count -eq 0) {
    $subMapArgs = @()
} else {
    $subMapArgs = New-MapArgs $engSubs
    if ($SubtitleInfo -eq 'Enable') {
        Write-Host "Keeping English subtitle stream(s): $($engSubs.Index -join ', ')" -ForegroundColor Green
    }
}

# Attachment map args
$attachMapArgs = @()
if ($AttachTracks.Count -gt 0) {
    $attachMapArgs = New-MapArgs $AttachTracks
}

#    -probesize        200M  = gives FFmpeg a larger initial scan of the file
#    -analyzeduration  200M  = about 200 seconds
#
#    -n : don't overwrite the output file if it already exists.
#
#    0:V? : maps all video streams except attached pictures.
#           Safe for audio-only MKVs. Use 0:v? if you also want
#           to keep cover art stored as an ATTACHED_PIC video stream.
$ffArgs = @(
    '-n'
    '-hide_banner'
    '-loglevel',         'warning'
    '-stats'
    '-probesize',        '200M'
    '-analyzeduration',  '200M'
    '-fflags',           '+discardcorrupt'   # discard corrupt packets instead of aborting
    '-i',                $fullInput
    '-avoid_negative_ts','make_zero'         # clamp negative PTS from source container
    '-map_metadata',     '0'                 # directly carry over container metadata
    '-map_chapters',     '0'                 # directly carry over chapters
    '-map',              '0:V?'
)

$ffArgs += $audioMapArgs
$ffArgs += $subMapArgs
$ffArgs += $attachMapArgs

# Give each audio track a single disposition setting.
# This avoids FFmpeg warnings about overlapping dispositions.
$ffArgs += New-AudioDispositionArgs $keepAudio.Count

$ffArgs += @(
    '-max_muxing_queue_size',  '14000'       # prevents video flooding mux queue on long files
    '-c',                      'copy'        # remux - no re-encode
    $OutputFile
)

# ======================
# ENGINE: Execute FFmpeg
# ======================
Write-Host "`nFFmpeg command (reference):"
$displayArgs = $ffArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }
Write-Host "ffmpeg $($displayArgs -join ' ')`n"

& $ffmpeg @ffArgs

if ($LASTEXITCODE -ne 0) {
    throw "FFmpeg exited with code $LASTEXITCODE."
}

Write-Host "`nDone. Output: $OutputFile" -ForegroundColor Green

# ======================
# ENGINE: Summary Engine
# ======================
Write-Host ""

$header = "{0,-4} {1,-20} {2,-10} {3}" -f "Idx", "Codec", "Action", "Rule"
Write-Host $header -ForegroundColor White
Write-Host ("-" * 52) -ForegroundColor Yellow
foreach ($t in $AudioTracks) {

    $Action = $t.IsAC3 ? "Removed"      : "Kept"
    $Rule   = $t.IsAC3 ? "AC3 removed"  : "Non-AC3 kept"
    $Color  = $t.IsAC3 ? "Red"          : "Green"

    $line = "{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, $Action, $Rule
    Write-Host $line -ForegroundColor $Color
}

# ---- SUBTITLES (English + untagged only) ----
if ($SubtitleInfo -eq 'Enable') {
    foreach ($t in $SubTracks) {

        if (-not $t.IsEnglish -and -not $t.IsUntagged) { continue }

        $Action = "Kept"
        $Rule   = $t.IsEnglish ? "English subtitle" : "Untagged subtitle"
        $Color  = "Cyan"

        $line = "{0,-4} {1,-20} {2,-10} {3}" -f $t.Index, $t.Codec, $Action, $Rule
        Write-Host $line -ForegroundColor $Color
    }

    $droppedSubCount = $SubTracks.Count - $engSubs.Count
    if ($droppedSubCount -gt 0) {
        $line   = "{0,-4} {1,-20} {2,-10} {3}" -f '—', 'subtitle(s)', 'Dropped', "$droppedSubCount Non-English sub(s)"
        Write-Host $line -ForegroundColor Yellow
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