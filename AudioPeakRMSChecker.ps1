# ============================================================
# Script: AudioPeakRMSChecker.ps1  -  (Version 4.11)
#
# Overview: 
# Probes all audio streams in an MKV file using ffprobe, then runs
# FFmpeg astats analysis on supported tracks (TrueHD 7.1, AAC 7.1, DTS-HD 7.1, DDP 5.1).
# Reports: Peak level dB, RMS level dB, Crest factor and Peak count
# per track with PASS/FAIL evaluation.
# Saves a full astats log file per track for further inspection.
# Supports single-track and dual-track (source + downmix) analysis modes.
#
# Usage: pwsh -ExecutionPolicy Bypass -File .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv"
#
# Utility for: Conversion Engines (DDP51.ps1 / Keep71.ps1)
# Compatible: PS 7.6 | FFmpeg 8.1
# ============================================================

#Requires -Version 7.6
using namespace System.Globalization
using namespace System.Collections.Generic

param(
    [Parameter(Mandatory)]
    [string]$InputFile,

    # Tunable timeout parameters
    [int]$MaxRuntimeMultiplier = 4,
    [int]$MinTimeoutSeconds    = 60
)

# JSON Output Toggle
$JsonMode = $false   # Set to $true for JSON-only output

# --- Helpers ---

function Write-Info { param([string]$m) Write-Host $m -ForegroundColor Cyan }
function Write-Warn { param([string]$m) Write-Host $m -ForegroundColor Yellow }
function Write-Err  { param([string]$m) Write-Host $m -ForegroundColor Red }

function Get-Metric {
    param([string[]]$Lines, [string]$Pattern)
    $m = $Lines | Select-String -Pattern $Pattern -SimpleMatch | Select-Object -First 1
    if ($m -and $m.Line -match ':\s*(\S+)\s*$') { return $matches[1] }
    return 'N/A'
}

function Get-FileDuration {
    param([string]$FilePath)
    $raw = & ffprobe -v error -show_entries format=duration `
        -of default=noprint_wrappers=1:nokey=1 "$FilePath" 2>$null
    $first = $raw | Where-Object { $_ -match '^\d' } | Select-Object -First 1
    return $first ? [double]::Parse($first.Trim(), [CultureInfo]::InvariantCulture) : 0.0
}

# Codec processing ratios (processing time / file duration).
function Get-TrackETA {
    param([string]$Codec, [double]$FileDuration, [double]$SizeMB)

    $sizeScale = ($Codec -in @("truehd","aac","dts")) ? [Math]::Max(1.0, $SizeMB / 15360) : 1.0

    $ratio = switch ($Codec) {
        "truehd" { 0.02550 }
        "aac"    { 0.00322 }
        "eac3"   { 0.00365 }
        "dts"    { 0.02400 }
        default  { 0.00365 }
    }
    return [int]($FileDuration * $ratio * $sizeScale)
}

function Measure-Track {
    param(
        [string]$InputFile,
        [int]$MapIndex,
        [string]$Label,
        [string]$LogSuffix,
        [string]$Codec,
        [double]$TotalDuration,
        [double]$FileSizeMB,
        [int]$Channels
    )

    $base   = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $log    = "$base-$LogSuffix-astats.txt"
    $tmpErr = [System.IO.Path]::GetTempFileName()
    # This temporary file is used to catch FFmpeg’s output so it doesn’t fill up your screen.
    # FFmpeg writes to this file, but the script never needs to read it.
    # The file is removed automatically when the script finishes.
    $tmpOut = [System.IO.Path]::GetTempFileName()

    $etaSec = Get-TrackETA $Codec $TotalDuration $FileSizeMB
    $etaStr = "$([Math]::Floor($etaSec / 60))m $(([int]($etaSec % 60)).ToString('D2'))s"
    Write-Info "Analyzing $Label (map 0:a:$MapIndex)  -  estimated ~$etaStr ..."

    $ffArgs = @(
        '-hide_banner'
        '-vn'
        '-i', $InputFile
        '-map', "0:a:$MapIndex"
        '-af', 'astats=measure_overall=all:measure_perchannel=none'
        '-f', 'null'
        '-'
    )

    $procParams = @{
        FilePath               = "ffmpeg"
        ArgumentList           = $ffArgs
        RedirectStandardError  = $tmpErr
        RedirectStandardOutput = $tmpOut
        NoNewWindow            = $true
        PassThru               = $true
    }
    $proc = Start-Process @procParams

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # This timeout check ensures the script behaves predictably.
    # If the file duration is unknown, the estimated time becomes zero and the timeout shrinks to the minimum value.
    # That minimum (60 seconds) is far too short for long files and would cause false timeout errors.
    # To avoid this, the script uses a safe 1‑hour limit whenever the duration cannot be determined.
    $maxSec = ($TotalDuration -gt 0) ?
        [Math]::Max($MinTimeoutSeconds, $etaSec * $MaxRuntimeMultiplier) :
        [Math]::Max($MinTimeoutSeconds, 3600)

    try {
        while (-not $proc.HasExited) {

            if ($sw.Elapsed.TotalSeconds -gt $maxSec) {
                Write-Err "FFmpeg exceeded maximum allowed runtime for $Label (map 0:a:$MapIndex)."
                Write-Err "Elapsed: $([int]$sw.Elapsed.TotalSeconds)s  |  Limit: $maxSec s"
                try { $proc.Kill(); $proc.WaitForExit(2000) | Out-Null } catch {}
                exit 5
            }

            $remaining = [Math]::Max(0, $etaSec - [int]$sw.Elapsed.TotalSeconds)
            $msg = ("  ETA: $([Math]::Floor($remaining / 60))m $(([int]($remaining % 60)).ToString('D2'))s remaining...").PadRight(35)
            [Console]::Write("`r" + $msg)

            Start-Sleep -Seconds 1
        }

        [Console]::Write("`r" + (" " * 40) + "`r")
        $e = $sw.Elapsed; Write-Info "  Done in $([Math]::Floor($e.TotalMinutes))m $($e.Seconds.ToString('D2'))s"
        $out = Get-Content $tmpErr -Encoding UTF8 -ErrorAction SilentlyContinue

        # ===========================
        # Guardrail: Validate stderr
        # ===========================

        if (-not $out -or $out.Count -eq 0) {
            Write-Err "FFmpeg produced no stderr output for $Label (map 0:a:$MapIndex). Analysis aborted."
            exit 2
        }

        # 2. Deterministic exit-code check
        $exit = ($proc.ExitCode -is [int]) ? $proc.ExitCode : 0
        if ($exit -ne 0) {
            Write-Err "FFmpeg exited with code $exit while analyzing $Label (map 0:a:$MapIndex)."
            Write-Err "Check the log file: $log"
            exit 3
        }

        # 3. Missing astats metrics
        $peakLine  = $out | Select-String -Pattern 'Peak level dB' -SimpleMatch
        $rmsLine   = $out | Select-String -Pattern 'RMS level dB' -SimpleMatch
        $clipLine  = $out | Select-String -Pattern '] Peak count'  -SimpleMatch

        if (-not $peakLine -or -not $rmsLine -or -not $clipLine) {
            Write-Err "FFmpeg produced incomplete astats output for $Label (map 0:a:$MapIndex)."
            Write-Err "Required metrics missing: Peak/RMS/Peak count."
            exit 6
        }
    }
    finally {
        Remove-Item $tmpErr, $tmpOut -Force -ErrorAction SilentlyContinue
    }

    $out | Out-File -LiteralPath $log -Encoding UTF8BOM
    Write-Info "  Saved astats log: $log"

    # Log verification
    if (-not (Test-Path -LiteralPath $log)) {
        Write-Err "Astats log file was not created for $Label (map 0:a:$MapIndex)."
        exit 10
    }

    $logInfo = Get-Item -LiteralPath $log -ErrorAction SilentlyContinue
    if (-not $logInfo -or $logInfo.Length -lt 10) {
        Write-Err "Astats log file is empty or incomplete for $Label (map 0:a:$MapIndex)."
        exit 11
    }

    $peak  = Get-Metric $out "Peak level dB"
    $rms   = Get-Metric $out "RMS level dB"
    $clip  = Get-Metric $out "] Peak count"

    # Guard: FFmpeg reports silent audio as “-inf,” which TryParse cannot read as a valid number.
    if ($peak -match '^-?inf') { $peak = '-100.0' }
    if ($rms  -match '^-?inf') { $rms  = '-100.0' }

    # Numeric validation
    $validPeak = 0.0
    $validRms  = 0.0
    $validClipDouble = 0.0

    if (-not [double]::TryParse($peak, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validPeak)) {
        Write-Err "Peak level dB value '$peak' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 7
    }
    if (-not [double]::TryParse($rms,  [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validRms)) {
        Write-Err "RMS level dB value '$rms' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 8
    }
    if (-not [double]::TryParse($clip, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validClipDouble)) {
        Write-Err "Peak count value '$clip' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 9
    }

    # [Math]::Round uses MidpointRounding.ToEven (banker's rounding) by default.
    # For example, 2.5 becomes 2 and 3.5 becomes 4, which keeps the results consistent and predictable.
    $validClip = [int][Math]::Round($validClipDouble, 0)

    # Crest factor = peak dBFS − RMS dBFS. Always ≥ 0 for valid audio.
    # This forces a neutral number format so commas never appear and parsing always stays reliable.
    $crest = ([Math]::Round($validPeak - $validRms, 6)).ToString([CultureInfo]::InvariantCulture)

    return [PSCustomObject]@{
        Label          = $Label
        MapIndex       = $MapIndex
        Codec          = $Codec
        Channels       = $Channels
        Peak           = $peak
        RMS            = $rms
        Crest          = $crest
        ClippedSamples = $validClip
        LogFile        = $log
    }
}

function Set-MetricStatus {
    param([PSCustomObject]$m)

    # The peak check uses “greater than 0.0 dB” because anything above that is definitely too loud.
    # The clipping check treats exactly 0.0 dB as clipping if any clipped samples are detected.
    # A case where the peak is exactly 0.0 dB but shows zero clipped samples cannot happen with FFmpeg’s astats.
    # A true 0.0 dB peak always produces at least one counted peak, so these checks fully cover all real situations.
    $peakFail  = ($m.Peak  -ne "N/A" -and [double]::Parse($m.Peak,  [CultureInfo]::InvariantCulture) -gt 0.0)
    $rmsFail   = ($m.RMS   -ne "N/A" -and [double]::Parse($m.RMS,   [CultureInfo]::InvariantCulture) -gt -12.0)
    $crestFail = ($m.Crest -ne "N/A" -and [double]::Parse($m.Crest, [CultureInfo]::InvariantCulture) -lt 6.0)

    # clipFail condition 1: peak at or above 0 dBFS with any peak count -> true clipping.
    # clipFail condition 2: peak count > 5 regardless of peak level -> persistent peak samples.
    # ClippedSamples is always stored as [int] — no "N/A" guard needed for that field.
    $clipFail  = ($m.Peak -ne "N/A" -and [double]::Parse($m.Peak, [CultureInfo]::InvariantCulture) -ge 0.0 -and
              [int]$m.ClippedSamples -gt 0) -or
             ([int]$m.ClippedSamples -gt 5)

    $status = ($peakFail -or $rmsFail -or $crestFail -or $clipFail) ? "FAIL" : "PASS"

    $m | Add-Member -MemberType NoteProperty -Name Status -Value $status -Force
    return $m
}

# --- Quick Interpretation ---

function Format-Metrics {
    param([PSCustomObject]$m)

    if ($m.Peak -eq "N/A" -or $m.RMS -eq "N/A" -or $m.Crest -eq "N/A" -or $m.ClippedSamples -eq "N/A") {
        return @("- Metrics unavailable for this track")
    }

    $peak  = [double]::Parse($m.Peak,  [CultureInfo]::InvariantCulture)
    $rms   = [double]::Parse($m.RMS,   [CultureInfo]::InvariantCulture)
    $crest = [double]::Parse($m.Crest, [CultureInfo]::InvariantCulture)
    $clip  = [int]$m.ClippedSamples

    $lines = @()

    # Peak interpretation
    if ($peak -ge 0) {
        $lines += "- Peak is at or above 0 dB (risk of clipping) [WARNING]"
        $lines += "  [0 dB and above range]"
    } elseif ($peak -ge -1) {
        $lines += "- Peak is very hot (close to clipping) [WARNING]"
        $lines += "  [-1 to 0 dB range]"
    } elseif ($peak -ge -3) {
        $lines += "- Peak is healthy (good headroom) [PASS]"
        $lines += "  [-3 to -1 dB range]"
    } else {
        $lines += "- Peak is low (very safe headroom) [PASS]"
        $lines += "  [below -3 dB range]"
    }

    # RMS interpretation
    if ($rms -gt -12) {
        $lines += "- RMS is very loud (compressed mix) [WARNING]"
        $lines += "  [above -12 dB range]"
    } elseif ($rms -gt -20) {
        $lines += "- RMS is moderately loud (typical TV mix) [PASS]"
        $lines += "  [-20 to -12 dB range]"
    } elseif ($rms -gt -30) {
        $lines += "- RMS is normal for movies (healthy dynamics) [PASS]"
        $lines += "  [-30 to -20 dB range]"
    } else {
        $lines += "- RMS is very low (quiet or highly dynamic) [PASS]"
        $lines += "  [below -30 dB range]"
    }

    # Crest factor interpretation
    if ($crest -lt 6) {
        $lines += "- Crest factor is low (heavily compressed audio) [WARNING]"
        $lines += "  [below 6 dB range]"
    } elseif ($crest -lt 12) {
        $lines += "- Crest factor is moderate (balanced dynamics) [PASS]"
        $lines += "  [6 to 12 dB range]"
    } elseif ($crest -lt 20) {
        $lines += "- Crest factor is high (good dynamic range) [PASS]"
        $lines += "  [12 to 20 dB range]"
    } else {
        $lines += "- Crest factor is extremely dynamic (very wide range) [PASS]"
        $lines += "  [above 20 dB range]"
    }

    # Clipping interpretation
    if ($peak -ge 0 -and $clip -gt 50) {
        $lines += "- Clipping detected (audio distortion likely) [WARNING]"
        $lines += "  [50 or more clips]"
    } elseif ($peak -ge 0 -and $clip -gt 0) {
        $lines += "- Peak at or above 0 dB with clipped samples detected [WARNING]"
        $lines += "  [1 or more clips]"
    } elseif ($clip -gt 5) {
        $lines += "- Some clipping detected (minor distortion possible) [WARNING]"
        $lines += "  [6 or more clips below 0 dB]"
    } elseif ($clip -gt 0) {
        $lines += "- A few peaks detected (not clipping) [PASS]"
        $lines += "  [1 to 5 clips]"
    } else {
        $lines += "- No clipping detected [PASS]"
        $lines += "  [0 clips]"
    }

    return $lines
}

# --- Validate environment ---

foreach ($bin in 'ffmpeg','ffprobe') {
    if (-not (Get-Command $bin -ErrorAction SilentlyContinue)) {
        Write-Err "$bin not found on PATH."
        exit 12
    }
}

# --- Validate input ---

if (-not (Test-Path -LiteralPath $InputFile)) {
    Write-Err "File not found: $InputFile"
    exit 1
}

$fileSizeMB = (Get-Item -LiteralPath $InputFile).Length / 1MB

# --- Get file duration for ETA ---

Write-Info "Probing file duration..."
$fileDuration = Get-FileDuration $InputFile
if ($fileDuration -gt 0) {
    Write-Info ("  Duration: {0:hh\:mm\:ss}" -f [TimeSpan]::FromSeconds($fileDuration))
} else {
    Write-Warn "  Could not determine duration  -  ETA estimates will show 0m 00s."
    Write-Warn "  Timeout guardrail will fall back to 3600 s per track."
}

# --- Probe audio streams ---

Write-Info "Probing audio streams in: $InputFile"

$probe = & ffprobe -v error -select_streams a `
    -show_entries stream=index,codec_name,channels `
    -of csv=p=0 "$InputFile"

if (-not $probe) {
    Write-Err "No audio streams found."
    exit 1
}

# Build audio stream list
$audioStreams = $probe | ForEach-Object {
    $p = $_.Split(",")
    if ($p.Count -lt 3) { return }
    $idx = 0; $ch = 0
    if (-not [int]::TryParse($p[0], [ref]$idx)) { return }
    if (-not [int]::TryParse($p[2], [ref]$ch))  { return }
    [PSCustomObject]@{
        FFProbeIndex = $idx
        Codec        = $p[1].ToLowerInvariant()
        Channels     = $ch
    }
} | Where-Object { $_ }

# Assign ffmpeg map index (0-based among audio streams)
$i = 0
foreach ($s in $audioStreams) {
    $s | Add-Member -MemberType NoteProperty -Name MapIndex -Value $i -Force
    $i++
}

# Filter supported codecs
$supported = $audioStreams | Where-Object {
    $_.Codec -in @("truehd", "aac", "eac3", "dts")
}

if ($supported.Count -eq 0) {
    Write-Err "No supported audio codecs found (truehd, aac, eac3, dts)."
    exit 1
}

Write-Info "Supported audio streams detected:"
foreach ($s in $supported) {
    Write-Host ("  FFProbeIndex={0}, MapIndex={1}, Codec={2}, Channels={3}" -f `
        $s.FFProbeIndex, $s.MapIndex, $s.Codec, $s.Channels)
}

# --- Identify source/downmix ---

$source  = $supported | Where-Object { $_.Codec -in @("truehd","aac","dts") -and $_.Channels -eq 8 } | Select-Object -First 1
$downmix = $supported | Where-Object { $_.Codec -eq "eac3" -and $_.Channels -eq 6 } | Select-Object -First 1

# The script sets the source label once so later analysis sections do not repeat that work.
$srcLabel = if ($source) {
    switch ($source.Codec) {
        'truehd' { 'Source TrueHD 7.1' }
        'dts'    { 'Source DTS-HD 7.1' }
        default  { 'Source AAC 7.1' }
    }
}

$results = [List[object]]::new()

# --- Analysis logic ---

if ($source -and $downmix) {
    Write-Info "Mode: Dual-track analysis (7.1 source + DDP 5.1)."

    $src = Measure-Track $InputFile $source.MapIndex  $srcLabel          "source"  $source.Codec  $fileDuration $fileSizeMB $source.Channels
    $dm  = Measure-Track $InputFile $downmix.MapIndex "Downmix DDP 5.1" "downmix" $downmix.Codec $fileDuration $fileSizeMB $downmix.Channels

    $results.Add((Set-MetricStatus $src))
    $results.Add((Set-MetricStatus $dm))

}
elseif ($source -and -not $downmix) {
    Write-Info "Mode: Single-track analysis (7.1 source only)."

    $src = Measure-Track $InputFile $source.MapIndex $srcLabel "single" $source.Codec $fileDuration $fileSizeMB $source.Channels
    $results.Add((Set-MetricStatus $src))

}
elseif (-not $source -and $downmix) {
    Write-Info "Mode: Single-track analysis (DDP 5.1 only)."

    $dm = Measure-Track $InputFile $downmix.MapIndex "Downmix DDP 5.1" "single" $downmix.Codec $fileDuration $fileSizeMB $downmix.Channels
    $results.Add((Set-MetricStatus $dm))

}
else {
    Write-Err "Supported codecs found, but no valid 7.1 or DDP 5.1 combination."
    exit 1
}

# ==================
# Build JSON Schema 
# ==================

# Build schema-driven structure
$schema = [PSCustomObject]@{
    file     = $InputFile
    duration = $fileDuration
    analysis = [ordered]@{}
}

# Insert source/downmix if present
foreach ($r in $results) {
    $role = if ($r.Label -like "Source*") { "source" }
            elseif ($r.Label -like "Downmix*") { "downmix" }
            else { "track$($r.MapIndex)" }

    $schema.analysis[$role] = [PSCustomObject]@{
        codec    = $r.Codec
        channels = $r.Channels
        metrics  = [PSCustomObject]@{
            peak    = $r.Peak
            rms     = $r.RMS
            crest   = $r.Crest
            clipped = $r.ClippedSamples
        }
        status = $r.Status
        log     = $r.LogFile
    }
}

# If JSON mode is enabled, output JSON and exit
if ($JsonMode) {
    $schema | ConvertTo-Json -Depth 10
    exit 0
}

# --- Summary ---

Write-Host ""
Write-Host "=== AUDIO VALIDATION SUMMARY ===" -ForegroundColor Yellow
Write-Host "File: $InputFile"
Write-Host ""

foreach ($r in $results) {
    $color = ($r.Status -eq "PASS") ? "Green" : "Red"

    Write-Host ("[{0}] (Track {1})" -f $r.Label, $r.MapIndex) -ForegroundColor Cyan
    Write-Host ("  Peak:   {0} dB" -f $r.Peak)
    Write-Host ("  RMS:    {0} dB" -f $r.RMS)
    Write-Host ("  Crest:  {0} dB" -f $r.Crest)
    Write-Host ("  Clips:  {0}" -f $(($r.ClippedSamples -eq "N/A") ? "N/A" : [int]$r.ClippedSamples))
    Write-Host ("  Status: {0}"    -f $r.Status) -ForegroundColor $color
    Write-Host ("  Log:    {0}"    -f $r.LogFile)
    Write-Host ""

    # --- Quick Interpretation ---
    $interpretation = Format-Metrics $r
    Write-Host "  Quick Interpretation:" -ForegroundColor DarkYellow
    foreach ($line in $interpretation) {
        Write-Host "    $line"
    }
    Write-Host ""
}

if ($results.Count -eq 2) {
    $src = $results | Where-Object { $_.Label -like "Source*" }
    $dm  = $results | Where-Object { $_.Label -like "Downmix*" }

    if ($src -and $dm) {
        Write-Host "=== SOURCE vs DOWNMIX COMPARISON ===" -ForegroundColor Yellow
        Write-Host ("Source Peak:   {0} dB" -f $src.Peak)
        Write-Host ("Downmix Peak:  {0} dB" -f $dm.Peak)
        Write-Host ("Source RMS:    {0} dB" -f $src.RMS)
        Write-Host ("Downmix RMS:   {0} dB" -f $dm.RMS)
        Write-Host ("Source Crest:  {0} dB" -f $src.Crest)
        Write-Host ("Downmix Crest: {0} dB" -f $dm.Crest)
        Write-Host ""
    }
}

Write-Host "=== End of Summary ===" -ForegroundColor Yellow
Write-Host ""