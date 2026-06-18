# ==============================================================================
# Script: AudioPeakRMSChecker.ps1  -  (Version 5.1.1)
# Compatible: PS 7.6 | FFmpeg 8.1
#
# Single-pass audio validation for MKV files.
#
#  In ONE ffmpeg decode pass per track, runs a chained filtergraph:
#  astats   -> Peak level dB, RMS level dB, Crest factor, Peak count
#              (PASS/FAIL gate; the contract DDP51.ps1 / Keep71.ps1 depend on)
#  ebur128  -> EBU R128 Integrated loudness (I), Loudness Range (LRA),
#              thresholds, LRA low/high, and true peak (dBTP)
#              (REPORTED, never gates)
#
#  Supported source/downmix/secondary tracks:
#    TrueHD 7.1, AAC 7.1, DTS-HD 7.1, DTS-HD MA 5.1, DDP (E-AC-3) 5.1,
#    DD (AC-3) 5.1, AAC 5.1, AAC 2.0, Opus 2.0, DDP 2.0
#    (single-track, multi-track, or combinations).
#
#  Reports A/V sync (primary source track) and, when a 7.1 source and a 5.1
#  downmix (E-AC-3 preferred, AC-3 accepted) are both present, source-vs-downmix
#  peak/RMS and loudness comparisons.
#
# USAGE:
#   Windows:     pwsh -ExecutionPolicy Bypass -File .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv"
#   macOS/Linux: pwsh -File ./AudioPeakRMSChecker.ps1 "./YourMovie.mkv"
#   Manual:      ... .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv" -Track 1 -JsonMode
# ==============================================================================

#Requires -Version 7.6
using namespace System.Globalization
using namespace System.Collections.Generic

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputFile,

    # Manual override: 0-based audio stream index (0:a:N). -1 = auto-detect.
    [int]$Track = -1,

    # LU window for the source-vs-downmix loudness "preserved" verdict.
    [double]$Tolerance = 1.0,

    # Tunable timeout guardrail.
    [int]$MaxRuntimeMultiplier = 4,
    [int]$MinTimeoutSeconds    = 60,

    [switch]$JsonMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# PS 7.4+ makes native commands throw on non-zero exit when EAP=Stop. Disable
# that so explicit $LASTEXITCODE / ExitCode checks keep control of exit codes.
$PSNativeCommandUseErrorActionPreference = $false

$ScriptVersion = '5.1.1'

# --- Constants ---
$INF_FLOOR        = -70.0   # loudness '-inf' substitute (R128 gate floor)
$ASTATS_INF_FLOOR = -100.0  # astats peak/RMS '-inf' substitute (silent audio)
$LoudnessOverhead = 1.25    # ebur128 K-weighting cost on top of the astats pass

# --- Binary resolution ---
$ext     = $IsWindows ? '.exe' : ''
$ffmpeg  = Join-Path $PSScriptRoot "ffmpeg$ext"
$ffprobe = Join-Path $PSScriptRoot "ffprobe$ext"

# --- JSON toggle (change $false to $true to default-on) ---
$DefaultJsonMode = $false
$JsonMode = ($JsonMode.IsPresent -or $DefaultJsonMode)

# ==========================================
#  Console styling (suppressed in JSON mode
# ==========================================
function Write-Info    { param([string]$m) if (-not $JsonMode) { Write-Host $m -ForegroundColor Cyan } }
function Write-Section { param([string]$m) if (-not $JsonMode) { Write-Host $m -ForegroundColor White } }
function Write-Warn    { param([string]$m) if (-not $JsonMode) { Write-Host $m -ForegroundColor Yellow } }
function Write-Err     { param([string]$m) Write-Host $m -ForegroundColor Red }
function Write-Plain   { param([string]$m) if (-not $JsonMode) { Write-Host $m } }

# ===================================
#  Invariant-culture numeric parsing
# ===================================
function Try-ParseDouble {
    # Parse using InvariantCulture / NumberStyles.Float only. Returns $null on
    # failure. '-inf'/'+inf'/'inf' map to the loudness floor sentinel.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()
    if ($t -match '^[+-]?inf(inity)?$') { return $script:INF_FLOOR }
    [double]$value = 0.0
    if ([double]::TryParse($t, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$value)) { return $value }
    return $null
}

function Format-Lufs {
    param($Value, [string]$Unit = 'LUFS')
    if ($null -eq $Value) { return 'N/A' }
    return ('{0} {1}' -f $Value.ToString('0.0', [CultureInfo]::InvariantCulture), $Unit)
}

function Get-Metric {
    param([string[]]$Lines, [string]$Pattern)
    $m = $Lines | Select-String -Pattern $Pattern -SimpleMatch | Select-Object -First 1
    if ($m -and $m.Line -match ':\s*(\S+)\s*$') { return $matches[1] }
    return 'N/A'
}

# ================================
#  Binary resolution / validation
# ================================
function Resolve-Binary {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-Err "[ERROR] Required binary not found next to script: $Path"
        exit 12
    }
    if (-not $IsWindows) {
        $mode = (Get-Item -LiteralPath $Path).UnixFileMode
        if (-not ($mode -band [System.IO.UnixFileMode]::UserExecute)) {
            Write-Err "[ERROR] Binary not executable: $Path  -  run: chmod +x `"$Path`""
            exit 12
        }
    }
}

# --- ffprobe (single combined call -> JSON, by-name access) ---
function Get-MediaProbe {
    param([string]$FFprobe, [string]$File)
    $raw = (& $FFprobe -v error -show_format -show_streams -of json $File) 2>$null
    if ($LASTEXITCODE -ne 0) { return $null }
    $jsonText = ($raw -join "`n")
    if ([string]::IsNullOrWhiteSpace($jsonText)) { return $null }
    try { return ($jsonText | ConvertFrom-Json) } catch { return $null }
}

function Get-ProbeProp {
    param($Node, [string]$Name)
    if (($null -ne $Node) -and ($Node.PSObject.Properties.Name -contains $Name)) {
        $val = $Node.$Name
        if ($null -ne $val) { return [string]$val }
    }
    return $null
}

function Get-ProbeStreams {
    param($Probe)
    if (($null -eq $Probe) -or (-not ($Probe.PSObject.Properties.Name -contains 'streams'))) { return @() }
    $s = $Probe.streams
    return ($null -eq $s) ? @() : @($s)
}

function Get-DurationFromProbe {
    param($Probe)
    if (($null -eq $Probe) -or (-not ($Probe.PSObject.Properties.Name -contains 'format'))) { return 0.0 }
    $d = Get-ProbeProp $Probe.format 'duration'
    if ($null -eq $d) { return 0.0 }
    $val = Try-ParseDouble $d.Trim()
    return ($null -ne $val) ? $val : 0.0
}

function Get-VideoStartFromProbe {
    # $null -> no video stream;  NaN -> video present but no timing;  double -> offset (s)
    param($Probe)
    $video = Get-ProbeStreams $Probe |
        Where-Object { (Get-ProbeProp $_ 'codec_type') -eq 'video' } | Select-Object -First 1
    if ($null -eq $video) { return $null }
    $st = Get-ProbeProp $video 'start_time'
    if (($null -eq $st) -or ($st.Trim() -eq 'N/A')) { return [double]::NaN }
    $val = Try-ParseDouble $st.Trim()
    return ($null -ne $val) ? $val : [double]::NaN
}

function Get-AudioStreamsFromProbe {
    # Audio streams only, re-indexed: MapIndex is the audio-relative position
    # (0,1,2...) matching ffmpeg's 0:a:N selector; FFprobeIndex keeps the
    # absolute container index. start_time feeds A/V sync + TrueHD preroll;
    # duration + nb_frames drive empty/placeholder-track detection.
    param($Probe)

    $list = [List[object]]::new()
    $mapIndex = 0
    foreach ($s in (Get-ProbeStreams $Probe)) {
        if ($null -eq $s) { continue }
        if ((Get-ProbeProp $s 'codec_type') -ne 'audio') { continue }

        $ffprobeIndex = Get-ProbeProp $s 'index'
        $codec        = Get-ProbeProp $s 'codec_name'
        $chStr        = Get-ProbeProp $s 'channels'
        $stStr        = Get-ProbeProp $s 'start_time'
        $durStr       = Get-ProbeProp $s 'duration'
        $nbStr        = Get-ProbeProp $s 'nb_frames'

        $ffprobeIndex = ($null -eq $ffprobeIndex) ? '' : $ffprobeIndex.Trim()
        $codec        = ($null -eq $codec) ? '' : $codec.Trim().ToLowerInvariant()

        $chParsed = 0
        if ($null -ne $chStr) {
            [int]::TryParse($chStr.Trim(), [NumberStyles]::Integer, [CultureInfo]::InvariantCulture, [ref]$chParsed) | Out-Null
        }

        $startTime = 0.0
        if ($null -ne $stStr -and $stStr.Trim() -ne 'N/A') {
            $stVal = Try-ParseDouble $stStr.Trim()
            if ($null -ne $stVal) { $startTime = $stVal }
        }

        $dur = ($null -ne $durStr) ? (Try-ParseDouble $durStr.Trim()) : $null
        $nb  = $null
        if ($null -ne $nbStr) {
            $nbTmp = 0L
            if ([long]::TryParse($nbStr.Trim(), [NumberStyles]::Integer, [CultureInfo]::InvariantCulture, [ref]$nbTmp)) { $nb = $nbTmp }
        }

        # Empty only when NO usable duration AND an explicit nb_frames of 0.
        # A missing/unparseable nb_frames (common for MKV copy streams) is NON-empty.
        $durEmpty = ($null -eq $dur) -or ($dur -le 0.0)
        $nbEmpty  = ($null -ne $nb) -and ($nb -eq 0)
        $isEmpty  = $durEmpty -and $nbEmpty

        $list.Add([PSCustomObject]@{
            MapIndex     = $mapIndex
            FFprobeIndex = $ffprobeIndex
            Codec        = $codec
            Channels     = $chParsed
            StartTime    = $startTime
            Duration     = $dur
            NbFrames     = $nb
            IsEmpty      = $isEmpty
        })
        $mapIndex++
    }
    return ,$list
}

# ==========================================================================
#  ETA / timeout. Per-codec processing ratios (proc time / file duration),
#  scaled by $LoudnessOverhead because each pass runs astats AND ebur128.
# ==========================================================================
function Get-TrackETA {
    param([string]$Codec, [double]$FileDuration, [int]$Channels, [double]$SizeMB)
    $sizeScale = ($Codec -in @('truehd','aac','dts')) ? [Math]::Max(1.0, $SizeMB / 15360) : 1.0
    $ratio = switch ($Codec) {
        'truehd' { 0.03500 }
        'aac'    { 0.00350 }
        'eac3'   { 0.00365 }
        'ac3'    { 0.00365 }
        'dts'    { 0.02400 }
        'opus'   { 0.00350 }
        default  { 0.00365 }
    }
    # astats+ebur128 cost tracks sample throughput (duration x channels), but the
    # lossy-codec ratios above were tuned on stereo, so they undershoot badly on
    # 5.1/7.1 content (e.g. AAC 7.1 ran ~4.5x over a stereo-baselined estimate).
    # Scale those by channels-over-2; the >=1 floor keeps mono/stereo unchanged.
    # truehd/dts ratios are already measured on multichannel material, so scaling
    # them would double-count -- left at 1.0. Anchored to a real AAC 7.1 sample:
    # 0.00350 x (8/2) x $LoudnessOverhead matches observed wall time within rounding.
    $chScale = ($Codec -in @('truehd','dts')) ? 1.0 : [Math]::Max(1.0, $Channels / 2.0)
    return [int]($FileDuration * $ratio * $chScale * $sizeScale * $script:LoudnessOverhead)
}

# =============================================================================
#  ebur128 Summary parsing (block-aware). 'Threshold:' appears under BOTH
#  'Integrated loudness:' and 'Loudness range:', so the sub-block is tracked.
#  Anchors on the last 'Summary:' line; ignores the ~10 Hz per-frame log lines.
#  Throws [FormatException] on a present-but-non-numeric token. Returns $null
#  if the Summary / required fields are absent.
# =============================================================================
function Parse-Ebur128Summary {
    param([string[]]$Lines)

    $summaryIdx = -1
    for ($i = $Lines.Count - 1; $i -ge 0; $i--) {
        if ($Lines[$i] -match 'Summary:\s*$') { $summaryIdx = $i; break }
    }
    if ($summaryIdx -lt 0) { return $null }

    $section = ''
    $out = @{
        I_LUFS               = $null
        Int_Threshold_LUFS   = $null
        LRA_LU               = $null
        Range_Threshold_LUFS = $null
        LRA_Low_LUFS         = $null
        LRA_High_LUFS        = $null
    }

    $extract = {
        param([string]$Line)
        if ($Line -match ':\s*([+-]?(?:inf(?:inity)?|\d+(?:\.\d+)?))\s*(?:LUFS|LU|dBFS)?\s*$') {
            $tok = $Matches[1]
            $v = Try-ParseDouble $tok
            if ($null -eq $v) { throw [System.FormatException]::new("Unparseable loudness token: '$tok'") }
            return $v
        }
        return $null
    }

    # True peak is optional: a build without libswresample may omit it, and that
    # must NOT invalidate the (required) integrated/range fields. Tracked
    # separately and attached only after the required-field check passes.
    $truePeak = $null

    for ($i = $summaryIdx + 1; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        if ($null -eq $line) { continue }
        if ($line -match 'Summary:\s*$') { break }

        if     ($line -match 'Integrated loudness:') { $section = 'integrated'; continue }
        elseif ($line -match 'Loudness range:')      { $section = 'range';      continue }
        elseif ($line -match 'True peak:')           { $section = 'truepeak';   continue }
        elseif ($line -match 'Sample peak:')         { $section = 'samplepeak'; continue }

        if ($section -eq 'integrated') {
            if     ($line -match '^\s*I:')         { $out.I_LUFS              = (& $extract $line) }
            elseif ($line -match '^\s*Threshold:') { $out.Int_Threshold_LUFS  = (& $extract $line) }
        }
        elseif ($section -eq 'range') {
            if     ($line -match '^\s*LRA:')       { $out.LRA_LU               = (& $extract $line) }
            elseif ($line -match '^\s*Threshold:') { $out.Range_Threshold_LUFS = (& $extract $line) }
            elseif ($line -match '^\s*LRA low:')   { $out.LRA_Low_LUFS         = (& $extract $line) }
            elseif ($line -match '^\s*LRA high:')  { $out.LRA_High_LUFS        = (& $extract $line) }
        }
        elseif ($section -eq 'truepeak') {
            if ($line -match '^\s*Peak:') { $truePeak = (& $extract $line) }
        }
        # 'samplepeak' section is intentionally not captured (astats already
        # provides sample peak); tracking the section just prevents its Peak:
        # line from being mistaken for the true-peak value.
    }

    # Required fields: integrated + range only. True peak is additive.
    foreach ($k in @($out.Keys)) {
        if ($null -eq $out[$k]) { return $null }
    }
    $out.TruePeak_dBFS = $truePeak
    return $out
}

# =============================================================================
#  Measure one track in a single decode pass (astats,ebur128).
#  astats failures are FATAL (the gate). ebur128 failures are NON-fatal:
#  loudness degrades to $null + warning and never alters the exit code or gate.
# =============================================================================
function Measure-Track {
    param(
        [string]$InputFile,
        [int]$MapIndex,
        [string]$Label,
        [string]$LogSuffix,
        [string]$Codec,
        [double]$TotalDuration,
        [double]$FileSizeMB,
        [int]$Channels,
        [double]$StartTime = 0.0
    )

    $outDir = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($InputFile))
    $base   = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $log    = Join-Path $outDir "$base-$LogSuffix-astats.txt"

    $tmpErr = [System.IO.Path]::GetTempFileName()
    $tmpOut = [System.IO.Path]::GetTempFileName()
    if (-not $IsWindows) { chmod 600 $tmpErr; chmod 600 $tmpOut }

    $etaSec = Get-TrackETA $Codec $TotalDuration $Channels $FileSizeMB
    $etaStr = "$([Math]::Floor($etaSec / 60))m $(([int]($etaSec % 60)).ToString('D2'))s"
    Write-Info "Analyzing $Label (map 0:a:$MapIndex)  -  estimated ~$etaStr ..."

    # astats upstream of ebur128: ebur128 converts to double-precision float
    # internally, so astats must measure the natively decoded signal first.
    $ffArgs = @(
        '-hide_banner'
        '-nostats'
        '-vn'
        '-sn'
        '-analyzeduration', '200M'
        '-probesize',       '200M'
        '-i', $InputFile
        '-map', "0:a:$MapIndex"
        '-af', 'astats=measure_overall=all:measure_perchannel=none,ebur128=peak=true'
        '-f', 'null'
        '-'
    )

    $procParams = @{
        FilePath               = $script:ffmpeg
        ArgumentList           = $ffArgs
        RedirectStandardError  = $tmpErr
        RedirectStandardOutput = $tmpOut
        NoNewWindow            = $true
        PassThru               = $true
    }
    $proc = Start-Process @procParams
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $maxSec = ($TotalDuration -gt 0) ?
        [Math]::Max($MinTimeoutSeconds, $etaSec * $MaxRuntimeMultiplier) :
        [Math]::Max($MinTimeoutSeconds, 3600)

    $out = $null
    try {
        while (-not $proc.HasExited) {
            if ($sw.Elapsed.TotalSeconds -gt $maxSec) {
                Write-Err "FFmpeg exceeded maximum allowed runtime for $Label (map 0:a:$MapIndex)."
                Write-Err "Elapsed: $([int]$sw.Elapsed.TotalSeconds)s  |  Limit: $maxSec s"
                try { $proc.Kill(); $proc.WaitForExit(2000) | Out-Null } catch {}
                exit 5
            }
            $remaining = [Math]::Max(0, $etaSec - [int]$sw.Elapsed.TotalSeconds)
            if (-not $JsonMode) {
                $msg = ("  ETA: $([Math]::Floor($remaining / 60))m $(([int]($remaining % 60)).ToString('D2'))s remaining...").PadRight(35)
                [Console]::Write("`r" + $msg)
            }
            Start-Sleep -Seconds 1
        }
        if (-not $JsonMode) { [Console]::Write("`r" + (" " * 40) + "`r") }
        $e = $sw.Elapsed
        Write-Info "  Done in $([Math]::Floor($e.TotalMinutes))m $($e.Seconds.ToString('D2'))s"

        $out = Get-Content $tmpErr -Encoding UTF8 -ErrorAction SilentlyContinue

        if (-not $out -or $out.Count -eq 0) {
            Write-Err "FFmpeg produced no stderr output for $Label (map 0:a:$MapIndex). Analysis aborted."
            exit 2
        }
        $exit = $proc.ExitCode
        if ($exit -ne 0) {
            Write-Err "FFmpeg exited with code $exit while analyzing $Label (map 0:a:$MapIndex)."
            Write-Err "Check the log file: $log"
            exit 3
        }

        $peakLine = $out | Select-String -Pattern 'Peak level dB' -SimpleMatch
        $rmsLine  = $out | Select-String -Pattern 'RMS level dB'  -SimpleMatch
        $clipLine = $out | Select-String -Pattern '] Peak count'  -SimpleMatch
        if (-not $peakLine -or -not $rmsLine -or -not $clipLine) {
            Write-Err "FFmpeg produced incomplete astats output for $Label (map 0:a:$MapIndex)."
            Write-Err "Required metrics missing: Peak/RMS/Peak count."
            exit 6
        }
    }
    finally {
        Remove-Item $tmpErr, $tmpOut -Force -ErrorAction SilentlyContinue
    }

    # Persist a per-track log: keep astats + the ebur128 Summary, drop the
    # ~10 Hz ebur128 per-frame lines (anchored by "] t: ") and container noise.
    $out | Where-Object {
        $_ -notmatch 'Subtitle:|Chapters:|Chapter #|^\s+start \d|^\s+title\s*:|^\s+Metadata:|\] t:\s'
    } | Out-File -LiteralPath $log -Encoding UTF8
    Write-Plain "  Saved log: $log"

    if (-not (Test-Path -LiteralPath $log)) {
        Write-Err "Log file was not created for $Label (map 0:a:$MapIndex)."
        exit 10
    }
    $logInfo = Get-Item -LiteralPath $log -ErrorAction SilentlyContinue
    if (-not $logInfo -or $logInfo.Length -lt 10) {
        Write-Err "Log file is empty or incomplete for $Label (map 0:a:$MapIndex)."
        exit 11
    }

    # --- astats (gate metrics) ---
    $peak = Get-Metric $out 'Peak level dB'
    $rms  = Get-Metric $out 'RMS level dB'
    $clip = Get-Metric $out '] Peak count'

    if ($peak -match '^-?inf') { $peak = $script:ASTATS_INF_FLOOR.ToString([CultureInfo]::InvariantCulture) }
    if ($rms  -match '^-?inf') { $rms  = $script:ASTATS_INF_FLOOR.ToString([CultureInfo]::InvariantCulture) }

    $validPeak = 0.0; $validRms = 0.0; $validClipDouble = 0.0
    if (-not [double]::TryParse($peak, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validPeak)) {
        Write-Err "Peak level dB value '$peak' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 7
    }
    if (-not [double]::TryParse($rms, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validRms)) {
        Write-Err "RMS level dB value '$rms' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 8
    }
    if (-not [double]::TryParse($clip, [NumberStyles]::Float, [CultureInfo]::InvariantCulture, [ref]$validClipDouble)) {
        Write-Err "Peak count value '$clip' failed numeric parse for $Label (map 0:a:$MapIndex)."
        exit 9
    }
    $validClip = [int][Math]::Round($validClipDouble, 0)
    $crest = ([Math]::Round($validPeak - $validRms, 6)).ToString([CultureInfo]::InvariantCulture)

    # --- ebur128 (loudness; NON-fatal) ---
    $loud = $null
    try {
        $loud = Parse-Ebur128Summary -Lines $out
    }
    catch [System.FormatException] {
        Write-Warn "  Loudness: ebur128 produced a non-numeric value for $Label (map 0:a:$MapIndex) - reported as N/A."
        $loud = $null
    }
    if ($null -eq $loud) {
        Write-Warn "  Loudness: ebur128 Summary unavailable for $Label (map 0:a:$MapIndex) - reported as N/A."
    }

    return [PSCustomObject]@{
        Label          = $Label
        MapIndex       = $MapIndex
        Codec          = $Codec
        Channels       = $Channels
        Peak           = $peak
        RMS            = $rms
        Crest          = $crest
        ClippedSamples = $validClip
        StartTime      = $StartTime
        LogFile        = $log
        Loudness       = $loud
    }
}

# ============================================================
#  Gate evaluation (astats only; loudness never participates).
# ============================================================
function Set-MetricStatus {
    param([PSCustomObject]$m)
    $peakFail  = ($m.Peak  -ne 'N/A' -and [double]::Parse($m.Peak,  [CultureInfo]::InvariantCulture) -gt 0.0)
    $rmsFail   = ($m.RMS   -ne 'N/A' -and [double]::Parse($m.RMS,   [CultureInfo]::InvariantCulture) -gt -12.0)
    $crestFail = ($m.Crest -ne 'N/A' -and [double]::Parse($m.Crest, [CultureInfo]::InvariantCulture) -lt 6.0)
    $clipFail  = (($m.Peak -ne 'N/A' -and [double]::Parse($m.Peak, [CultureInfo]::InvariantCulture) -ge 0.0 -and
                  [int]$m.ClippedSamples -gt 0) -or ([int]$m.ClippedSamples -gt 5))

    $status = ($peakFail -or $rmsFail -or $crestFail -or $clipFail) ? 'FAIL' : 'PASS'
    $m | Add-Member -MemberType NoteProperty -Name Status -Value $status -Force
    return $m
}

# ================================================================
#  Quick interpretation (peak / RMS / crest / clipping / preroll).
# ================================================================
function Format-Metrics {
    param([PSCustomObject]$m)
    if ($m.Peak -eq 'N/A' -or $m.RMS -eq 'N/A' -or $m.Crest -eq 'N/A') {
        return @('- Metrics unavailable for this track')
    }
    $peak  = [double]::Parse($m.Peak,  [CultureInfo]::InvariantCulture)
    $rms   = [double]::Parse($m.RMS,   [CultureInfo]::InvariantCulture)
    $crest = [double]::Parse($m.Crest, [CultureInfo]::InvariantCulture)
    $clip  = [int]$m.ClippedSamples
    $lines = @()

    if ($peak -ge 0)        { $lines += '- Peak is at or above 0 dB (risk of clipping) [WARNING]'; $lines += '  [0 dB and above range]' }
    elseif ($peak -ge -1)   { $lines += '- Peak is very hot (close to clipping) [WARNING]';        $lines += '  [-1 to 0 dB range]' }
    elseif ($peak -ge -3)   { $lines += '- Peak is healthy (good headroom) [PASS]';                $lines += '  [-3 to -1 dB range]' }
    else                    { $lines += '- Peak is low (very safe headroom) [PASS]';               $lines += '  [below -3 dB range]' }

    if ($rms -gt -12)       { $lines += '- RMS is very loud (compressed mix) [WARNING]';           $lines += '  [above -12 dB range]' }
    elseif ($rms -gt -20)   { $lines += '- RMS is moderately loud (typical TV mix) [PASS]';        $lines += '  [-20 to -12 dB range]' }
    elseif ($rms -gt -30)   { $lines += '- RMS is normal for movies (healthy dynamics) [PASS]';    $lines += '  [-30 to -20 dB range]' }
    else                    { $lines += '- RMS is very low (quiet or highly dynamic) [PASS]';      $lines += '  [below -30 dB range]' }

    if ($crest -lt 6)       { $lines += '- Crest factor is low (heavily compressed audio) [WARNING]'; $lines += '  [below 6 dB range]' }
    elseif ($crest -lt 12)  { $lines += '- Crest factor is moderate (balanced dynamics) [PASS]';      $lines += '  [6 to 12 dB range]' }
    elseif ($crest -lt 20)  { $lines += '- Crest factor is high (good dynamic range) [PASS]';         $lines += '  [12 to 20 dB range]' }
    else                    { $lines += '- Crest factor is extremely dynamic (very wide range) [PASS]'; $lines += '  [above 20 dB range]' }

    if ($peak -ge 0 -and $clip -gt 50)    { $lines += '- Clipping detected (audio distortion likely) [WARNING]';        $lines += '  [50 or more clips]' }
    elseif ($peak -ge 0 -and $clip -gt 0) { $lines += '- Peak at or above 0 dB with clipped samples detected [WARNING]'; $lines += '  [1 or more clips]' }
    elseif ($clip -gt 5)                  { $lines += '- Some clipping detected (minor distortion possible) [WARNING]';  $lines += '  [6 or more clips below 0 dB]' }
    elseif ($clip -gt 0)                  { $lines += '- A few peaks detected (not clipping) [PASS]';                    $lines += '  [1 to 5 clips]' }
    else                                  { $lines += '- No clipping detected [PASS]';                                   $lines += '  [0 clips]' }

    $stProp = $m.PSObject.Properties['StartTime']
    if ($null -ne $stProp -and $m.Codec -eq 'truehd') {
        $prerollMs = [int][Math]::Round([double]$stProp.Value * 1000)
        $lines += "- Preroll: $prerollMs ms - normal TrueHD PTS offset (source); not a sync issue"
    }
    return $lines
}

# =========================
#  A/V sync
# =========================
function Get-AVSyncStatus {
    param(
        [double]$VideoStartTime,
        [PSCustomObject]$AudioStream,
        [double]$TrueHDPrerollMaxMs = 120.0,
        [double]$GeneralToleranceMs = 5.0
    )
    $audioStartMs = $AudioStream.StartTime * 1000.0
    $videoStartMs = $VideoStartTime        * 1000.0
    $offsetMs     = [Math]::Round($audioStartMs - $videoStartMs, 1)
    $codec        = $AudioStream.Codec

    if ($codec -eq 'truehd') {
        if ($offsetMs -ge -$TrueHDPrerollMaxMs -and $offsetMs -le $GeneralToleranceMs) {
            return [PSCustomObject]@{ Status = 'OK'; Message = 'A/V Sync : OK'; OffsetMs = $offsetMs }
        }
    } else {
        if ([Math]::Abs($offsetMs) -le $GeneralToleranceMs) {
            return [PSCustomObject]@{ Status = 'OK'; Message = 'A/V Sync : OK'; OffsetMs = $offsetMs }
        }
    }
    $sign = if ($offsetMs -gt 0) { '+' } else { '' }
    $dir  = if ($offsetMs -gt 0) { 'after' } else { 'before' }
    $abs  = [Math]::Abs($offsetMs)
    return [PSCustomObject]@{
        Status   = 'WARNING'
        Message  = "A/V Sync : WARNING - audio starts $sign$([int]$abs) ms $dir video"
        OffsetMs = $offsetMs
    }
}

# ===============================
#  Validate environment + input
# ===============================
Resolve-Binary $ffprobe
Resolve-Binary $ffmpeg

if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Err "File not found: $InputFile"
    exit 1
}
$fileSizeMB = (Get-Item -LiteralPath $InputFile).Length / 1MB

Write-Section "AudioPeakRMSChecker v$ScriptVersion   (PS 7.6 | FFmpeg 8.1)"

# --- single ffprobe probe (duration + video start + audio streams) ---
Write-Info "Probing media (ffprobe)..."
$probe = Get-MediaProbe $ffprobe $InputFile

$fileDuration = Get-DurationFromProbe $probe
if ($fileDuration -gt 0) {
    Write-Info ("  Duration: {0:hh\:mm\:ss}" -f [TimeSpan]::FromSeconds($fileDuration))
} else {
    Write-Warn "  Could not determine duration  -  ETA estimates will show 0m 00s."
    Write-Warn "  Timeout guardrail will fall back to 3600 s per track."
}

Write-Plain "Probing audio streams in: $InputFile"
$audioStreams = Get-AudioStreamsFromProbe $probe
if (($null -eq $audioStreams) -or ($audioStreams.Count -eq 0)) {
    Write-Err "No audio streams found."
    exit 1
}

# --- video start_time (A/V sync) ---
$videoStartTime = Get-VideoStartFromProbe $probe
if ($null -ne $videoStartTime -and -not [double]::IsNaN($videoStartTime)) {
    Write-Info ("Video start_time: {0} ms" -f [int][Math]::Round($videoStartTime * 1000))
}

# --- supported codecs ---
$supported = $audioStreams | Where-Object { $_.Codec -in @('truehd','aac','eac3','ac3','dts','opus') }
if ($supported.Count -eq 0) {
    Write-Err "No supported audio codecs found (truehd, aac, eac3, ac3, dts, opus)."
    exit 1
}

$codecDisplayMap = @{
    'truehd:8' = 'TrueHD 7.1'
    'truehd:6' = 'TrueHD 5.1'
    'dts:8'    = 'DTS-HD MA 7.1'
    'dts:6'    = 'DTS-HD MA 5.1'
    'aac:8'    = 'AAC 7.1'
    'aac:6'    = 'AAC 5.1'
    'eac3:6'   = 'Dolby Digital Plus 5.1'
    'eac3:8'   = 'Dolby Digital Plus 7.1'
    'eac3:2'   = 'Dolby Digital Plus 2.0'
    'ac3:6'    = 'Dolby Digital 5.1'
    'ac3:8'    = 'Dolby Digital 7.1'
    'ac3:2'    = 'Dolby Digital 2.0'
    'opus:2'   = 'Opus 2.0'
    'aac:2'    = 'AAC 2.0'
}
function Get-CodecDisplay {
    param([string]$Codec, [int]$Channels)
    return $script:codecDisplayMap["$($Codec):$($Channels)"] ?? "$Codec ($Channels`ch)"
}

Write-Section "Supported audio streams detected:"
foreach ($s in $supported) {
    Write-Info ("  FFProbeIndex={0}, MapIndex={1}, Codec={2}, Channels={3}" -f `
        $s.FFprobeIndex, $s.MapIndex, (Get-CodecDisplay $s.Codec $s.Channels), $s.Channels)
}

# =================================
#  Build analyzable track list
# =================================
$analyzable = [List[object]]::new()

if ($Track -ge 0) {
    $match = $audioStreams | Where-Object { $_.MapIndex -eq $Track } | Select-Object -First 1
    if (-not $match) {
        Write-Err "Requested -Track $Track does not exist among the audio streams."
        exit 1
    }
    if ($match.IsEmpty) {
        Write-Err "Requested -Track $Track is an empty/placeholder track (no audio data)."
        exit 1
    }
    if ($match.Codec -notin @('truehd','aac','eac3','ac3','dts','opus')) {
        Write-Err "Requested -Track $Track codec '$($match.Codec)' is not supported."
        exit 1
    }
    $analyzable.Add([PSCustomObject]@{
        Stream   = $match
        Label    = (Get-CodecDisplay $match.Codec $match.Channels)
        Suffix   = "track$Track"
        ModeName = (Get-CodecDisplay $match.Codec $match.Channels)
    })
}
else {
    $valid   = $supported | Where-Object { -not $_.IsEmpty }
    $source  = $valid | Where-Object { $_.Codec -in @('truehd','aac','dts') -and $_.Channels -eq 8 } | Select-Object -First 1
    # Downmix peer: prefer E-AC-3 5.1 (the DDP51/Keep71 engine output), then
    # fall back to AC-3 5.1 (Dolby Digital) so DD downmixes are compared too.
    $downmix = $valid | Where-Object { $_.Codec -eq 'eac3' -and $_.Channels -eq 6 } | Select-Object -First 1
    if (-not $downmix) {
        $downmix = $valid | Where-Object { $_.Codec -eq 'ac3' -and $_.Channels -eq 6 } | Select-Object -First 1
    }
    $aac51   = $valid | Where-Object { $_.Codec -eq 'aac'  -and $_.Channels -eq 6 } | Select-Object -First 1
    $dts51   = $valid | Where-Object { $_.Codec -eq 'dts'  -and $_.Channels -eq 6 } | Select-Object -First 1
    $opus20  = $valid | Where-Object { $_.Codec -eq 'opus' -and $_.Channels -eq 2 } | Select-Object -First 1
    $eac320  = $valid | Where-Object { $_.Codec -eq 'eac3' -and $_.Channels -eq 2 } | Select-Object -First 1
    $aac20   = $valid | Where-Object { $_.Codec -eq 'aac'  -and $_.Channels -eq 2 } | Select-Object -First 1

    $srcLabel = if ($source) {
        switch ($source.Codec) {
            'truehd' { 'Source TrueHD 7.1' }
            'dts'    { 'Source DTS-HD 7.1' }
            default  { 'Source AAC 7.1' }
        }
    }

    if ($source) {
        $analyzable.Add([PSCustomObject]@{ Stream = $source; Label = $srcLabel; Suffix = 'source'; ModeName = '7.1 source' })
    }
    if ($downmix) {
        $dmxName  = ($downmix.Codec -eq 'ac3') ? 'DD 5.1'  : 'DDP 5.1'
        $analyzable.Add([PSCustomObject]@{ Stream = $downmix; Label = "Downmix $dmxName"; Suffix = 'downmix'; ModeName = $dmxName })
    }
    # When source AND downmix are both present, ignore aac51 (preserves v4.x behavior).
    if ($aac51 -and -not ($source -and $downmix)) {
        $analyzable.Add([PSCustomObject]@{ Stream = $aac51; Label = 'AAC 5.1'; Suffix = 'aac51'; ModeName = 'AAC 5.1' })
    }
    if ($dts51) {
        $analyzable.Add([PSCustomObject]@{ Stream = $dts51; Label = 'DTS-HD MA 5.1'; Suffix = 'dts51'; ModeName = 'DTS-HD MA 5.1' })
    }
    if ($opus20) {
        $analyzable.Add([PSCustomObject]@{ Stream = $opus20; Label = 'Opus 2.0'; Suffix = 'opus20'; ModeName = 'Opus 2.0' })
    }
    if ($eac320) {
        $analyzable.Add([PSCustomObject]@{ Stream = $eac320; Label = 'Dolby Digital Plus 2.0'; Suffix = 'eac320'; ModeName = 'DDP 2.0' })
    }
    if ($aac20) {
        $analyzable.Add([PSCustomObject]@{ Stream = $aac20; Label = 'AAC 2.0'; Suffix = 'aac20'; ModeName = 'AAC 2.0' })
    }
}

if ($analyzable.Count -eq 0) {
    Write-Err "Supported codecs found, but no valid 7.1, 5.1, or 2.0 combination."
    exit 1
}

$modeDesc = if ($analyzable.Count -eq 1) {
    "Single-track analysis ($($analyzable[0].ModeName) only)."
} elseif ($analyzable.Count -eq 2) {
    "Dual-track analysis ($(($analyzable | ForEach-Object { $_.ModeName }) -join ' + '))."
} else {
    "$($analyzable.Count)-track analysis."
}
Write-Warn "Mode: $modeDesc"

# Single auto-detected track uses the legacy "single" suffix.
if ($analyzable.Count -eq 1 -and $Track -lt 0) { $analyzable[0].Suffix = 'single' }

# --- A/V sync (primary source track; falls back to first audio stream) ---
$avSync = $null
if ($null -ne $videoStartTime -and -not [double]::IsNaN($videoStartTime)) {
    $syncStream = if ($Track -ge 0) {
        $analyzable[0].Stream
    } else {
        $srcCandidate = $supported | Where-Object { $_.Codec -in @('truehd','aac','dts') -and $_.Channels -eq 8 } | Select-Object -First 1
        if ($srcCandidate) { $srcCandidate } else { $audioStreams | Select-Object -First 1 }
    }
    if ($syncStream) { $avSync = Get-AVSyncStatus $videoStartTime $syncStream }
}

# =======================
#  Analyze
# =======================
$results = [List[object]]::new()
foreach ($a in $analyzable) {
    $r = Measure-Track $InputFile $a.Stream.MapIndex $a.Label $a.Suffix $a.Stream.Codec $fileDuration $fileSizeMB $a.Stream.Channels $a.Stream.StartTime
    $results.Add((Set-MetricStatus $r))
}

# ============================================================
#  Source-vs-downmix loudness comparison (when both measured)
# ============================================================
function Get-LoudnessComparison {
    param([PSCustomObject]$Src, [PSCustomObject]$Dmx, [double]$Tol)
    if (-not $Src -or -not $Dmx -or $null -eq $Src.Loudness -or $null -eq $Dmx.Loudness) { return $null }
    $sI = $Src.Loudness.I_LUFS; $dI = $Dmx.Loudness.I_LUFS
    $sL = $Src.Loudness.LRA_LU; $dL = $Dmx.Loudness.LRA_LU
    $deltaI   = ($null -ne $sI -and $null -ne $dI) ? [Math]::Round(($dI - $sI), 1) : $null
    $deltaLRA = ($null -ne $sL -and $null -ne $dL) ? [Math]::Round(($dL - $sL), 1) : $null
    $verdict  = ($null -ne $deltaI -and [Math]::Abs($deltaI) -le $Tol) ? 'preserved' : 'shifted'
    return [PSCustomObject]@{
        Reference = "0:a:$($Src.MapIndex)"; Target = "0:a:$($Dmx.MapIndex)"
        DeltaI = $deltaI; DeltaLRA = $deltaLRA; Tolerance = $Tol; Verdict = $verdict
    }
}

$srcResult = $results | Where-Object { $_.Label -like 'Source*' } | Select-Object -First 1
$dmxResult = $results | Where-Object { $_.Label -like 'Downmix*' } | Select-Object -First 1
$loudCompare = Get-LoudnessComparison $srcResult $dmxResult $Tolerance

# =========================
#  JSON schema
# =========================
function Get-LoudnessJson {
    param($L)
    if ($null -eq $L) { return $null }
    return [PSCustomObject]@{
        integrated = [PSCustomObject]@{ i_lufs = $L.I_LUFS; threshold_lufs = $L.Int_Threshold_LUFS }
        range      = [PSCustomObject]@{
            lra_lu = $L.LRA_LU; threshold_lufs = $L.Range_Threshold_LUFS
            lra_low_lufs = $L.LRA_Low_LUFS; lra_high_lufs = $L.LRA_High_LUFS
        }
        true_peak  = [PSCustomObject]@{ dbtp = $L.TruePeak_dBFS }
    }
}

$schema = [PSCustomObject]@{
    file     = $InputFile
    duration = $fileDuration
    analysis = [ordered]@{}
}

foreach ($r in $results) {
    $role = if ($r.Label -like 'Source*') { 'source' }
            elseif ($r.Label -like 'Downmix*') { 'downmix' }
            elseif ($r.Label -eq 'AAC 5.1') { 'aac51' }
            elseif ($r.Label -eq 'DTS-HD MA 5.1') { 'dts51' }
            elseif ($r.Label -eq 'Opus 2.0') { 'opus20' }
            elseif ($r.Label -eq 'Dolby Digital Plus 2.0') { 'eac320' }
            elseif ($r.Label -eq 'AAC 2.0') { 'aac20' }
            else { "track$($r.MapIndex)" }

    $metricsObj = [PSCustomObject]@{ peak = $r.Peak; rms = $r.RMS; crest = $r.Crest; clipped = $r.ClippedSamples }
    if ($r.Codec -eq 'truehd') {
        $metricsObj | Add-Member -MemberType NoteProperty -Name preroll_ms -Value ([int][Math]::Round($r.StartTime * 1000))
    }

    $schema.analysis[$role] = [PSCustomObject]@{
        codec    = $r.Codec
        channels = $r.Channels
        metrics  = $metricsObj
        loudness = (Get-LoudnessJson $r.Loudness)
        status   = $r.Status
        log      = $r.LogFile
    }
}

if ($avSync) {
    $schema | Add-Member -MemberType NoteProperty -Name av_sync -Value ([PSCustomObject]@{
        status = $avSync.Status; offset_ms = $avSync.OffsetMs; message = $avSync.Message
    }) -Force
}
if ($loudCompare) {
    $schema | Add-Member -MemberType NoteProperty -Name loudness_comparison -Value ([PSCustomObject]@{
        reference           = $loudCompare.Reference
        target              = $loudCompare.Target
        delta_integrated_lu = $loudCompare.DeltaI
        delta_lra_lu        = $loudCompare.DeltaLRA
        tolerance_lu        = $loudCompare.Tolerance
        verdict             = $loudCompare.Verdict
    }) -Force
}

if ($JsonMode) {
    $schema | ConvertTo-Json -Depth 10
    exit 0
}

# ========================
#  Human-readable summary
# ========================
Write-Host ""
Write-Host "=== AUDIO VALIDATION SUMMARY ===" -ForegroundColor Yellow
Write-Host "File: $InputFile"
Write-Host ""

# Per-track header color, keyed by displayed track number (MapIndex).
# Track 0 -> DarkCyan, Track 1 -> Blue, then distinct non-reserved colors;
# cycles via modulo. Avoids Green/Red (status) and Yellow (section dividers).
$trackPalette = @('DarkCyan','Blue','Magenta','Cyan','DarkMagenta','DarkGreen','Gray','DarkRed')

foreach ($r in $results) {
    $color = ($r.Status -eq 'PASS') ? 'Green' : 'Red'
    $trackColor = $trackPalette[$r.MapIndex % $trackPalette.Count]
    Write-Host ("[{0}] (Track {1})" -f $r.Label, $r.MapIndex) -ForegroundColor $trackColor
    Write-Host ("  Peak:   {0} dB" -f $r.Peak)
    Write-Host ("  RMS:    {0} dB" -f $r.RMS)
    Write-Host ("  Crest:  {0} dB" -f $r.Crest)
    Write-Host ("  Clips:  {0}" -f ([int]$r.ClippedSamples))
    if ($r.Codec -eq 'truehd') {
        Write-Host ("  Preroll: {0} ms" -f ([int][Math]::Round($r.StartTime * 1000)))
    }
    Write-Host ("  Status: {0}" -f $r.Status) -ForegroundColor $color
    Write-Host ("  Log:    {0}" -f $r.LogFile)

    # Loudness (informational; never gates). Full EBU R128 block.
    $roleToken = if ($r.Label -like 'Source*') { 'SOURCE' }
                 elseif ($r.Label -like 'Downmix*') { 'DOWNMIX' }
                 else { 'TRACK' }
    $codecLabel = "{0}, {1}ch" -f (Get-CodecDisplay $r.Codec $r.Channels), $r.Channels
    Write-Host ("--- {0}  ->  0:a:{1}  ({2}) ---" -f $roleToken, $r.MapIndex, $codecLabel) -ForegroundColor $trackColor
    if ($null -ne $r.Loudness) {
        Write-Host "  Integrated loudness:"
        Write-Host ("    I:         {0}" -f (Format-Lufs $r.Loudness.I_LUFS 'LUFS'))
        Write-Host ("    Threshold: {0}" -f (Format-Lufs $r.Loudness.Int_Threshold_LUFS 'LUFS'))
        Write-Host "  Loudness range:"
        Write-Host ("    LRA:       {0}" -f (Format-Lufs $r.Loudness.LRA_LU 'LU'))
        Write-Host ("    Threshold: {0}" -f (Format-Lufs $r.Loudness.Range_Threshold_LUFS 'LUFS'))
        Write-Host ("    LRA low:   {0}" -f (Format-Lufs $r.Loudness.LRA_Low_LUFS 'LUFS'))
        Write-Host ("    LRA high:  {0}" -f (Format-Lufs $r.Loudness.LRA_High_LUFS 'LUFS'))
        Write-Host "  True peak:"
        Write-Host ("    Peak:      {0}" -f (Format-Lufs $r.Loudness.TruePeak_dBFS 'dBTP'))
    } else {
        Write-Host "  Loudness: N/A (ebur128 Summary unavailable)"
    }
    Write-Host ""

    Write-Host "  Quick Interpretation:" -ForegroundColor DarkYellow
    foreach ($line in (Format-Metrics $r)) { Write-Host "    $line" }
    Write-Host ""
}

# --- source vs downmix: peak/RMS ---
if ($srcResult -and $dmxResult) {
    Write-Host "=== SOURCE vs DOWNMIX (peak / RMS) ===" -ForegroundColor Yellow
    Write-Host ("Source Peak:   {0} dB" -f $srcResult.Peak)
    Write-Host ("Downmix Peak:  {0} dB" -f $dmxResult.Peak)
    Write-Host ("Source RMS:    {0} dB" -f $srcResult.RMS)
    Write-Host ("Downmix RMS:   {0} dB" -f $dmxResult.RMS)
    Write-Host ("Source Crest:  {0} dB" -f $srcResult.Crest)
    Write-Host ("Downmix Crest: {0} dB" -f $dmxResult.Crest)
    Write-Host ""
}

# --- source vs downmix: loudness (informational; never gates) ---
if ($loudCompare) {
    $sI = $srcResult.Loudness.I_LUFS; $dI = $dmxResult.Loudness.I_LUFS
    $sL = $srcResult.Loudness.LRA_LU; $dL = $dmxResult.Loudness.LRA_LU
    $deltaI = $loudCompare.DeltaI; $deltaLRA = $loudCompare.DeltaLRA

    Write-Host "=== SOURCE vs DOWNMIX LOUDNESS ===" -ForegroundColor Yellow
    Write-Host ("  {0,-22}{1,18}{2,18}" -f '', 'SOURCE', 'DOWNMIX')
    Write-Host ("  {0,-22}{1,18}{2,18}" -f 'Integrated (I):', (Format-Lufs $sI 'LUFS'), (Format-Lufs $dI 'LUFS'))
    Write-Host ("  {0,-22}{1,18}{2,18}" -f 'Loudness range (LRA):', (Format-Lufs $sL 'LU'), (Format-Lufs $dL 'LU'))
    Write-Host ""
    Write-Host ("  delta integrated: {0} LU" -f $(if ($null -ne $deltaI) { $deltaI.ToString('+0.0;-0.0;0.0', [CultureInfo]::InvariantCulture) } else { 'N/A' }))
    Write-Host ("  delta LRA:        {0} LU" -f $(if ($null -ne $deltaLRA) { $deltaLRA.ToString('+0.0;-0.0;0.0', [CultureInfo]::InvariantCulture) } else { 'N/A' }))
    Write-Host ("  tolerance:        {0} LU" -f $Tolerance.ToString('0.0', [CultureInfo]::InvariantCulture))
    if ($loudCompare.Verdict -eq 'preserved') {
        Write-Host "  VERDICT: PRESERVED  (|delta I| <= tolerance)" -ForegroundColor Green
    } else {
        Write-Host "  VERDICT: SHIFTED    (downmix differs from source beyond tolerance)" -ForegroundColor Yellow
    }

    Write-Host "--- Interpretation (informational) ---" -ForegroundColor White
    if ($null -ne $dI) {
        Write-Host "  Reference targets: EBU R128 -23 LUFS | ATSC A/85 -24 LUFS | streaming -14 to -16 LUFS."
        Write-Host ("  Downmix integrated loudness is {0} LUFS." -f $dI.ToString('0.0', [CultureInfo]::InvariantCulture))
    }
    if ($null -ne $deltaI) {
        if ([Math]::Abs($deltaI) -le $Tolerance) {
            Write-Host "  Integrated: within tolerance -> downmix preserved source loudness."
        } elseif ($deltaI -gt 0) {
            Write-Host ("  Integrated: downmix is LOUDER than source by {0} LU." -f $deltaI.ToString('0.0', [CultureInfo]::InvariantCulture))
        } else {
            Write-Host ("  Integrated: downmix is QUIETER than source by {0} LU." -f ([Math]::Abs($deltaI)).ToString('0.0', [CultureInfo]::InvariantCulture))
        }
    }
    if ($null -ne $deltaLRA) {
        if ($deltaLRA -gt 0) {
            Write-Host ("  Dynamics: downmix loudness range is WIDER than source by {0} LU." -f $deltaLRA.ToString('0.0', [CultureInfo]::InvariantCulture))
        } elseif ($deltaLRA -lt 0) {
            Write-Host ("  Dynamics: downmix loudness range is NARROWER than source by {0} LU." -f ([Math]::Abs($deltaLRA)).ToString('0.0', [CultureInfo]::InvariantCulture))
        } else {
            Write-Host "  Dynamics: downmix loudness range matches source."
        }
    }
    Write-Host ""
}

if ($avSync) {
    Write-Host "=== A/V SYNC CHECK ===" -ForegroundColor Yellow
    $syncColor = ($avSync.Status -eq 'OK') ? 'Green' : 'Red'
    Write-Host $avSync.Message -ForegroundColor $syncColor
    if ($avSync.Status -eq 'OK') {
        Write-Host "  -> Safe to proceed with ConvertAudioEngine-DDP51 or ConvertAudioEngine-Keep71." -ForegroundColor Green
    } else {
        Write-Host "  -> Do NOT run ConvertAudioEngine-DDP51 or ConvertAudioEngine-Keep71 until this sync issue is fixed." -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "=== End of Summary ===" -ForegroundColor Yellow
Write-Host ""
exit 0