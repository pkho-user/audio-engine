# ========================================================================
#  ConvertAudioEngine.ps1 — (Version 1.vn) Production-daily use
#  PowerShell 5.1 + 7.5 Compatible
#  FFmpeg 8.1 Compatible
#  Audio tracks over 5.1 are downmixed (1024k)
#  EAC3/TrueHD 5.1 passthrough; other 5.1 tracks re-encoded (768k)
#  Supported audio codecs: AAC, EAC3-ATMOS, TrueHD, DTS, PCM, FLAC
#  Priority Mapping (default 0-100)
#  Run: ffmpeg -h encoder=eac3 to see available options
#  5.1 audio downmix using Pan Filter for channel mapping
#
#  De-sync for TrueHD 7.1 / long-duration files:
#  (1) Increase analyzeduration/probesize to 200M for full layout detection
#      Ensures 7.1 channel layout is fully parsed before encode starts
#  (2) -avoid_negative_ts make_zero
#      Clamps any negative initial PTS from TrueHD streams to zero.
#  (3) -max_muxing_queue_size 14000
#      Large mux queue to prevent video starving audio pipeline
#      14000 provides sufficient headroom for 3+ hour files.
#  (4) aformat=channel_layouts=7.1 prepended to downmix pan filter
#      Pins the TrueHD decoder output to the canonical 7.1 layout
#      (FL FR FC LFE BL BR SL SR) before the pan filter reads it.
#      Guards against Atmos decoder output layout ambiguity.
# ========================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

# ========================================
#  ENGINE: GLOBAL SETTINGS
# ========================================
$ThreadCount = 8 # User‑adjustable (4–16).
$CommentaryKeywords = @("commentary","director","producer","writer","cast","behind","bonus","alt","interview")

$ffprobe = Join-Path $PSScriptRoot "ffprobe.exe"
$ffmpeg  = Join-Path $PSScriptRoot "ffmpeg.exe"

foreach ($bin in $ffprobe, $ffmpeg) {
    if (-not (Test-Path $bin)) {
        Write-Error "Missing required binary: $bin"
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# ========================
#  ENGINE: CHANNEL FILTER
# ========================
class ChannelFilter {
    [string] $Op
    [int]    $Value

    ChannelFilter([string]$op, [int]$val) {
        $this.Op    = $op
        $this.Value = $val
    }

    [bool] Matches([int]$channels) {
        switch ($this.Op) {
            "gt" { return $channels -gt $this.Value }
            "lt" { return $channels -lt $this.Value }
            "ge" { return $channels -ge $this.Value }
            "le" { return $channels -le $this.Value }
            "eq" { return $channels -eq $this.Value }
            default { throw "Unknown ChannelFilter op: $($this.Op)" }
        }
        return $false  # Unreachable — satisfies PS 5.1 class method return-path analysis
    }

    static [ChannelFilter] $MoreThanTwo = [ChannelFilter]::new("gt", 2)
}

# ============================
#  ENGINE: AUDIO RULE GROUPS
# ============================

# AAC FAMILY
$Rules_AAC = @(
    # Rule1: AAC 7.1 → Downmix to 5.1 at 1024k
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = "1024k"
        PassthroughTag = $null
        Rule           = "AAC_7.1_Downmix_1024k"
        Priority       = 92
    },
    # Rule2: AAC 5.1 → Encode at 768k
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "768k"
        PassthroughTag = $null
        Rule           = "AAC_5.1_Encode_768k"
        Priority       = 91
    },
    # Rule3: AAC 2.0 → Encode at 256k
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "256k"
        PassthroughTag = $null
        Rule           = "AAC_2.0_Encode_256k"
        Priority       = 90
    }
)

# EAC3 / ATMOS FAMILY
$Rules_EAC3_Atmos = @(
    # Rule1: EAC3 Atmos (JOC) → Passthrough
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = [ChannelFilter]::new("gt", 2)
        ProfileRegex   = "JOC|Atmos"
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "EAC3_Atmos_Passthrough"
        Rule           = "EAC3_Atmos_Passthrough"
        Priority       = 100
    },
    # Rule2: EAC3 7.1 → Downmix to 5.1 at 1024k
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = "1024k"
        PassthroughTag = $null
        Rule           = "EAC3_7.1_Downmix_1024k"
        Priority       = 99
    },
    # Rule3: EAC3 5.1 → Passthrough
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "EAC3_5.1_Passthrough"
        Rule           = "EAC3_5.1_Passthrough"
        Priority       = 98
    },
    # Rule4: EAC3 2.0 → Passthrough
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "EAC3_2.0_Passthrough"
        Rule           = "EAC3_2.0_Passthrough"
        Priority       = 97
    }
)

# TRUEHD FAMILY
$Rules_TrueHD = @(
    # Rule1: TrueHD 5.1 → Passthrough
    [PSCustomObject]@{
        CodecRegex     = "^(mlp|truehd|true-hd)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "TrueHD_5.1_Passthrough"
        Rule           = "TrueHD_5.1_Passthrough"
        Priority       = 81
    },
    # Rule2: TrueHD 7.1 → Downmix to 5.1 at 1024k
    [PSCustomObject]@{
        CodecRegex     = "^(mlp|truehd|true-hd)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = "1024k"
        PassthroughTag = $null
        Rule           = "TrueHD_7.1_Downmix_1024k"
        Priority       = 80
    }
)

# DTS FAMILY
$Rules_DTS = @(
    # Rule1: DTS-HD Multichannel → Encode at 1024k
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = [ChannelFilter]::MoreThanTwo
        ProfileRegex   = "HD|MA|HRA"
        Action         = "Encode"
        Bitrate        = "1024k"
        PassthroughTag = $null
        Rule           = "DTSHD_Multichannel_Encode_1024k"
        Priority       = 73
    },
    # Rule2: DTS Multichannel → Encode at 768k
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = [ChannelFilter]::MoreThanTwo
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "768k"
        PassthroughTag = $null
        Rule           = "DTS_Multichannel_Encode_768k"
        Priority       = 72
    },
    # Rule3: DTS-HD 2.0 → Encode at 384k
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = 2
        ProfileRegex   = "HD|MA|HRA"
        Action         = "Encode"
        Bitrate        = "384k"
        PassthroughTag = $null
        Rule           = "DTSHD_2.0_Encode_384k"
        Priority       = 71
    },
    # Rule4: DTS 2.0 → Encode at 256k
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "256k"
        PassthroughTag = $null
        Rule           = "DTS_2.0_Encode_256k"
        Priority       = 70
    }
)

# PCM / FLAC FAMILY
$Rules_PCMFLAC = @(
    # Rule1: PCM/FLAC 7.1 → Downmix to 5.1 at 1024k
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = "1024k"
        PassthroughTag = $null
        Rule           = "PCMFLAC_7.1_Downmix_1024k"
        Priority       = 62
    },
    # Rule2: PCM/FLAC 5.1 → Encode at 768k
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "768k"
        PassthroughTag = $null
        Rule           = "PCMFLAC_5.1_Encode_768k"
        Priority       = 61
    },
    # Rule3: PCM/FLAC 2.0 → Encode at 384k
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = "384k"
        PassthroughTag = $null
        Rule           = "PCMFLAC_2.0_Encode_384k"
        Priority       = 60
    }
)

# ============================================================
#  ENGINE: MERGE ALL AUDIO RULE GROUPS
#  (Adding a new codec = define $Rules_XYZ, then append here)
# ============================================================
$AudioRules = $Rules_EAC3_Atmos + $Rules_TrueHD + $Rules_DTS + $Rules_AAC + $Rules_PCMFLAC

# Priority sort
# Replaced ForEach-Object, identical result
$AudioRules = $AudioRules | Sort-Object { $_.Priority } -Descending

# Precompile regex = correct.
$AudioRules = $AudioRules | Select-Object -Property *,
    @{ Name='CodecRegexObj';   Expression={
        if ($_.CodecRegex)   { [regex]::new($_.CodecRegex,   'IgnoreCase') } else { $null }
    }},
    @{ Name='ProfileRegexObj'; Expression={
        if ($_.ProfileRegex) { [regex]::new($_.ProfileRegex, 'IgnoreCase') } else { $null }
    }}

# ==========================
#  PROBE
# ==========================
function Get-AudioStreams {
    param([string]$File)

    $probeArgs = @(
        "-analyzeduration","200M", # match FFmpeg probe
        "-probesize","200M",       # match FFmpeg probe
        "-v","quiet","-print_format","json",
        "-show_streams","-select_streams","a",$File
    )

    $raw = & $script:ffprobe @probeArgs 2>$null
    if (-not $raw) {
        throw "ffprobe returned no output. Check the input file: $File"
    }
    try { return ($raw | ConvertFrom-Json).streams }
    catch {
        throw "Failed to parse ffprobe JSON: $_"
    }
}

# ===========================
#  COMMENTARY DETECTOR
# ===========================
function Test-IsCommentary {
    param($Channels, $Title)

    if ($Channels -ne 2 -or -not $Title) { return $false }

    $lower = $Title.ToLower()
    foreach ($kw in $script:CommentaryKeywords) {
        if ($lower -match $kw) { return $true }
    }
    return $false
}

# ===========================
#  RULE MATCHER
# ===========================
function Resolve-AudioRule {
    param($Codec, $Channels, $Profile, $Title, $Handler)

    foreach ($r in $script:AudioRules) {

        if ($r.CodecRegexObj -and -not $r.CodecRegexObj.IsMatch($Codec)) { continue }

        if ($null -ne $r.Channels) {
            if ($r.Channels -is [ChannelFilter]) {
                if (-not $r.Channels.Matches($Channels)) { continue }
            }
            elseif ($Channels -ne $r.Channels) {
                continue
            }
        }

        if ($r.ProfileRegexObj) {
            $match = $false
            if ($Profile -and $r.ProfileRegexObj.IsMatch($Profile)) { $match = $true }
            if (-not $match -and $Title   -and $r.ProfileRegexObj.IsMatch($Title))   { $match = $true }
            if (-not $match -and $Handler -and $r.ProfileRegexObj.IsMatch($Handler)) { $match = $true }
            if (-not $match) { continue }
        }

        return $r
    }

    return $null
}

# =============================
#  PROCESS TRACKS
# =============================
function Convert-AudioTracks {
    param($Streams)

    $Processed = [System.Collections.Generic.List[object]]::new()
    $TrackIndex = 0

    foreach ($s in $Streams) {

        $Codec    = $s.codec_name
        $Channels = [int]$s.channels
        $Profile  = $s.profile

        $Title    = if ($s.tags) { $s.tags.title } else { $null }
        $Lang     = if ($s.tags -and $s.tags.language) { $s.tags.language } else { "eng" }
        $Handler  = if ($s.tags) { $s.tags.handler_name } else { $null }

        $RealIndex = $s.index
        $Rule = $null

        # --- Malformed Layout Guards (include 7ch) ---
        if ($Codec -eq "aac" -and $Channels -eq 7) {
            $Channels = 6
            $Rule = "AAC_MalformedLayout_Guard"
        }

        if ($Codec -eq "eac3" -and $Channels -eq 7) {
            $Channels = 6
            $Rule = "EAC3_MalformedLayout_Guard"
        }

        if ($Codec -match "^(mlp|truehd|true-hd)$" -and $Channels -eq 7) {
            $Channels = 8
            $Rule = "TrueHD_MalformedLayout_Guard"
        }

        if ($Codec -match "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$" -and $Channels -eq 7) {
            $Channels = 8
            $Rule = "PCMFLAC_MalformedLayout_Guard"
        }

        # --- Commentary removal ---
        if (Test-IsCommentary -Channels $Channels -Title $Title) {
            $Processed.Add([PSCustomObject]@{
                Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
                Profile=$Profile; Title=$Title; Language=$Lang
                Action="Removed"; Passthrough=$false; Downmix=$false
                Bitrate=$null; PassthroughTag=$null; Output="Commentary"
                Rule="Commentary_Removed"; Priority=0
            })
            $TrackIndex++; continue
        }

        # Rule matching
        $match = Resolve-AudioRule -Codec $Codec -Channels $Channels -Profile $Profile -Title $Title -Handler $Handler

        if ($match) {
            $Action      = $match.Action
            $Bitrate     = $match.Bitrate
            $Passthrough = ($Action -eq "Passthrough")
            $Downmix     = ($Action -eq "Downmix")
            $Tag         = $match.PassthroughTag
            $Priority    = $match.Priority
            if (-not $Rule) { $Rule = $match.Rule }
        }
        else {
            if ($Channels -le 2)      { $Action="Encode"; $Bitrate="256k";  $Rule="Fallback_2.0_256k" }
            elseif ($Channels -gt 6)  { $Action="Downmix"; $Bitrate="1024k"; $Rule="Fallback_7.1_1024k" }
            else                      { $Action="Encode"; $Bitrate="768k";  $Rule="Fallback_5.1_768k" }
            $Passthrough=$false; $Downmix=($Action -eq "Downmix"); $Tag=$null; $Priority=0
        }

        # --- Safety Audit ---
        if ($Channels -eq 6 -and $Bitrate -eq "1024k") {
            Write-Warning ("Track {0} ({1}): 6ch rule had 1024k-clamped to 768k." -f $TrackIndex,$Codec)
            $Bitrate = "768k"
        }
        elseif ($Channels -eq 8 -and $Bitrate -eq "768k") {
            Write-Warning ("Track {0} ({1}): 8ch rule had 768k-clamped to 1024k." -f $TrackIndex,$Codec)
            $Bitrate = "1024k"
        }

        $Processed.Add([PSCustomObject]@{
            Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
            Profile=$Profile; Title=$Title; Language=$Lang; Action=$Action
            Passthrough=$Passthrough; Downmix=$Downmix; Bitrate=$Bitrate
            PassthroughTag=$Tag; Output=""; Rule=$Rule; Priority=$Priority
        })

        $TrackIndex++
    }

    return $Processed
}

# =============================
#  FFMPEG COMMAND BUILDER
# =============================
function Build-FFmpegCommand {
    param($Tracks, $InputFile, $ThreadCount)

    $ffArgs = New-Object System.Collections.Generic.List[string]

    # ------------------------------------------------------------------
    #  Global options
    #  NOTE — order matters for FFmpeg: all input options must appear
    #  before -i. Output options (-avoid_negative_ts, -max_muxing_queue_size)
    #  must appear after -i.
    # ------------------------------------------------------------------
    $ffArgs.AddRange([string[]](
        "-threads",             $ThreadCount,
        "-analyzeduration",     "200M",         # ensures TrueHD channel layout fully parsed
        "-probesize",           "200M",         # raised from 100M
        "-err_detect",          "ignore_err",
        "-drc_scale",           "0",
        "-i",                   $InputFile,
        "-avoid_negative_ts",   "make_zero",    # clamps negative TrueHD PTS to zero
        "-max_muxing_queue_size","14000",       # stops video flooding mux queue when audio is slow
        "-map",                 "0:v?",
        "-c:v",                 "copy",
        "-map_metadata",        "0",
        "-map_chapters",        "0"
    ))

    $i = 0
    foreach ($t in $Tracks) {
        if ($t.Action -eq "Removed") { continue }

        $ffArgs.AddRange([string[]]("-map","0:$($t.RealIndex)"))
        $LangTag = " [$($t.Language)]"

        if ($t.Passthrough) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i","copy",
                "-metadata:s:a:$i","title=$($t.PassthroughTag)$LangTag"
            ))
            $t.Output = "$($t.PassthroughTag)$LangTag"
        }
        elseif ($t.Downmix) {
            # --- 7.1 → 5.1 DOWNMIX PATH ---
            # (4) see header note
            # Downmixes 7.1 audio to EAC3 using ITU-R BS.775 matrix
            # Side (SL/SR) are folded into the rear (BL/BR) with -3 dB attenuation
            # Dolby DRC is disabled. Final output is encoded as DD+ 5.1.
            $panFilter = "aformat=channel_layouts=7.1,pan=5.1|FL=FL|FR=FR|FC=FC|LFE=LFE|BL=BL+0.707*SL|BR=BR+0.707*SR"

            $ffArgs.AddRange([string[]](
                "-filter:a:$i", $panFilter,
                "-c:a:$i",      "eac3",
                "-b:a:$i",      $t.Bitrate,
                "-dialnorm",    "-31",
                "-cutoff",      "20000",
                "-metadata:s:a:$i", "title=DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
        }
        elseif ($t.Channels -le 2) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
                "-ac","2",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-dsur_mode","1",           # Dolby Surround Mode, for EAC3 encoder.
                "-stereo_rematrixing","1",  # Explicitly enable rematrixing
                "-metadata:s:a:$i","title=DD+ 2.0 ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 2.0 ($($t.Bitrate))$LangTag"
        }
        else {
            # --- 5.1 RE-ENCODE PATH ---
            # Re-encodes any 5.1 audio track to EAC3 (DD+ 5.1).
            # No downmixing occurs here — input must already be 5.1.
            # Dolby DRC is disabled. Loudness signaling handled by -dialnorm -31.
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
                "-ac","6",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-cutoff","20000",
                "-metadata:s:a:$i","title=DD+ 5.1 ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 ($($t.Bitrate))$LangTag"
        }

        $ffArgs.AddRange([string[]]("-metadata:s:a:$i","language=$($t.Language)"))

        $disp = if ($i -eq 0) { "default" } else { "0" }
        $ffArgs.AddRange([string[]]("-disposition:a:$i",$disp))

        $i++
    }

    $ffArgs.AddRange([string[]]("-map","0:s?","-c:s","copy"))
    $outDir  = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($InputFile))
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile) + "_Processed.mkv"
    $ffArgs.Add([System.IO.Path]::Combine($outDir, $outName))

    return $ffArgs
}

# =======================================
#  MAIN EXECUTION
# =======================================
Write-Host "=== Probing Audio Streams ===" -ForegroundColor Cyan
$streams = Get-AudioStreams -File $InputFile

Write-Host "=== Processing Tracks ===" -ForegroundColor Cyan
$tracks = Convert-AudioTracks -Streams $streams

Write-Host "=== Building FFmpeg Command ===" -ForegroundColor Cyan
$cmd = Build-FFmpegCommand -Tracks $tracks -InputFile $InputFile -ThreadCount $ThreadCount

& $ffmpeg @cmd
if ($LASTEXITCODE -ne 0) {
    Write-Error "FFmpeg exited with code $LASTEXITCODE. Output may be incomplete."
    exit $LASTEXITCODE
}

# ==========================================
#  SUMMARY ENGINE
# ==========================================

Write-Host ""
Write-Host "=== AUDIO PROCESSING SUMMARY ===" -ForegroundColor Cyan
Write-Host ""

# Header
$header = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
    "Idx","Codec","Ch","Action","Output","Pri","Rule"

Write-Host $header -ForegroundColor White
Write-Host ("-" * $header.Length) -ForegroundColor DarkGray

foreach ($t in $tracks) {

    # Color selection
    $Color = switch ($t.Action) {
        "Passthrough" { "Green" }      # Keep green for passthrough (good)
        "Downmix"     { "Yellow" }     # Keep yellow for downmix (warning)
        "Encode"      { "Blue" }       # Blue for encoding (main action)
        "Removed"     { "Red" }        # Red for removed (important)
        default       { "White" }
    }

    # Build action label
    $ActionLabel = switch ($t.Action) {
        "Passthrough" { "Passthrough" }
        "Downmix"     { "Downmix 5.1" }
        "Encode"      {
            if     ($t.Bitrate -eq "256k")  { "Encode 256k" }
            elseif ($t.Bitrate -eq "384k")  { "Encode 384k" }
            elseif ($t.Bitrate -eq "768k")  { "Encode 768k" }
            elseif ($t.Bitrate -eq "1024k") { "Encode 1024k" }
            else                            { "Encode" }
        }
        "Removed"     { "Removed" }
        default       { $t.Action }
    }

    # Build output label
    $OutputLabel = if ($t.Output) { $t.Output } else { "(none)" }

    # Show "-" for removed tracks, otherwise priority.
    $PriLabel = if ($t.Action -eq "Removed") { "-" } else { $t.Priority }

    # Final formatted line
    $line = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
        $t.Index, $t.Codec, $t.Channels, $ActionLabel, $OutputLabel, $PriLabel, $t.Rule

    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
Write-Host "=== END OF SUMMARY ===" -ForegroundColor Cyan
Write-Host ""
