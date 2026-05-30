# ======================================================================
#  ConvertAudio2-Stereo (Version 1.4.0) — Production daily use
#  PowerShell 7.6 | FFmpeg 8.1 | Opus 1.6.1 | EAC3
#
#  Foundation: ConvertAudioEngine-DDP51 v3.0.5 (3-phase A/B/C SPN + unified
#  Stereo 2.0 only output:
#  EAC3 (default, broad device compatibility)
#  Opus (better quality-per-bit).
# ======================================================================
#Requires -Version 7.6

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputFile,

    [ValidateSet("EAC3","Opus")]
    [string]$StereoCodec
)

# ===================
#  GLOBAL SETTINGS
# ===================

# Default stereo codec used when -StereoCodec is not provided.  
# You can change this to "Opus" or "EAC3"
$DefaultStereoCodec = "EAC3"

# Param block binds before script body, so default is applied via $PSBoundParameters.
if (-not $PSBoundParameters.ContainsKey('StereoCodec')) {
    $StereoCodec = $DefaultStereoCodec
}

# Bitrates used when the script encodes stereo audio; higher values mean higher quality.  
# You can safely adjust these numbers, within "...k"
#   EAC3: 224,256,320,384,448,512   (256 = transparent for most stereo content)
#   Opus: 192,224,256,320,352,384   (256 = fully transparent for all content)
$StereoBitrates = @{
    "EAC3" = "384k"
    "Opus" = "320k"
}

# $StereoBitrates (EAC3/Opus) controls all bitrate values used throughout the script:
# the summary display, dynamic Rule strings, conversion rules, and FFmpeg encoding commands.
# Since the output is always stereo, every key in $StereoBitRateConfig points to the same bitrate,
# When you update the codec or bitrate setting above, everything else updates automatically.
$StereoBitRateConfig = @{
    Downmix71 = $StereoBitrates[$StereoCodec]   # 7.1 / Atmos source → stereo
    Encode51  = $StereoBitrates[$StereoCodec]   # 5.1 / multichannel source → stereo
    Encode20  = $StereoBitrates[$StereoCodec]   # 2.0 source → stereo (re-encode)
}

# LFE fold-in: $true folds LFE into FL+FR at +0.5 gain inside the pan matrix.
# $false discards LFE. [bool] enforces boolean semantics — strings would coerce.
[bool]$FoldLFE = $true

# Threading and SPN settings.
$ThreadCount     = 8    # User-adjustable (4-16).
if ($ThreadCount -lt 4 -or $ThreadCount -gt 16) { throw "ThreadCount must be between 4 and 16." }
# SPN scan concurrency. SSD/NVMe: safe up to 4. HDDs: keep at 1.
$ScanThrottle    = [Math]::Max(1, [int]([Environment]::ProcessorCount / $ThreadCount))
$PeakThresholdDB = -0.5 # SPN: loudnorm triggers when source peak exceeds this (dBFS)
$CommentaryPattern = [regex]::new(
    'commentary|director|producer|writer|cast|behind|bonus|alt|interview',
    'IgnoreCase,Compiled'
)

# Platform-aware binary resolution.
if ($IsWindows) {
    $ffmpeg  = Join-Path $PSScriptRoot "ffmpeg.exe"
    $ffprobe = Join-Path $PSScriptRoot "ffprobe.exe"
}
else {
    $ffmpeg  = Join-Path $PSScriptRoot "ffmpeg"
    $ffprobe = Join-Path $PSScriptRoot "ffprobe"
}

foreach ($bin in $ffprobe, $ffmpeg) {
    if (-not (Test-Path -LiteralPath $bin)) {
        Write-Error "Missing required binary: $bin"
        exit 1
    }
    if (-not $IsWindows -and -not ((Get-Item -LiteralPath $bin).UnixFileMode -band [System.IO.UnixFileMode]::UserExecute)) {
        Write-Error "Binary not executable: $bin — run: chmod +x `"$bin`""
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
        return $false  # Unreachable — required by PS class method return-path analysis
    }

    static [ChannelFilter] $MoreThanTwo = [ChannelFilter]::new("gt", 2)
}

# ==========================
#  ENGINE: AUDIO RULE GROUPS
# ==========================
# Rule engine assigns Passthrough/Encode/Downmix labels. Build-FFmpegCommand makes
# the actual stereo-output decision via channel-count routing.

# AAC FAMILY
$Rules_AAC = @(
    [PSCustomObject]@{
        CodecRegex = "^(aac)$"
        Channels = 8
        ProfileRegex = $null
        Action = "Downmix"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "AAC_7.1_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 92
    },
    [PSCustomObject]@{
        CodecRegex = "^(aac)$"
        Channels = 6
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode51
        PassthroughTag = $null
        Rule = "AAC_5.1_to_Stereo_$($StereoBitRateConfig.Encode51)"
        Priority = 91
    },
    [PSCustomObject]@{
        CodecRegex = "^(aac)$"
        Channels = 2
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode20
        PassthroughTag = $null
        Rule = "AAC_2.0_to_Stereo_$($StereoBitRateConfig.Encode20)"
        Priority = 90
    }
)

# EAC3 / ATMOS FAMILY
$Rules_EAC3_Atmos = @(
    # Atmos (JOC) — downmixed to stereo. Build-FFmpegCommand selects the 5.1 or 7.1 pan matrix.
    [PSCustomObject]@{
        CodecRegex = "^(eac3)$"
        Channels = [ChannelFilter]::new("gt", 2)
        ProfileRegex = "JOC|Atmos"
        Action = "Downmix"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "EAC3_Atmos_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 100
    },
    [PSCustomObject]@{
        CodecRegex = "^(eac3)$"
        Channels = 8
        ProfileRegex = $null
        Action = "Downmix"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "EAC3_7.1_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 99
    },
    [PSCustomObject]@{
        CodecRegex = "^(eac3)$"
        Channels = 6
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode51
        PassthroughTag = $null
        Rule = "EAC3_5.1_to_Stereo_$($StereoBitRateConfig.Encode51)"
        Priority = 98
    },
    # EAC3 2.0 — Conditional passthrough.(StereoCodec=EAC3 required).
    [PSCustomObject]@{
        CodecRegex = "^(eac3)$"
        Channels = 2
        ProfileRegex = $null
        Action = "Passthrough"
        Bitrate = $null
        PassthroughTag = "EAC3_2.0_Passthrough"
        Rule = "EAC3_2.0_Conditional_Passthrough"
        Priority = 97
    }
)

# TRUEHD FAMILY
$Rules_TrueHD = @(
    [PSCustomObject]@{
        CodecRegex = "^(mlp|truehd|true-hd)$"
        Channels = 6
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode51
        PassthroughTag = $null
        Rule = "TrueHD_5.1_to_Stereo_$($StereoBitRateConfig.Encode51)"
        Priority = 81
    },
    [PSCustomObject]@{
        CodecRegex = "^(mlp|truehd|true-hd)$"
        Channels = 8
        ProfileRegex = $null
        Action = "Downmix"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "TrueHD_7.1_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 80
    }
)

# DTS FAMILY
$Rules_DTS = @(
    [PSCustomObject]@{
        CodecRegex = "^(dts)$"
        Channels = [ChannelFilter]::MoreThanTwo
        ProfileRegex = "HD|MA|HRA"
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "DTSHD_Multichannel_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 73
    },
    [PSCustomObject]@{
        CodecRegex = "^(dts)$"
        Channels = [ChannelFilter]::MoreThanTwo
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode51
        PassthroughTag = $null
        Rule = "DTS_Multichannel_to_Stereo_$($StereoBitRateConfig.Encode51)"
        Priority = 72
    },
    [PSCustomObject]@{
        CodecRegex = "^(dts)$"
        Channels = 2
        ProfileRegex = "HD|MA|HRA"
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode20
        PassthroughTag = $null
        Rule = "DTSHD_2.0_to_Stereo_$($StereoBitRateConfig.Encode20)"
        Priority = 71
    },
    [PSCustomObject]@{
        CodecRegex = "^(dts)$"
        Channels = 2
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode20
        PassthroughTag = $null
        Rule = "DTS_2.0_to_Stereo_$($StereoBitRateConfig.Encode20)"
        Priority = 70
    }
)

# PCM / FLAC FAMILY
$Rules_PCMFLAC = @(
    [PSCustomObject]@{
        CodecRegex = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels = 8
        ProfileRegex = $null
        Action = "Downmix"
        Bitrate = $StereoBitRateConfig.Downmix71
        PassthroughTag = $null
        Rule = "PCMFLAC_7.1_to_Stereo_$($StereoBitRateConfig.Downmix71)"
        Priority = 62
    },
    [PSCustomObject]@{
        CodecRegex = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels = 6
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode51
        PassthroughTag = $null
        Rule = "PCMFLAC_5.1_to_Stereo_$($StereoBitRateConfig.Encode51)"
        Priority = 61
    },
    [PSCustomObject]@{
        CodecRegex = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels = 2
        ProfileRegex = $null
        Action = "Encode"
        Bitrate = $StereoBitRateConfig.Encode20
        PassthroughTag = $null
        Rule = "PCMFLAC_2.0_to_Stereo_$($StereoBitRateConfig.Encode20)"
        Priority = 60
    }
)

# ====================================
#  ENGINE: MERGE ALL AUDIO RULE GROUPS
# ====================================
$AudioRules = $Rules_EAC3_Atmos + $Rules_TrueHD + $Rules_DTS + $Rules_AAC + $Rules_PCMFLAC

# Priority sort
$AudioRules = $AudioRules | Sort-Object { $_.Priority } -Descending

# Precompile regex
$AudioRules = $AudioRules | Select-Object -Property *,
    @{ Name='CodecRegexObj';   Expression={
        $_.CodecRegex   ? [regex]::new($_.CodecRegex,   'IgnoreCase,Compiled') : $null
    }},
    @{ Name='ProfileRegexObj'; Expression={
        $_.ProfileRegex ? [regex]::new($_.ProfileRegex, 'IgnoreCase,Compiled') : $null
    }}

# ================================================
#  ENGINE: TRACK QUALITY RANK TABLE (PreTrack-QRS)
# ================================================
# Multichannel-only quality ladder. Independent of $AudioRules priority. First match wins.
# Channels matched against post-malformed-layout-guard count. Stereo codecs absent by design.
# Lossless ranks above lossy at equal channel count; TrueHD outranks PCM/FLAC (spatial metadata).
$TrackQualityRanks = @(
    [PSCustomObject]@{
        CodecRegex   = "^(mlp|truehd|true-hd)$"
        Channels     = 8
        ProfileRegex = "Atmos"
        QualityRank  = 100
        RankLabel    = "TrueHD Atmos"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(mlp|truehd|true-hd)$"
        Channels     = 8
        ProfileRegex = $null
        QualityRank  = 99
        RankLabel    = "TrueHD 7.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(mlp|truehd|true-hd)$"
        Channels     = 6
        ProfileRegex = $null
        QualityRank  = 98
        RankLabel    = "TrueHD 5.1"
    },

    [PSCustomObject]@{
        CodecRegex   = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels     = 8
        ProfileRegex = $null
        QualityRank  = 95
        RankLabel    = "PCM/FLAC 7.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels     = 6
        ProfileRegex = $null
        QualityRank  = 94
        RankLabel    = "PCM/FLAC 5.1"
    },

    [PSCustomObject]@{
        CodecRegex   = "^(dts)$"
        Channels     = 8
        ProfileRegex = "HD|MA"
        QualityRank  = 90
        RankLabel    = "DTS-HD MA 7.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(dts)$"
        Channels     = 6
        ProfileRegex = "HD|MA"
        QualityRank  = 89
        RankLabel    = "DTS-HD MA 5.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(dts)$"
        Channels     = [ChannelFilter]::MoreThanTwo
        ProfileRegex = "HRA"
        QualityRank  = 85
        RankLabel    = "DTS-HRA"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(dts)$"
        Channels     = [ChannelFilter]::MoreThanTwo
        ProfileRegex = $null
        QualityRank  = 80
        RankLabel    = "DTS Core"
    },

    [PSCustomObject]@{
        CodecRegex   = "^(eac3)$"
        Channels     = [ChannelFilter]::MoreThanTwo
        ProfileRegex = "JOC|Atmos"
        QualityRank  = 75
        RankLabel    = "EAC3 Atmos"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(eac3)$"
        Channels     = 8
        ProfileRegex = $null
        QualityRank  = 70
        RankLabel    = "EAC3 7.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(eac3)$"
        Channels     = 6
        ProfileRegex = $null
        QualityRank  = 69
        RankLabel    = "EAC3 5.1"
    },

    [PSCustomObject]@{
        CodecRegex   = "^(ac3)$"
        Channels     = 6
        ProfileRegex = $null
        QualityRank  = 60
        RankLabel    = "AC3 5.1"
    },

    [PSCustomObject]@{
        CodecRegex   = "^(aac)$"
        Channels     = 8
        ProfileRegex = $null
        QualityRank  = 50
        RankLabel    = "AAC 7.1"
    },
    [PSCustomObject]@{
        CodecRegex   = "^(aac)$"
        Channels     = 6
        ProfileRegex = $null
        QualityRank  = 49
        RankLabel    = "AAC 5.1"
    }
)

# Precompile regex objects for ranking table (mirrors AudioRules pattern)
$TrackQualityRanks = $TrackQualityRanks | Select-Object -Property *,
    @{ Name='CodecRegexObj';   Expression={
        $_.CodecRegex   ? [regex]::new($_.CodecRegex,   'IgnoreCase,Compiled') : $null
    }},
    @{ Name='ProfileRegexObj'; Expression={
        $_.ProfileRegex ? [regex]::new($_.ProfileRegex, 'IgnoreCase,Compiled') : $null
    }}

# ==========
#  PROBE
# ==========
function Get-AudioStreams {
    param([string]$File)

    $probeArgs = @(
        "-analyzeduration","150M",
        "-probesize","150M",
        "-v","quiet","-print_format","json",
        "-show_streams","-select_streams","a",$File
    )

    $raw = & $script:ffprobe @probeArgs 2>$null
    if (-not $raw) {
        throw "ffprobe returned no output. Check the input file: $File"
    }
    try {
        return ($raw | ConvertFrom-Json).streams
    }
    catch {
        throw "Failed to parse ffprobe JSON: $($_.Exception.Message)"
    }
}

# COMMENTARY DETECTOR
function Test-IsCommentary {
    param(
        [int]    $Channels,
        [string] $Title
    )

    if ($Channels -ne 2 -or -not $Title) { return $false }
    return $script:CommentaryPattern.IsMatch($Title)
}

# ================
#  RULE MATCHER
# ================
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
            if (-not (
                ($Profile -and $r.ProfileRegexObj.IsMatch($Profile)) -or
                ($Title   -and $r.ProfileRegexObj.IsMatch($Title))   -or
                ($Handler -and $r.ProfileRegexObj.IsMatch($Handler))
            )) { continue }
        }

        return $r
    }

    return $null
}

# ==========================
#  TRACK QUALITY RESOLVER
# ==========================
# Mirrors Resolve-AudioRule; returns QualityRank/RankLabel instead of action metadata.
function Resolve-TrackQuality {
    param($Codec, $Channels, $Profile, $Title, $Handler)

    foreach ($r in $script:TrackQualityRanks) {

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
            if (-not (
                ($Profile -and $r.ProfileRegexObj.IsMatch($Profile)) -or
                ($Title   -and $r.ProfileRegexObj.IsMatch($Title))   -or
                ($Handler -and $r.ProfileRegexObj.IsMatch($Handler))
            )) { continue }
        }

        return $r
    }

    return $null
}

# ==============================
#  LANGUAGE PRIORITY (tie-break)
# ==============================
# PreTrack-QRS tie-break: eng > und > others. Aligns with DDP51/KEEP71.
function Get-LanguagePriority {
    param([string]$Lang)
    switch (([string]$Lang).ToLower()) {
        'eng'   { return 100 }
        'und'   { return  50 }
        default { return   0 }
    }
}

# --- Malformed-Layout Guard Helpers ---
function New-RemovedTrack {
    param(
        $TrackIndex, $RealIndex, $Codec, $Channels,
        $Profile, $Title, $Handler, $Lang, $SourceBitrate,
        $Rule
    )

    [PSCustomObject]@{
        Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
        Profile=$Profile; Title=$Title; Handler=$Handler; Language=$Lang
        SourceBitrate=$SourceBitrate
        Action="Removed"; Passthrough=$false
        PassthroughTag=$null; Output="Malformed"
        Rule=$Rule; Priority=0
        NeedsNormalization=$false; NeedsSpnScan=$false
        QualityRank=$null; RankLabel=$null
    }
}

function Apply-MalformedLayoutGuards {
    param(
        [Parameter(Mandatory)][string]$Codec,
        [Parameter(Mandatory)][int]   $Channels,
        [Parameter(Mandatory)][int]   $RealIndex
    )

    $rule = $null

    switch -Regex ($Codec) {

        '^aac$' {
            if ($Channels -eq 7) { $Channels = 6; $rule = "AAC_MalformedLayout_Guard" }
        }

        '^eac3$' {
            if ($Channels -eq 7) { $Channels = 6; $rule = "EAC3_MalformedLayout_Guard" }
        }

        '^(mlp|truehd|true-hd)$' {
            if ($Channels -eq 7) { $Channels = 8; $rule = "TrueHD_MalformedLayout_Guard" }
        }

        '^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$' {
            if ($Channels -eq 7) { $Channels = 8; $rule = "PCMFLAC_MalformedLayout_Guard" }
        }

        '^ac3$' {
            if ($Channels -eq 7) {
                Write-Warning "Malformed AC3 7ch on track $RealIndex — removed."
                return [pscustomobject]@{
                    Channels  = $Channels
                    Rule      = 'AC3_Malformed_Removed'
                    SkipTrack = $true
                }
            }
        }
    }

    [pscustomobject]@{
        Channels  = $Channels
        Rule      = $rule
        SkipTrack = $false
    }
}

# ==================
#  PROCESS TRACKS
# ==================
function Convert-AudioTracks {
    param($Streams)

    $Processed  = [System.Collections.Generic.List[object]]::new()
    $TrackIndex = 0

    # --- Phase A ---
    $PendingTracks = [System.Collections.Generic.List[object]]::new()

    foreach ($s in $Streams) {

        $Codec    = $s.codec_name
        $Channels = [int]$s.channels
        $Profile  = $s.profile

        $Title    = $s.tags?.title
        $Lang     = if ($s.tags?.language) { $s.tags.language } else { "eng" }
        $Handler  = $s.tags?.handler_name

        # Source bitrate (bps int). PreTrack-QRS tie-break only; 0 if missing/N-A.
        $SourceBitrate = ($s.bit_rate -as [int]) ?? 0

        $RealIndex = $s.index
        $Rule = $null

        # --- Centralized Malformed-Layout Guard Module ---
        $guard = Apply-MalformedLayoutGuards -Codec $Codec -Channels $Channels -RealIndex $RealIndex
        $Channels = $guard.Channels
        if ($guard.Rule) { $Rule = $guard.Rule }

        if ($guard.SkipTrack) {
            $removed = New-RemovedTrack $TrackIndex $RealIndex $Codec $Channels $Profile $Title $Handler $Lang $SourceBitrate $Rule
            $PendingTracks.Add($removed)
            $TrackIndex++; continue
        }

        # --- Commentary removal ---
        if (Test-IsCommentary -Channels $Channels -Title $Title) {
            $PendingTracks.Add([PSCustomObject]@{
                Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
                Profile=$Profile; Title=$Title; Handler=$Handler; Language=$Lang
                SourceBitrate=$SourceBitrate
                Action="Removed"; Passthrough=$false
                PassthroughTag=$null; Output="Commentary"
                Rule="Commentary_Removed"; Priority=0
                NeedsNormalization=$false; NeedsSpnScan=$false
                QualityRank=$null; RankLabel=$null
            })
            $TrackIndex++; continue
        }

        $match = Resolve-AudioRule -Codec $Codec -Channels $Channels -Profile $Profile -Title $Title -Handler $Handler

        if ($match) {
            $Action      = $match.Action
            $Passthrough = ($Action -eq "Passthrough")
            $Tag         = $match.PassthroughTag
            $Priority    = $match.Priority
            if (-not $Rule) { $Rule = $match.Rule }
        }
        else {
            if ($Channels -le 2)      { $Action="Encode";  $Rule="Fallback_2.0_to_Stereo_$($StereoBitRateConfig.Encode20)" }
            elseif ($Channels -gt 6)  { $Action="Downmix"; $Rule="Fallback_7.1_to_Stereo_$($StereoBitRateConfig.Downmix71)" }
            else                      { $Action="Encode";  $Rule="Fallback_5.1_to_Stereo_$($StereoBitRateConfig.Encode51)" }
            $Passthrough=$false; $Tag=$null; $Priority=0
        }

        # SPN: every non-Removed track. Rule label is no longer reliable
        # (Atmos rule=Downmix, EAC3 2.0 rule=Passthrough both may encode).

        $PendingTracks.Add([PSCustomObject]@{
            Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
            Profile=$Profile; Title=$Title; Handler=$Handler; Language=$Lang
            SourceBitrate=$SourceBitrate
            Action=$Action; Passthrough=$Passthrough
            PassthroughTag=$Tag; Output=""; Rule=$Rule; Priority=$Priority
            NeedsNormalization=$false
            NeedsSpnScan=$true
            QualityRank=$null; RankLabel=$null
        })

        $TrackIndex++
    }

    # --- Phase B ---
    $spnCount = @($PendingTracks | Where-Object { $_.NeedsSpnScan }).Count
    if ($spnCount -gt 0) {
        Write-Host "[SPN] Peak scanning $spnCount track(s) — may take a few minutes..." -ForegroundColor Cyan
    }

    $PeakResults = [System.Collections.Concurrent.ConcurrentDictionary[int,object]]::new()
    $ScanOutput  = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    $throttleLimit = $ScanThrottle  # Local capture — $using: not valid on -ThrottleLimit
    $PendingTracks | Where-Object { $_.NeedsSpnScan } | ForEach-Object -Parallel {

        $t           = $_
        $ffmpegBin   = $using:ffmpeg
        $threadCount = $using:ThreadCount
        $threshold   = $using:PeakThresholdDB
        $inputFile   = $using:InputFile
        $peakDict    = $using:PeakResults
        $outputBag   = $using:ScanOutput

        $label    = "Track $($t.RealIndex) ($($t.Codec), $($t.Channels)ch)"
        $messages = [System.Collections.Generic.List[object]]::new()
        $messages.Add([PSCustomObject]@{
            Text  = "[SPN] Scanning peak: $label - may take a few minutes..."
            Color = "Cyan"
        })

        $raw = & $ffmpegBin -analyzeduration 200M -probesize 200M `
               -threads $threadCount -i $inputFile `
               -map "0:$($t.RealIndex)" -filter:a volumedetect -f null - 2>&1 `
               | ForEach-Object { "$_" }

        $peak = $null
        foreach ($line in $raw) {
            if ($line -match 'max_volume:\s*([-\d.]+)\s*dB') {
                $peak = [double]$Matches[1]; break
            }
        }

        if ($null -ne $peak -and $peak -gt $threshold) {
            $messages.Add([PSCustomObject]@{
                Text  = "[SPN] $label - Peak: $peak dBFS - normalization will be applied"
                Color = "Yellow"
            })
        } elseif ($null -ne $peak) {
            $messages.Add([PSCustomObject]@{
                Text  = "[SPN] $label - Peak: $peak dBFS - within threshold, skipping normalization"
                Color = "Green"
            })
        } else {
            $messages.Add([PSCustomObject]@{
                Text  = "[SPN] $label - Peak detection failed, skipping normalization"
                Color = "Red"
            })
        }

        [void]$peakDict.TryAdd($t.RealIndex, $peak)
        [void]$outputBag.Add([PSCustomObject]@{
            RealIndex = $t.RealIndex
            Messages  = $messages
        })

    } -ThrottleLimit $throttleLimit

    $ScanOutput | Sort-Object RealIndex | ForEach-Object {
        foreach ($msg in $_.Messages) {
            Write-Host $msg.Text -ForegroundColor $msg.Color
        }
    }

    # PRETRACK-QRS (between Phase B and Phase C)
    # Selects the single highest-quality multichannel source. Demotes losers
    # to Action=Removed so Build-FFmpegCommand skips them.
    # Tie-break: QualityRank → Language → SourceBitrate → RealIndex.
    $candidates = @($PendingTracks | Where-Object { $_.Action -ne 'Removed' })

    if ($candidates.Count -le 1) {
        if ($candidates.Count -eq 1) {
            Write-Host ""
            Write-Host "[PreTrack-QRS] Single non-Removed candidate — selection skipped" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host ""
        Write-Host "=== PreTrack-QRS ===" -ForegroundColor Cyan

        $mcCandidates = @($candidates | Where-Object { $_.Channels -gt 2 })

        if ($mcCandidates.Count -gt 0) {
            # Unranked MC (no ladder match) gets QualityRank=0 — loses to any match.
            foreach ($mc in $mcCandidates) {
                $qr = Resolve-TrackQuality -Codec $mc.Codec -Channels $mc.Channels `
                    -Profile $mc.Profile -Title $mc.Title -Handler $mc.Handler

                if ($qr) {
                    $mc.QualityRank = $qr.QualityRank
                    $mc.RankLabel   = $qr.RankLabel
                } else {
                    $mc.QualityRank = 0
                    $mc.RankLabel   = "Unranked Multichannel"
                }
            }

            $winner = $mcCandidates | Sort-Object `
                @{Expression = {$_.QualityRank};                   Descending = $true},
                @{Expression = {Get-LanguagePriority $_.Language}; Descending = $true},
                @{Expression = {$_.SourceBitrate};                  Descending = $true},
                @{Expression = {$_.RealIndex};                      Descending = $false} `
                | Select-Object -First 1

            $bitrateLabel = if ($winner.SourceBitrate -gt 0) { "$([math]::Round($winner.SourceBitrate / 1000))kbps" } else { "n/a" }
            Write-Host ("[PreTrack-QRS] Winner: Track {0} ({1}, {2}ch, {3}, {4}) — Rank {5} [{6}]" -f `
                $winner.RealIndex, $winner.Codec, $winner.Channels, $winner.Language, `
                $bitrateLabel, $winner.QualityRank, $winner.RankLabel) -ForegroundColor Green
        }
        else {
            # Stereo-only fallback: enforce single-track policy via tie-break only.
            $winner = $candidates | Sort-Object `
                @{Expression = {Get-LanguagePriority $_.Language}; Descending = $true},
                @{Expression = {$_.SourceBitrate};                  Descending = $true},
                @{Expression = {$_.RealIndex};                      Descending = $false} `
                | Select-Object -First 1

            $bitrateLabel = if ($winner.SourceBitrate -gt 0) { "$([math]::Round($winner.SourceBitrate / 1000))kbps" } else { "n/a" }
            Write-Host ("[PreTrack-QRS] No multichannel candidates — stereo fallback. Winner: Track {0} ({1}, {2}ch, {3}, {4})" -f `
                $winner.RealIndex, $winner.Codec, $winner.Channels, $winner.Language, $bitrateLabel) -ForegroundColor Yellow
        }

        # Demote all non-winners. NeedsSpnScan flipped false to suppress Phase C lookup.
        foreach ($t in $candidates) {
            if ($t.RealIndex -ne $winner.RealIndex) {
                $t.Action       = 'Removed'
                $t.Rule         = 'PreTrackQRS_Removed'
                $t.Output       = 'Removed (PreTrack-QRS)'
                $t.NeedsSpnScan = $false
                Write-Host ("[PreTrack-QRS] Removed: Track {0} ({1}, {2}ch, {3})" -f `
                    $t.RealIndex, $t.Codec, $t.Channels, $t.Language) -ForegroundColor DarkYellow
            }
        }
    }

    # --- Phase C ---
    foreach ($t in $PendingTracks) {
        if ($t.NeedsSpnScan) {
            $peakVal = $null
            [void]$PeakResults.TryGetValue($t.RealIndex, [ref]$peakVal)
            $t.NeedsNormalization = ($null -ne $peakVal -and $peakVal -gt $PeakThresholdDB)
        }

        $Processed.Add([PSCustomObject]@{
            Index=$t.Index; RealIndex=$t.RealIndex; Codec=$t.Codec; Channels=$t.Channels
            Profile=$t.Profile; Title=$t.Title; Handler=$t.Handler; Language=$t.Language
            SourceBitrate=$t.SourceBitrate
            Action=$t.Action; Passthrough=$t.Passthrough
            PassthroughTag=$t.PassthroughTag; Output=$t.Output
            Rule=$t.Rule; Priority=$t.Priority; NeedsNormalization=$t.NeedsNormalization
            QualityRank=$t.QualityRank; RankLabel=$t.RankLabel
        })
    }

    return $Processed
}


# ==========================
#  FFMPEG COMMAND BUILDER
# ==========================
function Build-FFmpegCommand {
    param(
        $Tracks,
        [string]    $InputFile,
        [int]       $ThreadCount,
        [string]    $StereoCodec,
        [hashtable] $StereoBitRateConfig,
        [bool]      $FoldLFE
    )

    # Output is always 2.0 stereo — all $StereoBitRateConfig keys resolve to the
    # same value, but Encode20 is semantically correct for the final output.
    $TargetBitrate = $StereoBitRateConfig.Encode20
    if ([string]::IsNullOrWhiteSpace($TargetBitrate)) {
        throw "StereoBitRateConfig does not contain a valid Encode20 entry. Check global settings."
    }

    $ffArgs   = [System.Collections.Generic.List[string]]::new()
    $loudnorm = "loudnorm=I=-23:TP=-1.5:LRA=11"  # SPN: applied only when NeedsNormalization=$true

    # LFE fold term — dynamic so $FoldLFE actually controls the matrix at runtime.
    $lfeTerm = $FoldLFE ? "+0.5*LFE" : ""

    # Pan matrices (5.1→2.0 and 7.1→2.0). Operator '=' (NOT '<') is deliberate.
    # '<' auto-normaliser counts non-existing channels in the denominator, silently
    # reducing volume on mixed-layout sources. With '=', non-existing channels
    # contribute zero signal. alimiter catches any post-fold peaks (theoretical
    # max ~2.9 with full LFE+BC fold, clamped to 0.948).
    #
    # 5.1 path: BL+SL+BC simultaneously covers '5.1', '5.1(side)', and 6.1 in one
    # matrix. FC at 0.707 (-3dB) is ITU-R BS.775. 7.1 path: BL+SL only (no BC in
    # 7.1), pinned via aformat to guard against Atmos decoder layout ambiguity.
    # SL fold-through 0.5 = 0.707*0.707 (chained SL→BL→stereo per BS.775).
    $pan51 = "pan=stereo|FL=FL+0.707*FC+0.707*BL+0.707*SL+0.5*BC${lfeTerm}|FR=FR+0.707*FC+0.707*BR+0.707*SR+0.5*BC${lfeTerm}"
    $pan71 = "aformat=channel_layouts=7.1,pan=stereo|FL=FL+0.707*FC+0.707*BL+0.5*SL${lfeTerm}|FR=FR+0.707*FC+0.707*BR+0.5*SR${lfeTerm}"

    $alimiter = "alimiter=limit=0.948:attack=5:release=50:level=disabled:latency=1"

    # Global FFmpeg options. Input options before -i, output options after.
    $ffArgs.AddRange([string[]](
        "-y",
        "-loglevel",             "error",   #changed from warning
        "-stats",
        "-threads",              $ThreadCount,
        "-analyzeduration",      "200M",
        "-probesize",            "200M",
        "-err_detect",           "ignore_err",
        "-drc_scale",            "0",
        "-i",                    $InputFile,
        "-avoid_negative_ts",    "make_zero",
        "-max_muxing_queue_size","14000",
        "-map",                  "0:v?",
        "-c:v",                  "copy",
        "-map_metadata",         "0",
        "-map_chapters",         "0"
    ))

    $i = 0
    foreach ($t in $Tracks) {
        if ($t.Action -eq "Removed") { continue }

        # Effective passthrough requires ALL:
        # (1) rule engine Passthrough  (2) source = 2ch  (3) StereoCodec = EAC3
        $isEffectivePassthrough = $t.Passthrough -and $t.Channels -eq 2 -and $StereoCodec -eq "EAC3"

        $ffArgs.AddRange([string[]]("-map", "0:$($t.RealIndex)"))
        $LangTag = " [$($t.Language)]"

        if ($isEffectivePassthrough) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i",          "copy",
                "-metadata:s:a:$i", "title=$($t.PassthroughTag)$LangTag"
            ))
            $t.Output = "$($t.PassthroughTag)$LangTag"
        }
        else {
            # Encode path: channel-count routing supersedes rule label.
            # >=8ch → 7.1 pan; >=3ch → 5.1 pan (covers 5.1/5.1side/6.1); <=2ch → no pan (mono via -ac 2).
            $pre = $t.NeedsNormalization ? "$loudnorm," : ""

            if ($t.Channels -ge 8) {
                $filterChain = "${pre}${pan71},${alimiter}"
                $ffArgs.AddRange([string[]]("-filter:a:$i", $filterChain))
                $opLabel = "Stereo Downmix 7.1->2.0"
            }
            elseif ($t.Channels -ge 3) {
                $filterChain = "${pre}${pan51},${alimiter}"
                $ffArgs.AddRange([string[]]("-filter:a:$i", $filterChain))
                $opLabel = "Stereo Downmix 5.1->2.0"
            }
            else {
                # Mono/stereo: no pan filter. Mono upmixed by -ac 2 in encoder block.
                if ($t.NeedsNormalization) {
                    $ffArgs.AddRange([string[]]("-filter:a:$i", $loudnorm))
                }
                $opLabel = "Stereo"
            }

            switch ($StereoCodec) {
                "EAC3" {
                    # EAC3 stereo flags:
                    $ffArgs.AddRange([string[]](
                        "-c:a:$i",            "eac3",
                        "-ac",                "2",
                        "-ar",                "48000",
                        "-b:a:$i",            $TargetBitrate,
                        "-dialnorm",          "-31",
                        "-dsur_mode",         "notindicated",
                        "-dmix_mode",         "loro",
                        "-stereo_rematrixing","1",
                        "-cutoff",            "20000",
                        "-metadata:s:a:$i",   "title=$opLabel EAC3 ($TargetBitrate)$LangTag"
                    ))
                    $t.Output = "$opLabel EAC3 ($TargetBitrate)$LangTag"
                }
                "Opus" {
                    # Opus stereo flags:
                    # -vbr on, Default but explicit
                    # -compression_level 10, Highest quality (locked across libopus builds)
                    # -frame_duration 60, Max frame size. best quality / stability
                    # -application audio, Correct for film/music mixed content
                    # -ar 48000, libopus internal output; explicit for clarity
                    # Excluded: -cutoff (libopus uses different syntax; auto at our bitrate)
                    # Excluded: -dialnorm -dsur_mode -dmix_mode -stereo_rematrixing (EAC3-only)
                    $ffArgs.AddRange([string[]](
                        "-c:a:$i",            "libopus",
                        "-ac",                "2",
                        "-ar",                "48000",
                        "-b:a:$i",            $TargetBitrate,
                        "-vbr",               "on",
                        "-compression_level", "10",
                        "-frame_duration",    "60",
                        "-application",       "audio",
                        "-metadata:s:a:$i",   "title=$opLabel Opus ($TargetBitrate)$LangTag"
                    ))
                    $t.Output = "$opLabel Opus ($TargetBitrate)$LangTag"
                }
                default {
                    throw "Unhandled StereoCodec value '$StereoCodec' in Build-FFmpegCommand. Valid: EAC3, Opus."
                }
            }
        }

        $ffArgs.AddRange([string[]]("-metadata:s:a:$i", "language=$($t.Language)"))

        $disp = $i -eq 0 ? "default" : "0"
        $ffArgs.AddRange([string[]]("-disposition:a:$i", $disp))

        $i++
    }

    # --- Subtitles, attachments, output path ---
    $ffArgs.AddRange([string[]]("-map", "0:s?", "-c:s", "copy"))
    $ffArgs.AddRange([string[]]("-map", "0:t?", "-c:t", "copy"))

    $codecSuffix = $StereoCodec.ToLower()
    $outDir  = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($InputFile))
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile) + "_stereo_$codecSuffix.mkv"
    $ffArgs.Add([System.IO.Path]::Combine($outDir, $outName))

    return $ffArgs
}

# ==================
#  MAIN EXECUTION
# ==================
Write-Host "=== ConvertAudioEngine-Stereo v1.4.0 ===" -ForegroundColor Cyan
Write-Host "Stereo codec: $StereoCodec @ $($StereoBitRateConfig.Encode20) | LFE fold: $FoldLFE" -ForegroundColor Cyan
Write-Host ""

Write-Host "=== Probing Audio Streams ===" -ForegroundColor Cyan
$streams = Get-AudioStreams -File $InputFile

if ($null -eq $streams -or @($streams).Count -eq 0) {
    Write-Error "No audio streams found in input file: $InputFile"
    exit 1
}

Write-Host "=== Processing Tracks ===" -ForegroundColor Cyan
$tracks = Convert-AudioTracks -Streams $streams

# Post-PreTrack-QRS sanity check (defensive — winner should always survive).
$survivors = @($tracks | Where-Object { $_.Action -ne 'Removed' })
if ($survivors.Count -eq 0) {
    Write-Error "All audio tracks marked Removed after PreTrack-QRS. Cannot produce output."
    exit 1
}

Write-Host ""
Write-Host "=== Building FFmpeg Command ===" -ForegroundColor Cyan
$cmd = Build-FFmpegCommand -Tracks $tracks -InputFile $InputFile -ThreadCount $ThreadCount `
                            -StereoCodec $StereoCodec -StereoBitRateConfig $StereoBitRateConfig `
                            -FoldLFE $FoldLFE

$timer = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "[Timer] Encoding in progress..." -ForegroundColor Cyan
& $ffmpeg @cmd
$timer.Stop()

if ($LASTEXITCODE -ne 0) {
    Write-Error "FFmpeg exited with code $LASTEXITCODE. Output may be incomplete."
    exit $LASTEXITCODE
}

$e        = $timer.Elapsed
$elapsed  = if ($e.Hours -gt 0) { "{0}h {1}m {2}s" -f $e.Hours, $e.Minutes, $e.Seconds } `
            else                 { "{0}m {1}s"       -f $e.Minutes, $e.Seconds }
$tColor   = if ($e.TotalMinutes -lt 5) { "Green" } elseif ($e.TotalMinutes -lt 15) { "Yellow" } else { "Red" }
Write-Host "[Timer] Encoding completed — $elapsed elapsed" -ForegroundColor $tColor

# ==================
#  SUMMARY ENGINE
# ==================

Write-Host ""
Write-Host "=== AUDIO PROCESSING SUMMARY ===" -ForegroundColor Cyan
Write-Host ""

# Header
$header = "{0,-4} {1,-8} {2,-5} {3,-26} {4,-44} {5,-5} {6}" -f `
    "Idx","Codec","Ch","Action","Output","Pri","Rule"

Write-Host $header -ForegroundColor White
Write-Host ("-" * ($header.Length + 10)) -ForegroundColor DarkGray

foreach ($t in $tracks) {

    # Color/ActionLabel from $t.Output (ground truth in Build-FFmpegCommand), not $t.Action.
    $Color = if ($t.Action -eq "Removed") {
        "Red"
    } elseif ($t.Output -match "Passthrough") {
        "Green"
    } elseif ($t.Output -match "Downmix") {
        "DarkCyan"
    } else {
        "Blue"
    }

    $ActionLabel = if ($t.Action -eq "Removed") {
        "Removed"
    } elseif ($t.Output -match "Passthrough") {
        "Passthrough"
    } elseif ($t.Output -match "Downmix 7\.1") {
        "Downmix 7.1->2.0 ($($StereoBitRateConfig.Downmix71))"
    } elseif ($t.Output -match "Downmix 5\.1") {
        "Downmix 5.1->2.0 ($($StereoBitRateConfig.Encode51))"
    } else {
        "Encode Stereo ($($StereoBitRateConfig.Encode20))"
    }

    $OutputLabel = $t.Output ? $t.Output : "(none)"
    $PriLabel    = $t.Action -eq "Removed" ? "-" : $t.Priority

    $line = "{0,-4} {1,-8} {2,-5} {3,-26} {4,-44} {5,-5} {6}" -f `
        $t.Index, $t.Codec, $t.Channels, $ActionLabel, $OutputLabel, $PriLabel, $t.Rule

    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * ($header.Length + 10)) -ForegroundColor DarkGray
Write-Host "=== END OF SUMMARY ===" -ForegroundColor Cyan
Write-Host ""