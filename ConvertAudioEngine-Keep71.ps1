# ==================================================================
#  ConvertAudioEngine-Keep71 v3.1.2 — Production Use
#  PowerShell 7.6, FFmpeg 8.1
#
#  Foundation: 3-phase SPN architecture (Phase A → Phase B → Phase C)
#  Includes SPN-TP (true-peak ceiling, attenuation-only) and malformed-layout guards.
#
#  Removes all 2.0 tracks (delegated to ConvertAudioEngine-Stereo)
#
#  Bitrate tiers (centralized via $BitRateConfig):
#    Downmix 7.1 → 5.1         = $BitRateConfig.Downmix51
#    Re-encode 5.1 sources     = $BitRateConfig.ReEncode51
#    Pass+Copy DDP 5.1         = $BitRateConfig.PassCopy51
#    DTS-HD MA/HRA multichannel encode = $BitRateConfig.EncodeDTSHD
#
#  Processing rules:
#    AAC 7.1 / TrueHD 7.1: keep original + add DDP 5.1 copy (Pass+Copy)
#    Other 7.1 sources → downmix to DDP 5.1 ($BitRateConfig.Downmix51)
#    5.1 sources → re-encode to DDP 5.1 ($BitRateConfig.ReEncode51)
#    EAC3/TrueHD 5.1 passthrough
#    DTS-HD MA/HRA → DDP 5.1 ($BitRateConfig.EncodeDTSHD)
#
#  Supported codecs:
#    AAC, EAC3-ATMOS, TrueHD, DTS, PCM, FLAC
#
#  Output:
#    MKV container with video copy + 7.1 passthrough + DDP 5.1 compatibility
#    (passthrough, Pass+Copy, downmix, or re-encode depending on rule outcome)
#
#  Downmix quality:
#    ITU-R BS.775 pan matrix for 7.1 → 5.1 accuracy
# ==================================================================
#Requires -Version 7.6

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputFile
)

# =========================================
#  ENGINE: GLOBAL SETTINGS
# =========================================
$ThreadCount  = 8 # User-adjustable (4-16).
$ScanThrottle = [Math]::Max(1, [int]([Environment]::ProcessorCount / $ThreadCount))  # SPN parallel scan throttle
[double]$TruePeakCeilingDB = -1.0  # SPN-TP: measured true peak above this ceiling (dBTP) gets a
                                   # static, attenuation-only trim back to it. Gain clamped <= 0 (never boosts);
                                   # the alimiter backstop on the fold-to-5.1 paths sits just below at ~-0.47 dBFS.
$CommentaryKeywords = @("commentary","director","producer","writer","cast","behind","bonus","alt","interview")
$CommentaryPattern  = [regex]::new($CommentaryKeywords -join '|', 'IgnoreCase,Compiled')

# Rule-matched keepers score >=61; fallbacks (plain AC-3 5.1, etc.) score 0.
# Losers below this threshold are removed. Winner always exempt. MaxValue = strict last-man-standing.
$QRSMinPriority = 60

# ========================================================================
#  ENGINE: DOWNMIX FILTERGRAPH — SINGLE SOURCE OF TRUTH
#  Used in lock-step by Phase-B true-peak measurement AND Build-FFmpegCommand
#  (panFilter71 for Pass+Copy, panFilter for Downmix). Any divergence means the
#  static safety trim would be computed against a different signal than ships.
# ========================================================================
$DownmixChain = 'aformat=channel_layouts=7.1,pan=5.1|FL=FL|FR=FR|FC=FC|LFE=LFE|BL=BL+0.707*SL|BR=BR+0.707*SR'

# ========================================
#  ENGINE: BITRATE CONFIGURATION TABLE
#  Single source of truth — changes here propagate to rule groups, fallback, and safety clamps.
#  Common EAC3 5.1: 384,448,512,640,768
#  Higher pipeline bitrates: 1024,1152,1280,1408,1536
#  You can safely adjust these numbers, within '....k'
# ========================================
$BitRateConfig = @{
    Downmix51   = '1152k'   # 7.1 → 5.1 downmix via pan filter
    ReEncode51  = '768k'    # 5.1 source re-encode
    PassCopy51  = '1152k'   # DDP 5.1 compatibility track (Pass+Copy)
    EncodeDTSHD = '1024k'   # DTS-HD MA/HRA lossless multichannel encode (5.1 and 7.1)
}

$ext     = $IsWindows ? ".exe" : ""
$ffmpeg  = Join-Path $PSScriptRoot "ffmpeg$ext"
$ffprobe = Join-Path $PSScriptRoot "ffprobe$ext"

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

# =========================
#  ENGINE: CHANNEL FILTER
# =========================
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

# =============================
#  ENGINE: AUDIO RULE GROUPS
# =============================

# AAC FAMILY
$Rules_AAC = @(
    # Rule1: AAC 7.1 Pass + DDP 5.1 copy (see header notes 3,4)
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Pass+Copy"
        Bitrate        = $BitRateConfig.PassCopy51
        PassthroughTag = "AAC_7.1_Pass"
        Rule           = "AAC_7.1_Plus_DDP5.1"
        Priority       = 93
    },
    # Rule2: AAC 5.1 → Encode
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = $BitRateConfig.ReEncode51
        PassthroughTag = $null
        Rule           = "AAC_5.1_Encode_$($BitRateConfig.ReEncode51)"
        Priority       = 91
    },
    # Rule3: AAC 2.0 → REMOVED (delegated to ConvertAudioEngine-Stereo)
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Removed"
        Bitrate        = $null
        PassthroughTag = $null
        Rule           = "AAC_2.0_Removed"
        Priority       = 90
    }
)

# EAC3 / ATMOS FAMILY
$Rules_EAC3_Atmos = @(
    # Rule1: EAC3 Atmos (JOC) pass
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = [ChannelFilter]::new("gt", 2)
        ProfileRegex   = "\bJOC\b|\bAtmos\b"
        TrustTitleProfile = $true   # JOC/Atmos under-reported by ffprobe; title/handler fallback (eac3-gated)
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "EAC3_Atmos_Passthrough"
        Rule           = "EAC3_Atmos_Passthrough"
        Priority       = 100
    },
    # Rule2: EAC3 7.1 → Downmix to 5.1
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = $BitRateConfig.Downmix51
        PassthroughTag = $null
        Rule           = "EAC3_7.1_Downmix_$($BitRateConfig.Downmix51)"
        Priority       = 99
    },
    # Rule3: EAC3 5.1 pass
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
    # Rule4: EAC3 2.0 → REMOVED (delegated to ConvertAudioEngine-Stereo)
    [PSCustomObject]@{
        CodecRegex     = "^(eac3)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Removed"
        Bitrate        = $null
        PassthroughTag = $null
        Rule           = "EAC3_2.0_Removed"
        Priority       = 97
    }
)

# TRUEHD FAMILY
$Rules_TrueHD = @(
    # Rule1: TrueHD 7.1 pass + DDP 5.1 copy (see header notes 3,4)
    [PSCustomObject]@{
        CodecRegex     = "^(mlp|truehd|true-hd)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Pass+Copy"
        Bitrate        = $BitRateConfig.PassCopy51
        PassthroughTag = "TrueHD_7.1_Pass"
        Rule           = "TrueHD_7.1_Plus_DDP5.1"
        Priority       = 110
    },
    # Rule2: TrueHD 5.1 pass
    [PSCustomObject]@{
        CodecRegex     = "^(mlp|truehd|true-hd)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Passthrough"
        Bitrate        = $null
        PassthroughTag = "TrueHD_5.1_Passthrough"
        Rule           = "TrueHD_5.1_Passthrough"
        Priority       = 81
    }
)

# DTS FAMILY
$Rules_DTS = @(
    # Rule1: DTS-HD multichannel → Encode
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = [ChannelFilter]::MoreThanTwo
        ProfileRegex   = "HD|MA|HRA"
        Action         = "Encode"
        Bitrate        = $BitRateConfig.EncodeDTSHD
        PassthroughTag = $null
        Rule           = "DTSHD_Multichannel_Encode_$($BitRateConfig.EncodeDTSHD)"
        Priority       = 73
    },
    # Rule2: DTS core multichannel → Encode
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = [ChannelFilter]::MoreThanTwo
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = $BitRateConfig.ReEncode51
        PassthroughTag = $null
        Rule           = "DTS_Multichannel_Encode_$($BitRateConfig.ReEncode51)"
        Priority       = 72
    },
    # Rule3: DTS-HD 2.0 → REMOVED (delegated to ConvertAudioEngine-Stereo)
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = 2
        ProfileRegex   = "HD|MA|HRA"
        Action         = "Removed"
        Bitrate        = $null
        PassthroughTag = $null
        Rule           = "DTSHD_2.0_Removed"
        Priority       = 71
    },
    # Rule4: DTS 2.0 → REMOVED (delegated to ConvertAudioEngine-Stereo)
    [PSCustomObject]@{
        CodecRegex     = "^(dts)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Removed"
        Bitrate        = $null
        PassthroughTag = $null
        Rule           = "DTS_2.0_Removed"
        Priority       = 70
    }
)

# PCM / FLAC FAMILY
$Rules_PCMFLAC = @(
    # Rule1: PCM/FLAC 7.1 → Downmix to 5.1
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = $BitRateConfig.Downmix51
        PassthroughTag = $null
        Rule           = "PCMFLAC_7.1_Downmix_$($BitRateConfig.Downmix51)"
        Priority       = 62
    },
    # Rule2: PCM/FLAC 5.1 → Encode
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 6
        ProfileRegex   = $null
        Action         = "Encode"
        Bitrate        = $BitRateConfig.ReEncode51
        PassthroughTag = $null
        Rule           = "PCMFLAC_5.1_Encode_$($BitRateConfig.ReEncode51)"
        Priority       = 61
    },
    # Rule3: PCM/FLAC 2.0 → REMOVED (delegated to ConvertAudioEngine-Stereo)
    [PSCustomObject]@{
        CodecRegex     = "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"
        Channels       = 2
        ProfileRegex   = $null
        Action         = "Removed"
        Bitrate        = $null
        PassthroughTag = $null
        Rule           = "PCMFLAC_2.0_Removed"
        Priority       = 60
    }
)

# ============================================================
#  ENGINE: MERGE ALL AUDIO RULE GROUPS
#  (Adding a new codec = define $Rules_XYZ, then append here)
# ============================================================
$AudioRules = $Rules_EAC3_Atmos + $Rules_TrueHD + $Rules_DTS + $Rules_AAC + $Rules_PCMFLAC

# Priority sort
$AudioRules = $AudioRules | Sort-Object Priority -Descending

# Precompile regex patterns — avoids repeated recompilation per-track at runtime
$AudioRules = $AudioRules | Select-Object -Property *,
    @{ Name='CodecRegexObj';   Expression={
        $_.CodecRegex   ? [regex]::new($_.CodecRegex,   'IgnoreCase,Compiled') : $null
    }},
    @{ Name='ProfileRegexObj'; Expression={
        $_.ProfileRegex ? [regex]::new($_.ProfileRegex, 'IgnoreCase,Compiled') : $null
    }}

# ==========================
#  PROBE
# ==========================
function Get-AudioStreams {
    param([string]$File)

    $probeArgs = @(
        "-analyzeduration","200M",  # match FFmpeg probe
        "-probesize","200M",        # match FFmpeg probe
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
    return $script:CommentaryPattern.IsMatch($Title)
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
            # Profile field is authoritative. Container title/handler are unreliable
            # and must NOT be allowed to promote a track into a higher tier, so they
            # are consulted only when the rule explicitly opts in via TrustTitleProfile
            # (currently EAC3 Atmos only). See v3.0.6 fix note in header.
            $profileMatched = $Profile -and $r.ProfileRegexObj.IsMatch($Profile)
            if (-not $profileMatched -and $r.TrustTitleProfile) {
                $profileMatched = ($Title   -and $r.ProfileRegexObj.IsMatch($Title)) -or
                                  ($Handler -and $r.ProfileRegexObj.IsMatch($Handler))
            }
            if (-not $profileMatched) { continue }
        }

        return $r
    }

    return $null
}

# --- Malformed-Layout Guard Helpers ---
function New-RemovedTrack {
    param(
        $TrackIndex, $RealIndex, $Codec, $Channels,
        $Profile, $Title, $Lang,
        $Rule
    )

    [PSCustomObject]@{
        Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
        Profile=$Profile; Title=$Title; Language=$Lang
        Action="Removed"; Passthrough=$false; Downmix=$false
        Bitrate=$null; PassthroughTag=$null; Output="Malformed"
        Rule=$Rule; Priority=0; MeasuredTP=$null; PeakSafetyGainDb=0.0; NeedsPeakSafety=$false
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

function Select-WinnerPerLanguage {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[object]]$Processed
    )

    $candidates = @($Processed | Where-Object { $_.Action -ne 'Removed' })
    if ($candidates.Count -eq 0) { return }

    $langGroups = $candidates | Group-Object -Property { $_.Language.ToLower() }

    foreach ($g in $langGroups) {
        $sorted = @($g.Group | Sort-Object `
            @{ Expression = 'Priority';  Descending = $true  }, `
            @{ Expression = 'RealIndex'; Descending = $false })

        $winner = $sorted[0]
        $losers = @($sorted | Select-Object -Skip 1)

        $profileLabel = $winner.Profile ? $winner.Profile : "n/a"
        Write-Host ("[PreTrack-QRS] Winner: Track {0} ({1}, {2}ch, {3}, {4}) - Rank {5} [{6}]" -f `
            $winner.Index, $winner.Codec, $winner.Channels, $winner.Language, `
            $profileLabel, $winner.Priority, $winner.Rule) -ForegroundColor Green

        foreach ($l in $losers) {
            if ($l.Priority -lt $script:QRSMinPriority) {
                $l.Action      = 'Removed'
                $l.Passthrough = $false
                $l.Downmix     = $false
                $l.Output      = 'QRS_Removed'
                $l.Rule        = 'PreTrack_QRS_Removed'
                $l.Priority    = 0

                Write-Host ("[PreTrack-QRS] Removed: Track {0} ({1}, {2}ch, {3})" -f `
                    $l.Index, $l.Codec, $l.Channels, $l.Language) -ForegroundColor Red
            }
        }

        $groupSurvivors = @($g.Group | Where-Object { $_.Action -ne 'Removed' })
        if ($groupSurvivors.Count -eq 0) {
            throw "PreTrack-QRS sanity check failed: language group '$($g.Name)' had candidates but zero survivors."
        }
    }
}

# =============================
#  PROCESS TRACKS
# =============================
function Convert-AudioTracks {
    param($Streams)

    $Processed  = [System.Collections.Generic.List[object]]::new()
    $TrackIndex = 0

    foreach ($s in $Streams) {

        $Codec    = $s.codec_name
        $Channels = [int]$s.channels
        $Profile  = $s.profile

        $Title   = $s.tags?.title
        $Lang    = $s.tags?.language ?? "eng"
        $Handler = $s.tags?.handler_name

        $RealIndex = $s.index
        $Rule      = $null

        # --- Centralized Malformed-Layout Guard Module ---
        $guard = Apply-MalformedLayoutGuards -Codec $Codec -Channels $Channels -RealIndex $RealIndex
        $Channels = $guard.Channels
        if ($guard.Rule) { $Rule = $guard.Rule }

        if ($guard.SkipTrack) {
            $removed = New-RemovedTrack $TrackIndex $RealIndex $Codec $Channels $Profile $Title $Lang $Rule
            $Processed.Add($removed)
            $TrackIndex++; continue
        }

        # --- Commentary removal ---
        if (Test-IsCommentary -Channels $Channels -Title $Title) {
            $Processed.Add([PSCustomObject]@{
                Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
                Profile=$Profile; Title=$Title; Language=$Lang
                Action="Removed"; Passthrough=$false; Downmix=$false
                Bitrate=$null; PassthroughTag=$null; Output="Commentary"
                Rule="Commentary_Removed"; Priority=0; MeasuredTP=$null; PeakSafetyGainDb=0.0; NeedsPeakSafety=$false
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
            $Rule       ??= $match.Rule
        }
        else {
            if      ($Channels -le 2) { $Action = "Removed"; $Bitrate = $null;                       $Rule ??= "Fallback_2.0_Removed" }
            elseif  ($Channels -gt 6) { $Action = "Downmix"; $Bitrate = $BitRateConfig.Downmix51;   $Rule ??= "Fallback_7.1_$($BitRateConfig.Downmix51)" }
            else                      { $Action = "Encode";  $Bitrate = $BitRateConfig.ReEncode51;  $Rule ??= "Fallback_5.1_$($BitRateConfig.ReEncode51)"  }
            $Passthrough = $false; $Downmix = ($Action -eq "Downmix"); $Tag = $null; $Priority = 0
        }

        # --- Safety Audit ---
        # Clamps compare against table values to remain in sync after bitrate changes.
        # DTS-HD MA/HRA at $BitRateConfig.EncodeDTSHD does NOT match $BitRateConfig.Downmix51,
        # so the 6ch clamp will not fire on legitimate DTS-HD 5.1 tracks.
        # Pass+Copy 8ch tracks carry $BitRateConfig.PassCopy51 which does NOT match $BitRateConfig.ReEncode51,
        # so the 8ch clamp will not fire on legitimate Pass+Copy tracks.
        if ($Action -eq "Pass+Copy" -and $Bitrate -ne $BitRateConfig.PassCopy51) {
            Write-Warning ("Track {0} ({1}): Pass+Copy bitrate {2} overridden to {3}." -f $TrackIndex, $Codec, $Bitrate, $BitRateConfig.PassCopy51)
            $Bitrate = $BitRateConfig.PassCopy51
        }
        elseif ($Channels -eq 6 -and $Bitrate -eq $BitRateConfig.Downmix51) {
            Write-Warning ("Track {0} ({1}): 6ch rule had {2}-clamped to {3}." -f $TrackIndex, $Codec, $BitRateConfig.Downmix51, $BitRateConfig.ReEncode51)
            $Bitrate = $BitRateConfig.ReEncode51
        }
        elseif ($Channels -eq 8 -and $Bitrate -eq $BitRateConfig.ReEncode51) {
            Write-Warning ("Track {0} ({1}): 8ch rule had {2}-clamped to {3}." -f $TrackIndex, $Codec, $BitRateConfig.ReEncode51, $BitRateConfig.Downmix51)
            $Bitrate = $BitRateConfig.Downmix51
        }

        # --- Safe Peak Normalizer (SPN-TP) ---
        # SPN-TP fields are placeholders initialized here; patched IN PLACE in Phase C
        # after the parallel true-peak scans. A [PSCustomObject] throws on a write to a
        # property that does not already exist, and a zero-scan file skips Phase C entirely,
        # so these defaults (no trim) are exactly what Build sees in that case.

        $Processed.Add([PSCustomObject]@{
            Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
            Profile=$Profile; Title=$Title; Language=$Lang; Action=$Action
            Passthrough=$Passthrough; Downmix=$Downmix; Bitrate=$Bitrate
            PassthroughTag=$Tag
            Output=($Action -eq "Removed" ? "2.0_Removed" : "")
            Rule=$Rule; Priority=$Priority
            MeasuredTP=$null; PeakSafetyGainDb=0.0; NeedsPeakSafety=$false
        })

        $TrackIndex++
    }
    # ── Phase A end ──────────────────────────────────────

    Select-WinnerPerLanguage -Processed $Processed

    # ── Phase B: Parallel SPN scans ──────────────────────
    $tracksToScan = @($Processed | Where-Object { $_.Action -notin "Passthrough","Removed" })

    if ($tracksToScan.Count -gt 0) {

        Write-Host "[SPN] True-peak scanning $($tracksToScan.Count) track(s) — may take a few minutes..." -ForegroundColor Cyan

        $spnResults     = [System.Collections.Concurrent.ConcurrentDictionary[int,object]]::new()

        # Capture all $script: references as locals for safe $using: transport
        $inputFile    = $script:InputFile
        $ffmpegBin    = $script:ffmpeg
        $threadCount  = $script:ThreadCount
        $downmixChain = $script:DownmixChain
        $scanThrottle = $script:ScanThrottle

        $tracksToScan | ForEach-Object -Parallel {
            $idx      = $_.RealIndex
            $codec    = $_.Codec
            $channels = $_.Channels
            $downmix  = $_.Downmix
            $action   = $_.Action
            $label    = "Track $idx ($codec, ${channels}ch)"
            Write-Host "[SPN] Measuring true peak: $label - may take a few minutes..." -ForegroundColor Cyan
            $log      = [System.Collections.Generic.List[object]]::new()

            # Pass+Copy's only encoded output is the 5.1 duplicate (panFilter71 = $DownmixChain),
            # so it is measured THROUGH the downmix chain alongside true Downmix tracks even though
            # $_.Downmix is $false. The 5.1 Encode/re-encode path measures as-is (6ch -> 6ch is a
            # no-op rematrix). A >6ch Encode source (DTS/DTS-HD 7.1) is reduced 7.1->5.1 INSIDE the
            # measure filtergraph via aresample=ochl=5.1 so loudnorm sees the same 5.1 signal Build
            # encodes with -ac 6 (the reduction MUST precede loudnorm; an output-side -ac would land
            # after it and leave the 7.1 source measured). Interpolate a PLAIN local for the chain —
            # never $using: inside a measure-filter string.
            $dmc = $using:downmixChain
            $measureFilter =
                if ($downmix -or $action -eq 'Pass+Copy') { "$dmc,loudnorm=print_format=json" }
                elseif ($channels -gt 6)                  { "aresample=ochl=5.1,loudnorm=print_format=json" }
                else                                      { "loudnorm=print_format=json" }

            $raw = & $using:ffmpegBin -analyzeduration 200M -probesize 200M `
                   -threads $using:threadCount -i $using:inputFile `
                   -map "0:$idx" -filter:a $measureFilter -f null - 2>&1

            # loudnorm prints its JSON summary to stderr; coerce the capture to one string first.
            $rawText = $raw -join "`n"
            $i_val  = if ($rawText -match '"input_i"\s*:\s*"([^"]+)"')  { $Matches[1] } else { $null }
            $tp_val = if ($rawText -match '"input_tp"\s*:\s*"([^"]+)"') { $Matches[1] } else { $null }

            if ($tp_val) {
                $msg   = "[SPN] $label - Measured true peak: $tp_val dBTP (integrated $i_val LUFS)"
                $color = "Green"
            } else {
                $msg   = "[SPN] $label - True-peak detection failed - no static trim, alimiter will backstop"
                $color = "Red"
            }

            $log.Add([PSCustomObject]@{ Message=$msg; Color=$color })
            [void]($using:spnResults).TryAdd($idx, [PSCustomObject]@{
                I   = $i_val
                TP  = $tp_val
                Log = $log.ToArray()
            })
        } -ThrottleLimit $scanThrottle

        # ── Phase C: Sequential merge — messages printed in RealIndex order ──
        foreach ($t in $Processed) {
            if (-not $spnResults.ContainsKey($t.RealIndex)) { continue }
            $r = $spnResults[$t.RealIndex]
            foreach ($entry in $r.Log) {
                Write-Host $entry.Message -ForegroundColor $entry.Color
            }
            # Deterministic true-peak safety decision: gainDb = min(0, ceiling - measured_tp).
            # Attenuate-only; never boosts. Parsed with InvariantCulture (loudnorm JSON always
            # uses '.'; a comma-decimal locale would otherwise mis-read the value).
            if ($r.TP) {
                $peakTP             = [double]::Parse($r.TP, [System.Globalization.CultureInfo]::InvariantCulture)
                $t.MeasuredTP       = $peakTP
                $t.PeakSafetyGainDb = [Math]::Min(0.0, $script:TruePeakCeilingDB - $peakTP)
                $t.NeedsPeakSafety  = ($t.PeakSafetyGainDb -lt 0.0)
            }
            # else: measurement missing -> fields keep their Phase-A defaults (no trim).
        }
    }

    return $Processed
}

# =============================
#  FFMPEG COMMAND BUILDER
# =============================
function Build-FFmpegCommand {
    param($Tracks, $InputFile, $ThreadCount)

    $ffArgs = [System.Collections.Generic.List[string]]::new()

    # ------------------------------------------------------------------
    #  Global options
    #  NOTE — order matters for FFmpeg: all input options must appear
    #  before -i. Output options (-avoid_negative_ts, -max_muxing_queue_size)
    #  must appear after -i.
    # ------------------------------------------------------------------
    $ffArgs.AddRange([string[]](
        "-y",
        "-loglevel",             "error",        # changed from warning
        "-stats",
        "-threads",              $ThreadCount,
        "-analyzeduration",      "200M",         # ensures 7.1 channel layout fully parsed (TrueHD+AAC)
        "-probesize",            "200M",         # raised from 100M
        "-err_detect",           "ignore_err",
        "-drc_scale",            "0",
        "-i",                    $InputFile,
        "-avoid_negative_ts",    "make_zero",    # clamps negative PTS from 7.1 streams to zero
        "-max_muxing_queue_size","14000",        # Pass+Copy = 3 streams: video c, 7.1 p, EAC3 enc
        "-map",                  "0:v?",
        "-c:v",                  "copy",
        "-map_metadata",         "0",
        "-map_chapters",         "0"
    ))

    $i = 0
    foreach ($t in $Tracks) {
        if ($t.Action -eq "Removed") { continue }

        $ffArgs.AddRange([string[]]("-map","0:$($t.RealIndex)"))
        $LangTag = " [$($t.Language)]"

        # Static attenuation-only true-peak trim (always <= 0 dB), built once per track and
        # injected per path below. InvariantCulture prevents a comma-decimal locale from emitting
        # an unparseable "volume=-0,7dB". Empty string when no trim is needed.
        $volTrim = ""
        if ($t.NeedsPeakSafety) {
            $g = $t.PeakSafetyGainDb.ToString('0.0##', [System.Globalization.CultureInfo]::InvariantCulture)
            $volTrim = "volume=${g}dB"
        }

        if ($t.Passthrough) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i","copy",
                "-metadata:s:a:$i","title=$($t.PassthroughTag)$LangTag"
            ))
            $t.Output = "$($t.PassthroughTag)$LangTag"
        }
        elseif ($t.Action -eq "Pass+Copy") {

            # --- 1. Original 7.1 passthrough (TrueHD or AAC)
            $ffArgs.AddRange([string[]](
                "-c:a:$i","copy",
                "-metadata:s:a:$i","title=$($t.PassthroughTag) [$($t.Language)]",
                "-metadata:s:a:$i","language=$($t.Language)"
            ))

            $disp = $i -eq 0 ? "default" : "0"
            $ffArgs.AddRange([string[]]("-disposition:a:$i",$disp))

            $t.Output = "$($t.PassthroughTag) [$($t.Language)]"
            $i++

            # --- 2. Duplicate as DDP 5.1 ---
            # (4) see header note. Order: pan (downmix) -> static true-peak trim (if any) ->
            # alimiter. $DownmixChain is the SAME graph measured in Phase B, so the trim matches
            # the signal. alimiter (limit=0.948) backstops post-pan peaks above -0.47 dBFS
            # (attack=5ms, release=50ms). The 7.1 copy stream above stays a pure -c:a copy.
            $post71      = $t.NeedsPeakSafety ? ",$volTrim" : ""
            $panFilter71 = "$($script:DownmixChain)${post71},alimiter=limit=0.948:attack=5:release=50:level=false:latency=1"
            $ffArgs.AddRange([string[]](
                "-map","0:$($t.RealIndex)",
                "-filter:a:$i", $panFilter71,
                "-c:a:$i","eac3",
                "-ar","48000",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-cutoff","20000",
                "-metadata:s:a:$i","title=DD+ 5.1 ($($t.Bitrate)) [$($t.Language)]",
                "-metadata:s:a:$i","language=$($t.Language)"
            ))

            $ffArgs.AddRange([string[]]("-disposition:a:$i","0"))

            $t.Output += " + DDP 5.1 ($($t.Bitrate)) [$($t.Language)]"
            $i++
        }
        elseif ($t.Downmix) {
            # --- 7.1 → 5.1 DOWNMIX PATH ---
            # (4) see header note
            # Downmixes any 7.1 audio track to EAC3 (DD+ 5.1) using an ITU-R BS.775
            # compliant pan matrix. Side surrounds (SL/SR) are folded into the rear
            # channels (BL/BR) with -3 dB attenuation to preserve spatial balance.
            # Dolby DRC is disabled, and the final output is encoded as DD+ 5.1.
            # Order: pan (downmix) -> static true-peak trim (if any) -> alimiter. $DownmixChain is
            # the SAME graph measured in Phase B. alimiter (limit=0.948) backstops post-pan peaks
            # above -0.47 dBFS (attack=5ms, release=50ms).
            $post      = $t.NeedsPeakSafety ? ",$volTrim" : ""
            $panFilter = "$($script:DownmixChain)${post},alimiter=limit=0.948:attack=5:release=50:level=false:latency=1"

            $ffArgs.AddRange([string[]](
                "-filter:a:$i", $panFilter,
                "-c:a:$i",      "eac3",
                "-ar",          "48000",
                "-b:a:$i",      $t.Bitrate,
                "-dialnorm",    "-31",
                "-cutoff",      "20000",
                "-metadata:s:a:$i", "title=DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
        }
        else {
            # --- 5.1 RE-ENCODE PATH ---
            # This block re-encodes any 5.1 audio track to EAC3 (DD+ 5.1).
            # No downmixing occurs here — input must already be 5.1.
            # Dolby DRC is disabled. Loudness signaling handled by -dialnorm -31.
            # Static trim alone guarantees the ceiling here — NO alimiter (the -ac 6 swr
            # rematrix is clip-protected by default; no channel summation creates new peaks).
            if ($t.NeedsPeakSafety) {
                $ffArgs.AddRange([string[]]("-filter:a:$i", $volTrim))
            }
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
                "-ac","6",
                "-ar","48000",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-cutoff","20000",
                "-metadata:s:a:$i","title=DD+ 5.1 ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 ($($t.Bitrate))$LangTag"
        }

        if ($t.Action -ne "Pass+Copy") {
            $ffArgs.AddRange([string[]]("-metadata:s:a:$i","language=$($t.Language)"))

            $disp = $i -eq 0 ? "default" : "0"
            $ffArgs.AddRange([string[]]("-disposition:a:$i",$disp))

            $i++
        }
    }

    $ffArgs.AddRange([string[]]("-map","0:s?","-c:s","copy"))
    $ffArgs.AddRange([string[]]("-map","0:t?","-c:t","copy"))
    $outDir  = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($InputFile))
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile) + "_keep71mix.mkv"
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
    Write-Error "FFmpeg failed with exit code $LASTEXITCODE. Aborting."
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

    # ======== Pass+Copy -- two logical streams → two summary rows ========
    if ($t.Action -eq "Pass+Copy") {
        # Row 1 — 7.1 passthrough stream (TrueHD or AAC)
        $line1 = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
            $t.Index, $t.Codec, $t.Channels, "Pass",
            "$($t.PassthroughTag) [$($t.Language)]", $t.Priority, $t.Rule
        Write-Host $line1 -ForegroundColor Green

        # Row 2 — DDP 5.1 duplicate stream (indent shows it's a sub-stream)
        $ddpOut = "DD+ 5.1 ($($t.Bitrate)) [$($t.Language)]"
        $line2 = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
            "", "", "", " \-- Copy", $ddpOut, "", ""
        Write-Host $line2 -ForegroundColor Yellow
        continue
    }

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
            $t.Bitrate ? "Encode $($t.Bitrate)" : "Encode"
        }
        "Removed"     { "Removed" }
        default       { $t.Action }
    }

    # Build output label
    $OutputLabel = $t.Output ? $t.Output : "(none)"

    # Show "-" for removed tracks, otherwise priority.
    $PriLabel = $t.Action -eq "Removed" ? "-" : $t.Priority

    # Final formatted line
    $line = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
        $t.Index, $t.Codec, $t.Channels, $ActionLabel, $OutputLabel, $PriLabel, $t.Rule

    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
Write-Host "=== END OF SUMMARY ===" -ForegroundColor Cyan
Write-Host ""