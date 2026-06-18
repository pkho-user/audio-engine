# ========================================================================
#  ConvertAudio2-DDP51 v1.0.8 — Production Use
#  PowerShell 7.6, FFmpeg 8.1
#
#  Removes low-bitrate audio and downmixes the remaining high-quality audio track.
#  Foundation: ConvertAudioEngine-DDP51 (3-phase A/B/C SPN architecture)
#  Includes Safe Peak Normalizer (SPN) and malformed-layout guards.
#
#  Input: any multichannel MKV; removes 2.0/stereo tracks
#  Output suffix: _ddp5only.mkv
# ========================================================================
#Requires -Version 7.6

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputFile
)

# ========================================
#  ENGINE: GLOBAL SETTINGS
# ========================================
$ThreadCount     = 8    # User-adjustable (4-16).
if ($ThreadCount -lt 4 -or $ThreadCount -gt 16) { throw "ThreadCount must be between 4 and 16." }
$ScanThrottle    = [Math]::Max(1, [int]([Environment]::ProcessorCount / $ThreadCount))
                        # Parallel SPN scan concurrency limit.
                        # Formula: ProcessorCount / ThreadCount (prevents CPU saturation).
                        # SSD/NVMe: safe to override up to 4. Spinning disk: keep at 1.
$TruePeakCeilingDB = -1.0  # SPN true-peak safety ceiling (dBTP). Tracks whose
                           # (downmixed) measured true peak exceeds this get a
                           # static, attenuation-only trim back down to it.
                           # Gain is always clamped to <= 0 dB — we never boost.
                           # The alimiter backstop sits below this at ~-0.46 dBFS.
$CommentaryPattern = [regex]::new(
    'commentary|director|producer|writer|cast|behind|bonus|alt|interview',
    'IgnoreCase,Compiled'
)

# ========================================================================
#  ENGINE: DOWNMIX FILTERGRAPH — SINGLE SOURCE OF TRUTH
#  Used in lock-step by Phase-B measurement AND Build-FFmpegCommand.
#  Any divergence means the safety trim is computed against the wrong signal.
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
    Downmix51   = '1152k'   # 7.1 → 5.1 downmix
    ReEncode51  = '768k'    # 5.1 source re-encode
    EncodeDTSHD = '1024k'   # DTS-HD MA/HRA lossless multichannel encode (5.1 and 7.1)
}

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
        return $false
    }

    static [ChannelFilter] $MoreThanTwo = [ChannelFilter]::new("gt", 2)
}

# ============================
#  ENGINE: AUDIO RULE GROUPS
# ============================

# AAC FAMILY
$Rules_AAC = @(
    # Rule1: AAC 7.1 → Downmix to 5.1
    [PSCustomObject]@{
        CodecRegex     = "^(aac)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = $BitRateConfig.Downmix51
        PassthroughTag = $null
        Rule           = "AAC_7.1_Downmix_$($BitRateConfig.Downmix51)"
        Priority       = 92
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
    # Rule2: TrueHD 7.1 → Downmix to 5.1
    [PSCustomObject]@{
        CodecRegex     = "^(mlp|truehd|true-hd)$"
        Channels       = 8
        ProfileRegex   = $null
        Action         = "Downmix"
        Bitrate        = $BitRateConfig.Downmix51
        PassthroughTag = $null
        Rule           = "TrueHD_7.1_Downmix_$($BitRateConfig.Downmix51)"
        Priority       = 80
    }
)

# DTS FAMILY
$Rules_DTS = @(
    # DTS-HD MA/HRA Multichannel (5.1 and 7.1) → Encode
    # MoreThanTwo (gt 2) intentionally covers both 6ch and 8ch DTS-HD.
    # DTS-HD MA/HRA is a lossless source — EncodeDTSHD is applied regardless of
    # whether the input is 5.1 or 7.1. This is a deliberate design choice;
    # DTS Core multichannel gets ReEncode51 (Rule2) as a separate lower-tier rule.
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
    # Rule2: DTS Multichannel → Encode
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
$AudioRules = $AudioRules | Sort-Object { $_.Priority } -Descending

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
        "-analyzeduration","200M",
        "-probesize","200M",
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

# ===========================
#  COMMENTARY DETECTOR
# ===========================
function Test-IsCommentary {
    param(
        [int]    $Channels,
        [string] $Title
    )

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
        Rule=$Rule; Priority=0
        NeedsSpnScan=$false
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

# =============================
#  PRETRACK-QRS (LANGUAGE-GROUPED)
# =============================
# Picks highest-Priority track per language group; marks losers Removed before Phase B.
# Tie-break: equal Priority → lower RealIndex wins. Already-Removed tracks are excluded.
function Select-WinnerPerLanguage {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[object]]$PendingTracks
    )

    $candidates = @($PendingTracks | Where-Object { $_.Action -ne 'Removed' })
    if ($candidates.Count -eq 0) { return }

    $langGroups = $candidates | Group-Object -Property { $_.Language.ToLower() }

    foreach ($g in $langGroups) {
        # Priority DESC, then RealIndex ASC (lower RealIndex wins tie).
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
            $l.Action       = 'Removed'
            $l.Passthrough  = $false
            $l.Downmix      = $false
            $l.Output       = 'QRS_Removed'
            $l.Rule         = 'PreTrack_QRS_Removed'
            $l.NeedsSpnScan = $false
            $l.Priority     = 0

            Write-Host ("[PreTrack-QRS] Removed: Track {0} ({1}, {2}ch, {3})" -f `
                $l.Index, $l.Codec, $l.Channels, $l.Language) -ForegroundColor Red
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

    # ----------------------------------------------------------------
    # PHASE A - Sequential
    # NeedsSpnScan marks which tracks must run a peak scan during the parallel Phase B stage.
    # ----------------------------------------------------------------
    $PendingTracks = [System.Collections.Generic.List[object]]::new()

    foreach ($s in $Streams) {

        $Codec    = $s.codec_name
        $Channels = [int]$s.channels
        $Profile  = $s.profile

        $Title    = $s.tags?.title
        $Lang     = if ($s.tags?.language) { $s.tags.language } else { "eng" }
        $Handler  = $s.tags?.handler_name

        $RealIndex = $s.index
        $Rule = $null

        # --- Centralized Malformed-Layout Guard Module ---
        $guard = Apply-MalformedLayoutGuards -Codec $Codec -Channels $Channels -RealIndex $RealIndex
        $Channels = $guard.Channels
        if ($guard.Rule) { $Rule = $guard.Rule }

        if ($guard.SkipTrack) {
            $removed = New-RemovedTrack $TrackIndex $RealIndex $Codec $Channels $Profile $Title $Lang $Rule
            $PendingTracks.Add($removed)
            $TrackIndex++; continue
        }

        # --- Commentary removal ---
        if (Test-IsCommentary -Channels $Channels -Title $Title) {
            $PendingTracks.Add([PSCustomObject]@{
                Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
                Profile=$Profile; Title=$Title; Language=$Lang
                Action="Removed"; Passthrough=$false; Downmix=$false
                Bitrate=$null; PassthroughTag=$null; Output="Commentary"
                Rule="Commentary_Removed"; Priority=0
                NeedsSpnScan=$false
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
            if ($Channels -le 2)      { $Action="Removed"; $Bitrate=$null;  $Rule="Fallback_2.0_Removed" }
            elseif ($Channels -gt 6)  { $Action="Downmix"; $Bitrate=$BitRateConfig.Downmix51;  $Rule="Fallback_7.1_$($BitRateConfig.Downmix51)" }
            else                      { $Action="Encode";  $Bitrate=$BitRateConfig.ReEncode51; $Rule="Fallback_5.1_$($BitRateConfig.ReEncode51)" }
            $Passthrough=$false; $Downmix=($Action -eq "Downmix"); $Tag=$null; $Priority=0
        }

        # --- Safety Audit ---
        # DTS-HD MA/HRA at $BitRateConfig.EncodeDTSHD won't match $BitRateConfig.Downmix51, so the 6ch clamp won't fire on legitimate DTS-HD 5.1 tracks.
        if ($Channels -eq 6 -and $Bitrate -eq $BitRateConfig.Downmix51) {
            Write-Warning ("Track {0} ({1}): 6ch rule had {2}-clamped to {3}." -f $TrackIndex,$Codec,$BitRateConfig.Downmix51,$BitRateConfig.ReEncode51)
            $Bitrate = $BitRateConfig.ReEncode51
        }
        elseif ($Channels -eq 7 -and $Bitrate -eq $BitRateConfig.ReEncode51) {
            Write-Warning ("Track {0} ({1}): 7ch rule had {2}-clamped to {3}." -f $TrackIndex,$Codec,$BitRateConfig.ReEncode51,$BitRateConfig.Downmix51)
            $Bitrate = $BitRateConfig.Downmix51
        }
        elseif ($Channels -eq 8 -and $Bitrate -eq $BitRateConfig.ReEncode51) {
            Write-Warning ("Track {0} ({1}): 8ch rule had {2}-clamped to {3}." -f $TrackIndex,$Codec,$BitRateConfig.ReEncode51,$BitRateConfig.Downmix51)
            $Bitrate = $BitRateConfig.Downmix51
        }

        $PendingTracks.Add([PSCustomObject]@{
            Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
            Profile=$Profile; Title=$Title; Language=$Lang; Action=$Action
            Passthrough=$Passthrough; Downmix=$Downmix; Bitrate=$Bitrate
            PassthroughTag=$Tag
            Output=($Action -eq "Removed" ? "2.0_Removed" : "")
            Rule=$Rule; Priority=$Priority
            NeedsSpnScan=($Action -ne "Passthrough" -and $Action -ne "Removed")
        })

        $TrackIndex++
    }

    # -------------------------------------------------------------------
    # PRETRACK-QRS - Language-grouped winner selection
    # Runs after Phase A so all rule matches and malformed-layout guards
    # are settled; runs before Phase B so SPN does not waste time scanning
    # tracks that QRS will kill.
    # -------------------------------------------------------------------
    Select-WinnerPerLanguage -PendingTracks $PendingTracks

    # -------------------------------------------------------------------
    # PHASE B - Parallel SPN true-peak scans
    # -------------------------------------------------------------------
    # Downmix tracks are measured through $DownmixChain (identical to the encoder path)
    # so input_tp reflects the 5.1 output peak, not the 7.1 source.
    # Runspace contract: block only READs $t; writes go to ConcurrentDictionary/ConcurrentBag only.
    $spnCount = @($PendingTracks | Where-Object { $_.NeedsSpnScan }).Count
    if ($spnCount -gt 0) {
        Write-Host "[SPN] True-peak scanning $spnCount track(s) — may take a few minutes..." -ForegroundColor Cyan
    }

    $PeakResults = [System.Collections.Concurrent.ConcurrentDictionary[int,object]]::new()
    $ScanOutput  = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    $throttleLimit = $ScanThrottle  
    $PendingTracks | Where-Object { $_.NeedsSpnScan } | ForEach-Object -Parallel {

        $t            = $_
        $ffmpegBin    = $using:ffmpeg
        $threadCount  = $using:ThreadCount
        $inputFile    = $using:InputFile
        $peakDict     = $using:PeakResults
        $outputBag    = $using:ScanOutput
        $downmixChain = $using:DownmixChain

        $label    = "Track $($t.RealIndex) ($($t.Codec), $($t.Channels)ch)"
        $messages = [System.Collections.Generic.List[object]]::new()
        
        $messages.Add([PSCustomObject]@{
            Text  = "[SPN] Measuring true peak: $label - may take a few minutes..."
            Color = "Cyan"
        })

        # Downmix tracks measured through $DownmixChain (same as encoder). >6ch "Encode" tracks
        # (DTS-HD/DTS-core) are measured through the SAME -ac 6 reduction Build applies (aresample
        # to 5.1 uses identical libswresample rematrix), so input_tp reflects the 6ch encoded signal.
        # Genuine 6ch sources are already 5.1, so they remain measured as-is.
        $measureFilter =
            if     ($t.Downmix)        { "$downmixChain,loudnorm=print_format=json" }
            elseif ($t.Channels -gt 6) { "aresample=ochl=5.1,loudnorm=print_format=json" }
            else                       { "loudnorm=print_format=json" }

        $raw = & $ffmpegBin -analyzeduration 200M -probesize 200M `
               -threads $threadCount -drc_scale 0 -i $inputFile `
               -map "0:$($t.RealIndex)" -filter:a $measureFilter -f null - 2>&1 `
               | ForEach-Object { "$_" }

        $rawText = $raw -join "`n"

        $i_val  = if ($rawText -match '"input_i"\s*:\s*"([^"]+)"')  { $Matches[1] } else { $null }
        $tp_val = if ($rawText -match '"input_tp"\s*:\s*"([^"]+)"') { $Matches[1] } else { $null }

        if ($tp_val) {
            $stats = [PSCustomObject]@{
                I  = $i_val
                TP = $tp_val
            }
            [void]$peakDict.TryAdd($t.RealIndex, $stats)
            
            $messages.Add([PSCustomObject]@{
                Text  = "[SPN] $label - Measured true peak: $tp_val dBTP (integrated $i_val LUFS)"
                Color = "Green"
            })
        } else {
            # Measurement failed — gainDb stays 0 (no trim). Never boosts; alimiter backstops.
            $messages.Add([PSCustomObject]@{
                Text  = "[SPN] $label - True-peak detection failed - no static trim, alimiter will backstop"
                Color = "Red"
            })
        }

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

    # -------------------------------------------------------------------
    # PHASE C - Sequential merge + deterministic true-peak safety decision
    # -------------------------------------------------------------------
    # gainDb = min(0, ceiling - measured_tp) — attenuate-only; never boosts.
    # input_tp parsed with InvariantCulture (loudnorm JSON always uses '.'; comma-decimal locales would mis-read it).
    foreach ($t in $PendingTracks) {
        $stats   = $null
        $peakTP  = $null
        $gainDb  = 0.0

        if ($t.NeedsSpnScan) {
            if ($PeakResults.TryGetValue($t.RealIndex, [ref]$stats) -and $stats -and $stats.TP) {
                $parsedTP = 0.0
                if ([double]::TryParse(
                        $stats.TP,
                        [System.Globalization.NumberStyles]::Float,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [ref]$parsedTP) -and [double]::IsFinite($parsedTP)) {
                    $peakTP = $parsedTP
                    $gainDb = [Math]::Min(0.0, $script:TruePeakCeilingDB - $peakTP)
                }
                # else: non-finite/unparseable TP (e.g. "-inf"/"inf"/"nan") -> gainDb stays 0; no trim, alimiter backstops.
            }
            # else: measurement missing -> gainDb stays 0; signal passes at native level.
        }

        $Processed.Add([PSCustomObject]@{
            Index=$t.Index; RealIndex=$t.RealIndex; Codec=$t.Codec; Channels=$t.Channels
            Profile=$t.Profile; Title=$t.Title; Language=$t.Language; Action=$t.Action
            Passthrough=$t.Passthrough; Downmix=$t.Downmix; Bitrate=$t.Bitrate
            PassthroughTag=$t.PassthroughTag; Output=$t.Output; Rule=$t.Rule
            Priority=$t.Priority
            MeasuredTP=$peakTP
            PeakSafetyGainDb=$gainDb
            NeedsPeakSafety=($gainDb -lt 0.0)
        })
    }

    return $Processed
}

# =============================
#  FFMPEG COMMAND BUILDER
# =============================
function Build-FFmpegCommand {
    param(
        $Tracks,
        [string] $InputFile,
        [int]    $ThreadCount
    )

    $ffArgs   = [System.Collections.Generic.List[string]]::new()

    $ffArgs.AddRange([string[]](
        "-y",
        "-loglevel",             "error",
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

        $ffArgs.AddRange([string[]]("-map", "0:$($t.RealIndex)"))

        # Static attenuation-only trim (always <= 0 dB). InvariantCulture prevents comma-decimal locale from
        # producing an unparseable "volume=-0,7dB". Emitted only when a trim is needed.
        $volTrim = ""
        if ($t.NeedsPeakSafety) {
            $g = $t.PeakSafetyGainDb.ToString('0.0##', [System.Globalization.CultureInfo]::InvariantCulture)
            $volTrim = "volume=${g}dB"
        }

        $LangTag = " [$($t.Language)]"

        if ($t.Passthrough) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i",          "copy",
                "-metadata:s:a:$i", "title=$($t.PassthroughTag)$LangTag"
            ))
            $t.Output = "$($t.PassthroughTag)$LangTag"
        }
        elseif ($t.Downmix) {
            # --- 7.1 → 5.1 DOWNMIX PATH ---
            # $DownmixChain is the SAME aformat+pan used by the Phase-B
            # measurement, so the trim below is computed from this exact signal.
            # Order is: downmix (pan) -> static true-peak trim (if any) -> limiter.
            # The trim sits AFTER the pan because side->rear summation is what
            # can push the peak up; the alimiter is the final sample-peak backstop.
            $post      = $t.NeedsPeakSafety ? ",$volTrim" : ""
            $panFilter = "$($script:DownmixChain)${post},alimiter=limit=0.948:attack=5:release=50:level=false:latency=1"

            $ffArgs.AddRange([string[]](
                "-filter:a:$i",     $panFilter,
                "-c:a:$i",          "eac3",
                "-ar",              "48000",
                "-b:a:$i",          $t.Bitrate,
                "-dialnorm",        "-31",
                "-cutoff",          "20000",
                "-metadata:s:a:$i", "title=DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
        }
        else {
            # --- 5.1 RE-ENCODE PATH ---
            # Re-encodes 5.1 input to EAC3 (DD+ 5.1). 6.1 input (DTS 7ch
            # after Safety Audit) also routes here — downmixed to 5.1 via -ac 6,
            # whose swr rematrix is clip-protected by default. Dolby DRC is
            # disabled; level signaling is left at -dialnorm -31 (no shift).
            # Static trim alone guarantees the true-peak ceiling — no limiter is
            # needed on this path (no channel summation to create new peaks).
            if ($t.NeedsPeakSafety) {
                $ffArgs.AddRange([string[]]("-filter:a:$i", $volTrim))
            }
            $ffArgs.AddRange([string[]](
                "-c:a:$i",          "eac3",
                "-ac",              "6",
                "-ar",              "48000",
                "-b:a:$i",          $t.Bitrate,
                "-dialnorm",        "-31",
                "-cutoff",          "20000",
                "-metadata:s:a:$i", "title=DD+ 5.1 ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 ($($t.Bitrate))$LangTag"
        }

        $ffArgs.AddRange([string[]]("-metadata:s:a:$i", "language=$($t.Language)"))

        $disp = $i -eq 0 ? "default" : "0"
        $ffArgs.AddRange([string[]]("-disposition:a:$i", $disp))

        $i++
    }

    $ffArgs.AddRange([string[]]("-map", "0:s?", "-c:s", "copy"))
    $ffArgs.AddRange([string[]]("-map", "0:t?", "-c:t", "copy"))
    $outDir  = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($InputFile))
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile) + "_ddp5only.mkv"
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

    # Color Selection
    $Color = switch ($t.Action) {
        "Passthrough" { "Green" }
        "Downmix"     { "Yellow" }
        "Encode"      { "Blue" }
        "Removed"     { "Red" }
        default       { "White" }
    }

    # Build action label
    $ActionLabel = switch ($t.Action) {
        "Passthrough" { "Passthrough" }
        "Downmix"     { "Downmix 5.1" }
        "Encode"      {
            if     ($t.Bitrate -eq $BitRateConfig.ReEncode51)  { "Encode $($BitRateConfig.ReEncode51)"  }
            elseif ($t.Bitrate -eq $BitRateConfig.Downmix51)   { "Encode $($BitRateConfig.Downmix51)"   }
            elseif ($t.Bitrate -eq $BitRateConfig.EncodeDTSHD) { "Encode $($BitRateConfig.EncodeDTSHD)" }
            elseif ($t.Bitrate)                                { "Encode $($t.Bitrate)"                  }
            else                                                { "Encode" }
        }
        "Removed"     { "Removed" }
        default       { $t.Action }
    }

    $OutputLabel = $t.Output ? $t.Output : "(none)"

    $PriLabel = $t.Action -eq "Removed" ? "-" : $t.Priority

    $line = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
        $t.Index, $t.Codec, $t.Channels, $ActionLabel, $OutputLabel, $PriLabel, $t.Rule

    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
Write-Host "=== END OF SUMMARY ===" -ForegroundColor Cyan
Write-Host ""