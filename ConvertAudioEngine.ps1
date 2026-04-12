# =====================================================================
#  ConvertAudioEngine.ps1 — (Version 1.vj) Production-daily use
#  PowerShell 5.1 + 7 Compatible
#  FFmpeg 8.1 Compatible
#  Modular Codec Groups
#  Rule table reorganized by codec family
#  Downmix is only used for sources with more than 5.1 channels.
#  Audio 5.1 sources are only re‑encoded.
#  Added Priority Mapping (default 0-100)
# =====================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

# =========================================
#  ENGINE: GLOBAL SETTINGS
# =========================================
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
        }
        return $false
    }

    static [ChannelFilter] $MoreThanTwo = [ChannelFilter]::new("gt", 2)
}

# =============================
#  ENGINE: AUDIO RULE GROUPS
# =============================

# AAC FAMILY
$Rules_AAC = @(
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=8; ProfileRegex=$null
        Action="Downmix"; Bitrate="1024k"; PassthroughTag=$null;
        Rule="AAC_7.1_Downmix_1024k"; Priority=92
    },
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=6; ProfileRegex=$null
        Action="Encode"; Bitrate="768k"; PassthroughTag=$null;
        Rule="AAC_5.1_Encode_768k"; Priority=91
    },
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=2; ProfileRegex=$null
        Action="Encode"; Bitrate="256k"; PassthroughTag=$null;
        Rule="AAC_2.0_Encode_256k"; Priority=90
    }
)

# EAC3 / ATMOS FAMILY
$Rules_EAC3_Atmos = @(
    [PSCustomObject]@{
        CodecRegex="^(eac3)$"; Channels=[ChannelFilter]::new("gt", 2); ProfileRegex="JOC|Atmos"
        Action="Passthrough"; Bitrate=$null; PassthroughTag="EAC3_Atmos_Passthrough";
        Rule="EAC3_Atmos_Passthrough"; Priority=100
    },
    [PSCustomObject]@{
        CodecRegex="^(eac3)$"; Channels=8; ProfileRegex=$null
        Action="Downmix"; Bitrate="1024k"; PassthroughTag=$null;
        Rule="EAC3_7.1_Downmix_1024k"; Priority=99
    },
    [PSCustomObject]@{
        CodecRegex="^(eac3)$"; Channels=6; ProfileRegex=$null
        Action="Passthrough"; Bitrate=$null; PassthroughTag="EAC3_5.1_Passthrough";
        Rule="EAC3_5.1_Passthrough"; Priority=98
    },
    [PSCustomObject]@{
        CodecRegex="^(eac3)$"; Channels=2; ProfileRegex=$null
        Action="Passthrough"; Bitrate=$null; PassthroughTag="EAC3_2.0_Passthrough";
        Rule="EAC3_2.0_Passthrough"; Priority=97
    }
)

# TRUEHD FAMILY
$Rules_TrueHD = @(
    [PSCustomObject]@{
        CodecRegex="^(mlp|truehd|true-hd)$"; Channels=6; ProfileRegex=$null
        Action="Passthrough"; Bitrate=$null; PassthroughTag="TrueHD_5.1_Passthrough";
        Rule="TrueHD_5.1_Passthrough"; Priority=81
    },
    [PSCustomObject]@{
        CodecRegex="^(mlp|truehd|true-hd)$"; Channels=8; ProfileRegex=$null
        Action="Downmix"; Bitrate="1024k"; PassthroughTag=$null;
        Rule="TrueHD_7.1_Downmix_1024k"; Priority=80
    }
)

# DTS FAMILY
$Rules_DTS = @(
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels=[ChannelFilter]::MoreThanTwo; ProfileRegex="HD|MA|HRA"
        Action="Encode"; Bitrate="1024k"; PassthroughTag=$null;
        Rule="DTSHD_Multichannel_Encode_1024k"; Priority=72
    },
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels=[ChannelFilter]::MoreThanTwo; ProfileRegex=$null
        Action="Encode"; Bitrate="768k"; PassthroughTag=$null;
        Rule="DTS_Multichannel_Encode_768k"; Priority=71
    },
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels=2; ProfileRegex=$null
        Action="Encode"; Bitrate="256k"; PassthroughTag=$null;
        Rule="DTS_2.0_Encode_256k"; Priority=70
    }
)

# PCM / FLAC FAMILY
$Rules_PCMFLAC = @(
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=8
        ProfileRegex=$null; Action="Downmix"; Bitrate="1024k"
        PassthroughTag=$null; Rule="PCMFLAC_7.1_Downmix_1024k"; Priority=62
    },
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=6
        ProfileRegex=$null; Action="Encode"; Bitrate="768k"
        PassthroughTag=$null; Rule="PCMFLAC_5.1_Encode_768k"; Priority=61
    },
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=2
        ProfileRegex=$null; Action="Encode"; Bitrate="256k"
        PassthroughTag=$null; Rule="PCMFLAC_2.0_Encode_256k"; Priority=60
    }
)

# ============================================================
#  ENGINE: MERGE ALL AUDIO RULE GROUPS
#  (Adding a new codec = define $Rules_XYZ, then append here)
# ============================================================
$AudioRules = $Rules_EAC3_Atmos + $Rules_TrueHD + $Rules_DTS + $Rules_AAC + $Rules_PCMFLAC

# Priority sort
$i = 0
$AudioRules = $AudioRules |
    ForEach-Object {
        [PSCustomObject]@{ Rule = $_; Index = $i++ }
    } |
    Sort-Object { $_.Rule.Priority } -Descending |
    ForEach-Object { $_.Rule }

# Precompile regex
$AudioRules = $AudioRules | ForEach-Object {
    $r = $_
    $codecVal   = if ($r.CodecRegex)   { [regex]::new($r.CodecRegex,'IgnoreCase') }   else { $null }
    $profileVal = if ($r.ProfileRegex) { [regex]::new($r.ProfileRegex,'IgnoreCase') } else { $null }
    $r | Add-Member -NotePropertyName CodecRegexObj   -NotePropertyValue $codecVal   -PassThru |
         Add-Member -NotePropertyName ProfileRegexObj -NotePropertyValue $profileVal -PassThru
}

# ==========================
#  PROBE
# ==========================
function Get-AudioStreams {
    param([string]$File)

    $probeArgs = @(
        "-v","quiet","-print_format","json",
        "-show_streams","-select_streams","a",$File
    )

    $raw = & $script:ffprobe @probeArgs 2>$null
    if (-not $raw) {
        Write-Error "ffprobe returned no output. Check the input file: $File"
        exit 1
    }
    try { return ($raw | ConvertFrom-Json).streams }
    catch {
        Write-Error "Failed to parse ffprobe JSON: $_"
        exit 1
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
        $Lang     = if ($s.tags -and $s.tags.language) { $s.tags.language } else { "und" }
        $Handler  = if ($s.tags) { $s.tags.handler_name } else { $null }

        $RealIndex = $s.index
        $Rule = $null

        # --- Malformed Layout Guards (include 7ch) ---
        if ($Codec -eq "aac" -and $Channels -eq 7) {
            $Channels = 6
            if (-not $Rule) { $Rule = "AAC_MalformedLayout_Guard" }
        }

        if ($Codec -eq "eac3" -and $Channels -eq 7) {
            $Channels = 6
            if (-not $Rule) { $Rule = "EAC3_MalformedLayout_Guard" }
        }

        if ($Codec -match "^(mlp|truehd|true-hd)$" -and $Channels -eq 7) {
            $Channels = 8
            if (-not $Rule) { $Rule = "TrueHD_MalformedLayout_Guard" }
        }

        if ($Codec -match "^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$" -and $Channels -eq 7) {
            $Channels = 8
            if (-not $Rule) { $Rule = "PCMFLAC_MalformedLayout_Guard" }
        }

        # Commentary removal — now with full schema
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
            elseif ($Channels -gt 6)  { $Action="Encode"; $Bitrate="1024k"; $Rule="Fallback_7.1_1024k" }
            else                      { $Action="Encode"; $Bitrate="768k";  $Rule="Fallback_5.1_768k" }
            $Passthrough=$false; $Downmix=$false; $Tag=$null; $Priority=0
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

    # Global options
    $ffArgs.AddRange([string[]](
        "-threads",$ThreadCount,
        "-analyzeduration","100M",
        "-probesize","100M",
        "-err_detect","ignore_err",
        "-i",$InputFile,
        "-map","0:v",
        "-c:v","copy",
        "-map_chapters","0"
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
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
                "-ac","6",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-cutoff","20000",
                "-metadata:s:a:$i","title=DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
        }
        elseif ($t.Channels -le 2) {
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
                "-ac","2",
                "-b:a:$i",$t.Bitrate,
                "-dialnorm","-31",
                "-dsur_mode","notindicated",
                "-metadata:s:a:$i","title=DD+ 2.0 ($($t.Bitrate))$LangTag"
            ))
            $t.Output = "DD+ 2.0 ($($t.Bitrate))$LangTag"
        }
        else {
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
        "Encode"      { "Cyan" }       # Cyan for encoding (main action)
        "Removed"     { "Red" }        # Red for removed (important)
        default       { "White" }
    }

    # Build action label
    $ActionLabel = switch ($t.Action) {
        "Passthrough" { "Passthrough" }
        "Downmix"     { "Downmix 5.1" }
        "Encode"      {
            if     ($t.Bitrate -eq "256k")  { "Encode 256k" }
            elseif ($t.Bitrate -eq "768k")  { "Encode 768k" }
            elseif ($t.Bitrate -eq "1024k") { "Encode 1024k" }
            else                            { "Encode" }
        }
        "Removed"     { "Removed" }
        default       { $t.Action }
    }

    # Build output label
    $OutputLabel = if ($t.Output) { $t.Output } else { "(none)" }

    # Priority label — show dash for removed/fallback tracks (priority not applicable)
    $PriLabel = if ($t.Action -eq "Removed") { "-" } else { $t.Priority }

    # Final formatted line
    $line = "{0,-4} {1,-8} {2,-5} {3,-16} {4,-40} {5,-5} {6}" -f `
        $t.Index, $t.Codec, $t.Channels, $ActionLabel, $OutputLabel, $PriLabel, $t.Rule

    Write-Host $line -ForegroundColor $Color
}

Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
Write-Host "=== END OF SUMMARY ===" -ForegroundColor Cyan
Write-Host ""
