# =====================================================================
#  ConvertAudioEngine.ps1 — (Version 1.v.f RC)
#  PowerShell 5.1 + 7 Compatible
#  FFmpeg 8.1 Compatible
#  Modular Codec Groups
#  Rule table reorganized by codec family
#  Downmix is only used for sources with more than 5.1 channels.
#  Audio 5.1 sources are only re‑encoded.
#  Added Priority Mapping (default 0-100)
#  Added dialnorm, dsur_mode
# =====================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

# =====================================================================
#  ENGINE: GLOBAL SETTINGS
# =====================================================================
$ThreadCount = 8    # User‑adjustable (4–16). Audio doesn't gain speed from more threads.
$CommentaryKeywords = @("commentary","director","producer","writer","cast","behind","bonus","alt","interview")

# Resolve ffprobe/ffmpeg relative to the script's own directory
# instead of the caller's working directory.
$ffprobe = Join-Path $PSScriptRoot "ffprobe.exe"
$ffmpeg  = Join-Path $PSScriptRoot "ffmpeg.exe"

# =====================================================================
#  ENGINE: AUDIO RULE GROUPS (BY CODEC FAMILY)
#  Goal: Adding new codecs is easy, rule table stays readable.
# =====================================================================

# AAC FAMILY
# - AAC 7.1 → Downmix to 5.1 @ 1024k
# - AAC 5.1 → Encode @ 768k
# - AAC 2.0 → Encode @ 256k
$Rules_AAC = @(
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=8; ProfileRegex=$null
        Action="Downmix"; Bitrate="1024k"; PassthroughTag=$null; Tag=$null
        Rule="AAC_7.1_Downmix_1024k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=6; ProfileRegex=$null
        Action="Encode"; Bitrate="768k"; PassthroughTag=$null; Tag=$null
        Rule="AAC_5.1_Encode_768k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(aac)$"; Channels=2; ProfileRegex=$null
        Action="Encode"; Bitrate="256k"; PassthroughTag=$null; Tag=$null
        Rule="AAC_2.0_Encode_256k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    }
)

# EAC3 / ATMOS FAMILY
# - Atmos (EAC3 JOC) → Passthrough
$Rules_EAC3_Atmos = @(
    [PSCustomObject]@{
        CodecRegex="^(eac3)$"; Channels=$null; ProfileRegex="JOC|Atmos"
        Action="Passthrough"; Bitrate=$null; PassthroughTag="EAC3_Atmos_Passthrough"; Tag=$null
        Rule="EAC3_Atmos_Passthrough"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    }
)

# TRUEHD FAMILY
# - TrueHD 5.1 → Passthrough
# - TrueHD 7.1 → Downmix to 5.1 @ 1024k
$Rules_TrueHD = @(
    [PSCustomObject]@{
        CodecRegex="^(mlp|truehd|true-hd)$"; Channels=6; ProfileRegex=$null
        Action="Passthrough"; Bitrate=$null; PassthroughTag="TrueHD_5.1_Passthrough"; Tag=$null
        Rule="TrueHD_5.1_Passthrough"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(mlp|truehd|true-hd)$"; Channels=8; ProfileRegex=$null
        Action="Downmix"; Bitrate="1024k"; PassthroughTag=$null; Tag=$null
        Rule="TrueHD_7.1_Downmix_1024k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    }
)

# DTS FAMILY (Core, MA, HRA)
# - Stereo DTS → Encode 256k
# - Multichannel DTS → Encode 768k
# - DTS-HD MA/HRA → Encode 1024k
$Rules_DTS = @(
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels=2; ProfileRegex=$null
        Action="Encode"; Bitrate="256k"; PassthroughTag=$null; Tag=$null
        Rule="DTS_2.0_Encode_256k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels={ param($c) $c -gt 2 }; ProfileRegex="HD|MA|HRA"
        Action="Encode"; Bitrate="1024k"; PassthroughTag=$null; Tag="DTSHD"
        Rule="DTSHD_Multichannel_Encode_1024k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(dts)$"; Channels={ param($c) $c -gt 2 }; ProfileRegex=$null
        Action="Encode"; Bitrate="768k"; PassthroughTag=$null; Tag=$null
        Rule="DTS_Multichannel_Encode_768k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    }
)

# PCM / FLAC FAMILY
# - 7.1 → Downmix to 5.1 @ 1024k
# - 5.1 → Encode @ 768k
# - 2.0 → Encode @ 256k
$Rules_PCMFLAC = @(
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=8
        ProfileRegex=$null; Action="Downmix"; Bitrate="1024k"
        PassthroughTag=$null; Tag=$null; Rule="PCMFLAC_7.1_Downmix_1024k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=6
        ProfileRegex=$null; Action="Encode"; Bitrate="768k"
        PassthroughTag=$null; Tag=$null; Rule="PCMFLAC_5.1_Encode_768k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    },
    [PSCustomObject]@{
        CodecRegex="^(pcm_s16le|pcm_s24le|pcm_f32le|pcm_f32be|flac)$"; Channels=2
        ProfileRegex=$null; Action="Encode"; Bitrate="256k"
        PassthroughTag=$null; Tag=$null; Rule="PCMFLAC_2.0_Encode_256k"; Priority=0
        CodecRegexObj=$null; ProfileRegexObj=$null
    }
)

# =====================================================================
#  ENGINE: MERGE ALL AUDIO RULE GROUPS
#  (Adding a new codec = define $Rules_XYZ, then append here)
# =====================================================================
$AudioRules = $Rules_EAC3_Atmos + $Rules_TrueHD + $Rules_DTS + $Rules_AAC + $Rules_PCMFLAC

# =====================================================================
#  ENGINE: PRIORITY SORT
#  Rules are sorted by Priority (highest first).
#  If priorities match, original rule order is preserved.
#  Default Priority is 0; raise it (0-100) to make a rule win.
# =====================================================================
$i = 0
$AudioRules = $AudioRules |
    ForEach-Object {
        [PSCustomObject]@{
            Rule  = $_
            Index = $i++
        }
    } |
    Sort-Object { $_.Rule.Priority } -Descending |
    ForEach-Object { $_.Rule }

# =====================================================================
#  ENGINE: PRECOMPILE REGEX
# =====================================================================
foreach ($r in $AudioRules) {
    $r.CodecRegexObj   = if ($r.CodecRegex)   { [regex]::new($r.CodecRegex,   'IgnoreCase') } else { $null }
    $r.ProfileRegexObj = if ($r.ProfileRegex) { [regex]::new($r.ProfileRegex, 'IgnoreCase') } else { $null }
}

# =====================================================================
#  ENGINE: PROBE FUNCTION
# =====================================================================
function Get-AudioStreams {
    param([string]$File)

    $probeArgs = @(
        "-v","quiet","-print_format","json",
        "-show_streams","-select_streams","a",$File
    )

    # Use $script:ffprobe resolved from $PSScriptRoot (set at top of script)
    $raw = & $script:ffprobe @probeArgs 2>$null
    try { return ($raw | ConvertFrom-Json).streams }
    catch {
        Write-Host "Failed to parse ffprobe JSON." -ForegroundColor Red
        exit
    }
}

# =====================================================================
#  ENGINE: COMMENTARY DETECTOR
# =====================================================================
function Test-IsCommentary {
    param($Channels, $Title)

    if ($Channels -ne 2 -or -not $Title) { return $false }

    $lower = ([string]$Title).ToLower()
    foreach ($kw in $CommentaryKeywords) {
        if ($lower -match $kw) { return $true }
    }
    return $false
}

# =====================================================================
#  ENGINE: RULE MATCHER
# =====================================================================
function Resolve-AudioRule {
    param($Codec, $Channels, $Profile, $Title, $Handler)

    foreach ($r in $AudioRules) {

        if ($r.CodecRegexObj -and -not $r.CodecRegexObj.IsMatch($Codec)) { continue }

        if ($r.Channels -ne $null) {
            if ($r.Channels -is [scriptblock]) {
                if (-not (& $r.Channels $Channels)) { continue }
            } elseif ($Channels -ne $r.Channels) {
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

# =====================================================================
#  ENGINE: PROCESS TRACKS
# =====================================================================
function Process-AudioTracks {
    param($Streams)

    $Processed = @()
    $TrackIndex = 0

    foreach ($s in $Streams) {

        $Codec    = $s.codec_name
        $Channels = $s.channels
        $Profile  = $s.profile

        # --- NULL-SAFE TAG ACCESS  ---
        $Title    = if ($s.tags) { $s.tags.title } else { $null }
        $Lang     = if ($s.tags -and $s.tags.language) { $s.tags.language } else { "und" }
        $Handler  = if ($s.tags) { $s.tags.handler_name } else { $null }

        $RealIndex = $s.index

        # Reset $Rule at the start of every iteration so it does
        # not bleed the previous track's value into the current one
        # when neither malformed-layout guard fires.
        $Rule = $null

        # --- Malformed Layout Guards  ---
        if ($Codec -eq "aac" -and $Channels -gt 6 -and $Channels -lt 8) {
            $Channels = 6
            $Rule = "AAC_MalformedLayout_Guard"
        }

        if ($Codec -eq "eac3" -and $Channels -gt 6 -and $Channels -lt 8) {
            $Channels = 6
            $Rule = "EAC3_MalformedLayout_Guard"
        }

        # Remove stereo commentary tracks
        if (Test-IsCommentary -Channels $Channels -Title $Title) {
            $Processed += [PSCustomObject]@{
                Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
                Profile=$Profile; Title=$Title; Language=$Lang; Action="Removed"
                Output="Commentary"; Rule="Commentary_Removed"
            }
            $TrackIndex++; continue
        }

        # Rule Matching
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

        # --- DTS-HD Bitrate Override  ---
        # Guard with -not $Passthrough so this block cannot
        # accidentally overwrite the bitrate/rule of a future
        if ($match -and $Codec -eq "dts" -and $Channels -gt 2 -and -not $Passthrough) {
            if ($match.Tag -eq "DTSHD") {
                $Bitrate = "1024k"
                $Rule    = "DTSHD_Multichannel_Encode_1024k"
            } else {
                $Bitrate = "768k"
                $Rule    = "DTS_Multichannel_Encode_768k"
            }
        }

        # --- 6/8ch Normalization  ---
        if ($Channels -eq 6 -and $Bitrate -eq "1024k") {
            $Bitrate = "768k"
        }
        elseif ($Channels -eq 8 -and $Bitrate -eq "768k") {
            $Bitrate = "1024k"
        }

        $Processed += [PSCustomObject]@{
            Index=$TrackIndex; RealIndex=$RealIndex; Codec=$Codec; Channels=$Channels
            Profile=$Profile; Title=$Title; Language=$Lang; Action=$Action
            Passthrough=$Passthrough; Downmix=$Downmix; Bitrate=$Bitrate
            PassthroughTag=$Tag; Output=""; Rule=$Rule; Priority=$Priority
        }

        $TrackIndex++
    }

    return $Processed
}

# =====================================================================
#  ENGINE: FFMPEG COMMAND BUILDER
# =====================================================================
function Build-FFmpegCommand {
    param($Tracks, $InputFile, $ThreadCount)

    # Renamed from $args to $ffArgs.
    $ffArgs = New-Object System.Collections.Generic.List[string]
    $ffArgs.AddRange([string[]](
        "-threads",$ThreadCount,
        "-err_detect","ignore_err",
        "-avioflags","direct",
        "-analyzeduration","100M",
        "-probesize","100M",
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
        	"-dsur_mode","0",
        	"-metadata:s:a:$i","title=DD+ 5.1 Downmix ($($t.Bitrate))$LangTag",
        	"-cutoff","20000"
            ))
            $t.Output = "DD+ 5.1 Downmix ($($t.Bitrate))$LangTag"
        }
        elseif ($t.Channels -le 2) {
            # Note: -cutoff is intentionally omitted for stereo encode;
            #       high-frequency limiting is only needed for multichannel downmixes.
            $ffArgs.AddRange([string[]](
                "-c:a:$i","eac3",
        	"-ac","2",
        	"-b:a:$i",$t.Bitrate,
        	"-dialnorm","-31",
        	"-dsur_mode","0",
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
        	"-dsur_mode","0",
        	"-metadata:s:a:$i","title=DD+ 5.1 ($($t.Bitrate))$LangTag",
        	"-cutoff","20000"
            ))
            $t.Output = "DD+ 5.1 ($($t.Bitrate))$LangTag"
        }

        # Language metadata
        $ffArgs.AddRange([string[]]("-metadata:s:a:$i","language=$($t.Language)"))

        # Disposition
        $disp = if ($i -eq 0) { "default" } else { "0" }
        $ffArgs.AddRange([string[]]("-disposition:a:$i",$disp))

        $i++
    }

    $ffArgs.AddRange([string[]]("-map","0:s?","-c:s","copy"))
    $ffArgs.Add("$([System.IO.Path]::GetFileNameWithoutExtension($InputFile))_Processed.mkv")

    return $ffArgs
}

# =====================================================================
#  ENGINE: MAIN EXECUTION
# =====================================================================
Write-Host "=== Probing Audio Streams ===" -ForegroundColor Cyan
$streams = Get-AudioStreams -File $InputFile

Write-Host "=== Processing Tracks ===" -ForegroundColor Cyan
$tracks = Process-AudioTracks -Streams $streams

Write-Host "=== Building FFmpeg Command ===" -ForegroundColor Cyan
$cmd = Build-FFmpegCommand -Tracks $tracks -InputFile $InputFile -ThreadCount $ThreadCount

# Use $ffmpeg resolved from $PSScriptRoot instead of .\ffmpeg.exe
& $ffmpeg @cmd

# =====================================================================
#  SUMMARY ENGINE
# =====================================================================

Write-Host ""
Write-Host "==================== AUDIO PROCESSING SUMMARY ====================" -ForegroundColor Cyan
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
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""
