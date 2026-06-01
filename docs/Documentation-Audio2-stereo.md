#### ConvertAudio2-Stereo.ps1 (v1.4.0)

> Downmixes any 5.1, 7.1, or Atmos track into a high‑quality EAC3 or Opus stereo track.   

- **PreTrack-QRS** — selects the best source using a quality ranking and a clear tie-break order (rank → language → bitrate → index) before any encoding decision is made.  
- **Safe Peak Normalizer (SPN)** — scans first, applies loudnorm only when the source truly needs it, preserving dynamics on clean tracks.  
- **Multichannel Quality Ladder** — ranks formats from TrueHD to AAC
- **Tie-Break Logic** — resolves equal scores using language, bitrate, and index
- **Malformed-layout correction** — fixes invalid 7-channel layouts
- **ITU-R BS.775 pan matrices** — accurate 5.1/7.1 > stereo downmixing.

**PowerShell 7.6** - **[Click here to Download PS 7.6 (Windows x64 .msi)](https://github.com/PowerShell/PowerShell/releases/download/v7.6.1/PowerShell-7.6.1-win-x64.msi)**  
**FFmpeg 8.1** - **https://ffmpeg.org/download.html**

####  Quick Start (Windows 11)

1. Place **`ConvertAudio2-Stereo.ps1`**, **`ffmpeg.exe`**, **`ffprobe.exe`** in the same folder.
2. Run the script:

> pwsh -ExecutionPolicy Bypass -File .\ConvertAudio2-Stereo.ps1 ".\YourMovie.mkv"

The -ExecutionPolicy Bypass only applies to this single run and doesn't change your system settings. 

#### macOS / Linux

1. Place the script, **`ffmpeg`**, and **`ffprobe`** in the same folder.
2. Make the binaries executable (first time only):

> chmod +x ffmpeg ffprobe

3. Run:

> pwsh -File ./ConvertAudio2-Stereo.ps1 "./YourMovie.mkv"

#### 1. Stereo Bitrate Selection

You can switch codecs at runtime by passing -StereoCodec Opus or EAC3 on the command line. Runtime switching is temporary, while changes within the script are permanent.
```powerhshell
pwsh -ExecutionPolicy Bypass -File .\ConvertAudio2-Stereo.ps1 ".\YourMovie.mkv" -StereoCodec Opus
or
pwsh -ExecutionPolicy Bypass -File .\ConvertAudio2-Stereo.ps1 ".\YourMovie.mkv" -StereoCodec EAC3
```
Remember to save file after changing Codecs or Bit-rates.
#### To change audio codec type, within the script. (line 25-27)  

```powershell
# Default stereo codec used when -StereoCodec is not provided.  
# You can change this to "Opus" or "EAC3"
$DefaultStereoCodec = "EAC3"  << Change Codec from "EAC3" to "OPUS"  
```

#### To change audio codec bit-rates, within the script (line 38-40)    

```powershell  
# You can safely adjust these numbers, within "...k"
# EAC3: 224, 256, 320, 384, 448, 512  
# Opus: 192, 224, 256, 320, 384, 448  
"EAC3" = "384k"  << Change your EAC3 Bitrate  
"Opus" = "320k"  << Change your OPUS Bitrate  
```
**EAC3** (384 = transparent 2.0 / 448 = transparent 5.1)  
**OPUS** (192 = transparent stereo / 320 = transparent 5.1)  

LFE fold-in is controlled by **`$FoldLFE`** in the Global Settings block (line 55).  

```powershell 
When set to $true, the LFE channel is folded into the stereo mix at +0.5 gain inside the pan matrix.   
When set to $false, the LFE channel is removed entirely.   
[bool]$FoldLFE = $true  # $false discards LFE
```

**Output filename:** The program saves a new file in the same folder, adding a codec suffix like (_stereo_eac3.mkv) or (_stereo_opus.mkv). Your original file stays untouched. 

<br>   

#### 2. Supported Codec Families

| Family | Formats |
|--------|---------|
| AAC | AAC LC, HE-AAC, AAC 5.1, AAC 7.1 |
| EAC3 / Atmos | EAC3, EAC3-JOC (Atmos) |
| TrueHD | TrueHD 5.1, TrueHD 7.1, TrueHD+Atmos |
| DTS | DTS Core, DTS-HD MA, DTS-HD HRA |
| PCM | pcm_s16le, pcm_s24le, pcm_f32le, pcm_f32be |
| FLAC | All FLAC layouts |

#### 3. FFmpeg Global Options

| Flag | Purpose |
|------|---------|
| **`-threads 8`** | Sets how many CPU threads the encoder uses. (4-16 user-adjustable) |
| **`-analyzeduration 200M`** | Deep probe for TrueHD/Atmos/AAC 7.1 layouts. |
| **`-probesize 200M`** | Prevents mis-detected channel layouts. |
| **`-err_detect ignore_err`** | Allows FFmpeg to continue through minor bitstream errors. |
| **`-drc_scale 0`** | Disables Dolby DRC for consistent loudness. |
| **`-max_muxing_queue_size 14000`** | Prevents muxer stalls on long files. |
| **`-map 0:v? -c:v copy`** | Always copy video losslessly. (no re‑encode). |
| **`-map_metadata 0 -map_chapters 0`** | Preserve metadata and chapters. |

#### 4. ScanThrottle + Threadcount

Sets how many CPU threads the encoder uses. (line 47-49)
```powershell
$ThreadCount     = 8    # User-adjustable (4-16).
```

Limits how many peak-scan tasks run at the same time. 
```powershell 
$ScanThrottle = [Math]::Max(1, [int]([Environment]::ProcessorCount / $ThreadCount))  
Recommended values: 1, 2, 3, or 4, (replace 1,)
```

#### What ScanThrottle does
- Limits **parallel SPN scans** to avoid CPU overload  
- Keeps the system responsive during peak detection  
- Does **not** affect encoding speed  

**Parallel SPN scan** — scans all tracks at the same time to detect peaks quickly.

---

#### 5. Safe Peak Normalizer (SPN)

SPN is a pre-filter safety system that checks each audio track for unsafe peaks before any downmix or encode occurs.

#### What SPN does
- Scans each non‑Removed track using **`volumedetect`**  
- Compares the detected peak against **−0.5 dBFS**  
- Marks the track with **NeedsNormalization** when the peak is too high  
- Inserts a light **`loudnorm`** step only when required  
- **Clean sources** are Left untouched. Original loudness and dynamics are preserved.  
- **Dirty sources** are Corrected safely without compression artifacts or “over-normalized” sound.  
- Skips passthrough and commentary tracks entirely  

#### 5.5 SPN Behavior

| Condition          | Behavior                               |
|-------------------|------------------------------------------|
| Peak is **greater than** −0.5 dBFS  | Apply **`loudnorm`** before pan matrix       |
| Peak is **less than or equal to** −0.5 dBFS  | No normalization                        |
| Passthrough track | SPN skipped                             |
| Commentary track  | SPN skipped                             |

#### 5.7 SPN Thresholds

- Peak threshold: −0.5 dBFS (line 52)
- Loudnorm target: −23 LUFS
- True peak ceiling: −1.5 dBTP
- Valid range: −1.5 dBFS to −0.3 dBFS (recommended: −0.5 dBFS)

 The threshold can be adjusted via **`$PeakThresholdDB`** in the Global Settings block. (line 52)
 
 ```powershell
 $PeakThresholdDB = -0.5 # SPN: loudnorm triggers when source peak exceeds this (dBFS)
```

#### 6. QRS Integration
(QRS) stands for **Quality**, **Resolution**, and **Scoring**.  
(SPN) stands for **Safe Peak Normalizer.**  

- **SPN runs before QRS**, scanning all non‑Removed tracks so peak and bitrate data are available for evaluation.
- SPN provides **peak**, **bitrate**, and **NeedsNormalization** flags to the QRS Quality, Resolution, and Scoring stages.
- **NeedsNormalization** is set from SPN peak data and controls whether the FFmpeg chain applies loudness correction.

<br>

#### 6.5 PreTrack-QRS

PreTrack‑QRS decides which single audio track becomes the stereo source.
It evaluates all candidates before any downmix or encode occurs.

#### How PreTrack-QRS works.

1. **Quality (Q)**  
- Identifies all multichannel candidates and ranks them using the Quality Ladder  
- (TrueHD > PCM/FLAC >> DTS‑HD >>> EAC3 >>>> AC3 >>>>> AAC).  
- Stereo codecs are ignored. 

2. **Resolution of ties (R)**  
- If two tracks share the same QualityRank, QRS applies a fixed chain:
- Language (eng > und >> others)  
- Higher bitrate  
- Lower stream index  

3. **Scoring (S)**  
- Applies structured checks (metadata, layout, profile) to finalize the winner.  

#### After QRS completes:

- All non‑winning tracks are marked **Removed** 
- Only the single best track proceeds to encoding  
- The script always produces exactly one stereo output

#### 7. Multichannel Audio Ladder (Quality)

Higher = better. The Quality Ladder defines the fixed ranking used by QRS.
| Score | Format / Channels       | Score | Format / Channels     |
|-------|--------------------------|--------|------------------------|
| 100   | TrueHD Atmos 7.1/Atmos  | 80     | DTS Core 5.1          |
| 99    | TrueHD 7.1              | 75     | EAC3 Atmos (JOC) 5.1/7.1 |
| 98    | TrueHD 5.1              | 70     | EAC3 7.1               |
| 95    | PCM/FLAC 7.1            | 69     | EAC3 5.1               |
| 94    | PCM/FLAC 5.1            | 60     | AC3 5.1                |
| 90    | DTS-HD MA 7.1           | 50     | AAC 7.1                |
| 89    | DTS-HD MA 5.1           | 49     | AAC 5.1                |
| 85    | DTS-HRA 5.1/7.1         |        |                        |

<div style="page-break-after: always;"></div>

#### 8. ITU-R BS.775 Pan-Matrix Routing  

The audio-engine uses ITU‑R BS.775 pan matrices for all multichannel >> stereo downmixing.

#### 5.1 > 2.0 Matrix

```powershell
pan=stereo|
FL=FL+0.707*FC+0.707*BL+0.707*SL+0.5*BC+0.5*LFE|
FR=FR+0.707*FC+0.707*BR+0.707*SR+0.5*BC+0.5*LFE
```

#### 7.1 > 2.0 Matrix

```powershell
aformat=channel_layouts=7.1,pan=stereo|
FL=FL+0.707*FC+0.707*BL+0.5*SL+0.5*LFE|
FR=FR+0.707*FC+0.707*BR+0.5*SR+0.5*LFE
```

#### Why these matrices matter

- **0.707** = −3 dB fold for center/surround channels (ITU‑R standard)   
- **0.5** = LFE fold-in (optional via `$FoldLFE`)  
- **aformat** ensures consistent 7.1 layout before Downmixing. 
- **Limiter** prevents true-peak overshoots  

> $alimiter = "alimiter=limit=0.948:attack=5:release=50:level=disabled:latency=1"

#### 9. EAC3 Conditional Passthrough

EAC3 2.0 is **only** passed through when:

1. The rule engine marks it as Passthrough  
2. The track has exactly 2 channels  
3. The user selected **`-StereoCodec EAC3`**  

If any condition fails, the track is **re-encoded** to the chosen stereo codec.

#### 10. Commentary Detection

Tracks are marked as Commentary when their metadata matches the commentary pattern:

Director, producer, writer, cast, behind, bonus, alt, interview, commentary,  

Commentary tracks are excluded from QRS selection and cannot be chosen as the primary stereo source.

#### 11. Malformed-Layout Guards  
#### (Fixes illegal 7-channel layouts)

Some sources report invalid 7‑channel layouts (AAC, EAC3, TrueHD, FLAC/PCM).

- When a malformed 7‑channel layout is detected, the engine replaces it with a valid layout before QRS and downmixing. 
- This prevents incorrect channel routing and ensures the pan matrix receives a consistent layout.

#### 12. Fallback Logic

If no valid stereo candidate is found after QRS filtering, the engine falls back to the highest‑ranked non‑commentary track.    
This ensures a consistent output when all stereo options are invalid or excluded.

#### 13. Summary Report

At the end of processing, the audio-engine prints a detailed summary table.

This table includes:

| Column | Description |
|--------|-------------|
| **Index** | Internal track index (0-based). |
| **Codec** | Source codec (TrueHD, DTS, AAC, etc.). |
| **Channels** | Corrected channel count (after malformed-layout guard). |
| **Action** | Removed / Encode / Downmix / Passthrough. |
| **Output** | Final codec + bitrate + metadata tag. |
| **Priority** | Rule priority (0-100). |
| **Rule** | The rule that matched (or fallback). |

Each row in the summary table is color-coded by action:  **Color->Action**  
Red->Removed, Green->Passthrough, Dark Cyan->Downmix, Blue->Encode  

<div style="page-break-after: always;"></div>

#### 14. Non-Critical FFmpeg Warnings (Safe to Ignore)

FFmpeg may print warnings during probing or decoding.  
These **do not** affect audio quality or sync.

#### Common warnings (safe to ignore)

| Warning | Meaning | Impact |
|---------|---------|--------|
| quant_step_size larger than huff_lsbs | TrueHD decoder diagnostic | Safe |
| Codec AVOption drc_scale ... has not been used | DRC disabled intentionally | Safe |
| Subtitle: hdmv_pgs_subtitle unspecified size | PGS metadata missing | Safe |
| Assuming an incorrectly encoded 7.1 channel layout | AAC 7.1 variant | Safe (layout is corrected) |
| Could not find codec parameters for stream (pgssub) | PGS subtitle track has incomplete metadata | Safe |
| Consider increasing analyzeduration/probesize | Generic suggestion caused by incomplete PGS metadata | Safe |
| unspecified size (PGS) | Subtitle bitmap dimensions not declared | Safe |

<div style="page-break-after: always;"></div>

#### 15. Audio-Script Rules That Must Not Change

#### a. Malformed-layout guards

AAC 7ch >> 6ch
EAC3 7ch >> 6ch
TrueHD 7ch >> 8ch
PCM/FLAC 7ch >> 8ch

#### b. Safe Peak Normalizer before PreTrack-QRS
Peak scanning completes before ranking begins.

#### c. Ranking table order
The Quality Ladder is evaluated top‑down. The first match is selected.

#### d. Tie-break sequence
QualityRank >> Language >> Bitrate >> Index.

#### e. Pan-matrix operator
Pan matrices use the '=' operator to avoid auto-normalization.

#### f. Limiter position
The limiter is applied last in the audio chain.

#### g. EAC3 passthrough conditions
EAC3 2.0 is passed through only when StereoCodec=EAC3.

#### h. Commentary detection
Commentary detection applies only to 2‑channel tracks.

#### i. Fallback rules must remain intact
Guarantees safe behavior on unknown codecs.

#### Disclaimer
- Always test on a few files before batch-processing your library.  
- Keep backups of original media.  
- Results may vary depending on source quality and hardware.  

