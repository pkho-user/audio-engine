### 1. ConvertAudioEngine-DDP51.ps1 v3.0.5
> Converts all audio tracks (7.1 → downmix, 5.1 → re-encode, passthrough where applicable).  
> Only stereo and commentary tracks are removed.

- Downmixes **7.1 audio** to **DDP 5.1 (1152k)** via ITU-R BS.775 pan matrix
- Re-encodes **5.1 audio** to **DDP 5.1 (768k)**; DTS-HD MA/HRA re-encodes at **1024k**
- Passthrough for **EAC3 Atmos**, **EAC3 5.1**, and **TrueHD 5.1**
- All **2.0 stereo tracks are removed** — stereo output is handled by ConvertAudioEngine-Stereo
- Maintains perfect sync on long-duration files
- Normalizes malformed channel layouts
- Removes 2-channel commentary tracks

### Requirements

**PowerShell 7.6** - **[Click here to Download PS 7.6 (Windows x64 .msi)](https://github.com/PowerShell/PowerShell/releases/download/v7.6.1/PowerShell-7.6.1-win-x64.msi)**  
**FFmpeg 8.1** - **https://ffmpeg.org/download.html**

### Quick Start (Windows 11)

1. Place **`ConvertAudioEngine-DDP51.ps1`**, **`ffmpeg.exe`**, and **`ffprobe.exe`** in the same folder.
2. Run the script:

> pwsh -ExecutionPolicy Bypass -File .\ConvertAudioEngine-DDP51.ps1 ".\YourMovie.mkv"

The -ExecutionPolicy Bypass only applies to this single run and doesn't change your system settings. 

### macOS / Linux

1. Place the script, **`ffmpeg`**, and **`ffprobe`** in the same folder.
2. Make the binaries executable (first time only):

> chmod +x ffmpeg ffprobe

3. Run:

> pwsh -File ./ConvertAudioEngine-DDP51.ps1 "./YourMovie.mkv"

<div style="page-break-after: always;"></div> 

#### 1 To Adjust Bitrate (line 66-70)

```powershell
You can safely adjust these numbers, within **'....k'** 
Common EAC3 5.1 bitrates: 384, 448, 512, 640, 768 
Higher pipeline bitrates: 1024, 1152, 1280, 1408, 1536

$BitRateConfig = @{
    Downmix51   = '1152k'   # << Change your Bitrate
    ReEncode51  = '768k'    # << Change your Bitrate
    EncodeDTSHD = '1024k'   # << Change your Bitrate
    # PassCopy51 entries added in Keep71 script
```

**Remember to save the file after making changes.**  
All three values can be adjusted safely — changes apply automatically throughout the script.

| Key | Default | Used for |
|-----|---------|----------|
| **Downmix51** | 1152k | All 7.1 sources downmixed to 5.1 |
| **ReEncode51** | 768k | 5.1 sources re-encoded to DD+ |
| **EncodeDTSHD** | 1024k | DTS-HD MA and HRA sources (5.1 and 7.1) |

### 🎧 Supported Audio Codecs

The engine supports and processes the following codec families:

| Codec Family | Formats Included |
|--------------|------------------|
| **AAC** | AAC LC, HE-AAC, AAC 5.1, AAC 7.1 |
| **EAC3 / Atmos** | EAC3, EAC3-JOC (Atmos) |
| **TrueHD** | TrueHD 5.1, TrueHD 7.1, TrueHD+Atmos |
| **DTS** | DTS Core, DTS-HD MA, DTS-HD HRA |
| **PCM** | pcm_s16le, pcm_s24le, pcm_f32le, pcm_f32be |
| **FLAC** | All FLAC channel layouts |

<div style="page-break-after: always;"></div>

### 2. FFmpeg Global Options
These are global flags to ensure FFmpeg handles large, high-resolution files without data corruption or sync drift.

* **`-y`** **:** Overwrites the output file without asking.
* **`-loglevel error`** **:** Suppresses all FFmpeg output except actual errors.
* **`-loglevel warning`** **:** Shows warnings and errors, but hides info, debug, and trace output.
* **`-analyzeduration 200M`** & **`-probesize 200M`**
    * **What:** Increases the initial data analysis buffer to 200 Megabytes.
    * **Why:** High-bitrate audio like TrueHD with Atmos metadata or AAC 7.1 requires a deep probe to fully parse the channel layout.   Standard 5MB, 10MB, 20MB probes often misidentify these tracks, causing the encode to fail or use the wrong channel map.
* **`-err_detect ignore_err`**
    * **What:** Instructs FFmpeg to bypass minor bitstream errors.
    * **Why:** Prevents a single corrupted audio frame (common in digital rips) from aborting a multi-hour conversion process.
* **`-drc_scale 0`**
    * **What:** Disables Dynamic Range Compression (DRC).
    * **Why:** Ensures the output audio maintains the full dynamic range of the source. It prevents the encoder from "flattening" the sound or lowering the volume of loud peaks.

### 2.5 ScanThrottle + ThreadCount

Sets how many CPU threads the encoder uses (line 47).

```powershell
$ThreadCount = 8    # User-adjustable (4-16).
```

Limits how many peak-scan tasks run at the same time (line 49).

```powershell
$ScanThrottle = [Math]::Max(1, [int]([Environment]::ProcessorCount / $ThreadCount))
# Recommended overrides: 1, 2, 3, or 4
```

**What ScanThrottle does:**
- Limits parallel SPN scans to avoid CPU overload.
- Keeps the system responsive during peak detection.
- Does not affect encoding speed.
- SSD/NVMe: safe to override up to 4. Spinning disk: keep at 1.

### 3. The "De-Sync" Safeguard
These flags are applied after the input to maintain perfect alignment on long-duration files.

* **`-avoid_negative_ts make_zero`**
    * **What:** Forces all stream timestamps to start at absolute zero.
    * **Why:** Many digital files contain negative initial "Presentation Time Stamps" (PTS); this clamps them to zero to prevent audio lag or lead.
* **`-max_muxing_queue_size 14000`** *(1024-100000) recommended range
    * **What:** Increases the muxing buffer headroom.
    * **Why:** Audio down-mixing and re-encoding (AAC 7.1, TrueHD 7.1, ...) can run slower than the muxer, causing FFmpeg’s queue to overflow and trigger errors. 14000 provides enough headroom to prevent these overflows on **all movie lengths.**
* **`-map 0:v? -c:v copy`**
    * **What:** Maps all video tracks and sets them to stream-copy mode.
    * **Why:** Ensures zero quality loss for the video while moving at maximum speed.
* **`-map_metadata 0 -map_chapters 0`**
    * **What:** Copies global metadata and chapter markers from the source.
    * **Why:** Preserves the original file's organizational structure (Title, Year, Scene markers).

### 3.2 Safe Peak Normalizer (SPN) 

* **What:** Safe Peak Normalizer checks each audio track before processing and only steps in when the volume is too close to clipping (above –0.5 dBFS).
* **Why:** Some movies arrive with peaks already above 0 dBFS. When that happens, even simple mixing of the channels can push those peaks higher and cause distortion in the final encode. SPN catches this early so the issue never reaches the rest of the audio processing.
* **How:** Safe Peak Normalizer runs right after the script decides each track’s action (Encode, Downmix, etc.). If the peak scan shows the track is too hot, a light loudness-correction step is added at the very start of that track’s filter chain.
* **Clean Sources:** Nothing is changed — the original loudness and dynamics stay exactly the same.
* **Passthrough tracks:** — Peak detection is skipped entirely. Copied streams cannot have filters applied.
* **Dirty Sources:** SPN fixes unsafe peaks without making the movie sound compressed or “over-normalized.”
* **Result:** More consistent, safer audio across all movies, keeping clean sources untouched.

<div style="page-break-after: always;"></div>

#### Threshold (line 53)
These are the values Safe Peak Normalizer uses to decide when to step in.

```powershell
$PeakThresholdDB = -0.5 # SPN: loudnorm triggers when source peak exceeds this (dBFS)
```

- Valid range: **−1.5 dBFS to −0.3 dBFS** (recommended: −0.5 dBFS)
- Lower values (e.g. −1.5) normalize more tracks — more conservative.
- Higher values (e.g. −0.3) normalize fewer tracks — only the hottest sources.

| Setting | Value |
|---------|-------|
| **loudnorm** target | −23 LUFS |
| **Loudness range target** | LRA=11 LU |
| **True peak ceiling** | −1.5 dBTP |

### Three‑Stage SPN Processing

The audio engine processes each audio track in three phases, with SPN peak scans running in parallel during Phase B to keep multi‑track processing fast.

```
3-phase SPN Architecture (Phase A > Phase B > Phase C)
Input MKV >> Probe
     > Phase A  (rule matching + malformed-layout guards + commentary removal)
     > Phase B  (parallel SPN peak scans)
     > Phase C  (sequential merge + NeedsNormalization resolution)
     > FFmpeg command builder
       >> Output MKV
```

**Phase A, Sequential rule matching**  
Every audio stream is probed and matched against the rule table. Malformed-layout guards and commentary detection run here. 2.0 tracks are marked Removed. Surviving tracks receive a pending action (Passthrough, Downmix, Encode).

**Phase B, Parallel SPN peak scans**  
All non-Removed, non-Passthrough tracks are scanned in parallel for peak levels using **volumedetect**.

**Phase C, Sequential merge**  
Peak results are merged back into the track list. **NeedsNormalization** is resolved from the scan data. The final track list is handed to the FFmpeg command builder.

### 4. Stereo (2.0) Tracks — Removed
All 2.0 stereo tracks are removed by this script, regardless of codec. Stereo output is delegated to **ConvertAudioEngine-Stereo**, which handles 2.0 encoding with the correct stereo-specific flags and bitrates. This keeps each script focused on a single job.

### 5. Encoding & Downmix Execution
These parameters determine how the audio is shaped and processed during encoding.

* **`-filter:a:$i $panFilter`**
    * **What:** Applies the custom channel-mapping matrix (the Pan Filter).
    * **Why:** Necessary for downmixing 7.1 to 5.1. It ensures that side channels (SL/SR) are folded into the rear speakers rather than simply deleted.
* **`-c:a:$i eac3`**
    * **What:** Sets the codec for audio stream `$i` to Dolby Digital Plus (EAC3).
    * **Why:** Chosen for its high efficiency and broad compatibility with modern smart TVs and home theater receivers.
* **`-ac 6`**
    * **What:** Forces the output to exactly 6 channels (5.1).
    * **Why:** Applied on the 5.1 re-encode path to guarantee the output channel count is correct, even if the source reports a non-standard layout.
* **`-b:a:$i $t.Bitrate`**
    * **What:** Sets the target bitrate (e.g., 1152k for 7.1 downmix, 768k for 5.1 re-encode, 1024k for DTS-HD).
    * **Why:** High bitrates ensure the "transparency" of the encode, making it difficult to distinguish from the lossless original.
* **`-dialnorm -31`**
    * **What:** Sets Dialogue Normalization to the reference "off" position.
    * **Why:** Prevents the playback device from applying its own internal volume leveling, preserving the original mix's intended loudness.
* **`-cutoff 20000`**
    * **What:** Sets the low-pass filter frequency to 20 kHz.
    * **Why:** Preserves the high-frequency clarity of the audio that standard encoders sometimes discard to save space.
    * **Applies to:** Downmix path and 5.1 re-encode path only.
* **`-metadata:s:a:$i "title=..."`**
    * **What:** Injects a custom string into the stream title metadata.
    * **Why:** Allows users to see exactly what the track is (e.g., "DD+ 5.1 Downmix (1152k)") in their media player's audio selection menu.

### 6. Audio Rules & Priority
The audio-engine uses a **Priority Mapping (0-100)** to evaluate tracks based on quality and codec type.

### **EAC3 / Atmos Rules**
| Priority | Condition | Action | Output |
|----------|-----------|--------|--------|
| 100 | **Atmos (JOC)** | Passthrough | Original EAC3 Atmos — preserves object metadata |
| 99 | **EAC3 7.1** | Downmix | DDP 5.1 @ 1152k — maps 8 channels to 6 using pan filter |
| 98 | **EAC3 5.1** | Passthrough | Original EAC3 5.1 |
| 97 | **EAC3 2.0** | Removed | Delegated to ConvertAudioEngine-Stereo |

### TrueHD Rules
| Priority | Channels | Action | Output |
|----------|----------|--------|--------|
| 81 | 5.1 | Passthrough | Original TrueHD 5.1 |
| 80 | 7.1 | Downmix | DDP 5.1 @ 1152k |

### AAC Rules
| Priority | Channels | Action | Output |
|----------|----------|--------|--------|
| 92 | 7.1 | Downmix | DDP 5.1 @ 1152k |
| 91 | 5.1 | Encode | DDP 5.1 @ 768k |
| 90 | 2.0 | Removed | Delegated to ConvertAudioEngine-Stereo |

### DTS Rules
| Priority | Variant | Action | Output |
|----------|---------|--------|--------|
| 73 | DTS-HD MA/HRA (any ch > 2) | Encode | DDP 5.1 @ 1024k |
| 72 | DTS Core (any ch > 2) | Encode | DDP 5.1 @ 768k |
| 71 | DTS-HD 2.0 | Removed | Delegated to ConvertAudioEngine-Stereo |
| 70 | DTS 2.0 | Removed | Delegated to ConvertAudioEngine-Stereo |

### PCM / FLAC Rules
| Priority | Channels | Action | Output |
|----------|----------|--------|--------|
| 62 | 7.1 | Downmix | DDP 5.1 @ 1152k |
| 61 | 5.1 | Encode | DDP 5.1 @ 768k |
| 60 | 2.0 | Removed | Delegated to ConvertAudioEngine-Stereo |

### 7. Spatial Audio: The Pan Filter Matrix
When 7.1 audio is reduced to 5.1, standard downmixing often loses spatial detail from the side channels.

### **The Formula**

> aformat=channel_layouts=7.1,pan=5.1|FL=FL|FR=FR|FC=FC|LFE=LFE|BL=BL+0.707*SL|BR=BR+0.707*SR  
> alimiter=limit=0.948:attack=5:release=50:level=disabled:latency=1

### Layout Locking (aformat)
**What:** Forces the decoder to output a strict 8-channel 7.1 layout *before* the Pan Filter.  
**Why:** Some codecs (TrueHD, AAC, EAC3) may decode into non-standard 7.1 variants.  
**Locking the layout ensures:**
- SL really is the Side Left channel
- BL really is the Back Left channel
- No channel swapping
- No silent corruption of the downmix

### The Math (0.707)
**What:** Applies a –3 dB coefficient to the Side Surrounds (SL/SR).  
**Why:** Prevents the side channels from overpowering the rear soundstage and avoids digital clipping.

### Limiter (0.948) — For All 7.1 Codecs  
**alimiter=limit=0.948:attack=5:release=50:level=disabled:latency=1**
**Why:** All 7.1 codecs can produce true peak overshoots after decoding:
-  AAC 7.1    >  AAC can create peaks slightly above 0 dB after decoding.
-  TrueHD 7.1 >  Combining similar channels can briefly push the volume higher.
-  DTS-HD MA  >  Decoder can rebuild peaks that go a bit above the original level.
-  PCM/FLAC   >  Short volume spikes can appear when channels are mixed together.
-  EAC3 7.1   >  Channels are merged into fewer speakers during downmixing.

Limiter ceiling of 0.948 provides about 0.5 dB of true-peak headroom, preventing:
1.) +0.0x dB overshoots
2.) clipping

**latency=1** allows the limiter a 1‑sample look‑ahead window to catch post‑downmix peaks without adding audible delay. 

#### Important:
- If the source audio already contains true-peak or inter-sample peaks (“clips”), these will also appear in **analysis** of the downmix.

- The limiter prevents new clipping during downmixing and encoding, but **it does not remove overshoots inherent in the source mix**. 

### 8. Commentary Filtering
**Only 2-channel commentary tracks** are removed to save space while preserving primary content.

* **Detection Methods:** Title matching, keyword detection, and case-insensitive search.
* **Keywords:** commentary, director, producer, writer, cast, behind, bonus, alt, interview.
* **Why 5.1 Commentary is NOT removed:**
    * It is extremely rare in retail media.
    * It is often hard to tell apart from legitimate "Alternate Mixes" or "Home Theater Optimized" tracks.
    * Removing it carries a high risk of "False Positives" (deleting real movie content).

<div style="page-break-after: always;"></div>

### 9. Malformed Layout Guards
The script corrects illegal 7-channel layouts before rule matching runs:

* **AAC 7ch** → 6ch (treated as 5.1)
* **EAC3 7ch** → 6ch (treated as 5.1)
* **TrueHD 7ch** → 8ch (treated as 7.1)
* **PCM/FLAC 7ch** → 8ch (treated as 7.1)
* **AC3 7ch** → Removed (no valid AC3 7-channel layout exists)

### **⛔ Fallback Rules**
If no specific rule matches, the following logic is applied:

| Channels | Fallback Action | Bitrate |
|----------|----------------|---------|
| 2 channels or fewer | Removed | (delegated to ConvertAudioEngine-Stereo) |
| 3 to 6 channels | Encode | 768k |
| more than 6 channels (7ch+) | Downmix | 1152k |

### 10. Summary Report
At === AUDIO PROCESSING SUMMARY ===, the script generates a detailed table with the following columns:

| Column | Description |
| :--- | :--- |
| **Index** | The original stream number (0, 1, 2, 3,). |
| **Codec** | The source format (e.g., TrueHD, DTS). |
| **Channels** | The detected number of channels (e.g., 7.1, 5.1). |
| **Action** | **Green**: Passthrough, **Yellow**: Downmix, **Blue**: Encode, **Red**: Removed. |
| **Output** | The final codec and bitrate configuration. |
| **Priority** | The assigned priority level (0–100). Removed tracks show `–`. |
| **Rule** | the logical rule that was matched during processing, such as "TrueHD 7.1 Rule" |

<div style="page-break-after: always;"></div>

### 11. Re-Enabling 2.0 Encoding (Optional)

All 2.0 rules are present in the engine but set to `Action = "Removed"` by default. This is intentional — stereo output is delegated to **ConvertAudioEngine-Stereo**. If you want this script to encode 2.0 tracks instead of removing them, you can re-enable any rule by changing three fields.

**Example — re-enabling AAC 2.0 at 256k:**

```powershell
# Find this block in $Rules_AAC (Rule3):
Action         = "Removed"
Bitrate        = $null
Rule           = "AAC_2.0_Removed"

# Change it to:
Action         = "Encode"
Bitrate        = "256k"
Rule           = "AAC_2.0_Encode_256k"
```

Only those three fields need to change. Nothing else in the engine requires modification.

Apply the same pattern to re-enable any other 2.0 rule:

| Rule block | Location | Default | Suggested bitrate |
|------------|----------|---------|-------------------|
| AAC 2.0 | Rule3 in `$Rules_AAC` | Removed | 256k |
| EAC3 2.0 | Rule4 in `$Rules_EAC3_Atmos` | Removed | 384k |
| DTS-HD 2.0 | Rule3 in `$Rules_DTS` | Removed | 384k |
| DTS 2.0 | Rule4 in `$Rules_DTS` | Removed | 256k |
| PCM/FLAC 2.0 | Rule3 in `$Rules_PCMFLAC` | Removed | 384k |

<div style="page-break-after: always;"></div>

### 12. Non-Critical FFmpeg Warnings (Safe to Ignore)

During processing, FFmpeg may display one or more of the following messages.  
These warnings are **normal**, and do not affect audio quality, sync, or re-encode.  
They come from FFmpeg’s internal decoders and subtitle parsers — not from the script!

#### **“quant_step_size larger than huff_lsbs”**
- **Cause:** Printed by the TrueHD decoder during decoding (not probe).  
- **Impact:** Safe to ignore. Does not affect decoding, channel layout, or the Pass+Copy path.

#### **“Codec AVOption drc_scale … has not been used for any stream”**
- **Cause:** (**`-drc_scale`**) applies only to AC-3/EAC3 decoders.  
  FFmpeg prints this when the input track is not AC-3 (e.g., TrueHD, DTS, FLAC).  
- **Impact:** Safe to ignore. The script disables DRC intentionally for consistent output.

#### **“(Subtitle: hdmv_pgs_subtitle) unspecified size”**
- **Cause:** PGS (Blu-ray) subtitles do not declare width/height in the container.  
  FFmpeg warns until the first subtitle packet is decoded.  
- **Impact:** Safe to ignore. Subtitles are copied untouched (**`-c:s copy`**).

#### **“Assuming an incorrectly encoded 7.1 channel layout...instead of 7.1(wide)”**
- **Cause:** AAC has two 7.1 layouts. Most encoders use the common (non‑spec) layout.  
  FFmpeg prints this when the file does not use the strict 7.1(wide) AAC layout.  
- **Impact:** Safe to ignore. The script already corrects and pins the layout using (**`aformat=channel_layouts=7.1`**)

#### Disclaimer
- Always test on a few files before batch-processing your library.  
- Keep backups of original media.  
- Results may vary depending on source quality and hardware.  