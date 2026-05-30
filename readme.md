# Converts 7.1 audio tracks to DDP 5.1

![PowerShell](https://img.shields.io/badge/PowerShell-7.6-blue)
![Linux](https://img.shields.io/badge/OS-Linux-yellow?logo=linux&logoColor=white)

> **Downmix is only used for sources with more than 5.1 channels.**  
> **Audio 5.1 sources are only re-encoded.**  
> **Supports TrueHD, DTS-HD, DTS Core, AAC, PCM, and FLAC.**  
> **No video re-encode — video is always passed through untouched.**

Designed for home theater users who want clean, consistent audio without writing complex FFmpeg commands.

---

## Requirements

- **PowerShell 7.6** **https://github.com/PowerShell/PowerShell/releases/download/v7.6.1/PowerShell-7.6.1-win-x64.msi**
- **FFmpeg 8.1 or newer** — Place **ffmpeg** and **ffprobe** in the same folder as the script. You can grab the latest builds from the official FFmpeg download page: **https://ffmpeg.org/download.html**
- **Runs on:** Windows 10/11, macOS, Linux

> **Note:** The scripts rely on features and stability improvements introduced in FFmpeg 8.1, particularly for reliable EAC3, TrueHD, and DTS-HD parsing.

---

## Audio Scripts

| Script | Purpose |
|--------|---------|
| **ConvertAudioEngine-DDP51.ps1** | Converts all audio to DDP 5.1 — passthrough, downmix, or re-encode depending on source |
| **ConvertAudioEngine-Keep71.ps1** | Preserves the original 7.1 source track and adds a DDP 5.1 compatibility track |
| **ConvertAudio2-Stereo.ps1**| Downmixes any 5.1, 7.1, or Atmos track into a high‑quality EAC3 or Opus stereo file.|
| **AudioRemove-AC3.ps1** | Removes low bit-rate AC3/E-AC3 streams to produce a lean master before conversion. **Uses FFmpeg** |
| **AudioRemove-AC3-mkvmerge.ps1** | Removes low bit-rate AC3/E-AC3 streams using **mkvmerge**
| **AudioPeakRMSChecker.ps1** | Validates peak, RMS, crest factor, and clipping on source and converted tracks |

## Quick Start

### Windows11

1. Place the script, **`ffmpeg.exe`**, and **`ffprobe.exe`** in the same folder.
2. Run:

```powershell
pwsh -ExecutionPolicy Bypass -File .\ConvertAudioEngine-DDP51.ps1 ".\YourMovie.mkv"
```

```powershell
pwsh -ExecutionPolicy Bypass -File .\ConvertAudioEngine-Keep71.ps1 ".\YourMovie.mkv"
```

> The **-ExecutionPolicy Bypass** flag only applies to this single run and does not change your system settings.

### macOS / Linux

1. Place the script, **`ffmpeg`**, and **`ffprobe`** in the same folder.
2. Make the binaries executable (first time only):

```bash
chmod +x ffmpeg ffprobe
```

3. Run:

```powershell
pwsh -File ./ConvertAudioEngine-DDP51.ps1 "./YourMovie.mkv"
```

```powershell
pwsh -File ./ConvertAudioEngine-Keep71.ps1 "./YourMovie.mkv"
```

### Output filenames

| Script | Output |
|--------|--------|
| **`ConvertAudioEngine-DDP51.ps1`** | `YourMovie_ddp5only.mkv` |
| **`ConvertAudioEngine-Keep71.ps1`** | `YourMovie_keep71mix.mkv` |

If an output file already exists, it will be overwritten automatically.

## Comparison — Audio Scripts vs. Typical FFmpeg/PowerShell Scripts

| Feature / Capability | Audio Script Set   | Typical GitHub Scripts |
|----------------------|-------------|------------------------|
| **De-sync Protection** | ✔ | ✖ |
| **Safe Peak Normalizer (SPN)** | ✔ | ✖ |
| **Sync-Stability System** | ✔ | ✖ |
| **PTS Correction** | ✔ | ✖ |
| **Mux-Queue Protection** | ✔ | ✖ |
| **No video re-encode** | ✔ | Sometimes broken |
| **TrueHD 5.1 passthrough** | ✔ | ✖ (often re-encodes) |
| **EAC3 Atmos (JOC) passthrough** | ✔ | ✖ |
| **EAC3 5.1 / 2.0 passthrough** | ✔ | ✖ |
| **DTS-HD MA/HRA detection (1024k)** | ✔ | ✖ |
| **DTS Core detection (768k)** | ✔ | ✖ |
| **PCM/FLAC stereo handled at 384k** | ✔ | ✖ |
| **7.1 → 5.1 intelligent downmixing** | ✔ | Partial |
| Rule-based audio engine | ✔ | ✖ |
| Modular architecture | ✔ | ✖ |
| Precompiled regex for performance | ✔ | ✖ |
| Malformed layout guards (audio codecs) | ✔ | ✖ |
| Safe fallback rules | ✔ | ✖ |
| Subtitle & chapter passthrough | ✔ | Partial |
| MKV attachment passthrough (fonts, cover art) | ✔ | ✖ |
| Adjustable audio bitrates | ✔ | ✖ |
| Add new codecs easily | ✔ | ✖ |
| Commentary detection & removal | ✔ | ✖ |
| Clean, color-coded summary output | ✔ | ✖ |
| Cross-platform (Windows, macOS, Linux) | ✔ | Rarely |
| Batch mode (folder processing) | ✖ | Sometimes |

---

## Why DD+?

EAC3 (DD+) plays nicely with almost everything — TVs, soundbars, streaming boxes, and headphones.  
It delivers solid surround quality without huge file sizes, making it a reliable format that works on virtually every device.

---

## How It Decides What to Do

| Original Format | Channels | Action | Output |
|-----------------|----------|--------|--------|
| EAC3 Atmos (JOC) | Any | Passthrough | Unchanged |
| EAC3 | 5.1 | Passthrough | Unchanged |
| EAC3 | 2.0 | Passthrough | Unchanged |
| TrueHD | 5.1 | Passthrough | Unchanged |
| TrueHD | 7.1 | Downmix | EAC3 5.1 @ 1024k |
| DTS-HD MA / DTS-HRA | >2 | Encode | EAC3 5.1 @ 1024k |
| DTS Core | >2 | Encode | EAC3 5.1 @ 768k |
| AAC / PCM / FLAC | 7.1 | Downmix | EAC3 5.1 @ 1024k |
| AAC / PCM / FLAC | 5.1 | Encode | EAC3 5.1 @ 768k |
| AAC | 2.0 | Encode | EAC3 2.0 @ 256k |
| PCM / FLAC | 2.0 | Encode | EAC3 2.0 @ 384k |
| Stereo + commentary title | 2.0 | Remove | — |

If something doesn't match any rule, it falls back to sensible defaults.

---

## Customization

You can tweak these settings directly in the script:

- **`CommentaryKeywords`** — Add or remove words that trigger commentary track removal.
- **Audio rule groups** (`$Rules_AAC`, `$Rules_DTS`, etc.) — Adjust bitrates, add new codecs, or change behavior per codec family.
- **`ThreadCount`** — Performance tuning. Defaults to 8 threads. Adjustable from **4 to 16**.
- **`ScanThrottle`** — Limits how many peak-scan tasks run simultaneously. Recommended values: **1 to 4**. Use 1 for Hard Drives, up to 4 for SSD/NVMe.

---

## Utility Scripts

### AudioRemove-AC3.ps1 - Uses FFMpeg 8.1

Removes all low-bitrate AC3 and E-AC3 audio streams from an MKV, producing a lean master file containing only high-fidelity audio (TrueHD / AAC 7.1 / DTS-HD MA). Run this before the conversion engine if your source contains low-bitrate AC3 compatibility tracks you want to strip first.

```powershell
# Windows
pwsh -ExecutionPolicy Bypass -File .\AudioRemove-AC3.ps1 ".\YourMovie.mkv"

# macOS / Linux
pwsh -File ./AudioRemove-AC3.ps1 "./YourMovie.mkv"
```
**Output:** 
Output: `YourMovie_remux.mkv`

### AudioRemove-AC3-mkvmerge.ps1
Same purpose as above, but uses **mkvmerge** for byte‑identical remuxing.  
Ideal for users who prefer MKVToolNix or want consistent output.

### AudioPeakRMSChecker.ps1

Validates audio levels on source and converted tracks. Reports peak level dB, RMS level dB, crest factor, and peak count — with PASS/FAIL evaluation per track. Useful for confirming the conversion engine produced clean, properly leveled output.
Also has:
- TrueHD preroll reporting  
- A/V sync offset detection  
- Per-track astats logs  
- JSON output mode for automation

```powershell
# Windows
pwsh -ExecutionPolicy Bypass -File .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv"

# macOS / Linux
pwsh -File ./AudioPeakRMSChecker.ps1 "./YourMovie.mkv"
```

---

## Disclaimer

Always test on a few files first before running on your whole library.  
Always keep backups of your original files.
