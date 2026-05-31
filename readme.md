### Converts 7.1 audio tracks to DDP 5.1

![PowerShell](https://img.shields.io/badge/PowerShell-7.6-blue)
![Linux](https://img.shields.io/badge/OS-Linux-yellow?logo=linux&logoColor=white)

> **Downmix is only used for sources with more than 5.1 channels.**  
> **Audio 5.1 sources are only re-encoded.**  
> **Supports TrueHD, DTS-HD, DTS Core, AAC, PCM, and FLAC.**  
> **No video re-encode — video is always passed through untouched.**

Designed for home theater users who want clean, consistent audio without writing complex FFmpeg commands.

---

### Requirements

- **PowerShell 7.6** **https://github.com/PowerShell/PowerShell/releases/download/v7.6.1/PowerShell-7.6.1-win-x64.msi**
- **FFmpeg 8.1 or newer** — Place **ffmpeg** and **ffprobe** in the same folder as the script. You can grab the latest builds from the official FFmpeg download page: **https://ffmpeg.org/download.html**
- **Runs on:** Windows 10/11, macOS, Linux

> **Note:** The scripts rely on features and stability improvements introduced in FFmpeg 8.1, particularly for reliable EAC3, TrueHD, and DTS-HD parsing.

---

### Audio Scripts

| Script | Purpose |
|--------|---------|
| **ConvertAudioEngine-DDP51.ps1** | Converts all audio to DDP 5.1, passthrough, downmix, or re-encode depending on source |
| **ConvertAudioEngine-Keep71.ps1** | Preserves the original 7.1 source track and adds a DDP 5.1 compatibility track |
| **ConvertAudio2-DDP51.ps1** | Single track DDP 5.1 output, selects the best audio per language, removes 2.0 tracks, downmixes 7.1/Atmos, re-encodes or passes through 5.1 |
| **ConvertAudio2-Stereo.ps1** | Converts any 7.1, Atmos, 5.1, or 2.0 track to a high quality EAC3 or Opus stereo file, downmix or re-encode depending on source |
| **AudioRemove-AC3.ps1** | Removes low bit-rate AC3/E-AC3 streams to produce a lean master before conversion |
| **AudioRemove-AC3-mkvmerge.ps1** | Removes low bit-rate AC3/E-AC3 streams using **mkvmerge** (MKVToolNix 98.0+) |
| **AudioPeakRMSChecker.ps1** | Validates peak, RMS, crest factor, and clipping on source and converted tracks |

### Windows11

1. Place the script, **`ffmpeg.exe`**, and **`ffprobe.exe`** in the same folder.
2. Run:

```powershell
pwsh -ExecutionPolicy Bypass -File .\ConvertAudioEngine-DDP51.ps1 ".\YourMovie.mkv"
```

> Replace the script name with any other script — the usage is the same for all of them.  
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

> Replace the script name with any other script — the usage is the same for all of them.

### Why HandBrake Falls Short for UHD Audio Workflows

**HandBrake's audio pipeline is extremely basic:**
- no video passthrough (always re‑encodes video)  
- no ITU‑R BS.775 downmix matrix — defaults to Dolby ProLogic II  
- no correct 7.1 layout pinning (`aformat=channel_layouts=7.1`)  
- no `alimiter=0.948` to prevent clipping  
- no peak scanning  
- no loudness logic  
- no codec‑aware rules for TrueHD / EAC3 / DTS‑HD  

**All DDP5.1 and KEEP7.1 scripts form a complete audio‑engineering pipeline:**
- video stream copy (bit‑perfect UHD preservation)  
- correct 7.1 layout pinning (`aformat=channel_layouts=7.1`)  
- ITU‑R BS.775 downmix matrix  
- tuned coefficients to avoid center‑heavy dialogue  
- SL/SR folded into BL/BR at 0.707  
- `alimiter=0.948` to prevent clipping  
- SPN peak scanning + conditional `loudnorm`  
- codec‑aware rules for TrueHD, DTS‑HD, AAC, FLAC, PCM, EAC3  
- Atmos passthrough logic  
- commentary detection/removal  
- malformed layout guards  
- deep probe (`200M`) for TrueHD/AAC 7.1  
- `avoid_negative_ts make_zero`  
- `max_muxing_queue_size 14000`  
- language + title tagging  

**HandBrake cannot do any of this.**

---

### Comparison — Audio Scripts vs. HandBrake / Typical FFmpeg Scripts

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

### Why DD+?

EAC3 (DD+) plays nicely with almost everything — TVs, soundbars, streaming boxes, and headphones.  
It delivers solid surround quality without huge file sizes, making it a reliable format that works on virtually every device.

---

### Customization

You can tweak these settings directly in the script:

- **`CommentaryKeywords`** — Add or remove words that trigger commentary track removal.
- **Audio rule groups** (`$Rules_AAC`, `$Rules_DTS`, etc.) — Adjust bitrates, add new codecs, or change behavior per codec family.
- **`ThreadCount`** — Performance tuning. Defaults to 8 threads. Adjustable from **4 to 16**.
- **`ScanThrottle`** — Limits how many peak-scan tasks run simultaneously. Recommended values: **1 to 4**. Use 1 for Hard Drives, up to 4 for SSD/NVMe.

---

### Disclaimer

Always test on a few files first before running on your whole library.  
Always keep backups of your original files.
