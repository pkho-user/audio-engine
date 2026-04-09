# Audio Converter Engine (DD+ 5.1 Focused)

> **Downmix is only used for sources with more than 5.1 channels.  
> Audio 5.1 sources are only re‑encoded.  
> No video re‑encode — video is always passed through untouched.**

A PowerShell script that processes MKV files with FFmpeg to clean up and standardize the audio tracks.

It automatically:
- Removes commentary and bonus audio tracks (stereo only)
- Passthroughs high-quality formats like EAC3 Atmos and TrueHD 5.1
- Downmixes 7.1 tracks to 5.1 when needed
- Re-encodes everything else to EAC3 (DD+) for good compatibility and quality

Optimized for **FFmpeg 8.1** and **PowerShell 5.1 / 7.x**. Designed for home theater users who want clean, consistent audio without writing complex FFmpeg commands.

## Features

- Smart commentary removal — Only removes clearly labeled stereo commentary tracks. Multichannel tracks are always kept safe.
- Passthrough support — Leaves EAC3 Atmos (JOC) and TrueHD 5.1 untouched for maximum quality.
- Automatic downmix — Converts 7.1 → 5.1 (EAC3) with a sensible 20 kHz cutoff.
- Consistent encoding — Everything else becomes EAC3 5.1 or 2.0 at reasonable bitrates (256k / 768k / 1024k).
- Clean metadata — Adds language tags, bitrate info, and sensible titles. First audio track is set as default.
- Safe & predictable — Preserves video, subtitles, and chapters.
  
## Comparison — This Engine vs. Typical FFmpeg/PowerShell Scripts

This script is a **true rule-based engine**, with safety, metadata integrity, and predictable behavior as core goals.

| Feature / Capability | This Engine | Typical GitHub Scripts |
|----------------------|-------------|-------------------------|
| Rule‑based audio engine | ✔ | ✖ |
| Precompiled regex for performance | ✔ | ✖ |
| Commentary detection & removal | ✔ | ✖ |
| TrueHD 5.1 passthrough | ✔ | ✖ (often re‑encodes) |
| EAC3 Atmos (JOC) passthrough | ✔ | ✖ |
| DTS‑HD MA/HRA detection | ✔ | ✖ |
| Malformed layout guards (AAC/EAC3) | ✔ | ✖ |
| 7.1 → 5.1 intelligent downmixing | ✔ | Partial |
| 5.1 sources preserved as 5.1 (re‑encoded only) | ✔ | ✖ |
| No video re‑encode | ✔ | Sometimes broken |
| Clean metadata preservation | ✔ | Partial |
| Language tag preservation | ✔ | ✖ |
| Subtitle & chapter passthrough | ✔ | Partial |
| Batch mode (folder processing) | ✖ | Sometimes |
| Modular architecture | ✔ | ✖ |
| Clean, color‑coded summary output | ✔ | ✖ |
| Safe fallback rules | ✔ | ✖ |
| Consistent output naming | ✔ | ✖ |
| Keeps original audio tracks | ✖ (always processes audio) | ✔ (often keeps originals) |
| Adjustable audio bitrates      | ✔ | ✖ |
| Add new Audio codecs easily          | ✔ | ✖ |

## Why DD+?

EAC3 (DD+) plays nicely with almost everything — TVs, Soundbars, Streaming boxes, and Headphones.  
It delivers solid surround quality without huge file sizes, making it an awesome “works everywhere” format.

## Requirements

- **PowerShell** 5.1 or 7.x (fully tested on both)
- **FFmpeg 8.1 or newer** — Must be in your PATH or the same folder as the script. You can grab the latest builds from the [official FFmpeg download page](https://ffmpeg.org/download.html).

**Note:** The script relies on features and stability improvements introduced in FFmpeg 8.1, particularly for reliable EAC3, TrueHD, and DTS-HD parsing.

## Quick Start

1. Place `ConvertAudioEngine.ps1`, `ffmpeg.exe`, and `ffprobe.exe` in the same folder.
2. Run the script:

```powershell
powershell -ExecutionPolicy Bypass -File .\ConvertAudioEngine.ps1 ".\YourMovie.mkv"
```

The script will create a new file: YourMovie_Processed.mkv  
Note: The -ExecutionPolicy Bypass only applies to this single run and doesn't change your system settings.

## How It Decides What to Do

The script uses a simple rule-based system:

| Original Format          | Channels | Action                          | Output                          |
|--------------------------|----------|---------------------------------|---------------------------------|
| EAC3 Atmos (JOC)         | Any      | Passthrough                     | Unchanged                       |
| TrueHD                   | 5.1      | Passthrough                     | Unchanged                       |
| TrueHD                   | 7.1      | Downmix                         | EAC3 5.1 @ 1024k                |
| DTS-HD / DTS-MA          | >2       | Encode                          | EAC3 5.1 @ 1024k                |
| DTS Core                 | >2       | Encode                          | EAC3 5.1 @ 768k                 |
| AAC / PCM / FLAC         | 7.1      | Downmix                         | EAC3 5.1 @ 1024k                |
| AAC / PCM / FLAC         | 5.1      | Encode                          | EAC3 5.1 @ 768k                 |
| Any stereo               | 2.0      | Encode                          | EAC3 2.0 @ 256k                 |
| Stereo + commentary title| 2.0      | Remove                          | —                               |
If something doesn't match any rule, it falls back to sensible defaults.

## Dolby Metadata Behavior (Per Track Type)

This engine applies Dolby‑safe metadata parameters depending on the output channel layout.
| Track Type            | dialnorm | dsur_mode |
|-----------------------|----------|-----------|
| **Stereo (2.0)**      | ✔        | ✔         |
| **5.1 Encode**        | ✔        | ✔         |
| **Downmix (7.1→5.1)** | ✔        | ✔         |
| **Passthrough**       | ❌        | ❌         |

## Customization

You can easily tweak these in the script:
- `$CommentaryKeywords` — Add or remove words that trigger removal
- **Audio rule groups** (`$Rules_AAC`, `$Rules_DTS`, etc.) — Adjust bitrates, add new codecs, or change behavior for each codec family.
- `$ThreadCount` - Performance Tuning. The script defaults to 8 threads for processing. If you have a high-end CPU, you can adjust the $ThreadCount variable from 4 to 16.

## Disclaimer

This script was initially created with AI assistance and then reviewed and customized.  
It works well for my use cases, but **test it on a few files first** before running it on your whole library. Always keep backups of your original files.

Feel free to open issues or submit pull requests if you find bugs or want to add features.

## Planned Improvements
- Batch/folder processing mode
- Keep all 7.1 Audio Tracks

