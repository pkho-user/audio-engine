### 1. Overview for AudioPeakRMSChecker.ps1

This script analyzes the audio tracks inside a movie file (MKV, MP4). It checks for:  
**Peak** (dB) loudest moment in the audio.  
**RMS** (dB) average loudness over time.  
**Crest Factor** (dB) difference between Peak and RMS  
**Clipping** (number of samples that hit or exceed 0 dB)

The function is to give you a quick PASS/FAIL result so you can tell whether the audio is clean, safe, and ready for encoding, downmixing, or playback. It supports common audio formats, including TrueHD 7.1, DTS‑HD 7.1, AAC (7.1 / 5.1 / 2.0), Dolby Digital Plus (5.1 / 2.0), and Opus 2.0.

This tool does **not** modify audio. It only analyzes it and reports the results.


### 2. Requirements
To run AudioPeakRMSChecker, you need:  
- PowerShell 7.6  
- FFmpeg 8.1 (ffmpeg and ffprobe must be placed in the same folder as the script.)

Official FFmpeg download page: https://ffmpeg.org/download.html

### Quick Start (Windows 11)

1. Place **`AudioPeakRMSChecker.ps1`**, **`ffmpeg.exe`**, and **`ffprobe.exe`** in the same folder.
2. Run the script:

```powershell
pwsh -ExecutionPolicy Bypass -File .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv"
```
ℹ️ The -ExecutionPolicy Bypass only applies to this single run and doesn't change your system settings. 

### macOS / Linux

1. Place the script, **`ffmpeg`**, and **`ffprobe`** in the same folder.
2. Make the binaries executable (first time only):

```bash
chmod +x ffmpeg ffprobe
```
<br>

3. Run:

```powershell
pwsh -File ./AudioPeakRMSChecker.ps1 "./YourMovie.mkv"
```

### 3. Supported Audio Formats

AudioPeakRMSChecker analyzes the following audio formats:

| Codec Family   | Formats Included                          |
|----------------|-------------------------------------------|
| **TrueHD**     | TrueHD 7.1                                |
| **AAC**        | AAC 7.1, AAC 5.1, AAC 2.0                 |
| **EAC3 / DDP** | Dolby Digital Plus 5.1, Dolby Digital Plus 2.0 |
| **DTS**        | DTS-HD MA 7.1                             |
| **Opus**       | Opus 2.0                                  |

The script automatically detects:

- the codec  
- the number of channels  
- whether the track is a source (7.1) or downmix (5.1)  
- whether the track is supported for analysis  
- Unsupported tracks are ignored safely.

### 4. How the Analysis Works

The script performs the following steps:

1. **Scans the movie file**  
   It looks for audio tracks and identifies their codec and channel layout.

2. **Runs FFmpeg’s “astats” filter**  
   This filter measures loudness, peaks, clipping, and dynamic range.

3. **Collects the results**  
   The script extracts the important numbers from FFmpeg’s output.

4. **Checks for problems**  
   It looks for:
   - peaks too close to 0 dB  
   - audio that is too loud overall  
   - low dynamic range  
   - clipped samples  

5. **Returns PASS or FAIL**  
   PASS means the audio is clean and safe.  
   FAIL means the audio may distort or has loudness issues.

6. **Saves a log file** (*.astats.txt)
   A text file is created with the full FFmpeg analysis for reference.  


### 5. What Each Metric Means

AudioPeakRMSChecker reports four main metrics:

### **Peak (dB)**
- The loudest moment in the audio.
- Values **near 0 dB** mean the audio is very loud.
- Values **above 0 dB** indicate clipping.

### **RMS (dB)**
- The average loudness over time.
- Higher RMS (closer to 0 dB) means the audio is louder.
- Lower RMS (–20 dB to –35 dB) means more dynamic range.

### **Crest Factor (dB)**
- The difference between Peak and RMS.
- Higher crest = more dynamic, cinematic sound.
- Low crest = compressed or flat audio.

### **Clipped Samples**
- The number of samples that hit or exceed 0 dB.
- A few clips are usually harmless.
- Many clips indicate distortion.

These metrics help determine whether the audio is clean or risky.

### 5.5 Quick Interpretation

Each track gets a **Quick Interpretation** that puts the numbers into plain language.

**Peak ranges:**
- 0 dB and above — at or above 0 dB (risk of clipping) [WARNING]
- –1 to 0 dB — very hot (close to clipping) [WARNING]
- –3 to –1 dB — healthy (good headroom) [PASS]
- Below –3 dB — low (very safe headroom) [PASS]

**RMS ranges:**
- Above –12 dB — very loud (compressed mix) [WARNING]
- –20 to –12 dB — moderately loud (typical TV mix) [PASS]
- –30 to –20 dB — normal for movies (healthy dynamics) [PASS]
- Below –30 dB — very low (quiet or highly dynamic) [PASS]

**Crest factor ranges:**
- Below 6 dB — heavily compressed [WARNING]
- 6 to 12 dB — balanced dynamics [PASS]
- 12 to 20 dB — good dynamic range [PASS]
- Above 20 dB — extremely dynamic (very wide range) [PASS]

**Clipping interpretation:**
- 50 or more clips at ≥ 0 dB — audio distortion likely [WARNING]
- 1 or more clips at ≥ 0 dB — clipped samples detected [WARNING]
- 6 or more clips below 0 dB — minor distortion possible [WARNING]
- 1 to 5 clips — not clipping [PASS]
- 0 clips — clean [PASS]

For TrueHD tracks, the Quick Interpretation also shows the preroll offset
in milliseconds. This is a normal PTS offset, not a sync issue.

### 6. PASS / FAIL Rules (Simple)

The script marks a track as **FAIL** if any of the following are true:

- **Peak > 0 dB**  
  The audio is clipping.

- **RMS louder than –12 dB**  
  The audio is extremely loud and may distort.

- **Crest < 6 dB**  
  The audio is heavily compressed.

- **Clipped samples > 5**  
  Too many clipped peaks.

- **Peak ≥ 0 dB with any clipped samples**  
  Even a single clipped sample triggers FAIL when the peak is at or above 0 dB.

If none of these conditions are met, the track is marked **PASS**.

PASS means:
- clean audio  
- safe headroom  
- healthy dynamics  
- no distortion  

FAIL means:
- the audio may distort  
- the mix may be too loud  
- the track may need correction  

### 7. JSON Output Mode

Useful for: Debugging and troubleshooting (easier to inspect raw values).  
In JSON mode: No summary is printed, No color output, No interpretation text.  

$DefaultJsonMode = $false (top of the script)  
To **enable** JSON mode by default: change `$false` to `$true` on line 43   
To **disable** JSON mode by default: change `$true` back to `$false` on line 43

You can also enable JSON mode for a single run without editing the script:
```powershell
pwsh -File .\AudioPeakRMSChecker.ps1 ".\YourMovie.mkv" -JsonMode
```  

### Exit Codes: 
  
|number|description|
|------|-----------|
| **0** | Success |
| **1** | Input validation: file missing / no audio streams / no supported codecs / no valid 7.1 or 5.1, or 2.0 combination |
| **2** | FFmpeg produced no stderr output (analysis aborted) |
| **3** | FFmpeg exited with non-zero exit code |
| **5** | FFmpeg exceeded maximum allowed runtime (timeout guardrail) |
| **6** | astats output missing required metrics (Peak / RMS / Peak count) |
| **7** | Peak level dB failed numeric parse |
| **8** | RMS level dB failed numeric parse |
| **9** | Peak count failed numeric parse  |
| **10** | astats log file was not created on disk |
| **11** | astats log file is empty or smaller than 10 bytes |
| **12** | ffmpeg or ffprobe binary missing or not executable in the script folder |

#### Tunable timeout parameters  
$MaxRuntimeMultiplier = 4,   Multiplier applied to the ETA to set the FFmpeg timeout.   
(user adjustable from 2 to 8)   

$MinTimeoutSeconds = 60,   Minimum timeout floor per track, regardless of ETA.  
(user adjustable from 30 to 300 seconds)     