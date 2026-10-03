---
name: oralslides-local
description: Convert a local PowerPoint or PDF into a narrated MP4 using the user's AI client for scripts, Edge TTS for speech, and local LibreOffice, Poppler, and FFmpeg tools. Use when the user wants a private, local PPT-to-video workflow without the OralSlides cloud service.
---

# OralSlides Local

Use `scripts/oralslides.py` for deterministic file processing. Run commands from
the skill directory or pass the script's absolute path.

## Workflow

1. Run `doctor` and resolve missing system dependencies.
2. Run `prepare INPUT --output WORK_DIR`. It renders slide images and writes
   `manifest.json` plus a `scripts.json` narration template.
3. Read `manifest.json` and write concise spoken narration into every matching
   entry in `scripts.json`. Preserve each `slideNumber`; do not leave blank
   scripts or add/remove slides.
4. Run `tts WORK_DIR`. Let language-based voice selection apply unless the user
   requested a voice, language, or rate.
5. Run `compose WORK_DIR --output OUTPUT.mp4`.
6. Report the output path and any warning emitted by the tool.

Generated files are resumable. Re-running `tts` reuses audio when the script,
voice, and rate have not changed. `compose` always validates slide/audio pairing
before writing the video.

Rendering and FFmpeg processing remain local. Edge TTS is network based and
receives the narration text. Tell the user this when privacy is material to the
request.
