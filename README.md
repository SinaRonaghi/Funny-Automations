# Automations nobody asked for

Tiny Windows automations for the last mile of academic slide production: renaming narration files, inserting AI-generated audio into PowerPoint, fixing playback timing, and exporting the result into a video.

This repo is intentionally practical and a little theatrical. The real goal is simple: take a stack of professor slides and turn them into a narrated, synchronized, exportable presentation with as little manual clicking as possible.

## A bit of a context

This project was born out of a European research initiative and a looming mountain of educational deliverables. The original plan was traditional: write slides, hire a human presenter, and record sessions one by one. It was a logistical nightmare and a massive time-sink.

To save our sanity, we pivoted. We drafted scripts for every slide and fed them into VertexAI’s Chirp 3 model. Suddenly, we had high-quality, human-like audio, but we also had a new problem: The Manual Integration Grind.

### The "Speech.wav" Problem

VertexAI is great at audio, but terrible at file naming. Every download is a repetitive `speech.wav`, `speech (1).wav`, and so on. To make matters worse, I had a universal `Final.wav` for the closing slides that needed to play last.

Rather than rewriting a complex sorting algorithm, I took the path of least resistance:

- `speech.wav` becomes `speech (0).wav`
- `Final.wav` becomes `zz.wav` (because alphabetical sorting is a developer's best friend)

### The Solution

This repo automates the "boring stuff" that happens between having an audio file and having a finished video:

- **rename.bat**: A quick-and-dirty script to fix the VertexAI naming conventions so they play nice with Windows sorting.
- **process_presentation.py**: The heavy lifter. Using `pywin32`, it injects the audio into PowerPoint, sets the transition to Push Up, and forces the playback sequence to "Start With Previous."

The result? A fully synchronized video export without the thousand-click headache.


## What lives here

- [process_presentation.py](process_presentation.py) opens a PowerPoint deck, inserts `.wav` narration files, sets each slide to a push-up transition, makes the audio play with previous, saves an intermediate deck, and exports a video.
- [rename.bat](rename.bat) normalizes audio file names before PowerPoint sees them.
- [Macro.txt](Macro.txt) contains a VBA fallback for repairing audio trigger timing and transitions directly inside PowerPoint.

## The workflow

1. Generate narration audio from your slide text using Vertex AI Chirp 3 or another TTS pipeline.
2. Put the audio files in the same folder as the PowerPoint deck.
3. Run [rename.bat](rename.bat) if you need the file names aligned with PowerPoint sorting.
4. Run [process_presentation.py](process_presentation.py) to inject the audio and export the final video.
5. If a presentation already has media objects but the timing is off, use the VBA macro in [Macro.txt](Macro.txt).

## Requirements

- Windows
- Microsoft PowerPoint desktop app
- Python 3.10+ with `pywin32`
- Narration files in `.wav` format

Install the Python dependency with:

```bash
pip install -r requirements.txt
```

## How to use it

### 1. Prepare the folder

Keep one `.pptx` deck in the working folder together with the narration files. The Python script ignores files ending in `_int.pptx` so it does not reprocess its own output.

### 2. Normalize audio names

Run [rename.bat](rename.bat) when you need the narration files to follow a predictable order. The current batch file renames two sample files:

- `speech.wav` to `speech (0).wav`
- `final.wav` to `zz.wav`

That looks oddly specific because it is. It is a tiny helper script, not a full asset manager.

### 3. Build the narrated deck

Run:

```bash
python process_presentation.py
```

The script will:

- open the first `.pptx` file in the current folder
- sort the `.wav` files using natural sorting
- remove existing media objects from each slide
- insert each audio file into the matching slide
- set the audio trigger to `With Previous`
- apply a `Push Up` transition to every slide
- save an intermediate presentation ending in `_int.pptx`
- export a video ending in `_video.mp4`

### 4. Repair an existing deck manually, if needed

If a deck already contains audio but the timing is wrong, paste the VBA from [Macro.txt](Macro.txt) into PowerPoint and run `FixAudioAndTransition`.

## Code walkthrough

### [process_presentation.py](process_presentation.py)

The script is built around a small PowerPoint COM automation flow.

`natural_sort_key()` keeps filenames like `1.wav`, `2.wav`, `10.wav` in the order people expect instead of the order Windows sometimes invents.

`main()` starts PowerPoint, discovers the first usable deck in the folder, and opens it for editing.

The slide loop does three things:

1. sets the slide transition to `ppEffectPushUp`
2. removes old media shapes so reruns stay clean
3. inserts the matching narration clip and adds a `MediaPlay` animation effect triggered `With Previous`

After the deck is updated, the script saves an intermediate `.pptx` and then calls PowerPoint's video export path so you end up with both the editable deck and the rendered video.

### [rename.bat](rename.bat)

The batch file is intentionally blunt. It changes specific filenames so the downstream sort order is stable. That is useful when the narration generator or TTS export uses names that are not slide-friendly.

### [Macro.txt](Macro.txt)

This VBA macro is the emergency screwdriver.

It loops through every slide, sets the transition to Push Up, and then scans the main animation sequence for audio play effects. When it finds one, it forces the trigger to `With Previous`.

In plain English: if PowerPoint got the audio timing wrong, this macro helps straighten it out without rebuilding the whole deck.

## Notes and limitations

- The Python script processes the first eligible `.pptx` in the folder.
- Audio files should be `.wav` for the current automation.
- The video export relies on the local PowerPoint installation, so this is a Windows desktop workflow.
- If you want the repo to scale beyond this playful prototype, the next step would be a config file that maps slide numbers to audio names.

## Why this exists

Because sometimes the most useful automation is the one that saves you from doing the same 40 clicks twice.
