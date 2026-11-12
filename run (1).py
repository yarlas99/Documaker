#!/usr/bin/env python3
"""
Video + Word transcript  ->  step-by-step how-to document (.docx) with screenshots.

Usage:
    python run.py video.mp4 transcript.docx -o howto.docx

LLM provider (auto-detected, or set LLM_PROVIDER=azure|anthropic):
    Azure OpenAI : AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY,
                   AZURE_OPENAI_DEPLOYMENT  (optional: AZURE_OPENAI_API_VERSION)
    Anthropic    : ANTHROPIC_API_KEY  (optional: ANTHROPIC_MODEL)

Requires ffmpeg + ffprobe on PATH.
"""
import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph

# Matches 0:05, 12:34, 1:02:03
TS = re.compile(r"(?<!\d)(?:(\d{1,2}):)?(\d{1,2}):(\d{2})(?!\d)")


# ----------------------------------------------------------------- transcript
def to_sec(m):
    return int(m.group(1) or 0) * 3600 + int(m.group(2)) * 60 + int(m.group(3))


def read_transcript(path):
    """Parse a .docx transcript into [{'t': seconds, 'text': str}, ...].

    Tolerant of Teams-style ("0:05 Speaker" then text on next line),
    "[00:05] text", and "Speaker  0:05" layouts, including tables.
    """
    doc = Document(path)
    lines = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            text = Paragraph(child, doc).text.strip()
            if text:
                lines.append(text)
        elif child.tag.endswith("}tbl"):
            for row in Table(child, doc).rows:
                cells = []
                for c in row.cells:
                    t = c.text.strip()
                    if t and t not in cells:
                        cells.append(t)
                if cells:
                    lines.append("  ".join(cells))

    segs = []
    for line in lines:
        m = TS.search(line)
        if m and m.start() < 60:
            rest = (line[: m.start()] + " " + line[m.end():]).strip(" -\u2013\u2014:[]()\t")
            segs.append({"t": to_sec(m), "text": rest})
        elif segs:
            segs[-1]["text"] = (segs[-1]["text"] + " " + line).strip()
    if not segs:
        sys.exit("Could not find any timestamps in the transcript. "
                 "Open it and check the format (expected things like 0:05 or 00:01:30).")
    return segs


def fmt(sec):
    sec = int(sec)
    return f"{sec // 60:02d}:{sec % 60:02d}"


# ------------------------------------------------------------------------ LLM
SYSTEM = (
    "You turn a timestamped transcript of someone demonstrating a process on screen "
    "into a clear, accurate step-by-step procedure. Reply with JSON only."
)

PROMPT = """Below is a transcript of a screen recording in which someone explains a process.
The video is {duration:.0f} seconds long. Each line starts with its timestamp in seconds.

Produce a how-to guide as JSON with exactly this shape:
{{
  "title": "short title for the process",
  "overview": "1-3 sentence summary of what the process achieves and who it is for",
  "prerequisites": ["optional list of things needed before starting"],
  "steps": [
    {{
      "title": "short imperative title, e.g. 'Open the Export menu'",
      "description": "clear instructions a new person could follow, 1-4 sentences. Name the exact buttons/menus/fields mentioned.",
      "start_sec": 0,
      "end_sec": 0,
      "screenshot_sec": 0,
      "needs_screenshot": true
    }}
  ]
}}

Rules:
- Only use information from the transcript. Do not invent steps, buttons or values.
- Drop filler, small talk and repetition. Merge fragments that belong to one action.
- Steps must be in order. start_sec/end_sec come from the transcript timestamps.
- screenshot_sec: the moment the screen shows the RESULT of the step (e.g. just after the
  click, when the dialog/page is open). Usually a second or two after the action is described,
  and never beyond the video length. Set needs_screenshot=false for steps that are purely
  explanatory with nothing useful to see.

TRANSCRIPT:
{transcript}
"""


def call_llm(prompt):
    provider = os.environ.get("LLM_PROVIDER") or (
        "azure" if os.environ.get("AZURE_OPENAI_ENDPOINT") else "anthropic"
    )
    if provider == "azure":
        from openai import AzureOpenAI

        client = AzureOpenAI(
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
            api_key=os.environ["AZURE_OPENAI_API_KEY"],
            api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2024-10-21"),
        )
        r = client.chat.completions.create(
            model=os.environ["AZURE_OPENAI_DEPLOYMENT"],
            messages=[{"role": "system", "content": SYSTEM},
                      {"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.2,
        )
        return r.choices[0].message.content
    import anthropic

    client = anthropic.Anthropic()
    r = client.messages.create(
        model=os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
        max_tokens=4000,
        system=SYSTEM,
        messages=[{"role": "user", "content": prompt}],
    )
    return r.content[0].text


def extract_json(text):
    a, b = text.find("{"), text.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("No JSON found in model output:\n" + text[:500])
    return json.loads(text[a : b + 1])


# ---------------------------------------------------------------------- video
def video_duration(video):
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=nw=1:nk=1", video],
        capture_output=True, text=True, check=True,
    ).stdout.strip()
    return float(out)


def grab_frame(video, sec, out_path):
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", "-ss", f"{sec:.2f}", "-i", video,
         "-frames:v", "1", "-q:v", "2", str(out_path)],
        check=True,
    )


# ------------------------------------------------------------------- document
def build_docx(guide, shots, out_path):
    doc = Document()
    doc.add_heading(guide.get("title", "Process Guide"), 0)
    if guide.get("overview"):
        doc.add_paragraph(guide["overview"])
    if guide.get("prerequisites"):
        doc.add_heading("Before you start", 1)
        for p in guide["prerequisites"]:
            doc.add_paragraph(p, style="List Bullet")
    doc.add_heading("Steps", 1)
    for i, step in enumerate(guide["steps"], 1):
        doc.add_heading(f"Step {i}: {step['title']}", 2)
        doc.add_paragraph(step["description"])
        shot = shots.get(i)
        if shot:
            doc.add_picture(str(shot), width=Inches(6))
            cap = doc.add_paragraph(f"Screenshot at {fmt(step['screenshot_sec'])}")
            cap.runs[0].italic = True
    doc.save(out_path)


# ----------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("video")
    ap.add_argument("transcript", help=".docx transcript with timestamps")
    ap.add_argument("-o", "--out", default="howto.docx")
    ap.add_argument("--offset", type=float, default=0.0,
                    help="seconds added to every screenshot time (tune if shots are early/late)")
    ap.add_argument("--prompt-only", action="store_true",
                    help="no API call: write work/prompt.txt to paste into any chat (e.g. free claude.ai)")
    ap.add_argument("--steps", metavar="JSON",
                    help="no API call: use a steps JSON file (e.g. the chat's answer saved as steps.json)")
    args = ap.parse_args()

    work = Path("work")
    work.mkdir(exist_ok=True)

    segs = read_transcript(args.transcript)
    dur = video_duration(args.video)
    print(f"Transcript: {len(segs)} segments | video: {dur:.0f}s")

    transcript_text = "\n".join(f"[{s['t']}s] {s['text']}" for s in segs)
    prompt = PROMPT.format(duration=dur, transcript=transcript_text)

    if args.prompt_only:
        (work / "prompt.txt").write_text(SYSTEM + "\n\n" + prompt, encoding="utf-8")
        print("Wrote work/prompt.txt. Paste it into a chat, save the JSON reply as steps.json, "
              "then rerun with --steps steps.json")
        return

    if args.steps:
        guide = extract_json(Path(args.steps).read_text(encoding="utf-8"))
        print(f"Loaded {len(guide['steps'])} steps from {args.steps} (no API call)")
    else:
        raw = call_llm(prompt)
        guide = extract_json(raw)
        (work / "steps.json").write_text(json.dumps(guide, indent=2))
        print(f"LLM produced {len(guide['steps'])} steps (saved work/steps.json)")

    shots = {}
    for i, step in enumerate(guide["steps"], 1):
        if not step.get("needs_screenshot", True):
            continue
        t = step.get("screenshot_sec", step.get("end_sec", step.get("start_sec", 0)))
        t = max(0.0, min(float(t) + args.offset, dur - 0.1))
        step["screenshot_sec"] = t
        path = work / f"step_{i:02d}.jpg"
        grab_frame(args.video, t, path)
        shots[i] = path

    build_docx(guide, shots, args.out)
    print(f"Done -> {args.out}  (frames in work/ for manual swaps)")


if __name__ == "__main__":
    main()
