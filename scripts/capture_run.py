#!/usr/bin/env python3
"""
scripts/capture_run.py
Executes a compiled binary or script, feeds interactive stdin inputs while echoing
them inline (simulating live terminal typing), auto-detects host VS Code terminal settings,
prepends an authentic terminal prompt matching the user's active folder,
and renders a crisp VS Code terminal screenshot.
"""

import os
import re
import sys
import time
import json
import argparse
import threading
import subprocess
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# Try importing auto-detector
try:
    from scripts.detect_vscode import detect_vscode_settings
except ImportError:
    try:
        from detect_vscode import detect_vscode_settings
    except ImportError:
        def detect_vscode_settings(profile_path=None):
            return {
                "font_family_name": "Consolas",
                "font_file_path": None,
                "font_size": 14,
                "terminal_bg_hex": "#181818",
                "background_hex": "#1E1E1E",
            }

# Terminal Palette
TERM_PALETTE = {
    "bg": "#181818",
    "top_bar_bg": "#1F1F1F",
    "top_bar_border": "#2B2B2B",
    "top_bar_text": "#858585",
    "prompt_ps": "#569CD6",
    "prompt_path": "#D4D4D4",
    "command": "#DCDCAA",
    "output_text": "#D4D4D4",
    "cursor": "#AEAFAD",
}


def hex_to_rgb(hex_code: str):
    hex_code = hex_code.lstrip("#")
    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))


def resolve_font(detected_settings, size=None):
    """Resolves font file from detected VS Code settings or system fallbacks."""
    if size is None:
        size = detected_settings.get("font_size", 14)

    # 1. First priority: detected VS Code font from user's settings
    font_file = detected_settings.get("font_file_path")
    if font_file and os.path.isfile(font_file):
        try:
            return ImageFont.truetype(font_file, size=size)
        except Exception:
            pass

    # 2. Candidate fallbacks
    candidates = [
        "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/cascadiamono.ttf",
        "/usr/share/fonts/TTF/JetBrainsMono-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    ]
    for fp in candidates:
        if os.path.isfile(fp):
            try:
                return ImageFont.truetype(fp, size=size)
            except Exception:
                continue

    return ImageFont.load_default()


def run_interactive_command(cmd: str, inputs: list = None, timeout=15):
    """
    Spawns process with piped I/O, feeds interactive inputs sequentially when
    the process pauses for input, and produces an integrated visual transcript.
    """
    if inputs is None:
        inputs = []

    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        shell=True,
        bufsize=0,
    )

    output_chunks = []
    finished = threading.Event()

    def reader():
        try:
            while True:
                chunk = proc.stdout.read(1)
                if not chunk:
                    break
                output_chunks.append((time.time(), chunk))
        except Exception:
            pass
        finally:
            finished.set()

    t = threading.Thread(target=reader, daemon=True)
    t.start()

    start_time = time.time()
    input_idx = 0
    last_chunk_len = 0
    last_activity = time.time()

    while proc.poll() is None and (time.time() - start_time < timeout):
        curr_len = len(output_chunks)
        if curr_len > last_chunk_len:
            last_chunk_len = curr_len
            last_activity = time.time()
        else:
            if (time.time() - last_activity > 0.18) and (input_idx < len(inputs)):
                inp_str = str(inputs[input_idx])
                input_idx += 1
                inp_bytes = (inp_str + "\n").encode("utf-8")
                output_chunks.append((time.time(), inp_bytes))
                try:
                    proc.stdin.write(inp_bytes)
                    proc.stdin.flush()
                except Exception:
                    pass
                last_activity = time.time()
                last_chunk_len = len(output_chunks)
        time.sleep(0.04)

    if proc.poll() is None:
        while input_idx < len(inputs):
            inp_str = str(inputs[input_idx])
            input_idx += 1
            inp_bytes = (inp_str + "\n").encode("utf-8")
            output_chunks.append((time.time(), inp_bytes))
            try:
                proc.stdin.write(inp_bytes)
                proc.stdin.flush()
            except Exception:
                pass

    try:
        proc.wait(timeout=3)
    except subprocess.TimeoutExpired:
        proc.kill()

    finished.wait(timeout=2)

    sorted_chunks = sorted(output_chunks, key=lambda x: x[0])
    raw_bytes = b"".join(c[1] for c in sorted_chunks)
    transcript = raw_bytes.decode("utf-8", errors="replace")
    return transcript


def sanitize_terminal_output(raw_output: str) -> str:
    """
    Ensures terminal output follows clean styling:
    Converts '----heading----' or '====heading====' into clean 'heading: '
    """
    lines = raw_output.splitlines()
    cleaned = []
    for line in lines:
        # Match e.g. "----Heading----" or "====Heading===="
        m = re.match(r"^[-=_]{2,}\s*([^-_=]+?)\s*[-=_]{2,}$", line.strip())
        if m:
            heading_name = m.group(1).strip()
            cleaned.append(f"{heading_name}:")
        else:
            cleaned.append(line)
    return "\n".join(cleaned)


def clean_display_command(raw_cmd: str) -> str:
    """Simplify verbose file paths for realistic terminal display."""
    cmd = raw_cmd.strip().replace("\\", "/")
    tokens = cmd.split(" ")
    cleaned_tokens = []
    for tok in tokens:
        if "/" in tok and not tok.startswith("-"):
            cleaned_tokens.append(Path(tok).name)
        else:
            cleaned_tokens.append(tok)
    return " ".join(cleaned_tokens)


def render_terminal_screenshot(
    prompt_str: str,
    command_str: str,
    output_text: str,
    out_path: str,
    detected_settings=None,
    shell_title="pwsh",
):
    """
    Renders terminal session output into a crisp VS Code dark terminal PNG image
    matching user's detected font, font size, and active terminal directory.
    """
    if detected_settings is None:
        detected_settings = detect_vscode_settings()

    font_size = detected_settings.get("font_size", 14)
    term_bg = detected_settings.get("terminal_bg_hex", TERM_PALETTE["bg"])

    font = resolve_font(detected_settings, size=font_size)
    bar_font = resolve_font(detected_settings, size=12)

    dummy_img = Image.new("RGB", (100, 100))
    dummy_draw = ImageDraw.Draw(dummy_img)

    bbox_m = dummy_draw.textbbox((0, 0), "M", font=font)
    char_w = max(bbox_m[2] - bbox_m[0], 9)
    line_h = max(bbox_m[3] - bbox_m[1] + 6, 22)

    # Clean headings: replace "----heading----" with "heading: "
    sanitized_output = sanitize_terminal_output(output_text)
    raw_lines = sanitized_output.splitlines()
    wrapped_lines = []
    max_line_len = max(len(prompt_str + command_str), 45)

    for rline in raw_lines:
        rline = rline.replace("\t", "    ")
        if len(rline) <= 95:
            wrapped_lines.append(rline)
            if len(rline) > max_line_len:
                max_line_len = len(rline)
        else:
            chunks = [rline[i : i + 90] for i in range(0, len(rline), 90)]
            for ch in chunks:
                wrapped_lines.append(ch)
                if len(ch) > max_line_len:
                    max_line_len = len(ch)

    total_lines = 1 + len(wrapped_lines)

    top_bar_h = 32
    pad_left = 18
    pad_right = 24
    pad_top = 12
    pad_bottom = 18

    canvas_w = max(int(pad_left + (max_line_len * char_w) + pad_right), 850)
    canvas_h = int(top_bar_h + pad_top + (total_lines * line_h) + pad_bottom)

    img = Image.new("RGB", (canvas_w, canvas_h), hex_to_rgb(term_bg))
    draw = ImageDraw.Draw(img)

    # 1. Top Terminal Tab Bar
    draw.rectangle([(0, 0), (canvas_w, top_bar_h)], fill=hex_to_rgb(TERM_PALETTE["top_bar_bg"]))
    draw.line([(0, top_bar_h - 1), (canvas_w, top_bar_h - 1)], fill=hex_to_rgb(TERM_PALETTE["top_bar_border"]), width=1)

    bar_text = f"TERMINAL  •  {shell_title}"
    draw.text((16, 8), bar_text, font=bar_font, fill=hex_to_rgb(TERM_PALETTE["top_bar_text"]))

    # 2. Prompt & Command Line
    curr_y = top_bar_h + pad_top
    curr_x = pad_left

    draw.text((curr_x, curr_y), prompt_str, font=font, fill=hex_to_rgb(TERM_PALETTE["prompt_ps"]))
    p_bbox = dummy_draw.textbbox((0, 0), prompt_str, font=font)
    curr_x += (p_bbox[2] - p_bbox[0])

    draw.text((curr_x, curr_y), command_str, font=font, fill=hex_to_rgb(TERM_PALETTE["command"]))
    curr_y += line_h

    # 3. Output Lines
    for oline in wrapped_lines:
        draw.text((pad_left, curr_y), oline, font=font, fill=hex_to_rgb(TERM_PALETTE["output_text"]))
        curr_y += line_h

    out_file = Path(out_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(out_file), "PNG", dpi=(150, 150))
    print(f"Terminal screenshot saved successfully: {out_file}")


def main():
    parser = argparse.ArgumentParser(description="Capture interactive command execution and render terminal PNG.")
    parser.add_argument("--cmd", required=True, help="Command to execute")
    parser.add_argument("--display-cmd", required=False, help="Cleaner command string to display in prompt")
    parser.add_argument("--inputs", nargs="*", default=[], help="Sequential lines of stdin input to feed")
    parser.add_argument("--cwd-label", required=False, help="Custom directory path to display in prompt")
    parser.add_argument("--out", required=True, help="Path to output PNG file")
    parser.add_argument("--profile", required=False, help="Path to assets/profile.json")

    args = parser.parse_args()

    # Auto-detect host VS Code settings
    detected = detect_vscode_settings(args.profile)

    # Resolve active terminal directory from host's opened workspace folder
    if args.cwd_label:
        active_cwd = args.cwd_label
    else:
        active_cwd = os.path.abspath(os.getcwd())

    if os.name == "nt":
        shell_title = "pwsh"
        prompt_str = f"PS {active_cwd}> "
    else:
        shell_title = "bash"
        prompt_str = f"user@host:{Path(active_cwd).name}$ "

    display_cmd = args.display_cmd or clean_display_command(args.cmd)

    print(f"Executing: {args.cmd}")
    print(f"Simulating inputs: {args.inputs}")
    transcript = run_interactive_command(args.cmd, inputs=args.inputs)
    print("Execution complete. Output:")
    print(transcript)

    render_terminal_screenshot(
        prompt_str=prompt_str,
        command_str=display_cmd,
        output_text=transcript,
        out_path=args.out,
        detected_settings=detected,
        shell_title=shell_title,
    )


if __name__ == "__main__":
    main()
