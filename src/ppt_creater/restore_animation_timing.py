from __future__ import annotations

import re
import zipfile
from pathlib import Path

TIMING_RE = re.compile(rb"<p:timing(?:\s[^>]*)?>.*?</p:timing\s*>", re.DOTALL)
TRANSITION_RE = re.compile(rb"<p:transition(?:\s[^>]*)?(?:/>|>.*?</p:transition\s*>)", re.DOTALL)


def _replace_part(master_xml: bytes, edited_xml: bytes, pattern: re.Pattern[bytes]) -> bytes:
    """Copy raw XML bytes so extension namespaces and Office markup survive."""
    source = pattern.search(master_xml)
    output = pattern.sub(b"", edited_xml)
    if source is None:
        return output
    # timing/transition are direct children of p:sld. Insert before closing tag.
    close = output.rfind(b"</p:sld>")
    if close < 0:
        raise ValueError("Expected a PowerPoint p:sld root element")
    return output[:close] + source.group(0) + output[close:]


def merge_timing(master_xml: bytes, edited_xml: bytes) -> bytes:
    """Replace only raw timing/transition blocks; preserve edited shape XML byte-for-byte."""
    output = _replace_part(master_xml, edited_xml, TIMING_RE)
    return _replace_part(master_xml, output, TRANSITION_RE)


def restore_animation_parts(master_pptx: Path, edited_pptx: Path, output_pptx: Path) -> int:
    with zipfile.ZipFile(master_pptx, "r") as master, zipfile.ZipFile(edited_pptx, "r") as edited:
        master_parts = set(master.namelist())
        edited_parts = set(edited.namelist())
        replacements: dict[str, bytes] = {}
        slide_parts = sorted(name for name in master_parts & edited_parts if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
        for part in slide_parts:
            replacements[part] = merge_timing(master.read(part), edited.read(part))
        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as output:
            for part in edited.namelist():
                output.writestr(part, replacements.get(part, edited.read(part)))
    return len(replacements)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("master", type=Path)
    parser.add_argument("edited", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    print({"slide_parts": restore_animation_parts(args.master, args.edited, args.output)})
