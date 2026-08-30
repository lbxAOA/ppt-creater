import sys
from pathlib import Path
import zipfile

sys.path.insert(0, str(Path(__file__).parents[1] / "src" / "ppt_creater"))
from restore_animation_timing import merge_timing  # noqa: E402


def test_merge_timing_restores_animation_without_replacing_edited_text():
    master = b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><a:t>old text</a:t></p:cSld><p:timing><p:tnLst><p:par/></p:tnLst></p:timing></p:sld>'
    edited = b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><a:t>new text</a:t></p:cSld></p:sld>'
    output = merge_timing(master, edited)
    assert b'new text' in output
    assert b'timing' in output
    assert b'par' in output
