import importlib.util
import json
from pathlib import Path


SCRIPT = Path(__file__).parents[1] / "oralslides-local/scripts/oralslides.py"
SPEC = importlib.util.spec_from_file_location("oralslides", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(MODULE)


def test_detect_language():
    assert MODULE.detect_language("你好，欢迎") == "zh"
    assert MODULE.detect_language("Welcome to the deck") == "en"


def test_validate_scripts_rejects_missing_narration(tmp_path):
    (tmp_path / "manifest.json").write_text(
        json.dumps({"slides": [{"slideNumber": 1}]})
    )
    (tmp_path / "scripts.json").write_text(
        json.dumps({"slides": [{"slideNumber": 1, "script": ""}]})
    )
    try:
        MODULE.validate_scripts(tmp_path)
    except MODULE.OralSlidesError as error:
        assert "empty narration" in str(error)
    else:
        raise AssertionError("expected validation error")
