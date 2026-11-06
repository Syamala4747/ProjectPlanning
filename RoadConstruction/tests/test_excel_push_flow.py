import json
from pathlib import Path
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts import actual_to_entity
from scripts.push_to_db import export_excel_to_json


def test_excel_export_writes_json_only_when_explicitly_requested(tmp_path, monkeypatch):
    workbook = tmp_path / "South.xlsx"
    workbook.touch()
    generated = {
        "project": [{"projectId": "south"}],
        "store": [],
    }
    monkeypatch.setattr(
        actual_to_entity,
        "build_entities",
        lambda path, project_type, seed: generated,
    )

    json_path = export_excel_to_json(str(workbook), "road")

    assert Path(json_path).name == "South_actual_data.json"
    assert json.loads(Path(json_path).read_text(encoding="utf-8")) == {
        "project": [{"projectId": "south"}],
    }


def test_excel_export_rejects_nonempty_unsupported_entity_type(tmp_path, monkeypatch):
    workbook = tmp_path / "South.xlsx"
    workbook.touch()
    monkeypatch.setattr(
        actual_to_entity,
        "build_entities",
        lambda path, project_type, seed: {
            "project": [{"projectId": "south"}],
            "store": [{"storeCode": "S1"}],
        },
    )

    with pytest.raises(ValueError, match="store=1"):
        export_excel_to_json(str(workbook), "road")

    assert not (tmp_path / "South_actual_data.json").exists()
