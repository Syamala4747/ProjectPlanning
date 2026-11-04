"""Workbook provenance and freshness metadata."""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone
from typing import Any, Dict

APPLICATION_VERSION = "RoadConstruction-planner-1.0"


def _stable(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): _stable(value[k]) for k in sorted(value)}
    if isinstance(value, (list, tuple)):
        return [_stable(item) for item in value]
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def input_summary(project_input: Any) -> Dict[str, Any]:
    return {
        key: _stable(value)
        for key, value in vars(project_input).items()
        if not key.startswith("_") and key != "title"
    }


def file_hash(*paths: str) -> str:
    digest = hashlib.sha256()
    for path in paths:
        digest.update(os.path.abspath(path).encode("utf-8"))
        with open(path, "rb") as stream:
            digest.update(stream.read())
    return digest.hexdigest()


def build_metadata(project_input: Any, guidelines_path: str, assumptions_path: str) -> Dict[str, Any]:
    summary = input_summary(project_input)
    return {
        "application_version": APPLICATION_VERSION,
        "project_type": str(getattr(project_input, "project_type", "road")),
        "input_summary": summary,
        "configuration_hash": file_hash(guidelines_path, assumptions_path),
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
    }


def attach_metadata(workbook: Any, metadata: Dict[str, Any]) -> None:
    payload = json.dumps(metadata, sort_keys=True, separators=(",", ":"))
    workbook.properties.keywords = "RoadConstructionMetadata:" + payload
    workbook.properties.description = (
        "RoadConstruction generated workbook; provenance metadata is stored "
        "in document properties."
    )


def read_metadata(workbook: Any) -> Dict[str, Any]:
    raw = str(getattr(workbook.properties, "keywords", "") or "")
    prefix = "RoadConstructionMetadata:"
    if not raw.startswith(prefix):
        return {}
    try:
        value = json.loads(raw[len(prefix):])
    except json.JSONDecodeError:
        return {}
    return value if isinstance(value, dict) else {}


def freshness_warnings(metadata: Dict[str, Any], project_input: Any,
                       guidelines_path: str, assumptions_path: str) -> list[str]:
    warnings = []
    if not metadata:
        return ["Workbook has no provenance metadata; freshness cannot be verified."]
    if metadata.get("project_type") != getattr(project_input, "project_type", "road"):
        warnings.append("Workbook project type does not match the current project input.")
    stored_inputs = {
        key: value
        for key, value in (metadata.get("input_summary") or {}).items()
        if key != "title"
    }
    current_inputs = input_summary(project_input)
    input_differences = {
        key for key in set(stored_inputs) | set(current_inputs)
        if stored_inputs.get(key, False) != current_inputs.get(key, False)
    }
    if input_differences:
        warnings.append("Workbook project inputs differ from the current project inputs.")
    try:
        current_hash = file_hash(guidelines_path, assumptions_path)
        if metadata.get("configuration_hash") != current_hash:
            warnings.append("Workbook guidelines or assumptions differ from current configuration.")
    except OSError as exc:
        warnings.append(f"Could not verify configuration hash: {exc}")
    return warnings
