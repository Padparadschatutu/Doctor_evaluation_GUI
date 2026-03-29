from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, session, url_for


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = Path("/home/yfan/benchmark/doctor/merged_selected_cases_with_source_shuffled.xlsx")
DEFAULT_IMAGE_ROOT = Path("/home/yfan/benchmark/bone/select_all_jpg")
DEFAULT_IMAGE_COL = "image_paths"
DEFAULT_TEXT_COL = "analysis_result"
OUTPUTS_DIR = BASE_DIR / "outputs"

ANNOTATION_KEY_SEP = "||"


def _patch_broken_colorama_for_werkzeug() -> None:
    """
    Work around a broken `colorama` installation that imports as a namespace package
    (missing `AnsiToWin32`), which crashes Werkzeug's startup logger.
    """
    try:
        import colorama  # type: ignore
    except Exception:
        return

    if hasattr(colorama, "AnsiToWin32"):
        return

    class AnsiToWin32:  # noqa: N801 - match Colorama's public API name
        def __init__(self, stream):
            self.stream = stream

        def __getattr__(self, name):
            return getattr(self.stream, name)

        def write(self, text):
            return self.stream.write(text)

        def flush(self):
            flush = getattr(self.stream, "flush", None)
            if flush is not None:
                return flush()
            return None

    try:
        colorama.AnsiToWin32 = AnsiToWin32  # type: ignore[attr-defined]
    except Exception:
        sys.modules.pop("colorama", None)


_patch_broken_colorama_for_werkzeug()

RATING_FIELDS: List[Tuple[str, str, List[str]]] = [
    ("overall_acceptability", "整体可接受性", ["yes", "no"]),
    ("factuality", "事实性", ["1", "2", "3", "4", "5"]),
    ("language_adaptation", "语言适配性", ["1", "2", "3", "4", "5"]),
]
NOTE_FIELD = "note"
NOTE_LABEL = "备注"
HIGHLIGHT_FIELD = "highlight_terms"
HIGHLIGHT_LABEL = "关注高亮"

BOUNDARY_PATTERNS: List[str] = [
    r"(?:边界|边缘|分界)[^，。；;\n]{0,8}(?:清楚|清晰|锐利|光整|光滑|欠清|不清|不清晰|模糊|毛糙|欠光整|欠光滑)",
    r"(?:清楚|清晰|锐利|光整|光滑|欠清|不清|不清晰|模糊|毛糙|欠光整|欠光滑)[^，。；;\n]{0,8}(?:边界|边缘|分界)",
    r"(?:边缘分叶|分叶征|边缘毛刺|毛刺征|浸润性边界)",
]

BOUNDARY_SENTENCE_PATTERNS: List[str] = BOUNDARY_PATTERNS + [
    r"(?:边界[^。；;\n]{0,12}地图样)",
    r"(?:地图样[^。；;\n]{0,24}硬化边)",
    r"(?:完整且清晰锐利的硬化边)",
    r"(?:过渡区窄)",
]

ENDOSTEAL_PATTERNS: List[str] = [
    r"(?:骨内膜|内膜)[^，。；;\n]{0,12}(?:侵蚀|吸收|破坏|受压|变薄)[^，。；;\n]{0,8}(?:轻度|中度|重度|明显)?",
    r"(?:轻度|中度|重度|明显|无明显)?[^，。；;\n]{0,8}(?:骨内膜|内膜)(?:侵蚀|吸收|破坏|受压|变薄)",
]

PERIOSTEAL_PATTERNS: List[str] = [
    r"(?:骨膜反应|层状骨膜反应|葱皮样骨膜反应|放射状骨膜反应|针状骨膜反应|Codman三角|袖套样骨膜反应)",
]

SOFT_TISSUE_MASS_PATTERNS: List[str] = [
    r"(?:软组织肿块|软组织包块|软组织肿胀|软组织肿块形成|邻近软组织肿块|周围软组织肿块|软组织影)",
    r"(?:软组织[^。；;\n]{0,12}(?:未见|无)[^，。；;\n]{0,8}(?:明确|明显)?[^，。；;\n]{0,8}(?:肿块影|肿块|包块|异常影))",
    r"(?:(?:未见|无)[^，。；;\n]{0,8}(?:明确|明显)?[^，。；;\n]{0,8}软组织(?:肿块影|肿块|包块|异常影))",
]

PATHOLOGIC_FRACTURE_PATTERNS: List[str] = [
    r"(?:病理性骨折)",
]

INTERNAL_FEATURE_PATTERNS: List[str] = [
    r"(?:未见明显[^，。；;\n]{0,8}基质(?:钙化)?(?:影)?)",
    r"(?:无明显[^，。；;\n]{0,8}基质(?:钙化)?(?:影)?)",
    r"(?:软骨样基质(?:钙化)?(?:影)?)",
    r"(?:骨样基质(?:钙化)?(?:影)?)",
    r"(?:纤维(?:性)?基质(?:钙化)?(?:影)?)",
    r"(?:混合基质(?:钙化)?(?:影)?)",
    r"(?:钙化基质(?:影)?)",
    r"(?:(?:病变)?内部[^。；;\n]{0,18}(?:多房(?:/分隔样)?改变|多房样改变|分隔样改变|可见分隔|透亮(?:为主)?))",
    r"(?:(?:病变)?内部[^。；;\n]{0,24}(?:多发不规则斑片状、云絮状高密度影|斑片状、云絮状高密度影|云絮状高密度影|斑片状高密度影))",
]

ANATOMY_PATTERNS: List[str] = [
    r"(?:左侧|右侧|双侧)?(?:股骨颈|股骨头|股骨粗隆间|股骨干|股骨远端|股骨近端|股骨)",
    r"(?:左侧|右侧|双侧)?(?:肱骨头|肱骨外科颈|肱骨干|肱骨远端|肱骨近端|肱骨)",
    r"(?:左侧|右侧|双侧)?(?:桡骨头|桡骨远端|桡骨近端|桡骨|尺骨鹰嘴|尺骨远端|尺骨近端|尺骨)",
    r"(?:左侧|右侧|双侧)?(?:胫骨平台|胫骨近端|胫骨远端|胫骨|腓骨头|腓骨远端|腓骨近端|腓骨)",
    r"(?:左侧|右侧|双侧)?(?:跟骨|距骨|跖骨|趾骨|掌骨|指骨|髌骨|锁骨|肩胛骨|肋骨)",
    r"(?:左侧|右侧|双侧)?(?:骨盆|髂骨|耻骨|坐骨|骶骨|髋臼|髋关节|膝关节|肩关节|踝关节|肘关节|腕关节)",
    r"(?:左侧|右侧|双侧)?(?:胸椎|腰椎|颈椎|椎体|椎弓根|椎间隙|关节面|骨端|骨干|骨皮质|髓腔|软组织)",
]


@dataclass
class AppState:
    operator_id: Optional[str] = None
    operator_slug: Optional[str] = None
    excel_path: Optional[Path] = None
    sheet_name: Optional[str] = None
    df: Optional[pd.DataFrame] = None
    id_col: Optional[str] = None
    image_col: Optional[str] = None
    text_col: Optional[str] = None
    highlight_col: Optional[str] = None
    boundary_col: Optional[str] = None
    lesion_structure_col: Optional[str] = None
    anatomy_col: Optional[str] = None
    filter_specs: List[Tuple[str, str]] = field(default_factory=list)
    image_root: Optional[Path] = None
    global_highlight_terms: str = ""
    annotation_path: Optional[Path] = None
    annotations: Dict[str, Dict[str, Any]] = field(default_factory=dict)

    def is_loaded(self) -> bool:
        return self.df is not None and self.excel_path is not None

    def is_configured(self) -> bool:
        return (
            self.is_loaded()
            and self.image_col is not None
            and self.text_col is not None
            and self.image_col in list(self.df.columns)  # type: ignore[union-attr]
            and self.text_col in list(self.df.columns)  # type: ignore[union-attr]
        )


SESSION_STATES: Dict[str, AppState] = {}
SESSION_STATES_LOCK = Lock()
BOOTSTRAP_STATE: Optional[AppState] = None


def _to_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if pd.isna(value):
        return ""
    return str(value)


def _load_annotations(path: Path) -> Dict[str, Dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data  # type: ignore[return-value]
        return {}
    except Exception:
        return {}


def _save_annotations(path: Path, data: Dict[str, Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_name(f".{path.name}.tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(path)


def _normalize_operator_id(value: str) -> Tuple[str, str]:
    operator_id = " ".join(value.split()).strip()
    if not operator_id:
        raise ValueError("操作者不能为空")

    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", operator_id).strip("._-").lower()
    digest = hashlib.sha1(operator_id.encode("utf-8")).hexdigest()[:8]
    operator_slug = f"{safe}-{digest}" if safe else f"operator-{digest}"
    return operator_id, operator_slug


def _get_or_create_state(operator_id: str, operator_slug: str) -> AppState:
    with SESSION_STATES_LOCK:
        state = SESSION_STATES.get(operator_slug)
        if state is None:
            state = AppState(operator_id=operator_id, operator_slug=operator_slug)
            SESSION_STATES[operator_slug] = state
        else:
            state.operator_id = operator_id
            state.operator_slug = operator_slug
    return state


def _set_operator_session(operator_id_raw: str) -> AppState:
    operator_id, operator_slug = _normalize_operator_id(operator_id_raw)
    session["operator_id"] = operator_id
    session["operator_slug"] = operator_slug
    return _get_or_create_state(operator_id, operator_slug)


def _current_state() -> Optional[AppState]:
    operator_id = _to_str(session.get("operator_id")).strip()
    operator_slug = _to_str(session.get("operator_slug")).strip()
    if not operator_id or not operator_slug:
        return None

    state = _get_or_create_state(operator_id, operator_slug)
    _apply_bootstrap_if_needed(state)
    return state


def _operator_output_dir(state: AppState) -> Path:
    if not state.operator_slug:
        raise ValueError("缺少操作者标识")
    return OUTPUTS_DIR / state.operator_slug


def _output_stem(state: AppState) -> str:
    if state.excel_path is None or not state.operator_slug:
        raise ValueError("导出前需要已加载 Excel 和操作者")
    return f"{state.excel_path.stem}_annotations_{state.operator_slug}"


def _annotation_path_for_state(state: AppState, excel_path: Path) -> Path:
    if not state.operator_slug:
        raise ValueError("缺少操作者标识")
    return _operator_output_dir(state) / f"{excel_path.stem}_annotations_{state.operator_slug}.json"


def _apply_bootstrap_if_needed(state: AppState) -> None:
    if BOOTSTRAP_STATE is None or state.is_loaded():
        return

    state.excel_path = BOOTSTRAP_STATE.excel_path
    state.sheet_name = BOOTSTRAP_STATE.sheet_name
    state.df = BOOTSTRAP_STATE.df
    state.id_col = BOOTSTRAP_STATE.id_col
    state.image_col = BOOTSTRAP_STATE.image_col
    state.text_col = BOOTSTRAP_STATE.text_col
    state.highlight_col = BOOTSTRAP_STATE.highlight_col
    state.boundary_col = BOOTSTRAP_STATE.boundary_col
    state.lesion_structure_col = BOOTSTRAP_STATE.lesion_structure_col
    state.anatomy_col = BOOTSTRAP_STATE.anatomy_col
    state.image_root = BOOTSTRAP_STATE.image_root
    state.global_highlight_terms = BOOTSTRAP_STATE.global_highlight_terms
    if state.excel_path is not None:
        state.annotation_path = _annotation_path_for_state(state, state.excel_path)
        state.annotations = _load_annotations(state.annotation_path)


def _apply_default_columns(state: AppState) -> None:
    if state.df is None:
        return

    columns = {str(c) for c in state.df.columns}
    if DEFAULT_IMAGE_COL in columns:
        state.image_col = DEFAULT_IMAGE_COL
    if DEFAULT_TEXT_COL in columns:
        state.text_col = DEFAULT_TEXT_COL


def _resolve_excel_path(path_str: str) -> Path:
    p = Path(path_str).expanduser()
    if not p.is_absolute():
        p = (Path.cwd() / p).resolve()
    return p


def _resolve_dir_path(path_str: str) -> Path:
    return _resolve_excel_path(path_str)


def _load_excel(excel_path: Path, sheet_name: Optional[str]) -> pd.DataFrame:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")
    if sheet_name:
        return pd.read_excel(excel_path, sheet_name=sheet_name)
    return pd.read_excel(excel_path)


def _row_key(state: AppState, row_index: int) -> str:
    if not state.is_loaded():
        return str(row_index)
    if state.id_col and state.id_col in list(state.df.columns):  # type: ignore[union-attr]
        value = state.df.iloc[row_index][state.id_col]  # type: ignore[union-attr]
        s = _to_str(value).strip()
        if s:
            return s
    return str(row_index)


def _annotation_key(state: AppState, row_index: int) -> str:
    """
    Annotation key is scoped by (row_key, text_col) so switching displayed column
    won't overwrite previous ratings from another model/column.
    """
    row_key = _row_key(state, row_index)
    text_col = state.text_col or ""
    return f"{row_key}{ANNOTATION_KEY_SEP}{text_col}"


def _parse_annotation_key(key: str) -> Tuple[str, Optional[str]]:
    if ANNOTATION_KEY_SEP not in key:
        return key, None
    row_key, text_col = key.split(ANNOTATION_KEY_SEP, 1)
    text_col = text_col.strip()
    return row_key, (text_col or None)


def _split_image_paths(cell: Any) -> List[str]:
    raw = _to_str(cell).strip()
    if not raw:
        return []
    for sep in ["\n", ";", "|"]:
        if sep in raw:
            parts = [p.strip() for p in raw.replace("\r", "").split(sep)]
            return [p for p in parts if p]
    return [raw]


@dataclass(frozen=True)
class ResolvedImage:
    raw: str
    path: Path
    exists: bool
    is_file: bool
    resolved_via: str


def _try_resolve_under_image_root(raw_path: Path, image_root: Path) -> Optional[Path]:
    if not image_root.exists() or not image_root.is_dir():
        return None

    parts = list(raw_path.parts)
    if raw_path.is_absolute() and parts and parts[0] == "/":
        parts = parts[1:]
    if not parts:
        return None

    # Prefer longer suffix matches first (more specific), then fall back to shorter ones.
    max_suffix = min(6, len(parts))
    for n in range(max_suffix, 0, -1):
        candidate = image_root.joinpath(*parts[-n:])
        if candidate.exists():
            return candidate
    return None


def _get_image_paths_for_row(state: AppState, row_index: int) -> List[ResolvedImage]:
    if not state.is_configured():
        return []
    base_dir = state.excel_path.parent if state.excel_path else Path.cwd()
    cell = state.df.iloc[row_index][state.image_col]  # type: ignore[union-attr]
    candidates = _split_image_paths(cell)
    out: List[ResolvedImage] = []
    for raw in candidates:
        p = Path(raw).expanduser()
        resolved_via = "as-is"

        if p.is_absolute():
            resolved = p
            if not resolved.exists() and state.image_root is not None:
                alt = _try_resolve_under_image_root(resolved, state.image_root)
                if alt is not None:
                    resolved = alt
                    resolved_via = "image_root(suffix)"
        else:
            resolved = (base_dir / p).resolve()
            resolved_via = "excel_dir"
            if not resolved.exists() and state.image_root is not None:
                alt = (state.image_root / p).resolve()
                if alt.exists():
                    resolved = alt
                    resolved_via = "image_root(rel)"
                else:
                    alt2 = _try_resolve_under_image_root(p, state.image_root)
                    if alt2 is not None:
                        resolved = alt2
                        resolved_via = "image_root(suffix)"

        out.append(
            ResolvedImage(
                raw=raw,
                path=resolved,
                exists=resolved.exists(),
                is_file=resolved.is_file(),
                resolved_via=resolved_via,
            )
        )

    return out


def _annotation_for_row(state: AppState, row_index: int) -> Dict[str, Any]:
    # Prefer the new (row_key||text_col) key; fall back to legacy row_key-only
    # annotations for backward compatibility.
    scoped = _annotation_key(state, row_index)
    if scoped in state.annotations:
        return state.annotations.get(scoped, {})
    legacy = _row_key(state, row_index)
    return state.annotations.get(legacy, {})


def _active_filter_specs(state: AppState) -> List[Tuple[str, str]]:
    if state.df is None:
        return []

    columns = set(list(state.df.columns))
    normalized: List[Tuple[str, str]] = []
    seen = set()
    for col, value in state.filter_specs:
        filter_col = _to_str(col).strip()
        filter_value = _to_str(value).strip()
        if not filter_col or not filter_value or filter_col not in columns:
            continue
        key = (filter_col, filter_value)
        if key in seen:
            continue
        seen.add(key)
        normalized.append(key)
    return normalized


def _filtered_row_indices(state: AppState) -> List[int]:
    if state.df is None:
        return []

    total = len(state.df)
    filter_specs = _active_filter_specs(state)
    if not filter_specs:
        return list(range(total))

    filtered: List[int] = []
    for idx in range(total):
        if all(
            _to_str(state.df.iloc[idx][filter_col]).strip() == filter_value  # type: ignore[union-attr]
            for filter_col, filter_value in filter_specs
        ):
            filtered.append(idx)
    return filtered


def _resolve_display_row_index(state: AppState, display_row_index: int) -> Tuple[List[int], int]:
    filtered_indices = _filtered_row_indices(state)
    if display_row_index < 0 or display_row_index >= len(filtered_indices):
        abort(404)
    return filtered_indices, filtered_indices[display_row_index]


def _display_row_index_for_source(state: AppState, source_row_index: int) -> Optional[int]:
    filtered_indices = _filtered_row_indices(state)
    try:
        return filtered_indices.index(source_row_index)
    except ValueError:
        return None


def _split_highlight_terms_text(value: str) -> List[str]:
    raw = _to_str(value).strip()
    if not raw:
        return []
    if re.search(r"[\n;；|]", raw):
        parts = re.split(r"[\n;；|]+", raw)
    else:
        comma_parts = [part.strip() for part in re.split(r"[,，]+", raw) if part.strip()]
        if len(comma_parts) > 1 and all(len(part) <= 12 for part in comma_parts):
            parts = comma_parts
        else:
            parts = [raw]
    out: List[str] = []
    seen = set()
    for part in parts:
        item = part.strip()
        if not item or item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def _merge_highlight_terms_text(*values: str) -> str:
    merged: List[str] = []
    seen = set()
    for value in values:
        for item in _split_highlight_terms_text(value):
            if item in seen:
                continue
            seen.add(item)
            merged.append(item)
    return "\n".join(merged)


def _extract_matches_by_patterns(text: str, patterns: List[str]) -> List[str]:
    found: List[str] = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            value = _to_str(match.group(0)).strip("，。；;：:、 ")
            if not value:
                continue
            if any(value in existing for existing in found):
                continue
            found = [existing for existing in found if existing not in value]
            found.append(value)
    return found


def _extract_first_match_by_patterns(text: str, patterns: List[str]) -> str:
    first_start: Optional[int] = None
    first_value = ""
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            value = _to_str(match.group(0)).strip("，。；;：:、 ")
            if not value:
                continue
            if first_start is None or match.start() < first_start:
                first_start = match.start()
                first_value = value
    return first_value


def _extract_sentence_matches_by_patterns(text: str, patterns: List[str]) -> List[str]:
    found: List[str] = []
    seen = set()
    clauses = [part.strip("，。；;：:、 ") for part in re.split(r"[。；;\n]+", text) if part.strip()]
    for clause in clauses:
        if any(re.search(pattern, clause, flags=re.IGNORECASE) for pattern in patterns):
            if clause not in seen:
                seen.add(clause)
                found.append(clause)
    return found


def _extract_rule_highlights(text: str) -> Dict[str, str]:
    return {
        "boundary": "\n".join(_extract_sentence_matches_by_patterns(text, BOUNDARY_SENTENCE_PATTERNS)),
        "endosteal": "\n".join(_extract_matches_by_patterns(text, ENDOSTEAL_PATTERNS)),
        "periosteal": "\n".join(_extract_sentence_matches_by_patterns(text, PERIOSTEAL_PATTERNS)),
        "soft_tissue": "\n".join(_extract_sentence_matches_by_patterns(text, SOFT_TISSUE_MASS_PATTERNS)),
        "pathologic_fracture": "\n".join(_extract_sentence_matches_by_patterns(text, PATHOLOGIC_FRACTURE_PATTERNS)),
        "internal_feature": "\n".join(_extract_matches_by_patterns(text, INTERNAL_FEATURE_PATTERNS)),
        "anatomy": _extract_first_match_by_patterns(text, ANATOMY_PATTERNS),
    }


def _annotated_count_for_current_text_col(state: AppState) -> int:
    if not state.is_configured() or not state.text_col:
        return len(state.annotations)
    suffix = f"{ANNOTATION_KEY_SEP}{state.text_col}"
    return sum(1 for k in state.annotations.keys() if k.endswith(suffix))


def _last_review_row_index(state: AppState) -> int:
    filtered_indices = _filtered_row_indices(state)
    if state.df is None or len(filtered_indices) == 0 or not state.annotations:
        return 0

    row_key_to_first_index = {_row_key(state, i): i for i in range(len(state.df))}
    filtered_index_set = set(filtered_indices)
    best_row_index = filtered_indices[0]
    best_updated_at: Optional[datetime] = None

    for ann_key, ann in state.annotations.items():
        _, text_col = _parse_annotation_key(ann_key)
        if state.text_col and text_col not in (None, state.text_col):
            continue

        row_index_value = ann.get("row_index")
        row_index: Optional[int] = None
        if isinstance(row_index_value, int) and 0 <= row_index_value < len(state.df):
            row_index = row_index_value
        else:
            row_key = _to_str(ann.get("row_key")).strip()
            if row_key:
                row_index = row_key_to_first_index.get(row_key)

        if row_index is None:
            continue
        if row_index not in filtered_index_set:
            continue

        updated_at_raw = _to_str(ann.get("updated_at")).strip()
        try:
            updated_at = datetime.fromisoformat(updated_at_raw) if updated_at_raw else None
        except Exception:
            updated_at = None

        if updated_at is None:
            if best_updated_at is None and row_index >= best_row_index:
                best_row_index = row_index
            continue

        if best_updated_at is None or updated_at > best_updated_at:
            best_updated_at = updated_at
            best_row_index = row_index

    display_row_index = _display_row_index_for_source(state, best_row_index)
    if display_row_index is None:
        return 0
    return display_row_index


def _build_annotation_payload(state: AppState, row_index: int, form_data: Any) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "row_index": row_index,
        "row_key": _row_key(state, row_index),
        "text_col": state.text_col,
        "operator_id": state.operator_id,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    for field, _, allowed in RATING_FIELDS:
        value = _to_str(form_data.get(field)).strip()
        if value and value not in allowed:
            raise ValueError(f"字段 {field} 值非法: {value}")
        payload[field] = value

    payload[NOTE_FIELD] = _to_str(form_data.get(NOTE_FIELD)).strip()
    payload[HIGHLIGHT_FIELD] = _to_str(form_data.get(HIGHLIGHT_FIELD)).strip()

    return payload


def _persist_annotation(state: AppState, row_index: int, payload: Dict[str, Any]) -> Dict[str, Any]:
    key = _annotation_key(state, row_index)
    state.annotations[key] = payload
    if state.annotation_path is not None:
        _save_annotations(state.annotation_path, state.annotations)
    return payload


def _build_export_dataframe(state: AppState) -> pd.DataFrame:
    if state.df is None:
        raise ValueError("尚未加载 Excel")

    df = state.df.copy()

    row_keys = [_row_key(state, i) for i in range(len(df))]
    row_key_to_first_index: Dict[str, int] = {}
    for idx, rk in enumerate(row_keys):
        row_key_to_first_index.setdefault(rk, idx)

    model_source_col = "模型来源"
    annotator_col = "标注人"
    if model_source_col not in df.columns:
        df[model_source_col] = ""
    if annotator_col not in df.columns:
        df[annotator_col] = ""

    def _short_source(value: Any) -> str:
        s = _to_str(value).strip()
        if not s:
            return ""
        if "/" in s or "\\" in s:
            try:
                return Path(s).name
            except Exception:
                return s
        return s

    per_row_sources: Dict[int, set[str]] = {}
    if "source_file" in df.columns:
        for i in range(len(df)):
            src = _short_source(df.at[i, "source_file"])
            if src:
                per_row_sources.setdefault(i, set()).add(src)

    annotated_text_cols: List[Optional[str]] = []
    for k in state.annotations.keys():
        _, text_col = _parse_annotation_key(k)
        annotated_text_cols.append(text_col)
    unique_text_cols = sorted({tc for tc in annotated_text_cols if tc is not None})

    generic_text_cols = {"analysis_result", "text", "report", "result", "output"}

    def _score_col(label: str, text_col: Optional[str]) -> str:
        if text_col is None:
            return label
        return f"{label} ({text_col})"

    for _, label, _ in RATING_FIELDS:
        if label not in df.columns:
            df[label] = ""
    if NOTE_LABEL not in df.columns:
        df[NOTE_LABEL] = ""
    if HIGHLIGHT_LABEL not in df.columns:
        df[HIGHLIGHT_LABEL] = ""
    for text_col in unique_text_cols:
        for _, label, _ in RATING_FIELDS:
            col = _score_col(label, text_col)
            if col not in df.columns:
                df[col] = ""
        note_col = _score_col(NOTE_LABEL, text_col)
        if note_col not in df.columns:
            df[note_col] = ""
        highlight_col = _score_col(HIGHLIGHT_LABEL, text_col)
        if highlight_col not in df.columns:
            df[highlight_col] = ""

    for ann_key, ann in state.annotations.items():
        rk, text_col = _parse_annotation_key(ann_key)

        row_index_value = ann.get("row_index")
        target_row: Optional[int] = None
        if isinstance(row_index_value, int) and 0 <= row_index_value < len(df):
            target_row = row_index_value
        else:
            target_row = row_key_to_first_index.get(rk)
        if target_row is None:
            continue

        if text_col is not None:
            tc = str(text_col).strip()
            if tc and tc.lower() not in generic_text_cols:
                per_row_sources.setdefault(target_row, set()).add(tc)

        df.at[target_row, annotator_col] = _to_str(ann.get("operator_id")) or _to_str(state.operator_id)
        for field, label, _ in RATING_FIELDS:
            col = _score_col(label, text_col)
            if col not in df.columns:
                df[col] = ""
            df.at[target_row, col] = _to_str(ann.get(field))
        note_col = _score_col(NOTE_LABEL, text_col)
        if note_col not in df.columns:
            df[note_col] = ""
        df.at[target_row, note_col] = _to_str(ann.get(NOTE_FIELD))
        highlight_col = _score_col(HIGHLIGHT_LABEL, text_col)
        if highlight_col not in df.columns:
            df[highlight_col] = ""
        df.at[target_row, highlight_col] = _to_str(ann.get(HIGHLIGHT_FIELD))

    for row_idx, sources in per_row_sources.items():
        df.at[row_idx, model_source_col] = ",".join(sorted(sources))

    return df


def _build_export_archive(state: AppState) -> Tuple[io.BytesIO, str]:
    if state.annotation_path is None:
        raise ValueError("标注文件路径不存在")

    export_df = _build_export_dataframe(state)
    output_stem = _output_stem(state)
    json_name = f"{output_stem}.json"
    csv_name = f"{output_stem}.csv"
    xlsx_name = f"{output_stem}.xlsx"
    zip_name = f"{output_stem}.zip"

    json_bytes = json.dumps(state.annotations, ensure_ascii=False, indent=2).encode("utf-8")
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")

    xlsx_buffer = io.BytesIO()
    with pd.ExcelWriter(xlsx_buffer) as writer:
        export_df.to_excel(writer, index=False)
    xlsx_buffer.seek(0)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(json_name, json_bytes)
        zf.writestr(csv_name, csv_bytes)
        zf.writestr(xlsx_name, xlsx_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer, zip_name


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = "vlm-anysis-gui-dev"

    @app.get("/")
    def index():
        state = _current_state()
        force_home = _to_str(request.args.get("force_home")).strip() == "1"
        if not force_home and state is not None and state.is_configured():
            return redirect(url_for("review", row_index=_last_review_row_index(state)))
        return render_template(
            "index.html",
            state=state,
            current_operator=_to_str(session.get("operator_id")).strip(),
            default_excel_path=str(BOOTSTRAP_STATE.excel_path) if BOOTSTRAP_STATE and BOOTSTRAP_STATE.excel_path else "",
            default_image_root=str(BOOTSTRAP_STATE.image_root) if BOOTSTRAP_STATE and BOOTSTRAP_STATE.image_root else str(DEFAULT_IMAGE_ROOT),
        )

    @app.post("/load")
    def load():
        operator_raw = _to_str(request.form.get("operator_id")).strip() or _to_str(session.get("operator_id")).strip()
        if not operator_raw:
            flash("请先填写操作者。")
            return redirect(url_for("index"))

        try:
            state = _set_operator_session(operator_raw)
        except ValueError as e:
            flash(str(e))
            return redirect(url_for("index"))

        excel_path_raw = _to_str(request.form.get("excel_path")).strip()
        sheet_name = _to_str(request.form.get("sheet_name")).strip() or None
        image_root_raw = _to_str(request.form.get("image_root")).strip()
        if not excel_path_raw and BOOTSTRAP_STATE and BOOTSTRAP_STATE.excel_path:
            excel_path_raw = str(BOOTSTRAP_STATE.excel_path)
        if not excel_path_raw:
            flash("请填写 Excel 路径。")
            return redirect(url_for("index"))

        try:
            excel_path = _resolve_excel_path(excel_path_raw)
            df = _load_excel(excel_path, sheet_name)
        except Exception as e:
            flash(f"加载失败: {e}")
            return redirect(url_for("index"))

        state.excel_path = excel_path
        state.sheet_name = sheet_name
        state.df = df
        state.id_col = None
        state.image_col = None
        state.text_col = None
        state.highlight_col = None
        state.boundary_col = None
        state.lesion_structure_col = None
        state.anatomy_col = None
        state.filter_specs = []
        _apply_default_columns(state)
        if image_root_raw:
            try:
                image_root = _resolve_dir_path(image_root_raw)
                if not image_root.exists() or not image_root.is_dir():
                    flash(f"图片根目录无效，将忽略：{image_root}")
                    state.image_root = None
                else:
                    state.image_root = image_root
            except Exception:
                state.image_root = None
        else:
            state.image_root = None
        state.annotation_path = _annotation_path_for_state(state, excel_path)
        state.annotations = _load_annotations(state.annotation_path)

        flash(f"操作者 {state.operator_id} 已加载: {excel_path}（共 {len(df)} 行）")
        return redirect(url_for("config"))

    @app.get("/config")
    def config():
        state = _current_state()
        if state is None:
            flash("请先填写操作者并加载 Excel。")
            return redirect(url_for("index"))
        if not state.is_loaded():
            return redirect(url_for("index"))
        columns = [str(c) for c in state.df.columns]  # type: ignore[union-attr]
        return render_template(
            "config.html",
            state=state,
            current_operator=_to_str(session.get("operator_id")).strip(),
            columns=columns,
        )

    @app.post("/config")
    def config_post():
        state = _current_state()
        if state is None:
            flash("请先填写操作者并加载 Excel。")
            return redirect(url_for("index"))
        if not state.is_loaded():
            return redirect(url_for("index"))

        id_col = _to_str(request.form.get("id_col")).strip() or None
        image_col = _to_str(request.form.get("image_col")).strip() or None
        text_col = _to_str(request.form.get("text_col")).strip() or None
        highlight_col = _to_str(request.form.get("highlight_col")).strip() or None
        boundary_col = _to_str(request.form.get("boundary_col")).strip() or None
        lesion_structure_col = _to_str(request.form.get("lesion_structure_col")).strip() or None
        anatomy_col = _to_str(request.form.get("anatomy_col")).strip() or None
        image_root_raw = _to_str(request.form.get("image_root")).strip() or None
        global_highlight_terms = _to_str(request.form.get("global_highlight_terms")).strip()

        columns = {str(c) for c in state.df.columns}  # type: ignore[union-attr]
        if image_col is None or image_col not in columns:
            flash("请选择有效的【图像路径列】。")
            return redirect(url_for("config"))
        if text_col is None or text_col not in columns:
            flash("请选择有效的【显示内容列】。")
            return redirect(url_for("config"))
        if id_col is not None and id_col not in columns:
            flash("【ID列】无效，将忽略。")
            id_col = None
        if highlight_col is not None and highlight_col not in columns:
            flash("【每病例自动高亮列】无效，将忽略。")
            highlight_col = None
        if boundary_col is not None and boundary_col not in columns:
            flash("【边界高亮列】无效，将忽略。")
            boundary_col = None
        if lesion_structure_col is not None and lesion_structure_col not in columns:
            flash("【病变结构高亮列】无效，将忽略。")
            lesion_structure_col = None
        if anatomy_col is not None and anatomy_col not in columns:
            flash("【解剖部位高亮列】无效，将忽略。")
            anatomy_col = None
        filter_specs: List[Tuple[str, str]] = []
        for idx in range(1, 4):
            filter_col = _to_str(request.form.get(f"filter_col_{idx}")).strip()
            filter_value = _to_str(request.form.get(f"filter_value_{idx}")).strip()
            if not filter_col and not filter_value:
                continue
            if not filter_col:
                flash(f"【筛选条件 {idx}】缺少筛选列，将忽略。")
                continue
            if filter_col not in columns:
                flash(f"【筛选条件 {idx}】筛选列无效，将忽略。")
                continue
            if not filter_value:
                flash(f"【筛选条件 {idx}】缺少筛选值，将忽略。")
                continue
            filter_specs.append((filter_col, filter_value))

        state.id_col = id_col
        state.image_col = image_col
        state.text_col = text_col
        state.highlight_col = highlight_col
        state.boundary_col = boundary_col
        state.lesion_structure_col = lesion_structure_col
        state.anatomy_col = anatomy_col
        state.filter_specs = filter_specs
        state.global_highlight_terms = global_highlight_terms
        if image_root_raw:
            try:
                image_root = _resolve_dir_path(image_root_raw)
                if not image_root.exists() or not image_root.is_dir():
                    flash(f"图片根目录无效，将忽略：{image_root}")
                    state.image_root = None
                else:
                    state.image_root = image_root
            except Exception:
                state.image_root = None
        else:
            state.image_root = None

        if not _filtered_row_indices(state):
            flash("筛选后没有匹配的病例，请重新设置筛选条件。")
            return redirect(url_for("config"))

        return redirect(url_for("review", row_index=_last_review_row_index(state)))

    @app.get("/review/<int:row_index>")
    def review(row_index: int):
        state = _current_state()
        if state is None:
            flash("请先填写操作者并加载 Excel。")
            return redirect(url_for("index"))
        if not state.is_loaded():
            return redirect(url_for("index"))
        if not state.is_configured():
            return redirect(url_for("config"))
        filtered_indices, source_row_index = _resolve_display_row_index(state, row_index)

        row = state.df.iloc[source_row_index]  # type: ignore[union-attr]
        key = _row_key(state, source_row_index)
        text_value = _to_str(row[state.text_col])  # type: ignore[index]
        images = _get_image_paths_for_row(state, source_row_index)
        annotated_count = _annotated_count_for_current_text_col(state)

        existing = _annotation_for_row(state, source_row_index)
        existing_values = {k: _to_str(existing.get(k)) for k, _, _ in RATING_FIELDS}
        existing_note = _to_str(existing.get(NOTE_FIELD))
        global_highlight_terms = _to_str(state.global_highlight_terms)
        auto_highlight_terms = ""
        boundary_highlight_terms = ""
        anatomy_highlight_terms = ""
        endosteal_highlight_terms = ""
        periosteal_highlight_terms = ""
        soft_tissue_highlight_terms = ""
        pathologic_fracture_highlight_terms = ""
        internal_feature_highlight_terms = ""
        smart_highlight_used = False
        if state.highlight_col and state.highlight_col in list(state.df.columns):  # type: ignore[union-attr]
            auto_highlight_terms = _to_str(row[state.highlight_col])  # type: ignore[index]
        if state.boundary_col and state.boundary_col in list(state.df.columns):  # type: ignore[union-attr]
            boundary_highlight_terms = _to_str(row[state.boundary_col])  # type: ignore[index]
        if state.anatomy_col and state.anatomy_col in list(state.df.columns):  # type: ignore[union-attr]
            anatomy_highlight_terms = _to_str(row[state.anatomy_col])  # type: ignore[index]
        if (
            not boundary_highlight_terms
            or not anatomy_highlight_terms
            or not endosteal_highlight_terms
            or not periosteal_highlight_terms
            or not soft_tissue_highlight_terms
            or not pathologic_fracture_highlight_terms
            or not internal_feature_highlight_terms
        ):
            extracted = _extract_rule_highlights(text_value)
            if any(extracted.values()):
                smart_highlight_used = True
            if not boundary_highlight_terms:
                boundary_highlight_terms = extracted["boundary"]
            else:
                boundary_highlight_terms = _merge_highlight_terms_text(boundary_highlight_terms, extracted["boundary"])
            if not anatomy_highlight_terms:
                anatomy_highlight_terms = extracted["anatomy"]
            else:
                anatomy_highlight_terms = _split_highlight_terms_text(anatomy_highlight_terms)[0] if _split_highlight_terms_text(anatomy_highlight_terms) else extracted["anatomy"]
            endosteal_highlight_terms = extracted["endosteal"]
            periosteal_highlight_terms = extracted["periosteal"]
            soft_tissue_highlight_terms = extracted["soft_tissue"]
            pathologic_fracture_highlight_terms = extracted["pathologic_fracture"]
            internal_feature_highlight_terms = extracted["internal_feature"]

        return render_template(
            "review.html",
            state=state,
            current_operator=_to_str(session.get("operator_id")).strip(),
            row_index=row_index,
            source_row_index=source_row_index,
            row_key=key,
            total_rows=len(filtered_indices),
            annotated_count=annotated_count,
            text_value=text_value,
            images=images,
            rating_fields=RATING_FIELDS,
            existing_values=existing_values,
            existing_note=existing_note,
            global_highlight_terms=global_highlight_terms,
            auto_highlight_terms=auto_highlight_terms,
            boundary_highlight_terms=boundary_highlight_terms,
            anatomy_highlight_terms=anatomy_highlight_terms,
            endosteal_highlight_terms=endosteal_highlight_terms,
            periosteal_highlight_terms=periosteal_highlight_terms,
            soft_tissue_highlight_terms=soft_tissue_highlight_terms,
            pathologic_fracture_highlight_terms=pathologic_fracture_highlight_terms,
            internal_feature_highlight_terms=internal_feature_highlight_terms,
            smart_highlight_used=smart_highlight_used,
            has_prev=row_index > 0,
            has_next=row_index + 1 < len(filtered_indices),
            prev_url=url_for("review", row_index=row_index - 1) if row_index > 0 else None,
            next_url=url_for("review", row_index=row_index + 1)
            if row_index + 1 < len(filtered_indices)
            else None,
        )

    @app.post("/save/<int:row_index>")
    def save(row_index: int):
        state = _current_state()
        if state is None:
            flash("请先填写操作者并加载 Excel。")
            return redirect(url_for("index"))
        if not state.is_loaded():
            return redirect(url_for("index"))
        if not state.is_configured():
            return redirect(url_for("config"))
        filtered_indices, source_row_index = _resolve_display_row_index(state, row_index)

        action = _to_str(request.form.get("action")).strip() or "stay"

        # Check if trying to navigate to next case - validate before saving
        if action == "next":
            missing_fields = []
            for field, label, _ in RATING_FIELDS:
                value = _to_str(request.form.get(field)).strip()
                if not value:
                    missing_fields.append(label)

            if missing_fields:
                flash(f"⚠️ 请完成所有评分项后再跳转到下一个案例！缺少：{', '.join(missing_fields)}", "error")
                return redirect(url_for("review", row_index=row_index))

        try:
            payload = _build_annotation_payload(state, source_row_index, request.form)
            _persist_annotation(state, source_row_index, payload)
        except ValueError as e:
            flash(str(e))
            return redirect(url_for("review", row_index=row_index))

        if action == "prev":
            return redirect(url_for("review", row_index=max(0, row_index - 1)))
        if action == "next":
            return redirect(url_for("review", row_index=min(len(filtered_indices) - 1, row_index + 1)))
        if action == "jump":
            try:
                target = int(_to_str(request.form.get("jump_to")).strip())
            except Exception:
                target = row_index
            target = max(0, min(len(filtered_indices) - 1, target))
            return redirect(url_for("review", row_index=target))

        return redirect(url_for("review", row_index=row_index))

    @app.post("/autosave/<int:row_index>")
    def autosave(row_index: int):
        state = _current_state()
        if state is None:
            return jsonify({"ok": False, "error": "请先填写操作者并加载 Excel。"}), 400
        if not state.is_loaded():
            return jsonify({"ok": False, "error": "请先加载 Excel。"}), 400
        if not state.is_configured():
            return jsonify({"ok": False, "error": "请先完成列配置。"}), 400
        _, source_row_index = _resolve_display_row_index(state, row_index)

        try:
            payload = _build_annotation_payload(state, source_row_index, request.form)
            saved = _persist_annotation(state, source_row_index, payload)
        except ValueError as e:
            return jsonify({"ok": False, "error": str(e)}), 400

        return jsonify(
            {
                "ok": True,
                "saved_at": _to_str(saved.get("updated_at")),
                "annotated_count": _annotated_count_for_current_text_col(state),
            }
        )

    @app.get("/image/<int:row_index>/<int:image_index>")
    def image(row_index: int, image_index: int):
        state = _current_state()
        if state is None or not state.is_configured():
            abort(404)
        _, source_row_index = _resolve_display_row_index(state, row_index)

        images = _get_image_paths_for_row(state, source_row_index)
        if image_index < 0 or image_index >= len(images):
            abort(404)
        p = images[image_index].path
        if not p.exists() or not p.is_file():
            abort(404)
        return send_file(p)

    @app.get("/export")
    def export():
        state = _current_state()
        if state is None:
            flash("请先填写操作者并加载 Excel。")
            return redirect(url_for("index"))
        if not state.is_loaded():
            return redirect(url_for("index"))
        if state.annotation_path is None or state.excel_path is None:
            abort(400)

        archive, download_name = _build_export_archive(state)
        return send_file(
            archive,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/zip",
        )

    return app


def main() -> None:
    global BOOTSTRAP_STATE

    parser = argparse.ArgumentParser(description="Excel 图片/文本标注 GUI（Flask）")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=5050, type=int)
    parser.add_argument("--excel", default=str(DEFAULT_EXCEL_PATH), help="可选：启动时直接加载 Excel 路径")
    parser.add_argument("--sheet", default="", help="可选：Sheet 名称")
    parser.add_argument("--image-root", default=str(DEFAULT_IMAGE_ROOT), help="可选：图片根目录（用于路径回退/映射）")
    parser.add_argument("--debug", action="store_true", help="启用 Flask debug；多人同时使用时不建议开启")
    args = parser.parse_args()

    if args.excel:
        try:
            excel_path = _resolve_excel_path(args.excel)
            df = _load_excel(excel_path, args.sheet or None)
            BOOTSTRAP_STATE = AppState(
                excel_path=excel_path,
                sheet_name=args.sheet or None,
                df=df,
            )
            _apply_default_columns(BOOTSTRAP_STATE)
            if args.image_root:
                image_root = _resolve_dir_path(args.image_root)
                if image_root.exists() and image_root.is_dir():
                    BOOTSTRAP_STATE.image_root = image_root
        except Exception:
            pass

    app = create_app()
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
