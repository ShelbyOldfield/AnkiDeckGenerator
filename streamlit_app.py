
from __future__ import annotations

import copy
import hashlib
import io
import json
import re
import tempfile
import uuid
import zipfile
from pathlib import Path, PurePosixPath

import genanki
import pandas as pd
import streamlit as st

APP_VERSION = "0.2.0"
SUPPORTED = {".xlsx", ".csv"}


def new_id() -> str:
    return str(uuid.uuid4())


def anki_id(stable_key: str) -> int:
    value = int.from_bytes(
        hashlib.sha256(stable_key.encode()).digest()[:4], "big"
    ) & 0x7FFFFFFF
    return value or 1


def safe_zip_path(name: str) -> str:
    p = PurePosixPath(name.replace("\\", "/"))
    if p.is_absolute() or ".." in p.parts:
        raise ValueError(f"Unsafe path in ZIP: {name}")
    return str(p)


def read_sheet(filename: str, content: bytes) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()

    if suffix == ".xlsx":
        df = pd.read_excel(io.BytesIO(content), engine="openpyxl")
    elif suffix == ".csv":
        df = pd.read_csv(io.BytesIO(content), encoding="utf-8-sig")
    else:
        raise ValueError(f"Unsupported file type: {filename}")

    df.columns = [str(c).strip() for c in df.columns]

    if any(not c for c in df.columns):
        raise ValueError(f"{filename} has a blank column header.")

    if len(set(df.columns)) != len(df.columns):
        raise ValueError(f"{filename} has duplicate column headers.")

    return df


def make_note_type(filename: str, fields: list[str]) -> dict:
    ntid = new_id()
    front = fields[0] if fields else "Front"
    back = fields[1] if len(fields) > 1 else front
    stem = Path(filename).stem or "Notes"

    return {
        "id": ntid,
        "name": f"{stem} Notes",
        "anki_model_id": anki_id("model:" + ntid),
        "model_type": "standard",
        "fields": fields or ["Front", "Back"],
        "css": (
            ".card { font-family: Arial; font-size: 22px; "
            "text-align: center; }"
        ),
        "templates": [
            {
                "id": new_id(),
                "name": "Card 1",
                "front": "{{" + front + "}}",
                "back": (
                    '{{FrontSide}}<hr id="answer">\\n{{'
                    + back
                    + "}}"
                ),
                "required_fields": [],
            }
        ],
    }


def read_zip(content: bytes):
    files, warnings = {}, []

    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        for info in zf.infolist():
            if not info.is_dir():
                name = safe_zip_path(info.filename)
                files[name] = zf.read(info.filename)

    config = None

    if "deck.json" in files:
        try:
            config = json.loads(
                files["deck.json"].decode("utf-8-sig")
            )
        except Exception as exc:
            warnings.append(f"Could not parse deck.json: {exc}")

    for filename, raw in files.items():
        path = PurePosixPath(filename)

        if path.suffix.lower() == ".json" and filename != "deck.json":
            try:
                obj = json.loads(raw.decode("utf-8-sig"))

                if isinstance(obj, dict) and (
                    "fields" in obj or "templates" in obj
                ):
                    config = config or {
                        "deck": {},
                        "datasets": [],
                        "note_types": [],
                    }
                    config.setdefault("note_types", [])

                    if not any(
                        x.get("id") == obj.get("id")
                        for x in config["note_types"]
                    ):
                        config["note_types"].append(obj)

            except Exception as exc:
                warnings.append(
                    f"Could not parse {filename}: {exc}"
                )

    sheets = {
        n: b
        for n, b in files.items()
        if Path(n).suffix.lower() in SUPPORTED
    }

    return config, sheets, warnings


def normalize(config, sheets, fallback_name):
    warnings = []
    frames = {}

    for filename, raw in sheets.items():
        try:
            frames[filename] = read_sheet(filename, raw)
        except Exception as exc:
            warnings.append(f"Skipped {filename}: {exc}")

    if config is None:
        datasets, note_types, schema_map = [], [], {}

        for filename, df in frames.items():
            schema = tuple(df.columns)

            if schema not in schema_map:
                nt = make_note_type(filename, list(schema))
                note_types.append(nt)
                schema_map[schema] = nt["id"]

            datasets.append({
                "id": new_id(),
                "name": Path(filename).stem,
                "file": filename,
                "included": True,
                "note_type_id": schema_map[schema],
                "guid_column": "",
                "tags_column": "",
                "default_tags": [],
            })

        deck_id = new_id()

        config = {
            "format_version": 1,
            "app_version": APP_VERSION,
            "deck": {
                "id": deck_id,
                "name": fallback_name,
                "anki_deck_id": anki_id("deck:" + deck_id),
                "description": "",
            },
            "datasets": datasets,
            "note_types": note_types,
        }

    else:
        config = copy.deepcopy(config)
        config.setdefault("deck", {})
        config.setdefault("datasets", [])
        config.setdefault("note_types", [])

        deck = config["deck"]
        deck.setdefault("id", new_id())
        deck.setdefault("name", fallback_name)
        deck.setdefault(
            "anki_deck_id",
            anki_id("deck:" + deck["id"]),
        )
        deck.setdefault("description", "")

        configured_files = {
            d.get("file") for d in config["datasets"]
        }

        for filename, df in frames.items():
            if filename not in configured_files:
                nt = next(
                    (
                        x for x in config["note_types"]
                        if x.get("fields") == list(df.columns)
                    ),
                    None,
                )

                if nt is None:
                    nt = make_note_type(filename, list(df.columns))
                    config["note_types"].append(nt)

                config["datasets"].append({
                    "id": new_id(),
                    "name": Path(filename).stem,
                    "file": filename,
                    "included": True,
                    "note_type_id": nt["id"],
                    "guid_column": "",
                    "tags_column": "",
                    "default_tags": [],
                })

        for ds in config["datasets"]:
            ds.setdefault("id", new_id())
            ds.setdefault(
                "name",
                Path(ds.get("file", "Dataset")).stem,
            )
            ds.setdefault("included", True)
            ds.setdefault("note_type_id", None)
            ds.setdefault("guid_column", "")
            ds.setdefault("tags_column", "")
            ds.setdefault("default_tags", [])

            if ds.get("file") not in frames:
                warnings.append(
                    f"Configured dataset '{ds['name']}' has no "
                    f"uploaded file '{ds.get('file')}'."
                )

            elif not ds.get("note_type_id"):
                df = frames[ds["file"]]

                nt = next(
                    (
                        x for x in config["note_types"]
                        if x.get("fields") == list(df.columns)
                    ),
                    None,
                )

                if nt is None:
                    nt = make_note_type(
                        ds["file"], list(df.columns)
                    )
                    config["note_types"].append(nt)

                ds["note_type_id"] = nt["id"]

    for nt in config["note_types"]:
        nt.setdefault("id", new_id())
        nt.setdefault("name", "Notes")
        nt.setdefault(
            "anki_model_id",
            anki_id("model:" + nt["id"]),
        )
        nt.setdefault("model_type", "standard")
        nt.setdefault("fields", ["Front", "Back"])
        nt.setdefault("css", "")
        nt.setdefault("templates", [])

        if nt["model_type"] == "standard" and not nt["templates"]:
            nt["templates"] = make_note_type(
                nt["name"], nt["fields"]
            )["templates"]

        for template in nt["templates"]:
            template.setdefault("id", new_id())
            template.setdefault("name", "Card")
            template.setdefault("front", "{{Front}}")
            template.setdefault(
                "back", "{{FrontSide}}<hr>{{Back}}"
            )
            template.setdefault("required_fields", [])

    return config, frames, warnings


def build_apkg(config, frames):
    errors = []
    deck_cfg = config["deck"]

    deck = genanki.Deck(
        int(deck_cfg["anki_deck_id"]),
        str(deck_cfg["name"]),
        description=str(deck_cfg.get("description", "")),
    )

    models = {}
    seen_model_ids = set()

    for nt in config["note_types"]:
        model_id = int(nt["anki_model_id"])

        if model_id in seen_model_ids:
            errors.append(
                f"Duplicate Anki model ID {model_id}; "
                f"note type '{nt['name']}' skipped."
            )
            continue

        seen_model_ids.add(model_id)

        fields = [{"name": str(x)} for x in nt["fields"]]
        templates = []

        if nt.get("model_type", "standard") == "standard":
            for t in nt.get("templates", []):
                tpl = {
                    "name": str(t.get("name", "Card")),
                    "qfmt": str(t.get("front", "")),
                    "afmt": str(t.get("back", "")),
                }

                if t.get("required_fields"):
                    tpl["req"] = [
                        int(x) for x in t["required_fields"]
                    ]

                templates.append(tpl)

            if not templates:
                errors.append(
                    f"'{nt['name']}' has no card templates."
                )
                continue

        model = genanki.Model(
            model_id,
            str(nt["name"]),
            fields=fields,
            templates=templates,
            css=str(nt.get("css", "")),
            model_type=(
                genanki.Model.CLOZE
                if nt.get("model_type") == "cloze"
                else 0
            ),
        )

        models[str(nt["id"])] = model

    count = 0

    for ds in config["datasets"]:
        if not ds.get("included", True):
            continue

        filename = ds.get("file")

        if filename not in frames:
            errors.append(
                f"Dataset '{ds.get('name')}' is missing its spreadsheet."
            )
            continue

        model = models.get(str(ds.get("note_type_id")))

        if model is None:
            errors.append(
                f"Dataset '{ds.get('name')}' references an invalid note type."
            )
            continue

        df = frames[filename]
        field_names = [f["name"] for f in model.fields]
        missing = [f for f in field_names if f not in df.columns]

        if missing:
            errors.append(
                f"{filename} is missing note fields/columns: "
                f"{', '.join(missing)}."
            )
            continue

        guid_col = ds.get("guid_column") or ""
        tags_col = ds.get("tags_column") or ""

        for col, label in (
            (guid_col, "GUID"),
            (tags_col, "tags"),
        ):
            if col and col not in df.columns:
                errors.append(
                    f"{filename}: configured {label} column "
                    f"'{col}' does not exist."
                )
                break

        else:
            for row_number, (_, row) in enumerate(df.iterrows()):
                values = [
                    "" if pd.isna(row[f]) else str(row[f])
                    for f in field_names
                ]

                if not any(v.strip() for v in values):
                    continue

                guid = None

                if guid_col:
                    value = (
                        ""
                        if pd.isna(row[guid_col])
                        else str(row[guid_col]).strip()
                    )
                    guid = value or None

                if not guid:
                    # Fallback is stable only while the dataset ID
                    # and row order stay stable.
                    guid = hashlib.sha256(
                        f"{ds['id']}:{row_number}".encode()
                    ).hexdigest()[:32]

                tags = list(ds.get("default_tags", []))

                if tags_col and not pd.isna(row[tags_col]):
                    tags += [
                        x
                        for x in re.split(
                            r"[,;\\s]+",
                            str(row[tags_col]).strip(),
                        )
                        if x
                    ]

                try:
                    deck.add_note(
                        genanki.Note(
                            model=model,
                            fields=values,
                            guid=guid,
                            tags=sorted(set(tags)),
                        )
                    )
                    count += 1

                except Exception as exc:
                    errors.append(
                        f"{filename}, row {row_number + 2}: {exc}"
                    )

    with tempfile.NamedTemporaryFile(suffix=".apkg") as tmp:
        genanki.Package(deck).write_to_file(tmp.name)
        tmp.seek(0)
        return tmp.read(), errors, count, len(models)


def make_project_zip(config, source_files):
    output = io.BytesIO()

    with zipfile.ZipFile(
        output, "w", zipfile.ZIP_DEFLATED
    ) as zf:
        zf.writestr(
            "deck.json",
            json.dumps(config, ensure_ascii=False, indent=2),
        )

        for nt in config.get("note_types", []):
            filename = re.sub(
                r"[^A-Za-z0-9_.-]+",
                "-",
                nt.get("name", "note-type"),
            ).strip("-") or "note-type"

            zf.writestr(
                f"note-types/{filename}.json",
                json.dumps(nt, ensure_ascii=False, indent=2),
            )

        for filename, content in source_files.items():
            target = safe_zip_path(filename)

            if (
                target == "deck.json"
                or target.startswith("note-types/")
            ):
                target = "datasets/" + target

            zf.writestr(target, content)

    return output.getvalue()


st.set_page_config(
    page_title="Anki Deck Generator",
    page_icon="🗂️",
    layout="wide",
)

st.title("Anki Deck Generator")
st.caption(
    "Generate an Anki package from spreadsheets, "
    "with reusable project configuration."
)

with st.sidebar:
    st.header("Input")

    uploaded_zip = st.file_uploader(
        "Project ZIP (optional)", type=["zip"]
    )

    uploaded_sheets = st.file_uploader(
        "Or upload spreadsheets",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
    )

if uploaded_zip is None and not uploaded_sheets:
    st.info(
        "Upload a ZIP containing spreadsheets/configuration, "
        "or upload one or more spreadsheets."
    )
    st.stop()

source_files = {}
loaded_config = None
warnings = []

if uploaded_zip is not None:
    try:
        loaded_config, source_files, warnings = read_zip(
            uploaded_zip.getvalue()
        )

        for file in uploaded_sheets or []:
            source_files[file.name] = file.getvalue()

    except Exception as exc:
        st.error(f"Could not read ZIP: {exc}")
        st.stop()

else:
    source_files = {
        f.name: f.getvalue()
        for f in uploaded_sheets
    }

source_files = {
    n: b
    for n, b in source_files.items()
    if Path(n).suffix.lower() in SUPPORTED
}

if not source_files:
    st.error("No .xlsx or .csv files were found.")
    st.stop()

fallback_name = (
    Path(uploaded_zip.name).stem
    if uploaded_zip
    else Path(next(iter(source_files))).stem
) or "My Anki Deck"

# Keep the normalized project config in session state. Streamlit reruns
# this script after every widget interaction; regenerating a fresh config
# on every rerun would create new internal IDs and reset edited Anki IDs.
fingerprint_hash = hashlib.sha256()

if uploaded_zip is not None:
    fingerprint_hash.update(uploaded_zip.getvalue())

for filename, content in sorted(source_files.items()):
    fingerprint_hash.update(filename.encode("utf-8"))
    fingerprint_hash.update(content)

input_fingerprint = fingerprint_hash.hexdigest()

if "project_config_override" in st.session_state:
    st.session_state["active_config"] = (
        st.session_state.pop("project_config_override")
    )
    st.session_state["active_config_fingerprint"] = input_fingerprint

if (
    st.session_state.get("active_config_fingerprint")
    != input_fingerprint
    or "active_config" not in st.session_state
):
    config, frames, normalize_warnings = normalize(
        loaded_config, source_files, fallback_name
    )

    st.session_state["active_config"] = config
    st.session_state["active_config_fingerprint"] = (
        input_fingerprint
    )

else:
    config = st.session_state["active_config"]

    # The uploaded bytes remain the source of truth for dataset contents.
    frames = {}
    normalize_warnings = []

    for filename, content in source_files.items():
        try:
            frames[filename] = read_sheet(filename, content)
        except Exception as exc:
            normalize_warnings.append(
                f"Skipped {filename}: {exc}"
            )

warnings += normalize_warnings

for warning in warnings:
    st.warning(warning)


st.header("Deck configuration")

deck_cfg = config["deck"]
c1, c2 = st.columns([2, 1])

with c1:
    deck_cfg["name"] = st.text_input(
        "Deck name", str(deck_cfg["name"])
    )

with c2:
    deck_cfg["anki_deck_id"] = int(
        st.number_input(
            "Anki deck ID",
            1,
            2147483647,
            int(deck_cfg["anki_deck_id"]),
            step=1,
        )
    )

deck_cfg["description"] = st.text_area(
    "Description (optional)",
    str(deck_cfg.get("description", "")),
    height=70,
)


st.header("Datasets")

st.caption(
    "A dataset is one spreadsheet. Fields used by its selected "
    "note type must match spreadsheet column headers."
)

note_types = config["note_types"]

for ds in config["datasets"]:
    with st.expander(
        f"{ds['name']} — {ds['file']}",
        expanded=False,
    ):
        ds["name"] = st.text_input(
            "Dataset name",
            str(ds["name"]),
            key=f"name_{ds['id']}",
        )

        ds["included"] = st.checkbox(
            "Include dataset",
            bool(ds.get("included", True)),
            key=f"include_{ds['id']}",
        )

        if note_types:
            ids = [str(nt["id"]) for nt in note_types]

            labels = {
                str(nt["id"]): (
                    f"{nt['name']} "
                    f"({', '.join(nt.get('fields', []))})"
                )
                for nt in note_types
            }

            current = str(ds.get("note_type_id") or ids[0])

            if current not in ids:
                current = ids[0]

            ds["note_type_id"] = st.selectbox(
                "Note type",
                ids,
                index=ids.index(current),
                format_func=lambda x: labels[x],
                key=f"ntsel_{ds['id']}",
            )

        df = frames.get(ds["file"])

        if df is not None:
            st.caption(
                f"{len(df)} rows; columns: {', '.join(df.columns)}"
            )

            columns = ["(none)"] + list(df.columns)

            g = ds.get("guid_column") or "(none)"
            t = ds.get("tags_column") or "(none)"

            x, y = st.columns(2)

            with x:
                choice = st.selectbox(
                    "Stable note GUID column",
                    columns,
                    index=columns.index(g) if g in columns else 0,
                    key=f"guid_{ds['id']}",
                    help=(
                        "Recommended for stable updates when rows "
                        "are sorted or inserted."
                    ),
                )

            with y:
                tag_choice = st.selectbox(
                    "Tags column",
                    columns,
                    index=columns.index(t) if t in columns else 0,
                    key=f"tags_{ds['id']}",
                )

            ds["guid_column"] = (
                "" if choice == "(none)" else choice
            )

            ds["tags_column"] = (
                "" if tag_choice == "(none)" else tag_choice
            )

            ds["default_tags"] = [
                x.strip()
                for x in st.text_input(
                    "Default tags (comma-separated)",
                    ", ".join(ds.get("default_tags", [])),
                    key=f"deftags_{ds['id']}",
                ).split(",")
                if x.strip()
            ]

        else:
            st.warning("Spreadsheet not found in this upload.")


st.header("Note types and card templates")

st.caption(
    "HTML snippets use Anki's native template syntax. "
    "For standard notes, each card template creates a card."
)

for nt in note_types:
    with st.expander(
        f"{nt['name']} ({nt.get('model_type', 'standard')})",
        expanded=False,
    ):
        nt["name"] = st.text_input(
            "Note type name",
            str(nt["name"]),
            key=f"ntname_{nt['id']}",
        )

        nt["anki_model_id"] = int(
            st.number_input(
                "Anki model ID",
                1,
                2147483647,
                int(nt["anki_model_id"]),
                step=1,
                key=f"modelid_{nt['id']}",
            )
        )

        nt["model_type"] = st.selectbox(
            "Note type kind",
            ["standard", "cloze"],
            index=1 if nt.get("model_type") == "cloze" else 0,
            key=f"kind_{nt['id']}",
        )

        nt["fields"] = [
            x.strip()
            for x in st.text_input(
                "Fields in order (comma-separated)",
                ", ".join(nt.get("fields", [])),
                key=f"fields_{nt['id']}",
            ).split(",")
            if x.strip()
        ]

        nt["css"] = st.text_area(
            "CSS",
            str(nt.get("css", "")),
            height=90,
            key=f"css_{nt['id']}",
        )

        if nt["model_type"] == "standard":
            for template in nt["templates"]:
                st.markdown(
                    f"**Card template: {template['name']}**"
                )

                template["name"] = st.text_input(
                    "Template name",
                    str(template["name"]),
                    key=f"tname_{template['id']}",
                )

                left, right = st.columns(2)

                with left:
                    template["front"] = st.text_area(
                        "Front HTML",
                        str(template["front"]),
                        height=120,
                        key=f"front_{template['id']}",
                    )

                with right:
                    template["back"] = st.text_area(
                        "Back HTML",
                        str(template["back"]),
                        height=120,
                        key=f"back_{template['id']}",
                    )

            if st.button(
                "Add card template",
                key=f"addtpl_{nt['id']}",
            ):
                fields = nt["fields"] or ["Front", "Back"]

                nt["templates"].append({
                    "id": new_id(),
                    "name": f"Card {len(nt['templates']) + 1}",
                    "front": "{{" + fields[0] + "}}",
                    "back": (
                        '{{FrontSide}}<hr id="answer">\\n{{'
                        + (
                            fields[1]
                            if len(fields) > 1
                            else fields[0]
                        )
                        + "}}"
                    ),
                    "required_fields": [],
                })

                st.rerun()

        else:
            st.info(
                "Cloze notes use markers such as {{c1::answer}} "
                "in a field; Anki generates deletion cards automatically."
            )

        if st.button(
            "Remove note type",
            key=f"rmnt_{nt['id']}",
        ):
            config["note_types"] = [
                x for x in config["note_types"]
                if x["id"] != nt["id"]
            ]

            for ds in config["datasets"]:
                if ds.get("note_type_id") == nt["id"]:
                    ds["note_type_id"] = None

            st.rerun()


if st.button("Add note type"):
    config["note_types"].append(
        make_note_type("NewNotes", ["Front", "Back"])
    )
    st.rerun()


with st.expander("Advanced: edit the complete configuration JSON"):
    json_edit = st.text_area(
        "deck.json",
        json.dumps(config, ensure_ascii=False, indent=2),
        height=320,
    )

    if st.button("Apply JSON"):
        try:
            parsed = json.loads(json_edit)

            if not isinstance(parsed, dict):
                raise ValueError(
                    "Top-level JSON must be an object."
                )

            st.session_state["project_config_override"] = parsed
            st.rerun()

        except Exception as exc:
            st.error(f"Invalid JSON: {exc}")


st.header("Generate")

st.warning(
    "For robust regeneration, add a stable GUID column to each "
    "spreadsheet and select it above. Without one, the app uses "
    "dataset ID + row position, which is not robust to sorting "
    "or inserting rows."
)

if st.button(
    "Build Anki package",
    type="primary",
    use_container_width=True,
):
    errors = []
    ids = {}

    for nt in config["note_types"]:
        mid = int(nt["anki_model_id"])

        if mid in ids:
            errors.append(
                f"Note types '{ids[mid]}' and '{nt['name']}' "
                f"share model ID {mid}."
            )

        ids[mid] = nt["name"]

        if len(nt["fields"]) != len(set(nt["fields"])):
            errors.append(
                f"Note type '{nt['name']}' has duplicate field names."
            )

    if errors:
        for err in errors:
            st.error(err)

    else:
        try:
            with st.spinner("Building package..."):
                (
                    package_bytes,
                    build_errors,
                    note_count,
                    model_count,
                ) = build_apkg(config, frames)

            for err in build_errors:
                st.warning(err)

            if note_count == 0:
                st.error(
                    "No notes were generated. "
                    "Check the dataset field mappings."
                )

            else:
                st.success(
                    f"Generated {note_count} notes using "
                    f"{model_count} note type(s)."
                )

                safe_name = (
                    re.sub(
                        r"[^A-Za-z0-9_.-]+",
                        "_",
                        deck_cfg["name"],
                    ).strip("_")
                    or "anki_deck"
                )

                st.download_button(
                    "Download Anki package (.apkg)",
                    package_bytes,
                    f"{safe_name}.apkg",
                    mime="application/octet-stream",
                    use_container_width=True,
                )

                st.download_button(
                    "Download reusable project ZIP",
                    make_project_zip(config, source_files),
                    f"{safe_name}_project.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

        except Exception as exc:
            st.error(f"Could not build package: {exc}")


with st.expander("Current limitations"):
    st.markdown("""
    - Supports `.xlsx` and `.csv`, multiple datasets, multiple note types,
      multiple standard card templates, basic cloze notes, CSS, and tags.
    - Media packaging and the full range of Anki deck options are not
      implemented yet.
    - `genanki` may not expose every newer Anki feature.
    - Test repeated imports in a disposable Anki profile before relying
      on update behavior to preserve scheduling.
    """)
