import hashlib
import io
import os
import re
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

import genanki
import pandas as pd
import streamlit as st
from jinja2 import Environment, StrictUndefined, TemplateError


st.set_page_config(page_title="ZIP to Anki", page_icon="🗂️")

st.title("ZIP to Anki")
st.write(
    "Upload a ZIP containing matching `.html` and `.xlsx` or `.csv` files. "
    "Spreadsheet rows become cards, and HTML templates define their appearance."
)

# --------------------------------------------------
# Upload ZIP and select deck name
# --------------------------------------------------

uploaded_zip = st.file_uploader(
    "Choose a ZIP file",
    type=["zip"],
    help=(
        "The ZIP can contain subfolders. Matching files must have "
        "the same relative path and filename stem."
    ),
)

# Remember which ZIP was selected so we only auto-fill the name
# when a new ZIP is selected.
if "previous_upload_id" not in st.session_state:
    st.session_state.previous_upload_id = None

if "deck_name" not in st.session_state:
    st.session_state.deck_name = ""

if uploaded_zip is not None:
    upload_id = (uploaded_zip.name, uploaded_zip.size)

    if upload_id != st.session_state.previous_upload_id:
        st.session_state.previous_upload_id = upload_id
        st.session_state.deck_name = Path(uploaded_zip.name).stem

deck_name = st.text_input(
    "Anki deck name",
    key="deck_name",
    placeholder="Enter a deck name",
)


# --------------------------------------------------
# Helper functions
# --------------------------------------------------

def stable_id(text: str, namespace: str) -> int:
    """Generate a deterministic numeric ID for an Anki deck or model."""
    digest = hashlib.sha256(
        f"{namespace}:{text}".encode("utf-8")
    ).digest()

    return int.from_bytes(digest[:4], "big") & 0x7FFFFFFF


def read_spreadsheet(filename: str, data: bytes) -> pd.DataFrame:
    """Read an XLSX or CSV spreadsheet into a DataFrame."""
    suffix = Path(filename).suffix.lower()

    if suffix == ".xlsx":
        frame = pd.read_excel(
            io.BytesIO(data),
            dtype=str,
            keep_default_na=False,
        )

    elif suffix == ".csv":
        frame = pd.read_csv(
            io.BytesIO(data),
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )

    else:
        raise ValueError(
            f"Unsupported spreadsheet type: {suffix}"
        )

    # Normalize column headers.
    frame.columns = [
        str(column).strip()
        for column in frame.columns
    ]

    if any(not column for column in frame.columns):
        raise ValueError(
            f"{filename}: every spreadsheet column "
            "must have a non-empty header."
        )

    if len(set(frame.columns)) != len(frame.columns):
        raise ValueError(
            f"{filename}: column headers must be unique."
        )

    return frame


def html_to_anki(html_template: str, row: dict) -> str:
    """
    Render an HTML template using values from a spreadsheet row.

    Example:
        <h2>{{ Front }}</h2>
        <p>{{ Back }}</p>
    """
    env = Environment(
        autoescape=True,
        undefined=StrictUndefined,
    )

    template = env.from_string(html_template)

    return template.render(**row)


def discover_pairs(archive: zipfile.ZipFile):
    """
    Find matching HTML/spreadsheet pairs in a ZIP archive.

    Files must share the same relative path and filename stem.

    Example:
        Animals.html
        Animals.xlsx

    Also supported:
        Grammar/PastTense.html
        Grammar/PastTense.csv

    Each pair must contain exactly one HTML file and exactly
    one spreadsheet file.
    """
    files = {}

    for info in archive.infolist():
        if info.is_dir():
            continue

        path = PurePosixPath(info.filename)

        # Ignore absolute paths and path traversal entries.
        if path.is_absolute() or ".." in path.parts:
            continue

        suffix = path.suffix.lower()

        if suffix not in {".html", ".xlsx", ".csv"}:
            continue

        key = str(path.with_suffix(""))

        files.setdefault(key, {}).setdefault(
            suffix, []
        ).append(info.filename)

    pairs = []

    for key, kinds in sorted(files.items()):
        html_files = kinds.get(".html", [])

        spreadsheet_files = (
            kinds.get(".xlsx", [])
            + kinds.get(".csv", [])
        )

        # Require exactly one HTML file and one spreadsheet.
        if (
            len(html_files) == 1
            and len(spreadsheet_files) == 1
        ):
            pairs.append(
                (
                    key,
                    html_files[0],
                    spreadsheet_files[0],
                )
            )

    return pairs


# --------------------------------------------------
# Generate the Anki package
# --------------------------------------------------

def build_apkg(
    zip_bytes: bytes,
    deck_title: str,
    progress_bar,
    status,
):
    """Create an Anki .apkg package from a ZIP archive."""

    try:
        archive = zipfile.ZipFile(
            io.BytesIO(zip_bytes)
        )

    except zipfile.BadZipFile as exc:
        raise ValueError(
            "The uploaded file is not a valid ZIP archive."
        ) from exc

    with archive:
        pairs = discover_pairs(archive)

        if not pairs:
            raise ValueError(
                "No matching file pairs were found.\n\n"
                "Each pair needs exactly one `.html` file "
                "and exactly one `.xlsx` or `.csv` file "
                "with the same relative path and filename stem."
            )

        # Create the deck.
        deck = genanki.Deck(
            stable_id(deck_title, "anki-deck"),
            deck_title,
        )

        # All cards use the same Anki note model.
        model = genanki.Model(
            stable_id(deck_title, "anki-model"),
            f"{deck_title} Note",
            fields=[
                {"name": "Content"},
            ],
            templates=[
                {
                    "name": "Card",
                    "qfmt": "{{Content}}",
                    "afmt": "{{FrontSide}}",
                }
            ],
            css="""
.card {
    font-family: Arial, sans-serif;
    font-size: 20px;
    text-align: left;
    color: #222;
    background-color: #fff;
    padding: 1em;
}

img {
    max-width: 100%;
    height: auto;
}
""",
        )

        errors = []
        total_cards = 0
        completed_pairs = 0
        total_pairs = len(pairs)

        # Process each matched pair.
        for key, html_name, sheet_name in pairs:
            status.text(
                f"Processing {key} "
                f"({completed_pairs + 1} of {total_pairs})..."
            )

            try:
                # Read and decode the HTML template.
                html_template = archive.read(
                    html_name
                ).decode("utf-8-sig")

                # Read the corresponding spreadsheet.
                frame = read_spreadsheet(
                    sheet_name,
                    archive.read(sheet_name),
                )

                if frame.empty:
                    errors.append(
                        f"{key}: spreadsheet contains no data rows."
                    )

                    completed_pairs += 1
                    progress_bar.progress(
                        completed_pairs / total_pairs
                    )
                    continue

                # Create one Anki note per spreadsheet row.
                rows = frame.to_dict(orient="records")

                for row_number, row_data in enumerate(
                    rows,
                    start=2,
                ):
                    row = {
                        str(column): (
                            "" if value is None
                            else str(value)
                        )
                        for column, value in row_data.items()
                    }

                    try:
                        rendered_html = html_to_anki(
                            html_template,
                            row,
                        )

                    except TemplateError as exc:
                        raise ValueError(
                            "HTML template error on "
                            f"spreadsheet row {row_number}: {exc}"
                        ) from exc

                    # Use a deterministic GUID for each note.
                    guid_seed = (
                        f"{key}:{row_number}:{row!r}"
                    )

                    note = genanki.Note(
                        model=model,
                        fields=[rendered_html],
                        guid=genanki.guid_for(guid_seed),
                    )

                    deck.add_note(note)
                    total_cards += 1

            except Exception as exc:
                errors.append(f"{key}: {exc}")

            completed_pairs += 1

            progress_bar.progress(
                completed_pairs / total_pairs
            )

        # Don't produce an empty deck if all pairs failed.
        if total_cards == 0:
            detail = "\n".join(
                f"- {error}"
                for error in errors
            )

            raise ValueError(
                "No cards could be generated from the matching pairs."
                + (
                    f"\n\nDetails:\n{detail}"
                    if detail
                    else ""
                )
            )

        # Package the deck into an .apkg file.
        status.text(
            f"Packaging {total_cards} card(s) into an Anki deck..."
        )

        package = genanki.Package(deck)

        with tempfile.NamedTemporaryFile(
            suffix=".apkg",
            delete=False,
        ) as temp_file:
            temp_path = temp_file.name

        try:
            package.write_to_file(temp_path)

            with open(temp_path, "rb") as generated:
                apkg_bytes = generated.read()

        finally:
            try:
                os.remove(temp_path)
            except OSError:
                pass

        return (
            apkg_bytes,
            total_cards,
            pairs,
            errors,
        )


# --------------------------------------------------
# Generate button
# --------------------------------------------------

with st.form("generate_deck_form"):
    submitted = st.form_submit_button(
        "Generate Anki deck",
        type="primary",
        disabled=(
            uploaded_zip is None
            or not deck_name.strip()
        ),
    )


if submitted:
    if uploaded_zip is None:
        st.error("Please upload a ZIP file first.")

    elif not deck_name.strip():
        st.error("Please enter a deck name.")

    else:
        progress = st.progress(
            0,
            text="Starting...",
        )

        status = st.empty()

        try:
            (
                apkg,
                card_count,
                pairs,
                warnings,
            ) = build_apkg(
                uploaded_zip.getvalue(),
                deck_name.strip(),
                progress,
                status,
            )

            progress.progress(1.0)

            status.success(
                f"Deck ready! Generated {card_count} card(s) "
                f"from {len(pairs)} matched pair(s)."
            )

            # Keep the generated package available for download.
            st.session_state.generated_apkg = apkg

            # Sanitize the filename for the download.
            safe_filename = re.sub(
                r'[\\/:*?"<>|]+',
                "_",
                deck_name.strip(),
            ).strip(" .")

            st.session_state.generated_filename = (
                safe_filename or "anki_deck"
            )

            st.session_state.generated_card_count = card_count
            st.session_state.generated_warnings = warnings

        except Exception as exc:
            st.session_state.pop(
                "generated_apkg",
                None,
            )

            st.error(str(exc))


# --------------------------------------------------
# Download the generated deck
# --------------------------------------------------

if st.session_state.get("generated_apkg") is not None:
    st.download_button(
        label="Download Anki deck (.apkg)",
        data=st.session_state.generated_apkg,
        file_name=(
            f"{st.session_state.generated_filename}.apkg"
        ),
        mime="application/octet-stream",
        type="primary",
    )

    st.caption(
        f"{st.session_state.generated_card_count} card(s) generated."
    )

    for warning in st.session_state.get(
        "generated_warnings",
        [],
    ):
        st.warning(warning)


# --------------------------------------------------
# Instructions
# --------------------------------------------------

with st.expander("How the ZIP should be organized"):
    st.markdown(
        """
        Example:

        ```
        my-materials.zip
        ├── Animals.html
        ├── Animals.xlsx
        ├── Grammar/
        │   ├── PastTense.html
        │   └── PastTense.csv
        ```

        `Animals.html` pairs with `Animals.xlsx`.

        `Grammar/PastTense.html` pairs with
        `Grammar/PastTense.csv`.

        In each HTML file, use spreadsheet column names as
        placeholders, for example:

        ```html
        <h2>{{ Front }}</h2>
        <p>{{ Back }}</p>
        ```

        Each spreadsheet row creates one Anki note/card.

        Use UTF-8 encoding for HTML and CSV files.
        """
    )