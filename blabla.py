import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
import sys
import os
import shutil
import subprocess
import importlib.util
import logging
from pathlib import Path
from string import Template
import warnings
import re
from datetime import datetime
from openpyxl import load_workbook
from matplotlib.patches import Patch

warnings.filterwarnings(
    "ignore",
    message=".*extension is not supported and will be removed.*",
    category=UserWarning,
)
warnings.filterwarnings(
    "ignore",
    message=".*wmf image format is not supported so the image is being dropped.*",
    category=UserWarning,
)


logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("stps_report")

WORKBOOK_FILE = "stps_tolk.xlsx"
SHEET_NAME = "Hovedtabell"
REPORT_TEMPLATE_PATH = Path("data") / "report_template.html"
LOGO_FILENAME = "st_logo_hvit_bgr.png"
LOGO_FILENAME_PATH = Path("data") / LOGO_FILENAME
LOGO_ENV_VAR = "STPS_LOGO_PATH"
SOURCE_CANDIDATES = [
    "stps_2026_RTM.xlsm",
    "STPS_2026_RTM.xlsm",
    "stps_2027_RTM.xlsm",
    "STPS_2027_RTM.xlsm",
    "stps_2027.xlsm",
    "STPS_2027.xlsm",
    "STPS 2027.xlsm",
    "stps_tolk.xlsx.xlsm",
]
WEEKLY_CANDIDATES = [
    "Sist_uke",
    "Alle_Siste_Ukes_Poeng",
    "Ukevinnere",
    "UkasVinner",
    "ukevinneretabell",
]

file = WORKBOOK_FILE
sheet_name = SHEET_NAME
refresh_on_start = "--refresh" in sys.argv


def resolve_workbook_path(preferred_file: str) -> str:
    preferred_path = Path(preferred_file)
    if preferred_path.exists():
        return str(preferred_path)

    xlsm_files = sorted(
        p for p in Path(".").glob("*.xlsm") if not p.name.startswith("~$")
    )
    if not xlsm_files:
        raise FileNotFoundError("Fant ingen .xlsm-filer i arbeidsmappen.")

    # Fuzzy match filename regardless of case, spaces, underscores, etc.
    def normalize_name(name: str) -> str:
        return re.sub(r"[^a-z0-9]", "", name.lower())

    preferred_norm = normalize_name(preferred_path.stem)
    exact_like = [p for p in xlsm_files if normalize_name(p.stem) == preferred_norm]
    if exact_like:
        return str(exact_like[0])

    # Prefer files that look like STPS annual workbooks.
    stps_candidates = [p for p in xlsm_files if "stps" in p.name.lower()]
    selected = stps_candidates[0] if stps_candidates else xlsm_files[0]

    logger.warning(f"Advarsel: '{preferred_file}' ble ikke funnet. Bruker '{selected.name}' i stedet.")
    return str(selected)


def update_stps_tolk_on_start(target_file: str, target_sheet: str) -> pd.DataFrame | None:
    source_file = None
    for candidate in SOURCE_CANDIDATES:
        if Path(candidate).exists():
            source_file = candidate
            break

    if source_file is None:
        logger.info("Info: Fant ingen 2027-kilde for oppdatering. Bruker eksisterende stps_tolk.xlsx.")
        return None

    try:
        source_df = load_named_range_to_df(source_file, "Sammendrag")

        # Write plain values (no formatting) so stps_tolk.xlsx stays a simple data store.
        write_df = source_df.copy().where(pd.notna(source_df), "")

        with pd.ExcelWriter(target_file, engine="openpyxl", mode="w") as writer:
            write_df.to_excel(writer, sheet_name=target_sheet, index=False)
        logger.info(f"Oppdaterte '{target_file}' fra '{source_file}' (område: Sammendrag).")
        return source_df
    except PermissionError:
        logger.warning(f"Advarsel: '{target_file}' er låst og kunne ikke skrives.")
        logger.info("Bruker oppdaterte data i minnet for denne kjøringen.")
        return source_df
    except Exception as exc:
        logger.warning(f"Advarsel: Kunne ikke oppdatere '{target_file}' fra '{source_file}': {exc}")
        logger.info("Bruker eksisterende stps_tolk.xlsx videre i kjøringen.")
        return None


def find_source_workbook() -> str | None:
    for candidate in SOURCE_CANDIDATES:
        if Path(candidate).exists():
            return candidate

    return None


def refresh_linked_workbooks_via_excel(source_file: str, target_file: str) -> bool:
    """Refresh linked Excel data by opening source first, then target, through Excel COM."""
    if os.name != "nt":
        return False

    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except Exception as exc:
        logger.info(f"Info: Excel COM ikke tilgjengelig for link-oppdatering: {exc}")
        return False

    source_path = str(Path(source_file).resolve())
    target_path = str(Path(target_file).resolve())

    excel = None
    source_wb = None
    target_wb = None

    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        # Open source first so external references in target can resolve against an open workbook.
        source_wb = excel.Workbooks.Open(source_path, UpdateLinks=3, ReadOnly=True)
        source_wb.RefreshAll()
        excel.CalculateFullRebuild()

        target_wb = excel.Workbooks.Open(target_path, UpdateLinks=3, ReadOnly=False)
        target_wb.RefreshAll()
        excel.CalculateFullRebuild()
        target_wb.Save()

        logger.info("Info: Oppdaterte linker i stps_tolk.xlsx via Excel.")
        return True
    except Exception as exc:
        logger.warning(f"Advarsel: Klarte ikke å oppdatere linker via Excel: {exc}")
        return False
    finally:
        if target_wb is not None:
            try:
                target_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if source_wb is not None:
            try:
                source_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()  # type: ignore[name-defined]
        except Exception:
            pass


def build_weekly_filename(
    base_name: str,
    extension: str,
    report_week: int,
    report_year: int,
) -> str:
    # Avoid accidental double separators when callers pass names like "report_".
    clean_base = re.sub(r"[_\-\s]+$", "", base_name)
    if not clean_base:
        clean_base = "report"
    return f"{clean_base}_uke_{report_week}_{report_year}{extension}"


def create_weekly_workbook_copy(
    file_path: str,
    report_week: int,
    report_year: int,
) -> str | None:
    source_path = Path(file_path)
    if not source_path.exists():
        return None

    copy_name = build_weekly_filename(
        source_path.stem,
        source_path.suffix,
        report_week,
        report_year,
    )
    copy_path = source_path.with_name(copy_name)
    shutil.copy2(source_path, copy_path)
    return str(copy_path)


def render_report_html(
    template_path: Path,
    logo_src: str,
    report_week: int,
    report_pdf_filename: str,
    pdf_ready_hint: bool,
    report_updated_at: str,
    summary_section_html: str,
    weekly_section_html: str,
) -> str:
    template = Template(template_path.read_text(encoding="utf-8"))
    return template.substitute(
        logo_src=logo_src,
        report_week=report_week,
        report_pdf_filename=report_pdf_filename,
        pdf_status_text=(
            "PDF-eksport klar" if pdf_ready_hint else "PDF krever weasyprint eller wkhtmltopdf"
        ),
        report_updated_at=report_updated_at,
        summary_section_html=summary_section_html,
        weekly_section_html=weekly_section_html,
    )


def resolve_logo_path() -> Path:
    """Resolve logo path from env var and common project locations."""
    candidates: list[Path] = []

    env_logo_path = os.environ.get(LOGO_ENV_VAR)
    if env_logo_path:
        candidates.append(Path(env_logo_path).expanduser())

    base_dir = Path(__file__).resolve().parent
    cwd = Path.cwd()
    candidates.extend(
        [
            base_dir / LOGO_FILENAME,
            base_dir / "data" / LOGO_FILENAME,
            cwd / LOGO_FILENAME,
            cwd / "data" / LOGO_FILENAME,
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    return base_dir / LOGO_FILENAME


def build_logo_src(logo_path: Path) -> str:
    """Return a data URI for the logo to avoid filesystem encoding issues in PDF renderers."""
    if logo_path.is_file():
        try:
            encoded_logo = base64.b64encode(logo_path.read_bytes()).decode("ascii")
            return f"data:image/png;base64,{encoded_logo}"
        except Exception as exc:
            logger.warning(f"Advarsel: Klarte ikke å base64-enkode logo '{logo_path}': {exc}")

    logger.warning(
        f"Advarsel: Fant ikke logo '{logo_path}'. "
        f"Sett {LOGO_ENV_VAR} til riktig filsti hvis logoen ligger et annet sted. "
        "Fortsetter uten logobilde."
    )
    empty_svg = "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='1'></svg>"
    encoded_svg = base64.b64encode(empty_svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded_svg}"


def table_to_df(
    wb,
    table_name: str | None = None,
    allow_any_totalt_table: bool = True,
) -> pd.DataFrame | None:
    for ws in wb.worksheets:
        tables = ws.tables
        if not tables:
            continue

        if table_name:
            for name, table in tables.items():
                if name.lower() == table_name.lower():
                    table_ref = table.ref if hasattr(table, "ref") else str(table)
                    values = [[cell.value for cell in row] for row in ws[table_ref]]
                    if values:
                        header, *data = values
                        return pd.DataFrame(data, columns=header)

        if allow_any_totalt_table:
            for table in tables.values():
                table_ref = table.ref if hasattr(table, "ref") else str(table)
                values = [[cell.value for cell in row] for row in ws[table_ref]]
                if not values:
                    continue
                header, *data = values
                if "Totalt" in [str(h) if h is not None else "" for h in header]:
                    return pd.DataFrame(data, columns=header)

    return None


def load_table_from_workbook(file_path: str, table_name: str) -> pd.DataFrame:
    wb = load_workbook(filename=file_path, data_only=True, read_only=False)
    try:
        table_df = table_to_df(wb, table_name=table_name, allow_any_totalt_table=False)
        if table_df is None:
            raise ValueError(f"Fant ikke tabell '{table_name}' i {file_path}.")
        return table_df
    finally:
        wb.close()


def load_named_range_to_df(
    file_path: str,
    range_name: str,
    allow_generic_totalt_fallback: bool = True,
) -> pd.DataFrame:
    wb = load_workbook(filename=file_path, data_only=True, read_only=False)

    def resolve_defined_name(name: str) -> str | None:
        all_names = list(wb.defined_names.keys())

        if name in wb.defined_names:
            return name

        lowered = {n.lower(): n for n in all_names}
        if name.lower() in lowered:
            return lowered[name.lower()]

        # Common case: user writes "Sammendrag", workbook contains "sammendraget".
        for n in all_names:
            if n.lower().startswith(name.lower()) or name.lower().startswith(n.lower()):
                return n

        # Flexible match for names with separators/extra words, e.g. "Sist_uke" -> "Alle_Siste_Ukes_Poeng".
        def normalize_token(text: str) -> str:
            return re.sub(r"[^a-z0-9]", "", text.lower())

        target = normalize_token(name)
        for n in all_names:
            candidate = normalize_token(n)
            if target and target in candidate:
                return n

        # Token match: "Sist_uke" should match names like "Alle_Siste_Ukes_Poeng".
        query_tokens = [t for t in re.split(r"[^a-z0-9]+", name.lower()) if t]
        for n in all_names:
            candidate_tokens = [t for t in re.split(r"[^a-z0-9]+", n.lower()) if t]
            if query_tokens and all(
                any(ct.startswith(qt) or qt.startswith(ct) for ct in candidate_tokens)
                for qt in query_tokens
            ):
                return n

        return None

    resolved_name = resolve_defined_name(range_name)
    if resolved_name is None:
        wb.close()
        raise ValueError(
            f"Fant ikke navngitt område '{range_name}' i {file_path}. "
            f"Tilgjengelige navn: {', '.join(wb.defined_names.keys())}"
        )

    defn = wb.defined_names[resolved_name]
    rows = []

    for sheet_name, cell_range in defn.destinations:
        ws = wb[sheet_name]
        for row in ws[cell_range]:
            rows.append([cell.value for cell in row])

    if not rows:
        # Some workbooks keep stale defined names (#REF!). Fall back to tables.
        fallback_df = table_to_df(
            wb,
            table_name=range_name,
            allow_any_totalt_table=allow_generic_totalt_fallback,
        )
        wb.close()
        if fallback_df is not None:
            return fallback_df
        raise ValueError(f"Navngitt område '{range_name}' er tomt i {file_path}.")

    header, *data = rows
    df = pd.DataFrame(data, columns=header)
    wb.close()
    return df


def load_weekly_winner_counts_from_named_range(
    file_path: str,
    range_name: str = "Ukevinnere",
) -> pd.DataFrame | None:
    """Build a fallback weekly scoreboard from a winner history range.

    Expected shape is date in first column and winner code/name in second column.
    The function aggregates winner occurrences into columns Navn/Totalt so it can
    reuse the regular section renderer.
    """
    wb = load_workbook(filename=file_path, data_only=True, read_only=False)
    try:
        if range_name not in wb.defined_names:
            lowered = {n.lower(): n for n in wb.defined_names.keys()}
            resolved = lowered.get(range_name.lower())
        else:
            resolved = range_name

        if not resolved:
            return None

        defn = wb.defined_names[resolved]
        rows: list[list[object]] = []
        for sheet_name, cell_range in defn.destinations:
            ws = wb[sheet_name]
            for row in ws[cell_range]:
                rows.append([cell.value for cell in row])

        if not rows:
            return None

        winners: list[str] = []
        for row in rows:
            if len(row) < 2:
                continue
            winner = row[1]
            if winner is None:
                continue

            winner_text = str(winner).strip()
            if not winner_text or winner_text.lower() in {"nan", "none", "null"}:
                continue

            winners.append(winner_text.upper())

        if not winners:
            return None

        counts = pd.Series(winners, dtype="object").value_counts().reset_index()
        counts.columns = ["Navn", "Totalt"]
        return counts
    finally:
        wb.close()


def load_latest_weekly_from_player_tables(file_path: str) -> pd.DataFrame | None:
    """Build a weekly table from the latest row in each per-player history table.

    Expected source tables are in sheets like f_rb/f_og and have columns such as
    Ukenr, Dato, hp, up, bp, straff, tp, Rank, Endring.
    """
    wb = load_workbook(filename=file_path, data_only=True, read_only=False)
    try:
        rows_out: list[dict[str, object]] = []

        for ws in wb.worksheets:
            if not ws.tables:
                continue

            for table_name, table in ws.tables.items():
                table_lower = table_name.lower()
                if not table_lower.startswith("t_"):
                    continue

                table_ref = table.ref if hasattr(table, "ref") else str(table)
                values = [[cell.value for cell in row] for row in ws[table_ref]]
                if not values:
                    continue

                header, *data = values
                if not data:
                    continue

                df_local = pd.DataFrame(data, columns=header)
                df_local.columns = [
                    str(c).replace("\xa0", " ").replace("\n", " ").strip() if c is not None else ""
                    for c in df_local.columns
                ]

                if "tp" not in df_local.columns or "Ukenr" not in df_local.columns:
                    continue

                df_local["tp"] = pd.to_numeric(df_local["tp"], errors="coerce")
                df_local["Ukenr"] = pd.to_numeric(df_local["Ukenr"], errors="coerce")
                df_local = df_local.dropna(subset=["Ukenr", "tp"])
                if df_local.empty:
                    continue

                latest_idx = df_local["Ukenr"].idxmax()
                latest_row = df_local.loc[latest_idx]

                player_code = table_name.split("_", 1)[1] if "_" in table_name else ws.title
                antall_12 = latest_row.get("Antall 12", None)
                antall_11 = latest_row.get("Antall 11", None)
                antall_10 = latest_row.get("Antall 10", None)
                premie_12 = latest_row.get("12-premie", None)
                premie_11 = latest_row.get("11-premie", None)
                premie_10 = latest_row.get("10-premie", None)
                lev_bonus = latest_row.get("Premie2", latest_row.get("Premie", None))

                output_row: dict[str, object] = {
                    "Navn": str(player_code).upper(),
                    "hp": latest_row.get("hp", None),
                    "up": latest_row.get("up", None),
                    "bp": latest_row.get("bp", None),
                    "straff": latest_row.get("straff", None),
                    "tp": latest_row.get("tp", None),
                    "Tolvere": antall_12,
                    "Elvere": antall_11,
                    "Tiere": antall_10,
                    "12p": premie_12,
                    "11p": premie_11,
                    "10p": premie_10,
                    "12LEV": None,
                    "11LEV": None,
                    "10LEV": None,
                    "LEVp": lev_bonus,
                }
                tp_val = pd.to_numeric(pd.Series([output_row.get("tp")]), errors="coerce").iloc[0]
                levp_val = pd.to_numeric(pd.Series([output_row.get("LEVp")]), errors="coerce").iloc[0]
                output_row["Totalt"] = (
                    (0.0 if pd.isna(tp_val) else float(tp_val))
                    + (0.0 if pd.isna(levp_val) else float(levp_val))
                )
                rows_out.append(output_row)

        if not rows_out:
            return None

        out_df = pd.DataFrame(rows_out)
        for col in [
            "hp",
            "up",
            "bp",
            "straff",
            "tp",
            "Tolvere",
            "Elvere",
            "Tiere",
            "12p",
            "11p",
            "10p",
            "12LEV",
            "11LEV",
            "10LEV",
            "LEVp",
            "Totalt",
        ]:
            if col in out_df.columns:
                out_df[col] = pd.to_numeric(out_df[col], errors="coerce")

        out_df = out_df.sort_values(by="tp", ascending=False).reset_index(drop=True)
        return out_df
    finally:
        wb.close()


def load_weekly_matrix_from_header_range(file_path: str) -> pd.DataFrame | None:
    """Find a weekly matrix by matching the expected header row in workbook cells.

    This targets layouts like:
    Navn, hp, up, bp, straff, tp, Tolvere, Elvere, Tiere, 12p, 11p, 10p,
    12LEV, 11LEV, 10LEV, LEVp, Totalt.
    """
    wb = load_workbook(filename=file_path, data_only=True, read_only=False)
    try:
        required = ["Navn", "hp", "up", "bp", "straff", "tp"]
        extras = ["Tolvere", "Elvere", "Tiere", "12p", "11p", "10p", "12LEV", "11LEV", "10LEV", "LEVp", "Totalt"]

        for ws in wb.worksheets:
            max_row = min(ws.max_row, 600)
            max_col = min(ws.max_column, 80)

            for r in range(1, max_row + 1):
                row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
                row_text = [str(v).strip() if v is not None else "" for v in row_vals]
                lowered = [t.lower() for t in row_text]

                def find_col(name: str) -> int | None:
                    needle = name.lower()
                    for idx, text in enumerate(lowered):
                        if text == needle:
                            return idx
                    return None

                required_indices = [find_col(col_name) for col_name in required]
                if any(idx is None for idx in required_indices):
                    continue

                has_extra = any(find_col(col_name) is not None for col_name in extras)
                if not has_extra:
                    continue

                all_cols = required + extras
                selected_cols: list[tuple[str, int]] = []
                for col_name in all_cols:
                    idx = find_col(col_name)
                    if idx is not None:
                        selected_cols.append((col_name, idx))

                if not selected_cols:
                    continue

                data_rows: list[list[object]] = []
                for rr in range(r + 1, max_row + 1):
                    row_data = [ws.cell(rr, col_idx + 1).value for _, col_idx in selected_cols]
                    navn_idx = next((i for i, (name, _) in enumerate(selected_cols) if name == "Navn"), None)
                    navn_val = row_data[navn_idx] if navn_idx is not None else None
                    if navn_val is None or str(navn_val).strip() == "":
                        if data_rows:
                            break
                        continue
                    data_rows.append(row_data)

                if not data_rows:
                    continue

                out_df = pd.DataFrame(data_rows, columns=[name for name, _ in selected_cols])
                for numeric_col in [
                    "hp",
                    "up",
                    "bp",
                    "straff",
                    "tp",
                    "Tolvere",
                    "Elvere",
                    "Tiere",
                    "12p",
                    "11p",
                    "10p",
                    "12LEV",
                    "11LEV",
                    "10LEV",
                    "LEVp",
                    "Totalt",
                ]:
                    if numeric_col in out_df.columns:
                        out_df[numeric_col] = pd.to_numeric(out_df[numeric_col], errors="coerce")

                if "Totalt" in out_df.columns:
                    out_df = out_df.sort_values(by="Totalt", ascending=False).reset_index(drop=True)
                elif "tp" in out_df.columns:
                    out_df = out_df.sort_values(by="tp", ascending=False).reset_index(drop=True)

                # Heuristic: reject season-summary tables when looking for weekly table.
                metric_col = "tp" if "tp" in out_df.columns else ("Totalt" if "Totalt" in out_df.columns else None)
                if metric_col is not None:
                    metric_max = pd.to_numeric(out_df[metric_col], errors="coerce").max()
                    if pd.notna(metric_max) and float(metric_max) > 3000.0:
                        continue

                return out_df

        return None
    finally:
        wb.close()


def prepare_report_df(input_df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    df_local = input_df.copy()
    df_local.columns = [
        str(c).replace("\xa0", " ").replace("\n", " ").strip() if c is not None else ""
        for c in df_local.columns
    ]

    non_empty_headers = [c for c in df_local.columns if c and c.lower() != "nan"]
    unnamed_count = sum(1 for c in df_local.columns if c.lower().startswith("unnamed:"))
    if (len(non_empty_headers) <= 2 or unnamed_count >= max(1, len(df_local.columns) // 2)) and len(df_local) > 0:
        promoted = [
            str(c).replace("\xa0", " ").replace("\n", " ").strip()
            if c is not None
            else ""
            for c in df_local.iloc[0].tolist()
        ]
        if any(promoted):
            df_local.columns = promoted
            df_local = df_local.iloc[1:].reset_index(drop=True)

    df_local = df_local.rename(columns={df_local.columns[0]: "Navn"})
    df_local["Navn"] = df_local["Navn"].astype(str).str.strip()
    df_local = df_local[
        ~df_local["Navn"].str.lower().isin({"", "none", "nan", "null"})
    ]

    total_col_local = None
    for c in df_local.columns:
        lc = c.lower().strip()
        if lc == "totalt" or "totalt" in lc:
            total_col_local = c
            break

    if total_col_local is None:
        raise KeyError(f"Fant ikke 'Totalt'-kolonne. Kolonner i datasettet: {list(df_local.columns)}")

    df_local[total_col_local] = pd.to_numeric(df_local[total_col_local], errors="coerce")
    df_local = df_local.dropna(subset=[total_col_local])

    df_local = (
        df_local.sort_values(by=total_col_local, ascending=False)
        .drop_duplicates(subset=["Navn"], keep="first")
        .sort_values(by=total_col_local, ascending=False)
        .reset_index(drop=True)
    )

    return df_local, total_col_local


def render_section(
    df_local: pd.DataFrame,
    total_col_local: str,
    title: str,
    y_max: float = 25000.0,
) -> tuple[str, str]:
    fig, ax = plt.subplots(figsize=(7, 4))

    bar_colors = ["#d4af37", "#c0c0c0", "#cd7f32"] + ["#175c5f"] * max(0, len(df_local) - 3)
    bars = ax.bar(df_local["Navn"], df_local[total_col_local], color=bar_colors[: len(df_local)], width=0.6)

    legend_handles = []
    if len(df_local) >= 1:
        legend_handles.append(Patch(facecolor="#d4af37", label="1. plass"))
    if len(df_local) >= 2:
        legend_handles.append(Patch(facecolor="#c0c0c0", label="2. plass"))
    if len(df_local) >= 3:
        legend_handles.append(Patch(facecolor="#cd7f32", label="3. plass"))
    if legend_handles:
        ax.legend(handles=legend_handles, loc="upper right", frameon=True)

    ax.set_title(title)
    ax.set_xlabel(df_local.columns[0])
    ax.set_ylabel(total_col_local)

    y_min = 0.0
    ax.set_ylim(y_min, y_max)

    label_offset = max((y_max - y_min) * 0.01, 1.0)
    for bar, value in zip(bars, df_local[total_col_local]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            float(value) + label_offset,
            f"{int(round(float(value)))}",
            ha="center",
            va="bottom",
            fontsize=7,
            color="#222222",
        )

    plt.tight_layout()

    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format="png")
    img_buffer.seek(0)
    img_base64_local = base64.b64encode(img_buffer.getvalue()).decode("utf-8")
    plt.close(fig)

    display_df_local = df_local.copy()
    numeric_view = display_df_local.apply(lambda col: pd.to_numeric(col, errors="coerce"))
    display_df_local = display_df_local.mask(numeric_view.eq(0))
    display_df_local = display_df_local.replace({None: ""}).fillna("")
    display_df_local = display_df_local.replace(
        to_replace=r"^\s*(none|null|nan)\s*$",
        value="",
        regex=True,
    )
    display_df_local = display_df_local.replace(to_replace=r"^\s*0(?:\.0+)?\s*$", value="", regex=True)

    html_table_local = display_df_local.to_html(
        classes="styled-table",
        index=False,
        na_rep="",
        float_format="{:.0f}".format,
    )

    section_html_local = f"""
    <div class="chart-container">
        <img src="data:image/png;base64,{img_base64_local}" alt="{title}">
    </div>
    <h2>{title}</h2>
    <div class="table-wrap">
        {html_table_local}
    </div>
    """

    return section_html_local, img_base64_local


def build_weekly_section_html(source_workbook: str | None, report_week: int) -> str:
    """Build the weekly section using a prioritized fallback chain."""
    if source_workbook is None:
        return (
            "<p style='color:#555; text-align:left;'>"
            "Fant ingen kildefil for navngitt område 'Sist_uke'."
            "</p>"
        )

    weekly_error = None
    weekly_df = None
    weekly_total_col = None

    # 1) Preferred source: named ranges that should contain this week's full table.
    for candidate in WEEKLY_CANDIDATES:
        try:
            weekly_raw_df = load_named_range_to_df(
                source_workbook,
                candidate,
                allow_generic_totalt_fallback=False,
            )
            weekly_df, weekly_total_col = prepare_report_df(weekly_raw_df)
            if len(weekly_df) > 0:
                logger.info(f"Info: Bruker ukesdata fra navngitt område '{candidate}'.")
                break
        except Exception as exc:
            weekly_error = exc

    if weekly_df is not None and weekly_total_col is not None and len(weekly_df) > 0:
        weekly_section_html, _ = render_section(
            weekly_df,
            weekly_total_col,
            f"Sist uke - Uke {report_week} - Tippernes resultater",
            y_max=2000.0,
        )
        return weekly_section_html

    # 2) Secondary source: legacy workbook table used in some files.
    try:
        weekly_raw_df = load_table_from_workbook(source_workbook, "Summasummarium5")
        weekly_df, weekly_total_col = prepare_report_df(weekly_raw_df)
        weekly_section_html, _ = render_section(
            weekly_df,
            weekly_total_col,
            f"Sist uke - Uke {report_week} - Tippernes resultater",
            y_max=2000.0,
        )
        logger.info("Info: Bruker ukesdata fra tabell 'Summasummarium5'.")
        return weekly_section_html
    except Exception as table_exc:
        # 3) Preferred fallback: look for the exact weekly matrix layout in workbook cells.
        weekly_matrix_df = load_weekly_matrix_from_header_range(source_workbook)
        if weekly_matrix_df is None and Path(WORKBOOK_FILE).exists():
            weekly_matrix_df = load_weekly_matrix_from_header_range(WORKBOOK_FILE)

        if weekly_matrix_df is not None and len(weekly_matrix_df) > 0:
            metric_col = "Totalt" if "Totalt" in weekly_matrix_df.columns else "tp"
            weekly_section_html, _ = render_section(
                weekly_matrix_df,
                metric_col,
                f"Sist uke - Uke {report_week} - Tippernes resultater",
                y_max=max(float(pd.to_numeric(weekly_matrix_df[metric_col], errors='coerce').max()) + 200.0, 1000.0),
            )
            logger.info("Info: Bruker fallback-data fra ukematrise med Navn/hp/up/bp/straff/tp-kolonner.")
            return weekly_section_html

        # 4) Fallback: reconstruct this week's table from player history tables.
        player_weekly_df = load_latest_weekly_from_player_tables(source_workbook)
        if player_weekly_df is not None and len(player_weekly_df) > 0:
            weekly_section_html, _ = render_section(
                player_weekly_df,
                "tp",
                f"Sist uke - Uke {report_week} - Tippernes resultater",
                y_max=max(float(player_weekly_df["tp"].max()) + 200.0, 1000.0),
            )
            logger.info("Info: Bruker fallback-data fra spillerhistorikk-tabeller (f_*/t_*).")
            return weekly_section_html

        # 5) Last fallback: winner counts only (reduced column set).
        winner_counts_df = load_weekly_winner_counts_from_named_range(source_workbook, "Ukevinnere")
        if winner_counts_df is not None and len(winner_counts_df) > 0:
            weekly_section_html, _ = render_section(
                winner_counts_df,
                "Totalt",
                f"Ukevinnere hittil - Uke {report_week}",
                y_max=max(float(winner_counts_df["Totalt"].max()) + 1.0, 5.0),
            )
            logger.info("Info: Bruker fallback-data fra navngitt område 'Ukevinnere'.")
            return weekly_section_html

        if weekly_error is not None:
            logger.warning(f"Advarsel: Klarte ikke å bygge seksjon for 'Sist_uke': {weekly_error}")
        logger.warning(f"Advarsel: Klarte heller ikke å bruke tabell 'Summasummarium5': {table_exc}")
        return (
            "<p style='color:#555; text-align:left;'>"
            "Fant ikke gyldige data for navngitt område 'Sist_uke'."
            "</p>"
        )


def export_html_to_pdf(html_path: str, pdf_path: str) -> None:
    try:
        from weasyprint import HTML  # type: ignore

        HTML(filename=html_path).write_pdf(pdf_path)
        logger.info(f"PDF eksportert: '{pdf_path}' (weasyprint).")
        return
    except Exception:
        pass

    wkhtmltopdf_path = resolve_wkhtmltopdf_path()
    if wkhtmltopdf_path:
        try:
            subprocess.run(
                [
                    wkhtmltopdf_path,
                    "--enable-local-file-access",
                    "--print-media-type",
                    "--orientation",
                    "Portrait",
                    html_path,
                    pdf_path,
                ],
                check=True,
            )
            logger.info(f"PDF eksportert: '{pdf_path}' (wkhtmltopdf).")
            return
        except Exception as exc:
            logger.warning(f"Advarsel: PDF-eksport feilet via wkhtmltopdf: {exc}")
            return

    logger.info(
        "Info: PDF ikke laget automatisk. Installer 'weasyprint' eller 'wkhtmltopdf' "
        "for PDF-eksport."
    )


def resolve_wkhtmltopdf_path() -> str | None:
    candidates: list[Path] = []

    env_path = os.environ.get("WKHTMLTOPDF_PATH")
    if env_path:
        candidates.append(Path(env_path))

    which_path = shutil.which("wkhtmltopdf")
    if which_path:
        candidates.append(Path(which_path))

    candidates.extend(
        [
            Path(r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"),
            Path(r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    return None


def pdf_export_available() -> bool:
    return importlib.util.find_spec("weasyprint") is not None or resolve_wkhtmltopdf_path() is not None


file_exists = Path(file).exists()
should_refresh = refresh_on_start or not file_exists

# --- 1) Load and refresh input data from workbook ---
if should_refresh:
    if not file_exists:
        logger.info("Info: stps_tolk.xlsx mangler. Prøver å opprette den fra 2027-kilde.")
        startup_df = update_stps_tolk_on_start(file, sheet_name)
        if startup_df is not None:
            df = startup_df.copy()
        else:
            if not Path(file).exists():
                raise FileNotFoundError(
                    "stps_tolk.xlsx finnes ikke, og automatisk oppretting fra 2027-kilde feilet. "
                    "Kjør med en gyldig kildefil tilgjengelig (f.eks. STPS_2027.xlsm)."
                )
            df = pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")
    else:
        source_workbook_for_refresh = find_source_workbook()
        if source_workbook_for_refresh is not None:
            refresh_linked_workbooks_via_excel(source_workbook_for_refresh, file)
        else:
            logger.info("Info: Fant ingen kildefil å åpne før stps_tolk.xlsx for link-oppdatering.")

        logger.info("Info: Leser data direkte fra stps_tolk.xlsx (ingen manuell omskriving av arket).")
        df = pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")
else:
    logger.info("Info: Leser direkte fra stps_tolk.xlsx. Bruk --refresh for å hente fra 2027-kilde.")
    df = pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")

# --- 2) Prepare runtime metadata and output filenames ---
now = datetime.now()
report_updated_at = now.strftime("%d.%m.%Y %H:%M:%S")
report_week = now.isocalendar().week
report_year = now.year
pdf_ready_hint = pdf_export_available()
latest_report_html_filename = "report_2027.html"
latest_report_pdf_filename = "report_2027.pdf"
report_html_filename = build_weekly_filename("report", ".html", report_week, report_year)
report_pdf_filename = build_weekly_filename("report", ".pdf", report_week, report_year)

# --- 3) Build summary section (chart + table) ---
summary_df, summary_total_col = prepare_report_df(df)
summary_section_html, _ = render_section(
    summary_df,
    summary_total_col,
    f"Sammendrag - Uke {report_week} - Tippernes resultater 2026/2027",
)

# --- 4) Build weekly section via fallback chain ---
source_workbook = find_source_workbook()
weekly_section_html = build_weekly_section_html(source_workbook, report_week)

# --- 5) Render complete report HTML from template ---
html_content = render_report_html(
    REPORT_TEMPLATE_PATH,
    build_logo_src(resolve_logo_path()),
    report_week,
    latest_report_pdf_filename,
    pdf_ready_hint,
    report_updated_at,
    summary_section_html,
    weekly_section_html,
)

# --- 6) Write HTML outputs and export PDF ---
with open(report_html_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

with open(latest_report_html_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

export_html_to_pdf(report_html_filename, report_pdf_filename)
shutil.copy2(report_pdf_filename, latest_report_pdf_filename)

# --- 7) Save weekly workbook snapshot when source is xlsm ---
if source_workbook is not None and source_workbook.lower().endswith(".xlsm"):
    workbook_copy = create_weekly_workbook_copy(source_workbook, report_week, report_year)
    if workbook_copy is not None:
        logger.info(f"Kopierte kildearbeidsbok til '{workbook_copy}'.")

logger.info(f"HTML report successfully generated as '{report_html_filename}'!")
