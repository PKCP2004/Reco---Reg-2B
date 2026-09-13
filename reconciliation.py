from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
from typing import Iterable
from difflib import SequenceMatcher

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter


FIELD_ALIASES = {
    "gstin": ["gstin", "gst no", "gst number", "supplier gstin", "ctin"],
    "invoice_no": ["invoice no", "invoice number", "inv no", "document no", "bill no", "inum"],
    "invoice_date": ["invoice date", "date", "document date", "idt"],
    "two_b_month": ["2b month", "return period", "period", "month"],
    "taxable": ["taxable value", "taxable", "taxable amount"],
    "igst": ["igst", "igst amount"],
    "cgst": ["cgst", "cgst amount"],
    "sgst": ["sgst", "sgst amount", "utgst", "sgst/utgst"],
    "cess": ["cess", "cess amount"],
    "total": ["invoice value", "invoice amount", "total value", "total amount"],
}

VALUE_FIELDS = ["taxable", "igst", "cgst", "sgst", "cess", "total"]
COMPARISON_FIELDS = ["gstin", "invoice_no", "invoice_date", "two_b_month", *VALUE_FIELDS]


def _clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def _number(value) -> float:
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    cleaned = re.sub(r"[^0-9.\-]", "", str(value).replace(",", ""))
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return 0.0


def _find_column(columns: Iterable, aliases: list[str]):
    normalized = {str(column).strip().lower(): column for column in columns}
    for alias in aliases:
        for name, original in normalized.items():
            if alias == name or alias in name:
                return original
    return None


def standardize_frame(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    frame = frame.copy()
    output = pd.DataFrame(index=frame.index)
    for field, aliases in FIELD_ALIASES.items():
        column = _find_column(frame.columns, aliases)
        output[field] = frame[column] if column is not None else ""

    output["gstin"] = output["gstin"].map(_clean_text)
    output["invoice_no"] = output["invoice_no"].map(_clean_text)
    output["invoice_date"] = pd.to_datetime(output["invoice_date"], errors="coerce", dayfirst=True).dt.strftime("%Y-%m-%d").fillna("")
    output["two_b_month"] = output["two_b_month"].fillna("").astype(str).str.strip()
    for field in VALUE_FIELDS:
        output[field] = output[field].map(_number)

    output["source"] = source
    output["source_row"] = frame.index + 2
    output = output[(output["gstin"] != "") | (output["invoice_no"] != "")].reset_index(drop=True)
    output["match_key"] = output["gstin"] + "|" + output["invoice_no"]
    return output


def reconcile(gstr2b: pd.DataFrame, register: pd.DataFrame, tolerance: float = 1.0, progress_callback=None) -> dict[str, pd.DataFrame]:
    two_b = gstr2b.copy()
    inward = register.copy()
    two_b["_side"] = "GSTR-2B"
    inward["_side"] = "Inward Register"

    two_b = two_b.reset_index(drop=True)
    inward = inward.reset_index(drop=True)
    two_b["_row_id"] = two_b.index
    inward["_row_id"] = inward.index
    inward["reco_status"] = ""
    inward["likely_2b_invoice_no"] = ""
    inward["matched_2b_month"] = ""
    inward["matched_2b_row_id"] = pd.Series([None] * len(inward), index=inward.index, dtype="object")
    inward["matched_2b_link"] = ""
    inward["reco_confidence"] = pd.Series([None] * len(inward), index=inward.index, dtype="object")
    inward["reco_basis"] = ""

    matched_register_rows = set()
    results = []
    two_b_records = two_b.to_dict("records")
    inward_records = inward.to_dict("index")
    register_by_key = {}
    register_by_invoice = {}
    register_by_gstin = {}
    register_by_invoice_prefix = {}
    for row_id, register_row in inward_records.items():
        register_by_key.setdefault(register_row["match_key"], []).append(row_id)
        register_by_invoice.setdefault(register_row["invoice_no"], []).append(row_id)
        register_by_gstin.setdefault(register_row["gstin"], []).append(row_id)
        invoice_prefix = register_row["invoice_no"][:4]
        if invoice_prefix:
            register_by_invoice_prefix.setdefault(invoice_prefix, []).append(row_id)

    def comparison_fields(two_b_row=None, register_row=None):
        values = {}
        for field in COMPARISON_FIELDS:
            values[f"2B_{field}"] = "" if two_b_row is None else two_b_row.get(field, "")
            values[f"Register_{field}"] = "" if register_row is None else register_row.get(field, "")
        return values

    def value_differences(two_b_row, register_row):
        return {
            field: round(two_b_row[field] - register_row[field], 2)
            for field in VALUE_FIELDS
            if abs(two_b_row[field] - register_row[field]) > tolerance
        }

    def date_warning(two_b_row, register_row):
        left = two_b_row.get("invoice_date", "")
        right = register_row.get("invoice_date", "")
        if not left or not right or left == right:
            return ""
        return f"Date differs: 2B {left} | Register {right}"

    def append_result(two_b_row=None, register_row=None, status="", differences=None, confidence=0, basis="", action=""):
        primary = two_b_row if two_b_row is not None else register_row
        differences = differences or {}
        result = {
            **(primary if isinstance(primary, dict) else primary.to_dict()),
            **comparison_fields(two_b_row, register_row),
            "status": status,
            "confidence": confidence,
            "match_basis": basis,
            "difference": "; ".join(f"{field}: {value:+.2f}" for field, value in differences.items()),
            "date_warning": "" if two_b_row is None or register_row is None else date_warning(two_b_row, register_row),
            "register_row": "" if register_row is None else str(register_row["source_row"]),
            "recommended_action": action,
        }
        results.append(result)

    def annotate_register(register_row, two_b_row, status, confidence, basis):
        row_id = register_row["_row_id"]
        likely_invoice = two_b_row["invoice_no"] if two_b_row is not None else ""
        inward.at[row_id, "reco_status"] = status
        inward.at[row_id, "likely_2b_invoice_no"] = likely_invoice
        inward.at[row_id, "matched_2b_month"] = two_b_row["two_b_month"] if two_b_row is not None else ""
        inward.at[row_id, "matched_2b_row_id"] = two_b_row["_row_id"] if two_b_row is not None else None
        inward.at[row_id, "matched_2b_link"] = "Open GSTR-2B invoice" if two_b_row is not None else ""
        inward.at[row_id, "reco_confidence"] = confidence
        inward.at[row_id, "reco_basis"] = basis
        register_row["reco_status"] = status
        register_row["likely_2b_invoice_no"] = likely_invoice
        register_row["matched_2b_month"] = two_b_row["two_b_month"] if two_b_row is not None else ""
        register_row["matched_2b_row_id"] = two_b_row["_row_id"] if two_b_row is not None else None
        register_row["matched_2b_link"] = "Open GSTR-2B invoice" if two_b_row is not None else ""
        register_row["reco_confidence"] = confidence
        register_row["reco_basis"] = basis

    def combined_register_row(row_ids):
        combined = dict(inward_records[row_ids[0]])
        combined["source_row"] = ", ".join(str(inward_records[row_id]["source_row"]) for row_id in row_ids)
        combined["split_count"] = len(row_ids)
        for field in VALUE_FIELDS:
            combined[field] = round(sum(inward_records[row_id][field] for row_id in row_ids), 2)
        return combined

    total_two_b = max(len(two_b), 1)
    processed_two_b = 0
    progress_interval = max(1, len(two_b) // 100)
    for source_row in two_b_records:
        key = source_row["match_key"]
        exact_ids = [row_id for row_id in register_by_key.get(key, []) if row_id not in matched_register_rows]
        target_ids = exact_ids
        target = combined_register_row(target_ids) if target_ids else None
        if target is not None:
            matched_register_rows.update(target_ids)
            differences = value_differences(source_row, target)
            status = "Value Mismatch" if differences else "Matched"
            basis = "Exact GSTIN + invoice number" if len(target_ids) == 1 else f"Split inward invoice ({len(target_ids)} lines) summed to 2B"
            for target_id in target_ids:
                annotate_register(inward_records[target_id], source_row, status, 100, basis)
            append_result(
                two_b_row=source_row,
                register_row=target,
                status=status,
                differences=differences,
                confidence=100,
                basis=basis,
                action="Review tax values" if differences else ("Review date" if date_warning(source_row, target) else "No action"),
            )
            processed_two_b += 1
            if progress_callback and (processed_two_b % progress_interval == 0 or processed_two_b == len(two_b)):
                progress_callback(0.25 + 0.60 * (processed_two_b / total_two_b), f"Matching GSTR-2B invoice {processed_two_b:,} of {len(two_b):,}")
            continue

        # Score same-invoice and near-invoice candidates before declaring a row missing.
        candidates = []
        candidate_ids = set(register_by_invoice.get(source_row["invoice_no"], []))
        candidate_ids.update(register_by_gstin.get(source_row["gstin"], []))
        candidate_ids.update(register_by_invoice_prefix.get(source_row["invoice_no"][:4], []))
        for row_id in candidate_ids:
            if row_id in matched_register_rows:
                continue
            possible = inward_records[row_id]
            invoice_similarity = SequenceMatcher(None, source_row["invoice_no"], possible["invoice_no"]).ratio()
            same_invoice = source_row["invoice_no"] and source_row["invoice_no"] == possible["invoice_no"]
            same_gstin = source_row["gstin"] and source_row["gstin"] == possible["gstin"]
            differences = value_differences(source_row, possible)
            force_match = same_gstin and not differences and not same_invoice
            if same_invoice or force_match or (invoice_similarity >= 0.88 and (same_gstin or not differences)):
                score = (100 if same_invoice else 92 if force_match else 88) + (5 if same_gstin else 0) + (5 if not differences else 0)
                candidates.append((score, possible, differences, invoice_similarity, same_gstin, same_invoice))

        if candidates:
            _, target, differences, similarity, same_gstin, same_invoice = max(candidates, key=lambda item: item[0])
            matched_register_rows.add(target["_row_id"])
            if not same_gstin and not differences:
                status = "GSTIN Mismatch"
                basis = "Invoice number + tax values; GSTIN differs"
                action = "Verify supplier GSTIN"
            elif same_gstin and not same_invoice:
                status = "Likely Match - Register Invoice Error"
                basis = "Same GSTIN + identical taxable/tax values; invoice number differs"
                action = f"Confirm inward invoice number should be {source_row['invoice_no']}"
            elif differences:
                status = "Value Mismatch"
                basis = "Invoice number candidate"
                action = "Review tax values"
            else:
                status = "Matched"
                basis = "Invoice number + tax values"
                action = "Review date" if date_warning(source_row, target) else "No action"
            annotate_register(target, source_row, status, min(99, int(max(candidates, key=lambda item: item[0])[0])), basis)
            append_result(two_b_row=source_row, register_row=target, status=status, differences=differences, confidence=min(99, int(max(candidates, key=lambda item: item[0])[0])), basis=basis, action=action)
        else:
            append_result(source_row, status="In 2B Only", confidence=0, basis="No reliable register candidate", action="Check invoice number, GSTIN, or register entry")

        processed_two_b += 1
        if progress_callback and (processed_two_b % progress_interval == 0 or processed_two_b == len(two_b)):
            progress_callback(0.25 + 0.60 * (processed_two_b / total_two_b), f"Matching GSTR-2B invoice {processed_two_b:,} of {len(two_b):,}")

    if progress_callback:
        progress_callback(0.88, f"Checking {len(inward):,} inward-register rows for register-only items")

    for _, register_row in inward.iterrows():
        if register_row["_row_id"] not in matched_register_rows:
            annotate_register(register_row, None, "In Register Only", 0, "No reliable 2B candidate")
            append_result(register_row=register_row, status="In Register Only", confidence=0, basis="No reliable 2B candidate", action="Check 2B month or supplier filing")

    result = pd.DataFrame(results).drop(columns=["_side", "_row_id"], errors="ignore")
    if "two_b_month" not in result.columns:
        result["two_b_month"] = ""
    if not result.empty:
        result["_register_order"] = pd.to_numeric(result["register_row"], errors="coerce")
        result["_register_only"] = result["_register_order"].isna()
        result = result.sort_values(
            ["_register_only", "_register_order"], kind="stable"
        ).drop(columns=["_register_order", "_register_only"])
    inward_source = inward.drop(columns=["_side", "_row_id"], errors="ignore")
    matched_details = inward_source[inward_source["reco_status"] == "Matched"].copy()
    if not matched_details.empty:
        two_b_details = gstr2b.drop(columns=["_side", "_row_id"], errors="ignore")[[
            "match_key", "gstin", "invoice_no", "invoice_date", "two_b_month",
            *VALUE_FIELDS,
        ]].rename(columns={
            column: f"2B_{column}" for column in [
                "gstin", "invoice_no", "invoice_date", "two_b_month", *VALUE_FIELDS,
            ]
        })
        matched_details = matched_details.merge(two_b_details, on="match_key", how="left")
        matched_details = matched_details.rename(columns={
            "gstin": "Register GSTIN",
            "invoice_no": "Register Invoice No",
            "source_row": "Register Source Row",
            "reco_status": "Status",
            "likely_2b_invoice_no": "Matched 2B Invoice No",
            "reco_confidence": "Match Confidence",
            "reco_basis": "Match Basis",
        })
    monthly_analysis = inward_source.copy()
    monthly_analysis["matched_2b_month"] = monthly_analysis["matched_2b_month"].replace("", "Not found in 2B")
    monthly_analysis = monthly_analysis.groupby(["matched_2b_month", "reco_status"], dropna=False).agg(
        Invoice_Lines=("invoice_no", "size"),
        Taxable_Value=("taxable", "sum"),
        IGST=("igst", "sum"),
        CGST=("cgst", "sum"),
        SGST=("sgst", "sum"),
        Cess=("cess", "sum"),
        Invoice_Value=("total", "sum"),
    ).reset_index().rename(columns={"matched_2b_month": "2B Month", "reco_status": "Status"})
    duplicate_keys = two_b[two_b.duplicated("match_key", keep=False)]["match_key"]
    if not duplicate_keys.empty:
        result.loc[result["match_key"].isin(duplicate_keys), "status"] = "Duplicate in 2B"
    if progress_callback:
        progress_callback(0.93, "Building reconciliation result tables")

    return {
        "reconciliation": result,
        "gstr2b_source": gstr2b.drop(columns=["_side", "_row_id"], errors="ignore"),
        "inward_register_source": inward_source,
        "matched_details": matched_details,
        "monthly_analysis": monthly_analysis,
        "matched": result[result["status"] == "Matched"].copy(),
        "mismatches": result[result["status"].isin(["Value Mismatch", "GSTIN Mismatch", "Likely Match - Register Invoice Error", "Likely Match - Invoice Format", "Duplicate in 2B"])].copy(),
        "missing_in_2b": result[result["status"] == "In Register Only"].copy(),
        "missing_in_register": result[result["status"] == "In 2B Only"].copy(),
    }


def read_upload(upload) -> pd.DataFrame:
    suffix = Path(upload.name).suffix.lower()
    content = upload.getvalue()
    if suffix == ".csv":
        return pd.read_csv(BytesIO(content))
    return pd.read_excel(BytesIO(content), sheet_name=0)


def export_results(results: dict[str, pd.DataFrame], progress_callback=None) -> bytes:
    output = BytesIO()
    brand = "9C002C"
    summary = pd.DataFrame([{
        "GSTR-2B rows": len(results["missing_in_register"]) + len(results["matched"]) + len(results["mismatches"]),
        "Inward register rows": len(results["missing_in_2b"]) + len(results["matched"]) + len(results["mismatches"]),
        "Matched": len(results["matched"]),
        "Review mismatches": len(results["mismatches"]),
        "GSTIN mismatches": int((results["reconciliation"]["status"] == "GSTIN Mismatch").sum()),
        "Missing in 2B": len(results["missing_in_2b"]),
        "Missing in register": len(results["missing_in_register"]),
    }])

    logo_path = Path(__file__).parent / "assets" / "PKP_logo.png"

    sheets = [("Summary", summary)]
    sheets.extend(
        (name[:31], frame)
        for name, frame in results.items()
        if name not in {"reconciliation", "matched", "gstr2b_source", "inward_register_source"}
    )
    sheets.extend([
        ("GSTR_2B_Upload", results["gstr2b_source"]),
        ("Inward_Register_Upload", results["inward_register_source"]),
        ("All_Reconciliation", results["reconciliation"]),
    ])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        workbook.set_properties({
            "title": "GST Reconciliation Studio",
            "subject": "GSTR-2B vs Inward Register Reconciliation",
            "author": "PUSHPAK KUMAR",
            "comments": "Professional GST reconciliation and exception reporting workbook.",
        })
        header_format = workbook.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": f"#{brand}",
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        review_format = workbook.add_format({"bg_color": "#FFF2CC"})
        title_format = workbook.add_format({"bold": True, "font_size": 18, "font_color": f"#{brand}"})
        subtitle_format = workbook.add_format({"italic": True, "font_color": "#667085"})
        brand_format = workbook.add_format({"bold": True, "font_color": f"#{brand}", "font_size": 9})
        banner_format = workbook.add_format({"bg_color": "#F9EEF2"})
        link_format = workbook.add_format({"font_color": "#0563C1", "underline": 1})

        for sheet_name, frame in sheets:
            if progress_callback:
                progress_callback(f"Writing {sheet_name} sheet...")
            start_row = 3 if sheet_name == "Summary" else 0
            frame.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_row)
            worksheet = writer.sheets[sheet_name]
            rows, columns = frame.shape
            header_row = start_row
            worksheet.hide_gridlines(2)
            worksheet.freeze_panes(header_row + 1, 0)
            worksheet.set_tab_color(f"#{brand}")
            worksheet.set_footer("PUSHPAK KUMAR  |  GST Reconciliation Studio")
            worksheet.set_row(header_row, 28)
            worksheet.set_column(0, max(columns - 1, 0), 12)

            for column_index, column_name in enumerate(frame.columns):
                longest = max([len(str(column_name)), *[len(str(value or "")) for value in frame[column_name].head(100)]])
                worksheet.set_column(column_index, column_index, min(max(longest + 2, 12), 30))
                worksheet.write(header_row, column_index, column_name, header_format)

            if "matched_2b_link" in frame.columns and "matched_2b_row_id" in frame.columns:
                link_column = frame.columns.get_loc("matched_2b_link")
                row_column = frame.columns.get_loc("matched_2b_row_id")
                for frame_row, row_id in enumerate(frame.iloc[:, row_column]):
                    if pd.notna(row_id) and str(frame.iloc[frame_row, link_column]).strip():
                        target_excel_row = int(row_id) + 2
                        worksheet.write_url(
                            header_row + 1 + frame_row,
                            link_column,
                            f"internal:'GSTR_2B_Upload'!A{target_excel_row}",
                            link_format,
                            string="Open GSTR-2B invoice",
                        )

            if columns:
                worksheet.autofilter(header_row, 0, header_row + rows, columns - 1)

            if "status" in frame.columns and rows:
                status_column = frame.columns.get_loc("status")
                worksheet.conditional_format(
                    header_row + 1, 0, header_row + rows, columns - 1,
                    {"type": "formula", "criteria": f'=${get_column_letter(status_column + 1)}{header_row + 2}<>"Matched"', "format": review_format},
                )

            if sheet_name == "Summary":
                worksheet.merge_range("B1:G1", "GST RECONCILIATION STUDIO", title_format)
                worksheet.merge_range("B2:G2", "GSTR-2B vs Inward Register | Reconciliation & Exception Report", subtitle_format)
                worksheet.merge_range("B3:G3", "Prepared by PUSHPAK KUMAR  |  GST Reconciliation & Reporting", brand_format)
                worksheet.set_row(0, 32)
                worksheet.set_row(1, 22)
                worksheet.set_row(2, 20)
                for row in range(3):
                    worksheet.set_row(row, None, banner_format)
                if logo_path.exists():
                    worksheet.insert_image("A1", str(logo_path), {"x_scale": 0.45, "y_scale": 0.45})

        if progress_callback:
            progress_callback("Excel workbook finalized.")

    return output.getvalue()
