# -*- coding: utf-8 -*-
import base64
import os
import re
import zipfile
from io import BytesIO

from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError
from odoo.modules.module import get_module_path  # to save/read inside the module repo

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

# --- limits / guards ---
MAX_FETCH_BYTES = 10_000_000  # 10 MB hard limit when reading an image from disk
PREVIEW_MAX = 24
THUMB_MAX_SIDE = 160
THUMB_TARGET_BYTES = 120_000

# services
from ..services import constants as C
from ..services import xlsx_builder as XB
from ..services import collectors as COL
from ..services import importers as IMP
from ..services import resolvers as R


class CmMasterImportWizard(models.TransientModel):
    _name = "cm.master.import.wizard"
    _description = "Master Import Wizard (All-in-One XLSX)"

    # ===== File + results =====
    file = fields.Binary("Excel (.xlsx)")
    filename = fields.Char()
    log_text = fields.Text(readonly=True)
    line_count = fields.Integer(readonly=True)
    error_count = fields.Integer(readonly=True)
    tested_ok = fields.Boolean(readonly=True)
    # kept for view compatibility; we don't fill it anymore
    log_preview_html = fields.Html(readonly=True, sanitize=False, translate=False)

    # ===== Image repo flow (ZIP → static/import_images) =====
    has_local_images = fields.Boolean(
        string="Excel contains image filenames?",
        help="If checked, upload the ZIP of images and click “Import Image Repo” "
             "before running Test/Import."
    )
    image_zip = fields.Binary("Images (ZIP)")
    image_zip_filename = fields.Char()
    image_repo_report = fields.Html(readonly=True, sanitize=False)
    images_repo_ok = fields.Boolean(readonly=True, help="True when images were extracted and match Excel needs.")

    # ------------------------------------------------------------
    # Basics
    # ------------------------------------------------------------
    @api.model
    def action_open_master_wizard(self):
        wiz = self.create({})
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": wiz.id,
            "view_mode": "form",
            "target": "new",
        }

    def _require_openpyxl(self):
        if load_workbook is None:
            raise UserError(_("Missing Python dependency: openpyxl"))

    def _read_workbook(self, *, read_only=True):
        self._require_openpyxl()
        if not self.file:
            raise UserError(_("Please upload an .xlsx file."))
        try:
            stream = BytesIO(base64.b64decode(self.file))
            return load_workbook(stream, read_only=read_only, data_only=True)
        except Exception as e:
            raise UserError(_("Could not read file: %s") % e)

    # ------------------------------------------------------------
    # Minimal image rules: ONLY a bare filename in Excel (e.g. a.png)
    # ------------------------------------------------------------
    @staticmethod
    def _is_bare_filename(s: str) -> bool:
        """
        Accept only simple filenames with image extension.
        Examples: background.png, image.jpg, photo.jpeg, logo.webp
        """
        if not s:
            return False
        t = s.strip()
        # No slashes/drive chars; a valid image extension
        return bool(re.match(r'^[^\\/:*?"<>|]+\.(png|jpe?g|webp)$', t, re.I))

    # ------------------------------------------------------------
    # Image repository helpers
    # ------------------------------------------------------------
    def _repo_base_dir(self):
        """<module>/static/import_images"""
        mod = getattr(self, "_module", None) or self._name.split(".")[0]
        base = get_module_path(mod)
        return os.path.join(base, "static", "import_images")

    def _ensure_dir(self, d):
        os.makedirs(d, exist_ok=True)
        if not os.path.isdir(d):
            raise UserError(_("Cannot create directory: %s") % d)

    def _safe_write(self, dest_dir, rel_name, raw_bytes):
        base = os.path.basename(rel_name)
        dest = os.path.normpath(os.path.join(dest_dir, base))
        if not dest.startswith(os.path.normpath(dest_dir)):
            raise UserError(_("Blocked unsafe path: %s") % rel_name)
        with open(dest, "wb") as f:
            f.write(raw_bytes)
        return dest

    def _collect_required_filenames(self, wb):
        """
        Scan the product sheets and return the set of image filenames required,
        using ONLY the typed text of the Image column.
        """
        required = set()
        for sheet_key in ("product.raw", "product.semi_finished"):
            title = C.SHEET_TITLES.get(sheet_key, sheet_key)
            if title not in wb.sheetnames:
                continue
            ws = wb[title]
            try:
                hmap, body = XB.parse_worksheet_strict(ws, sheet_key)
            except Exception:
                continue
            img_col_idx = (hmap.get("image") or hmap.get("Image") or hmap.get("image_embed"))
            if img_col_idx is None:
                continue

            for row in body:
                val = (row[img_col_idx] or "").strip() if img_col_idx < len(row) else ""
                if not val:
                    continue
                if self._is_bare_filename(val):
                    required.add(val)
        return required

    def _repo_image_bytes(self, filename: str) -> bytes:
        if not filename:
            return b""
        path = os.path.join(self._repo_base_dir(), filename)
        if not os.path.exists(path):
            return b""
        try:
            size = os.path.getsize(path)
            if 0 < size <= MAX_FETCH_BYTES:
                with open(path, "rb") as f:
                    return f.read()
        except Exception:
            return b""
        return b""

    def _make_thumb_png(self, raw: bytes) -> bytes:
        """Return a small PNG for preview; falls back to raw if Pillow missing."""
        if not raw:
            return b""
        try:
            from PIL import Image
            im = Image.open(BytesIO(raw));
            im.load()
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA")
            im.thumbnail((THUMB_MAX_SIDE, THUMB_MAX_SIDE))
            out = BytesIO()
            im.save(out, format="PNG", optimize=True)
            if out.tell() > THUMB_TARGET_BYTES:
                out.seek(0);
                out.truncate(0)
                im.save(out, format="PNG", optimize=True, compress_level=9)
            return out.getvalue()
        except Exception:
            return raw  # browser will scale it

    def _preview_tile_html(self, filename: str, raw: bytes) -> str:
        thumb = self._make_thumb_png(raw) or raw
        b64 = base64.b64encode(thumb).decode()
        safe = filename.replace("'", "&#39;").replace('"', "&quot;")
        return (
            "<div style='margin:6px'>"
            f"<img src='data:image/png;base64,{b64}' alt='{safe}' "
            "style='height:88px;max-width:180px;object-fit:contain;"
            "border:1px solid #ddd;padding:2px;border-radius:8px;background:#fff'/>"
            f"<div style='font-size:11px;color:#666;text-align:center;margin-top:4px'>{safe}</div>"
            "</div>"
        )

    # ------------------------------------------------------------
    # Button: import ZIP → static/import_images + report
    # ------------------------------------------------------------
    def action_import_image_repo(self):
        self.ensure_one()
        if not self.image_zip:
            raise UserError(_("Please upload a ZIP file containing the images."))

        dest_dir = self._repo_base_dir()
        self._ensure_dir(dest_dir)

        raw_zip = base64.b64decode(self.image_zip)
        imported, skipped = [], []
        try:
            with zipfile.ZipFile(BytesIO(raw_zip)) as zf:
                total_uncompressed = sum(zi.file_size for zi in zf.infolist())
                if total_uncompressed > 500 * 1024 * 1024:
                    raise UserError(_("ZIP too large (>%s MB).") % 500)
                for zi in zf.infolist():
                    if zi.is_dir():
                        continue
                    if not re.search(r"\.(png|jpe?g|webp)$", zi.filename, re.I):
                        skipped.append(zi.filename)
                        continue
                    data = zf.read(zi)
                    path = self._safe_write(dest_dir, zi.filename, data)
                    imported.append(os.path.basename(path))
        except zipfile.BadZipFile:
            raise UserError(_("Not a valid ZIP file."))
        except Exception as e:
            raise UserError(_("Could not extract ZIP: %s") % e)

        # If an Excel file is already uploaded, compute what's required
        required = set()
        try:
            wb = self._read_workbook(read_only=True)
            required = self._collect_required_filenames(wb)
        except Exception:
            # No Excel yet or unreadable; ignore
            pass

        imported_set = set(imported)
        missing = sorted(list(required - imported_set))

        self.images_repo_ok = not bool(missing)
        lines = []
        lines.append("<b>Images directory:</b> %s" % dest_dir)
        lines.append("<br/><b>Imported files:</b> %s" % len(imported))
        if skipped:
            lines.append("<br/><b>Skipped (non-images):</b> %s" % len(skipped))
        if required:
            lines.append("<br/><b>Excel requires:</b> %s" % len(required))
        if missing:
            lines.append("<br/><span style='color:#b00'>Missing:</span> %s" %
                         (", ".join(missing[:50]) + ("…" if len(missing) > 50 else "")))
        elif required:
            lines.append("<br/><span style='color:#080'>All required images are present.</span>")
        self.image_repo_report = "<div style='font-family:Inter,Segoe UI,Arial,sans-serif'>" + "".join(lines) + "</div>"

        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "new",
        }

    # ------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------
    def action_download_master_template(self):
        """
        Build & download the template; add a comment on the Image header:
        'Type only a filename like background.png / image.jpg / logo.webp'.
        """
        self.ensure_one()
        filename, content = XB.build_master_template(self.env)

        # Add a helper comment on Image column
        try:
            from openpyxl import load_workbook as _lb
            from openpyxl.comments import Comment
            wb = _lb(BytesIO(content))
            for sheet_key in ("product.raw", "product.semi_finished"):
                title = C.SHEET_TITLES.get(sheet_key, sheet_key)
                if title in wb.sheetnames:
                    ws = wb[title]
                    for cell in ws[1]:
                        val = (cell.value or "")
                        val = val.strip().lower() if isinstance(val, str) else str(val or "").lower()
                        if val in ("image", "image_embed", "image url", "image path"):
                            try:
                                cell.comment = Comment(
                                    "Saisissez uniquement le nom du fichier (ex: background.png, image.jpg, logo.webp).\n"
                                    "Les images seront recherchées dans le répertoire du module: static/import_images/",
                                    "helper",
                                )
                            except Exception:
                                pass
                            break
            buf = BytesIO()
            wb.save(buf)
            content = buf.getvalue()
        except Exception:
            pass

        # Use a one-off attachment to let the browser download (standard Odoo pattern)
        att = self.env["ir.attachment"].create({
            "name": filename,
            "type": "binary",
            "datas": base64.b64encode(content),
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "res_model": self._name,
            "res_id": self.id,
        })
        return {"type": "ir.actions.act_url", "url": f"/web/content/{att.id}?download=true", "target": "self"}

    def action_test_master(self):
        """Validate the Excel and check only filename-style images exist in static/import_images/."""
        self.ensure_one()
        wb = self._read_workbook()

        if self.has_local_images and not self.images_repo_ok:
            # Not blocking, but helpful: you can make it a UserError if you prefer hard block
            pass

        total_rows = 0
        total_errors = 0
        lines = []
        found_titles = []
        gallery = []
        bad_filename = 0
        missing_files = 0
        repo = self._repo_base_dir()

        for sheet_key in C.SHEET_IMPORT_ORDER:
            sheet_title = C.SHEET_TITLES.get(sheet_key, sheet_key)
            if sheet_title not in wb.sheetnames:
                continue
            found_titles.append(sheet_title)
            ws = wb[sheet_title]

            try:
                hmap, body = XB.parse_worksheet_strict(ws, sheet_key)
            except Exception as e:
                lines.append(f"[{sheet_title}]")
                lines.append(f"✖ {sheet_title}: header error — {e}")
                lines.append("")
                total_errors += 1
                continue

            # Check Image column (typed text only)
            if sheet_key in ("product.raw", "product.semi_finished"):
                img_col_idx = (hmap.get("image") or hmap.get("Image") or hmap.get("image_embed"))
                if img_col_idx is not None:
                    for i, row in enumerate(body, start=2):
                        if not any((cell or "").strip() for cell in row):
                            continue
                        val = (row[img_col_idx] or "").strip() if img_col_idx < len(row) else ""
                        if not val:
                            continue
                        if not self._is_bare_filename(val):
                            lines.append(
                                f"Line {i} [{sheet_title}] Image: invalid value '{val}'. "
                                f"Put only a filename like background.png / image.jpg / logo.webp."
                            )
                            total_errors += 1
                            bad_filename += 1
                            continue
                        p = os.path.join(repo, val)
                        if not os.path.exists(p):
                            lines.append(
                                f"Line {i} [{sheet_title}] Image file '{val}' not found in {repo}."
                            )
                            total_errors += 1
                            missing_files += 1
                        else:
                            if len(gallery) < PREVIEW_MAX:
                                raw = self._repo_image_bytes(val)
                                if raw:
                                    gallery.append(self._preview_tile_html(val, raw))

            # Count data rows + normal collector validation
            if not body or not any(any(c for c in r) for r in body):
                lines.append(f"✔ {sheet_title}: no data rows — OK")
                lines.append("")
                continue

            records, log, count, errors = COL.collect(self.env, hmap, body, sheet_key)
            total_rows += count
            total_errors += errors

            if errors == 0:
                lines.append(f"✔ {sheet_title}: OK — {count} row(s) checked")
            else:
                lines.append(f"[{sheet_title}]")
                lines.extend(log)
                lines.append(f"✖ {sheet_title}: {errors} issue(s) in {count} row(s)")
            lines.append("")

        header = [
            _("=== MASTER TEST RESULT ==="),
            _("Sheets found: %s") % (", ".join(found_titles) if found_titles else "-"),
            _("Total rows checked: %s") % total_rows,
            _("Image: invalid filenames (errors): %s") % bad_filename,
            _("Image: missing files in repo (errors): %s") % missing_files,
            _("Total issues: %s") % total_errors,
            "",
        ]

        self.write({
            "log_text": "\n".join(header + lines),
            "log_preview_html": (
                    "<div style='font-family:Inter,Segoe UI,Arial,sans-serif'>"
                    f"<div style='margin:6px 0 8px;color:#666'>Image previews (inline, {len(gallery)} shown)</div>"
                    "<div style='display:flex;flex-wrap:wrap'>" + (
                                "".join(gallery) or "<em>No previews</em>") + "</div>"
                                                                              "</div>"
            ),
            "line_count": total_rows,
            "error_count": total_errors,
            "tested_ok": total_errors == 0 and bool(found_titles),
        })
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "new",
        }

    def action_import_master(self):
        """
        Import everything; for product sheets, read image filename from the Image column
        and load bytes from static/import_images/<filename>.
        """
        self.ensure_one()
        wb = self._read_workbook()

        if self.has_local_images and not self.images_repo_ok:
            raise ValidationError(_("Please import the images ZIP (Import Image Repo) before importing."))

        created_total = updated_total = scanned_total = 0
        details_all = []
        processed_titles = []

        semi_sheet_names = None
        raw_sheet_names = None

        repo = self._repo_base_dir()

        for sheet_key in C.SHEET_IMPORT_ORDER:
            sheet_title = C.SHEET_TITLES.get(sheet_key, sheet_key)
            if sheet_title not in wb.sheetnames:
                continue

            processed_titles.append(sheet_title)
            ws = wb[sheet_title]
            hmap, body = XB.parse_worksheet_strict(ws, sheet_key)
            if not body or not any(any(c for c in r) for r in body):
                details_all.append(f"[{sheet_title}] (empty) — skipped")
                continue

            records, log, count, errors = COL.collect(self.env, hmap, body, sheet_key)
            scanned_total += count
            if errors:
                raise ValidationError(_("Sheet '%s': please fix errors first. Issues found: %s\n\n%s") %
                                      (sheet_title, errors, "\n".join(log[:50])))

            if sheet_key == "product.category":
                c, u, det = IMP.import_categories(self.env, records)

            elif sheet_key in ("product.raw", "product.semi_finished", "product.finished"):
                if sheet_key == "product.semi_finished":
                    semi_sheet_names = {R.norm(r.get("name")) for r in records if r.get("name")}
                if sheet_key == "product.raw":
                    raw_sheet_names = {R.norm(r.get("name")) for r in records if r.get("name")}

                c, u, det = IMP.import_products(self.env, records, sheet_key)

                # --- image setting from filename only ---
                if sheet_key in ("product.raw", "product.semi_finished"):
                    img_col_idx = (hmap.get("image") or hmap.get("Image") or hmap.get("image_embed"))
                    name_col_idx = hmap.get("name")
                    if img_col_idx is not None and name_col_idx is not None:
                        PT = self.env["product.template"]

                        for excel_row_idx, row in enumerate(body, start=2):
                            if not any((cell or "").strip() for cell in row):
                                continue

                            name = (row[name_col_idx] or "").strip() if name_col_idx < len(row) else ""
                            if not name:
                                continue

                            val = (row[img_col_idx] or "").strip() if img_col_idx < len(row) else ""
                            if not val:
                                continue
                            if not self._is_bare_filename(val):
                                details_all.append(f"[{sheet_title}] row {excel_row_idx}: image ignored "
                                                   f"(invalid filename '{val}')")
                                continue

                            path = os.path.join(repo, val)
                            if not os.path.exists(path):
                                details_all.append(f"[{sheet_title}] row {excel_row_idx}: image file '{val}' "
                                                   f"not found in repo {repo}")
                                continue
                            try:
                                size = os.path.getsize(path)
                                if not (0 < size <= MAX_FETCH_BYTES):
                                    details_all.append(f"[{sheet_title}] row {excel_row_idx}: image '{val}' too big")
                                    continue
                                with open(path, "rb") as f:
                                    img_bytes = f.read()
                            except Exception as e:
                                details_all.append(f"[{sheet_title}] row {excel_row_idx}: cannot read '{val}' ({e})")
                                continue

                            tmpl = (PT.search([("name", "=", name)], limit=1)
                                    or PT.search([("name", "ilike", name)], limit=1))
                            if not tmpl:
                                continue

                            tmpl.write({"image_1920": base64.b64encode(img_bytes)})
                            det = det or []
                            det.append(f"product.template({tmpl.id}) image set (row {excel_row_idx})")

            elif sheet_key == "mrp.bom.semi_finished":
                c, u, det = IMP.import_boms_semi_finished(self.env, records, semi_sheet_names, raw_sheet_names)

            elif sheet_key == "mrp.bom.finished":
                c, u, det = IMP.import_boms_finished(self.env, records, raw_sheet_names)

            elif sheet_key == "stock.location":
                c, u, det = IMP.import_locations(self.env, records)

            else:
                details_all.append(f"[{sheet_title}] unsupported sheet key '{sheet_key}' — skipped")
                continue

            created_total += c
            updated_total += u
            details_all.append(f"[{sheet_title}] OK — rows: {count}, created: {c}, updated: {u}")
            details_all.extend([f"[{sheet_title}] {d}" for d in (det[:200] if det else [])])

        summary = [
            _("=== MASTER IMPORT RESULT ==="),
            _("Sheets processed (order): %s") % (", ".join(processed_titles) or "-"),
            _("Rows scanned: %s") % scanned_total,
            _("Created: %s") % created_total,
            _("Updated: %s") % updated_total,
            "",
        ]
        self.write({
            "log_text": "\n".join(summary + details_all),
            "tested_ok": True,
            "line_count": scanned_total,
            "error_count": 0,
        })
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "new",
        }
