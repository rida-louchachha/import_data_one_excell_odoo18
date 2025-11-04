# -*- coding: utf-8 -*-
from odoo import _
from . import constants as C


def parse_worksheet_strict(ws, target_key):
    """
    Strict parser that enforces exact headers (labels and order).
    Returns (hmap, body) where:
      hmap: {internal_key -> column_index_zero_based}
      body: list of data rows with trimmed strings (header removed)
    """
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([("" if v is None else str(v).strip()) for v in row])
    if not rows:
        return {}, []

    raw_header = [h if h is not None else "" for h in rows[0]]
    expected_pairs = C.HEADER_LABELS.get(target_key, [])
    expected_labels = [lbl for (_k, lbl) in expected_pairs]

    # Trim trailing empties
    while raw_header and (raw_header[-1] or "").strip() == "":
        raw_header.pop()

    # Column count must match exactly
    if len(raw_header) != len(expected_labels):
        raise ValueError(
            _("Sheet '%(sheet)s' header mismatch.\nExpected exactly %(exp_n)d columns: %(exp)s\n"
              "But found %(got_n)d columns: %(got)s\n\nHeaders cannot be renamed, moved, added or removed.") % {
                "sheet": ws.title,
                "exp_n": len(expected_labels),
                "exp": ", ".join(expected_labels),
                "got_n": len(raw_header),
                "got": ", ".join(raw_header) or "-",
            }
        )

    # Each label must match exactly
    for i, (got, want) in enumerate(zip(raw_header, expected_labels), start=1):
        if (got or "").strip() != want:
            raise ValueError(
                _("Sheet '%(sheet)s' header mismatch at column %(col)d.\n"
                  "Expected: '%(want)s'\nFound:    '%(got)s'\n\nHeaders cannot be renamed, moved, added or removed.") %
                {"sheet": ws.title, "col": i, "want": want, "got": got}
            )

    hmap = {key: idx for idx, (key, _lbl) in enumerate(expected_pairs)}
    body = rows[1:]
    return hmap, body


def build_master_template(env):
    """
    Returns (filename, bytes) of the master import Excel template.

    Guarantees:
    - Row 1 locked (headers), all other cells unlocked.
    - Examples are written on EVERY sheet (row 2).
    - BoM sheets: a merged, visual example block for the first product only.
    - BoM Type column is hidden + pre-filled:
        * mrp.bom.semi_finished → "normal"
        * mrp.bom.finished     → "phantom"
    - Validations apply to ALL rows (2..MAX_ROWS).
    - product.raw includes 3-state barcode mode + conditional manual entry.
    """
    import io
    output = io.BytesIO()

    try:
        import xlsxwriter
        use_xlsxwriter = True
    except Exception:
        use_xlsxwriter = False

    # Models
    Uom = env["uom.uom"]
    PP  = env["product.product"]
    PT  = env["product.template"]
    Cat = env["product.category"]
    Loc = env["stock.location"]

    def _dedup(seq):
        seen = set(); out = []
        for x in seq:
            if x not in seen:
                out.append(x); seen.add(x)
        return out

    # Source data
    uom_names = _dedup(Uom.search([]).mapped("name"))
    existing_categories = _dedup(Cat.search([]).mapped("name"))

    purchasable_variants  = _dedup(PP.search([("purchase_ok", "=", True)]).mapped("name"))
    pos_finished_variants = _dedup(PP.search([("product_tmpl_id.available_in_pos", "=", True)]).mapped("name"))

    # Components list for finished BoM = ALL products
    component_variants = _dedup(PP.search([]).mapped("name"))

    # Semi-finished templates (include true semi even if flags are False)
    if "available_in_pos" in PT._fields:
        semi_domain = ["|", ("product_type_cm", "=", "semi"),
                       "&", "&", ("purchase_ok", "=", False), ("sale_ok", "=", False), ("available_in_pos", "=", False)]
    else:
        semi_domain = ["|", ("product_type_cm", "=", "semi"),
                       "&", ("purchase_ok", "=", False), ("sale_ok", "=", False)]
    semi_finished_variants = _dedup(PT.search(semi_domain).mapped("name"))

    raw_variants = _dedup(PP.search([("product_tmpl_id.product_type_cm", "=", "raw")]).mapped("name"))

    # Default semi-finished category (for examples)
    semi_cat_rec = env.ref("import_wizards_caisse_manager_18.category_semi_finished_cm", raise_if_not_found=False)
    semi_cat_name = semi_cat_rec.name if semi_cat_rec else "Produits semi-finis"
    if semi_cat_name and semi_cat_name not in existing_categories:
        existing_categories = [semi_cat_name] + existing_categories

    # Locations
    parent_loc_names = _dedup(Loc.search([]).mapped("name"))
    if "WH" not in parent_loc_names:
        parent_loc_names = ["WH"] + parent_loc_names
    else:
        parent_loc_names = ["WH"] + [n for n in parent_loc_names if n != "WH"]

    MAX_ROWS = 10000  # validations/apply range

    if use_xlsxwriter:
        import xlsxwriter  # noqa: F401
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        # Formats
        fmt_header_locked = wb.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'text_wrap': True, 'locked': True,
        })
        fmt_unlocked = wb.add_format({'locked': False})
        fmt_unlocked_center = wb.add_format({'locked': False, 'align': 'center', 'valign': 'vcenter'})
        border_unlocked = wb.add_format({'border': 1, 'locked': False})
        colname = xlsxwriter.utility.xl_col_to_name

        # Helpers
        def _helper_fill_column(ws, col_idx_zero, values):
            """Write helper list in a hidden column for validations."""
            for r, v in enumerate(values, start=1):
                ws.write(r, col_idx_zero, v)
            ws.set_column(col_idx_zero, col_idx_zero, None, None, {'hidden': True})

        def _mirror_or_blank(sheet_name, col_letter, row_1b):
            """Mirror 'Sheet'!$A$n but yield "" if blank to avoid '0' in lists."""
            return f"=IF(TRIM('{sheet_name}'!${col_letter}${row_1b})=\"\",\"\",'{sheet_name}'!${col_letter}${row_1b})"

        def _protect_sheet_row1_readonly(ws, num_cols, default_col_width=28, allow_user_merge=False):
            """Row 1 locked (format), other cells unlocked.
            NOTE: Excel disables Merge on protected sheets. Set allow_user_merge=True
            to skip protection so users can merge any cells except you should warn them
            not to touch row 1.
            """
            ws.set_column(0, max(0, max(num_cols - 1, 0) + 20), default_col_width, fmt_unlocked)
            ws.set_row(0, None, fmt_header_locked)  # row 1 LOOKS locked

            if allow_user_merge:
                # No sheet protection → users can merge cells via UI.
                # (Lock flags aren’t enforced without protection, so rely on guidance/UI.)
                ws.freeze_panes(1, 0)
                # Optional: nudge users not to touch headers
                ws.write_comment(0, 0, "Ne modifiez/mergez pas la ligne d’en-tête.")
                return

            # Protected mode → users CANNOT merge via UI (Excel limitation)
            ws.protect('', {
                'select_locked_cells': False,  # row 1 (locked) not selectable
                'select_unlocked_cells': True,
                'format_cells': True,
                'format_columns': True,
                'format_rows': True,
                'insert_columns': True,
                'insert_rows': True,
                'delete_columns': True,
                'delete_rows': True,
                'sort': True,
                'autofilter': True,
                'objects': True,
                'scenarios': True,
            })
            ws.freeze_panes(1, 0)

        # Build each sheet
        def _build_sheet(sheet_key):
            title = C.SHEET_TITLES[sheet_key]
            headers = C.HEADER_LABELS.get(sheet_key, [])
            ws = wb.add_worksheet(title)

            # Headers (locked)
            for c, (_k, lbl) in enumerate(headers):
                ws.write(0, c, lbl, fmt_header_locked)

            # Protect: row 1 locked, others unlocked
            _protect_sheet_row1_readonly(ws, num_cols=len(headers), allow_user_merge=True)

            next_hcol = len(headers)
            u0 = uom_names[0] if uom_names else ""

            # ------------------ PRODUCTS (MP & semi-finis & Finished) ------------------
            if sheet_key in ("product.raw", "product.semi_finished", "product.finished"):
                # Example row on every sheet (row 2)
                if sheet_key == "product.semi_finished":
                    cat_default = semi_cat_name or (existing_categories[0] if existing_categories else "")
                else:
                    cat_default = existing_categories[0] if existing_categories else ""

                if sheet_key == "product.raw":
                    # name, uom, price, cat, image, barcode_mode, barcode_manual
                    ws.write_row(1, 0, ["Produit A", u0, "199.99", cat_default, "", "Auto", ""], fmt_unlocked)
                else:
                    # name, uom, price, cat, image
                    ws.write_row(1, 0, ["Produit A", u0, "199.99", cat_default, ""], fmt_unlocked)

                if sheet_key == "product.finished":
                    ws.write_row(1, 0, ["Produit Finished 1", u0, "199.99", cat_default, ""], fmt_unlocked)
                else:
                    ws.write_row(1, 0, ["Produit Finished 1", u0, "199.99", cat_default, ""], fmt_unlocked)

                # UoM dropdown (all rows, col=1)
                if uom_names:
                    _helper_fill_column(ws, next_hcol, uom_names)
                    ws.data_validation(1, 1, MAX_ROWS, 1, {
                        'validate': 'list',
                        'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${len(uom_names)}",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })
                next_hcol += 1

                # Categories dropdown (all rows, col=3) + mirror from categories
                cat_sheet = C.SHEET_TITLES["product.category"]
                existing_n = len(existing_categories)
                _helper_fill_column(ws, next_hcol, existing_categories)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(cat_sheet, "A", j + 1))
                ws.data_validation(1, 3, MAX_ROWS, 3, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${existing_n + MAX_ROWS}",
                    'ignore_blank': True, 'error_type': 'stop',
                    'error_title': 'Catégorie invalide',
                    'error_message': "Choisissez une catégorie depuis la liste.",
                })
                next_hcol += 1

                # Image column
                img_col = 4
                ws.set_column(img_col, img_col, 36, fmt_unlocked)
                ws.write_comment(0, img_col, "Insérez une image (PNG/JPG) ou laissez vide.")

                # MP only: barcode mode + manual (apply to all rows)
                if sheet_key == "product.raw":
                    _helper_fill_column(ws, next_hcol, ["Auto", "Manuel", "Aucun"])
                    mode_choices_col = next_hcol
                    mode_col = 5   # "Code-barres (Auto / Manuel / Aucun)"
                    manual_col = 6 # "Code-barres manuel (si Manuel)"

                    ws.data_validation(1, mode_col, MAX_ROWS, mode_col, {
                        'validate': 'list',
                        'source': f"${colname(mode_choices_col)}$1:${colname(mode_choices_col)}$3",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Mode invalide',
                        'error_message': "Choisissez Auto / Manuel / Aucun.",
                    })
                    next_hcol += 1

                    ws.set_column(manual_col, manual_col, 24, fmt_unlocked)
                    # Per-row custom rule: manual allowed only if mode="Manuel"
                    for r in range(2, MAX_ROWS + 1):
                        ws.data_validation(r - 1, manual_col, r - 1, manual_col, {
                            'validate': 'custom',
                            'value': f'=${colname(mode_col)}{r}="Manuel"',
                            'ignore_blank': True, 'error_type': 'stop',
                            'error_title': 'Non autorisé',
                            'error_message': "Saisie autorisée uniquement si Mode = 'Manuel'.",
                        })

                # Draw borders but DO NOT overwrite example row (start at row 3)
                last_col = len(headers) - 1
                for rr in range(3, min(MAX_ROWS, 200) + 1):
                    for cc in range(0, last_col + 1):
                        ws.write_blank(rr - 1, cc, None, border_unlocked)
                return ws

            # ------------------ CATEGORIES ------------------
            if sheet_key == "product.category":
                # Example row (row 2)
                ws.write_row(1, 0, ["Catégorie Test A", ""], fmt_unlocked)

                existing_n = len(existing_categories)
                _helper_fill_column(ws, next_hcol, existing_categories)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(title, "A", j + 1))
                ws.data_validation(1, 1, MAX_ROWS, 1, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${existing_n + MAX_ROWS}",
                    'ignore_blank': True, 'error_type': 'stop',
                    'error_title': 'Parent invalide',
                    'error_message': "Choisissez une catégorie parente de la liste.",
                })

                # Borders (start at row 3)
                for rr in range(3, min(MAX_ROWS, 200) + 1):
                    for cc in range(0, len(headers)):
                        ws.write_blank(rr - 1, cc, None, border_unlocked)
                return ws

            # ------------------ SEMI-FINISHED BoM ------------------
            if sheet_key == "mrp.bom.semi_finished":
                COMPONENT_ROWS_DEFAULT = 10
                BLOCK_ROWS = 1 + COMPONENT_ROWS_DEFAULT  # header + 10 comps
                START_ROW = 1  # example block starts on row index 1 (Excel row 2)

                p0 = (semi_finished_variants[0] if semi_finished_variants else "Produit semi-fini A")
                u0 = (uom_names[0] if uom_names else "")
                comp0 = raw_variants[0] if raw_variants else "Matière Première A"
                example = [p0, "BOM-001", "1", u0, "normal", comp0, "2", u0]
                ws.write_row(START_ROW, 0, example, fmt_unlocked)

                # Merge ONLY the example header block across rows 2.. for A..D
                ws.merge_range(START_ROW, 0, START_ROW + BLOCK_ROWS - 1, 0, example[0], fmt_unlocked_center)
                ws.merge_range(START_ROW, 1, START_ROW + BLOCK_ROWS - 1, 1, example[1], fmt_unlocked_center)
                ws.merge_range(START_ROW, 2, START_ROW + BLOCK_ROWS - 1, 2, example[2], fmt_unlocked_center)
                ws.merge_range(START_ROW, 3, START_ROW + BLOCK_ROWS - 1, 3, example[3], fmt_unlocked_center)

                next_hcol = len(headers)

                # UoM validations ALL rows (bom_uom col=3, component_uom col=7)
                if uom_names:
                    _helper_fill_column(ws, next_hcol, uom_names)
                    ucol = next_hcol
                    ws.data_validation(1, 3, MAX_ROWS, 3, {
                        'validate': 'list',
                        'source': f"${colname(ucol)}$1:${colname(ucol)}${len(uom_names)}",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })
                    ws.data_validation(1, 7, MAX_ROWS, 7, {
                        'validate': 'list',
                        'source': f"${colname(ucol)}$1:${colname(ucol)}${len(uom_names)}",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })
                next_hcol += 1

                # Semi-finished product list: ALL + mirror (col=0)
                semi_sheet = C.SHEET_TITLES["product.semi_finished"]
                existing_n = len(semi_finished_variants)
                _helper_fill_column(ws, next_hcol, semi_finished_variants)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(semi_sheet, "A", j + 1))
                ws.data_validation(1, 0, MAX_ROWS, 0, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${max(1, existing_n + MAX_ROWS)}",
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Produit invalide',
                    'error_message': "Choisissez un produit semi-fini existant ou saisi dans l’onglet.",
                })
                next_hcol += 1

                # Components ALL rows (col=5) = purchasable + mirror of MP
                raw_sheet = C.SHEET_TITLES["product.raw"]
                existing_n = len(purchasable_variants)
                _helper_fill_column(ws, next_hcol, purchasable_variants)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(raw_sheet, "A", j + 1))
                ws.data_validation(1, 5, MAX_ROWS, 5, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${existing_n + MAX_ROWS}",
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Composant invalide',
                    'error_message': "Choisissez une MP existante ou saisie dans l’onglet MP.",
                })
                next_hcol += 1

                # BoM type (col=4) HIDDEN + default 'normal' on ALL rows
                ws.set_column(4, 4, None, fmt_unlocked, {'hidden': True})
                for r in range(2, MAX_ROWS + 1):
                    ws.write(r - 1, 4, "normal", fmt_unlocked)

                # (Optional) light defaults for visible slice
                for r in range(2, min(MAX_ROWS, 200) + 1):
                    if u0:
                        ws.write(r - 1, 3, u0, fmt_unlocked)  # bom_uom
                        ws.write(r - 1, 7, u0, fmt_unlocked)  # component_uom

                # Borders: start at row 3 (do not touch example block)
                for rr in range(3, min(MAX_ROWS, 200) + 1):
                    for cc in range(0, len(headers)):
                        ws.write_blank(rr - 1, cc, None, border_unlocked)
                return ws

            # ------------------ FINISHED BoM ------------------
            if sheet_key == "mrp.bom.finished":
                COMPONENT_ROWS_DEFAULT = 10
                BLOCK_ROWS = 1 + COMPONENT_ROWS_DEFAULT
                START_ROW = 1  # example block starts on row index 1 (Excel row 2)

                p0 = (pos_finished_variants[0] if pos_finished_variants else "Produit fini POS A")
                u0 = (uom_names[0] if uom_names else "")
                comp0 = (component_variants[0] if component_variants else "Composant A")
                # cols: 0 prod, 1 barcode_mode, 2 barcode_manual, 3 code, 4 qty, 5 uom, 6 type, 7 comp, 8 c_qty, 9 c_uom
                example = [p0, "Auto", "", "BOM-001", "1", u0, "phantom", comp0, "2", u0]
                ws.write_row(START_ROW, 0, example, fmt_unlocked)

                # Merge ONLY the example header block across rows 2.. for A..F (0..5)
                ws.merge_range(START_ROW, 0, START_ROW + BLOCK_ROWS - 1, 0, example[0], fmt_unlocked_center)
                ws.merge_range(START_ROW, 1, START_ROW + BLOCK_ROWS - 1, 1, example[1], fmt_unlocked_center)
                ws.merge_range(START_ROW, 2, START_ROW + BLOCK_ROWS - 1, 2, example[2], fmt_unlocked_center)
                ws.merge_range(START_ROW, 3, START_ROW + BLOCK_ROWS - 1, 3, example[3], fmt_unlocked_center)
                ws.merge_range(START_ROW, 4, START_ROW + BLOCK_ROWS - 1, 4, example[4], fmt_unlocked_center)
                ws.merge_range(START_ROW, 5, START_ROW + BLOCK_ROWS - 1, 5, example[5], fmt_unlocked_center)

                next_hcol = len(headers)

                # Mode dropdown on ALL rows (col=1)
                _helper_fill_column(ws, next_hcol, ["Auto", "Manuel", "Aucun"])
                ws.data_validation(1, 1, MAX_ROWS, 1, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}$3",
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Mode invalide',
                    'error_message': "Choisissez Auto / Manuel / Aucun.",
                })
                next_hcol += 1

                # Manual barcode allowed only if current-row mode = "Manuel" (col=2)
                ws.set_column(2, 2, 24, fmt_unlocked)
                for r in range(2, MAX_ROWS + 1):
                    ws.data_validation(r - 1, 2, r - 1, 2, {
                        'validate': 'custom',
                        'value': f'=${colname(1)}{r}="Manuel"',
                        'ignore_blank': True, 'error_type': 'stop',
                        'error_title': 'Non autorisé',
                        'error_message': "Saisie autorisée uniquement si Mode = 'Manuel'.",
                    })

                # UoM validations on ALL rows (cols 5 and 9)
                if uom_names:
                    _helper_fill_column(ws, next_hcol, uom_names)
                    ucol = next_hcol
                    ws.data_validation(1, 5, MAX_ROWS, 5, {
                        'validate': 'list',
                        'source': f"${colname(ucol)}$1:${colname(ucol)}${len(uom_names)}",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })
                    ws.data_validation(1, 9, MAX_ROWS, 9, {
                        'validate': 'list',
                        'source': f"${colname(ucol)}$1:${colname(ucol)}${len(uom_names)}",
                        'ignore_blank': False, 'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })
                next_hcol += 1

                # Finished product selector on ALL rows (col=0)
                finished_sheet = C.SHEET_TITLES["product.finished"]

                existing_n = len(pos_finished_variants)
                _helper_fill_column(ws, next_hcol, pos_finished_variants)

                # Mirror entries typed in the "Produits – Finis" tab (column A)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(finished_sheet, "A", j + 1))

                # Validation covers: DB finished products + MAX_ROWS newly typed finished products
                ws.data_validation(1, 0, MAX_ROWS, 0, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${max(1, existing_n + MAX_ROWS)}",
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Produit fini invalide',
                    'error_message': "Choisissez un produit fini existant ou saisi dans l’onglet Produits – Finis.",
                })
                next_hcol += 1
                # Components (col=7) = ALL DB products + mirror of RAW + mirror of SEMI-FINISHED
                raw_sheet  = C.SHEET_TITLES["product.raw"]
                semi_sheet = C.SHEET_TITLES["product.semi_finished"]

                existing_n = len(component_variants)
                _helper_fill_column(ws, next_hcol, component_variants)

                # mirror RAW entries (from MP tab)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(raw_sheet, "A", j + 1))

                # mirror SEMI-FINISHED entries (from Produits semi-finis tab)
                for j in range(1, MAX_ROWS + 1):
                    ws.write_formula(existing_n + MAX_ROWS + j, next_hcol, _mirror_or_blank(semi_sheet, "A", j + 1))

                # Validation covers: DB products + MAX_ROWS RAW + MAX_ROWS SEMI-FINISHED
                ws.data_validation(1, 7, MAX_ROWS, 7, {
                    'validate': 'list',
                    'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${max(1, existing_n + 2*MAX_ROWS)}",
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Composant invalide',
                    'error_message': "Choisissez un produit existant ou saisi dans les onglets MP / Produits semi-finis.",
                })
                next_hcol += 1

                # BoM type (col=6) HIDDEN + default 'phantom' on ALL rows
                ws.set_column(6, 6, None, fmt_unlocked, {'hidden': True})
                for r in range(2, MAX_ROWS + 1):
                    ws.write(r - 1, 6, "phantom", fmt_unlocked)

                # Borders (start at row 3; don't touch example block)
                for rr in range(3, min(MAX_ROWS, 200) + 1):
                    for cc in range(0, len(headers)):
                        ws.write_blank(rr - 1, cc, None, border_unlocked)
                return ws

            # ------------------ LOCATIONS ------------------
            if sheet_key == "stock.location":
                # Example row (row 2)
                ws.write_row(1, 0, ["Test", "WH", "internal"], fmt_unlocked)

                # Build Parent list = DB parents + mirror of names typed in column A of this same sheet
                existing_n = len(parent_loc_names)
                if parent_loc_names:
                    _helper_fill_column(ws, next_hcol, parent_loc_names)  # hidden helper column

                    # Mirror A2..A{MAX_ROWS+1} from THIS sheet ('title'), turning blanks into ""
                    for j in range(1, MAX_ROWS + 1):
                        ws.write_formula(existing_n + j, next_hcol, _mirror_or_blank(title, "A", j + 1))

                    # Validate column B (Parent) against the combined helper range
                    ws.data_validation(1, 1, MAX_ROWS, 1, {
                        'validate': 'list',
                        'source': f"${colname(next_hcol)}$1:${colname(next_hcol)}${max(1, existing_n + MAX_ROWS)}",
                        'ignore_blank': True, 'error_type': 'stop',
                        'error_title': 'Parent invalide',
                        'error_message': 'Choisissez un parent de la liste.',
                    })
                    next_hcol += 1

                # Usage dropdown (unchanged)
                ws.data_validation(1, 2, MAX_ROWS, 2, {
                    'validate': 'list',
                    'source': C.USAGE_CHOICES,
                    'ignore_blank': False, 'error_type': 'stop',
                    'error_title': 'Type invalide',
                    'error_message': 'Choisissez un type de la liste.',
                })

                # Borders (start at row 3)
                for rr in range(3, min(MAX_ROWS, 200) + 1):
                    for cc in range(0, len(headers)):
                        ws.write_blank(rr - 1, cc, None, border_unlocked)
                return ws
        # Build all sheets
        for key in C.SHEET_IMPORT_ORDER:
            _build_sheet(key)
        wb.close()

    else:
        # openpyxl fallback (locks row 1; minimal but with examples + hidden BoM type)
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Protection

        wb = Workbook()
        first = True

        def _lock_header_unlock_data(ws, num_cols, max_rows=MAX_ROWS):
            ws.protection.sheet = True
            ws.freeze_panes = "A2"
            unprot = Protection(locked=False)
            for r in range(2, max_rows + 1):
                for c in range(1, num_cols + 1):
                    ws.cell(row=r, column=c).protection = unprot

        for key in C.SHEET_IMPORT_ORDER:
            if first:
                ws = wb.active; ws.title = C.SHEET_TITLES[key]; first = False
            else:
                ws = wb.create_sheet(C.SHEET_TITLES[key])

            headers = C.HEADER_LABELS.get(key, [])
            num_cols = len(headers)

            # headers
            for c, (_k, lbl) in enumerate(headers, start=1):
                ws.cell(row=1, column=c, value=lbl)
                ws.column_dimensions[get_column_letter(c)].width = 28

            _lock_header_unlock_data(ws, num_cols)

            # Examples on EVERY sheet (row 2)
            if key == "product.category":
                ex = ["Catégorie A", ""]
            elif key == "product.raw":
                u0 = (uom_names[0] if uom_names else "")
                c_default = existing_categories[0] if existing_categories else ""
                ex = ["Produit A", u0, "199.99", c_default, "", "Auto", ""]
            elif key == "product.semi_finished":
                u0 = (uom_names[0] if uom_names else "")
                c_default = semi_cat_name or (existing_categories[0] if existing_categories else "")
                ex = ["Produit A", u0, "199.99", c_default, ""]
            elif key == "mrp.bom.semi_finished":
                u0 = (uom_names[0] if uom_names else "")
                p0 = semi_finished_variants[0] if semi_finished_variants else "Produit Semi-fini A"
                comp0 = raw_variants[0] if raw_variants else "Matière Première A"
                ex = [p0, "BOM-001", "1", u0, "normal", comp0, "2", u0]
            elif key == "product.finished":
                u0 = (uom_names[0] if uom_names else "")
                c_default = existing_categories[0] if existing_categories else ""
                ex = ["Produit Finished 1", u0, "199.99", c_default, ""]

            elif key == "mrp.bom.finished":
                u0 = (uom_names[0] if uom_names else "")
                comp0 = (component_variants[0] if component_variants else "Composant A")
                ex = ["Produit Fini A", "Auto", "", "BOMF-001", "1", u0, "phantom", comp0, "2", u0]
            else:
                ex = ["Test", "WH", "internal"]

            for c, v in enumerate(ex, start=1):
                cell = ws.cell(row=2, column=c, value=v)
                cell.protection = Protection(locked=False)

            if key == "mrp.bom.semi_finished":
                col_letter = get_column_letter(5)
                ws.column_dimensions[col_letter].hidden = True
                for r in range(2, MAX_ROWS + 1):
                    ws.cell(row=r, column=5, value="normal").protection = Protection(locked=False)

            if key == "mrp.bom.finished":
                # type column is 7th: index 7 => zero-based 6
                col_letter = get_column_letter(7)
                ws.column_dimensions[col_letter].hidden = True
                for r in range(2, MAX_ROWS + 1):
                    ws.cell(row=r, column=7, value="phantom").protection = Protection(locked=False)

        wb.save(output)

    content = output.getvalue()
    output.close()
    return "Modele_Master_Import.xlsx", content
