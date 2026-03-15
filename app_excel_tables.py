"""
app_excel_tables.py  — v4
─────────────────────────
Détection hybride de tableaux Excel avec 5 stratégies :

  1. Header Score souple (bold, fill, bordures, types de données)
  2. Tolérance aux lignes vides (lookahead 2 lignes)
  3. Détection par grille de bordures
  4. Profilage de types de données (cohérence verticale)
  5. Matrice de cellules pré-calculée (performance)

Lancement :
    pip install streamlit openpyxl pandas
    streamlit run app_excel_tables.py
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  CellInfo : structure légère pré-calculée
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass(slots=True)
class CellInfo:
    value: object = None
    bold: bool = False
    fill_key: Optional[str] = None
    font_size: float = 11.0
    border_top: bool = False
    border_bottom: bool = False
    border_left: bool = False
    border_right: bool = False
    has_any_border: bool = False
    data_type: str = "empty"  # empty | text | number | date | formula

    @property
    def has_fill(self) -> bool:
        return self.fill_key is not None

    @property
    def all_borders(self) -> bool:
        return (self.border_top and self.border_bottom
                and self.border_left and self.border_right)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  DetectedTable
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass
class DetectedTable:
    sheet: str
    title: str
    top_left: str
    bottom_right: str
    num_rows: int
    num_cols: int
    headers: list[str] = field(default_factory=list)
    source: str = ""
    score: float = 0.0
    has_header_fill: bool = False
    has_total_row: bool = False
    has_grid_borders: bool = False
    expanded_left: bool = False
    expanded_right: bool = False
    has_empty_rows: bool = False

    @property
    def range_str(self):
        return f"{self.top_left}:{self.bottom_right}"

    @property
    def badge(self):
        return {
            "excel_table":      "📊 Table Excel",
            "hybrid_detected":  "🎯 Hybride",
            "grid_detected":    "🔲 Grille bordures",
            "contiguous_block": "📦 Bloc contigu",
        }.get(self.source, self.source)

    @property
    def confidence(self):
        if self.score >= 70:
            return "haute"
        if self.score >= 40:
            return "moyenne"
        return "basse"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  ExcelTableDetector v4
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

LOOKAHEAD = 2  # lignes vides tolérées

class ExcelTableDetector:

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        wb_light = load_workbook(filepath, read_only=True, data_only=True)
        self.sheetnames = wb_light.sheetnames
        wb_light.close()
        self._wb = None
        self._wb_values = None

    def _ensure_wb(self):
        if self._wb is None:
            self._wb = load_workbook(str(self.filepath), data_only=False)
        if self._wb_values is None:
            self._wb_values = load_workbook(str(self.filepath), data_only=True)

    # ────────────────────────────────────────────
    #  Stratégie 5 : Matrice pré-calculée
    # ────────────────────────────────────────────
    def _build_matrix(self, ws) -> dict[tuple[int, int], CellInfo]:
        matrix: dict[tuple[int, int], CellInfo] = {}
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1

        for row in ws.iter_rows(min_row=1, max_row=max_r,
                                min_col=1, max_col=max_c):
            for cell in row:
                r, c = cell.row, cell.column

                # Fill key
                fk = None
                f = cell.fill
                if f and f.patternType and f.patternType != "none" and f.start_color:
                    sc = f.start_color
                    if sc.type == "rgb" and sc.rgb:
                        fk = f"rgb:{sc.rgb}"
                    elif sc.type == "theme":
                        fk = f"theme:{sc.theme}:{round(sc.tint or 0, 4)}"
                    elif sc.type == "indexed":
                        fk = f"idx:{sc.value}"

                # Borders (with None safety)
                b = cell.border
                bt = b.top.style is not None if b and b.top else False
                bb = b.bottom.style is not None if b and b.bottom else False
                bl = b.left.style is not None if b and b.left else False
                br_ = b.right.style is not None if b and b.right else False

                # Data type
                v = cell.value
                if v is None:
                    dt = "empty"
                elif isinstance(v, str) and v.startswith("="):
                    dt = "formula"
                elif isinstance(v, (int, float)):
                    dt = "number"
                elif isinstance(v, (datetime, date)):
                    dt = "date"
                elif isinstance(v, str):
                    dt = "text"
                else:
                    dt = "text"

                info = CellInfo(
                    value=v,
                    bold=bool(cell.font and cell.font.bold),
                    fill_key=fk,
                    font_size=float(cell.font.size) if cell.font and cell.font.size else 11.0,
                    border_top=bt, border_bottom=bb,
                    border_left=bl, border_right=br_,
                    has_any_border=(bt or bb or bl or br_),
                    data_type=dt,
                )
                # N'ajouter que si la cellule n'est pas totalement vide
                if v is not None or fk is not None or info.has_any_border:
                    matrix[(r, c)] = info

        # Marquer les cellules fusionnées
        for m in ws.merged_cells.ranges:
            top_cell = matrix.get((m.min_row, m.min_col))
            for mr in range(m.min_row, m.max_row + 1):
                for mc in range(m.min_col, m.max_col + 1):
                    if (mr, mc) not in matrix and top_cell:
                        matrix[(mr, mc)] = CellInfo(
                            value=top_cell.value,
                            bold=top_cell.bold,
                            fill_key=top_cell.fill_key,
                            font_size=top_cell.font_size,
                            data_type=top_cell.data_type,
                        )

        return matrix

    def _get(self, matrix, r, c) -> CellInfo:
        return matrix.get((r, c), CellInfo())

    # ────────────────────────────────────────────
    #  API publique
    # ────────────────────────────────────────────
    def detect_sheet(self, sheet_name: str) -> list[DetectedTable]:
        self._ensure_wb()
        ws = self._wb[sheet_name]
        matrix = self._build_matrix(ws)
        results: list[DetectedTable] = []
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1

        content_cells = {pos for pos, ci in matrix.items() if ci.value is not None}

        # Passe 1 : Tables Excel déclarées
        covered = self._declared_tables(ws, sheet_name, results, matrix)

        # Passe 2 : Détection hybride (header score + expansion + lookahead)
        covered |= self._hybrid_tables(matrix, sheet_name, results, covered,
                                        max_r, max_c, ws)

        # Passe 3 : Détection par grille de bordures
        covered |= self._grid_tables(matrix, sheet_name, results, covered,
                                      max_r, max_c, ws)

        # Passe 4 : Blocs contigus résiduels (gap=0 strict)
        self._residual_blocks(matrix, sheet_name, results, covered,
                              content_cells, max_r, max_c, ws)

        return results

    def load_table(self, table: DetectedTable) -> pd.DataFrame:
        self._ensure_wb()
        ws_val = self._wb_values[table.sheet]
        mc, mr, xc, xr = range_boundaries(f"{table.top_left}:{table.bottom_right}")
        data = []
        for row in ws_val.iter_rows(min_row=mr, max_row=xr, min_col=mc, max_col=xc):
            data.append([cell.value for cell in row])
        if not data:
            return pd.DataFrame()
        raw = data[0]
        headers, seen = [], {}
        for h in raw:
            name = str(h).strip() if h is not None else "Sans titre"
            name = name.replace("\n", " ")
            if name in seen:
                seen[name] += 1
                name = f"{name}_{seen[name]}"
            else:
                seen[name] = 0
            headers.append(name)
        return pd.DataFrame(data[1:], columns=headers)

    # ────────────────────────────────────────────
    #  Passe 1 : Tables Excel déclarées
    # ────────────────────────────────────────────
    def _declared_tables(self, ws, sheet, results, matrix) -> set:
        covered = set()
        for tbl in ws.tables.values():
            min_c, min_r, max_c, max_r = range_boundaries(tbl.ref)
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    covered.add((r, c))

            has_total = False
            tr = max_r + 1
            ci_t = self._get(matrix, tr, min_c)
            if ci_t.value is not None and ci_t.bold and ci_t.border_top:
                has_total = True
                max_r = tr
                for c in range(min_c, max_c + 1):
                    covered.add((tr, c))

            title = self._find_section_title(matrix, min_r, min_c, max_c, ws)
            if not title:
                title = tbl.displayName

            headers = [self._get(matrix, min_r, c).value or ""
                       for c in range(min_c, max_c + 1)]
            hdr_fill = self._get(matrix, min_r, min_c).fill_key

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(min_c)}{min_r}",
                bottom_right=f"{get_column_letter(max_c)}{max_r}",
                num_rows=max_r - min_r, num_cols=max_c - min_c + 1,
                headers=[str(h) for h in headers], source="excel_table", score=100,
                has_header_fill=hdr_fill is not None, has_total_row=has_total,
            ))
        return covered

    # ────────────────────────────────────────────
    #  Stratégie 1 : Header Score
    # ────────────────────────────────────────────
    def _header_score(self, matrix, row, c1, c2, max_r) -> float:
        """Évalue la probabilité qu'une ligne soit un en-tête de tableau."""
        score = 0.0
        cells = [self._get(matrix, row, c) for c in range(c1, c2 + 1)]
        non_empty = [ci for ci in cells if ci.value is not None]

        if not non_empty:
            return 0

        # Critère 1 : Cellules en gras (+3)
        bold_ratio = sum(1 for ci in non_empty if ci.bold) / len(non_empty)
        if bold_ratio >= 0.8:
            score += 3
        elif bold_ratio >= 0.5:
            score += 1.5

        # Critère 2 : Fill différent de la ligne suivante (+3)
        fills_hdr = set(ci.fill_key for ci in non_empty if ci.fill_key)
        if fills_hdr:
            next_r = row + 1
            fills_next = set(self._get(matrix, next_r, c).fill_key
                            for c in range(c1, c2 + 1)
                            if self._get(matrix, next_r, c).fill_key)
            if fills_hdr != fills_next:
                score += 3

        # Critère 3 : Bordure inférieure (+2)
        bottom_ratio = sum(1 for ci in non_empty if ci.border_bottom) / len(non_empty)
        if bottom_ratio >= 0.5:
            score += 2

        # Critère 4 : Type de données = texte, et ligne suivante = nombres/dates (+5)
        hdr_types = set(ci.data_type for ci in non_empty)
        if hdr_types <= {"text"}:
            if row + 1 <= max_r:
                next_cells = [self._get(matrix, row + 1, c)
                              for c in range(c1, c2 + 1)]
                next_non_empty = [ci for ci in next_cells if ci.value is not None]
                if next_non_empty:
                    next_types = set(ci.data_type for ci in next_non_empty)
                    if next_types & {"number", "date", "formula"}:
                        score += 5

        return score

    # ────────────────────────────────────────────
    #  Passe 2 : Détection hybride
    #    Header Score → expansion + lookahead
    # ────────────────────────────────────────────
    def _hybrid_tables(self, matrix, sheet, results, covered,
                       max_r, max_c, ws) -> set:
        new_covered = set()
        HEADER_THRESHOLD = 3.0

        # Trouver les header runs : séquences de cellules non-vides + même fill
        # OU séquences bold contigues, évaluées par header_score
        header_runs = []  # (row, c1, c2, fill_key_or_none)

        for r in range(1, max_r + 1):
            # Trouver des segments contigus de cellules non-vides
            segments = []
            seg_start = None
            for c in range(1, max_c + 2):
                ci = self._get(matrix, r, c) if c <= max_c else CellInfo()
                in_cov = (r, c) in covered or (r, c) in new_covered
                has_val = ci.value is not None and not in_cov and c <= max_c

                if has_val:
                    if seg_start is None:
                        seg_start = c
                else:
                    if seg_start is not None and c - seg_start >= 2:
                        segments.append((seg_start, c - 1))
                    seg_start = None

            # Évaluer chaque segment comme potentiel header
            for c1, c2 in segments:
                hs = self._header_score(matrix, r, c1, c2, max_r)
                if hs >= HEADER_THRESHOLD:
                    fk = self._get(matrix, r, c1).fill_key
                    header_runs.append((r, c1, c2, fk, hs))

        # Trier par score décroissant pour traiter les meilleurs en premier
        header_runs.sort(key=lambda x: -x[4])
        used_cells = set()

        for hdr_row, style_c1, style_c2, fk, hs in header_runs:
            if any((hdr_row, c) in used_cells or (hdr_row, c) in covered
                   for c in range(style_c1, style_c2 + 1)):
                continue

            # ── Expansion vers le bas avec LOOKAHEAD structurel ──
            last_data_row = hdr_row
            total_row = None
            empty_streak = 0
            has_empty_rows = False
            # Nombre de colonnes remplies typique (calculé sur les premières lignes)
            typical_fill = None

            r = hdr_row + 1
            while r <= max_r:
                filled = sum(1 for c in range(style_c1, style_c2 + 1)
                            if self._get(matrix, r, c).value is not None)

                if filled == 0:
                    empty_streak += 1
                    if empty_streak > LOOKAHEAD:
                        break
                    r += 1
                    continue

                # Après un gap vide, vérifier la cohérence structurelle
                if empty_streak > 0:
                    # Les données qui reprennent doivent remplir un nombre
                    # de colonnes similaire (±1) aux lignes précédentes
                    if typical_fill is not None and abs(filled - typical_fill) > 1:
                        break  # structure différente → c'est un autre tableau
                    has_empty_rows = True
                    empty_streak = 0

                # Mettre à jour le remplissage typique
                if typical_fill is None:
                    typical_fill = filled
                else:
                    # Moyenne glissante simple
                    typical_fill = round((typical_fill + filled) / 2)

                # Vérifier ligne de total
                row_cells = [self._get(matrix, r, c)
                             for c in range(style_c1, style_c2 + 1)]
                non_empty_rc = [ci for ci in row_cells if ci.value is not None]
                all_bold = non_empty_rc and all(ci.bold for ci in non_empty_rc)
                any_top = any(ci.border_top for ci in row_cells)

                if all_bold and any_top:
                    total_row = r
                    break

                last_data_row = r
                r += 1

            end_row = total_row if total_row else last_data_row
            if end_row - hdr_row < 1:
                continue

            # ── Expansion latérale par contenu ──
            final_c1, final_c2 = style_c1, style_c2
            expanded_left = expanded_right = False

            c = style_c1 - 1
            while c >= 1:
                col_filled = sum(
                    1 for r2 in range(hdr_row, end_row + 1)
                    if self._get(matrix, r2, c).value is not None
                    and (r2, c) not in covered
                )
                if col_filled >= max((end_row - hdr_row + 1) * 0.3, 1):
                    final_c1 = c
                    expanded_left = True
                    c -= 1
                else:
                    break

            c = style_c2 + 1
            while c <= max_c:
                col_filled = sum(
                    1 for r2 in range(hdr_row, end_row + 1)
                    if self._get(matrix, r2, c).value is not None
                    and (r2, c) not in covered
                )
                if col_filled >= max((end_row - hdr_row + 1) * 0.3, 1):
                    final_c2 = c
                    expanded_right = True
                    c += 1
                else:
                    break

            num_cols = final_c2 - final_c1 + 1
            num_rows = end_row - hdr_row

            # Vérifier chevauchement
            cells_set = set()
            overlap = 0
            total_ct = 0
            for r2 in range(hdr_row, end_row + 1):
                for c2 in range(final_c1, final_c2 + 1):
                    total_ct += 1
                    if (r2, c2) in covered or (r2, c2) in new_covered:
                        overlap += 1
                    else:
                        cells_set.add((r2, c2))
            if total_ct > 0 and overlap / total_ct > 0.3:
                continue

            # ── Scoring ──────────────────────
            score = min(hs * 5, 30)  # header score converti (max 30)
            score += 15 if num_cols >= 3 else 8 if num_cols >= 2 else 0
            score += 15 if num_rows >= 3 else 8 if num_rows >= 1 else 0
            score += 10 if total_row else 0

            # Cohérence de types verticale (Stratégie 4)
            type_bonus = self._type_consistency_score(matrix, hdr_row, end_row,
                                                      final_c1, final_c2)
            score += type_bonus

            # Densité
            body_end = (total_row - 1) if total_row else end_row
            body_cells = max((body_end - hdr_row) * num_cols, 1)
            filled_ct = sum(
                1 for r2 in range(hdr_row + 1, body_end + 1)
                for c2 in range(final_c1, final_c2 + 1)
                if self._get(matrix, r2, c2).value is not None
            )
            density = filled_ct / body_cells
            score += 10 if density >= 0.4 else 5 if density >= 0.2 else 0

            if expanded_left or expanded_right:
                score += 5

            score = min(score, 95)

            title = self._find_section_title(matrix, hdr_row, final_c1,
                                              final_c2, ws)
            if not title:
                title = f"Tableau_{get_column_letter(final_c1)}{hdr_row}"

            headers = []
            for c2 in range(final_c1, final_c2 + 1):
                v = self._get(matrix, hdr_row, c2).value
                headers.append(str(v) if v is not None else "")

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(final_c1)}{hdr_row}",
                bottom_right=f"{get_column_letter(final_c2)}{end_row}",
                num_rows=num_rows, num_cols=num_cols,
                headers=headers, source="hybrid_detected", score=score,
                has_header_fill=fk is not None,
                has_total_row=total_row is not None,
                expanded_left=expanded_left,
                expanded_right=expanded_right,
                has_empty_rows=has_empty_rows,
            ))

            new_covered |= cells_set
            for r2 in range(hdr_row, end_row + 1):
                for c2 in range(final_c1, final_c2 + 1):
                    used_cells.add((r2, c2))

        return new_covered

    # ────────────────────────────────────────────
    #  Stratégie 4 : Cohérence de types verticale
    # ────────────────────────────────────────────
    def _type_consistency_score(self, matrix, hdr_row, end_row, c1, c2) -> float:
        if end_row - hdr_row < 2:
            return 0

        consistent_cols = 0
        total_cols = 0

        for c in range(c1, c2 + 1):
            types_in_col = []
            for r in range(hdr_row + 1, end_row + 1):
                ci = self._get(matrix, r, c)
                if ci.value is not None:
                    types_in_col.append(ci.data_type)

            if len(types_in_col) >= 2:
                total_cols += 1
                dominant = max(set(types_in_col), key=types_in_col.count)
                ratio = types_in_col.count(dominant) / len(types_in_col)
                if ratio >= 0.7:
                    consistent_cols += 1

        if total_cols == 0:
            return 0
        return 10 * (consistent_cols / total_cols)

    # ────────────────────────────────────────────
    #  Passe 3 : Détection par grille de bordures
    #    (Stratégie 3)
    # ────────────────────────────────────────────
    def _grid_tables(self, matrix, sheet, results, covered,
                     max_r, max_c, ws) -> set:
        new_covered = set()

        # Trouver les cellules avec 4 bordures, non couvertes
        bordered = {
            pos for pos, ci in matrix.items()
            if ci.all_borders and pos not in covered and pos not in new_covered
        }

        if not bordered:
            return new_covered

        # Flood-fill strict (gap=0) sur les cellules bordurées
        blocks = self._flood_fill(bordered, gap=0)

        for block in blocks:
            rows_b = [r for r, c in block]
            cols_b = [c for r, c in block]
            r1, r2, c1, c2 = min(rows_b), max(rows_b), min(cols_b), max(cols_b)
            num_rows = r2 - r1
            num_cols = c2 - c1 + 1

            if num_rows < 1 or num_cols < 2:
                continue

            # Vérifier que c'est un rectangle bien rempli de bordures
            expected = (r2 - r1 + 1) * num_cols
            border_ratio = len(block) / max(expected, 1)
            if border_ratio < 0.6:
                continue

            # Scoring
            score = 30  # base grille
            score += 15 if num_cols >= 3 else 8
            score += 15 if num_rows >= 3 else 8
            score += self._type_consistency_score(matrix, r1, r2, c1, c2)

            # Densité de contenu
            filled = sum(1 for r, c in block
                         if self._get(matrix, r, c).value is not None)
            density = filled / max(len(block), 1)
            score += 10 if density >= 0.5 else 5 if density >= 0.3 else 0

            # Première ligne bold → bonus
            first_bold = all(
                self._get(matrix, r1, c).bold
                for c in range(c1, c2 + 1)
                if self._get(matrix, r1, c).value is not None
            )
            if first_bold:
                score += 5

            score = min(score, 90)

            title = self._find_section_title(matrix, r1, c1, c2, ws)
            if not title:
                title = f"Grille_{get_column_letter(c1)}{r1}"

            headers = [str(self._get(matrix, r1, c).value or "")
                       for c in range(c1, c2 + 1)]

            cells_set = set()
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cells_set.add((r, c))

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(c1)}{r1}",
                bottom_right=f"{get_column_letter(c2)}{r2}",
                num_rows=num_rows, num_cols=num_cols,
                headers=headers, source="grid_detected", score=score,
                has_grid_borders=True,
            ))
            new_covered |= cells_set

        return new_covered

    # ────────────────────────────────────────────
    #  Passe 4 : Blocs résiduels (gap=0 strict)
    # ────────────────────────────────────────────
    def _residual_blocks(self, matrix, sheet, results, covered,
                         content_cells, max_r, max_c, ws):
        remaining = {pos for pos in content_cells if pos not in covered}
        blocks = self._flood_fill(remaining, gap=0)

        # Tenter fusion de blocs proches avec même structure d'en-tête
        blocks = self._try_merge_blocks(blocks, matrix)

        for block in blocks:
            rows_b = [r for r, c in block]
            cols_b = [c for r, c in block]
            r1, r2, c1, c2 = min(rows_b), max(rows_b), min(cols_b), max(cols_b)
            num_rows = r2 - r1
            num_cols = c2 - c1 + 1

            if num_rows < 1 and num_cols < 2:
                continue

            score = 0.0
            score += 15 if (num_rows >= 3 and num_cols >= 2) else 8 if num_rows >= 1 else 0

            first_vals = [self._get(matrix, r1, c)
                          for c in range(c1, c2 + 1)
                          if self._get(matrix, r1, c).value is not None]
            if first_vals and all(ci.bold for ci in first_vals):
                score += 10

            fills = set(self._get(matrix, r1, c).fill_key
                        for c in range(c1, c2 + 1)
                        if self._get(matrix, r1, c).fill_key)
            if len(fills) == 1:
                score += 8

            total_cells = max((r2 - r1 + 1) * num_cols, 1)
            density = len(block) / total_cells
            score += 8 if density >= 0.4 else 4 if density >= 0.2 else 0

            score += self._type_consistency_score(matrix, r1, r2, c1, c2)

            # Pénalité décorative
            all_fills = set(self._get(matrix, r, c).fill_key
                            for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                            if self._get(matrix, r, c).fill_key)
            has_numbers = any(self._get(matrix, r, c).data_type in ("number", "date")
                              for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
            if len(all_fills) == 1 and not has_numbers and density < 0.3:
                score = max(score - 25, 5)

            font_sizes = [self._get(matrix, r, c).font_size
                          for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                          if self._get(matrix, r, c).value is not None]
            if font_sizes and min(font_sizes) >= 18:
                score = max(score - 15, 5)

            score = min(score, 55)

            title = self._find_section_title(matrix, r1, c1, c2, ws)
            if not title:
                title = f"Bloc_{get_column_letter(c1)}{r1}"

            headers = [str(self._get(matrix, r1, c).value or "")
                       for c in range(c1, c2 + 1)
                       if self._get(matrix, r1, c).value is not None]

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(c1)}{r1}",
                bottom_right=f"{get_column_letter(c2)}{r2}",
                num_rows=num_rows, num_cols=num_cols,
                headers=headers, source="contiguous_block", score=score,
            ))

    # ────────────────────────────────────────────
    #  Stratégie 5 (flood-fill) : Fusion de blocs proches
    # ────────────────────────────────────────────
    def _try_merge_blocks(self, blocks, matrix):
        if len(blocks) < 2:
            return blocks

        def block_bounds(b):
            rows = [r for r, c in b]
            cols = [c for r, c in b]
            return min(rows), max(rows), min(cols), max(cols)

        def header_sig(b, mat):
            r1, _, c1, c2 = block_bounds(b)
            return tuple(
                mat.get((r1, c), CellInfo()).data_type
                for c in range(c1, c2 + 1)
            )

        merged = []
        used = set()

        for i, b1 in enumerate(blocks):
            if i in used:
                continue
            r1_1, r2_1, c1_1, c2_1 = block_bounds(b1)
            sig1 = header_sig(b1, matrix)
            current = set(b1)

            for j, b2 in enumerate(blocks):
                if j <= i or j in used:
                    continue
                r1_2, r2_2, c1_2, c2_2 = block_bounds(b2)
                # Même colonnes et proches verticalement (≤2 lignes)
                if c1_2 == c1_1 and c2_2 == c2_1 and abs(r1_2 - r2_1) <= 2:
                    sig2 = header_sig(b2, matrix)
                    if sig1 == sig2 or len(sig1) == len(sig2):
                        current |= b2
                        used.add(j)

            merged.append(current)
            used.add(i)

        return merged

    # ────────────────────────────────────────────
    #  Utilitaires
    # ────────────────────────────────────────────
    def _find_section_title(self, matrix, top_row, c1, c2, ws) -> Optional[str]:
        # Fusion au-dessus
        for m in ws.merged_cells.ranges:
            if m.min_row >= top_row - 2 and m.max_row < top_row:
                if m.min_col <= c2 and m.max_col >= c1:
                    v = self._get(matrix, m.min_row, m.min_col).value
                    if v and isinstance(v, str):
                        return v.strip().replace("\n", " ")
        # Cellule bold grande au-dessus
        for offset in [1, 2]:
            r = top_row - offset
            if r < 1:
                break
            for c in range(c1, c2 + 1):
                ci = self._get(matrix, r, c)
                if (ci.value and isinstance(ci.value, str)
                        and ci.bold and ci.font_size >= 14):
                    return str(ci.value).strip().replace("\n", " ")
        # Texte simple au-dessus
        if top_row > 1:
            ci = self._get(matrix, top_row - 1, c1)
            if ci.value and isinstance(ci.value, str) and len(str(ci.value).strip()) > 1:
                return str(ci.value).strip()
        return None

    @staticmethod
    def _flood_fill(cells, gap=0):
        remaining, blocks = set(cells), []
        while remaining:
            seed = next(iter(remaining))
            block, queue = set(), [seed]
            while queue:
                cur = queue.pop()
                if cur in block or cur not in remaining:
                    continue
                block.add(cur)
                remaining.discard(cur)
                r, c = cur
                for dr in range(-gap - 1, gap + 2):
                    for dc in range(-gap - 1, gap + 2):
                        if dr == 0 and dc == 0:
                            continue
                        nb = (r + dr, c + dc)
                        if nb in remaining:
                            queue.append(nb)
            if block:
                blocks.append(block)
        return blocks


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Interface Streamlit
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.set_page_config(page_title="Excel Table Detector", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,400;0,500;0,700&family=JetBrains+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'DM Sans', sans-serif; font-weight: 700; }

.header-bar {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
    padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem; color: white;
}
.header-bar h1 { color: white; margin: 0 0 .3rem 0; font-size: 1.8rem; }
.header-bar p  { color: #94a3b8; margin: 0; font-size: .92rem; line-height: 1.5; }

.stat-row { display: flex; gap: .8rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.stat-card {
    flex: 1; min-width: 100px; background: #f8fafc; border: 1px solid #e2e8f0;
    border-radius: 12px; padding: .9rem 1rem; text-align: center;
}
.stat-card .num { font-size: 1.6rem; font-weight: 700; color: #1e293b; }
.stat-card .lbl { font-size: .72rem; color: #64748b; text-transform: uppercase;
                   letter-spacing: .04em; margin-top: .1rem; }

.table-card {
    background: #fff; border: 1px solid #e2e8f0; border-radius: 14px;
    padding: 1.2rem 1.4rem; margin-bottom: .8rem; transition: box-shadow .2s;
}
.table-card:hover { box-shadow: 0 4px 20px rgba(0,0,0,.06); }
.table-card .tc-header { display: flex; justify-content: space-between;
                          align-items: center; margin-bottom: .4rem; }
.table-card .tc-title { font-weight: 700; font-size: 1rem; color: #1e293b; }
.badge { display: inline-block; font-size: .65rem; font-weight: 600;
         padding: .15rem .5rem; border-radius: 6px; letter-spacing: .02em; }
.badge-excel  { background: #dcfce7; color: #166534; }
.badge-hybrid { background: #fef3c7; color: #92400e; }
.badge-grid   { background: #e0e7ff; color: #3730a3; }
.badge-block  { background: #f1f5f9; color: #475569; }
.meta { font-family: 'JetBrains Mono', monospace; font-size: .74rem; color: #64748b; }

.conf-bar { display: inline-block; height: 5px; border-radius: 3px; vertical-align: middle; }
.conf-high   { background: #22c55e; }
.conf-medium { background: #f59e0b; }
.conf-low    { background: #ef4444; }

.sheet-chip {
    display: inline-block; background: #f1f5f9; border: 1px solid #cbd5e1;
    border-radius: 8px; padding: .22rem .65rem; font-size: .8rem;
    color: #475569; margin: 0 .25rem .25rem 0;
}

.style-pills { display: flex; gap: .35rem; margin-top: .3rem; flex-wrap: wrap; }
.pill { display: inline-block; font-size: .64rem; padding: .1rem .4rem;
        border-radius: 4px; background: #f1f5f9; color: #475569; }
.pill-fill    { background: #fef3c7; color: #92400e; }
.pill-total   { background: #dcfce7; color: #166534; }
.pill-expand  { background: #ede9fe; color: #6d28d9; }
.pill-grid    { background: #e0e7ff; color: #3730a3; }
.pill-gap     { background: #fce7f3; color: #9d174d; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
    <h1>📊 Excel Table Detector v4</h1>
    <p>Détection hybride avancée — Header Score souple · Lookahead lignes vides ·
       Grilles de bordures · Profilage de types · Matrice pré-calculée</p>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Glissez un fichier Excel ici",
                            type=["xlsx", "xlsm"],
                            label_visibility="collapsed")
if not uploaded:
    st.info("⬆️  Chargez un fichier **.xlsx** pour commencer.")
    st.stop()

tmp_path = Path("/tmp") / uploaded.name
tmp_path.write_bytes(uploaded.getvalue())
detector = ExcelTableDetector(str(tmp_path))

st.markdown(f"### 📑 {len(detector.sheetnames)} onglet(s)")
chips = " ".join(f'<span class="sheet-chip">{s}</span>' for s in detector.sheetnames)
st.markdown(chips, unsafe_allow_html=True)
st.markdown("")

selected_sheets = st.multiselect(
    "Feuilles à analyser", options=detector.sheetnames, default=[],
    placeholder="Choisissez…",
)
if not selected_sheets:
    st.warning("👈  Sélectionnez au moins une feuille.")
    st.stop()

tables: list[DetectedTable] = []
for s in selected_sheets:
    with st.spinner(f"Analyse de « {s} »…"):
        tables += detector.detect_sheet(s)
tables.sort(key=lambda t: t.score, reverse=True)

n_excel  = sum(1 for t in tables if t.source == "excel_table")
n_hybrid = sum(1 for t in tables if t.source == "hybrid_detected")
n_grid   = sum(1 for t in tables if t.source == "grid_detected")
n_block  = sum(1 for t in tables if t.source == "contiguous_block")
avg = sum(t.score for t in tables) / max(len(tables), 1)

st.markdown(f"""
<div class="stat-row">
    <div class="stat-card"><div class="num">{n_excel}</div><div class="lbl">Tables Excel</div></div>
    <div class="stat-card"><div class="num">{n_hybrid}</div><div class="lbl">Hybrides</div></div>
    <div class="stat-card"><div class="num">{n_grid}</div><div class="lbl">Grilles</div></div>
    <div class="stat-card"><div class="num">{n_block}</div><div class="lbl">Blocs</div></div>
    <div class="stat-card"><div class="num">{len(tables)}</div><div class="lbl">Total</div></div>
    <div class="stat-card"><div class="num">{avg:.0f}%</div><div class="lbl">Score moy.</div></div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Filtres")
    sel_types = st.multiselect(
        "Type",
        ["excel_table", "hybrid_detected", "grid_detected", "contiguous_block"],
        ["excel_table", "hybrid_detected", "grid_detected", "contiguous_block"],
        format_func=lambda x: {"excel_table": "📊 Table Excel",
                                "hybrid_detected": "🎯 Hybride",
                                "grid_detected": "🔲 Grille",
                                "contiguous_block": "📦 Bloc"}[x],
    )
    min_score = st.slider("Score minimum", 0, 100, 0, step=5)
    min_rows = st.slider("Lignes minimum", 0, 20, 0)
    if len(selected_sheets) > 1:
        filter_sheets = st.multiselect("Feuille", selected_sheets, selected_sheets)
    else:
        filter_sheets = selected_sheets

filtered = [t for t in tables
            if t.sheet in filter_sheets and t.source in sel_types
            and t.score >= min_score and t.num_rows >= min_rows]

st.markdown(f"### {len(filtered)} tableau(x)")
if not filtered:
    st.warning("Aucun résultat.")
    st.stop()

for i, tbl in enumerate(filtered):
    bcls = {"excel_table": "badge-excel", "hybrid_detected": "badge-hybrid",
            "grid_detected": "badge-grid", "contiguous_block": "badge-block"
            }.get(tbl.source, "badge-block")
    ccls = {"haute": "conf-high", "moyenne": "conf-medium",
            "basse": "conf-low"}.get(tbl.confidence, "conf-low")
    bw = max(int(tbl.score * 0.8), 5)

    pills = ""
    if tbl.has_header_fill:
        pills += '<span class="pill pill-fill">Fill en-tête</span>'
    if tbl.has_total_row:
        pills += '<span class="pill pill-total">Ligne total</span>'
    if tbl.has_grid_borders:
        pills += '<span class="pill pill-grid">Grille bordures</span>'
    if tbl.expanded_left:
        pills += '<span class="pill pill-expand">← index</span>'
    if tbl.expanded_right:
        pills += '<span class="pill pill-expand">→ étendu</span>'
    if tbl.has_empty_rows:
        pills += '<span class="pill pill-gap">Lignes vides tolérées</span>'

    st.markdown(f"""
    <div class="table-card">
        <div class="tc-header">
            <span class="tc-title">{tbl.title}</span>
            <span class="badge {bcls}">{tbl.badge}</span>
        </div>
        <span class="meta">
            {tbl.sheet} &nbsp;·&nbsp; <b>{tbl.range_str}</b> &nbsp;·&nbsp;
            {tbl.num_rows}l × {tbl.num_cols}c &nbsp;·&nbsp;
            Score <b>{tbl.score:.0f}%</b>
            <span class="conf-bar {ccls}" style="width:{bw}px"></span>
            ({tbl.confidence})
        </span>
        <div class="style-pills">{pills}</div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🔍  {tbl.title}", expanded=False):
        df = detector.load_table(tbl)
        if df.empty:
            st.caption("_(vide)_")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.download_button(
                f"⬇️  CSV", df.to_csv(index=False).encode("utf-8"),
                file_name=f"{tbl.title}.csv", mime="text/csv", key=f"dl_{i}",
            )
