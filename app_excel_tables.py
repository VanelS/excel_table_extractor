"""
app_excel_tables.py
───────────────────
Application Streamlit — Détection hybride de tableaux Excel.

Combine 3 stratégies complémentaires :
  1. Tables Excel déclarées (ListObjects)
  2. Détection hybride style + contenu :
     - Style (bold+fill) pour ancrer les en-têtes
     - Contenu (flood-fill) pour étendre aux colonnes d'index et lignes manquantes
  3. Blocs contigus résiduels (fallback)

Lancement :
    pip install streamlit openpyxl pandas
    streamlit run app_excel_tables.py
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.cell.cell import Cell


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Helpers : style d'une cellule
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _has_fill(cell: Cell) -> bool:
    f = cell.fill
    return f is not None and f.patternType is not None and f.patternType != "none"


def _fill_key(cell: Cell) -> Optional[str]:
    if not _has_fill(cell):
        return None
    sc = cell.fill.start_color
    if sc is None:
        return None
    if sc.type == "rgb" and sc.rgb:
        return f"rgb:{sc.rgb}"
    if sc.type == "theme":
        return f"theme:{sc.theme}:tint:{round(sc.tint, 4) if sc.tint else 0}"
    if sc.type == "indexed":
        return f"indexed:{sc.value}"
    return None


def _has_border_top(cell: Cell) -> bool:
    b = cell.border
    return b is not None and b.top is not None and b.top.style is not None


def _has_any_border(cell: Cell) -> bool:
    b = cell.border
    if b is None:
        return False
    return any([
        b.left and b.left.style,
        b.right and b.right.style,
        b.top and b.top.style,
        b.bottom and b.bottom.style,
    ])


def _is_bold(cell: Cell) -> bool:
    return bool(cell.font and cell.font.bold)


def _font_size(cell: Cell) -> float:
    return float(cell.font.size) if cell.font and cell.font.size else 11.0


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Dataclass : tableau détecté
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
    source: str = ""          # excel_table | hybrid_detected | contiguous_block
    score: float = 0.0
    has_header_fill: bool = False
    has_total_row: bool = False
    header_fill_color: str = ""
    expanded_left: bool = False   # a-t-on étendu vers la gauche (index) ?
    expanded_right: bool = False

    @property
    def range_str(self):
        return f"{self.top_left}:{self.bottom_right}"

    @property
    def badge(self):
        return {
            "excel_table":      "📊 Table Excel",
            "hybrid_detected":  "🎯 Hybride style+contenu",
            "contiguous_block": "🔲 Bloc contigu",
        }.get(self.source, self.source)

    @property
    def confidence(self):
        if self.score >= 70:
            return "haute"
        if self.score >= 40:
            return "moyenne"
        return "basse"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Détecteur hybride
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

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
    #  API publique
    # ────────────────────────────────────────────
    def detect_sheet(self, sheet_name: str) -> list[DetectedTable]:
        self._ensure_wb()
        ws = self._wb[sheet_name]
        results: list[DetectedTable] = []

        # Pré-calculer la carte de contenu (cellules non-vides + fusions)
        content_map = self._build_content_map(ws)

        # Passe 1 : Tables Excel déclarées
        covered = self._declared_tables(ws, sheet_name, results, content_map)

        # Passe 2 : Détection hybride (style anchor + content expansion)
        covered |= self._hybrid_tables(ws, sheet_name, results, covered, content_map)

        # Passe 3 : Blocs contigus résiduels
        self._residual_blocks(ws, sheet_name, results, covered, content_map)

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
    #  Carte de contenu
    # ────────────────────────────────────────────
    def _build_content_map(self, ws) -> set[tuple[int, int]]:
        filled = set()
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1
        for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
            for cell in row:
                if cell.value is not None:
                    filled.add((cell.row, cell.column))
        for m in ws.merged_cells.ranges:
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    filled.add((r, c))
        return filled

    # ────────────────────────────────────────────
    #  Passe 1 : Tables Excel déclarées
    # ────────────────────────────────────────────
    def _declared_tables(self, ws, sheet, results, content_map) -> set:
        covered = set()
        for tbl in ws.tables.values():
            min_c, min_r, max_c, max_r = range_boundaries(tbl.ref)

            # Absorber cellules de la table
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    covered.add((r, c))

            # Vérifier ligne de total juste après
            has_total = False
            tr = max_r + 1
            if tr <= (ws.max_row or max_r):
                cell_t = ws.cell(row=tr, column=min_c)
                if cell_t.value and _is_bold(cell_t) and _has_border_top(cell_t):
                    has_total = True
                    max_r = tr
                    for c in range(min_c, max_c + 1):
                        covered.add((tr, c))

            # Titre de section
            title = self._find_section_title(ws, min_r, min_c, max_c)
            if not title:
                title = tbl.displayName

            headers = [ws.cell(row=min_r, column=c).value or ""
                       for c in range(min_c, max_c + 1)]
            hdr_fill = _fill_key(ws.cell(row=min_r, column=min_c))

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(min_c)}{min_r}",
                bottom_right=f"{get_column_letter(max_c)}{max_r}",
                num_rows=max_r - min_r, num_cols=max_c - min_c + 1,
                headers=headers, source="excel_table", score=100,
                has_header_fill=hdr_fill is not None,
                has_total_row=has_total,
                header_fill_color=hdr_fill or "",
            ))
        return covered

    # ────────────────────────────────────────────
    #  Passe 2 : Détection hybride
    #
    #  Étape A — Ancrage par style :
    #    Trouver les "runs" de cellules bold+fill
    #    sur une même ligne (= en-têtes candidats)
    #
    #  Étape B — Expansion par contenu :
    #    À partir de l'en-tête, étendre :
    #    - Vers le BAS tant qu'il y a du contenu
    #    - Vers la GAUCHE/DROITE si des colonnes
    #      adjacentes ont du contenu aligné
    #    - Identifier la ligne de total (bold + border-top)
    # ────────────────────────────────────────────
    def _hybrid_tables(self, ws, sheet, results, covered, content_map) -> set:
        new_covered = set()
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        # ── Étape A : trouver les header runs ───
        header_runs = []  # (row, col_start, col_end, fill_key)

        for r in range(1, max_row + 1):
            run_start = None
            run_fill = None
            for c in range(1, max_col + 2):  # +1 pour clôturer
                in_bounds = c <= max_col
                cell = ws.cell(row=r, column=c) if in_bounds else None

                skip = (not in_bounds
                        or (r, c) in covered
                        or (r, c) in new_covered)

                is_hdr_cell = False
                fk = None
                if not skip and cell is not None:
                    fk = _fill_key(cell)
                    is_hdr_cell = (_is_bold(cell) and fk is not None
                                   and cell.value is not None)

                if is_hdr_cell:
                    if run_start is None:
                        run_start = c
                        run_fill = fk
                    elif fk != run_fill:
                        if c - run_start >= 2:
                            header_runs.append((r, run_start, c - 1, run_fill))
                        run_start = c
                        run_fill = fk
                else:
                    if run_start is not None:
                        length = (c if in_bounds else c) - run_start
                        if length >= 2:
                            header_runs.append((r, run_start, c - 1, run_fill))
                    run_start = None
                    run_fill = None

        # ── Étape B : pour chaque header run, expansion par contenu ───
        used_cells = set()

        for hdr_row, style_c1, style_c2, fk in header_runs:
            # Vérifier que ces colonnes d'en-tête ne sont pas déjà traitées
            if any((hdr_row, c) in used_cells for c in range(style_c1, style_c2 + 1)):
                continue
            # Vérifier pas déjà couvert
            if any((hdr_row, c) in covered or (hdr_row, c) in new_covered
                   for c in range(style_c1, style_c2 + 1)):
                continue

            # B.1 — Extension vers le bas : trouver la dernière ligne de données
            last_data_row = hdr_row
            total_row = None

            for r in range(hdr_row + 1, max_row + 1):
                filled_count = 0
                is_bold_row = True
                has_top_border = False
                any_bold = False

                for c in range(style_c1, style_c2 + 1):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is not None:
                        filled_count += 1
                    if not _is_bold(cell):
                        is_bold_row = False
                    else:
                        any_bold = True
                    if _has_border_top(cell):
                        has_top_border = True

                if filled_count == 0:
                    break  # ligne vide → fin

                # Ligne de total = bold sur toutes les cellules remplies + border top
                if any_bold and has_top_border and is_bold_row:
                    total_row = r
                    break

                last_data_row = r

            end_row = total_row if total_row else last_data_row
            if end_row - hdr_row < 1:
                continue

            # B.2 — Extension latérale par contenu
            #        On regarde les colonnes adjacentes aux bornes du style :
            #        si elles ont du contenu sur les lignes de données → on les inclut
            final_c1 = style_c1
            final_c2 = style_c2
            expanded_left = False
            expanded_right = False

            # Extension vers la gauche
            c = style_c1 - 1
            while c >= 1:
                # Compter les cellules remplies dans cette colonne
                # sur la plage hdr_row → end_row
                col_filled = sum(
                    1 for r in range(hdr_row, end_row + 1)
                    if (r, c) in content_map and (r, c) not in covered
                )
                data_rows = end_row - hdr_row + 1
                # Inclure si ≥30% de remplissage ou si l'en-tête est rempli
                if col_filled >= max(data_rows * 0.3, 1):
                    final_c1 = c
                    expanded_left = True
                    c -= 1
                else:
                    break

            # Extension vers la droite
            c = style_c2 + 1
            while c <= max_col:
                col_filled = sum(
                    1 for r in range(hdr_row, end_row + 1)
                    if (r, c) in content_map and (r, c) not in covered
                )
                data_rows = end_row - hdr_row + 1
                if col_filled >= max(data_rows * 0.3, 1):
                    final_c2 = c
                    expanded_right = True
                    c += 1
                else:
                    break

            # B.3 — Extension vers le bas avec les colonnes élargies
            #        (il se peut que les colonnes d'index aient des données
            #         plus bas que les colonnes style)
            for r in range(end_row + 1, max_row + 1):
                if r == total_row:
                    continue
                filled_count = sum(
                    1 for c in range(final_c1, final_c2 + 1)
                    if (r, c) in content_map and (r, c) not in covered
                )
                if filled_count == 0:
                    break

                # Re-vérifier total row sur les nouvelles colonnes
                all_bold = all(
                    _is_bold(ws.cell(row=r, column=c))
                    for c in range(final_c1, final_c2 + 1)
                    if ws.cell(row=r, column=c).value is not None
                )
                any_top = any(
                    _has_border_top(ws.cell(row=r, column=c))
                    for c in range(final_c1, final_c2 + 1)
                )
                if all_bold and any_top:
                    total_row = r
                    end_row = r
                    break
                last_data_row = r
                end_row = max(end_row, last_data_row)

            if total_row and total_row > end_row:
                end_row = total_row

            num_cols = final_c2 - final_c1 + 1
            num_rows = end_row - hdr_row

            # Vérifier pas de chevauchement massif
            cells_set = set()
            overlap_count = 0
            total_count = 0
            for r in range(hdr_row, end_row + 1):
                for c in range(final_c1, final_c2 + 1):
                    total_count += 1
                    if (r, c) in covered or (r, c) in new_covered:
                        overlap_count += 1
                    else:
                        cells_set.add((r, c))

            if total_count > 0 and overlap_count / total_count > 0.3:
                continue

            # ── Scoring ──────────────────────
            score = 0.0
            score += 25  # fill d'en-tête
            score += 15 if num_cols >= 3 else 8 if num_cols >= 2 else 0
            score += 15 if num_rows >= 3 else 8 if num_rows >= 1 else 0
            score += 15 if total_row else 0

            # Cohérence de police dans les données
            data_sizes = set()
            for r in range(hdr_row + 1, end_row + 1):
                for c in range(final_c1, final_c2 + 1):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is not None:
                        data_sizes.add(_font_size(cell))
            score += 10 if len(data_sizes) <= 2 else 0

            # Densité de remplissage
            body_end = (total_row - 1) if total_row else end_row
            body_cells = max((body_end - hdr_row) * num_cols, 1)
            filled_ct = sum(
                1 for r in range(hdr_row + 1, body_end + 1)
                for c in range(final_c1, final_c2 + 1)
                if (r, c) in content_map
            )
            density = filled_ct / body_cells
            score += 10 if density >= 0.4 else 5 if density >= 0.2 else 0

            # Bordures
            border_ct = sum(
                1 for r in range(hdr_row, end_row + 1)
                for c in range(final_c1, final_c2 + 1)
                if _has_any_border(ws.cell(row=r, column=c))
            )
            border_ratio = border_ct / max(total_count, 1)
            score += 10 if border_ratio >= 0.3 else 0

            # Bonus : expansion latérale réussie (signe de robustesse)
            if expanded_left or expanded_right:
                score += 5

            score = min(score, 95)

            # ── Titre ────────────────────────
            title = self._find_section_title(ws, hdr_row, final_c1, final_c2)
            if not title:
                title = f"Tableau_{get_column_letter(final_c1)}{hdr_row}"

            headers = []
            for c in range(final_c1, final_c2 + 1):
                v = ws.cell(row=hdr_row, column=c).value
                headers.append(str(v) if v is not None else "")

            tl = f"{get_column_letter(final_c1)}{hdr_row}"
            br = f"{get_column_letter(final_c2)}{end_row}"

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=tl, bottom_right=br,
                num_rows=num_rows, num_cols=num_cols,
                headers=headers, source="hybrid_detected", score=score,
                has_header_fill=True,
                has_total_row=total_row is not None,
                header_fill_color=fk or "",
                expanded_left=expanded_left,
                expanded_right=expanded_right,
            ))

            new_covered |= cells_set
            for r in range(hdr_row, end_row + 1):
                for c in range(final_c1, final_c2 + 1):
                    used_cells.add((r, c))

        return new_covered

    # ────────────────────────────────────────────
    #  Passe 3 : Blocs contigus résiduels
    # ────────────────────────────────────────────
    def _residual_blocks(self, ws, sheet, results, covered, content_map):
        remaining = {pos for pos in content_map if pos not in covered}
        blocks = self._flood_fill(remaining, gap=1)

        for block in blocks:
            rows_b = [r for r, c in block]
            cols_b = [c for r, c in block]
            r1, r2, c1, c2 = min(rows_b), max(rows_b), min(cols_b), max(cols_b)
            num_rows = r2 - r1
            num_cols = c2 - c1 + 1

            if num_rows < 1 and num_cols < 2:
                continue

            # ── Scoring ──────────────────────
            score = 0.0
            score += 20 if (num_rows >= 3 and num_cols >= 2) else 10 if num_rows >= 1 else 0

            # Première ligne bold ?
            first_row_vals = [
                ws.cell(row=r1, column=c)
                for c in range(c1, c2 + 1)
                if ws.cell(row=r1, column=c).value is not None
            ]
            if first_row_vals and all(_is_bold(c) for c in first_row_vals):
                score += 15

            # Première ligne fill uniforme ?
            fills = set(_fill_key(ws.cell(row=r1, column=c))
                        for c in range(c1, c2 + 1)
                        if _fill_key(ws.cell(row=r1, column=c)) is not None)
            if len(fills) == 1:
                score += 10

            # Densité
            total_cells = max((r2 - r1 + 1) * num_cols, 1)
            density = len(block) / total_cells
            score += 10 if density >= 0.4 else 5 if density >= 0.2 else 0

            # Pénalités : zone décorative
            all_fills = set()
            has_numbers = False
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    fk = _fill_key(ws.cell(row=r, column=c))
                    if fk:
                        all_fills.add(fk)
                    if isinstance(ws.cell(row=r, column=c).value, (int, float)):
                        has_numbers = True
            if len(all_fills) == 1 and not has_numbers and density < 0.3:
                score = max(score - 30, 5)

            # Pénalité : police très grande → titre/déco
            font_sizes = [
                _font_size(ws.cell(row=r, column=c))
                for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                if ws.cell(row=r, column=c).value is not None
            ]
            if font_sizes and min(font_sizes) >= 18:
                score = max(score - 20, 5)

            score = min(score, 60)

            title = self._find_section_title(ws, r1, c1, c2)
            if not title:
                title = f"Bloc_{get_column_letter(c1)}{r1}"

            headers = [
                str(ws.cell(row=r1, column=c).value or "")
                for c in range(c1, c2 + 1)
                if ws.cell(row=r1, column=c).value is not None
            ]

            results.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=f"{get_column_letter(c1)}{r1}",
                bottom_right=f"{get_column_letter(c2)}{r2}",
                num_rows=num_rows, num_cols=num_cols,
                headers=headers, source="contiguous_block", score=score,
            ))

    # ────────────────────────────────────────────
    #  Utilitaires
    # ────────────────────────────────────────────
    def _find_section_title(self, ws, top_row, c1, c2) -> Optional[str]:
        # 1. Fusion au-dessus
        for m in ws.merged_cells.ranges:
            if m.min_row >= top_row - 2 and m.max_row < top_row:
                if m.min_col <= c2 and m.max_col >= c1:
                    v = ws.cell(row=m.min_row, column=m.min_col).value
                    if v and isinstance(v, str):
                        return v.strip().replace("\n", " ")
        # 2. Cellule bold + grande police au-dessus
        for offset in [1, 2]:
            r = top_row - offset
            if r < 1:
                break
            for c in range(c1, c2 + 1):
                cell = ws.cell(row=r, column=c)
                if (cell.value and isinstance(cell.value, str)
                        and _is_bold(cell) and _font_size(cell) >= 14):
                    return cell.value.strip().replace("\n", " ")
        # 3. Texte simple au-dessus
        if top_row > 1:
            v = ws.cell(row=top_row - 1, column=c1).value
            if v and isinstance(v, str) and len(v.strip()) > 1:
                return v.strip()
        return None

    @staticmethod
    def _flood_fill(cells, gap=1):
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
                for dr in range(-gap, gap + 1):
                    for dc in range(-gap, gap + 1):
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
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem; color: white;
}
.header-bar h1 { color: white; margin: 0 0 .3rem 0; font-size: 1.8rem; }
.header-bar p  { color: #94a3b8; margin: 0; font-size: .95rem; }

.stat-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.stat-card {
    flex: 1; min-width: 110px; background: #f8fafc; border: 1px solid #e2e8f0;
    border-radius: 12px; padding: 1rem 1.2rem; text-align: center;
}
.stat-card .num { font-size: 1.7rem; font-weight: 700; color: #1e293b; }
.stat-card .lbl { font-size: .75rem; color: #64748b; text-transform: uppercase;
                   letter-spacing: .04em; margin-top: .1rem; }

.table-card {
    background: #fff; border: 1px solid #e2e8f0; border-radius: 14px;
    padding: 1.3rem 1.5rem; margin-bottom: 1rem; transition: box-shadow .2s;
}
.table-card:hover { box-shadow: 0 4px 20px rgba(0,0,0,.06); }
.table-card .tc-header { display: flex; justify-content: space-between;
                          align-items: center; margin-bottom: .5rem; }
.table-card .tc-title { font-weight: 700; font-size: 1.05rem; color: #1e293b; }
.badge { display: inline-block; font-size: .68rem; font-weight: 600;
         padding: .18rem .55rem; border-radius: 6px; letter-spacing: .02em; }
.badge-excel  { background: #dcfce7; color: #166534; }
.badge-hybrid { background: #fef3c7; color: #92400e; }
.badge-block  { background: #dbeafe; color: #1e40af; }
.meta { font-family: 'JetBrains Mono', monospace; font-size: .76rem; color: #64748b; }

.conf-bar { display: inline-block; height: 6px; border-radius: 3px; vertical-align: middle; }
.conf-high   { background: #22c55e; }
.conf-medium { background: #f59e0b; }
.conf-low    { background: #ef4444; }

.sheet-chip {
    display: inline-block; background: #f1f5f9; border: 1px solid #cbd5e1;
    border-radius: 8px; padding: .25rem .7rem; font-size: .82rem;
    color: #475569; margin: 0 .3rem .3rem 0;
}

.style-pills { display: flex; gap: .4rem; margin-top: .35rem; flex-wrap: wrap; }
.pill { display: inline-block; font-size: .66rem; padding: .12rem .45rem;
        border-radius: 4px; background: #f1f5f9; color: #475569; }
.pill-fill    { background: #fef3c7; color: #92400e; }
.pill-total   { background: #dcfce7; color: #166534; }
.pill-expand  { background: #ede9fe; color: #6d28d9; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>📊 Excel Table Detector v3</h1>
    <p>Détection hybride : ancrage par style (fills, bordures, gras) +
       expansion par contenu (flood-fill, colonnes d'index). Scoring de confiance.</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ──────────────────────────────────────
uploaded = st.file_uploader("Glissez un fichier Excel ici",
                            type=["xlsx", "xlsm"],
                            label_visibility="collapsed")
if not uploaded:
    st.info("⬆️  Chargez un fichier **.xlsx** pour commencer.")
    st.stop()

tmp_path = Path("/tmp") / uploaded.name
tmp_path.write_bytes(uploaded.getvalue())
detector = ExcelTableDetector(str(tmp_path))

# ── Sélection des feuilles ──────────────────────
st.markdown(f"### 📑 Le classeur contient **{len(detector.sheetnames)}** onglet(s)")
chips = " ".join(f'<span class="sheet-chip">{s}</span>' for s in detector.sheetnames)
st.markdown(chips, unsafe_allow_html=True)
st.markdown("")

selected_sheets = st.multiselect(
    "Sélectionnez les feuilles à analyser",
    options=detector.sheetnames, default=[],
    placeholder="Choisissez une ou plusieurs feuilles…",
)
if not selected_sheets:
    st.warning("👈  Sélectionnez au moins une feuille.")
    st.stop()

# ── Détection ───────────────────────────────────
tables: list[DetectedTable] = []
for s in selected_sheets:
    with st.spinner(f"Analyse de « {s} »…"):
        tables += detector.detect_sheet(s)
tables.sort(key=lambda t: t.score, reverse=True)

# ── Stats ───────────────────────────────────────
n_excel  = sum(1 for t in tables if t.source == "excel_table")
n_hybrid = sum(1 for t in tables if t.source == "hybrid_detected")
n_block  = sum(1 for t in tables if t.source == "contiguous_block")
avg_score = sum(t.score for t in tables) / max(len(tables), 1)

st.markdown(f"""
<div class="stat-row">
    <div class="stat-card"><div class="num">{len(selected_sheets)}</div>
         <div class="lbl">Feuille(s)</div></div>
    <div class="stat-card"><div class="num">{n_excel}</div>
         <div class="lbl">Tables Excel</div></div>
    <div class="stat-card"><div class="num">{n_hybrid}</div>
         <div class="lbl">Hybrides</div></div>
    <div class="stat-card"><div class="num">{n_block}</div>
         <div class="lbl">Blocs résiduels</div></div>
    <div class="stat-card"><div class="num">{len(tables)}</div>
         <div class="lbl">Total</div></div>
    <div class="stat-card"><div class="num">{avg_score:.0f}%</div>
         <div class="lbl">Score moyen</div></div>
</div>
""", unsafe_allow_html=True)

# ── Filtres (sidebar) ──────────────────────────
with st.sidebar:
    st.markdown("### Filtres")
    sel_types = st.multiselect(
        "Type de détection",
        options=["excel_table", "hybrid_detected", "contiguous_block"],
        default=["excel_table", "hybrid_detected", "contiguous_block"],
        format_func=lambda x: {"excel_table": "📊 Table Excel",
                                "hybrid_detected": "🎯 Hybride",
                                "contiguous_block": "🔲 Bloc contigu"}[x],
    )
    min_score = st.slider("Score minimum", 0, 100, 0, step=5)
    min_rows = st.slider("Nb min. de lignes", 0, 20, 0)
    if len(selected_sheets) > 1:
        filter_sheets = st.multiselect("Feuille", selected_sheets, selected_sheets)
    else:
        filter_sheets = selected_sheets

filtered = [t for t in tables
            if t.sheet in filter_sheets
            and t.source in sel_types
            and t.score >= min_score
            and t.num_rows >= min_rows]

# ── Affichage ──────────────────────────────────
st.markdown(f"### {len(filtered)} tableau(x) affiché(s)")
if not filtered:
    st.warning("Aucun tableau ne correspond aux filtres.")
    st.stop()

for i, tbl in enumerate(filtered):
    badge_cls = {"excel_table": "badge-excel",
                 "hybrid_detected": "badge-hybrid",
                 "contiguous_block": "badge-block"}.get(tbl.source, "badge-block")
    c_cls = {"haute": "conf-high", "moyenne": "conf-medium",
             "basse": "conf-low"}.get(tbl.confidence, "conf-low")
    bar_w = max(int(tbl.score * 0.8), 5)

    pills = ""
    if tbl.has_header_fill:
        pills += '<span class="pill pill-fill">En-tête colorée</span>'
    if tbl.has_total_row:
        pills += '<span class="pill pill-total">Ligne de total</span>'
    if tbl.expanded_left:
        pills += '<span class="pill pill-expand">Étendu ← (index)</span>'
    if tbl.expanded_right:
        pills += '<span class="pill pill-expand">Étendu →</span>'

    st.markdown(f"""
    <div class="table-card">
        <div class="tc-header">
            <span class="tc-title">{tbl.title}</span>
            <span class="badge {badge_cls}">{tbl.badge}</span>
        </div>
        <span class="meta">
            Onglet : <b>{tbl.sheet}</b> &nbsp;·&nbsp;
            Plage : <b>{tbl.range_str}</b> &nbsp;·&nbsp;
            {tbl.num_rows} ligne(s) × {tbl.num_cols} col.
            &nbsp;·&nbsp;
            Score : <b>{tbl.score:.0f}%</b>
            <span class="conf-bar {c_cls}" style="width:{bar_w}px"></span>
            ({tbl.confidence})
        </span>
        <div class="style-pills">{pills}</div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🔍  Voir les données — {tbl.title}", expanded=False):
        df = detector.load_table(tbl)
        if df.empty:
            st.caption("_(tableau vide)_")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"⬇️  Télécharger « {tbl.title} » en CSV",
                csv, file_name=f"{tbl.title}.csv",
                mime="text/csv", key=f"dl_{i}",
            )
