"""
app_excel_tables.py
───────────────────
Application Streamlit pour détecter et afficher
tous les tableaux d'un classeur Excel.

Lancement :
    pip install streamlit openpyxl pandas
    streamlit run app_excel_tables.py
"""

import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Détecteur de tableaux (repris de detect_excel_tables.py)
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

    @property
    def range_str(self):
        return f"{self.top_left}:{self.bottom_right}"

    @property
    def badge(self):
        return {"excel_table": "📊 Table Excel",
                "contiguous_block": "🔲 Bloc détecté"}.\
            get(self.source, self.source)


class ExcelTableDetector:
    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        self.wb = load_workbook(filepath, data_only=True)

    def detect_all(self) -> list[DetectedTable]:
        tables: list[DetectedTable] = []
        for name in self.wb.sheetnames:
            ws = self.wb[name]
            tables += self._excel_tables(ws, name)
            tables += self._contiguous_blocks(ws, name)
        return tables

    # ── Tables Excel déclarées ──────────────────
    def _excel_tables(self, ws, sheet):
        out = []
        for tbl in ws.tables.values():
            min_c, min_r, max_c, max_r = range_boundaries(tbl.ref)
            headers = [ws.cell(row=min_r, column=c).value or ""
                       for c in range(min_c, max_c + 1)]
            tl = f"{get_column_letter(min_c)}{min_r}"
            br = f"{get_column_letter(max_c)}{max_r}"
            out.append(DetectedTable(
                sheet=sheet, title=tbl.displayName,
                top_left=tl, bottom_right=br,
                num_rows=max_r - min_r, num_cols=max_c - min_c + 1,
                headers=headers, source="excel_table"))
        return out

    # ── Blocs contigus ──────────────────────────
    def _contiguous_blocks(self, ws, sheet):
        if ws.max_row is None or ws.max_column is None:
            return []

        filled = set()
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    filled.add((cell.row, cell.column))
        for m in ws.merged_cells.ranges:
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    filled.add((r, c))

        # Exclure cellules des Tables Excel
        excel_cells = set()
        for tbl in ws.tables.values():
            mc, mr, xc, xr = range_boundaries(tbl.ref)
            for r in range(mr, xr + 1):
                for c in range(mc, xc + 1):
                    excel_cells.add((r, c))
        candidates = filled - excel_cells

        blocks, out = self._flood_fill(candidates, gap=2), []
        for block in blocks:
            rows = [r for r, c in block]
            cols = [c for r, c in block]
            r1, r2, c1, c2 = min(rows), max(rows), min(cols), max(cols)
            # Ignorer les blocs d'une seule cellule
            if r2 == r1 and c2 == c1:
                continue
            title = self._find_title(ws, r1, c1, c2) or \
                f"Bloc_{get_column_letter(c1)}{r1}"
            headers = [str(ws.cell(row=r1, column=c).value or "")
                       for c in range(c1, c2 + 1)
                       if ws.cell(row=r1, column=c).value is not None]
            tl = f"{get_column_letter(c1)}{r1}"
            br = f"{get_column_letter(c2)}{r2}"
            out.append(DetectedTable(
                sheet=sheet, title=title,
                top_left=tl, bottom_right=br,
                num_rows=r2 - r1, num_cols=c2 - c1 + 1,
                headers=headers, source="contiguous_block"))
        return out

    def _find_title(self, ws, top_row, c1, c2):
        for m in ws.merged_cells.ranges:
            if m.max_row >= top_row - 2 and m.max_row <= top_row:
                if m.min_col <= c2 and m.max_col >= c1:
                    v = ws.cell(row=m.min_row, column=m.min_col).value
                    if v and isinstance(v, str):
                        return v.strip().replace("\n", " ")
        if top_row > 1:
            v = ws.cell(row=top_row - 1, column=c1).value
            if v and isinstance(v, str):
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

    def load_table(self, table: DetectedTable) -> pd.DataFrame:
        ws = self.wb[table.sheet]
        mc, mr, xc, xr = range_boundaries(
            f"{table.top_left}:{table.bottom_right}")
        data = []
        for row in ws.iter_rows(min_row=mr, max_row=xr,
                                min_col=mc, max_col=xc):
            data.append([cell.value for cell in row])
        if not data:
            return pd.DataFrame()

        def sanitize_cols(raw):
            seen: dict[str, int] = {}
            result = []
            for i, col in enumerate(raw):
                name = str(col).strip() if col is not None else ""
                if not name or name == "None":
                    name = f"Col_{i + 1}"
                if name in seen:
                    seen[name] += 1
                    name = f"{name}_{seen[name]}"
                else:
                    seen[name] = 0
                result.append(name)
            return result

        def row_is_empty(row):
            return all(v is None or str(v).strip() == "" for v in row)

        # Supprimer les lignes vides en début et en fin
        while data and row_is_empty(data[0]):
            data.pop(0)
        while data and row_is_empty(data[-1]):
            data.pop()

        if not data:
            return pd.DataFrame()

        # ── Tableau à une seule ligne : afficher comme données sans en-tête ──
        if len(data) == 1:
            cols = [f"Col_{i+1}" for i in range(len(data[0]))]
            return pd.DataFrame(data, columns=cols)

        # ── Cas normal : première ligne = en-tête ──
        clean_cols = sanitize_cols(data[0])
        rows = data[1:]
        # Filtrer les lignes vides intercalaires
        rows = [r for r in rows if not row_is_empty(r)]
        return pd.DataFrame(rows, columns=clean_cols)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Interface Streamlit
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# ── Configuration de la page ────────────────────
st.set_page_config(
    page_title="Excel Table Detector",
    page_icon="📊",
    layout="wide",
)

# ── Style personnalisé ──────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,400;0,500;0,700&family=JetBrains+Mono:wght@400;500&display=swap');

/* Global */
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'DM Sans', sans-serif; font-weight: 700; }

/* Header */
.header-bar {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    padding: 2rem 2.5rem;
    border-radius: 16px;
    margin-bottom: 1.5rem;
    color: white;
}
.header-bar h1 { color: white; margin: 0 0 .3rem 0; font-size: 1.8rem; }
.header-bar p  { color: #94a3b8; margin: 0; font-size: .95rem; }

/* Stat cards */
.stat-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; }
.stat-card {
    flex: 1;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.1rem 1.4rem;
    text-align: center;
}
.stat-card .num  { font-size: 1.8rem; font-weight: 700; color: #1e293b; }
.stat-card .lbl  { font-size: .8rem; color: #64748b; text-transform: uppercase;
                    letter-spacing: .04em; margin-top: .15rem; }

/* Table card */
.table-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1rem;
    transition: box-shadow .2s;
}
.table-card:hover { box-shadow: 0 4px 20px rgba(0,0,0,.06); }
.table-card .tc-header { display: flex; justify-content: space-between;
                          align-items: center; margin-bottom: .6rem; }
.table-card .tc-title  { font-weight: 700; font-size: 1.05rem; color: #1e293b; }
.badge {
    display: inline-block;
    font-size: .7rem; font-weight: 600;
    padding: .2rem .6rem;
    border-radius: 6px;
    letter-spacing: .02em;
}
.badge-excel  { background: #dcfce7; color: #166534; }
.badge-block  { background: #dbeafe; color: #1e40af; }
.meta { font-family: 'JetBrains Mono', monospace; font-size: .78rem;
        color: #64748b; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>📊 Excel Table Detector</h1>
    <p>Chargez un classeur Excel pour détecter et explorer automatiquement
       tous ses tableaux — tables déclarées et blocs de données contigus.</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ──────────────────────────────────────
uploaded = st.file_uploader(
    "Glissez un fichier Excel ici",
    type=["xlsx", "xlsm"],
    label_visibility="collapsed",
)

if not uploaded:
    st.info("⬆️  Chargez un fichier **.xlsx** pour commencer l'analyse.")
    st.stop()

# ── Sauvegarde temporaire & détection ───────────
tmp_path = Path("/tmp") / uploaded.name
tmp_path.write_bytes(uploaded.getvalue())

detector = ExcelTableDetector(str(tmp_path))
tables = detector.detect_all()

# ── Stats (calculées après sélection de l'onglet, donc plus bas) ───────────
n_sheets = len(detector.wb.sheetnames)

# ── Sélecteur d'onglet (zone principale) ───────
all_sheets = detector.wb.sheetnames

if len(all_sheets) == 1:
    active_sheet = all_sheets[0]
else:
    active_sheet = st.selectbox(
        "📋 Onglet à analyser",
        options=all_sheets,
        index=0,
        help="Choisissez l'onglet du classeur à explorer",
    )

# Filtrer les tableaux sur l'onglet actif
tables_on_sheet = [t for t in tables if t.sheet == active_sheet]

# ── Filtres (sidebar) ──────────────────────────
with st.sidebar:
    st.markdown("### Filtres")

    st.markdown(f"**Onglet actif :** `{active_sheet}`")
    st.caption(f"{len(all_sheets)} onglet(s) dans le classeur")
    st.divider()

    sel_types = st.multiselect(
        "Type de tableau",
        options=["excel_table", "contiguous_block"],
        default=["excel_table", "contiguous_block"],
        format_func=lambda x: "📊 Table Excel" if x == "excel_table"
                               else "🔲 Bloc détecté",
    )
    min_rows = st.slider("Nb min. de lignes (hors en-tête)", 0, 20, 0)

filtered = [t for t in tables_on_sheet
            if t.source in sel_types
            and t.num_rows >= min_rows]

# ── Stats dynamiques ──────────────────────────
n_excel_sheet = sum(1 for t in tables_on_sheet if t.source == "excel_table")
n_blocks_sheet = sum(1 for t in tables_on_sheet if t.source == "contiguous_block")

st.markdown(f"""
<div class="stat-row">
    <div class="stat-card"><div class="num">{n_sheets}</div>
         <div class="lbl">Onglets</div></div>
    <div class="stat-card"><div class="num">{n_excel_sheet}</div>
         <div class="lbl">Tables Excel</div></div>
    <div class="stat-card"><div class="num">{n_blocks_sheet}</div>
         <div class="lbl">Blocs détectés</div></div>
    <div class="stat-card"><div class="num">{len(tables_on_sheet)}</div>
         <div class="lbl">Total (onglet)</div></div>
</div>
""", unsafe_allow_html=True)

# ── Liste des tableaux ─────────────────────────
st.markdown(f"### {len(filtered)} tableau(x) affiché(s)")

if not filtered:
    st.warning("Aucun tableau ne correspond aux filtres sélectionnés.")
    st.stop()

for i, tbl in enumerate(filtered):
    badge_cls = "badge-excel" if tbl.source == "excel_table" else "badge-block"
    st.markdown(f"""
    <div class="table-card">
        <div class="tc-header">
            <span class="tc-title">{tbl.title}</span>
            <span class="badge {badge_cls}">{tbl.badge}</span>
        </div>
        <span class="meta">
            Onglet : <b>{tbl.sheet}</b> &nbsp;·&nbsp;
            Plage : <b>{tbl.range_str}</b> &nbsp;·&nbsp;
            {tbl.num_rows} ligne(s) × {tbl.num_cols} colonne(s)
        </span>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🔍  Voir les données — {tbl.title}", expanded=False):
        df = detector.load_table(tbl)
        if df.empty:
            st.caption("_(tableau vide — aucune donnée sous l'en-tête)_")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

            # Télécharger en CSV
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"⬇️  Télécharger « {tbl.title} » en CSV",
                csv,
                file_name=f"{tbl.title}.csv",
                mime="text/csv",
                key=f"dl_{i}",
            )