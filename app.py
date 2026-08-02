#!/usr/bin/env python3
"""Sleek CSV/Excel table and chart builder.

Run the web app:
    streamlit run app.py

Create a demo HTML file without Streamlit:
    python app.py --demo-output demo_table.html
"""

from __future__ import annotations

import argparse
import csv
import html
import io
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


META_KEYS = {
    "title",
    "subtitle",
    "footer",
    "source",
    "source_url",
    "brand",
    "accent",
}

TITLE_SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "in", "of", "on", "or", "the", "to", "vs", "with"}

THEMES = {
    "Action Green": {
        "accent": "#00b67a",
        "accent_bright": "#00e09a",
        "accent_dark": "#00865b",
        "navy": "#101820",
        "navy_2": "#152630",
        "background": "#ffffff",
        "surface": "#f7faf8",
        "text": "#20252a",
        "muted": "#6d7880",
        "border": "#dfe7e3",
    },
    "Midnight Editorial": {
        "accent": "#7ee8c2",
        "accent_bright": "#b6ffe5",
        "accent_dark": "#35b98c",
        "navy": "#07111b",
        "navy_2": "#122433",
        "background": "#ffffff",
        "surface": "#f4f7f9",
        "text": "#1e2933",
        "muted": "#657480",
        "border": "#d9e2e8",
    },
    "Royal Blue": {
        "accent": "#3978ff",
        "accent_bright": "#70a2ff",
        "accent_dark": "#1f55c8",
        "navy": "#0c1830",
        "navy_2": "#16294d",
        "background": "#ffffff",
        "surface": "#f5f8ff",
        "text": "#1d2738",
        "muted": "#69758a",
        "border": "#dce4f2",
    },
    "Warm Editorial": {
        "accent": "#e66b38",
        "accent_bright": "#ff9a67",
        "accent_dark": "#b64920",
        "navy": "#211b19",
        "navy_2": "#3a2a25",
        "background": "#fffdf9",
        "surface": "#faf4ec",
        "text": "#302a27",
        "muted": "#7d7069",
        "border": "#e8ddd2",
    },
}

VISUALS = [
    {
        "key": "editorial_table",
        "name": "Editorial Table",
        "kind": "table",
        "description": "Full interactive table with search, sorting and a data download.",
    },
    {
        "key": "ranking_table",
        "name": "Ranking Table",
        "kind": "table",
        "description": "Podium emphasis and a visual data bar on the chosen metric.",
    },
    {
        "key": "compact_table",
        "name": "Compact Table",
        "kind": "table",
        "description": "Dense, restrained presentation for wide or detailed datasets.",
    },
    {
        "key": "bar_horizontal",
        "name": "Horizontal Bars",
        "kind": "chart",
        "description": "Best for rankings, long labels and six or more categories.",
    },
    {
        "key": "bar_vertical",
        "name": "Vertical Bars",
        "kind": "chart",
        "description": "Best for a short categorical comparison.",
    },
    {
        "key": "line",
        "name": "Line Chart",
        "kind": "chart",
        "description": "Best for ordered values and change over time.",
    },
    {
        "key": "area",
        "name": "Area Chart",
        "kind": "chart",
        "description": "A more visual treatment for a single ordered series.",
    },
    {
        "key": "donut",
        "name": "Donut Chart",
        "kind": "chart",
        "description": "Use only when categories form a meaningful whole.",
    },
    {
        "key": "scatter",
        "name": "Scatter Plot",
        "kind": "chart",
        "description": "Shows the relationship between two numeric measures.",
    },
]


@dataclass
class Dataset:
    frame: pd.DataFrame
    metadata: dict[str, str]


def clean_column_name(value: Any) -> str:
    text = str(value).strip()
    return text or "Untitled column"


def smart_title_case(value: Any) -> str:
    """Apply strict editorial title case without all-capital words."""
    text = str(value).strip()
    if not text:
        return ""
    parts = re.split(r"(\s+|[·|:–—/\-])", text)
    word_positions = [i for i, part in enumerate(parts) if part and not re.fullmatch(r"\s+|[·|:–—/\-]", part)]
    last_position = word_positions[-1] if word_positions else -1
    for position in word_positions:
        token = parts[position]
        match = re.match(r"^([^A-Za-z0-9]*)(.*?)([^A-Za-z0-9]*)$", token)
        if not match:
            continue
        prefix, core, suffix = match.groups()
        lower = core.lower()
        if position != word_positions[0] and position != last_position and lower in TITLE_SMALL_WORDS:
            converted = lower
        else:
            converted = lower[:1].upper() + lower[1:]
        parts[position] = f"{prefix}{converted}{suffix}"
    return "".join(parts)


def parse_metadata_csv(raw: bytes, encoding: str = "utf-8-sig") -> Dataset:
    """Read optional '# key: value' lines, then parse the normal CSV table."""
    text = raw.decode(encoding, errors="replace")
    lines = text.splitlines()
    metadata: dict[str, str] = {}
    data_lines: list[str] = []
    table_started = False

    for line in lines:
        stripped = line.strip()
        if not table_started and (not stripped or stripped.startswith("#")):
            if stripped.startswith("#"):
                match = re.match(r"^#\s*([A-Za-z_]+)\s*:\s*(.*)$", stripped)
                if match:
                    key = match.group(1).lower()
                    if key in META_KEYS:
                        metadata[key] = match.group(2).strip()
            continue
        table_started = True
        data_lines.append(line)

    if not data_lines:
        raise ValueError("No table header row or data was found.")

    data_text = "\n".join(data_lines)
    try:
        dialect = csv.Sniffer().sniff(data_text[:5000], delimiters=",;\t|")
        separator = dialect.delimiter
    except csv.Error:
        separator = ","

    frame = pd.read_csv(io.StringIO(data_text), sep=separator)
    return Dataset(normalize_frame(frame), metadata)


def parse_uploaded_file(raw: bytes, filename: str, sheet_name: str | int = 0) -> Dataset:
    suffix = Path(filename).suffix.lower()
    if suffix in {".csv", ".txt", ".tsv"}:
        return parse_metadata_csv(raw)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        frame = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name)
        return Dataset(normalize_frame(frame), {})
    raise ValueError("Please upload a supported data file.")


def normalize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [clean_column_name(c) for c in frame.columns]
    frame = frame.dropna(how="all").dropna(axis=1, how="all")
    if frame.empty or len(frame.columns) == 0:
        raise ValueError("The uploaded file does not contain a usable table.")

    for column in frame.columns:
        if frame[column].dtype == object:
            stripped = frame[column].astype(str).str.strip()
            numeric_candidate = (
                stripped
                .str.replace(r"[,$£€%]", "", regex=True)
                .str.replace(r"^\+", "", regex=True)
                .str.replace(r"\((.*)\)", r"-\1", regex=True)
            )
            converted = pd.to_numeric(numeric_candidate, errors="coerce")
            non_empty = stripped.ne("") & stripped.ne("nan")
            if non_empty.sum() and converted[non_empty].notna().mean() >= 0.9:
                frame[column] = converted
            else:
                frame[column] = frame[column].where(frame[column].notna(), "")
    return frame


def numeric_columns(frame: pd.DataFrame) -> list[str]:
    return [c for c in frame.columns if pd.api.types.is_numeric_dtype(frame[c])]


def categorical_columns(frame: pd.DataFrame) -> list[str]:
    numeric = set(numeric_columns(frame))
    return [c for c in frame.columns if c not in numeric]


def infer_format(column: str, series: pd.Series) -> str:
    name = column.lower()
    if any(token in name for token in ("percent", "percentage", "probability", "rate", "share", "%")):
        return "percent"
    if any(token in name for token in ("price", "cost", "revenue", "income", "salary", "bonus", "usd", "cad", "gbp")):
        return "currency"
    values = pd.to_numeric(series, errors="coerce").dropna()
    if not values.empty and (values % 1 == 0).all():
        return "integer"
    return "decimal"


def format_value(value: Any, column: str, series: pd.Series) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    if not pd.api.types.is_numeric_dtype(series):
        return str(value)

    number = float(value)
    kind = infer_format(column, series)
    if kind == "percent":
        return f"{number:,.1f}%"
    if kind == "currency":
        symbol = "£" if "gbp" in column.lower() else "€" if "eur" in column.lower() else "$"
        return f"{symbol}{number:,.0f}"
    if kind == "integer":
        return f"{number:,.0f}"
    return f"{number:,.2f}".rstrip("0").rstrip(".")


def animated_number_markup(value: Any, column: str, series: pd.Series, extra_class: str = "") -> str:
    """Return a safely formatted number that counts up in generated HTML."""
    if value is None or pd.isna(value):
        return "—"
    if not pd.api.types.is_numeric_dtype(series):
        return html.escape(str(value))
    number = float(value)
    kind = infer_format(column, series)
    symbol = ""
    if kind == "currency":
        symbol = "£" if "gbp" in column.lower() else "€" if "eur" in column.lower() else "$"
    classes = "animated-number" + (f" {extra_class}" if extra_class else "")
    fallback = html.escape(format_value(number, column, series))
    return (
        f'<span class="{classes}" data-animate-number="{number}" '
        f'data-number-kind="{kind}" data-number-prefix="{html.escape(symbol, quote=True)}">{fallback}</span>'
    )


def animation_script() -> str:
    return """
    (() => {
      const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      const formatNumber = (value, kind, prefix) => {
        if (kind === 'percent') return value.toLocaleString(undefined,{maximumFractionDigits:1,minimumFractionDigits:1}) + '%';
        if (kind === 'currency') return prefix + value.toLocaleString(undefined,{maximumFractionDigits:0});
        if (kind === 'integer') return value.toLocaleString(undefined,{maximumFractionDigits:0});
        if (kind === 'compact') return new Intl.NumberFormat(undefined,{notation:'compact',maximumFractionDigits:1}).format(value);
        return value.toLocaleString(undefined,{maximumFractionDigits:2});
      };
      document.querySelectorAll('[data-animate-number]').forEach((element,index) => {
        const target = Number(element.dataset.animateNumber);
        const kind = element.dataset.numberKind || 'decimal';
        const prefix = element.dataset.numberPrefix || '';
        if (!Number.isFinite(target) || reduceMotion) {
          if (Number.isFinite(target)) element.textContent = formatNumber(target,kind,prefix);
          return;
        }
        const duration = 1050;
        const delay = Math.min(index * 45, 360);
        const start = performance.now() + delay;
        const tick = now => {
          if (now < start) { requestAnimationFrame(tick); return; }
          const progress = Math.min((now-start)/duration,1);
          const eased = 1-Math.pow(1-progress,4);
          element.textContent = formatNumber(target*eased,kind,prefix);
          if (progress < 1) requestAnimationFrame(tick);
        };
        requestAnimationFrame(tick);
      });
    })();
    """


def safe_id(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "-", value).strip("-").lower() or "data-viz"


def json_records(frame: pd.DataFrame) -> list[dict[str, Any]]:
    safe = frame.where(pd.notna(frame), None)
    return safe.to_dict(orient="records")


def json_safe(value: Any) -> Any:
    """Convert spreadsheet values into JSON-safe primitives."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "item"):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            pass
    return str(value)


def safe_json_dumps(value: Any) -> str:
    """Serialize safely for an inline script element."""
    return (
        json.dumps(value, ensure_ascii=False)
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
    )


def merge_theme(theme_name: str, accent_override: str = "") -> dict[str, str]:
    theme = dict(THEMES.get(theme_name, THEMES["Action Green"]))
    if re.fullmatch(r"#[0-9a-fA-F]{6}", accent_override or ""):
        theme["accent"] = accent_override
        theme["accent_dark"] = accent_override
    return theme


def base_css(theme: dict[str, str], compact: bool = False, preview: bool = False) -> str:
    pad = "10px 12px" if compact else "14px 15px"
    font_size = "13px" if compact else "14.5px"
    preview_css = """
    .viz-shell{box-shadow:none;border-radius:4px}.hero{padding:18px 22px 17px}h1{font-size:28px;margin:5px 0 6px}
    .brand{font-size:9px}.subtitle{font-size:11px;line-height:1.35}.content{padding:12px}.toolbar{margin-bottom:8px}
    .toolbar .btn{display:none}.search input{padding:7px 10px 7px 30px;font-size:11px}.search:before{left:10px;top:3px}
    .table-wrap{max-height:185px}table{min-width:560px;font-size:9.5px}thead th,tbody td{padding:7px 8px}
    thead th{font-size:8px}.rank-badge{min-width:21px;height:21px}.metric{min-width:115px;grid-template-columns:1fr 45px}
    .chart-wrap{padding:13px}.chart-title{font-size:14px;margin-bottom:8px}.chart-svg{min-width:0}.footer{padding:7px 12px;font-size:8.5px}
    """ if preview else ""
    return f"""
    :root{{--accent:{theme['accent']};--accent-bright:{theme['accent_bright']};--accent-dark:{theme['accent_dark']};
      --navy:{theme['navy']};--navy-2:{theme['navy_2']};--bg:{theme['background']};--surface:{theme['surface']};
      --text:{theme['text']};--muted:{theme['muted']};--border:{theme['border']};}}
    *{{box-sizing:border-box}} html,body{{margin:0;padding:0;background:var(--bg);color:var(--text)}}
    body{{font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;line-height:1.55}}
    .viz-shell{{width:100%;max-width:1180px;margin:0 auto;background:var(--bg);overflow:hidden;border:1px solid var(--border);border-radius:6px;
      box-shadow:0 18px 46px rgba(16,24,32,.10);animation:shellRise .72s cubic-bezier(.2,.72,.2,1) both}}
    .hero{{position:relative;padding:34px clamp(20px,4vw,52px) 30px;color:#fff;background:
      radial-gradient(750px 260px at 85% -15%,color-mix(in srgb,var(--accent) 25%,transparent),transparent 65%),
      linear-gradient(145deg,var(--navy),var(--navy-2));overflow:hidden}}
    .hero:after{{content:"";position:absolute;inset:0;pointer-events:none;background-image:
      linear-gradient(rgba(255,255,255,.035) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.035) 1px,transparent 1px);
      background-size:48px 48px;mask-image:linear-gradient(to left,#000,transparent 75%)}}
    .brand{{position:relative;z-index:1;font-weight:800;font-size:12px;letter-spacing:.14em;color:var(--accent-bright);animation:fadeUp .72s .08s both}}
    h1{{position:relative;z-index:1;margin:9px 0 10px;font-family:"Arial Narrow",Impact,Inter,sans-serif;font-size:clamp(34px,5vw,62px);
      line-height:.98;letter-spacing:-.01em;animation:fadeUp .78s .14s both}}
    .subtitle{{position:relative;z-index:1;max-width:850px;margin:0;color:rgba(255,255,255,.78);font-size:16px;animation:fadeUp .78s .22s both}}
    .content{{padding:clamp(18px,3vw,34px);animation:fadeUp .82s .28s both}}
    .toolbar{{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:14px}}
    .search{{position:relative;flex:1 1 260px}} .search input{{width:100%;border:1px solid var(--border);background:#fff;color:var(--text);
      padding:11px 14px 11px 38px;font:inherit}} .search:before{{content:"⌕";position:absolute;left:13px;top:8px;color:var(--muted);font-size:20px}}
    .btn{{appearance:none;border:1px solid var(--border);border-radius:5px;background:#fff;color:var(--text);font-weight:750;letter-spacing:.04em;padding:10px 14px;cursor:pointer;transition:transform .2s,border-color .2s,color .2s,box-shadow .2s}}
    .btn:hover{{border-color:var(--accent);color:var(--accent-dark);transform:translateY(-2px);box-shadow:0 8px 20px rgba(16,24,32,.08)}} .count{{font-size:12px;color:var(--muted);margin-left:auto}}
    .table-wrap{{overflow:auto;border:1px solid var(--border);max-height:720px;scrollbar-color:var(--accent) #edf2ef}}
    table{{width:100%;border-collapse:collapse;min-width:760px;background:#fff;font-size:{font_size}}}
    thead th{{position:sticky;top:0;z-index:2;text-align:left;background:var(--navy);color:#fff;padding:{pad};white-space:nowrap;
      font-size:12px;letter-spacing:.035em;cursor:pointer;user-select:none}}
    thead th:hover{{background:var(--navy-2)}} .sort{{margin-left:6px;color:var(--accent-bright);opacity:.45}}
    tbody td{{padding:{pad};border-bottom:1px solid var(--border);vertical-align:middle}}
    tbody tr{{animation:rowRise .52s cubic-bezier(.2,.72,.2,1) both}} tbody tr:nth-child(1){{animation-delay:.04s}} tbody tr:nth-child(2){{animation-delay:.08s}}
    tbody tr:nth-child(3){{animation-delay:.12s}} tbody tr:nth-child(4){{animation-delay:.16s}} tbody tr:nth-child(5){{animation-delay:.20s}}
    tbody tr:nth-child(even){{background:var(--surface)}} tbody tr:hover{{background:color-mix(in srgb,var(--accent) 10%,#fff)}}
    tbody tr.podium{{background:linear-gradient(90deg,color-mix(in srgb,var(--accent) 13%,#fff),#fff 60%)}}
    .rank-badge{{display:inline-flex;align-items:center;justify-content:center;min-width:29px;height:29px;background:var(--navy);color:#fff;font-weight:900}}
    tr.podium .rank-badge{{background:var(--accent);color:var(--navy)}}
    .metric{{display:grid;grid-template-columns:minmax(90px,1fr) 70px;align-items:center;gap:9px;min-width:170px}}
    .track{{height:10px;background:#e8efec;overflow:hidden}} .fill{{height:100%;background:linear-gradient(90deg,var(--accent-dark),var(--accent));transform-origin:left center;animation:fillGrow 1.15s .25s cubic-bezier(.2,.72,.2,1) both}}
    .metric-value{{font-weight:800;color:var(--accent-dark);text-align:right}}
    .footer{{padding:15px clamp(20px,3vw,34px);border-top:1px solid var(--border);background:var(--surface);color:var(--muted);font-size:12.5px}}
    .footer a{{color:var(--accent-dark)}} .empty{{padding:32px;text-align:center;color:var(--muted)}}
    .chart-wrap{{position:relative;overflow:auto;border:1px solid var(--border);background:linear-gradient(180deg,#fff,var(--surface));padding:24px}}
    .chart-title{{font-family:"Arial Narrow",Impact,Inter,sans-serif;font-size:21px;letter-spacing:.01em;color:var(--navy);margin:0 0 18px}}
    .chart-svg{{display:block;width:100%;height:auto;min-width:620px;overflow:visible}} .grid{{stroke:var(--border);stroke-width:1}}
    .axis-label{{fill:var(--muted);font-size:12px}} .cat-label{{fill:var(--text);font-size:12.5px;font-weight:650}}
    .value-label{{fill:var(--accent-dark);font-size:12px;font-weight:800}} .bar{{fill:url(#barGradient);transform-box:fill-box}}
    .bar-h{{transform-origin:left center;animation:barGrowX 1.15s .2s cubic-bezier(.2,.72,.2,1) both}}
    .bar-v{{transform-origin:center bottom;animation:barGrowY 1.15s .2s cubic-bezier(.2,.72,.2,1) both}}
    .line{{fill:none;stroke:var(--accent);stroke-width:4;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:2600;stroke-dashoffset:2600;animation:lineDraw 1.8s .18s cubic-bezier(.2,.72,.2,1) forwards}}
    .area{{fill:url(#areaGradient);animation:areaFade 1.05s .35s both}} .point{{fill:#fff;stroke:var(--accent);stroke-width:3;transform-box:fill-box;transform-origin:center;animation:pointPop .55s .8s both}}
    .slice{{stroke:#fff;stroke-width:3;transform-box:fill-box;transform-origin:center;animation:sliceIn .85s cubic-bezier(.2,.72,.2,1) both}}
    .legend{{display:flex;flex-wrap:wrap;gap:9px 18px;margin-top:15px;font-size:12px;color:var(--muted)}}
    .legend i{{display:inline-block;width:9px;height:9px;margin-right:6px}}
    @keyframes shellRise{{from{{opacity:0;transform:translateY(14px)}}to{{opacity:1;transform:none}}}}
    @keyframes fadeUp{{from{{opacity:0;transform:translateY(12px)}}to{{opacity:1;transform:none}}}}
    @keyframes rowRise{{from{{opacity:0;transform:translateY(8px)}}to{{opacity:1;transform:none}}}}
    @keyframes fillGrow{{from{{transform:scaleX(0)}}to{{transform:scaleX(1)}}}}
    @keyframes barGrowX{{from{{transform:scaleX(0);opacity:.35}}to{{transform:scaleX(1);opacity:1}}}}
    @keyframes barGrowY{{from{{transform:scaleY(0);opacity:.35}}to{{transform:scaleY(1);opacity:1}}}}
    @keyframes lineDraw{{to{{stroke-dashoffset:0}}}} @keyframes areaFade{{from{{opacity:0}}to{{opacity:1}}}}
    @keyframes pointPop{{from{{opacity:0;transform:scale(.2)}}to{{opacity:1;transform:scale(1)}}}}
    @keyframes sliceIn{{from{{opacity:0;transform:scale(.82)}}to{{opacity:1;transform:scale(1)}}}}
    @media(max-width:640px){{.hero{{padding:26px 20px 23px}} h1{{font-size:36px}} .content{{padding:14px}}
      .count{{width:100%;margin:0}} .chart-wrap{{padding:16px}}}}
    @media(prefers-reduced-motion:reduce){{*,*:before,*:after{{animation-duration:.01ms!important;animation-delay:0s!important;scroll-behavior:auto!important}}}}
    @media print{{.toolbar{{display:none}} .viz-shell{{box-shadow:none}} .table-wrap{{max-height:none;overflow:visible}}}}
    {preview_css}
    """


def hero_html(metadata: dict[str, str]) -> str:
    title = metadata.get("title", "").strip()
    subtitle = metadata.get("subtitle", "").strip()
    brand = metadata.get("brand", "Data Studio").strip()
    if not title and not subtitle:
        return ""
    return (
        '<header class="hero">'
        + (f'<div class="brand">{html.escape(smart_title_case(brand))}</div>' if brand else "")
        + (f"<h1>{html.escape(smart_title_case(title))}</h1>" if title else "")
        + (f'<p class="subtitle">{html.escape(subtitle)}</p>' if subtitle else "")
        + "</header>"
    )


def footer_html(metadata: dict[str, str]) -> str:
    footer = metadata.get("footer", "").strip()
    source = metadata.get("source", "").strip()
    source_url = metadata.get("source_url", "").strip()
    bits: list[str] = []
    if footer:
        bits.append(html.escape(footer))
    if source:
        source_text = f"Source: {html.escape(source)}"
        if source_url.startswith(("https://", "http://")):
            source_text = f'<a href="{html.escape(source_url, quote=True)}" target="_blank" rel="noopener">{source_text}</a>'
        bits.append(source_text)
    return f'<footer class="footer">{" · ".join(bits)}</footer>' if bits else ""


def table_markup(frame: pd.DataFrame, visual_key: str, metric_column: str | None) -> tuple[str, str]:
    columns = list(frame.columns)
    rank_column = next((c for c in columns if c.lower() in {"rank", "ranking", "position"}), None)
    metric_column = metric_column if metric_column in columns else None
    values = pd.to_numeric(frame[metric_column], errors="coerce") if metric_column else pd.Series(dtype=float)
    max_metric = float(values.max()) if not values.empty and values.notna().any() else 1.0

    heads = "".join(
        f'<th data-column="{i}" data-type="{"number" if pd.api.types.is_numeric_dtype(frame[c]) else "text"}">'
        f'{html.escape(smart_title_case(c))}<span class="sort">↕</span></th>'
        for i, c in enumerate(columns)
    )

    rows: list[str] = []
    search_records: list[dict[str, Any]] = []
    for row_index, (_, row) in enumerate(frame.iterrows()):
        rank_value = row.get(rank_column, row_index + 1) if rank_column else row_index + 1
        try:
            podium = int(float(rank_value)) <= 3
        except (TypeError, ValueError):
            podium = row_index < 3
        cells: list[str] = []
        raw_values: list[Any] = []
        for column in columns:
            raw = row[column]
            raw_values.append(None if pd.isna(raw) else json_safe(raw))
            display = animated_number_markup(raw, column, frame[column])
            if visual_key == "ranking_table" and column == rank_column:
                display = f'<span class="rank-badge">{display}</span>'
            elif visual_key == "ranking_table" and column == metric_column and not pd.isna(raw):
                width = max(0.0, min(100.0, float(raw) / max_metric * 100 if max_metric else 0))
                display = (
                    '<div class="metric"><div class="track">'
                    f'<div class="fill" style="width:{width:.2f}%"></div></div>'
                    f'<span class="metric-value">{display}</span></div>'
                )
            cells.append(f"<td>{display}</td>")
        rows.append(f'<tr class="{"podium" if podium and visual_key == "ranking_table" else ""}">{"".join(cells)}</tr>')
        search_records.append({"search": " ".join(str(v) for v in raw_values if v is not None).lower(), "values": raw_values})

    table = f'<table id="dataTable"><thead><tr>{heads}</tr></thead><tbody>{"".join(rows)}</tbody></table>'
    script = f"""
    const DATA = {safe_json_dumps(search_records)};
    const table = document.getElementById('dataTable');
    const tbody = table.querySelector('tbody');
    const originalRows = [...tbody.rows];
    const search = document.getElementById('tableSearch');
    const count = document.getElementById('rowCount');
    let sortColumn = -1, sortDirection = 1;
    function updateCount(){{const shown=[...tbody.rows].filter(r=>r.style.display!=='none').length;count.textContent=`${{shown}} of ${{DATA.length}} rows`;}}
    search.addEventListener('input',()=>{{const q=search.value.trim().toLowerCase();originalRows.forEach((row,i)=>{{row.style.display=DATA[i].search.includes(q)?'':'none'}});updateCount();}});
    table.querySelectorAll('th').forEach(th=>th.addEventListener('click',()=>{{
      const col=Number(th.dataset.column), type=th.dataset.type;sortDirection=sortColumn===col?-sortDirection:1;sortColumn=col;
      [...tbody.rows].sort((a,b)=>{{let x=a.cells[col].innerText.trim(),y=b.cells[col].innerText.trim();
        if(type==='number'){{x=Number(x.replace(/[^0-9.+-]/g,''))||0;y=Number(y.replace(/[^0-9.+-]/g,''))||0;return (x-y)*sortDirection}}
        return x.localeCompare(y,undefined,{{numeric:true}})*sortDirection;}}).forEach(r=>tbody.appendChild(r));
      table.querySelectorAll('.sort').forEach(s=>s.textContent='↕');th.querySelector('.sort').textContent=sortDirection===1?'▲':'▼';
    }}));
    document.getElementById('downloadCsv').addEventListener('click',()=>{{
      const rows=[...table.rows].filter((r,i)=>i===0||r.style.display!=='none');
      const csv=rows.map(r=>[...r.cells].map(c=>'"'+c.innerText.replaceAll('"','""')+'"').join(',')).join('\\n');
      const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{{type:'text/csv'}}));a.download='filtered-data.csv';a.click();URL.revokeObjectURL(a.href);
    }}); updateCount();
    """
    return table, script


def nice_number(value: float) -> str:
    absolute = abs(value)
    if absolute >= 1_000_000_000:
        return f"{value / 1_000_000_000:.1f}B".rstrip("0").rstrip(".")
    if absolute >= 1_000_000:
        return f"{value / 1_000_000:.1f}M".rstrip("0").rstrip(".")
    if absolute >= 1_000:
        return f"{value / 1_000:.1f}K".rstrip("0").rstrip(".")
    return f"{value:,.1f}".rstrip("0").rstrip(".")


def chart_data(frame: pd.DataFrame, label_column: str, value_column: str, second_value: str | None, limit: int) -> pd.DataFrame:
    selected = [label_column, value_column] + ([second_value] if second_value and second_value not in {label_column, value_column} else [])
    data = frame[selected].copy().head(limit)
    data[value_column] = pd.to_numeric(data[value_column], errors="coerce")
    if second_value and second_value in data:
        data[second_value] = pd.to_numeric(data[second_value], errors="coerce")
    return data.dropna(subset=[value_column])


def chart_markup(
    frame: pd.DataFrame,
    visual_key: str,
    label_column: str,
    value_column: str,
    second_value: str | None,
    limit: int,
) -> str:
    data = chart_data(frame, label_column, value_column, second_value, limit)
    if data.empty:
        return '<div class="empty">The selected columns do not contain chartable numeric data.</div>'

    labels = [str(v) for v in data[label_column].tolist()]
    values = [float(v) for v in data[value_column].tolist()]
    width = 1040
    height = max(430, 88 + 44 * len(values)) if visual_key == "bar_horizontal" else 540
    left, right, top, bottom = 92, 50, 40, 80
    plot_w, plot_h = width - left - right, height - top - bottom
    max_value = max(values) if values else 1
    min_value = min(0, min(values) if values else 0)
    span = max(max_value - min_value, 1e-9)
    parts = [
        f'<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img" aria-label="{html.escape(smart_title_case(value_column))} by {html.escape(smart_title_case(label_column))}">',
        '<defs><linearGradient id="barGradient" x1="0" x2="1"><stop offset="0" stop-color="var(--accent-dark)"/><stop offset="1" stop-color="var(--accent)"/></linearGradient>',
        '<linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="var(--accent)" stop-opacity=".45"/><stop offset="1" stop-color="var(--accent)" stop-opacity=".03"/></linearGradient></defs>',
    ]

    if visual_key == "bar_horizontal":
        gap = plot_h / max(len(values), 1)
        bar_h = min(24, gap * 0.58)
        label_x = left - 14
        for i, (label, value) in enumerate(zip(labels, values)):
            y = top + i * gap + (gap - bar_h) / 2
            bar_w = max(1, (value - min_value) / span * plot_w)
            parts.append(f'<text class="cat-label" x="{label_x}" y="{y + bar_h * .72:.1f}" text-anchor="end">{html.escape(label[:28])}</text>')
            parts.append(f'<rect class="bar bar-h" x="{left}" y="{y:.1f}" width="{bar_w:.1f}" height="{bar_h:.1f}" rx="2"><title>{html.escape(label)}: {html.escape(format_value(value, value_column, data[value_column]))}</title></rect>')
            parts.append(f'<text class="value-label" data-animate-number="{value}" data-number-kind="compact" x="{min(left + bar_w + 8, width - right)}" y="{y + bar_h * .72:.1f}">{html.escape(nice_number(value))}</text>')
    elif visual_key == "bar_vertical":
        gap = plot_w / max(len(values), 1)
        bar_w = min(56, gap * 0.62)
        for i, (label, value) in enumerate(zip(labels, values)):
            bar_h = max(1, (value - min_value) / span * plot_h)
            x = left + i * gap + (gap - bar_w) / 2
            y = top + plot_h - bar_h
            parts.append(f'<rect class="bar bar-v" x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{bar_h:.1f}" rx="3"><title>{html.escape(label)}: {html.escape(format_value(value, value_column, data[value_column]))}</title></rect>')
            parts.append(f'<text class="value-label" data-animate-number="{value}" data-number-kind="compact" x="{x + bar_w/2:.1f}" y="{max(16,y-8):.1f}" text-anchor="middle">{html.escape(nice_number(value))}</text>')
            parts.append(f'<text class="cat-label" x="{x + bar_w/2:.1f}" y="{top + plot_h + 22:.1f}" text-anchor="middle">{html.escape(label[:14])}</text>')
    elif visual_key in {"line", "area"}:
        count = max(len(values), 1)
        points: list[tuple[float, float]] = []
        for i, value in enumerate(values):
            x = left + (plot_w * i / max(count - 1, 1))
            y = top + plot_h - ((value - min_value) / span * plot_h)
            points.append((x, y))
        point_str = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
        if visual_key == "area":
            polygon = f"{left},{top + plot_h} {point_str} {points[-1][0]:.1f},{top + plot_h}"
            parts.append(f'<polygon class="area" points="{polygon}"/>')
        parts.append(f'<polyline class="line" points="{point_str}"/>')
        for (x, y), label, value in zip(points, labels, values):
            parts.append(f'<circle class="point" cx="{x:.1f}" cy="{y:.1f}" r="5"><title>{html.escape(label)}: {html.escape(format_value(value, value_column, data[value_column]))}</title></circle>')
            parts.append(f'<text class="cat-label" x="{x:.1f}" y="{top + plot_h + 24}" text-anchor="middle">{html.escape(label[:12])}</text>')
    elif visual_key == "donut":
        positive = [max(0, v) for v in values]
        total = sum(positive)
        if total <= 0:
            return '<div class="empty">A donut chart requires positive values.</div>'
        colors = ["var(--accent)", "var(--navy)", "#5b8def", "#f2b84b", "#e46b6b", "#8f73d8", "#43a6a1", "#8a9a5b"]
        cx, cy, radius, stroke = 370, 248, 150, 62
        circumference = 2 * math.pi * radius
        offset = 0.0
        for i, (label, value) in enumerate(zip(labels, positive)):
            length = value / total * circumference
            parts.append(f'<circle cx="{cx}" cy="{cy}" r="{radius}" fill="none" stroke="{colors[i % len(colors)]}" stroke-width="{stroke}" stroke-dasharray="{length:.3f} {circumference-length:.3f}" stroke-dashoffset="{-offset:.3f}" transform="rotate(-90 {cx} {cy})" class="slice"><title>{html.escape(label)}: {value/total*100:.1f}%</title></circle>')
            offset += length
        parts.append(f'<text x="{cx}" y="{cy-3}" text-anchor="middle" data-animate-number="{total}" data-number-kind="compact" style="font-size:30px;font-weight:900;fill:var(--navy)">{html.escape(nice_number(total))}</text>')
        parts.append(f'<text x="{cx}" y="{cy+22}" text-anchor="middle" class="axis-label">Total</text>')
        legend_x, legend_y = 650, 95
        for i, (label, value) in enumerate(zip(labels[:8], positive[:8])):
            y = legend_y + i * 38
            parts.append(f'<rect x="{legend_x}" y="{y-11}" width="12" height="12" fill="{colors[i % len(colors)]}"/>')
            parts.append(f'<text class="cat-label" x="{legend_x+20}" y="{y}">{html.escape(label[:24])}</text>')
            parts.append(f'<text class="value-label" data-animate-number="{value/total*100}" data-number-kind="percent" x="{width-right}" y="{y}" text-anchor="end">{value/total*100:.1f}%</text>')
    elif visual_key == "scatter":
        if not second_value or second_value not in data:
            return '<div class="empty">Choose a second numeric measure for a scatter plot.</div>'
        ys = pd.to_numeric(data[second_value], errors="coerce")
        valid = ys.notna()
        xs = data.loc[valid, value_column].astype(float).tolist()
        y_values = ys[valid].astype(float).tolist()
        point_labels = data.loc[valid, label_column].astype(str).tolist()
        if not xs:
            return '<div class="empty">The selected measures contain no paired numeric values.</div>'
        min_x, max_x = min(xs), max(xs); min_y, max_y = min(y_values), max(y_values)
        span_x, span_y = max(max_x-min_x,1e-9), max(max_y-min_y,1e-9)
        for label, x_value, y_value in zip(point_labels, xs, y_values):
            x = left + (x_value-min_x)/span_x*plot_w
            y = top + plot_h - (y_value-min_y)/span_y*plot_h
            parts.append(f'<circle class="point" cx="{x:.1f}" cy="{y:.1f}" r="7"><title>{html.escape(label)} — {html.escape(smart_title_case(value_column))}: {x_value:g}; {html.escape(smart_title_case(second_value))}: {y_value:g}</title></circle>')
        parts.append(f'<text class="axis-label" x="{left+plot_w/2}" y="{height-20}" text-anchor="middle">{html.escape(smart_title_case(value_column))}</text>')
        parts.append(f'<text class="axis-label" x="22" y="{top+plot_h/2}" text-anchor="middle" transform="rotate(-90 22 {top+plot_h/2})">{html.escape(smart_title_case(second_value))}</text>')

    parts.append("</svg>")
    return "".join(parts)


def generate_html(
    frame: pd.DataFrame,
    metadata: dict[str, str],
    visual_key: str = "editorial_table",
    theme_name: str = "Action Green",
    label_column: str | None = None,
    value_column: str | None = None,
    second_value: str | None = None,
    metric_column: str | None = None,
    row_limit: int = 20,
    preview: bool = False,
) -> str:
    metadata = {k: str(v) for k, v in metadata.items() if v is not None}
    theme = merge_theme(theme_name, metadata.get("accent", ""))
    visual = next((v for v in VISUALS if v["key"] == visual_key), VISUALS[0])
    compact = visual_key == "compact_table"
    css = base_css(theme, compact=compact, preview=preview)
    render_frame = frame.head(6).copy() if preview else frame
    title = metadata.get("title", "Data visualisation")
    hero = hero_html(metadata)
    footer = footer_html(metadata)

    if visual["kind"] == "table":
        table, table_script = table_markup(render_frame, visual_key, metric_column)
        body = f"""
        <main class="content">
          <div class="toolbar">
            <label class="search"><input id="tableSearch" type="search" placeholder="Search this table" aria-label="Search this table"></label>
            <button class="btn" id="downloadCsv" type="button">Download Data</button>
            <span class="count" id="rowCount"></span>
          </div>
          <div class="table-wrap">{table}</div>
        </main>
        """
        script = table_script + animation_script()
    else:
        numeric = numeric_columns(render_frame)
        label_column = label_column if label_column in render_frame.columns else str(render_frame.columns[0])
        value_column = value_column if value_column in numeric else (numeric[0] if numeric else str(frame.columns[-1]))
        chart = chart_markup(render_frame, visual_key, label_column, value_column, second_value, min(row_limit, 8) if preview else row_limit)
        chart_heading = f"{smart_title_case(value_column)} by {smart_title_case(label_column)}"
        body = f'<main class="content"><div class="chart-wrap"><h2 class="chart-title">{html.escape(chart_heading)}</h2>{chart}</div></main>'
        script = animation_script()

    return f"""<!doctype html>
<html lang="en">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="{html.escape(metadata.get('subtitle','Interactive data visualisation'), quote=True)}">
<title>{html.escape(smart_title_case(title))}</title><style>{css}</style></head>
<body><article class="viz-shell">{hero}{body}{footer}</article><script>{script}</script></body></html>"""


def sample_dataset() -> Dataset:
    frame = pd.DataFrame(
        {
            "Rank": [1, 2, 3, 4, 5, 6, 7, 8],
            "State": ["California", "Arizona", "Utah", "Texas", "Colorado", "New Mexico", "Oregon", "Nevada"],
            "Overall Score": [92.4, 88.6, 85.1, 82.8, 81.4, 78.9, 76.2, 74.7],
            "Implied Probability": [12.8, 11.9, 11.1, 10.4, 10.0, 9.2, 8.0, 6.6],
            "Venues": [1842, 978, 734, 1560, 921, 502, 844, 611],
        }
    )
    metadata = {
        "title": "The 2026 State Sports Index",
        "subtitle": "A demonstration dataset showing how the builder turns an ordinary data file into an editorial interactive.",
        "footer": "Scores shown for demonstration only.",
        "source": "Example data",
        "brand": "Action Network · Data Studio",
    }
    return Dataset(frame, metadata)


def run_app() -> None:
    import streamlit as st
    import streamlit.components.v1 as components

    st.set_page_config(page_title="Table & Chart Builder", page_icon="▦", layout="wide")
    st.markdown(
        """<style>
        .stApp{background:#f2f6f4}.block-container{max-width:1480px;padding-top:1.5rem}
        [data-testid="stSidebar"]{background:#101820;color:#fff}
        [data-testid="stSidebar"] label,[data-testid="stSidebar"] p{color:#e9f4ef!important}
        div[data-testid="stFileUploader"]{background:#fff;border:1px solid #dfe7e3;padding:12px}
        div[data-testid="stFileUploader"] small{display:none!important}
        .carousel-card{padding:17px 22px;border:1px solid #dfe7e3;border-top:4px solid #00b67a;border-radius:7px;background:#fff;box-shadow:0 14px 34px rgba(16,24,32,.09);animation:carouselRise .55s cubic-bezier(.2,.72,.2,1) both}
        .carousel-card h3{margin:0 0 5px;color:#101820;font-size:1.55rem}.carousel-card p{margin:0;color:#66747d}.dots{text-align:center;color:#a9b5af;letter-spacing:5px;margin-top:5px}.dots b{color:#00b67a}
        div[data-testid="stIFrame"]{border:1px solid #dfe7e3;border-radius:7px;overflow:hidden;box-shadow:0 12px 30px rgba(16,24,32,.08)}
        @keyframes carouselRise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
        </style>""",
        unsafe_allow_html=True,
    )

    st.title("Automatic Table & Chart Builder")
    st.caption("Upload a data file, choose a sleek layout, preview it and download a standalone web page.")

    uploaded = st.file_uploader("Upload a Data File", type=["csv", "tsv", "txt", "xls", "xlsx", "xlsm"])
    use_demo = st.checkbox("Use Demonstration Data", value=uploaded is None)

    if uploaded is None and not use_demo:
        st.info("Upload a file or turn on demonstration data to begin.")
        return

    try:
        dataset = sample_dataset() if uploaded is None else parse_uploaded_file(uploaded.getvalue(), uploaded.name)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        return

    frame = dataset.frame
    metadata = dataset.metadata

    with st.sidebar:
        st.header("Content")
        title = st.text_input("Heading", value=metadata.get("title", ""), placeholder="Optional")
        subtitle = st.text_area("Subtitle", value=metadata.get("subtitle", ""), placeholder="Optional", height=90)
        footer = st.text_area("Footer Note", value=metadata.get("footer", ""), placeholder="Optional", height=75)
        source = st.text_input("Source Label", value=metadata.get("source", ""), placeholder="Optional")
        source_url = st.text_input("Source Link", value=metadata.get("source_url", ""), placeholder="Optional")
        brand = st.text_input("Brand Line", value=metadata.get("brand", "Action Network · Data Studio"))
        st.header("Appearance")
        theme_name = st.selectbox("Colour Theme", list(THEMES), index=0)

    final_metadata = {
        "title": title,
        "subtitle": subtitle,
        "footer": footer,
        "source": source,
        "source_url": source_url,
        "brand": brand,
        "accent": metadata.get("accent", ""),
    }
    numerics = numeric_columns(frame)
    categoricals = categorical_columns(frame)
    label_default = categoricals[0] if categoricals else str(frame.columns[0])
    value_default = numerics[0] if numerics else str(frame.columns[-1])
    second_default = numerics[1] if len(numerics) > 1 else None

    if "visual_index" not in st.session_state:
        st.session_state.visual_index = 0
    previous, card, nxt = st.columns([1, 6, 1])
    with previous:
        if st.button("←", use_container_width=True, help="Previous Design"):
            st.session_state.visual_index = (st.session_state.visual_index - 1) % len(VISUALS)
            st.session_state.visual_jump = st.session_state.visual_index
    with nxt:
        if st.button("→", use_container_width=True, help="Next Design"):
            st.session_state.visual_index = (st.session_state.visual_index + 1) % len(VISUALS)
            st.session_state.visual_jump = st.session_state.visual_index
    selected = VISUALS[st.session_state.visual_index]
    with card:
        st.markdown(f'<div class="carousel-card"><h3>{selected["name"]}</h3><p>{selected["description"]}</p></div>', unsafe_allow_html=True)
        miniature_html = generate_html(
            frame,
            final_metadata,
            visual_key=selected["key"],
            theme_name=theme_name,
            label_column=label_default,
            value_column=value_default,
            second_value=second_default,
            metric_column=value_default if value_default in numerics else None,
            row_limit=8,
            preview=True,
        )
        components.html(miniature_html, height=350, scrolling=False)
        dots = " ".join("<b>●</b>" if i == st.session_state.visual_index else "●" for i in range(len(VISUALS)))
        st.markdown(f'<div class="dots">{dots}</div>', unsafe_allow_html=True)

    def jump_to_visual() -> None:
        st.session_state.visual_index = st.session_state.visual_jump

    st.selectbox(
        "Jump Directly to a Design",
        range(len(VISUALS)),
        index=st.session_state.visual_index,
        format_func=lambda i: VISUALS[i]["name"],
        key="visual_jump",
        on_change=jump_to_visual,
    )
    selected = VISUALS[st.session_state.visual_index]

    control_cols = st.columns(4)
    label_column = control_cols[0].selectbox(
        "Category / Label",
        list(frame.columns),
        index=list(frame.columns).index(label_default),
        format_func=smart_title_case,
    )
    value_column = control_cols[1].selectbox("Primary Measure", numerics or list(frame.columns), index=0, format_func=smart_title_case)
    second_options = ["None"] + numerics
    second_pick = control_cols[2].selectbox("Second Measure", second_options, index=0, format_func=smart_title_case)
    second_value = None if second_pick == "None" else second_pick
    metric_column = control_cols[3].selectbox(
        "Ranking Data Bar",
        ["None"] + numerics,
        index=1 if numerics else 0,
        format_func=smart_title_case,
    )
    metric_column = None if metric_column == "None" else metric_column
    row_limit = st.slider("Maximum Chart Categories", 3, 30, min(12, max(3, len(frame))))
    output_html = generate_html(
        frame,
        final_metadata,
        visual_key=selected["key"],
        theme_name=theme_name,
        label_column=label_column,
        value_column=value_column,
        second_value=second_value,
        metric_column=metric_column,
        row_limit=row_limit,
    )

    preview_tab, data_tab, help_tab = st.tabs(["Preview", "Data", "File Format"])
    with preview_tab:
        components.html(output_html, height=820, scrolling=True)
    with data_tab:
        display_frame = frame.rename(columns={column: smart_title_case(column) for column in frame.columns})
        st.dataframe(display_frame, use_container_width=True, height=520)
        st.caption(f"{len(frame):,} rows · {len(frame.columns)} columns")
    with help_tab:
        st.code(
            """# title: Optional heading
# subtitle: Optional explanatory line
# footer: Optional note shown below the visual
# source: Optional source label
# source_url: https://example.com/optional-source
# brand: Action Network · Data Studio

Rank,State,Score,Implied Probability
1,California,92.4,12.8
2,Arizona,88.6,11.9
3,Utah,85.1,11.1""",
            language="text",
        )
        st.write("All `#` metadata lines are optional. A normal data file beginning with its header row works too.")

    filename = safe_id(title or "data-visualisation") + ".html"
    st.download_button("Download Selected Web Page", output_html.encode("utf-8"), filename, "text/html", type="primary")
    st.download_button("Download Cleaned Data", frame.to_csv(index=False).encode("utf-8"), "cleaned-data.csv", "text/csv")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--demo-output", type=Path, help="Write a demonstration web page and exit.")
    args = parser.parse_args()
    if args.demo_output:
        sample = sample_dataset()
        args.demo_output.write_text(
            generate_html(
                sample.frame,
                sample.metadata,
                visual_key="ranking_table",
                metric_column="Overall Score",
            ),
            encoding="utf-8",
        )
        print(f"Wrote {args.demo_output}")
    else:
        run_app()


if __name__ == "__main__":
    main()
