
import io
import re
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st
from bs4 import BeautifulSoup
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

TARGET_REASONS = [
    "RDF DESSERT",
    "RDF CAFE",
    "RDF BOISSON",
    "RDF ELSASSICH",
]

LOT_TO_REASON = {
    "dessert": "RDF DESSERT",
    "cafe": "RDF CAFE",
    "boisson": "RDF BOISSON",
    "formule elsassich": "RDF ELSASSICH",
}


def normalize_amount(value: str) -> float:
    value = str(value).replace("\xa0", " ").strip()
    value = value.replace(" ", "").replace(",", ".")
    value = re.sub(r"[^0-9.\-]", "", value)
    return float(value) if value not in {"", "-", ".", "-."} else 0.0


def format_eur(value: float) -> str:
    return f"{value:,.2f} €".replace(",", " ").replace(".", ",")


def format_pct(value: float) -> str:
    return f"{value:,.1f} %".replace(",", " ").replace(".", ",")


def safe_date_str(value) -> str:
    if pd.isna(value):
        return "-"
    return pd.to_datetime(value).strftime("%Y-%m-%d")


def compute_roi(total_to_pay: float, discount_total: float) -> float:
    if discount_total == 0:
        return 0.0
    return (total_to_pay / discount_total) * 100.0


def read_people_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=None, engine="python")
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)
    else:
        raise ValueError("Format non supporté pour le fichier personnes.")

    df.columns = [str(c).strip().upper() for c in df.columns]
    rename_map = {
        "PRENOM": "PRENOM",
        "PRÉNOM": "PRENOM",
        "FIRSTNAME": "PRENOM",
        "NOM": "NOM",
        "LASTNAME": "NOM",
        "LOT": "LOT",
        "DATE": "DATE",
        "ADDED_TIME": "DATE",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    required = {"NOM", "PRENOM", "LOT", "DATE"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans le fichier personnes : {', '.join(sorted(missing))}")

    df = df[list(required)].copy()
    df["PRENOM"] = df["PRENOM"].astype(str).str.strip()
    df["NOM"] = df["NOM"].astype(str).str.strip()
    df["LOT"] = df["LOT"].astype(str).str.strip().str.lower()
    df["DATE"] = pd.to_datetime(df["DATE"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["DATE"]).copy()
    df["DATE_JOUR"] = pd.to_datetime(df["DATE"].dt.date)
    df["LOT_LABEL"] = df["LOT"].str.title()
    return df


def extract_text_from_html(html_bytes: bytes) -> str:
    raw = html_bytes.decode("utf-8", errors="ignore")
    soup = BeautifulSoup(raw, "html.parser")
    text = soup.get_text("\n")
    text = text.replace("\xa0", " ")
    text = text.replace("\r", "\n")
    lines = [line.rstrip() for line in text.splitlines()]
    return "\n".join(lines)


def split_ticket_blocks(text: str) -> List[str]:
    starts = [m.start() for m in re.finditer(r"(?m)^Table:", text)]
    blocks = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        blocks.append(text[start:end].strip())
    return blocks


def parse_discount_lines(block: str) -> List[Dict]:
    matches = []
    pattern = re.compile(
        r"(?m)^\s*(?P<qty>-?\d+(?:[.,]\d+)?)\s+"
        r"(?P<designation>.+?)\s+"
        r"(?P<discount>-?\d+(?:[.,]\d+)?)\s+"
        r"(?P<reason>RDF DESSERT|RDF CAFE|RDF BOISSON|RDF ELSASSICH)\s+"
        r"(?P<when>\d{2}\.\d{2}\.\d{4}\s+\d{2}:)"
    )
    for m in pattern.finditer(block):
        matches.append(
            {
                "qty": normalize_amount(m.group("qty")),
                "designation": m.group("designation").strip(),
                "discount_amount": normalize_amount(m.group("discount")),
                "reason": m.group("reason").strip(),
                "action_time_raw": m.group("when").strip(),
            }
        )
    return matches


def parse_total_to_pay(block: str) -> float:
    m = re.search(r"TOTAL TO PAY:\s*([0-9,.-]+)", block)
    return normalize_amount(m.group(1)) if m else 0.0


def parse_ticket_opened_date(block: str):
    m = re.search(r"Table opened by:.*?@\s*(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}", block, flags=re.DOTALL)
    if m:
        dt = pd.to_datetime(m.group(1), dayfirst=True, errors="coerce")
        return pd.to_datetime(dt.date()) if pd.notna(dt) else pd.NaT

    fallback = re.search(r"(\d{2}\.\d{2}\.\d{4})", block)
    if fallback:
        dt = pd.to_datetime(fallback.group(1), dayfirst=True, errors="coerce")
        return pd.to_datetime(dt.date()) if pd.notna(dt) else pd.NaT
    return pd.NaT


def parse_note_number(block: str) -> str:
    m = re.search(r"Note number:(\d+)", block)
    return m.group(1) if m else ""


def parse_table_number(block: str) -> str:
    m = re.search(r"^Table:([^\s]+)", block, flags=re.MULTILINE)
    return m.group(1).strip() if m else ""


def parse_number_of_covers(block: str) -> int:
    m = re.search(r"Number of covers:(\d+)", block)
    return int(m.group(1)) if m else 0


def parse_html_tickets(uploaded_html) -> Tuple[pd.DataFrame, pd.DataFrame]:
    text = extract_text_from_html(uploaded_html.getvalue())
    blocks = split_ticket_blocks(text)

    ticket_rows = []
    discount_rows = []
    source_name = uploaded_html.name

    for block in blocks:
        discounts = parse_discount_lines(block)
        if not discounts:
            continue

        note_number = parse_note_number(block)
        table_number = parse_table_number(block)
        ticket_date = parse_ticket_opened_date(block)
        total_to_pay = parse_total_to_pay(block)
        covers = parse_number_of_covers(block)
        ticket_key = f"{source_name}__{note_number or table_number}"

        ticket_rows.append(
            {
                "ticket_key": ticket_key,
                "source_file": source_name,
                "note_number": note_number,
                "table_number": table_number,
                "date": ticket_date,
                "number_of_covers": covers,
                "total_to_pay": total_to_pay,
                "discount_count": len(discounts),
                "discount_total": sum(d["discount_amount"] for d in discounts),
                "discount_types": ", ".join(sorted({d["reason"] for d in discounts})),
            }
        )

        for d in discounts:
            discount_rows.append(
                {
                    "ticket_key": ticket_key,
                    "source_file": source_name,
                    "note_number": note_number,
                    "table_number": table_number,
                    "date": ticket_date,
                    "number_of_covers": covers,
                    "total_to_pay": total_to_pay,
                    **d,
                }
            )

    return pd.DataFrame(ticket_rows), pd.DataFrame(discount_rows)


def parse_multiple_html_tickets(uploaded_html_files) -> Tuple[pd.DataFrame, pd.DataFrame]:
    ticket_frames = []
    discount_frames = []

    for html_file in uploaded_html_files:
        df_tickets_one, df_discounts_one = parse_html_tickets(html_file)
        if not df_tickets_one.empty:
            ticket_frames.append(df_tickets_one)
        if not df_discounts_one.empty:
            discount_frames.append(df_discounts_one)

    df_tickets = pd.concat(ticket_frames, ignore_index=True) if ticket_frames else pd.DataFrame()
    df_discounts = pd.concat(discount_frames, ignore_index=True) if discount_frames else pd.DataFrame()

    if not df_tickets.empty:
        df_tickets = df_tickets.drop_duplicates(subset=["ticket_key"]).copy()
    if not df_discounts.empty:
        df_discounts = df_discounts.drop_duplicates(
            subset=["ticket_key", "reason", "designation", "discount_amount", "action_time_raw"]
        ).copy()

    return df_tickets, df_discounts


def build_people_dashboard(df_people: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    daily_lot = (
        df_people.groupby(["DATE_JOUR", "LOT_LABEL"], dropna=False)
        .size()
        .reset_index(name="NB_PERSONNES")
        .sort_values(["DATE_JOUR", "LOT_LABEL"])
    )

    pivot = (
        daily_lot.pivot_table(
            index="DATE_JOUR",
            columns="LOT_LABEL",
            values="NB_PERSONNES",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    pivot.columns.name = None
    value_cols = [c for c in pivot.columns if c != "DATE_JOUR"]
    pivot["TOTAL"] = pivot[value_cols].sum(axis=1) if value_cols else 0
    return daily_lot, pivot


def build_html_covers_dashboard(df_tickets: pd.DataFrame) -> pd.DataFrame:
    if df_tickets.empty:
        return pd.DataFrame(columns=["DATE_TICKET", "PERSONNES_RESTAURANT"])
    result = (
        df_tickets.groupby("date", dropna=False)["number_of_covers"]
        .sum()
        .reset_index()
        .rename(columns={"date": "DATE_TICKET", "number_of_covers": "PERSONNES_RESTAURANT"})
        .sort_values("DATE_TICKET")
    )
    return result


def build_discount_summary(df_discounts: pd.DataFrame, df_tickets: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df_discounts.groupby("reason", dropna=False)
        .agg(
            nb_remises=("reason", "size"),
            montant_total_remises=("discount_amount", "sum"),
            nb_tickets=("ticket_key", "nunique"),
        )
        .reset_index()
        .sort_values("reason")
    )

    ticket_map = df_discounts.groupby("reason")["ticket_key"].apply(list).to_dict()
    totals = []
    rois = []
    for reason in summary["reason"]:
        ticket_keys = set(ticket_map.get(reason, []))
        total_to_pay = df_tickets[df_tickets["ticket_key"].isin(ticket_keys)]["total_to_pay"].sum()
        discount_total = summary.loc[summary["reason"] == reason, "montant_total_remises"].iloc[0]
        totals.append(total_to_pay)
        rois.append(compute_roi(total_to_pay, discount_total))

    summary["montant_total_to_pay_tickets"] = totals
    summary["roi_pct"] = rois
    return summary


def build_reconciliation(df_people: pd.DataFrame, df_discounts: pd.DataFrame) -> pd.DataFrame:
    expected = (
        df_people.assign(reason=df_people["LOT"].map(LOT_TO_REASON))
        .dropna(subset=["reason"])
        .groupby(["DATE_JOUR", "reason"])
        .size()
        .reset_index(name="nb_personnes_attendues")
        .rename(columns={"DATE_JOUR": "date"})
    )

    actual = (
        df_discounts.groupby(["date", "reason"])
        .size()
        .reset_index(name="nb_remises_trouvees")
    )

    merged = expected.merge(actual, on=["date", "reason"], how="outer")
    merged["nb_personnes_attendues"] = merged["nb_personnes_attendues"].fillna(0).astype(int)
    merged["nb_remises_trouvees"] = merged["nb_remises_trouvees"].fillna(0).astype(int)
    merged["ecart"] = merged["nb_remises_trouvees"] - merged["nb_personnes_attendues"]
    return merged.sort_values(["date", "reason"])


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = name[:31]
            export_df = df.copy()
            for col in export_df.columns:
                if pd.api.types.is_datetime64_any_dtype(export_df[col]):
                    export_df[col] = export_df[col].dt.strftime("%Y-%m-%d")
            export_df.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output.read()


def df_to_table_data(df: pd.DataFrame, max_rows: int = 35) -> List[List[str]]:
    if df.empty:
        return [["Aucune donnée"]]

    export_df = df.head(max_rows).copy()
    for col in export_df.columns:
        if pd.api.types.is_datetime64_any_dtype(export_df[col]):
            export_df[col] = export_df[col].dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_float_dtype(export_df[col]):
            export_df[col] = export_df[col].map(lambda x: f"{x:,.2f}".replace(",", " ").replace(".", ","))
        else:
            export_df[col] = export_df[col].astype(str)

    return [list(export_df.columns)] + export_df.values.tolist()


def styled_pdf_table(data: List[List[str]], col_widths=None) -> Table:
    table = Table(data, colWidths=col_widths, repeatRows=1 if len(data) > 1 else 0)
    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("LEADING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ])
    table.setStyle(style)
    return table


def build_pdf_report(
    df_people: pd.DataFrame,
    df_tickets: pd.DataFrame,
    df_discounts: pd.DataFrame,
    daily_pivot: pd.DataFrame,
    html_covers_daily: pd.DataFrame,
    discount_summary: pd.DataFrame,
    reconciliation: pd.DataFrame,
    selected_sources: List[str],
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Rapport Analyse RDF Flams",
    )

    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.textColor = colors.HexColor("#0f172a")
    title_style.fontName = "Helvetica-Bold"

    h_style = styles["Heading2"]
    h_style.textColor = colors.HexColor("#1d4ed8")
    h_style.spaceBefore = 8
    h_style.spaceAfter = 6

    body = ParagraphStyle(
        "BodyCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        spaceAfter=5,
    )

    total_remises = df_discounts["discount_amount"].sum() if not df_discounts.empty else 0.0
    total_to_pay = df_tickets["total_to_pay"].sum() if not df_tickets.empty else 0.0
    roi_pct = compute_roi(total_to_pay, total_remises)

    period_values = []
    if not df_people.empty:
        period_values.extend(df_people["DATE_JOUR"].dropna().tolist())
    if not df_tickets.empty:
        period_values.extend(df_tickets["date"].dropna().tolist())
    period_label = "-"
    if period_values:
        period_label = f"{min(period_values).strftime('%Y-%m-%d')} - {max(period_values).strftime('%Y-%m-%d')}"

    source_label = ", ".join(selected_sources) if selected_sources else "Tous les fichiers HTML"

    story = [
        Paragraph("Rapport Analyse RDF Flams", title_style),
        Spacer(1, 0.15 * cm),
        Paragraph(f"Période analysée : <b>{period_label}</b>", body),
        Paragraph(f"Fichiers HTML inclus : <b>{source_label}</b>", body),
        Paragraph(
            f"Personnes Excel : <b>{len(df_people)}</b> | Tickets RDF : <b>{df_tickets['ticket_key'].nunique() if not df_tickets.empty else 0}</b> | "
            f"Remises RDF : <b>{len(df_discounts)}</b> | Montant remises : <b>{format_eur(total_remises)}</b> | "
            f"Total to pay : <b>{format_eur(total_to_pay)}</b> | ROI : <b>{format_pct(roi_pct)}</b>",
            body,
        ),
        Spacer(1, 0.25 * cm),
        Paragraph("Synthèse des remises RDF", h_style),
        styled_pdf_table(df_to_table_data(discount_summary), col_widths=[5 * cm, 2.2 * cm, 3.4 * cm, 2.2 * cm, 3.8 * cm, 2.5 * cm]),
        Spacer(1, 0.3 * cm),
        Paragraph("Personnes par jour et par lot - fichier Excel", h_style),
        styled_pdf_table(df_to_table_data(daily_pivot), col_widths=None),
        Spacer(1, 0.3 * cm),
        Paragraph("Personnes par jour en restaurant - tickets HTML RDF", h_style),
        styled_pdf_table(df_to_table_data(html_covers_daily), col_widths=[4 * cm, 4.5 * cm]),
        PageBreak(),
        Paragraph("Contrôle attendu vs trouvé", h_style),
        styled_pdf_table(df_to_table_data(reconciliation), col_widths=None),
        Spacer(1, 0.3 * cm),
        Paragraph("Tickets RDF", h_style),
        styled_pdf_table(df_to_table_data(df_tickets.sort_values(["date", "note_number"])), col_widths=None),
        Spacer(1, 0.3 * cm),
        Paragraph("Détails des remises RDF", h_style),
        styled_pdf_table(df_to_table_data(df_discounts.sort_values(["date", "note_number", "reason"])), col_widths=None),
    ]

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


def apply_filters(
    df_people: pd.DataFrame,
    df_tickets: pd.DataFrame,
    df_discounts: pd.DataFrame,
    selected_dates,
    selected_reasons,
    selected_lots,
    selected_sources,
    note_query: str,
):
    people_filtered = df_people.copy()
    tickets_filtered = df_tickets.copy()
    discounts_filtered = df_discounts.copy()

    if selected_dates:
        selected_dates = pd.to_datetime(list(selected_dates))
        people_filtered = people_filtered[people_filtered["DATE_JOUR"].isin(selected_dates)]
        tickets_filtered = tickets_filtered[tickets_filtered["date"].isin(selected_dates)]
        discounts_filtered = discounts_filtered[discounts_filtered["date"].isin(selected_dates)]

    if selected_reasons:
        discounts_filtered = discounts_filtered[discounts_filtered["reason"].isin(selected_reasons)]
        allowed_keys = set(discounts_filtered["ticket_key"].astype(str))
        tickets_filtered = tickets_filtered[tickets_filtered["ticket_key"].astype(str).isin(allowed_keys)]

    if selected_lots:
        people_filtered = people_filtered[people_filtered["LOT_LABEL"].isin(selected_lots)]

    if selected_sources:
        tickets_filtered = tickets_filtered[tickets_filtered["source_file"].isin(selected_sources)]
        allowed_keys = set(tickets_filtered["ticket_key"].astype(str))
        discounts_filtered = discounts_filtered[discounts_filtered["ticket_key"].astype(str).isin(allowed_keys)]

    if note_query:
        nq = note_query.strip().lower()
        tickets_filtered = tickets_filtered[
            tickets_filtered["note_number"].astype(str).str.lower().str.contains(nq, na=False)
        ]
        allowed_keys = set(tickets_filtered["ticket_key"].astype(str))
        discounts_filtered = discounts_filtered[discounts_filtered["ticket_key"].astype(str).isin(allowed_keys)]

    return people_filtered, tickets_filtered, discounts_filtered


def render_header():
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.2rem;
                padding-bottom: 2rem;
            }
            .main-card {
                background: linear-gradient(135deg, rgba(15, 23, 42, 0.96) 0%, rgba(30, 41, 59, 0.96) 100%);
                color: #ffffff;
                padding: 1.35rem 1.5rem;
                border-radius: 18px;
                margin-bottom: 1rem;
                border: 1px solid rgba(255,255,255,0.08);
            }
            .subtle {
                color: rgba(255,255,255,0.86);
                margin-top: 0.35rem;
                font-size: 0.95rem;
            }
            div[data-testid="stMetric"] {
                background: var(--secondary-background-color);
                border: 1px solid rgba(128,128,128,0.25);
                padding: 0.8rem 1rem;
                border-radius: 14px;
            }
            div[data-testid="stMetric"] label,
            div[data-testid="stMetric"] [data-testid="stMetricLabel"],
            div[data-testid="stMetric"] [data-testid="stMetricValue"],
            div[data-testid="stMetric"] [data-testid="stMetricDelta"],
            div[data-testid="stMetric"] * {
                color: var(--text-color) !important;
            }
            .filter-note, .period-banner {
                padding: 0.9rem 1rem;
                border-radius: 14px;
                margin-bottom: 0.75rem;
                color: var(--text-color);
                border: 1px solid rgba(59, 130, 246, 0.22);
                background: color-mix(in srgb, var(--secondary-background-color) 90%, #2563eb 10%);
            }
            .period-banner strong {
                font-size: 1rem;
            }
            div[data-baseweb="select"] > div,
            div[data-baseweb="input"] > div,
            .stTextInput > div > div > input {
                color: var(--text-color) !important;
            }
            .stTabs [data-baseweb="tab-list"] button {
                color: var(--text-color);
            }
        </style>
        <div class="main-card">
            <h1 style="margin:0; font-size:2rem;">Analyse RDF Flams</h1>
            <div class="subtle">Import des personnes + plusieurs rapports HTML, filtres par date, recherche par note, tableau de bord, export Excel et rapport PDF.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_overview(df_people, df_tickets, df_discounts):
    if df_people.empty and df_discounts.empty:
        st.warning("Aucune donnée après filtrage.")
        return

    period_parts = []
    if not df_people.empty:
        period_parts.extend(df_people["DATE_JOUR"].dropna().tolist())
    if not df_discounts.empty:
        period_parts.extend(df_discounts["date"].dropna().tolist())

    period_label = "-"
    if period_parts:
        period_label = f"{min(period_parts).strftime('%Y-%m-%d')} - {max(period_parts).strftime('%Y-%m-%d')}"

    st.markdown(
        f'<div class="period-banner"><strong>Période analysée :</strong> {period_label}</div>',
        unsafe_allow_html=True,
    )

    total_remises = df_discounts["discount_amount"].sum() if not df_discounts.empty else 0.0
    total_to_pay = df_tickets["total_to_pay"].sum() if not df_tickets.empty else 0.0
    roi_pct = compute_roi(total_to_pay, total_remises)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Personnes Excel", f"{len(df_people):,}".replace(",", " "))
    c2.metric("Tickets RDF", f"{df_tickets['ticket_key'].nunique() if not df_tickets.empty else 0:,}".replace(",", " "))
    c3.metric("Remises RDF", f"{len(df_discounts):,}".replace(",", " "))
    c4.metric("Montant remises", format_eur(total_remises))
    c5.metric("Total to pay", format_eur(total_to_pay))

    c6, c7, c8 = st.columns(3)
    c6.metric("ROI", format_pct(roi_pct))
    c7.metric("Notes visibles", f"{df_tickets['note_number'].astype(str).nunique() if not df_tickets.empty else 0:,}".replace(",", " "))
    c8.metric("Fichiers HTML", f"{df_tickets['source_file'].astype(str).nunique() if not df_tickets.empty else 0:,}".replace(",", " "))


def main():
    st.set_page_config(page_title="Analyse RDF Flams", layout="wide")
    render_header()

    with st.sidebar:
        st.header("Imports")
        people_file = st.file_uploader(
            "Fichier personnes",
            type=["xlsx", "xls", "csv"],
            help="Colonnes attendues : NOM, PRENOM, LOT, DATE/ADDED_TIME",
        )
        html_files = st.file_uploader(
            "Rapports HTML tickets",
            type=["html", "htm"],
            accept_multiple_files=True,
        )
        st.markdown("---")
        st.markdown("**Remises suivies**")
        for reason in TARGET_REASONS:
            st.write(f"• {reason}")

    if not people_file or not html_files:
        st.info("Ajoute le fichier personnes et un ou plusieurs fichiers HTML dans la barre latérale pour démarrer l'analyse.")
        st.stop()

    try:
        df_people = read_people_file(people_file)
        df_tickets, df_discounts = parse_multiple_html_tickets(html_files)
    except Exception as exc:
        st.error(f"Erreur pendant l'import : {exc}")
        st.stop()

    if df_discounts.empty:
        st.warning("Aucune remise RDF ciblée n'a été trouvée dans les fichiers HTML.")
        st.stop()

    all_dates = sorted(
        set(df_people["DATE_JOUR"].dropna().tolist()) | set(df_discounts["date"].dropna().tolist())
    )
    available_reasons = sorted(df_discounts["reason"].dropna().unique().tolist())
    available_lots = sorted(df_people["LOT_LABEL"].dropna().unique().tolist())
    available_sources = sorted(df_tickets["source_file"].dropna().unique().tolist())

    st.markdown(
        '<div class="filter-note"><strong>Filtres</strong> : dates, types RDF, lots, fichiers HTML et recherche par note. Le tableau Excel reste séparé du calcul HTML basé sur la date du ticket et le Number of covers.</div>',
        unsafe_allow_html=True,
    )

    f1, f2, f3, f4, f5 = st.columns([1.2, 1, 1, 1.2, 1.4])
    selected_dates = f1.multiselect("Dates", options=all_dates, default=all_dates)
    selected_reasons = f2.multiselect("Types RDF", options=available_reasons, default=available_reasons)
    selected_lots = f3.multiselect("Lots", options=available_lots, default=available_lots)
    selected_sources = f4.multiselect("Fichiers HTML", options=available_sources, default=available_sources)
    note_query = f5.text_input("Recherche par note", placeholder="Ex: 8091")

    df_people_f, df_tickets_f, df_discounts_f = apply_filters(
        df_people,
        df_tickets,
        df_discounts,
        selected_dates,
        selected_reasons,
        selected_lots,
        selected_sources,
        note_query,
    )

    daily_lot, daily_pivot = (
        build_people_dashboard(df_people_f)
        if not df_people_f.empty
        else (pd.DataFrame(columns=["DATE_JOUR", "LOT_LABEL", "NB_PERSONNES"]), pd.DataFrame())
    )
    html_covers_daily = build_html_covers_dashboard(df_tickets_f)
    discount_summary = (
        build_discount_summary(df_discounts_f, df_tickets_f)
        if not df_discounts_f.empty
        else pd.DataFrame(columns=["reason", "nb_remises", "montant_total_remises", "nb_tickets", "montant_total_to_pay_tickets", "roi_pct"])
    )
    reconciliation = (
        build_reconciliation(df_people_f, df_discounts_f)
        if not (df_people_f.empty and df_discounts_f.empty)
        else pd.DataFrame(columns=["date", "reason", "nb_personnes_attendues", "nb_remises_trouvees", "ecart"])
    )

    render_overview(df_people_f, df_tickets_f, df_discounts_f)

    tab1, tab2, tab3, tab4 = st.tabs(["Vue d'ensemble", "Contrôle", "Détails", "Exports"])

    with tab1:
        left, right = st.columns([1.15, 1])
        with left:
            st.subheader("Personnes par jour et par lot - Excel")
            if daily_pivot.empty:
                st.info("Aucune donnée personnes avec les filtres actuels.")
            else:
                display_pivot = daily_pivot.copy()
                if "DATE_JOUR" in display_pivot.columns:
                    display_pivot["DATE_JOUR"] = pd.to_datetime(display_pivot["DATE_JOUR"]).dt.strftime("%Y-%m-%d")
                st.dataframe(display_pivot, use_container_width=True, hide_index=True)

        with right:
            st.subheader("Synthèse remises RDF")
            if discount_summary.empty:
                st.info("Aucune remise RDF avec les filtres actuels.")
            else:
                summary_view = discount_summary.copy()
                summary_view["montant_total_remises"] = summary_view["montant_total_remises"].map(format_eur)
                summary_view["montant_total_to_pay_tickets"] = summary_view["montant_total_to_pay_tickets"].map(format_eur)
                summary_view["roi_pct"] = summary_view["roi_pct"].map(format_pct)
                st.dataframe(summary_view, use_container_width=True, hide_index=True)

        st.subheader("Personnes par jour en restaurant - HTML RDF")
        if html_covers_daily.empty:
            st.info("Aucune donnée HTML RDF avec les filtres actuels.")
        else:
            html_table = html_covers_daily.copy()
            html_table["DATE_TICKET"] = pd.to_datetime(html_table["DATE_TICKET"]).dt.strftime("%Y-%m-%d")
            st.dataframe(html_table, use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            if not daily_lot.empty:
                fig_people = px.bar(
                    daily_lot.assign(DATE_JOUR=daily_lot["DATE_JOUR"].dt.strftime("%Y-%m-%d")),
                    x="DATE_JOUR",
                    y="NB_PERSONNES",
                    color="LOT_LABEL",
                    barmode="stack",
                    title="Répartition des personnes par jour - Excel",
                )
                st.plotly_chart(fig_people, use_container_width=True)
        with c2:
            if not html_covers_daily.empty:
                fig_html = px.bar(
                    html_covers_daily.assign(DATE_TICKET=html_covers_daily["DATE_TICKET"].dt.strftime("%Y-%m-%d")),
                    x="DATE_TICKET",
                    y="PERSONNES_RESTAURANT",
                    title="Personnes par jour en restaurant - tickets HTML RDF",
                )
                st.plotly_chart(fig_html, use_container_width=True)

        if not discount_summary.empty:
            fig_discounts = px.bar(
                discount_summary,
                x="reason",
                y="montant_total_remises",
                title="Montant des remises par type RDF",
            )
            st.plotly_chart(fig_discounts, use_container_width=True)

    with tab2:
        st.subheader("Contrôle attendu vs trouvé")
        st.caption("Comparaison entre les lots du fichier personnes et les remises RDF retrouvées dans les tickets.")
        if reconciliation.empty:
            st.info("Aucune ligne de contrôle avec les filtres actuels.")
        else:
            rec_view = reconciliation.copy()
            rec_view["date"] = pd.to_datetime(rec_view["date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(rec_view, use_container_width=True, hide_index=True)

            fig_gap = px.bar(
                reconciliation.assign(date=reconciliation["date"].dt.strftime("%Y-%m-%d")),
                x="date",
                y="ecart",
                color="reason",
                barmode="group",
                title="Écart entre attendu et trouvé",
            )
            st.plotly_chart(fig_gap, use_container_width=True)

    with tab3:
        st.subheader("Détails des remises trouvées")
        if df_discounts_f.empty:
            st.info("Aucune remise RDF à afficher.")
        else:
            discounts_view = df_discounts_f.sort_values(["date", "note_number", "reason"]).copy()
            discounts_view["date"] = pd.to_datetime(discounts_view["date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(discounts_view, use_container_width=True, hide_index=True)

        st.subheader("Tickets contenant au moins une remise RDF")
        if df_tickets_f.empty:
            st.info("Aucun ticket RDF à afficher.")
        else:
            tickets_view = df_tickets_f.sort_values(["date", "note_number"]).copy()
            tickets_view["date"] = pd.to_datetime(tickets_view["date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(tickets_view, use_container_width=True, hide_index=True)

        st.subheader("Personnes importées - Excel")
        people_view = df_people_f.sort_values(["DATE_JOUR", "LOT_LABEL", "NOM", "PRENOM"]).copy()
        people_view["DATE_JOUR"] = pd.to_datetime(people_view["DATE_JOUR"]).dt.strftime("%Y-%m-%d")
        st.dataframe(
            people_view[["DATE_JOUR", "LOT_LABEL", "PRENOM", "NOM"]],
            use_container_width=True,
            hide_index=True,
        )

    with tab4:
        st.subheader("Exports")
        export_bytes = to_excel_bytes(
            {
                "synthese_remises": discount_summary,
                "personnes_par_jour_excel": daily_pivot,
                "personnes_restaurant_html": html_covers_daily,
                "controle": reconciliation,
                "details_remises": df_discounts_f,
                "tickets_rdf": df_tickets_f,
                "personnes_excel": df_people_f,
            }
        )
        st.download_button(
            "Télécharger l'export Excel",
            data=export_bytes,
            file_name="analyse_rdf_flams_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        pdf_bytes = build_pdf_report(
            df_people=df_people_f,
            df_tickets=df_tickets_f,
            df_discounts=df_discounts_f,
            daily_pivot=daily_pivot,
            html_covers_daily=html_covers_daily,
            discount_summary=discount_summary,
            reconciliation=reconciliation,
            selected_sources=selected_sources,
        )
        st.download_button(
            "Extraire un rapport PDF",
            data=pdf_bytes,
            file_name="rapport_analyse_rdf_flams.pdf",
            mime="application/pdf",
        )

        st.info(
            "Pour partager l'app à quelqu'un sans installation de son côté, il faut l'héberger en ligne. "
            "Une fois publiée sur Streamlit Community Cloud, la personne n'aura plus qu'à ouvrir le lien et importer ses fichiers."
        )


if __name__ == "__main__":
    main()
