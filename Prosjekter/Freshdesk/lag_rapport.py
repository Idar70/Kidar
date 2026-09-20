"""
lag_rapport.py — Freshdesk statistikkrapport for Compute AS
============================================================
Leser data fra freshdesk.db og bygger Excel-rapport.

Kjor forst:  python hent_data.py      (henter/oppdaterer data)
Deretter:    python lag_rapport.py     (bygger rapport)

Krever: pip install openpyxl matplotlib
"""

import os
import sys
import sqlite3
import json
import tempfile
from datetime import date
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, LineChart, Reference

# ── Konfigurasjon ─────────────────────────────────────────────────────────────
FROM_DATE  = "2010-01-01"  # Tidlig nok til å fange all historikk
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE    = os.path.join(SCRIPT_DIR, "freshdesk.db")
OUT_FILE   = os.path.join(SCRIPT_DIR, "freshdesk_rapport.xlsx")

ERP_GROUP_ID = 13000002529

STATUS_MAP   = {2: "Open", 3: "Pending", 4: "Resolved", 5: "Closed"}
PRIORITY_MAP = {1: "Low",  2: "Medium",  3: "High",     4: "Urgent"}
SOURCE_MAP   = {1: "E-post", 2: "Portal", 3: "Telefon", 7: "Chat", 9: "Feedback"}

# ── Farger ────────────────────────────────────────────────────────────────────
C_HEADER_BG  = "1F4E79"
C_HEADER_FG  = "FFFFFF"
C_SUBHEAD_BG = "2E75B6"
C_ALT_ROW    = "D6E4F0"
C_ERP_HEADER = "375623"
C_ERP_ALT    = "E2EFDA"

thin   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


# ── Database-hjelper ──────────────────────────────────────────────────────────

def get_db() -> sqlite3.Connection:
    if not os.path.exists(DB_FILE):
        print(f"FEIL: Databasen finnes ikke: {DB_FILE}")
        print("Kjor forst:  python hent_data.py")
        sys.exit(1)
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def load_tickets(conn: sqlite3.Connection) -> list[dict]:
    """
    Henter alle tickets fra databasen som dicts, sortert pa created_at.
    Konverterer tags/custom_fields fra JSON-strenger til Python-objekter.
    """
    rows = conn.execute("""
        SELECT * FROM tickets
        ORDER BY created_at ASC
    """).fetchall()

    tickets = []
    for r in rows:
        t = dict(r)
        # Parse JSON-felter
        try:
            t["tags"]          = json.loads(t["tags"] or "[]")
        except Exception:
            t["tags"]          = []
        try:
            t["custom_fields"] = json.loads(t["custom_fields"] or "{}")
        except Exception:
            t["custom_fields"] = {}
        tickets.append(t)
    return tickets


def load_agents(conn: sqlite3.Connection) -> dict[int, str]:
    """Returnerer {agent_id: navn}."""
    rows = conn.execute("SELECT id, name FROM agents").fetchall()
    return {r["id"]: (r["name"] or "") for r in rows}


def load_groups(conn: sqlite3.Connection) -> dict[int, str]:
    """Returnerer {group_id: navn} fra databasen (kombinert med hardkodede fallbacks)."""
    fallback = {
        13000002529: "Avd. ERP",
        13000009896: "Avd. Insight",
        13000002530: "Avd. IT",
        13000008850: "Intern (adm)",
        13000009133: "Intern (teknisk)",
        13000010520: "IT-sikkerhet",
    }
    try:
        rows = conn.execute("SELECT id, name FROM groups_list").fetchall()
        if rows:
            return {r["id"]: r["name"] for r in rows}
    except Exception:
        pass
    return fallback


# ── Excel-hjelpere ────────────────────────────────────────────────────────────

def header_cell(cell, text, bg=C_HEADER_BG, fg=C_HEADER_FG, bold=True, size=10):
    cell.value = text
    cell.font  = Font(name="Arial", bold=bold, color=fg, size=size)
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def data_cell(cell, value, bold=False, bg=None, align="left", fmt=None):
    cell.value = value
    cell.font  = Font(name="Arial", bold=bold, size=9)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = BORDER
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        cell.number_format = fmt


def set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def freeze(ws, cell="A2"):
    ws.freeze_panes = cell


# ── Ark 1: Sammendrag ─────────────────────────────────────────────────────────

def build_summary_sheet(ws, tickets: list[dict], groups: dict):
    ws.title = "Sammendrag"
    ws.sheet_view.showGridLines = False

    today       = date.today().isoformat()
    erp_tickets = [t for t in tickets if t.get("group_id") == ERP_GROUP_ID]
    open_all    = sum(1 for t in tickets      if t.get("status") in (2, 3))
    open_erp    = sum(1 for t in erp_tickets  if t.get("status") in (2, 3))
    urgent_all  = sum(1 for t in tickets      if t.get("priority") == 4)
    urgent_erp  = sum(1 for t in erp_tickets  if t.get("priority") == 4)

    # Banner
    ws.merge_cells("B2:G2")
    banner       = ws["B2"]
    banner.value = "Freshdesk Statistikkrapport — Compute AS"
    banner.font  = Font(name="Arial", bold=True, size=16, color="FFFFFF")
    banner.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
    banner.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:G3")
    sub       = ws["B3"]
    sub.value = f"Periode: {FROM_DATE}  –  {today}  |  Generert: {today}"
    sub.font  = Font(name="Arial", size=10, color="FFFFFF")
    sub.fill  = PatternFill("solid", fgColor=C_SUBHEAD_BG)
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20

    def kpi_block(row, col, label, value, bg=C_HEADER_BG):
        ws.merge_cells(f"{get_column_letter(col)}{row}:{get_column_letter(col+1)}{row}")
        ws.merge_cells(f"{get_column_letter(col)}{row+1}:{get_column_letter(col+1)}{row+1}")
        c1 = ws.cell(row, col)
        c1.value = label
        c1.font  = Font(name="Arial", size=9, color="FFFFFF", bold=True)
        c1.fill  = PatternFill("solid", fgColor=bg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2 = ws.cell(row + 1, col)
        c2.value = value
        c2.font  = Font(name="Arial", size=20, bold=True, color=bg)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height     = 18
        ws.row_dimensions[row + 1].height = 34

    kpi_block(5, 2, "Totalt tickets",   len(tickets))
    kpi_block(5, 4, "Apne / Pending",   open_all,   "C55A11")
    kpi_block(5, 6, "Urgent (alle)",    urgent_all, "C00000")
    kpi_block(8, 2, "ERP — totalt",     len(erp_tickets), C_ERP_HEADER)
    kpi_block(8, 4, "ERP — Apne",       open_erp,   "375623")
    kpi_block(8, 6, "ERP — Urgent",     urgent_erp, "C00000")

    ws.cell(12, 2).value = "Navigasjon"
    ws.cell(12, 2).font  = Font(name="Arial", bold=True, size=11)

    nav = [
        ("Radata",                "Alle tickets som radata med filter"),
        ("ERP-oversikt",          "Tickets kun for Avd. ERP"),
        ("Statistikk per gruppe", "Antall tickets per avdeling, status og prioritet"),
        ("Manedlig utvikling",    "Manedlig volum per avdeling med linjediagram"),
        ("ERP Manedlig",          "Detaljert manedlig rapport for ERP"),
        ("ERP Losning",           "Losning og lukking per maned med losningsgrad"),
    ]
    for row_idx, (sheet, desc) in enumerate(nav, 13):
        ws.cell(row_idx, 2).value = f"->  {sheet}"
        ws.cell(row_idx, 2).font  = Font(name="Arial", size=9, color="1F4E79", bold=True)
        ws.cell(row_idx, 3).value = desc
        ws.cell(row_idx, 3).font  = Font(name="Arial", size=9)
        ws.row_dimensions[row_idx].height = 16

    for col in ["A","B","C","D","E","F","G","H"]:
        ws.column_dimensions[col].width = 18


# ── Ark 2: Rådata ─────────────────────────────────────────────────────────────

def build_rawdata_sheet(ws, tickets: list[dict], agents: dict, groups: dict):
    ws.title = "Radata"
    ws.sheet_view.showGridLines = False

    headers = [
        "ID", "Opprettet", "Oppdatert", "Lost", "Lukket",
        "Gruppe", "Status", "Prioritet", "Kilde", "Agent", "Emne",
    ]
    for col, h in enumerate(headers, 1):
        header_cell(ws.cell(1, col), h)
    ws.row_dimensions[1].height = 22

    def fmt_dt(s):
        return s[:10] if s else ""

    for row_idx, t in enumerate(tickets, 2):
        bg         = C_ALT_ROW if row_idx % 2 == 0 else None
        group_name = groups.get(t.get("group_id"), f"Ukjent ({t.get('group_id')})")
        agent_name = agents.get(t.get("responder_id", 0), "")

        vals = [
            t.get("id"),
            fmt_dt(t.get("created_at")),
            fmt_dt(t.get("updated_at")),
            fmt_dt(t.get("resolved_at")),
            fmt_dt(t.get("closed_at")),
            group_name,
            STATUS_MAP.get(t.get("status"),   str(t.get("status"))),
            PRIORITY_MAP.get(t.get("priority"), str(t.get("priority"))),
            SOURCE_MAP.get(t.get("source"),   str(t.get("source", ""))),
            agent_name,
            t.get("subject", ""),
        ]
        for col, v in enumerate(vals, 1):
            data_cell(ws.cell(row_idx, col), v, bg=bg,
                      align="center" if col <= 9 else "left")

    set_col_widths(ws, {
        "A": 9, "B": 12, "C": 12, "D": 12, "E": 12,
        "F": 18, "G": 10, "H": 10, "I": 10, "J": 20, "K": 50,
    })
    freeze(ws, "A2")
    ws.auto_filter.ref = f"A1:K{len(tickets)+1}"


# ── Ark 3: ERP-oversikt ───────────────────────────────────────────────────────

def build_erp_sheet(ws, tickets: list[dict], agents: dict):
    ws.title = "ERP-oversikt"
    ws.sheet_view.showGridLines = False

    erp = [t for t in tickets if t.get("group_id") == ERP_GROUP_ID]

    ws.merge_cells("A1:K1")
    title       = ws["A1"]
    title.value = f"Avd. ERP — Ticketoversikt  ({FROM_DATE} – {date.today()})"
    title.font  = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    title.fill  = PatternFill("solid", fgColor=C_ERP_HEADER)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["ID","Opprettet","Oppdatert","Lost","Lukket",
               "Status","Prioritet","Kilde","Agent","Emne"]
    for col, h in enumerate(headers, 1):
        header_cell(ws.cell(2, col), h, bg=C_ERP_HEADER)
    ws.row_dimensions[2].height = 20

    def fmt_dt(s):
        return s[:10] if s else ""

    for row_idx, t in enumerate(erp, 3):
        bg         = C_ERP_ALT if row_idx % 2 == 0 else None
        agent_name = agents.get(t.get("responder_id", 0), "")

        vals = [
            t.get("id"),
            fmt_dt(t.get("created_at")),
            fmt_dt(t.get("updated_at")),
            fmt_dt(t.get("resolved_at")),
            fmt_dt(t.get("closed_at")),
            STATUS_MAP.get(t.get("status"),    str(t.get("status"))),
            PRIORITY_MAP.get(t.get("priority"), str(t.get("priority"))),
            SOURCE_MAP.get(t.get("source"),    str(t.get("source", ""))),
            agent_name,
            t.get("subject", ""),
        ]
        for col, v in enumerate(vals, 1):
            data_cell(ws.cell(row_idx, col), v, bg=bg,
                      align="center" if col <= 8 else "left")

    set_col_widths(ws, {
        "A": 9, "B": 12, "C": 12, "D": 12, "E": 12,
        "F": 12, "G": 10, "H": 10, "I": 20, "J": 55,
    })
    freeze(ws, "A3")
    ws.auto_filter.ref = f"A2:J{len(erp)+2}"


# ── Ark 4: Statistikk per gruppe ──────────────────────────────────────────────

def build_stats_sheet(ws, tickets: list[dict], groups: dict):
    ws.title = "Statistikk per gruppe"
    ws.sheet_view.showGridLines = False

    by_group: dict[int, dict] = defaultdict(lambda: defaultdict(int))
    for t in tickets:
        gid    = t.get("group_id", 0)
        status = STATUS_MAP.get(t.get("status"),    "Ukjent")
        prio   = PRIORITY_MAP.get(t.get("priority"), "Ukjent")
        by_group[gid]["total"]           += 1
        by_group[gid][f"status_{status}"] += 1
        by_group[gid][f"prio_{prio}"]     += 1

    ws.merge_cells("A1:J1")
    title       = ws["A1"]
    title.value = f"Statistikk per gruppe  ({FROM_DATE} – {date.today()})"
    title.font  = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    title.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["Gruppe","Totalt","Open","Pending","Resolved","Closed",
               "Low","Medium","High","Urgent"]
    for col, h in enumerate(headers, 1):
        header_cell(ws.cell(2, col), h)
    ws.row_dimensions[2].height = 20

    # Bruk grupper fra DB, sorter pa ID for konsistent rekkefolge
    group_items = sorted(groups.items())
    row = 3
    for gid, gname in group_items:
        d  = by_group.get(gid, {})
        bg = C_ALT_ROW if row % 2 == 0 else None
        vals = [
            gname,
            d.get("total", 0),
            d.get("status_Open",     0),
            d.get("status_Pending",  0),
            d.get("status_Resolved", 0),
            d.get("status_Closed",   0),
            d.get("prio_Low",    0),
            d.get("prio_Medium", 0),
            d.get("prio_High",   0),
            d.get("prio_Urgent", 0),
        ]
        for col, v in enumerate(vals, 1):
            data_cell(ws.cell(row, col), v, bg=bg,
                      align="left" if col == 1 else "center")
        row += 1

    total_row = row
    data_cell(ws.cell(total_row, 1), "TOTALT", bold=True, bg="D9D9D9")
    for col in range(2, 11):
        cl = get_column_letter(col)
        ws.cell(total_row, col).value     = f"=SUM({cl}3:{cl}{total_row-1})"
        ws.cell(total_row, col).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(total_row, col).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(total_row, col).fill      = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(total_row, col).border    = BORDER

    set_col_widths(ws, {
        "A": 20, "B": 9, "C": 9, "D": 10, "E": 11, "F": 9,
        "G": 9, "H": 10, "I": 9, "J": 9,
    })
    freeze(ws, "A3")

    # Stolpediagram
    chart_row = total_row + 3
    ws.cell(chart_row, 1).value = "Fordeling per gruppe"
    ws.cell(chart_row, 1).font  = Font(name="Arial", bold=True, size=11)

    chart = BarChart()
    chart.type    = "col"
    chart.title   = "Antall tickets per gruppe"
    chart.y_axis.title = "Antall"
    chart.x_axis.title = "Gruppe"
    chart.style   = 10
    chart.width   = 22
    chart.height  = 13

    data_ref = Reference(ws, min_col=2, max_col=2, min_row=2, max_row=total_row - 1)
    cats     = Reference(ws, min_col=1, min_row=3, max_row=total_row - 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"A{chart_row + 1}")


# ── Ark 5: Månedlig utvikling ─────────────────────────────────────────────────

def build_monthly_sheet(ws, tickets: list[dict], groups: dict):
    ws.title = "Manedlig utvikling"
    ws.sheet_view.showGridLines = False

    monthly: dict[str, dict[int, int]] = defaultdict(lambda: defaultdict(int))
    for t in tickets:
        ym = (t.get("created_at") or "")[:7]
        if ym:
            monthly[ym][t.get("group_id", 0)] += 1

    months    = sorted(monthly.keys())
    group_ids = sorted(groups.keys())
    n_cols    = 1 + len(group_ids) + 1

    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    title       = ws["A1"]
    title.value = f"Manedlig ticketutvikling  ({FROM_DATE} – {date.today()})"
    title.font  = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    title.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    header_cell(ws.cell(2, 1), "Maaned")
    for col, gid in enumerate(group_ids, 2):
        header_cell(ws.cell(2, col), groups[gid])
    header_cell(ws.cell(2, len(group_ids) + 2), "Totalt")
    ws.row_dimensions[2].height = 32

    for row_idx, ym in enumerate(months, 3):
        bg = C_ALT_ROW if row_idx % 2 == 0 else None
        data_cell(ws.cell(row_idx, 1), ym, bg=bg, align="center")
        for col, gid in enumerate(group_ids, 2):
            data_cell(ws.cell(row_idx, col), monthly[ym].get(gid, 0), bg=bg, align="center")
        fc = get_column_letter(2)
        lc = get_column_letter(1 + len(group_ids))
        ws.cell(row_idx, len(group_ids) + 2).value     = f"=SUM({fc}{row_idx}:{lc}{row_idx})"
        ws.cell(row_idx, len(group_ids) + 2).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(row_idx, len(group_ids) + 2).alignment = Alignment(horizontal="center")
        ws.cell(row_idx, len(group_ids) + 2).border    = BORDER
        if bg:
            ws.cell(row_idx, len(group_ids) + 2).fill = PatternFill("solid", fgColor=bg)

    total_row = len(months) + 3
    data_cell(ws.cell(total_row, 1), "TOTALT", bold=True, bg="D9D9D9")
    for col in range(2, len(group_ids) + 3):
        cl = get_column_letter(col)
        ws.cell(total_row, col).value     = f"=SUM({cl}3:{cl}{total_row-1})"
        ws.cell(total_row, col).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(total_row, col).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(total_row, col).fill      = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(total_row, col).border    = BORDER

    ws.column_dimensions["A"].width = 11
    for col in range(2, len(group_ids) + 3):
        ws.column_dimensions[get_column_letter(col)].width = 17
    freeze(ws, "A3")

    # Linjediagram
    chart_row = total_row + 3
    chart = LineChart()
    chart.title        = "Manedlig ticketvolum per gruppe"
    chart.y_axis.title = "Antall tickets"
    chart.x_axis.title = "Maaned"
    chart.style        = 10
    chart.width        = 28
    chart.height       = 14
    for col, _ in enumerate(group_ids, 2):
        data_ref = Reference(ws, min_col=col, max_col=col, min_row=2, max_row=total_row - 1)
        chart.add_data(data_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=3, max_row=total_row - 1)
    chart.set_categories(cats)
    ws.add_chart(chart, f"A{chart_row}")


# ── Ark 6: ERP Månedlig ───────────────────────────────────────────────────────

def build_erp_monthly_sheet(ws, tickets: list[dict], agents: dict):
    ws.title = "ERP Manedlig"
    ws.sheet_view.showGridLines = False

    erp = [t for t in tickets if t.get("group_id") == ERP_GROUP_ID]

    monthly_status: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    monthly_prio:   dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    agent_counts:   dict[str, int]            = defaultdict(int)

    for t in erp:
        ym     = (t.get("created_at") or "")[:7]
        status = STATUS_MAP.get(t.get("status"),    "Ukjent")
        prio   = PRIORITY_MAP.get(t.get("priority"), "Ukjent")
        agent  = agents.get(t.get("responder_id", 0), "Ikke tildelt")
        monthly_status[ym][status] += 1
        monthly_prio[ym][prio]     += 1
        agent_counts[agent]        += 1

    months     = sorted(set(monthly_status) | set(monthly_prio))
    statuses   = ["Open","Pending","Resolved","Closed"]
    priorities = ["Low","Medium","High","Urgent"]

    total_cols = 1 + len(statuses) + 1 + len(priorities) + 1
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    title       = ws["A1"]
    title.value = f"ERP — Manedlig detaljrapport  ({FROM_DATE} – {date.today()})"
    title.font  = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    title.fill  = PatternFill("solid", fgColor=C_ERP_HEADER)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:A3")
    header_cell(ws.cell(2, 1), "Maaned", bg=C_ERP_HEADER)

    status_start = 2
    ws.merge_cells(f"{get_column_letter(status_start)}2:{get_column_letter(status_start+len(statuses)-1)}2")
    header_cell(ws.cell(2, status_start), "Status", bg=C_ERP_HEADER)
    for col, s in enumerate(statuses, status_start):
        header_cell(ws.cell(3, col), s, bg=C_ERP_HEADER)
    total_col_s = status_start + len(statuses)
    header_cell(ws.cell(2, total_col_s), "Total", bg=C_ERP_HEADER)
    header_cell(ws.cell(3, total_col_s), "",      bg=C_ERP_HEADER)

    prio_start = total_col_s + 2
    ws.merge_cells(f"{get_column_letter(prio_start)}2:{get_column_letter(prio_start+len(priorities)-1)}2")
    header_cell(ws.cell(2, prio_start), "Prioritet", bg=C_ERP_HEADER)
    for col, p in enumerate(priorities, prio_start):
        header_cell(ws.cell(3, col), p, bg=C_ERP_HEADER)

    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18

    for row_idx, ym in enumerate(months, 4):
        bg = C_ERP_ALT if row_idx % 2 == 0 else None
        data_cell(ws.cell(row_idx, 1), ym, bg=bg, align="center")
        for col, s in enumerate(statuses, status_start):
            data_cell(ws.cell(row_idx, col), monthly_status[ym].get(s, 0), bg=bg, align="center")
        s_first = get_column_letter(status_start)
        s_last  = get_column_letter(status_start + len(statuses) - 1)
        ws.cell(row_idx, total_col_s).value     = f"=SUM({s_first}{row_idx}:{s_last}{row_idx})"
        ws.cell(row_idx, total_col_s).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(row_idx, total_col_s).alignment = Alignment(horizontal="center")
        ws.cell(row_idx, total_col_s).border    = BORDER
        if bg:
            ws.cell(row_idx, total_col_s).fill = PatternFill("solid", fgColor=bg)
        for col, p in enumerate(priorities, prio_start):
            data_cell(ws.cell(row_idx, col), monthly_prio[ym].get(p, 0), bg=bg, align="center")

    data_end = len(months) + 4
    data_cell(ws.cell(data_end, 1), "TOTALT", bold=True, bg="D9D9D9")
    for col in range(status_start, total_col_s + 1):
        cl = get_column_letter(col)
        ws.cell(data_end, col).value     = f"=SUM({cl}4:{cl}{data_end-1})"
        ws.cell(data_end, col).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(data_end, col).alignment = Alignment(horizontal="center")
        ws.cell(data_end, col).fill      = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(data_end, col).border    = BORDER
    for col in range(prio_start, prio_start + len(priorities)):
        cl = get_column_letter(col)
        ws.cell(data_end, col).value     = f"=SUM({cl}4:{cl}{data_end-1})"
        ws.cell(data_end, col).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(data_end, col).alignment = Alignment(horizontal="center")
        ws.cell(data_end, col).fill      = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(data_end, col).border    = BORDER

    # Agent-tabell
    agent_col = prio_start + len(priorities) + 2
    header_cell(ws.cell(2, agent_col),     "Agent",   bg=C_ERP_HEADER)
    header_cell(ws.cell(2, agent_col + 1), "Tickets", bg=C_ERP_HEADER)
    for row_idx, (agent, count) in enumerate(
        sorted(agent_counts.items(), key=lambda x: -x[1]), 3
    ):
        bg = C_ERP_ALT if row_idx % 2 == 0 else None
        data_cell(ws.cell(row_idx, agent_col),     agent, bg=bg)
        data_cell(ws.cell(row_idx, agent_col + 1), count, bg=bg, align="center")

    for col in range(1, agent_col + 2):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions[get_column_letter(agent_col)].width = 22
    freeze(ws, "A4")


# ── Ark 7: ERP Løsning ────────────────────────────────────────────────────────

def _try_matplotlib() -> bool:
    try:
        import matplotlib  # noqa: F401
        return True
    except ImportError:
        pass
    print("    matplotlib mangler — forsoker automatisk installasjon ...")
    import subprocess
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", "matplotlib"]
        )
        import matplotlib  # noqa: F401
        print("    matplotlib installert.")
        return True
    except Exception as e:
        print(f"    Klarte ikke installere matplotlib: {e}")
        return False


def build_erp_resolution_sheet(ws, tickets: list[dict]):
    """
    ERP-losningsark.
    Aggregerer bade pa opprettelsesmaaned (for 'Opprettet'-kolonnen)
    og pa losningsmaaned/lukkemaaned (for 'Lost'/'Lukket'-kolonnene),
    slik at statistikken reflekterer nar arbeidet faktisk ble ferdig.
    """
    ws.title = "ERP Losning"
    ws.sheet_view.showGridLines = False

    erp = [t for t in tickets if t.get("group_id") == ERP_GROUP_ID]

    # -- Opprettet per maaned --
    opprettet: dict[str, int] = defaultdict(int)
    for t in erp:
        ym = (t.get("created_at") or "")[:7]
        if ym:
            opprettet[ym] += 1

    # -- Lost per losningsmaaned (bruker resolved_at hvis tilgjengelig,
    #    ellers created_at som fallback for eldre cache-data) --
    lost:   dict[str, int] = defaultdict(int)
    lukket: dict[str, int] = defaultdict(int)

    for t in erp:
        status = t.get("status")
        if status == 4:   # Resolved
            ym = (t.get("resolved_at") or t.get("created_at") or "")[:7]
            if ym:
                lost[ym] += 1
        elif status == 5:  # Closed
            ym = (t.get("closed_at") or t.get("resolved_at") or t.get("created_at") or "")[:7]
            if ym:
                lukket[ym] += 1

    months = sorted(set(opprettet) | set(lost) | set(lukket))

    # -- Bygg tabell --
    ws.merge_cells("A1:H1")
    title       = ws["A1"]
    title.value = f"ERP — Losning per maaned  ({FROM_DATE} – {date.today()})"
    title.font  = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    title.fill  = PatternFill("solid", fgColor=C_ERP_HEADER)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Underskrift: forklarer aggregeringslogikk
    ws.merge_cells("A2:H2")
    note       = ws["A2"]
    note.value = ("'Opprettet' = antall nye tickets den maaneden. "
                  "'Lost'/'Lukket' = nar de faktisk ble ferdig (resolved_at / closed_at).")
    note.font  = Font(name="Arial", italic=True, size=8, color="595959")
    note.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 14

    cols = ["Maaned", "Opprettet", "Lost", "Lukket",
            "Lost + Lukket", "Losningsgrad %", "Netto inn"]
    for col, h in enumerate(cols, 1):
        header_cell(ws.cell(3, col), h, bg=C_ERP_HEADER)
    ws.row_dimensions[3].height = 36

    for row_idx, ym in enumerate(months, 4):
        bg       = C_ERP_ALT if row_idx % 2 == 0 else None
        opp      = opprettet.get(ym, 0)
        l        = lost.get(ym, 0)
        lk       = lukket.get(ym, 0)
        ll       = l + lk
        r        = row_idx

        data_cell(ws.cell(r, 1), ym,  bg=bg, align="center")
        data_cell(ws.cell(r, 2), opp, bg=bg, align="center")
        data_cell(ws.cell(r, 3), l,   bg=bg, align="center")
        data_cell(ws.cell(r, 4), lk,  bg=bg, align="center")
        data_cell(ws.cell(r, 5), ll,  bold=True, bg=bg, align="center")

        # Losningsgrad: lost+lukket / opprettet
        ws.cell(r, 6).value        = f"=IF(B{r}=0,\"-\",E{r}/B{r})"
        ws.cell(r, 6).number_format = "0.0%"
        ws.cell(r, 6).font         = Font(name="Arial", size=9)
        ws.cell(r, 6).alignment    = Alignment(horizontal="center")
        ws.cell(r, 6).border       = BORDER
        if bg:
            ws.cell(r, 6).fill = PatternFill("solid", fgColor=bg)

        # Netto inn: nye tickets minus ferdigstilte denne maaneden
        ws.cell(r, 7).value        = f"=B{r}-E{r}"
        ws.cell(r, 7).font         = Font(name="Arial", size=9)
        ws.cell(r, 7).alignment    = Alignment(horizontal="center")
        ws.cell(r, 7).border       = BORDER
        if bg:
            ws.cell(r, 7).fill = PatternFill("solid", fgColor=bg)

    total_row = len(months) + 4
    data_cell(ws.cell(total_row, 1), "TOTALT", bold=True, bg="D9D9D9")
    for col in (2, 3, 4):
        cl = get_column_letter(col)
        ws.cell(total_row, col).value     = f"=SUM({cl}4:{cl}{total_row-1})"
        ws.cell(total_row, col).font      = Font(name="Arial", bold=True, size=9)
        ws.cell(total_row, col).alignment = Alignment(horizontal="center")
        ws.cell(total_row, col).fill      = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(total_row, col).border    = BORDER
    ws.cell(total_row, 5).value     = f"=C{total_row}+D{total_row}"
    ws.cell(total_row, 5).font      = Font(name="Arial", bold=True, size=9)
    ws.cell(total_row, 5).alignment = Alignment(horizontal="center")
    ws.cell(total_row, 5).fill      = PatternFill("solid", fgColor="D9D9D9")
    ws.cell(total_row, 5).border    = BORDER
    ws.cell(total_row, 6).value        = f"=IF(B{total_row}=0,\"-\",E{total_row}/B{total_row})"
    ws.cell(total_row, 6).number_format = "0.0%"
    ws.cell(total_row, 6).font         = Font(name="Arial", bold=True, size=9)
    ws.cell(total_row, 6).alignment    = Alignment(horizontal="center")
    ws.cell(total_row, 6).fill         = PatternFill("solid", fgColor="D9D9D9")
    ws.cell(total_row, 6).border       = BORDER
    ws.cell(total_row, 7).value        = f"=B{total_row}-E{total_row}"
    ws.cell(total_row, 7).font         = Font(name="Arial", bold=True, size=9)
    ws.cell(total_row, 7).alignment    = Alignment(horizontal="center")
    ws.cell(total_row, 7).fill         = PatternFill("solid", fgColor="D9D9D9")
    ws.cell(total_row, 7).border       = BORDER

    set_col_widths(ws, {
        "A": 11, "B": 13, "C": 10, "D": 11,
        "E": 15, "F": 16, "G": 14,
    })
    freeze(ws, "A4")

    # ── Diagram (matplotlib PNG, BarChart som fallback) ──
    chart_row = total_row + 3

    if months and _try_matplotlib():
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import matplotlib.ticker as mticker

            y_opp = [opprettet.get(ym, 0)               for ym in months]
            y_ll  = [lost.get(ym, 0) + lukket.get(ym, 0) for ym in months]

            fig, ax = plt.subplots(figsize=(13, 6))
            ax.plot(months, y_opp, color="#1F4E79", linewidth=1.8,
                    marker="s", markersize=4, label="Opprettet")
            ax.plot(months, y_ll,  color="#375623", linewidth=2.4,
                    marker="o", markersize=5, label="Lost + Lukket")
            ax.fill_between(months, y_ll, alpha=0.10, color="#375623")
            ax.set_xlabel("Maaned", fontsize=10)
            ax.set_ylabel("Antall tickets", fontsize=10)
            ax.set_title("ERP — Opprettet vs. Lost/Lukket per maaned",
                         fontsize=12, fontweight="bold")
            ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
            ax.grid(axis="y", linestyle="--", alpha=0.4)
            ax.legend(loc="upper left", fontsize=9, frameon=True)
            plt.xticks(rotation=45, ha="right", fontsize=8)
            fig.tight_layout()

            tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            tmp_path = tmp.name
            tmp.close()
            fig.savefig(tmp_path, dpi=150, bbox_inches="tight")
            plt.close(fig)

            img        = XLImage(tmp_path)
            img.width  = 820
            img.height = 380
            ws.add_image(img, f"A{chart_row}")
            os.unlink(tmp_path)
            print("    Diagram: PNG-bilde (matplotlib).")
            return
        except Exception as e:
            print(f"    matplotlib-feil: {e} — bruker BarChart.")

    # Fallback: BarChart
    if months:
        chart = BarChart()
        chart.type        = "col"
        chart.style       = 11
        chart.title       = "ERP — Opprettet vs. Lost/Lukket per maaned"
        chart.y_axis.title = "Antall tickets"
        chart.x_axis.title = "Maaned"
        chart.width        = 26
        chart.height       = 13
        opp_ref = Reference(ws, min_col=2, max_col=2, min_row=3, max_row=total_row - 1)
        ll_ref  = Reference(ws, min_col=5, max_col=5, min_row=3, max_row=total_row - 1)
        chart.add_data(opp_ref, titles_from_data=True)
        chart.add_data(ll_ref,  titles_from_data=True)
        cats = Reference(ws, min_col=1, min_row=4, max_row=total_row - 1)
        chart.set_categories(cats)
        ws.add_chart(chart, f"A{chart_row}")
        print("    Diagram: BarChart (openpyxl fallback).")


# ── Hovedprogram ──────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  Freshdesk Rapport  —  Compute AS")
    print("=" * 55)
    print(f"  Kilde  : {DB_FILE}")
    print(f"  Utdata : {OUT_FILE}")
    print("=" * 55)

    conn    = get_db()
    tickets = load_tickets(conn)
    agents  = load_agents(conn)
    groups  = load_groups(conn)
    conn.close()

    if not tickets:
        print("Ingen tickets funnet i databasen.")
        print("Kjor forst:  python hent_data.py")
        return

    print(f"Leste {len(tickets)} tickets, {len(agents)} agenter, {len(groups)} grupper.")

    wb = Workbook()
    del wb[wb.sheetnames[0]]

    print("Bygger Excel-rapport ...")
    build_summary_sheet(wb.create_sheet(),   tickets, groups)
    build_rawdata_sheet(wb.create_sheet(),   tickets, agents, groups)
    build_erp_sheet(wb.create_sheet(),       tickets, agents)
    build_stats_sheet(wb.create_sheet(),     tickets, groups)
    build_monthly_sheet(wb.create_sheet(),   tickets, groups)
    build_erp_monthly_sheet(wb.create_sheet(), tickets, agents)
    build_erp_resolution_sheet(wb.create_sheet(), tickets)

    wb.save(OUT_FILE)
    print(f"\nRapport lagret: {OUT_FILE}")
    print("Apne filen i Excel.")


if __name__ == "__main__":
    main()
