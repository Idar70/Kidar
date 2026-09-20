"""
hent_data.py — Freshdesk fullstendig datahenting for Compute AS
================================================================
Henter historiske data fra Freshdesk og lagrer i SQLite-database
(freshdesk.db) som grunnlag for rapporter og analyser.

Bruk:
  python hent_data.py                # Hent tickets + referansedata
  python hent_data.py --samtaler     # Inkluder samtaler per ticket
  python hent_data.py --tidslogg     # Inkluder tidsloggposter per ticket
  python hent_data.py --csat         # Inkluder kundetilfredshetsvurderinger
  python hent_data.py --alt          # Hent absolutt alt

Inkrementell oppdatering: scriptet husker sist hentede tidspunkt og
henter kun nye/endrede data ved neste kjoring.

Krever: pip install httpx
"""

import httpx
import os
import sys
import time
import json
import sqlite3
from datetime import datetime, timezone

# ── Konfigurasjon ─────────────────────────────────────────────────────────────
DOMAIN   = os.getenv("FRESHDESK_DOMAIN",  "compute.freshdesk.com")
API_KEY  = os.getenv("FRESHDESK_API_KEY", "")
BASE_URL = f"https://{DOMAIN}/api/v2"
FROM_DATE = "2010-01-01"  # Tidlig nok til å fange all historikk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE    = os.path.join(SCRIPT_DIR, "freshdesk.db")

# CLI-flagg
ARGS          = set(sys.argv[1:])
HENT_SAMTALER = "--samtaler" in ARGS or "--alt" in ARGS
HENT_TIDSLOGG = "--tidslogg" in ARGS or "--alt" in ARGS
HENT_CSAT     = "--csat"     in ARGS or "--alt" in ARGS


# ── API-hjelper ───────────────────────────────────────────────────────────────

def api_get(path: str, params: dict | None = None) -> list | dict | None:
    """GET mot Freshdesk API med automatisk retry ved rate-limiting/feil."""
    auth = (API_KEY, "X")
    url  = f"{BASE_URL}/{path}"
    last_exc = None
    for attempt in range(5):
        try:
            r = httpx.get(url, auth=auth, params=params or {}, timeout=30)
            if r.status_code in (429, 502, 503, 504):
                wait = 10 * (attempt + 1)
                print(f"    HTTP {r.status_code} — venter {wait}s og prover igjen ...")
                time.sleep(wait)
                continue
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.json()
        except httpx.TimeoutException as e:
            last_exc = e
            wait = 10 * (attempt + 1)
            print(f"    Timeout — venter {wait}s og prover igjen ...")
            time.sleep(wait)
        except httpx.HTTPStatusError as e:
            raise
    if last_exc:
        raise last_exc


# ── Database-oppsett ───────────────────────────────────────────────────────────

def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def setup_schema(conn: sqlite3.Connection):
    """Oppretter alle tabeller og indekser hvis de ikke finnes."""
    conn.executescript("""
    -- ── Tickets ──────────────────────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS tickets (
        id                       INTEGER PRIMARY KEY,
        subject                  TEXT,
        status                   INTEGER,   -- 2=Open, 3=Pending, 4=Resolved, 5=Closed
        priority                 INTEGER,   -- 1=Low, 2=Medium, 3=High, 4=Urgent
        source                   INTEGER,   -- 1=Email, 2=Portal, 3=Phone, 7=Chat, 9=Feedback
        type                     TEXT,      -- Question / Problem / Incident / Feature Request
        spam                     INTEGER DEFAULT 0,
        group_id                 INTEGER,
        responder_id             INTEGER,   -- agent som har ticket
        requester_id             INTEGER,   -- kontakt som sendte inn
        company_id               INTEGER,
        email_config_id          INTEGER,
        product_id               INTEGER,
        due_by                   TEXT,      -- SLA-frist (ISO 8601)
        fr_due_by                TEXT,      -- Forste svar-frist (ISO 8601)
        fr_escalated             INTEGER DEFAULT 0,
        is_escalated             INTEGER DEFAULT 0,
        nr_due_by                TEXT,
        nr_escalated             INTEGER DEFAULT 0,
        association_type         INTEGER,
        associated_tickets_count INTEGER,
        tags                     TEXT,      -- JSON-array
        custom_fields            TEXT,      -- JSON-objekt
        created_at               TEXT,      -- ISO 8601
        updated_at               TEXT,
        resolved_at              TEXT,      -- Fra stats-feltet (krever include=stats)
        closed_at                TEXT,      -- Fra stats-feltet
        first_responded_at       TEXT,      -- Fra stats-feltet
        synced_at                TEXT       -- Tidspunkt for siste sync
    );

    CREATE INDEX IF NOT EXISTS idx_tickets_group     ON tickets(group_id);
    CREATE INDEX IF NOT EXISTS idx_tickets_status    ON tickets(status);
    CREATE INDEX IF NOT EXISTS idx_tickets_priority  ON tickets(priority);
    CREATE INDEX IF NOT EXISTS idx_tickets_created   ON tickets(created_at);
    CREATE INDEX IF NOT EXISTS idx_tickets_updated   ON tickets(updated_at);
    CREATE INDEX IF NOT EXISTS idx_tickets_responder ON tickets(responder_id);
    CREATE INDEX IF NOT EXISTS idx_tickets_requester ON tickets(requester_id);

    -- ── Samtaler (replies og notater per ticket) ───────────────────────────
    CREATE TABLE IF NOT EXISTS conversations (
        id           INTEGER PRIMARY KEY,
        ticket_id    INTEGER NOT NULL,
        body_text    TEXT,
        from_email   TEXT,
        user_id      INTEGER,
        support_email TEXT,
        incoming     INTEGER DEFAULT 0,  -- 1 = fra kunde, 0 = fra agent
        private_note INTEGER DEFAULT 0,  -- 1 = intern notat
        source       INTEGER,
        created_at   TEXT,
        updated_at   TEXT,
        synced_at    TEXT,
        FOREIGN KEY (ticket_id) REFERENCES tickets(id)
    );

    CREATE INDEX IF NOT EXISTS idx_conv_ticket  ON conversations(ticket_id);
    CREATE INDEX IF NOT EXISTS idx_conv_created ON conversations(created_at);
    CREATE INDEX IF NOT EXISTS idx_conv_user    ON conversations(user_id);

    -- ── Tidslogg (tid brukt per ticket) ───────────────────────────────────
    CREATE TABLE IF NOT EXISTS time_entries (
        id                   INTEGER PRIMARY KEY,
        ticket_id            INTEGER NOT NULL,
        agent_id             INTEGER,
        time_spent           TEXT,     -- Format: "HH:MM"
        time_spent_in_seconds INTEGER,
        billable             INTEGER DEFAULT 0,
        note                 TEXT,
        executed_at          TEXT,
        created_at           TEXT,
        updated_at           TEXT,
        FOREIGN KEY (ticket_id) REFERENCES tickets(id)
    );

    CREATE INDEX IF NOT EXISTS idx_time_ticket ON time_entries(ticket_id);
    CREATE INDEX IF NOT EXISTS idx_time_agent  ON time_entries(agent_id);

    -- ── Agenter ────────────────────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS agents (
        id         INTEGER PRIMARY KEY,
        name       TEXT,
        email      TEXT,
        agent_type INTEGER,
        role_id    INTEGER,
        active     INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT,
        synced_at  TEXT
    );

    -- ── Grupper (avdelinger) ───────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS groups_list (
        id          INTEGER PRIMARY KEY,
        name        TEXT,
        description TEXT,
        agent_count INTEGER,
        created_at  TEXT,
        updated_at  TEXT,
        synced_at   TEXT
    );

    -- ── Kontakter (requestere) ─────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS contacts (
        id         INTEGER PRIMARY KEY,
        name       TEXT,
        email      TEXT,
        phone      TEXT,
        mobile     TEXT,
        company_id INTEGER,
        active     INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT,
        synced_at  TEXT
    );

    CREATE INDEX IF NOT EXISTS idx_contacts_company ON contacts(company_id);
    CREATE INDEX IF NOT EXISTS idx_contacts_email   ON contacts(email);

    -- ── Selskaper ──────────────────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS companies (
        id          INTEGER PRIMARY KEY,
        name        TEXT,
        description TEXT,
        created_at  TEXT,
        updated_at  TEXT,
        synced_at   TEXT
    );

    -- ── Kundetilfredshet (CSAT) ────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS csat_ratings (
        ticket_id     INTEGER PRIMARY KEY,
        survey_remark TEXT,
        ratings       TEXT,    -- JSON-objekt med detaljerte ratings
        agent_id      INTEGER,
        group_id      INTEGER,
        created_at    TEXT,
        updated_at    TEXT,
        synced_at     TEXT,
        FOREIGN KEY (ticket_id) REFERENCES tickets(id)
    );

    -- ── Sync-logg ──────────────────────────────────────────────────────────
    CREATE TABLE IF NOT EXISTS sync_log (
        table_name     TEXT PRIMARY KEY,
        last_synced_at TEXT,
        records_total  INTEGER
    );
    """)
    conn.commit()


def get_last_sync(conn: sqlite3.Connection, table_name: str) -> str | None:
    row = conn.execute(
        "SELECT last_synced_at FROM sync_log WHERE table_name = ?", (table_name,)
    ).fetchone()
    return row["last_synced_at"] if row else None


def set_last_sync(conn: sqlite3.Connection, table_name: str, synced_at: str, count: int):
    conn.execute(
        "INSERT OR REPLACE INTO sync_log VALUES (?, ?, ?)",
        (table_name, synced_at, count)
    )
    conn.commit()


# ── Tickets ────────────────────────────────────────────────────────────────────

def _upsert_ticket_batch(conn: sqlite3.Connection, batch: list[dict], synced_at: str) -> tuple[int, int, int]:
    """Lagrer en batch tickets. Returnerer (nye, oppdaterte, hoppet_over)."""
    new_c = updated_c = skipped_c = 0
    for t in batch:
        # Ingen datofilter — ta med all historikk

        stats = t.get("stats") or {}
        existing = conn.execute(
            "SELECT id FROM tickets WHERE id = ?", (t["id"],)
        ).fetchone()

        if existing:
            updated_c += 1
        else:
            new_c += 1

        conn.execute("""
            INSERT OR REPLACE INTO tickets VALUES (
                :id, :subject, :status, :priority, :source, :type, :spam,
                :group_id, :responder_id, :requester_id, :company_id,
                :email_config_id, :product_id,
                :due_by, :fr_due_by, :fr_escalated, :is_escalated,
                :nr_due_by, :nr_escalated,
                :association_type, :associated_tickets_count,
                :tags, :custom_fields,
                :created_at, :updated_at,
                :resolved_at, :closed_at, :first_responded_at,
                :synced_at
            )
        """, {
            "id":                       t["id"],
            "subject":                  t.get("subject"),
            "status":                   t.get("status"),
            "priority":                 t.get("priority"),
            "source":                   t.get("source"),
            "type":                     t.get("type"),
            "spam":                     int(bool(t.get("spam"))),
            "group_id":                 t.get("group_id"),
            "responder_id":             t.get("responder_id"),
            "requester_id":             t.get("requester_id"),
            "company_id":               t.get("company_id"),
            "email_config_id":          t.get("email_config_id"),
            "product_id":               t.get("product_id"),
            "due_by":                   t.get("due_by"),
            "fr_due_by":                t.get("fr_due_by"),
            "fr_escalated":             int(bool(t.get("fr_escalated"))),
            "is_escalated":             int(bool(t.get("is_escalated"))),
            "nr_due_by":                t.get("nr_due_by"),
            "nr_escalated":             int(bool(t.get("nr_escalated"))),
            "association_type":         t.get("association_type"),
            "associated_tickets_count": t.get("associated_tickets_count"),
            "tags":                     json.dumps(t.get("tags") or []),
            "custom_fields":            json.dumps(t.get("custom_fields") or {}),
            "created_at":               t.get("created_at"),
            "updated_at":               t.get("updated_at"),
            # Stats-felter (krever include=stats i API-kallet)
            "resolved_at":              stats.get("resolved_at"),
            "closed_at":                stats.get("closed_at"),
            "first_responded_at":       stats.get("first_responded_at"),
            "synced_at":                synced_at,
        })

    conn.commit()
    return new_c, updated_c, skipped_c


def sync_tickets(conn: sqlite3.Connection, synced_at: str):
    """
    Henter alle tickets fra Freshdesk med inkrementell oppdatering.

    Freshdesk tillater maks 1000 tickets (10 sider x 100) per sporringsresultat.
    Ved mer enn 1000 endringer siden sist henting lopper scriptet automatisk
    og bruker siste updated_at som ny 'siden'-dato for neste batch.
    """
    last_sync = get_last_sync(conn, "tickets")
    since     = last_sync if last_sync else FROM_DATE + "T00:00:00Z"

    if last_sync:
        print(f"\nTickets — inkrementell henting siden {last_sync[:16]} ...")
    else:
        print(f"\nTickets — full henting fra {FROM_DATE} (forste kjoring) ...")

    total_new = total_updated = batch_num = 0

    while True:
        batch_num += 1
        batch_new = batch_updated = 0
        last_updated_in_batch = since
        page = 1

        while True:
            batch = api_get("tickets", {
                "updated_since": since,
                "order_by":      "updated_at",
                "order_type":    "asc",
                "per_page":      100,
                "page":          page,
                "include":       "stats",   # gir resolved_at, closed_at, first_responded_at
            }) or []

            if not batch:
                break

            new_c, upd_c, skip_c = _upsert_ticket_batch(conn, batch, synced_at)
            batch_new     += new_c
            batch_updated += upd_c

            # Husk siste updated_at for eventuell neste lopp
            last_updated_in_batch = batch[-1].get("updated_at", since)

            print(f"  Side {page}: +{new_c} nye, ~{upd_c} oppdatert, {skip_c} hoppet over")

            if len(batch) < 100:
                break   # Siste side

            page += 1

            if page > 10:
                # Treffer Freshdesk-grensen pa 1000 tickets
                # — fortsett med ny sporringsrunde fra siste kjente updated_at
                print(f"  Treffer 1000-ticketsgrensen — starter ny runde fra {last_updated_in_batch[:16]} ...")
                since = last_updated_in_batch
                break

        else:
            # Ferdig — ingen page > 10
            total_new     += batch_new
            total_updated += batch_updated
            break

        total_new     += batch_new
        total_updated += batch_updated

        if page <= 10:
            break   # Siste batch var ferdig

    count = conn.execute("SELECT COUNT(*) FROM tickets").fetchone()[0]
    set_last_sync(conn, "tickets", synced_at, count)
    print(f"  Totalt: +{total_new} nye, ~{total_updated} oppdatert | {count} tickets i database")


# ── Agenter ────────────────────────────────────────────────────────────────────

def sync_agents(conn: sqlite3.Connection, synced_at: str):
    print("\nAgenter ...")
    agents = api_get("agents") or []
    for a in agents:
        contact = a.get("contact") or {}
        conn.execute(
            "INSERT OR REPLACE INTO agents VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                a["id"],
                contact.get("name"),
                contact.get("email"),
                a.get("type"),
                a.get("role_id"),
                int(bool(a.get("available", True))),
                a.get("created_at"),
                a.get("updated_at"),
                synced_at,
            )
        )
    conn.commit()
    set_last_sync(conn, "agents", synced_at, len(agents))
    print(f"  {len(agents)} agenter lagret.")


# ── Grupper ────────────────────────────────────────────────────────────────────

def sync_groups(conn: sqlite3.Connection, synced_at: str):
    print("\nGrupper ...")
    groups = api_get("groups") or []
    for g in groups:
        conn.execute(
            "INSERT OR REPLACE INTO groups_list VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                g["id"],
                g.get("name"),
                g.get("description"),
                len(g.get("agent_ids") or []),
                g.get("created_at"),
                g.get("updated_at"),
                synced_at,
            )
        )
    conn.commit()
    set_last_sync(conn, "groups_list", synced_at, len(groups))
    print(f"  {len(groups)} grupper lagret.")


# ── Kontakter ──────────────────────────────────────────────────────────────────

def sync_contacts(conn: sqlite3.Connection, synced_at: str):
    """
    Henter alle kontakter (requestere). Freshdesk kontakt-API
    stoetter ikke updated_since direkte, sa vi henter all paginerte
    data og UPSERT-er (eksisterende poster overskrives ved endring).
    """
    last_sync = get_last_sync(conn, "contacts")
    if last_sync:
        print(f"\nKontakter — henter endringer siden {last_sync[:16]} ...")
    else:
        print("\nKontakter — full henting (forste kjoring) ...")

    page  = 0
    total = 0
    while True:
        page += 1
        params = {"page": page, "per_page": 100}
        if last_sync:
            params["updated_since"] = last_sync
        batch = api_get("contacts", params) or []
        if not batch:
            break
        for c in batch:
            conn.execute(
                "INSERT OR REPLACE INTO contacts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    c["id"],
                    c.get("name"),
                    c.get("email"),
                    c.get("phone"),
                    c.get("mobile"),
                    c.get("company_id"),
                    int(bool(c.get("active", True))),
                    c.get("created_at"),
                    c.get("updated_at"),
                    synced_at,
                )
            )
        conn.commit()
        total += len(batch)
        print(f"  Side {page}: {len(batch)} kontakter (totalt: {total})")
        if len(batch) < 100:
            break

    count = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
    set_last_sync(conn, "contacts", synced_at, count)
    print(f"  {count} kontakter totalt i database.")


# ── Selskaper ──────────────────────────────────────────────────────────────────

def sync_companies(conn: sqlite3.Connection, synced_at: str):
    print("\nSelskaper ...")
    page  = 0
    total = 0
    while True:
        page += 1
        batch = api_get("companies", {"page": page, "per_page": 100}) or []
        if not batch:
            break
        for c in batch:
            conn.execute(
                "INSERT OR REPLACE INTO companies VALUES (?, ?, ?, ?, ?, ?)",
                (
                    c["id"],
                    c.get("name"),
                    c.get("description"),
                    c.get("created_at"),
                    c.get("updated_at"),
                    synced_at,
                )
            )
        conn.commit()
        total += len(batch)
        print(f"  Side {page}: {len(batch)} selskaper (totalt: {total})")
        if len(batch) < 100:
            break

    count = conn.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
    set_last_sync(conn, "companies", synced_at, count)
    print(f"  {count} selskaper totalt i database.")


# ── Samtaler ───────────────────────────────────────────────────────────────────

def sync_conversations(conn: sqlite3.Connection, synced_at: str):
    """
    Henter samtaler (replies + notater) per ticket via API-kall per ticket.
    Henter kun for tickets som er nye eller endret siden sist samtale-sync.
    """
    last_sync = get_last_sync(conn, "conversations")

    if last_sync:
        rows = conn.execute(
            "SELECT id FROM tickets WHERE updated_at > ? ORDER BY updated_at ASC",
            (last_sync,)
        ).fetchall()
        print(f"\nSamtaler — henter for {len(rows)} oppdaterte tickets ...")
    else:
        rows = conn.execute(
            "SELECT id FROM tickets ORDER BY created_at ASC"
        ).fetchall()
        print(f"\nSamtaler — full henting for {len(rows)} tickets (forste kjoring) ...")

    ticket_ids  = [r["id"] for r in rows]
    total_convs = 0

    for i, ticket_id in enumerate(ticket_ids):
        convs = api_get(f"tickets/{ticket_id}/conversations") or []
        for c in convs:
            conn.execute("""
                INSERT OR REPLACE INTO conversations VALUES (
                    :id, :ticket_id, :body_text, :from_email, :user_id,
                    :support_email, :incoming, :private_note, :source,
                    :created_at, :updated_at, :synced_at
                )
            """, {
                "id":           c["id"],
                "ticket_id":    ticket_id,
                "body_text":    c.get("body_text"),
                "from_email":   c.get("from_email"),
                "user_id":      c.get("user_id"),
                "support_email": c.get("support_email"),
                "incoming":     int(bool(c.get("incoming"))),
                "private_note": int(bool(c.get("private"))),
                "source":       c.get("source"),
                "created_at":   c.get("created_at"),
                "updated_at":   c.get("updated_at"),
                "synced_at":    synced_at,
            })
        conn.commit()
        total_convs += len(convs)

        # Vis fremdrift hvert 50. ticket, og vent litt for rate limiting
        if (i + 1) % 50 == 0 or (i + 1) == len(ticket_ids):
            print(f"  {i+1}/{len(ticket_ids)} tickets behandlet | {total_convs} samtaler hentet")
        time.sleep(0.05)

    count = conn.execute("SELECT COUNT(*) FROM conversations").fetchone()[0]
    set_last_sync(conn, "conversations", synced_at, count)
    print(f"  {count} samtaler totalt i database.")


# ── Tidslogg ───────────────────────────────────────────────────────────────────

def sync_time_entries(conn: sqlite3.Connection, synced_at: str):
    """Henter tidsloggposter per ticket."""
    last_sync = get_last_sync(conn, "time_entries")

    if last_sync:
        rows = conn.execute(
            "SELECT id FROM tickets WHERE updated_at > ? ORDER BY updated_at ASC",
            (last_sync,)
        ).fetchall()
        print(f"\nTidslogg — henter for {len(rows)} oppdaterte tickets ...")
    else:
        rows = conn.execute(
            "SELECT id FROM tickets ORDER BY created_at ASC"
        ).fetchall()
        print(f"\nTidslogg — full henting for {len(rows)} tickets (forste kjoring) ...")

    ticket_ids    = [r["id"] for r in rows]
    total_entries = 0

    for i, ticket_id in enumerate(ticket_ids):
        entries = api_get(f"tickets/{ticket_id}/time_entries") or []
        for e in entries:
            ts      = e.get("time_spent") or "00:00"
            parts   = ts.split(":")
            seconds = int(parts[0]) * 3600 + int(parts[1]) * 60 if len(parts) == 2 else 0
            conn.execute("""
                INSERT OR REPLACE INTO time_entries VALUES (
                    :id, :ticket_id, :agent_id, :time_spent, :time_spent_in_seconds,
                    :billable, :note, :executed_at, :created_at, :updated_at
                )
            """, {
                "id":                   e["id"],
                "ticket_id":            ticket_id,
                "agent_id":             e.get("agent_id"),
                "time_spent":           ts,
                "time_spent_in_seconds": seconds,
                "billable":             int(bool(e.get("billable"))),
                "note":                 e.get("note"),
                "executed_at":          e.get("executed_at"),
                "created_at":           e.get("created_at"),
                "updated_at":           e.get("updated_at"),
            })
        conn.commit()
        total_entries += len(entries)

        if (i + 1) % 50 == 0 or (i + 1) == len(ticket_ids):
            print(f"  {i+1}/{len(ticket_ids)} tickets behandlet | {total_entries} tidsloggposter")
        time.sleep(0.05)

    count = conn.execute("SELECT COUNT(*) FROM time_entries").fetchone()[0]
    set_last_sync(conn, "time_entries", synced_at, count)
    print(f"  {count} tidsloggposter totalt i database.")


# ── CSAT ───────────────────────────────────────────────────────────────────────

def sync_csat(conn: sqlite3.Connection, synced_at: str):
    """Henter kundetilfredshetsvurderinger for lukkede/losende tickets."""
    last_sync = get_last_sync(conn, "csat_ratings")

    if last_sync:
        rows = conn.execute(
            "SELECT id FROM tickets WHERE status IN (4,5) AND updated_at > ?",
            (last_sync,)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT id FROM tickets WHERE status IN (4,5)"
        ).fetchall()

    ticket_ids = [r["id"] for r in rows]
    print(f"\nCSAT — henter for {len(ticket_ids)} losende/lukkede tickets ...")

    total = 0
    for i, ticket_id in enumerate(ticket_ids):
        result = api_get(f"tickets/{ticket_id}/satisfaction_ratings")
        ratings_list = result if isinstance(result, list) else ([result] if result else [])
        for r in ratings_list:
            if not r:
                continue
            conn.execute("""
                INSERT OR REPLACE INTO csat_ratings VALUES (
                    :ticket_id, :survey_remark, :ratings, :agent_id, :group_id,
                    :created_at, :updated_at, :synced_at
                )
            """, {
                "ticket_id":     ticket_id,
                "survey_remark": r.get("survey_remark"),
                "ratings":       json.dumps(r.get("ratings") or {}),
                "agent_id":      r.get("agent_id"),
                "group_id":      r.get("group_id"),
                "created_at":    r.get("created_at"),
                "updated_at":    r.get("updated_at"),
                "synced_at":     synced_at,
            })
            total += 1
        conn.commit()

        if (i + 1) % 100 == 0 or (i + 1) == len(ticket_ids):
            print(f"  {i+1}/{len(ticket_ids)} sjekket | {total} CSAT-svar funnet")
        time.sleep(0.05)

    count = conn.execute("SELECT COUNT(*) FROM csat_ratings").fetchone()[0]
    set_last_sync(conn, "csat_ratings", synced_at, count)
    print(f"  {count} CSAT-svar totalt i database.")


# ── Oppsummering ───────────────────────────────────────────────────────────────

def print_summary(conn: sqlite3.Connection):
    tabeller = [
        ("tickets",       "Tickets"),
        ("conversations", "Samtaler"),
        ("time_entries",  "Tidsloggposter"),
        ("agents",        "Agenter"),
        ("groups_list",   "Grupper"),
        ("contacts",      "Kontakter"),
        ("companies",     "Selskaper"),
        ("csat_ratings",  "CSAT-svar"),
    ]
    print()
    print("=" * 55)
    print("  Oppsummering  —  freshdesk.db")
    print("=" * 55)
    for table, label in tabeller:
        try:
            n = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            print(f"  {label:<22} {n:>6} poster")
        except Exception:
            pass

    print()
    rows = conn.execute(
        "SELECT table_name, last_synced_at FROM sync_log ORDER BY table_name"
    ).fetchall()
    if rows:
        print("  Sist synkronisert:")
        for row in rows:
            ts = (row["last_synced_at"] or "")[:16].replace("T", " ")
            print(f"    {row['table_name']:<22} {ts}")

    print("=" * 55)
    print()
    print("  Anbefalte SQL-sporringseksempler:")
    print()
    print("  -- Alle ERP-tickets med agent og kontakt")
    print("  SELECT t.id, t.subject, t.status, t.created_at,")
    print("         a.name AS agent, c.name AS requester")
    print("  FROM tickets t")
    print("  LEFT JOIN agents a   ON a.id = t.responder_id")
    print("  LEFT JOIN contacts c ON c.id = t.requester_id")
    print("  WHERE t.group_id = 13000002529;")
    print()
    print("  -- Ticket-volum per maned per gruppe")
    print("  SELECT substr(t.created_at,1,7) AS maaned,")
    print("         g.name AS gruppe,")
    print("         COUNT(*) AS antall")
    print("  FROM tickets t")
    print("  LEFT JOIN groups_list g ON g.id = t.group_id")
    print("  GROUP BY maaned, gruppe")
    print("  ORDER BY maaned, gruppe;")
    print()
    print("  -- Gjennomsnittlig losingstid per gruppe (timer)")
    print("  SELECT g.name AS gruppe,")
    print("         ROUND(AVG((julianday(t.resolved_at) -")
    print("               julianday(t.created_at)) * 24), 1) AS snitt_timer")
    print("  FROM tickets t")
    print("  LEFT JOIN groups_list g ON g.id = t.group_id")
    print("  WHERE t.resolved_at IS NOT NULL")
    print("  GROUP BY gruppe;")
    print()
    print("=" * 55)


# ── Hovedprogram ───────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  Freshdesk Datahenting  —  Compute AS")
    print("=" * 55)
    print(f"  Database : {DB_FILE}")
    if HENT_SAMTALER:
        print("  Inkluderer: samtaler")
    if HENT_TIDSLOGG:
        print("  Inkluderer: tidslogg")
    if HENT_CSAT:
        print("  Inkluderer: CSAT-vurderinger")
    print("=" * 55)

    conn = get_db()
    setup_schema(conn)

    synced_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    # ── Kjernadata (hentes alltid) ──
    sync_tickets(conn, synced_at)
    sync_agents(conn, synced_at)
    sync_groups(conn, synced_at)
    sync_contacts(conn, synced_at)
    sync_companies(conn, synced_at)

    # ── Tilleggsdata (valgfritt via CLI-flagg) ──
    if HENT_SAMTALER:
        sync_conversations(conn, synced_at)
    if HENT_TIDSLOGG:
        sync_time_entries(conn, synced_at)
    if HENT_CSAT:
        sync_csat(conn, synced_at)

    print_summary(conn)
    conn.close()

    print(f"\nFerdig! Databasen er lagret: {DB_FILE}")
    print()
    print("Anbefalt visningsvektoy: DB Browser for SQLite")
    print("  https://sqlitebrowser.org  (gratis, Windows)")
    print()
    print("Neste kjoring:")
    print("  python hent_data.py               (hent kun endringer)")
    print("  python hent_data.py --samtaler     (inkluder samtaler)")
    print("  python hent_data.py --alt          (alt inkludert)")


if __name__ == "__main__":
    main()
