#!/usr/bin/env python3
"""
Freshdesk MCP Server

Gir Claude tilgang til Freshdesk sitt API v2 for ticket-rapportering
og avdelingsoppfølging (f.eks. ERP-avdelingen hos Compute AS).

Konfigurasjon via miljøvariabler:
    FRESHDESK_DOMAIN  — f.eks. compute.freshdesk.com
    FRESHDESK_API_KEY — API-nøkkel fra Freshdesk Profile Settings
"""

import json
import os
from typing import Optional, List, Dict, Any
from enum import Enum
from datetime import datetime, date

import httpx
from pydantic import BaseModel, Field, field_validator, ConfigDict
from mcp.server.fastmcp import FastMCP

# ── Initialisering ────────────────────────────────────────────────────────────

mcp = FastMCP("freshdesk_mcp")

FRESHDESK_DOMAIN = os.environ.get("FRESHDESK_DOMAIN", "compute.freshdesk.com")
FRESHDESK_API_KEY = os.environ.get("FRESHDESK_API_KEY", "")
API_BASE_URL = f"https://{FRESHDESK_DOMAIN}/api/v2"

# ── Oppslagstabeller ──────────────────────────────────────────────────────────

TICKET_STATUS: Dict[int, str] = {
    2: "Open",
    3: "Pending",
    4: "Resolved",
    5: "Closed",
    6: "Waiting on Customer",
    7: "Waiting on Third Party",
}

TICKET_PRIORITY: Dict[int, str] = {
    1: "Low",
    2: "Medium",
    3: "High",
    4: "Urgent",
}

TICKET_SOURCE: Dict[int, str] = {
    1: "Email",
    2: "Portal",
    3: "Phone",
    7: "Chat",
    9: "Feedback Widget",
    10: "Outbound Email",
}

# ── Delte hjelpefunksjoner ────────────────────────────────────────────────────

def _auth() -> tuple[str, str]:
    """Returnerer (api_key, 'X') for Basic-autentisering mot Freshdesk."""
    if not FRESHDESK_API_KEY:
        raise ValueError(
            "Mangler FRESHDESK_API_KEY. Sett miljøvariabelen og start MCP-serveren på nytt."
        )
    return (FRESHDESK_API_KEY, "X")


async def _get(endpoint: str, params: Optional[Dict[str, Any]] = None) -> Any:
    """
    Utfører en GET-forespørsel mot Freshdesk API v2.

    Args:
        endpoint: API-endepunkt uten base-URL, f.eks. '/tickets'
        params:   Valgfrie query-parametere

    Returns:
        Parset JSON-respons (dict eller liste)

    Raises:
        httpx.HTTPStatusError: Ved HTTP-feil
    """
    url = f"{API_BASE_URL}{endpoint}"
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            url,
            auth=_auth(),
            params={k: v for k, v in (params or {}).items() if v is not None},
            timeout=30.0,
            headers={"Content-Type": "application/json"},
        )
        resp.raise_for_status()
        return resp.json()


def _handle_error(e: Exception) -> str:
    """Formaterer feil til lesbar melding med handlingsforslag."""
    if isinstance(e, httpx.HTTPStatusError):
        code = e.response.status_code
        if code == 401:
            return "Feil: Ugyldig API-nøkkel. Sjekk FRESHDESK_API_KEY."
        if code == 403:
            return "Feil: Ingen tilgang til ressursen. Kontroller agentrettigheter i Freshdesk."
        if code == 404:
            return "Feil: Ressursen finnes ikke. Sjekk at ID-en er korrekt."
        if code == 429:
            return "Feil: API-grensen er nådd. Vent litt og prøv igjen."
        try:
            body = e.response.json()
            msg = body.get("description") or body.get("message") or str(body)
        except Exception:
            msg = e.response.text[:300]
        return f"Feil: API returnerte HTTP {code} — {msg}"
    if isinstance(e, httpx.TimeoutException):
        return "Feil: Forespørselen tok for lang tid. Prøv igjen."
    if isinstance(e, ValueError):
        return f"Konfigurasjonsfeil: {e}"
    return f"Uventet feil ({type(e).__name__}): {e}"


def _fmt_ticket(t: Dict[str, Any], include_description: bool = False) -> str:
    """Formaterer ett ticket som Markdown-kortutsnitt."""
    lines = [
        f"### Ticket #{t.get('id')} — {t.get('subject', '(uten tittel)')}",
        f"- **Status**: {TICKET_STATUS.get(t.get('status', 0), t.get('status'))}",
        f"- **Prioritet**: {TICKET_PRIORITY.get(t.get('priority', 0), t.get('priority'))}",
        f"- **Kilde**: {TICKET_SOURCE.get(t.get('source', 0), t.get('source'))}",
    ]
    if t.get("responder_id"):
        lines.append(f"- **Ansvarlig agent (ID)**: {t['responder_id']}")
    if t.get("group_id"):
        lines.append(f"- **Gruppe (ID)**: {t['group_id']}")
    if t.get("requester_id"):
        lines.append(f"- **Rekvirent (ID)**: {t['requester_id']}")
    if t.get("created_at"):
        lines.append(f"- **Opprettet**: {t['created_at'][:10]}")
    if t.get("updated_at"):
        lines.append(f"- **Oppdatert**: {t['updated_at'][:10]}")
    if t.get("due_by"):
        lines.append(f"- **Frist**: {t['due_by'][:16].replace('T', ' ')}")
    if t.get("tags"):
        lines.append(f"- **Tagger**: {', '.join(t['tags'])}")
    if include_description and t.get("description_text"):
        desc = t["description_text"][:500]
        lines += ["", "**Beskrivelse:**", desc]
    return "\n".join(lines)


# ── Enum-modeller ─────────────────────────────────────────────────────────────

class ResponseFormat(str, Enum):
    MARKDOWN = "markdown"
    JSON = "json"

class TicketStatus(int, Enum):
    OPEN = 2
    PENDING = 3
    RESOLVED = 4
    CLOSED = 5

class TicketPriority(int, Enum):
    LOW = 1
    MEDIUM = 2
    HIGH = 3
    URGENT = 4

class TicketOrderBy(str, Enum):
    CREATED_AT = "created_at"
    UPDATED_AT = "updated_at"
    DUE_BY = "due_by"

# ── Input-modeller ────────────────────────────────────────────────────────────

class ListTicketsInput(BaseModel):
    """Input for listing/filtering Freshdesk tickets."""
    model_config = ConfigDict(str_strip_whitespace=True, validate_assignment=True, extra="forbid")

    status: Optional[TicketStatus] = Field(
        default=None,
        description="Filtrer på status: 2=Open, 3=Pending, 4=Resolved, 5=Closed"
    )
    priority: Optional[TicketPriority] = Field(
        default=None,
        description="Filtrer på prioritet: 1=Low, 2=Medium, 3=High, 4=Urgent"
    )
    group_id: Optional[int] = Field(
        default=None,
        description="Filtrer på gruppe-ID (bruk freshdesk_list_groups for å finne ID-er)"
    )
    agent_id: Optional[int] = Field(
        default=None,
        description="Filtrer på agent-ID (bruk freshdesk_list_agents for å finne ID-er)"
    )
    updated_since: Optional[str] = Field(
        default=None,
        description="Hent tickets oppdatert etter denne dato. Format: YYYY-MM-DD"
    )
    tag: Optional[str] = Field(
        default=None,
        description="Filtrer på tag-navn"
    )
    order_by: TicketOrderBy = Field(
        default=TicketOrderBy.CREATED_AT,
        description="Sorter etter: created_at, updated_at, due_by"
    )
    order_type: Optional[str] = Field(
        default="desc",
        description="Sorteringsretning: 'asc' eller 'desc'"
    )
    page: int = Field(default=1, ge=1, description="Sidenummer (30 tickets per side)")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class GetTicketInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    ticket_id: int = Field(..., description="Ticket-ID", ge=1)
    include_conversations: bool = Field(
        default=False,
        description="Om True: henter også samtalehistorikk for ticketen"
    )
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class SearchTicketsInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    query: str = Field(
        ...,
        description=(
            'Søkestreng i Freshdesk Query Language. Eksempler: '
            '"status:2 AND priority:3", '
            '"subject:\'ERP\'", '
            '"agent_id:12345 AND status:2"'
        ),
        min_length=1,
        max_length=512,
    )
    page: int = Field(default=1, ge=1, le=10, description="Sidenummer (maks 10)")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ListAgentsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class ListGroupsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class TicketStatsInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    group_id: Optional[int] = Field(
        default=None,
        description="Beregn statistikk kun for denne gruppen (bruk freshdesk_list_groups)"
    )
    agent_id: Optional[int] = Field(
        default=None,
        description="Beregn statistikk kun for denne agenten"
    )
    updated_since: Optional[str] = Field(
        default=None,
        description="Kun tickets oppdatert etter denne datoen. Format: YYYY-MM-DD"
    )


# ── Verktøy ───────────────────────────────────────────────────────────────────

@mcp.tool(
    name="freshdesk_list_tickets",
    annotations={
        "title": "List Freshdesk Tickets",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_list_tickets(params: ListTicketsInput) -> str:
    """
    Henter en liste med Freshdesk-tickets med valgfri filtrering.

    Returnerer 30 tickets per side sortert etter valgt felt.
    Støtter filtrering på status, prioritet, gruppe, agent, dato og tag.

    Args:
        params (ListTicketsInput):
            - status (Optional[int]): 2=Open, 3=Pending, 4=Resolved, 5=Closed
            - priority (Optional[int]): 1=Low, 2=Medium, 3=High, 4=Urgent
            - group_id (Optional[int]): Filtrer på gruppe-ID
            - agent_id (Optional[int]): Filtrer på agent-ID
            - updated_since (Optional[str]): ISO-dato, f.eks. '2025-01-01'
            - tag (Optional[str]): Tag-navn
            - order_by (str): created_at | updated_at | due_by
            - order_type (str): asc | desc
            - page (int): Sidenummer
            - response_format (str): markdown | json

    Returns:
        str: Formatert Markdown-liste eller JSON-array med tickets
    """
    try:
        query_params: Dict[str, Any] = {
            "page": params.page,
            "order_by": params.order_by.value,
            "order_type": params.order_type,
        }
        if params.status is not None:
            query_params["status"] = params.status.value
        if params.priority is not None:
            query_params["priority"] = params.priority.value
        if params.group_id is not None:
            query_params["group_id"] = params.group_id
        if params.agent_id is not None:
            query_params["responder_id"] = params.agent_id
        if params.updated_since:
            query_params["updated_since"] = params.updated_since
        if params.tag:
            query_params["tag"] = params.tag

        tickets: List[Dict] = await _get("/tickets", query_params)

        if not tickets:
            return "Ingen tickets funnet med de valgte filtrene."

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(tickets, indent=2, ensure_ascii=False)

        lines = [f"# Freshdesk Tickets (side {params.page})", f"Viser {len(tickets)} ticket(s)\n"]
        for t in tickets:
            lines.append(_fmt_ticket(t))
            lines.append("")
        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="freshdesk_get_ticket",
    annotations={
        "title": "Get Freshdesk Ticket",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_get_ticket(params: GetTicketInput) -> str:
    """
    Henter ett spesifikt Freshdesk-ticket med full detaljer.

    Kan også inkludere samtalehistorikk (alle svar og notater).

    Args:
        params (GetTicketInput):
            - ticket_id (int): Ticket-ID
            - include_conversations (bool): Hent samtalehistorikk
            - response_format (str): markdown | json

    Returns:
        str: Ticket-detaljer, eventuelt med samtaler
    """
    try:
        ticket: Dict = await _get(f"/tickets/{params.ticket_id}")

        conversations: List[Dict] = []
        if params.include_conversations:
            conversations = await _get(f"/tickets/{params.ticket_id}/conversations")

        if params.response_format == ResponseFormat.JSON:
            result: Dict[str, Any] = {"ticket": ticket}
            if params.include_conversations:
                result["conversations"] = conversations
            return json.dumps(result, indent=2, ensure_ascii=False)

        lines = [_fmt_ticket(ticket, include_description=True)]

        if params.include_conversations and conversations:
            lines += ["", f"## Samtalehistorikk ({len(conversations)} innlegg)"]
            for c in conversations:
                author = c.get("user_id") or c.get("agent_id") or "ukjent"
                ts = (c.get("created_at") or "")[:16].replace("T", " ")
                body = c.get("body_text") or c.get("body") or ""
                body_short = body[:400].strip()
                private = " *(notat)*" if c.get("private") else ""
                lines += [f"\n**{ts} — ID {author}**{private}", body_short]

        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="freshdesk_search_tickets",
    annotations={
        "title": "Search Freshdesk Tickets",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_search_tickets(params: SearchTicketsInput) -> str:
    """
    Søker etter Freshdesk-tickets med Freshdesk Query Language.

    Returnerer inntil 30 treff per side (maks 10 sider = 300 totalt).

    Args:
        params (SearchTicketsInput):
            - query (str): FQL-søkestreng, f.eks. 'status:2 AND priority:3'
              Eksempler på felter: status, priority, agent_id, group_id,
              created_at, updated_at, due_by, tag, subject
            - page (int): Sidenummer (1–10)
            - response_format (str): markdown | json

    Returns:
        str: Søkeresultater med antall treff og ticketliste
    """
    try:
        data: Dict = await _get(
            "/search/tickets",
            {"query": f'"{params.query}"', "page": params.page},
        )

        tickets: List[Dict] = data.get("results", [])
        total: int = data.get("total", 0)

        if not tickets:
            return f"Ingen tickets funnet for søket: {params.query}"

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(
                {"total": total, "page": params.page, "results": tickets},
                indent=2, ensure_ascii=False,
            )

        lines = [
            f"# Søkeresultater for: `{params.query}`",
            f"Totalt **{total}** treff — side {params.page}\n",
        ]
        for t in tickets:
            lines.append(_fmt_ticket(t))
            lines.append("")
        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="freshdesk_list_agents",
    annotations={
        "title": "List Freshdesk Agents",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_list_agents(params: ListAgentsInput) -> str:
    """
    Lister alle agenter (supportmedarbeidere) i Freshdesk-kontoen.

    Nyttig for å finne agent-ID-er til bruk i andre verktøy.

    Args:
        params (ListAgentsInput):
            - response_format (str): markdown | json

    Returns:
        str: Liste med agenter, inkludert ID, navn og e-post
    """
    try:
        agents: List[Dict] = await _get("/agents")

        if not agents:
            return "Ingen agenter funnet."

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(agents, indent=2, ensure_ascii=False)

        lines = [f"# Freshdesk Agenter ({len(agents)} totalt)\n"]
        for a in agents:
            contact = a.get("contact", {})
            name = contact.get("name") or a.get("name") or "ukjent"
            email = contact.get("email") or a.get("email") or "—"
            agent_id = a.get("id")
            active = "✓" if a.get("available") else "–"
            lines.append(f"- **{name}** (ID: `{agent_id}`) — {email}  tilgjengelig: {active}")
        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="freshdesk_list_groups",
    annotations={
        "title": "List Freshdesk Groups",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_list_groups(params: ListGroupsInput) -> str:
    """
    Lister alle grupper/avdelinger i Freshdesk-kontoen.

    Nyttig for å finne gruppe-ID til bruk i filtrering (f.eks. ERP-avdelingen).

    Args:
        params (ListGroupsInput):
            - response_format (str): markdown | json

    Returns:
        str: Liste med grupper og deres ID-er
    """
    try:
        groups: List[Dict] = await _get("/groups")

        if not groups:
            return "Ingen grupper funnet."

        if params.response_format == ResponseFormat.JSON:
            return json.dumps(groups, indent=2, ensure_ascii=False)

        lines = [f"# Freshdesk Grupper ({len(groups)} totalt)\n"]
        for g in groups:
            gid = g.get("id")
            name = g.get("name") or "ukjent"
            desc = g.get("description") or ""
            desc_str = f" — {desc}" if desc else ""
            lines.append(f"- **{name}** (ID: `{gid}`){desc_str}")
        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="freshdesk_get_ticket_stats",
    annotations={
        "title": "Get Freshdesk Ticket Statistics",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def freshdesk_get_ticket_stats(params: TicketStatsInput) -> str:
    """
    Beregner aggregert statistikk over tickets: antall per status og prioritet.

    Henter opp til 300 tickets (10 sider) og teller opp fordeling.
    Kan filtreres på gruppe og/eller agent, og begrenset til en tidsperiode.

    Args:
        params (TicketStatsInput):
            - group_id (Optional[int]): Kun statistikk for denne gruppen
            - agent_id (Optional[int]): Kun statistikk for denne agenten
            - updated_since (Optional[str]): Fra-dato, format YYYY-MM-DD

    Returns:
        str: Markdown-oppsummering med fordeling per status og prioritet
    """
    try:
        query_params: Dict[str, Any] = {"per_page": 100}
        if params.group_id is not None:
            query_params["group_id"] = params.group_id
        if params.agent_id is not None:
            query_params["responder_id"] = params.agent_id
        if params.updated_since:
            query_params["updated_since"] = params.updated_since

        # Hent opp til 3 sider = 300 tickets
        all_tickets: List[Dict] = []
        for page in range(1, 4):
            query_params["page"] = page
            batch: List[Dict] = await _get("/tickets", query_params)
            if not batch:
                break
            all_tickets.extend(batch)
            if len(batch) < 100:
                break

        if not all_tickets:
            return "Ingen tickets funnet med de valgte filtrene."

        # Tell opp status
        status_counts: Dict[str, int] = {}
        priority_counts: Dict[str, int] = {}
        for t in all_tickets:
            s = TICKET_STATUS.get(t.get("status", 0), str(t.get("status")))
            p = TICKET_PRIORITY.get(t.get("priority", 0), str(t.get("priority")))
            status_counts[s] = status_counts.get(s, 0) + 1
            priority_counts[p] = priority_counts.get(p, 0) + 1

        total = len(all_tickets)
        lines = [f"# Freshdesk Ticket-statistikk", f"Totalt analysert: **{total}** tickets\n"]

        filter_parts = []
        if params.group_id:
            filter_parts.append(f"Gruppe-ID: {params.group_id}")
        if params.agent_id:
            filter_parts.append(f"Agent-ID: {params.agent_id}")
        if params.updated_since:
            filter_parts.append(f"Fra: {params.updated_since}")
        if filter_parts:
            lines.append(f"*Filter: {' | '.join(filter_parts)}*\n")

        lines.append("## Fordeling per status")
        for status_label in ["Open", "Pending", "Resolved", "Closed",
                              "Waiting on Customer", "Waiting on Third Party"]:
            count = status_counts.get(status_label, 0)
            if count > 0:
                pct = round(count / total * 100)
                bar = "█" * (pct // 5)
                lines.append(f"- **{status_label}**: {count} ({pct}%) {bar}")

        lines.append("\n## Fordeling per prioritet")
        for prio_label in ["Urgent", "High", "Medium", "Low"]:
            count = priority_counts.get(prio_label, 0)
            if count > 0:
                pct = round(count / total * 100)
                lines.append(f"- **{prio_label}**: {count} ({pct}%)")

        if total >= 300:
            lines.append(
                "\n> ⚠️ Statistikken er basert på de 300 nyeste tickets. "
                "Det kan finnes flere."
            )

        return "\n".join(lines)

    except Exception as e:
        return _handle_error(e)


# ── Inngang ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run()
