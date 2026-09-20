# Freshdesk MCP Server — Installasjonsguide

## Forutsetninger
- Python 3.10 eller nyere
- Claude Desktop installert

---

## Steg 1 — Installer avhengigheter

Åpne terminal (PowerShell eller cmd) og kjør:

```powershell
pip install "mcp[cli]" httpx pydantic
```

---

## Steg 2 — Test at serveren starter

```powershell
cd "C:\Data\AI\Prosjekter\Freshdesk"

$env:FRESHDESK_DOMAIN="compute.freshdesk.com"
$env:FRESHDESK_API_KEY="<din-api-nøkkel>"

python freshdesk_mcp.py
```

Du skal se noe slikt:
```
Starting MCP server 'freshdesk_mcp' with transport 'stdio'
```
Trykk Ctrl+C for å stoppe — det er nok å bekrefte at den starter.

---

## Steg 3 — Koble til Claude Desktop

Åpne Claude Desktop sin konfigurasjonsfil. Den ligger typisk her:

```
C:\Users\Idar.COMPUTE\AppData\Roaming\Claude\claude_desktop_config.json
```

Legg til (eller flett inn i) `mcpServers`-blokken:

```json
{
  "mcpServers": {
    "freshdesk": {
      "command": "python",
      "args": [
        "C:\\Data\\AI\\Prosjekter\\Freshdesk\\freshdesk_mcp.py"
      ],
      "env": {
        "FRESHDESK_DOMAIN": "compute.freshdesk.com",
        "FRESHDESK_API_KEY": "<din-api-nøkkel>"
      }
    }
  }
}
```

> **Tips:** Hvis filen allerede har en `mcpServers`-blokk, legg bare til `"freshdesk": { ... }` inni den.

---

## Steg 4 — Start Claude Desktop på nytt

Lukk og åpne Claude Desktop. MCP-serveren starter automatisk i bakgrunnen.

---

## Tilgjengelige verktøy

Når serveren er koblet til vil Claude ha tilgang til disse verktøyene:

| Verktøy | Beskrivelse |
|---|---|
| `freshdesk_list_tickets` | Hent tickets med filter (status, prioritet, gruppe, agent, dato) |
| `freshdesk_get_ticket` | Hent ett ticket med full detaljer og samtalehistorikk |
| `freshdesk_search_tickets` | Søk med Freshdesk Query Language |
| `freshdesk_list_agents` | List alle agenter med ID-er |
| `freshdesk_list_groups` | List alle grupper/avdelinger med ID-er |
| `freshdesk_get_ticket_stats` | Statistikk: fordeling per status og prioritet |

---

## Eksempler på spørsmål du kan stille Claude

- *"Vis meg alle åpne tickets for ERP-gruppen"*
- *"Hva er statistikken for tickets siste måned?"*
- *"Finn alle tickets med høy prioritet som ikke er løst"*
- *"Hvilke grupper finnes i Freshdesk?"*
- *"Vis ticket #1234 med samtalehistorikk"*

---

## Feilsøking

**"Mangler FRESHDESK_API_KEY"** — `env`-blokken i konfigurasjonsfilen er ikke riktig satt.

**"Ugyldig API-nøkkel"** — Kontroller nøkkelen i Freshdesk under *Profile → API Key*.

**Serveren vises ikke i Claude** — Kontroller JSON-syntaks i `claude_desktop_config.json` (ingen komma etter siste element).
