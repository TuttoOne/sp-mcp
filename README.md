# SharePoint Bridge for Claude

> Control SharePoint with natural language through Claude's Model Context Protocol (MCP).

SharePoint Bridge is an open-source MCP server that connects Claude to your SharePoint environment via the Microsoft Graph API. Ask questions, create lists, update items, query data, and trigger Power Automate flows — all through conversation.

## What It Does

Instead of clicking through SharePoint's UI, you talk to Claude:

```
You: "Show me all active projects with a deadline in the next two weeks"
Claude: [queries SharePoint → synthesises results → presents actionable summary]

You: "Create a new task for Sarah on the Henderson case, due Friday"
Claude: [creates item in SharePoint → confirms with link]

You: "What's the compliance status across all our open investigations?"
Claude: [queries multiple lists → cross-references lookups → produces briefing]
```

Claude reads your SharePoint schema, understands relationships between lists, and works across your entire site — not just one list at a time.

## Features

- **Read & Write** — Query, create, update, and delete list items
- **Schema-Aware** — Understands your column types, lookups, and relationships
- **Multi-List Intelligence** — Cross-references data across related lists in a single conversation
- **Audit & Restructure** — Claude can audit your entire SharePoint environment, diagnose structural problems, and rebuild it for optimal AI use
- **Power Automate Integration** — List, inspect, and trigger flows
- **Microsoft Learn Integration** — Built-in documentation search for accurate Graph API guidance
- **Full Column Support** — Text, number, date, choice, lookup, calculated, boolean, currency, person
- **Dual Transport** — HTTP mode for Claude.ai / remote connections, stdio mode for Claude Code / local use

## Start Here: Audit Your SharePoint

Once the bridge is connected, try this as your first conversation with Claude:

```
"Audit my SharePoint site. List every list, inspect every schema, and tell me:
1. What's well-structured and what's messy
2. Where I'm missing lookup relationships between related lists
3. Which columns should be choice fields instead of free-text
4. What data quality issues you can see
5. How you'd restructure it to work better with AI"
```

Claude will use the bridge tools to scan your entire site and come back with a detailed assessment. This is the fastest way to understand the value of AI-powered SharePoint — and to see exactly what's holding your data back.

If you want help acting on the recommendations, see [Custom Solutions](#custom-solutions) below.

## Quick Start

### Prerequisites

- Node.js 18+
- A Microsoft 365 tenant with SharePoint Online
- An Azure App Registration with Graph API permissions
- A Claude account (Pro, Team, or Enterprise) for Claude.ai, or Claude Code for local use

### 1. Azure App Registration

Create an app registration in [Azure Portal](https://portal.azure.com):

1. Go to **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name: `SharePoint Bridge for Claude`
3. Supported account types: **Single tenant** (or Multi-tenant if serving multiple orgs)
4. Under **API permissions**, add **Application** permissions (not Delegated):
   - `Sites.ReadWrite.All` — for list and item operations
   - `Files.ReadWrite.All` — for document library access
5. Optionally add **Delegated** permission:
   - `User.Read` — for user profile resolution
6. Click **Grant admin consent** for your tenant
7. Under **Certificates & secrets**, create a new client secret and note it down
8. Note your **Application (client) ID** and **Directory (tenant) ID** from the Overview page

### 2. Install & Configure

```bash
git clone https://github.com/TuttoOne/sp-mcp.git
cd sp-mcp
npm install
npm run build
```

Create a `.env` file:

```env
# Azure App Registration
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret

# SharePoint
SP_HOSTNAME=yourcompany.sharepoint.com
SP_DEFAULT_SITE_PATH=/sites/YourSiteName

# Server
PORT=3500
TRANSPORT=http
NODE_ENV=production
```

### 3. Run

```bash
# HTTP mode — for Claude.ai and remote connections
TRANSPORT=http npm start

# stdio mode — for Claude Code and local CLI use
TRANSPORT=stdio npm start
```

In HTTP mode, the MCP server starts at `http://localhost:3500/mcp` (or your configured port).

### 4. Connect to Claude

**Claude.ai:** Go to the conversation's MCP tools section and add your bridge URL (e.g., `https://sp-mcp.yourdomain.com/mcp`).

**Claude Code:** Add to your MCP config:
```bash
claude mcp add --transport http sharepoint "https://sp-mcp.yourdomain.com/mcp"
```

**Claude Desktop:** Add to `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["dist/index.js"],
      "cwd": "/path/to/sp-mcp",
      "env": {
        "TRANSPORT": "stdio",
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "SP_HOSTNAME": "yourcompany.sharepoint.com",
        "SP_DEFAULT_SITE_PATH": "/sites/YourSiteName"
      }
    }
  }
}
```

## Available Tools (15)

### SharePoint Lists (8 tools)

| Tool | Description |
|------|-------------|
| `sp_list_lists` | List all lists and document libraries on a site |
| `sp_list_get` | Get a list's full schema including column definitions, types, and relationships |
| `sp_list_create` | Create a new list (add columns separately via `sp_column_add`) |
| `sp_column_add` | Add columns: text, number, dateTime, choice, lookup, calculated, boolean, currency, personOrGroup |
| `sp_items_query` | Query items with OData filtering, sorting, field selection, and pagination |
| `sp_item_create` | Create new items with field values (supports lookup IDs) |
| `sp_item_update` | Update existing item fields |
| `sp_item_delete` | Delete items |

### Power Automate (4 tools)

| Tool | Description |
|------|-------------|
| `pa_flows_list` | List all flows in an environment |
| `pa_flow_get` | Get full flow definition including triggers and actions |
| `pa_flow_runs` | View run history with status and error details |
| `pa_flow_trigger` | Manually trigger a flow |

### Microsoft Learn Documentation (3 tools)

| Tool | Description |
|------|-------------|
| `ms_docs_search` | Search Microsoft Learn docs (Graph API, Power Automate, SharePoint) |
| `ms_docs_fetch` | Fetch full content of a documentation page |
| `ms_docs_context` | Build comprehensive API reference for a topic (combines search + fetch) |

## Example Use Cases

### Private Investigation Firms
Case management, evidence chain of custody, subject tracking, compliance (GDPR/DPA), expense tracking — all queryable by Claude across related lists.

### Legal Practices
Matter management, document tracking, court deadlines, billing, conflict checks, client relationship management.

### Property Management
Tenancy records, compliance certificates, contractor management, maintenance scheduling, rent tracking.

### Professional Services
Project tracking, resource allocation, time logging, client deliverables, knowledge management.

## Architecture

```
┌──────────────┐     MCP Protocol      ┌─────────────────────┐     Graph API     ┌──────────────┐
│  Claude.ai   │ ◄──────────────────► │  SharePoint Bridge  │ ◄───────────────► │  SharePoint  │
│  Claude Code │  (HTTP or stdio)      │  (Node.js/TS MCP)   │    (REST/HTTP)    │  Online      │
│  Claude      │                       │  Port 3500          │                   │              │
│  Desktop     │                       │                     │                   │              │
└──────────────┘                       └─────────────────────┘                   └──────────────┘
                                              │         │
                                              │         ▼
                                              │  ┌─────────────────────┐
                                              │  │   Power Automate    │
                                              │  │   (Flow Management) │
                                              │  └─────────────────────┘
                                              ▼
                                       ┌─────────────────────┐
                                       │   Microsoft Learn   │
                                       │   (Live doc search) │
                                       └─────────────────────┘
```

The bridge authenticates with Microsoft Graph using the OAuth 2.0 client credentials flow (via MSAL). It translates Claude's MCP tool calls into Graph API requests and returns structured responses that Claude can reason about.

## Deployment

### Build First

The server is written in TypeScript. Always build before running in production:

```bash
npm run build
```

This compiles to `dist/index.js`.

### PM2 (Recommended for Production)

```bash
npm install -g pm2
pm2 start dist/index.js --name sp-mcp
pm2 save
pm2 startup
```

Or use the included ecosystem config:

```bash
pm2 start ecosystem.config.cjs
pm2 save
```

### Nginx Reverse Proxy (for HTTPS)

```nginx
server {
    listen 443 ssl;
    server_name sp-mcp.yourdomain.com;

    ssl_certificate /etc/letsencrypt/live/sp-mcp.yourdomain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/sp-mcp.yourdomain.com/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:3500;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;

        # MCP streaming support — critical for SSE transport
        proxy_set_header Connection '';
        proxy_buffering off;
        proxy_cache off;
        chunked_transfer_encoding on;

        # Extended timeouts for Graph API calls
        proxy_read_timeout 120s;
        proxy_send_timeout 120s;
    }
}
```

Then obtain SSL:
```bash
certbot --nginx -d sp-mcp.yourdomain.com
```

## Important Notes

### Lookup Columns
Lookup columns are not indexed by default. OData filters on lookup fields (`{ColumnName}LookupId`) will return 400 errors on large lists. For small-to-medium lists (< 5,000 items), Claude can pull all items and filter client-side. For larger lists, index the lookup column in SharePoint's list settings.

### List Creation & Columns
When creating lists via `sp_list_create`, column types are limited to: text, number, dateTime, choice, boolean, currency. For lookup, calculated, and personOrGroup columns, create the list first, then add columns separately using `sp_column_add`.

### Permissions
The app registration requires **Application** permissions (not Delegated) for the client credentials flow. `Sites.ReadWrite.All` provides read and write access to all sites in the tenant. For tighter security, use `Sites.Selected` and grant per-site access via PowerShell.

### Token Caching
The bridge caches access tokens for their full lifetime (typically 1 hour). If you change permissions in Azure, restart the bridge to force a token refresh:
```bash
pm2 restart sp-mcp
```

### Known Limitations
- Power Automate flow **creation/editing** is not supported — the Flow API supports listing, triggering, and inspecting, but not authoring flow definitions programmatically.
- Lookup columns return as `{ColumnName}LookupId` (integer), not display values. Cross-reference with a second query to resolve.
- `Address` is a reserved column name in SharePoint — use alternatives like `ClientAddress`.
- Hyperlink columns may return 500 errors when creating items (known Microsoft Graph bug).

## Custom Solutions

This open-source bridge is the foundation. Tutto.one offers a full service ladder built on top of it:

- **SharePoint AI Readiness Audit** — We connect the bridge and Claude audits your entire SharePoint environment. You get a detailed report on what's working, what's broken, and exactly what to fix. Free self-service, or guided with a walkthrough call.
- **SharePoint Cleanup & Restructure** — Claude diagnosed the problems; now we fix them. We restructure your lists, retype your columns, create proper lookup relationships, and migrate your data. Your SharePoint goes from messy to AI-ready.
- **Industry-specific data architectures** — Pre-built SharePoint schemas for PI firms, legal practices, property management, professional services, and recruitment. Proven structures with demo data and Claude prompt templates.
- **Custom software for your organisation** — Not just industry templates, but bespoke systems designed for exactly how your organisation operates. Custom MCP tools, custom workflows, custom integrations — built around your specific processes, terminology, and requirements.
- **Power Automate integration** — Automated workflows triggered by Claude: notifications, approvals, document generation, cross-system orchestration.
- **Multi-system integration** — Connect Claude to SharePoint AND your other tools (Gmail, Calendar, Slack, Linear) through MCP. One conversation, multiple systems.
- **Managed service** — We host the bridge, monitor it, keep it updated, and provide ongoing support. You focus on your business.

Contact **[Tutto.one](https://tutto.one)** — the team that built this bridge.

Email: daniel@tutto.one

## Contributing

Contributions welcome. Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

We especially welcome industry template contributions — if you've built a SharePoint architecture for a specific sector and want to share the schema, open an issue and let's collaborate.

## License

MIT — see [LICENSE](LICENSE) for details.

---

Built by [Tutto.one](https://tutto.one) | Powered by [Anthropic's MCP](https://modelcontextprotocol.io)
