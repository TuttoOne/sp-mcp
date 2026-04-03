# SharePoint MCP Server

An MCP (Model Context Protocol) server that gives Claude direct access to SharePoint and Power Automate via Microsoft Graph API. Built to solve two problems:

1. **CLI-style SharePoint management** — create lists, columns, lookups, items, and query data programmatically
2. **Accurate documentation** — live Microsoft docs search so Claude uses current API info, not stale training data

## Tools

### SharePoint Lists
| Tool | Description |
|------|-------------|
| `sp_list_lists` | List all lists on a site |
| `sp_list_get` | Get list schema (columns, types, relationships) |
| `sp_list_create` | Create a new list with optional columns |
| `sp_column_add` | Add column (text, number, lookup, calculated, choice, person, etc.) |
| `sp_items_query` | Query items with OData filter/sort/select |
| `sp_item_create` | Create a list item |
| `sp_item_update` | Update item fields |
| `sp_item_delete` | Delete an item |

### Power Automate
| Tool | Description |
|------|-------------|
| `pa_flows_list` | List flows in an environment |
| `pa_flow_get` | Get full flow definition |
| `pa_flow_runs` | Get run history with error details |
| `pa_flow_trigger` | Manually trigger a flow |

### Documentation
| Tool | Description |
|------|-------------|
| `ms_docs_search` | Search Microsoft Learn docs (Graph, PA, SharePoint) |
| `ms_docs_fetch` | Fetch full content of a doc page |
| `ms_docs_context` | Build comprehensive reference for a topic |

## Setup

### 1. Azure AD App Registration

1. Go to [portal.azure.com](https://portal.azure.com) > Azure Active Directory > App registrations > New
2. Name: anything you like (e.g. `SharePoint MCP Bridge`)
3. Account type: Single tenant
4. Register, then on the app page:
   - **API Permissions** > Add > Microsoft Graph > Application:
     - `Sites.ReadWrite.All`
     - `Lists.ReadWrite.All`
     - `Files.ReadWrite.All`
   - Click **Grant admin consent**
   - **Certificates & secrets** > New client secret > copy value
5. Note: Tenant ID, Client ID, Client Secret

### 2. Install & Configure

```bash
git clone https://github.com/TuttoOne/sp-mcp.git
cd sp-mcp
npm install
cp .env.example .env
# Edit .env with your Azure AD credentials
npm run build
```

### 3. Run

```bash
# HTTP mode (for remote/Claude.ai connector)
TRANSPORT=http PORT=3500 npm start

# stdio mode (for local Claude Code)
TRANSPORT=stdio npm start
```

### 4. Connect to Claude

**Claude Code (local, stdio):**
Add to your Claude Code MCP config:
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["dist/index.js"],
      "cwd": "/path/to/sp-mcp",
      "env": {
        "TRANSPORT": "stdio"
      }
    }
  }
}
```

**Claude.ai (remote, HTTP):**
Deploy to a server (see `deploy.sh`), then add as a custom MCP connector:
- URL: `https://your-domain.com/mcp`
- Name: `SharePoint`

### 5. Deploy (Optional — PM2 + Nginx)

```bash
# Set your domain and install directory
export SP_MCP_DOMAIN=sp-mcp.yourdomain.com
export SP_MCP_DIR=/opt/sp-mcp

# Run the deploy script
./deploy.sh
```

## Usage Examples

Once connected, Claude can:

```
"List all SharePoint lists on the ProjectDashboard site"
→ Claude calls sp_list_lists

"Create a lookup column on the Tasks list pointing to the Clients list"
→ Claude calls sp_list_lists (to get Clients list ID), then sp_column_add with lookup type

"Show me the last 10 failed Power Automate flow runs"
→ Claude calls pa_flow_runs with status_filter="Failed"

"What's the correct Graph API format for creating a calculated column?"
→ Claude calls ms_docs_search, gets current docs, gives accurate answer
```

## Architecture

```
Claude.ai / Claude Code
       │
       ▼
  MCP Protocol (HTTP/stdio)
       │
       ▼
┌──────────────────────┐
│  sharepoint-mcp-server│
│                      │
│  ┌─ SharePoint tools─┐│
│  │ Lists, Columns,   ││
│  │ Items, Lookups    ││
│  └───────────────────┘│
│  ┌─ PA tools─────────┐│
│  │ Flows, Runs,      ││
│  │ Triggers          ││
│  └───────────────────┘│
│  ┌─ Doc tools────────┐│
│  │ Search, Fetch,    ││
│  │ Context builder   ││
│  └───────────────────┘│
│         │             │
│    MSAL Auth          │
│         │             │
└─────────┼─────────────┘
          ▼
   Microsoft Graph API
   + Flow Management API
   + Microsoft Learn
```

## Known Limitations

- **Power Automate flow creation**: The Flow API supports listing, triggering, and inspecting flows, but creating/editing flow definitions programmatically is complex (involves raw JSON definition manipulation). For flow creation, use the Power Automate web UI.
- **Lookup column values**: Graph API returns lookup fields as `{ColumnName}LookupId` (integer ID), not the display value. You need a second query to resolve display values.
- **Person columns**: Similar to lookups — returned as IDs, need resolution.
- **SharePoint column sorting**: The `$orderby` parameter on list items requires indexed columns and may need the `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly` header.
- **Hyperlink columns**: Graph API may return 500 errors when creating items with hyperlink column data (known MS bug).

## License

MIT
