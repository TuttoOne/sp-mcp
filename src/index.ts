import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import express from "express";
import { initGraphClient } from "./services/graph-client.js";
import { registerListTools } from "./tools/sharepoint-lists.js";
import { registerFlowTools } from "./tools/power-automate.js";
import { registerDocTools } from "./tools/documentation.js";

// ── Load config from env ──

function loadConfig(): void {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const spHostname = process.env.SP_HOSTNAME;

  if (!tenantId || !clientId || !clientSecret || !spHostname) {
    console.error("Missing required environment variables:");
    if (!tenantId) console.error("  - AZURE_TENANT_ID");
    if (!clientId) console.error("  - AZURE_CLIENT_ID");
    if (!clientSecret) console.error("  - AZURE_CLIENT_SECRET");
    if (!spHostname) console.error("  - SP_HOSTNAME");
    console.error("\nCopy .env.example to .env and fill in your Azure AD app registration details.");
    process.exit(1);
  }

  initGraphClient({
    tenantId,
    clientId,
    clientSecret,
    spHostname,
    defaultSitePath: process.env.SP_DEFAULT_SITE_PATH,
  });

  console.error(`Graph client initialized for ${spHostname}`);
}

// ── Create and configure MCP server ──

function createServer(): McpServer {
  const server = new McpServer({
    name: "sharepoint-mcp-server",
    version: "1.0.0",
  });

  // Register all tool groups
  registerListTools(server);
  registerFlowTools(server);
  registerDocTools(server);

  console.error("Registered tools: SharePoint Lists, Power Automate, Documentation");

  return server;
}

// ── Transport: Streamable HTTP ──

async function runHTTP(): Promise<void> {
  const server = createServer();
  const app = express();
  app.use(express.json());

  // Health check
  app.get("/health", (_req, res) => {
    res.json({ status: "ok", server: "sharepoint-mcp-server", version: "1.0.0" });
  });

  // MCP endpoint
  app.post("/mcp", async (req, res) => {
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    res.on("close", () => transport.close());
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  });

  const port = parseInt(process.env.PORT || "3500");
  app.listen(port, () => {
    console.error(`SharePoint MCP server running on http://localhost:${port}/mcp`);
    console.error(`Health check: http://localhost:${port}/health`);
  });
}

// ── Transport: stdio ──

async function runStdio(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("SharePoint MCP server running on stdio");
}

// ── Main ──

loadConfig();

const transport = process.env.TRANSPORT || "http";
if (transport === "http") {
  runHTTP().catch((error) => {
    console.error("Server error:", error);
    process.exit(1);
  });
} else {
  runStdio().catch((error) => {
    console.error("Server error:", error);
    process.exit(1);
  });
}
