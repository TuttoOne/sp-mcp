import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { searchMSDocs, fetchDocPage, buildDocContext } from "../services/docs-service.js";

export function registerDocTools(server: McpServer): void {

  // ─── ms_docs_search ──────────────────────────────────────────────
  server.registerTool(
    "ms_docs_search",
    {
      title: "Search Microsoft Documentation",
      description: `Search Microsoft Learn documentation for SharePoint, Graph API, and Power Automate topics.

This tool helps Claude give ACCURATE, CURRENT answers instead of relying on potentially outdated training data.

ALWAYS use this tool before answering questions about:
- SharePoint column types, list operations, or API endpoints
- Power Automate connector capabilities, trigger types, or actions
- Graph API permissions, query parameters, or response formats
- Any SharePoint/Power Automate behavior you're not 100% sure about

Args:
  - query (string): Search topic, e.g. "lookup column", "calculated column formula", "SharePoint trigger"
  - scope (string, optional): "graph" | "power-automate" | "sharepoint" (default: "graph")
  - fetch_content (boolean, optional): Also fetch the full content of the top result (default: false)

Returns: Relevant documentation pages with titles, URLs, descriptions, and optionally full page content.`,
      inputSchema: {
        query: z.string().min(1).describe("Search query"),
        scope: z.enum(["graph", "power-automate", "sharepoint"]).default("graph").describe("Documentation scope"),
        fetch_content: z.boolean().default(false).describe("Fetch full content of top result"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async ({ query, scope, fetch_content }) => {
      try {
        const results = await searchMSDocs(query, scope, 5);

        let output = `# Documentation Search: "${query}" (scope: ${scope})\n\n`;

        if (results.length === 0) {
          output += "No results found. Try broadening your search or changing the scope.\n";
        } else {
          for (const doc of results) {
            output += `## ${doc.title}\n`;
            output += `URL: ${doc.url}\n`;
            if (doc.lastUpdatedDate) output += `Last updated: ${doc.lastUpdatedDate}\n`;
            output += `${doc.description}\n\n`;
          }
        }

        // Optionally fetch full content of top result
        if (fetch_content && results.length > 0) {
          output += `\n---\n# Full Content: ${results[0].title}\n\n`;
          const content = await fetchDocPage(results[0].url);
          output += content;
        }

        return {
          content: [{ type: "text", text: output }],
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Documentation search error: ${error instanceof Error ? error.message : String(error)}`,
          }],
        };
      }
    }
  );

  // ─── ms_docs_fetch ───────────────────────────────────────────────
  server.registerTool(
    "ms_docs_fetch",
    {
      title: "Fetch Documentation Page",
      description: `Fetch the full content of a specific Microsoft Learn documentation page.

Use this when you need the complete reference for a specific API endpoint or feature.

Args:
  - url (string): Full URL of the Microsoft Learn page

Returns: Extracted text content from the page (HTML stripped, truncated to ~15k chars if needed).`,
      inputSchema: {
        url: z.string().url().describe("Full URL of the documentation page"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async ({ url }) => {
      try {
        const content = await fetchDocPage(url);
        return {
          content: [{ type: "text", text: content }],
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error fetching page: ${error instanceof Error ? error.message : String(error)}`,
          }],
        };
      }
    }
  );

  // ─── ms_docs_context ─────────────────────────────────────────────
  server.registerTool(
    "ms_docs_context",
    {
      title: "Build Documentation Context",
      description: `Build a comprehensive documentation reference for a topic, combining search results from Graph API,
Power Automate, and SharePoint documentation plus key API endpoint reference.

Use this before starting any complex SharePoint/Power Automate task to ensure Claude has current,
accurate API reference material.

Args:
  - topic (string): The topic to build context for, e.g. "lookup columns between lists",
    "calculated column formulas", "Power Automate SharePoint triggers"

Returns: Combined documentation reference with relevant pages and API endpoint quick-reference.`,
      inputSchema: {
        topic: z.string().min(1).describe("Topic to build context for"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async ({ topic }) => {
      try {
        const context = await buildDocContext(topic);
        return {
          content: [{ type: "text", text: context }],
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error building context: ${error instanceof Error ? error.message : String(error)}`,
          }],
        };
      }
    }
  );
}
