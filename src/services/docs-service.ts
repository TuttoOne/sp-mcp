/**
 * Documentation search service
 * Fetches latest Microsoft Graph / Power Automate docs on demand
 * so Claude has current info instead of stale training data.
 */

const MS_LEARN_SEARCH_API = "https://learn.microsoft.com/api/search";

interface DocSearchResult {
  title: string;
  url: string;
  description: string;
  lastUpdatedDate?: string;
}

interface MSLearnSearchResponse {
  results: Array<{
    title: string;
    url: string;
    descriptions: string[];
    lastUpdatedDate?: string;
  }>;
}

/**
 * Search Microsoft Learn documentation
 */
export async function searchMSDocs(
  query: string,
  scope: "graph" | "power-automate" | "sharepoint" = "graph",
  maxResults: number = 5
): Promise<DocSearchResult[]> {
  // Build scoped query
  const scopeMap: Record<string, string> = {
    graph: "Microsoft Graph API",
    "power-automate": "Power Automate",
    sharepoint: "SharePoint",
  };

  const fullQuery = `${scopeMap[scope]} ${query}`;

  // Use Microsoft Learn's search API
  const params = new URLSearchParams({
    search: fullQuery,
    locale: "en-us",
    $top: String(maxResults),
    facet: "category",
    "category": "Documentation",
  });

  try {
    const res = await fetch(`${MS_LEARN_SEARCH_API}?${params}`);
    if (!res.ok) {
      // Fallback: construct direct URLs for common topics
      return getFallbackDocs(query, scope);
    }
    const data = (await res.json()) as MSLearnSearchResponse;

    return (data.results || []).map((r) => ({
      title: r.title,
      url: r.url,
      description: r.descriptions?.[0] || "",
      lastUpdatedDate: r.lastUpdatedDate,
    }));
  } catch {
    return getFallbackDocs(query, scope);
  }
}

/**
 * Fetch the content of a Microsoft Learn page
 */
export async function fetchDocPage(url: string): Promise<string> {
  try {
    const res = await fetch(url, {
      headers: { Accept: "text/html" },
    });
    if (!res.ok) return `Failed to fetch ${url}: ${res.status}`;

    const html = await res.text();

    // Extract main content — MS Learn uses <main> tag
    const mainMatch = html.match(/<main[^>]*>([\s\S]*?)<\/main>/i);
    const content = mainMatch ? mainMatch[1] : html;

    // Strip HTML tags, normalize whitespace
    const text = content
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, " ")
      .trim();

    // Truncate to reasonable size for context
    if (text.length > 15000) {
      return text.slice(0, 15000) + "\n\n[Content truncated. Visit the full page for more details.]";
    }

    return text;
  } catch (error) {
    return `Error fetching doc page: ${error instanceof Error ? error.message : String(error)}`;
  }
}

/**
 * Fallback documentation links for common topics
 */
function getFallbackDocs(query: string, scope: string): DocSearchResult[] {
  const graphDocs: Record<string, DocSearchResult> = {
    lists: {
      title: "Working with SharePoint Lists via Graph API",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/list?view=graph-rest-1.0",
      description: "List resource type - create, read, update lists in SharePoint",
    },
    columns: {
      title: "Column definitions in SharePoint via Graph API",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0",
      description: "ColumnDefinition resource - text, number, lookup, calculated, choice columns",
    },
    lookup: {
      title: "Lookup column definition",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0",
      description: "lookupColumn resource type for creating relationships between lists",
    },
    items: {
      title: "Working with list items via Graph API",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/listitem?view=graph-rest-1.0",
      description: "listItem resource - CRUD operations on SharePoint list items",
    },
    sites: {
      title: "SharePoint sites via Graph API",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0",
      description: "Working with SharePoint sites in Microsoft Graph",
    },
    calculated: {
      title: "Calculated column definition",
      url: "https://learn.microsoft.com/en-us/graph/api/resources/calculatedcolumn?view=graph-rest-1.0",
      description: "calculatedColumn resource for formula-based columns",
    },
  };

  const paDocs: Record<string, DocSearchResult> = {
    flows: {
      title: "Work with cloud flows using code",
      url: "https://learn.microsoft.com/en-us/power-automate/manage-flows-with-code",
      description: "Manage Power Automate flows programmatically",
    },
    connectors: {
      title: "Power Automate Management connector",
      url: "https://learn.microsoft.com/en-us/connectors/flowmanagement/",
      description: "Management connector for listing, triggering, and managing flows",
    },
    triggers: {
      title: "Triggers in Power Automate",
      url: "https://learn.microsoft.com/en-us/power-automate/triggers-introduction",
      description: "Understanding triggers in Power Automate",
    },
    sharepoint: {
      title: "SharePoint connector for Power Automate",
      url: "https://learn.microsoft.com/en-us/connectors/sharepointonline/",
      description: "SharePoint connector actions and triggers",
    },
  };

  const docs = scope === "power-automate" ? paDocs : graphDocs;
  const queryLower = query.toLowerCase();

  // Find matching docs
  const matches = Object.entries(docs)
    .filter(([key]) => queryLower.includes(key) || key.includes(queryLower))
    .map(([, doc]) => doc);

  // If no matches, return all docs for the scope as reference
  if (matches.length === 0) {
    return Object.values(docs).slice(0, 3);
  }

  return matches;
}

/**
 * Build a comprehensive documentation context for a topic
 * Combines search results with key reference pages
 */
export async function buildDocContext(topic: string): Promise<string> {
  const lines: string[] = [
    `# Microsoft Documentation Reference: ${topic}`,
    "",
    "## Relevant Documentation Pages",
    "",
  ];

  // Search across scopes
  const [graphResults, paResults, spResults] = await Promise.all([
    searchMSDocs(topic, "graph", 3),
    searchMSDocs(topic, "power-automate", 2),
    searchMSDocs(topic, "sharepoint", 2),
  ]);

  const allResults = [...graphResults, ...paResults, ...spResults];

  for (const doc of allResults) {
    lines.push(`### ${doc.title}`);
    lines.push(`URL: ${doc.url}`);
    if (doc.lastUpdatedDate) lines.push(`Last updated: ${doc.lastUpdatedDate}`);
    lines.push(doc.description);
    lines.push("");
  }

  lines.push("## Key API Endpoints Reference");
  lines.push("");
  lines.push("**SharePoint Lists (Graph API v1.0):**");
  lines.push("- GET /sites/{site-id}/lists — List all lists");
  lines.push("- POST /sites/{site-id}/lists — Create a list");
  lines.push("- GET /sites/{site-id}/lists/{list-id}?expand=columns,items(expand=fields) — Get list with schema and data");
  lines.push("- POST /sites/{site-id}/lists/{list-id}/columns — Create a column");
  lines.push("- POST /sites/{site-id}/lists/{list-id}/items — Create an item");
  lines.push("- PATCH /sites/{site-id}/lists/{list-id}/items/{item-id}/fields — Update item fields");
  lines.push("");
  lines.push("**Lookup columns** require: `{ name, lookup: { columnName, listId } }` in the column definition POST body.");
  lines.push("**Calculated columns** require: `{ name, calculated: { formula, outputType } }` where outputType is 'boolean', 'currency', 'dateTime', 'number', or 'text'.");
  lines.push("");
  lines.push("**Power Automate (api.flow.microsoft.com):**");
  lines.push("- GET /environments/{env-id}/flows — List flows");
  lines.push("- GET /environments/{env-id}/flows/{flow-id} — Get flow details");
  lines.push("- POST /environments/{env-id}/flows/{flow-id}/triggers/{trigger}/run — Trigger a flow");
  lines.push("- GET /environments/{env-id}/flows/{flow-id}/runs — List flow runs");
  lines.push("- POST /environments/{env-id}/flows/{flow-id}/runs/{run-id}/cancel — Cancel a run");

  return lines.join("\n");
}
