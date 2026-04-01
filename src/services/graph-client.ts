import { ConfidentialClientApplication } from "@azure/msal-node";
import { GRAPH_API_BASE, GRAPH_SCOPES, FLOW_API_BASE, FLOW_SCOPES } from "../constants.js";
import type { GraphResponse } from "../types.js";

// ── Configuration ──

interface GraphClientConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  spHostname: string;
  defaultSitePath?: string;
}

let config: GraphClientConfig;
let msalClient: ConfidentialClientApplication;
let cachedGraphToken: { token: string; expiresAt: number } | null = null;
let cachedFlowToken: { token: string; expiresAt: number } | null = null;

export function initGraphClient(cfg: GraphClientConfig): void {
  config = cfg;
  msalClient = new ConfidentialClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      clientSecret: cfg.clientSecret,
    },
  });
}

// ── Token acquisition ──

async function getGraphToken(): Promise<string> {
  if (cachedGraphToken && Date.now() < cachedGraphToken.expiresAt - 60000) {
    return cachedGraphToken.token;
  }
  const result = await msalClient.acquireTokenByClientCredential({
    scopes: GRAPH_SCOPES,
  });
  if (!result?.accessToken) throw new Error("Failed to acquire Graph API token");
  cachedGraphToken = {
    token: result.accessToken,
    expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600000,
  };
  return cachedGraphToken.token;
}

async function getFlowToken(): Promise<string> {
  if (cachedFlowToken && Date.now() < cachedFlowToken.expiresAt - 60000) {
    return cachedFlowToken.token;
  }
  const result = await msalClient.acquireTokenByClientCredential({
    scopes: FLOW_SCOPES,
  });
  if (!result?.accessToken) throw new Error("Failed to acquire Flow API token");
  cachedFlowToken = {
    token: result.accessToken,
    expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600000,
  };
  return cachedFlowToken.token;
}

// ── HTTP helpers ──

export async function graphFetch<T>(
  path: string,
  method: "GET" | "POST" | "PATCH" | "DELETE" = "GET",
  body?: unknown
): Promise<T> {
  const token = await getGraphToken();
  const url = path.startsWith("http") ? path : `${GRAPH_API_BASE}${path}`;

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  const res = await fetch(url, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const errorBody = await res.text();
    throw new GraphApiError(res.status, res.statusText, errorBody, path);
  }

  if (res.status === 204) return {} as T;
  return res.json() as Promise<T>;
}

export async function flowFetch<T>(
  path: string,
  method: "GET" | "POST" | "PATCH" | "DELETE" = "GET",
  body?: unknown
): Promise<T> {
  const token = await getFlowToken();
  const url = path.startsWith("http") ? path : `${FLOW_API_BASE}${path}`;

  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const errorBody = await res.text();
    throw new GraphApiError(res.status, res.statusText, errorBody, path);
  }

  if (res.status === 204) return {} as T;
  return res.json() as Promise<T>;
}

// ── Pagination helper ──

export async function graphFetchAll<T>(path: string): Promise<T[]> {
  const results: T[] = [];
  let url: string | undefined = path;

  while (url) {
    const response: GraphResponse<T> = await graphFetch<GraphResponse<T>>(url);
    if (response.value) results.push(...response.value);
    url = response["@odata.nextLink"];
  }

  return results;
}

// ── Site resolution ──

export async function resolveSiteId(sitePath?: string): Promise<string> {
  const hostname = config.spHostname;
  const path = sitePath || config.defaultSitePath;

  if (path) {
    const site = await graphFetch<{ id: string }>(
      `/sites/${hostname}:${path}`
    );
    return site.id;
  }

  // Root site
  const site = await graphFetch<{ id: string }>(`/sites/${hostname}`);
  return site.id;
}

export function getConfig(): GraphClientConfig {
  return config;
}

// ── Error class ──

export class GraphApiError extends Error {
  constructor(
    public status: number,
    public statusText: string,
    public body: string,
    public path: string
  ) {
    super(
      `Graph API Error ${status} ${statusText} on ${path}: ${body}`
    );
    this.name = "GraphApiError";
  }

  toUserMessage(): string {
    try {
      const parsed = JSON.parse(this.body);
      const msg = parsed?.error?.message || this.body;
      return `Error ${this.status}: ${msg}\n\nPath: ${this.path}\n\nCommon fixes:\n` +
        (this.status === 401
          ? "- Check AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID\n- Ensure admin consent was granted"
          : this.status === 403
          ? "- The app registration may lack required permissions\n- Ensure Sites.ReadWrite.All and Lists.ReadWrite.All are granted with admin consent"
          : this.status === 404
          ? "- Check site hostname and path\n- Verify list/item IDs exist"
          : `- Status ${this.status}: ${this.statusText}`);
    } catch {
      return `Error ${this.status} on ${this.path}: ${this.body}`;
    }
  }
}
