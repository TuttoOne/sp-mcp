import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { flowFetch, GraphApiError } from "../services/graph-client.js";
import type { PAFlow, PAFlowRun } from "../types.js";

export function registerFlowTools(server: McpServer): void {

  // ─── pa_flows_list ───────────────────────────────────────────────
  server.registerTool(
    "pa_flows_list",
    {
      title: "List Power Automate Flows",
      description: `List all Power Automate flows in an environment.

Args:
  - environment_id (string): Power Platform environment ID
  - filter (string, optional): Filter by state, e.g. "properties/state eq 'Started'"

Returns: Array of flows with name (GUID), displayName, state, triggers, creation/modification dates.

To find your environment ID: Go to admin.powerplatform.microsoft.com > Environments, or use the
Power Automate Management connector's "List Environments" action.

NOTE: This uses the Flow Management API (api.flow.microsoft.com) which requires separate
authentication. If you get auth errors, the app may need Power Automate service permissions.`,
      inputSchema: {
        environment_id: z.string().describe("Power Platform environment ID"),
        filter: z.string().optional().describe("OData filter"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async ({ environment_id, filter }) => {
      try {
        let url = `/environments/${environment_id}/flows?api-version=2016-11-01`;
        if (filter) url += `&$filter=${encodeURIComponent(filter)}`;

        const response = await flowFetch<{ value: PAFlow[] }>(url);

        const flows = (response.value || []).map((f) => ({
          flowId: f.name,
          displayName: f.properties.displayName,
          state: f.properties.state,
          created: f.properties.createdTime,
          lastModified: f.properties.lastModifiedTime,
          triggers: f.properties.definitionSummary?.triggers?.map((t) => ({
            type: t.type,
            kind: t.kind,
          })) || [],
          actionCount: f.properties.definitionSummary?.actions?.length || 0,
        }));

        return {
          content: [{
            type: "text",
            text: JSON.stringify({ count: flows.length, flows }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── pa_flow_get ─────────────────────────────────────────────────
  server.registerTool(
    "pa_flow_get",
    {
      title: "Get Flow Details",
      description: `Get detailed information about a specific Power Automate flow including its full definition.

Args:
  - environment_id (string): Environment ID
  - flow_id (string): Flow ID (GUID from pa_flows_list)

Returns: Complete flow metadata, trigger configuration, action definitions, and connection references.
This gives you the full JSON definition of the flow.`,
      inputSchema: {
        environment_id: z.string().describe("Environment ID"),
        flow_id: z.string().describe("Flow ID (GUID)"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ environment_id, flow_id }) => {
      try {
        const flow = await flowFetch<PAFlow>(
          `/environments/${environment_id}/flows/${flow_id}?api-version=2016-11-01`
        );

        return {
          content: [{ type: "text", text: JSON.stringify(flow, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── pa_flow_runs ────────────────────────────────────────────────
  server.registerTool(
    "pa_flow_runs",
    {
      title: "List Flow Runs",
      description: `Get recent run history for a Power Automate flow. Shows status, timing, and errors.

Args:
  - environment_id (string): Environment ID
  - flow_id (string): Flow ID
  - top (number, optional): Number of runs to return (default 25, max 50)
  - status_filter (string, optional): Filter by status: "Succeeded", "Failed", "Running", "Cancelled"

Returns: Array of runs with status, start/end times, trigger info, and error details for failed runs.

Use this to diagnose flow failures — the error code and message fields are crucial.`,
      inputSchema: {
        environment_id: z.string().describe("Environment ID"),
        flow_id: z.string().describe("Flow ID"),
        top: z.number().int().min(1).max(50).default(25).describe("Number of runs"),
        status_filter: z.enum(["Succeeded", "Failed", "Running", "Cancelled"]).optional().describe("Filter by run status"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ environment_id, flow_id, top, status_filter }) => {
      try {
        let url = `/environments/${environment_id}/flows/${flow_id}/runs?api-version=2016-11-01&$top=${top}`;
        if (status_filter) {
          url += `&$filter=status eq '${status_filter}'`;
        }

        const response = await flowFetch<{ value: PAFlowRun[] }>(url);

        const runs = (response.value || []).map((r) => ({
          runId: r.name,
          status: r.properties.status,
          startTime: r.properties.startTime,
          endTime: r.properties.endTime || null,
          durationMs: r.properties.endTime
            ? new Date(r.properties.endTime).getTime() - new Date(r.properties.startTime).getTime()
            : null,
          triggerName: r.properties.trigger?.name,
          error: r.properties.error || null,
        }));

        const summary = {
          total: runs.length,
          succeeded: runs.filter((r) => r.status === "Succeeded").length,
          failed: runs.filter((r) => r.status === "Failed").length,
          running: runs.filter((r) => r.status === "Running").length,
        };

        return {
          content: [{
            type: "text",
            text: JSON.stringify({ summary, runs }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── pa_flow_trigger ─────────────────────────────────────────────
  server.registerTool(
    "pa_flow_trigger",
    {
      title: "Trigger a Flow",
      description: `Manually trigger a Power Automate flow (only works for flows with manual/HTTP triggers).

Args:
  - environment_id (string): Environment ID
  - flow_id (string): Flow ID
  - trigger_name (string, optional): Trigger name (default "manual")
  - body (object, optional): Trigger body/payload

Returns: Run ID and status.

This only works for flows that have a "Manually trigger a flow" or HTTP request trigger.
SharePoint-triggered or scheduled flows cannot be triggered this way.`,
      inputSchema: {
        environment_id: z.string().describe("Environment ID"),
        flow_id: z.string().describe("Flow ID"),
        trigger_name: z.string().default("manual").describe("Trigger name"),
        body: z.record(z.unknown()).optional().describe("Trigger payload"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    async ({ environment_id, flow_id, trigger_name, body }) => {
      try {
        const result = await flowFetch<Record<string, unknown>>(
          `/environments/${environment_id}/flows/${flow_id}/triggers/${trigger_name}/run?api-version=2016-11-01`,
          "POST",
          body || {}
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              message: "Flow triggered successfully",
              result,
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );
}

function handleError(error: unknown): string {
  if (error instanceof GraphApiError) return error.toUserMessage();
  if (error instanceof Error) return `Error: ${error.message}`;
  return `Unknown error: ${String(error)}`;
}
