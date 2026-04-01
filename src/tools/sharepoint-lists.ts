import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphFetch, graphFetchAll, resolveSiteId, GraphApiError } from "../services/graph-client.js";
import type { SPList, SPColumnDefinition, SPListItem, GraphResponse, ColumnCreatePayload } from "../types.js";

export function registerListTools(server: McpServer): void {

  // ─── sp_list_lists ───────────────────────────────────────────────
  server.registerTool(
    "sp_list_lists",
    {
      title: "List SharePoint Lists",
      description: `List all lists (and document libraries) on a SharePoint site.

Args:
  - site_path (string, optional): Site relative path, e.g. "/sites/ProjectDashboard". Uses default site if omitted.

Returns: Array of lists with id, displayName, description, webUrl, template, lastModified.

Use this first to discover what lists exist before working with columns or items.`,
      inputSchema: {
        site_path: z.string().optional().describe("Site path e.g. /sites/ProjectDashboard"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        const lists = await graphFetchAll<SPList>(`/sites/${siteId}/lists`);

        const output = lists.map((l) => ({
          id: l.id,
          displayName: l.displayName,
          description: l.description || "",
          webUrl: l.webUrl,
          template: l.list?.template || "unknown",
          lastModified: l.lastModifiedDateTime,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify(output, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_list_get ─────────────────────────────────────────────────
  server.registerTool(
    "sp_list_get",
    {
      title: "Get List Schema",
      description: `Get a SharePoint list's full schema including all column definitions, types, and relationships.

Args:
  - list_id (string): List ID or display name
  - site_path (string, optional): Site path
  - include_items (boolean, optional): Also return first 100 items. Default false.

Returns: List metadata + array of column definitions showing name, type, lookup targets, formulas, choices etc.

This is essential for understanding entity structure and relationships before making changes.`,
      inputSchema: {
        list_id: z.string().describe("List ID (GUID) or display name"),
        site_path: z.string().optional().describe("Site path"),
        include_items: z.boolean().default(false).describe("Include first 100 items"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ list_id, site_path, include_items }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        let expand = "columns";
        if (include_items) expand += ",items(expand=fields)";

        const list = await graphFetch<SPList & { columns?: SPColumnDefinition[]; items?: SPListItem[] }>(
          `/sites/${siteId}/lists/${list_id}?expand=${expand}`
        );

        const columns = (list.columns || []).map(formatColumn);

        const result: Record<string, unknown> = {
          id: list.id,
          displayName: list.displayName,
          description: list.description,
          webUrl: list.webUrl,
          template: list.list?.template,
          columns,
        };

        if (include_items && list.items) {
          result.items = list.items.map((item) => ({
            id: item.id,
            fields: item.fields,
            created: item.createdDateTime,
            modified: item.lastModifiedDateTime,
          }));
          result.itemCount = list.items.length;
        }

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_list_create ──────────────────────────────────────────────
  server.registerTool(
    "sp_list_create",
    {
      title: "Create SharePoint List",
      description: `Create a new SharePoint list with optional initial columns.

Args:
  - display_name (string): List display name
  - description (string, optional): List description
  - columns (array, optional): Initial columns to create. Each: { name, type, ...typeSpecificProps }
    Supported types: text, number, dateTime, choice, boolean, currency
    (Lookup columns must be added separately after list creation using sp_column_add)
  - site_path (string, optional): Site path

Returns: Created list metadata with ID.

Note: Lookup and calculated columns should be added after creation using sp_column_add.`,
      inputSchema: {
        display_name: z.string().min(1).describe("List display name"),
        description: z.string().optional().describe("List description"),
        columns: z.array(z.object({
          name: z.string().describe("Internal column name"),
          type: z.enum(["text", "number", "dateTime", "choice", "boolean", "currency"]).describe("Column type"),
          choices: z.array(z.string()).optional().describe("For choice type: array of options"),
          maxLength: z.number().optional().describe("For text type: max characters"),
          allowMultipleLines: z.boolean().optional().describe("For text type: allow multiline"),
        })).optional().describe("Initial columns"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async ({ display_name, description, columns, site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);

        const body: Record<string, unknown> = {
          displayName: display_name,
          list: { template: "genericList" },
        };
        if (description) body.description = description;

        if (columns && columns.length > 0) {
          body.columns = columns.map((col) => {
            const colDef: ColumnCreatePayload = { name: col.name };
            switch (col.type) {
              case "text":
                colDef.text = {
                  allowMultipleLines: col.allowMultipleLines || false,
                  appendChangesToExistingText: false,
                  linesForEditing: 0,
                  maxLength: col.maxLength || 255,
                };
                break;
              case "number":
                colDef.number = { decimalPlaces: "automatic", maximum: 1e10, minimum: 0 };
                break;
              case "dateTime":
                colDef.dateTime = { displayAs: "default", format: "dateTime" };
                break;
              case "choice":
                colDef.choice = {
                  allowTextEntry: false,
                  choices: col.choices || [],
                  displayAs: "dropDownMenu",
                };
                break;
              case "boolean":
                colDef.boolean = {};
                break;
              case "currency":
                colDef.currency = { locale: "en-GB" };
                break;
            }
            return colDef;
          });
        }

        const created = await graphFetch<SPList>(`/sites/${siteId}/lists`, "POST", body);

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              id: created.id,
              displayName: created.displayName,
              webUrl: created.webUrl,
              message: "List created successfully. Use sp_column_add to add lookup or calculated columns.",
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_column_add ───────────────────────────────────────────────
  server.registerTool(
    "sp_column_add",
    {
      title: "Add Column to List",
      description: `Add a column to an existing SharePoint list. Supports all column types including lookup and calculated.

Args:
  - list_id (string): Target list ID
  - name (string): Internal column name (no spaces)
  - display_name (string, optional): Display name (can have spaces)
  - type (string): One of: text, number, dateTime, choice, lookup, calculated, boolean, currency, personOrGroup
  - site_path (string, optional): Site path

Type-specific args:
  - For lookup: lookup_list_id (required), lookup_column_name (default "Title")
  - For calculated: formula (required, e.g. "=[Column1]*[Column2]"), output_type (text|number|dateTime|boolean|currency)
  - For choice: choices (string array), allow_text_entry (boolean)
  - For text: max_length (number), allow_multiple_lines (boolean)
  - For personOrGroup: allow_multiple (boolean)

Returns: Created column definition.

IMPORTANT for lookups: The lookup_list_id must be the GUID of the target list. Use sp_list_lists to find it first.`,
      inputSchema: {
        list_id: z.string().describe("Target list ID"),
        name: z.string().describe("Internal column name"),
        display_name: z.string().optional().describe("Display name"),
        type: z.enum(["text", "number", "dateTime", "choice", "lookup", "calculated", "boolean", "currency", "personOrGroup"]).describe("Column type"),
        // Lookup-specific
        lookup_list_id: z.string().optional().describe("For lookup: target list GUID"),
        lookup_column_name: z.string().optional().describe("For lookup: column name in target list (default: Title)"),
        // Calculated-specific
        formula: z.string().optional().describe("For calculated: formula e.g. =[Price]*[Quantity]"),
        output_type: z.enum(["text", "number", "dateTime", "boolean", "currency"]).optional().describe("For calculated: output type"),
        // Choice-specific
        choices: z.array(z.string()).optional().describe("For choice: array of options"),
        allow_text_entry: z.boolean().optional().describe("For choice: allow free text"),
        // Text-specific
        max_length: z.number().optional().describe("For text: max characters"),
        allow_multiple_lines: z.boolean().optional().describe("For text: multiline"),
        // PersonOrGroup-specific
        allow_multiple: z.boolean().optional().describe("For personOrGroup: allow multiple"),
        // General
        required: z.boolean().optional().describe("Is field required"),
        indexed: z.boolean().optional().describe("Index this column"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async (params) => {
      try {
        const siteId = await resolveSiteId(params.site_path);

        const body: ColumnCreatePayload = {
          name: params.name,
          displayName: params.display_name || params.name,
          enforceUniqueValues: false,
          indexed: params.indexed || false,
          required: params.required || false,
        };

        switch (params.type) {
          case "text":
            body.text = {
              allowMultipleLines: params.allow_multiple_lines || false,
              appendChangesToExistingText: false,
              linesForEditing: 0,
              maxLength: params.max_length || 255,
            };
            break;
          case "number":
            body.number = { decimalPlaces: "automatic", maximum: 1e10, minimum: 0 };
            break;
          case "dateTime":
            body.dateTime = { displayAs: "default", format: "dateTime" };
            break;
          case "boolean":
            body.boolean = {};
            break;
          case "currency":
            body.currency = { locale: "en-GB" };
            break;
          case "choice":
            if (!params.choices?.length) throw new Error("choices array is required for choice columns");
            body.choice = {
              allowTextEntry: params.allow_text_entry || false,
              choices: params.choices,
              displayAs: "dropDownMenu",
            };
            break;
          case "lookup":
            if (!params.lookup_list_id) throw new Error("lookup_list_id is required for lookup columns");
            body.lookup = {
              allowMultipleValues: false,
              columnName: params.lookup_column_name || "Title",
              listId: params.lookup_list_id,
            };
            break;
          case "calculated":
            if (!params.formula) throw new Error("formula is required for calculated columns");
            body.calculated = {
              formula: params.formula,
              outputType: params.output_type || "text",
            };
            break;
          case "personOrGroup":
            body.personOrGroup = {
              allowMultipleSelection: params.allow_multiple || false,
              chooseFromType: "peopleAndGroups",
              displayAs: "nameWithPresence",
            };
            break;
        }

        const created = await graphFetch<SPColumnDefinition>(
          `/sites/${siteId}/lists/${params.list_id}/columns`,
          "POST",
          body
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              column: formatColumn(created),
              message: `Column "${created.displayName}" (${params.type}) added to list.`,
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_items_query ──────────────────────────────────────────────
  server.registerTool(
    "sp_items_query",
    {
      title: "Query List Items",
      description: `Query items in a SharePoint list with optional OData filtering and field selection.

Args:
  - list_id (string): List ID
  - filter (string, optional): OData filter, e.g. "fields/Status eq 'Active'"
  - select_fields (string array, optional): Fields to return, e.g. ["Title", "Status", "DueDate"]
  - top (number, optional): Max items to return (default 100, max 5000)
  - orderby (string, optional): Sort field, e.g. "fields/Created desc"
  - site_path (string, optional): Site path

Returns: Array of items with their field values.

IMPORTANT: Filter and orderby work best on indexed columns. For non-indexed columns on large lists,
you may need the header Prefer: HonorNonIndexedQueriesWarningMayFailRandomly.
Lookup columns return as {ColumnName}LookupId (integer ID) in the fields — not the display value.`,
      inputSchema: {
        list_id: z.string().describe("List ID"),
        filter: z.string().optional().describe("OData filter expression"),
        select_fields: z.array(z.string()).optional().describe("Fields to select"),
        top: z.number().int().min(1).max(5000).default(100).describe("Max items"),
        orderby: z.string().optional().describe("Sort expression"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ list_id, filter, select_fields, top, orderby, site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        let url = `/sites/${siteId}/lists/${list_id}/items?expand=fields`;

        if (select_fields?.length) {
          url += `(select=${select_fields.join(",")})`;
        }
        if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
        if (top) url += `&$top=${top}`;
        if (orderby) url += `&$orderby=${encodeURIComponent(orderby)}`;

        const response = await graphFetch<GraphResponse<SPListItem>>(url);
        const items = (response.value || []).map((item) => ({
          id: item.id,
          fields: item.fields,
          created: item.createdDateTime,
          modified: item.lastModifiedDateTime,
        }));

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              count: items.length,
              hasMore: !!response["@odata.nextLink"],
              items,
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_item_create ──────────────────────────────────────────────
  server.registerTool(
    "sp_item_create",
    {
      title: "Create List Item",
      description: `Create a new item in a SharePoint list.

Args:
  - list_id (string): List ID
  - fields (object): Key-value pairs of field names and values.
    For lookup columns, use the {ColumnName}LookupId key with the integer ID.
    For person columns, use the {ColumnName}LookupId key with the user's integer ID.
  - site_path (string, optional): Site path

Returns: Created item with its ID and field values.`,
      inputSchema: {
        list_id: z.string().describe("List ID"),
        fields: z.record(z.unknown()).describe("Field name-value pairs"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async ({ list_id, fields, site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        const created = await graphFetch<SPListItem>(
          `/sites/${siteId}/lists/${list_id}/items`,
          "POST",
          { fields }
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              id: created.id,
              fields: created.fields,
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_item_update ──────────────────────────────────────────────
  server.registerTool(
    "sp_item_update",
    {
      title: "Update List Item",
      description: `Update fields on an existing SharePoint list item.

Args:
  - list_id (string): List ID
  - item_id (string): Item ID
  - fields (object): Field name-value pairs to update (only include fields you want to change)
  - site_path (string, optional): Site path

Returns: Confirmation of update.`,
      inputSchema: {
        list_id: z.string().describe("List ID"),
        item_id: z.string().describe("Item ID"),
        fields: z.record(z.unknown()).describe("Fields to update"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ list_id, item_id, fields, site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        await graphFetch(
          `/sites/${siteId}/lists/${list_id}/items/${item_id}/fields`,
          "PATCH",
          fields
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({ success: true, message: `Item ${item_id} updated.` }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );

  // ─── sp_item_delete ──────────────────────────────────────────────
  server.registerTool(
    "sp_item_delete",
    {
      title: "Delete List Item",
      description: `Delete an item from a SharePoint list.

Args:
  - list_id (string): List ID
  - item_id (string): Item ID to delete
  - site_path (string, optional): Site path

Returns: Confirmation of deletion.`,
      inputSchema: {
        list_id: z.string().describe("List ID"),
        item_id: z.string().describe("Item ID"),
        site_path: z.string().optional(),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    },
    async ({ list_id, item_id, site_path }) => {
      try {
        const siteId = await resolveSiteId(site_path);
        await graphFetch(`/sites/${siteId}/lists/${list_id}/items/${item_id}`, "DELETE");

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: `Item ${item_id} deleted.` }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleError(error) }] };
      }
    }
  );
}

// ── Helpers ──

function formatColumn(col: SPColumnDefinition): Record<string, unknown> {
  const result: Record<string, unknown> = {
    id: col.id,
    name: col.name,
    displayName: col.displayName,
    description: col.description || "",
    hidden: col.hidden || false,
    readOnly: col.readOnly || false,
    required: col.required || false,
    indexed: col.indexed || false,
  };

  // Determine type and add type-specific info
  if (col.text) { result.type = "text"; result.config = col.text; }
  else if (col.number) { result.type = "number"; result.config = col.number; }
  else if (col.dateTime) { result.type = "dateTime"; result.config = col.dateTime; }
  else if (col.choice) { result.type = "choice"; result.config = col.choice; }
  else if (col.lookup) { result.type = "lookup"; result.config = col.lookup; }
  else if (col.calculated) { result.type = "calculated"; result.config = col.calculated; }
  else if (col.boolean !== undefined) { result.type = "boolean"; }
  else if (col.currency) { result.type = "currency"; result.config = col.currency; }
  else if (col.personOrGroup) { result.type = "personOrGroup"; result.config = col.personOrGroup; }
  else { result.type = "unknown"; }

  return result;
}

function handleError(error: unknown): string {
  if (error instanceof GraphApiError) return error.toUserMessage();
  if (error instanceof Error) return `Error: ${error.message}`;
  return `Unknown error: ${String(error)}`;
}
