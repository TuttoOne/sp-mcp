// ── SharePoint types ──

export interface SPSite {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
}

export interface SPList {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  list?: {
    template: string;
    contentTypesEnabled: boolean;
  };
}

export interface SPColumnDefinition {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  type?: string;
  indexed?: boolean;
  enforceUniqueValues?: boolean;
  hidden?: boolean;
  readOnly?: boolean;
  required?: boolean;
  // Type-specific properties
  text?: { allowMultipleLines: boolean; maxLength: number };
  number?: { decimalPlaces: string; maximum: number; minimum: number };
  dateTime?: { displayAs: string; format: string };
  choice?: { allowTextEntry: boolean; choices: string[]; displayAs: string };
  lookup?: { allowMultipleValues: boolean; columnName: string; listId: string; primaryLookupColumnId?: string };
  calculated?: { formula: string; outputType: string };
  boolean?: Record<string, unknown>;
  currency?: { locale: string };
  personOrGroup?: { allowMultipleSelection: boolean; chooseFromType: string; displayAs: string };
}

export interface SPListItem {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl?: string;
  fields?: Record<string, unknown>;
}

// ── Power Automate types ──

export interface PAFlow {
  name: string; // GUID
  id: string;
  type: string;
  properties: {
    displayName: string;
    state: string;
    createdTime: string;
    lastModifiedTime: string;
    environment?: { name: string };
    definitionSummary?: {
      triggers: Array<{ type: string; kind?: string }>;
      actions: Array<{ type: string }>;
    };
  };
}

export interface PAFlowRun {
  name: string;
  id: string;
  type: string;
  properties: {
    startTime: string;
    endTime?: string;
    status: string;
    trigger: {
      name: string;
      outputsLink?: { uri: string };
    };
    error?: {
      code: string;
      message: string;
    };
  };
}

// ── Graph API response wrapper ──

export interface GraphResponse<T> {
  "@odata.context"?: string;
  "@odata.nextLink"?: string;
  value?: T[];
}

// ── Column creation types ──

export type ColumnType =
  | "text"
  | "number"
  | "dateTime"
  | "choice"
  | "lookup"
  | "calculated"
  | "boolean"
  | "currency"
  | "personOrGroup";

export interface ColumnCreatePayload {
  name: string;
  displayName?: string;
  description?: string;
  enforceUniqueValues?: boolean;
  indexed?: boolean;
  required?: boolean;
  hidden?: boolean;
  [key: string]: unknown;
}
