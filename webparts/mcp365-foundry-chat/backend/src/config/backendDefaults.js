// Runtime defaults for the proxy. Used as fallbacks when the corresponding
// app settings are missing. Deployer-side defaults (resource group name,
// region, function app name, …) live in the webpart's deploy.config.json.

export const BACKEND_DEFAULTS = {
  apiVersion: "2025-01-01-preview",
  requestLimits: {
    perMinute: 30,
    perDay: 1000,
  },
};

export const MODEL_DEPLOYMENT = {
  label: "Reasoning",
  deploymentName: "mcp365-fc-gpt5mini",
  modelName: "gpt-5-mini",
  modelVersion: "2025-08-07",
  modelFormat: "OpenAI",
  skuName: "GlobalStandard",
  skuCapacity: 40,
  defaultReasoningEffort: "low",
  usesMaxCompletionTokens: true,
};
