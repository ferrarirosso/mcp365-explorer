import { BACKEND_DEFAULTS, MODEL_DEPLOYMENT } from "./backendDefaults.js";

function parseIntWithDefault(value, fallback) {
  const parsed = parseInt(String(value || ""), 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseBoolean(value, fallback) {
  if (typeof value === "boolean") return value;
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return fallback;
  return normalized === "true";
}

// Sec-6 guard: production must use managed identity. WEBSITE_INSTANCE_ID is
// injected by Azure App Service / Functions in deployed environments and is
// absent locally, so this only fires on a real Function App.
function assertNoApiKeyInProduction() {
  if (process.env.AZURE_OPENAI_API_KEY && process.env.WEBSITE_INSTANCE_ID) {
    throw new Error(
      "AZURE_OPENAI_API_KEY is set on a deployed Function App. Production must use " +
        "managed identity. Remove the app setting and redeploy:\n" +
        "  az functionapp config appsettings delete --name <function-app> " +
        "--resource-group <rg> --setting-names AZURE_OPENAI_API_KEY"
    );
  }
}

// Sec-2 guard: production must enforce authorization. The deployer always
// sets REQUIRED_APP_ROLE; if it's missing on a deployed Function App the
// admin has explicitly removed authorization and we fail loudly.
// ALLOW_ANONYMOUS_AUTHZ=true is the deliberate opt-out (not recommended).
function assertAuthorizationConfigured() {
  if (
    process.env.WEBSITE_INSTANCE_ID &&
    !process.env.REQUIRED_APP_ROLE &&
    String(process.env.ALLOW_ANONYMOUS_AUTHZ || "").toLowerCase() !== "true"
  ) {
    throw new Error(
      "No REQUIRED_APP_ROLE configured on this deployed Function App. " +
        "Production deployments must enforce authorization via an app role. " +
        "Either redeploy (the deployer sets this automatically) or, if you " +
        "explicitly want an open backend, set ALLOW_ANONYMOUS_AUTHZ=true " +
        "as an app setting (not recommended)."
    );
  }
}

export function getRuntimeConfig() {
  assertNoApiKeyInProduction();
  assertAuthorizationConfigured();
  return {
    azureOpenAiEndpoint: (process.env.AZURE_OPENAI_ENDPOINT || "").replace(/\/$/, ""),
    azureOpenAiApiKey: process.env.AZURE_OPENAI_API_KEY || "",
    azureOpenAiApiVersion: process.env.AZURE_OPENAI_API_VERSION || BACKEND_DEFAULTS.apiVersion,
    requestLimits: {
      perMinute: parseIntWithDefault(
        process.env.REQUESTS_PER_MINUTE,
        BACKEND_DEFAULTS.requestLimits.perMinute
      ),
      perDay: parseIntWithDefault(
        process.env.REQUESTS_PER_DAY,
        BACKEND_DEFAULTS.requestLimits.perDay
      ),
    },
    cors: {
      allowedOrigin: (process.env.ALLOWED_ORIGIN || "").trim(),
      allowPermissiveLocalCors: parseBoolean(process.env.ALLOW_PERMISSIVE_LOCAL_CORS, false),
    },
    authorization: {
      requiredAppRole: (process.env.REQUIRED_APP_ROLE || "").trim(),
      allowAnonymousAuthz: parseBoolean(process.env.ALLOW_ANONYMOUS_AUTHZ, false),
    },
    model: {
      ...MODEL_DEPLOYMENT,
      deploymentName: process.env.AZURE_OPENAI_DEPLOYMENT || MODEL_DEPLOYMENT.deploymentName,
      defaultReasoningEffort:
        process.env.AZURE_OPENAI_REASONING_EFFORT || MODEL_DEPLOYMENT.defaultReasoningEffort,
      usesMaxCompletionTokens: parseBoolean(
        process.env.AZURE_OPENAI_USES_MAX_COMPLETION_TOKENS,
        MODEL_DEPLOYMENT.usesMaxCompletionTokens
      ),
    },
  };
}
