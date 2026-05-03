import { app } from "@azure/functions";
import { DefaultAzureCredential } from "@azure/identity";
import { getRuntimeConfig } from "../config/runtimeConfig.js";

// =============================================================================
// CONFIGURATION
// =============================================================================

const runtimeConfig = getRuntimeConfig();
const AZURE_OPENAI_ENDPOINT = runtimeConfig.azureOpenAiEndpoint;
const AZURE_OPENAI_API_KEY = runtimeConfig.azureOpenAiApiKey;
const AZURE_OPENAI_API_VERSION = runtimeConfig.azureOpenAiApiVersion;
const REQUESTS_PER_MINUTE = runtimeConfig.requestLimits.perMinute;
const REQUESTS_PER_DAY = runtimeConfig.requestLimits.perDay;
const ALLOWED_ORIGIN = runtimeConfig.cors.allowedOrigin;
const ALLOW_PERMISSIVE_LOCAL_CORS = runtimeConfig.cors.allowPermissiveLocalCors;
const REQUIRED_APP_ROLE = runtimeConfig.authorization.requiredAppRole;
const ALLOW_ANONYMOUS_AUTHZ = runtimeConfig.authorization.allowAnonymousAuthz;
const MODEL = runtimeConfig.model;

// =============================================================================
// MANAGED IDENTITY TOKEN CACHE
// =============================================================================

const AZURE_OPENAI_SCOPE = "https://cognitiveservices.azure.com/.default";
let credential = null;
let tokenCache = { token: null, expiresAt: 0 };

async function getAzureToken() {
  const now = Date.now();
  if (tokenCache.token && tokenCache.expiresAt > now + 5 * 60 * 1000) {
    return tokenCache.token;
  }
  if (!credential) {
    credential = new DefaultAzureCredential();
  }
  const tokenResponse = await credential.getToken(AZURE_OPENAI_SCOPE);
  tokenCache = { token: tokenResponse.token, expiresAt: tokenResponse.expiresOnTimestamp };
  return tokenCache.token;
}

// =============================================================================
// CALLER IDENTITY — Easy Auth principal is required (function key auth removed)
// =============================================================================

const rateLimitStore = new Map();

function decodePrincipal(request) {
  const principalHeader = request.headers.get("x-ms-client-principal");
  if (!principalHeader) return null;
  try {
    return JSON.parse(Buffer.from(principalHeader, "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function extractEasyAuthObjectId(request) {
  const explicitObjectId = request.headers.get("x-ms-client-principal-id");
  if (explicitObjectId && explicitObjectId.trim()) {
    return explicitObjectId.trim();
  }

  const decoded = decodePrincipal(request);
  if (!decoded) return undefined;

  const claims = Array.isArray(decoded.claims) ? decoded.claims : [];
  for (let i = 0; i < claims.length; i++) {
    const type = String(claims[i]?.typ || "").toLowerCase();
    const value = String(claims[i]?.val || "");
    if (!value) continue;
    if (
      type === "http://schemas.microsoft.com/identity/claims/objectidentifier" ||
      type === "oid"
    ) {
      return value;
    }
  }
  return undefined;
}

// Sec-2: extract every `roles` claim from the principal. Multiple `roles`
// claims are emitted (one per assigned role), so iterate the whole array.
function extractRoles(request) {
  const decoded = decodePrincipal(request);
  if (!decoded) return [];
  const claims = Array.isArray(decoded.claims) ? decoded.claims : [];
  const roles = [];
  for (const claim of claims) {
    if (String(claim?.typ || "").toLowerCase() === "roles" && claim?.val) {
      roles.push(String(claim.val));
    }
  }
  return roles;
}

function checkAuthorization(request) {
  if (ALLOW_ANONYMOUS_AUTHZ) {
    // Explicit opt-out (set via app setting). Defended at startup so this
    // only fires in dev / explicit-opt-out scenarios.
    return { allowed: true };
  }
  if (!REQUIRED_APP_ROLE) {
    // Local dev with no role configured — allow. The startup check in
    // runtimeConfig prevents this branch from being reached in production.
    return { allowed: true };
  }
  const roles = extractRoles(request);
  if (!roles.includes(REQUIRED_APP_ROLE)) {
    return {
      allowed: false,
      error:
        "Caller is authenticated but not authorized. Assign the user (or their group) " +
        "the '" +
        REQUIRED_APP_ROLE +
        "' role on the Backend API in Entra → Enterprise applications.",
    };
  }
  return { allowed: true };
}

function checkRateLimit(callerId) {
  const now = Date.now();
  const minuteAgo = now - 60 * 1000;
  const dayAgo = now - 24 * 60 * 60 * 1000;

  const requests = rateLimitStore.get(callerId) || [];
  const validRequests = requests.filter((ts) => ts > dayAgo);
  rateLimitStore.set(callerId, validRequests);

  const minuteCount = validRequests.filter((ts) => ts > minuteAgo).length;
  const dayCount = validRequests.length;

  if (minuteCount >= REQUESTS_PER_MINUTE) {
    return { allowed: false, error: `Rate limit exceeded. Max ${REQUESTS_PER_MINUTE} req/min.` };
  }
  if (dayCount >= REQUESTS_PER_DAY) {
    return { allowed: false, error: `Daily limit exceeded. Max ${REQUESTS_PER_DAY} req/day.` };
  }

  validRequests.push(now);
  return { allowed: true, minuteCount: minuteCount + 1, dayCount: dayCount + 1 };
}

// =============================================================================
// CORS HELPERS
// =============================================================================

function isLocalOrigin(origin) {
  return /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(origin);
}

function stripTrailingSlash(value) {
  if (!value) return value;
  return value.replace(/\/+$/, "");
}

function isOriginAllowed(origin) {
  if (!origin) return false;
  if (ALLOW_PERMISSIVE_LOCAL_CORS && isLocalOrigin(origin)) return true;
  return !!ALLOWED_ORIGIN && stripTrailingSlash(origin) === stripTrailingSlash(ALLOWED_ORIGIN);
}

function getCorsHeaders(request) {
  const origin = request.headers.get("origin");
  const headers = {};
  if (isOriginAllowed(origin)) {
    headers["Access-Control-Allow-Origin"] = origin;
    headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
    headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization";
    headers["Access-Control-Max-Age"] = "3600";
    // Expose response headers to the browser so the agent can honour
    // 429 backoff signals from Azure OpenAI / Foundry. Without this list
    // fetch() returns those headers as `null`.
    headers["Access-Control-Expose-Headers"] =
      "Retry-After, X-Correlation-Id, X-RateLimit-Remaining-Minute, X-RateLimit-Remaining-Day, " +
      "X-RateLimit-Reset-Requests, X-RateLimit-Reset-Tokens";
  }
  return headers;
}

// Forward Azure OpenAI's rate-limit signal headers (Retry-After plus the
// AOAI-specific ones) so the browser-side agent can honour them. Header
// names from upstream `Headers` are case-insensitive; we re-emit them in
// the canonical CamelCase the consumer expects.
function forwardRateLimitHeaders(upstream) {
  const out = {};
  const retryAfter = upstream.headers.get("retry-after");
  if (retryAfter) out["Retry-After"] = retryAfter;
  const resetReq = upstream.headers.get("x-ratelimit-reset-requests");
  if (resetReq) out["X-RateLimit-Reset-Requests"] = resetReq;
  const resetTok = upstream.headers.get("x-ratelimit-reset-tokens");
  if (resetTok) out["X-RateLimit-Reset-Tokens"] = resetTok;
  return out;
}

function handlePreflight(request) {
  if (request.method === "OPTIONS") {
    return { status: 200, body: "", headers: getCorsHeaders(request) };
  }
  return null;
}

// =============================================================================
// AZURE OPENAI
// =============================================================================

async function getBackendAuthHeaders(context) {
  if (AZURE_OPENAI_API_KEY) {
    return { "api-key": AZURE_OPENAI_API_KEY };
  }
  try {
    const token = await getAzureToken();
    return { Authorization: `Bearer ${token}` };
  } catch (error) {
    context.error(`Managed identity token error: ${error.message}`);
    throw new Error(
      "No API key configured and managed identity token failed. " +
        "Set AZURE_OPENAI_API_KEY (local dev only) or enable system-assigned managed identity."
    );
  }
}

function buildUpstreamChatBody(requestBody) {
  const body = { ...(requestBody && typeof requestBody === "object" ? requestBody : {}), model: MODEL.deploymentName };
  if (MODEL.defaultReasoningEffort && typeof body.reasoning_effort !== "string") {
    body.reasoning_effort = MODEL.defaultReasoningEffort;
  }
  if (
    MODEL.usesMaxCompletionTokens &&
    typeof body.max_completion_tokens !== "number" &&
    typeof body.max_tokens === "number"
  ) {
    body.max_completion_tokens = body.max_tokens;
    delete body.max_tokens;
  }
  return body;
}

async function callAzureOpenAi(body, context) {
  if (!AZURE_OPENAI_ENDPOINT) {
    throw new Error("LLM backend not configured.");
  }
  const targetUrl =
    `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${MODEL.deploymentName}/chat/completions` +
    `?api-version=${AZURE_OPENAI_API_VERSION}`;
  const authHeaders = await getBackendAuthHeaders(context);
  return fetch(targetUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify(body),
  });
}

function rateLimitExceededResponse(request, errorMessage) {
  // When *our own* per-caller rate limit fires, hint at when capacity returns.
  // 60s is a safe default for the per-minute window; the daily limit doesn't
  // expose a reset, so we still send the same hint — caller sees it as a
  // generic backoff instruction.
  const headers = { ...getCorsHeaders(request), "Retry-After": "60" };
  return { status: 429, headers, jsonBody: { error: errorMessage } };
}

function unauthenticatedResponse(request) {
  // Easy Auth should reject unauthenticated traffic before the handler runs;
  // this is the defensive net inside the function in case the platform
  // middleware is misconfigured.
  return {
    status: 401,
    headers: getCorsHeaders(request),
    jsonBody: { error: "Authentication required." },
  };
}

function unauthorizedResponse(request, errorMessage) {
  return {
    status: 403,
    headers: getCorsHeaders(request),
    jsonBody: { error: errorMessage },
  };
}

// =============================================================================
// HANDLERS
// =============================================================================

async function chatCompletionsHandler(request, context) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const callerId = extractEasyAuthObjectId(request);
  if (!callerId) {
    return unauthenticatedResponse(request);
  }

  const authzCheck = checkAuthorization(request);
  if (!authzCheck.allowed) {
    context.warn(`Authorization failed for caller: ${callerId}`);
    return unauthorizedResponse(request, authzCheck.error);
  }

  const corsHeaders = getCorsHeaders(request);
  const rateCheck = checkRateLimit(`easyauth:${callerId}`);
  if (!rateCheck.allowed) {
    context.warn(`Rate limit exceeded for caller: ${callerId}`);
    return rateLimitExceededResponse(request, rateCheck.error);
  }

  if (!AZURE_OPENAI_ENDPOINT) {
    context.error("Azure OpenAI endpoint not configured");
    return { status: 503, headers: corsHeaders, jsonBody: { error: "LLM backend not configured." } };
  }

  // Server-side correlation id (caller does not see backend internals).
  const correlationId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  context.log(
    `chat/completions cid=${correlationId} caller=${callerId} ` +
      `minute=${rateCheck.minuteCount}/${REQUESTS_PER_MINUTE} day=${rateCheck.dayCount}/${REQUESTS_PER_DAY}`
  );

  try {
    const body = await request.json();
    const response = await callAzureOpenAi(buildUpstreamChatBody(body), context);
    const responseContentType = response.headers.get("content-type") || "";
    const responseText = await response.text();

    // Browser-facing headers: correlation id, our per-caller rate-limit
    // hints, and the upstream rate-limit signals (Retry-After + AOAI-
    // specific reset headers) so the browser agent can back off correctly
    // when the model deployment throttles.
    const responseHeaders = {
      ...corsHeaders,
      ...forwardRateLimitHeaders(response),
      "X-Correlation-Id": correlationId,
      "X-RateLimit-Remaining-Minute": String(REQUESTS_PER_MINUTE - rateCheck.minuteCount),
      "X-RateLimit-Remaining-Day": String(REQUESTS_PER_DAY - rateCheck.dayCount),
    };

    if (responseContentType.includes("application/json")) {
      try {
        const data = JSON.parse(responseText);
        if (data.usage?.total_tokens) {
          context.log(`cid=${correlationId} tokens=${data.usage.total_tokens}`);
        }
        return { status: response.status, jsonBody: data, headers: responseHeaders };
      } catch {
        context.warn(`cid=${correlationId} upstream returned invalid JSON`);
      }
    }

    return {
      status: response.status,
      body: responseText,
      headers: { ...responseHeaders, ...(responseContentType ? { "Content-Type": responseContentType } : {}) },
    };
  } catch (error) {
    context.error(`cid=${correlationId} proxy error: ${error.message}`);
    return {
      status: 502,
      headers: { ...corsHeaders, "X-Correlation-Id": correlationId },
      jsonBody: { error: "Proxy error.", correlationId },
    };
  }
}

async function healthHandler(request) {
  // Public readiness only: no endpoint, deployment, model, or auth method.
  const preflight = handlePreflight(request);
  if (preflight) return preflight;
  const corsHeaders = getCorsHeaders(request);
  return {
    headers: corsHeaders,
    jsonBody: { status: "ok", time: new Date().toISOString() },
  };
}

// =============================================================================
// ROUTES
// =============================================================================
//
// authLevel is anonymous because Easy Auth (configured by the deployer in
// pass-through-required mode) is the actual gate. The handler defensively
// 401s if no Easy Auth principal is present.

app.http("health", {
  methods: ["GET", "OPTIONS"],
  authLevel: "anonymous",
  route: "health",
  handler: healthHandler,
});

app.http("chatCompletions", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  route: "chat/completions",
  handler: chatCompletionsHandler,
});
