/**
 * Browser-compatible MCP client using raw fetch + JSON-RPC 2.0.
 * No @modelcontextprotocol/sdk dependency — just HTTP POST with JSON-RPC bodies.
 *
 * Protocol reference: MCP Streamable HTTP transport
 * - POST JSON-RPC to server URL
 * - Track Mcp-Session-Id header across requests
 * - Response may be application/json or text/event-stream (SSE)
 */

export const MCP_AUDIENCE = 'ea9ffc3e-8a23-4a7d-836d-234d7c7565c1';
export const GATEWAY_BASE = 'https://agent365.svc.cloud.microsoft/mcp/environments';
export const PROTOCOL_VERSION = '2025-03-26';

export interface IMcpTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
}

export interface IMcpPrompt {
  name: string;
  description?: string;
  arguments?: Array<{ name: string; description?: string; required?: boolean }>;
}

export interface IMcpResource {
  uri: string;
  name: string;
  description?: string;
  mimeType?: string;
}

export interface IMcpServerInfo {
  name: string;
  version: string;
}

export interface IMcpCapabilities {
  logging?: Record<string, unknown>;
  prompts?: Record<string, unknown>;
  resources?: Record<string, unknown>;
  tools?: Record<string, unknown>;
}

export interface IMcpConnectResult {
  protocolVersion: string;
  serverInfo: IMcpServerInfo;
  capabilities: IMcpCapabilities;
  sessionId: string | undefined;
  tools: IMcpTool[];
  prompts: IMcpPrompt[];
  resources: IMcpResource[];
}

export interface IMcpCallResult {
  content: Array<{ type: string; text?: string; [key: string]: unknown }>;
  isError?: boolean;
}

export interface IMcpLogEntry {
  timestamp: Date;
  level: 'info' | 'error' | 'debug' | 'warning';
  category: 'mcp' | 'http';
  direction: 'send' | 'receive' | 'none';
  method: string;
  message: string;
  data?: unknown;
}

export type McpLogHandler = (entry: IMcpLogEntry) => void;

export class McpBrowserClient {
  private serverUrl: string;
  private getToken: () => Promise<string>;
  private mcpSessionId: string | undefined;
  private requestCounter: number = 0;
  private onLog: McpLogHandler | undefined;

  constructor(config: {
    environmentId: string;
    serverId: string;
    getToken: () => Promise<string>;
    onLog?: McpLogHandler;
  }) {
    this.serverUrl = `${GATEWAY_BASE}/${config.environmentId}/servers/${config.serverId}`;
    this.getToken = config.getToken;
    this.onLog = config.onLog;
  }

  public getServerUrl(): string {
    return this.serverUrl;
  }

  /**
   * Connect to the MCP server and discover all capabilities.
   * Flow: initialize → notifications/initialized → tools/list → prompts/list → resources/list
   */
  public async connect(): Promise<IMcpConnectResult> {
    this.log('info', 'mcp', 'Connecting...', { serverUrl: this.serverUrl }, 'none', 'connect');

    // 1. Initialize
    const initResult = await this.sendRequest('initialize', {
      protocolVersion: PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: 'mcp365-explorer', version: '1.0.0' }
    }) as {
      protocolVersion: string;
      capabilities: IMcpCapabilities;
      serverInfo: IMcpServerInfo;
    };

    this.log('info', 'mcp', `Server: ${initResult.serverInfo.name} v${initResult.serverInfo.version}`, initResult, 'receive', 'initialize');

    // 2. Send initialized notification
    await this.sendNotification('notifications/initialized');
    this.log('debug', 'mcp', 'Sent initialized notification', undefined, 'send', 'notifications/initialized');

    const capabilities = initResult.capabilities || {};

    // 3. List tools (if capability advertised)
    let tools: IMcpTool[] = [];
    if (capabilities.tools) {
      const toolsResult = await this.sendRequest('tools/list') as { tools?: IMcpTool[] };
      tools = toolsResult.tools || [];
      this.log('info', 'mcp', `Discovered ${tools.length} tools`, tools.map(t => t.name), 'receive', 'tools/list');
    }

    // 4. List prompts (if capability advertised)
    let prompts: IMcpPrompt[] = [];
    if (capabilities.prompts) {
      try {
        const promptsResult = await this.sendRequest('prompts/list') as { prompts?: IMcpPrompt[] };
        prompts = promptsResult.prompts || [];
        this.log('info', 'mcp', `Discovered ${prompts.length} prompts`, prompts.map(p => p.name), 'receive', 'prompts/list');
      } catch (err) {
        this.log('warning', 'mcp', `prompts/list failed: ${(err as Error).message}`, undefined, 'receive', 'prompts/list');
      }
    }

    // 5. List resources (if capability advertised)
    let resources: IMcpResource[] = [];
    if (capabilities.resources) {
      try {
        const resourcesResult = await this.sendRequest('resources/list') as { resources?: IMcpResource[] };
        resources = resourcesResult.resources || [];
        this.log('info', 'mcp', `Discovered ${resources.length} resources`, resources.map(r => r.name), 'receive', 'resources/list');
      } catch (err) {
        this.log('warning', 'mcp', `resources/list failed: ${(err as Error).message}`, undefined, 'receive', 'resources/list');
      }
    }

    return {
      protocolVersion: initResult.protocolVersion,
      serverInfo: initResult.serverInfo,
      capabilities,
      sessionId: this.mcpSessionId,
      tools,
      prompts,
      resources
    };
  }

  /**
   * Get a prompt by name with arguments.
   */
  public async getPrompt(name: string, args?: Record<string, string>): Promise<unknown> {
    this.log('info', 'mcp', `Getting prompt: ${name}`, args, 'send', 'prompts/get');
    const result = await this.sendRequest('prompts/get', { name, arguments: args || {} });
    this.log('info', 'mcp', `Prompt ${name} received`, result, 'receive', 'prompts/get');
    return result;
  }

  /**
   * Read a resource by URI.
   */
  public async readResource(uri: string): Promise<unknown> {
    this.log('info', 'mcp', `Reading resource: ${uri}`, undefined, 'send', 'resources/read');
    const result = await this.sendRequest('resources/read', { uri });
    this.log('info', 'mcp', `Resource received`, result, 'receive', 'resources/read');
    return result;
  }

  /**
   * Call an MCP tool by name with arguments.
   */
  public async callTool(name: string, args: Record<string, unknown>): Promise<IMcpCallResult> {
    this.log('info', 'mcp', `Calling tool: ${name}`, args, 'send', 'tools/call');
    const startTime = Date.now();

    const result = await this.sendRequest('tools/call', { name, arguments: args }) as IMcpCallResult;

    const durationMs = Date.now() - startTime;
    this.log('info', 'mcp', `Tool ${name} completed (${durationMs}ms)`, result, 'receive', 'tools/call');

    return result;
  }

  /**
   * Send a JSON-RPC request (expects a response).
   */
  private async sendRequest(method: string, params?: unknown): Promise<unknown> {
    const id = ++this.requestCounter;
    const body = {
      jsonrpc: '2.0' as const,
      id,
      method,
      params: params || {}
    };

    this.log('debug', 'http', `POST ${method}`, body, 'send', method);
    const response = await this.doFetch(body);

    const contentType = response.headers.get('content-type') || '';

    if (contentType.indexOf('text/event-stream') !== -1) {
      return this.readSseResponse(response, method);
    }

    // Regular JSON response
    const json = await response.json();

    if (json.error) {
      const errMsg = json.error.message || JSON.stringify(json.error);
      this.log('error', 'http', `JSON-RPC error: ${errMsg}`, json.error, 'receive', method);
      throw new Error(`MCP error: ${errMsg}`);
    }

    this.log('debug', 'http', `HTTP ${response.status} OK`, undefined, 'receive', method);
    return json.result;
  }

  /**
   * Send a JSON-RPC notification (no response expected).
   */
  private async sendNotification(method: string, params?: unknown): Promise<void> {
    const body = {
      jsonrpc: '2.0' as const,
      method,
      params: params || {}
    };
    // Notifications may return 202 Accepted or 200 with empty body
    await this.doFetch(body);
  }

  /**
   * Execute the HTTP POST with auth and session headers.
   */
  private async doFetch(body: unknown): Promise<Response> {
    const token = await this.getToken();

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Accept': 'application/json, text/event-stream',
      'Authorization': `Bearer ${token}`
    };

    if (this.mcpSessionId) {
      headers['Mcp-Session-Id'] = this.mcpSessionId;
    }

    const response = await fetch(this.serverUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(body)
    });

    // Track session ID from response
    const sessionId = response.headers.get('mcp-session-id');
    if (sessionId) {
      this.mcpSessionId = sessionId;
    }

    if (!response.ok) {
      const errorText = await response.text();
      this.log('error', 'http', `HTTP ${response.status}: ${errorText}`, undefined, 'receive', '');
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    return response;
  }

  /**
   * Read a Server-Sent Events response and extract the JSON-RPC result.
   */
  private async readSseResponse(response: Response, method: string): Promise<unknown> {
    const text = await response.text();
    const lines = text.split('\n');
    let lastData = '';

    for (const line of lines) {
      if (line.indexOf('data: ') === 0) {
        lastData = line.substring(6);
      }
    }

    if (!lastData) {
      return {};
    }

    const json = JSON.parse(lastData);
    if (json.error) {
      const errMsg = json.error.message || JSON.stringify(json.error);
      this.log('error', 'http', `SSE error: ${errMsg}`, json.error, 'receive', method);
      throw new Error(`MCP error: ${errMsg}`);
    }

    this.log('debug', 'http', 'SSE response received', undefined, 'receive', method);
    return json.result;
  }

  private log(
    level: IMcpLogEntry['level'],
    category: IMcpLogEntry['category'],
    message: string,
    data?: unknown,
    direction: IMcpLogEntry['direction'] = 'none',
    method: string = ''
  ): void {
    if (this.onLog) {
      this.onLog({ timestamp: new Date(), level, category, direction, method, message, data });
    }
  }
}
