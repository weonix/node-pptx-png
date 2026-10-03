/** Latest MCP protocol revision this server implements. */
export declare const LATEST_PROTOCOL_VERSION = "2025-06-18";
/** Protocol revisions the server accepts (echoed back when requested). */
export declare const SUPPORTED_PROTOCOL_VERSIONS: readonly string[];
/** JSON-RPC 2.0 error codes used by this server. */
export declare const JSONRPC_ERRORS: {
    readonly PARSE_ERROR: -32700;
    readonly INVALID_REQUEST: -32600;
    readonly METHOD_NOT_FOUND: -32601;
    readonly INVALID_PARAMS: -32602;
    readonly INTERNAL_ERROR: -32603;
};
/** A JSON-RPC response or error message produced by the server. */
export interface JsonRpcResponse {
    jsonrpc: '2.0';
    id: string | number | null;
    result?: unknown;
    error?: {
        code: number;
        message: string;
        data?: unknown;
    };
}
/**
 * Thrown by argument validation; mapped to a JSON-RPC -32602 (invalid
 * params) error so misuse is distinguishable from tool execution failures
 * (which are reported in-band with `isError: true` so the calling model
 * can read and correct them).
 */
export declare class McpInvalidParamsError extends Error {
}
/**
 * Parses a slides list like "1-5,8" into 1-based slide numbers.
 *
 * Minimal duplicate of the CLI's private parseSlideList (src/cli.ts does
 * not export it); throws {@link McpInvalidParamsError} instead of the
 * CLI's usage error.
 */
export declare function parseSlideList(value: string): number[];
/** Options for {@link McpStdioServer}. */
export interface McpServerOptions {
    /** Input stream (default: process.stdin). */
    input?: NodeJS.ReadableStream;
    /** Output stream (default: process.stdout). */
    output?: NodeJS.WritableStream;
    /**
     * Called after the input stream ends and all in-flight requests have
     * been answered (default: exit the process with code 0, per the MCP
     * stdio shutdown model).
     */
    onClose?: () => void;
}
/**
 * MCP server over newline-delimited JSON-RPC 2.0 (the MCP stdio
 * transport). Construct and call {@link start}; or, for in-process tests,
 * call {@link handleMessage} directly with decoded JSON-RPC messages.
 */
export declare class McpStdioServer {
    private readonly input;
    private readonly output;
    private readonly onClose;
    private readonly serverVersion;
    private buffered;
    /** Serializes message handling so responses keep request order. */
    private queue;
    constructor(options?: McpServerOptions);
    /** Starts reading newline-delimited JSON-RPC messages from the input. */
    start(): void;
    /** Splits buffered input into complete lines and queues each. */
    private drainBuffer;
    /** Queues one raw line for sequential handling. */
    private enqueueLine;
    /** Parses one line; returns the response message, if any. */
    private handleLine;
    /**
     * Handles one decoded JSON-RPC message. Returns the response for
     * requests, or undefined for notifications (and for client responses,
     * which this server never solicits). Exposed for in-process tests.
     */
    handleMessage(message: unknown): Promise<JsonRpcResponse | undefined>;
    /** Builds the initialize result (version negotiation + capabilities). */
    private handleInitialize;
    /**
     * Dispatches tools/call. Unknown tool names and invalid arguments are
     * JSON-RPC errors (thrown); tool execution failures are reported
     * in-band with isError so the calling model can read them.
     */
    private handleToolCall;
    /** Writes one message as a single line (MCP stdio framing). */
    private writeMessage;
}
//# sourceMappingURL=server.d.ts.map