#!/usr/bin/env node
/**
 * node-pptx-png MCP server entry point.
 *
 * Starts the MCP stdio server (newline-delimited JSON-RPC 2.0 on
 * stdin/stdout). Register it with an MCP client (Claude Desktop, Claude
 * Code, etc.) as:
 *
 *   node /path/to/node_modules/node-pptx-png/dist/mcp/cli.js
 *
 * See docs/mcp.md for client configuration and the tool reference.
 * The process exits when its stdin closes (the MCP stdio shutdown model).
 */
import * as path from 'node:path';
import { McpStdioServer } from './server.js';
/**
 * Detects direct execution without import.meta (which the CJS transpile
 * cannot parse): the server starts when the executed script is this file
 * or an installed bin alias.
 */
function isMainModule() {
    const entry = process.argv[1];
    if (!entry) {
        return false;
    }
    const base = path.basename(entry);
    return base === 'cli.js' || base === 'cli.ts' || base === 'node-pptx-png-mcp';
}
if (isMainModule()) {
    const server = new McpStdioServer();
    server.start();
}
//# sourceMappingURL=cli.js.map