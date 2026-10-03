#!/usr/bin/env node
"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
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
const path = __importStar(require("node:path"));
const server_js_1 = require("./server.js");
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
    const server = new server_js_1.McpStdioServer();
    server.start();
}
