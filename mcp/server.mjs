#!/usr/bin/env node
/**
 * dsh-connectors MCP server - zero dependencies.
 *
 * Transport: stdio, newline-delimited JSON-RPC 2.0 (the MCP stdio frame format).
 * Nothing but protocol messages is ever written to stdout; diagnostics go to
 * stderr, because one stray line breaks the host's parser.
 *
 * Each tool call is delegated to ../bridge.ps1, which reuses the PowerShell
 * functions in this repository. Tool arguments travel on the child's stdin as
 * JSON: quoting them into an argv would corrupt any note text containing
 * quotes, newlines or non-ASCII characters.
 *
 * Environment overrides:
 *   DSH_CONNECTORS_POWERSHELL   interpreter to use (default: powershell.exe)
 *   DSH_CONNECTORS_TIMEOUT_MS   per-call timeout   (default: 600000)
 */

import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BRIDGE = path.join(HERE, 'bridge.ps1');

const SERVER_INFO = { name: 'dsh-connectors', version: '0.1.0' };
const PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05'];
const DEFAULT_PROTOCOL = '2025-06-18';

const POWERSHELL = process.env.DSH_CONNECTORS_POWERSHELL || 'powershell.exe';
const TIMEOUT_MS = Number(process.env.DSH_CONNECTORS_TIMEOUT_MS || 600000);

const log = (...args) => process.stderr.write('[dsh-connectors] ' + args.join(' ') + '\n');

// ---------------------------------------------------------------------------
// Tool catalog. Descriptions are written for a model, not for a human skimming
// a README: they say when to reach for the tool and what comes back.
// ---------------------------------------------------------------------------
const TOOLS = [
    {
        name: 'connector_status',
        description:
            'Report which local applications dsh-connectors can reach on this machine ' +
            '(Obsidian vault, MATLAB, COMSOL, STM32, Keil, VS Code, Notion, draw.io, and a ' +
            'Python with the Office libraries). Start here when a connector reports an error: ' +
            'it distinguishes "not installed" from "installed but not resolved".',
        inputSchema: { type: 'object', properties: {}, additionalProperties: false },
    },
    {
        name: 'obsidian_list',
        description:
            'List recent Markdown notes from the Obsidian vault, newest first. ' +
            'Returns vault-relative paths, byte sizes and modification times.',
        inputSchema: {
            type: 'object',
            properties: {
                newest: { type: 'integer', minimum: 1, maximum: 500, description: 'How many notes to return (default 20).' },
                folder: { type: 'string', description: 'Vault-relative folder to restrict the listing to.' },
            },
            additionalProperties: false,
        },
    },
    {
        name: 'obsidian_read',
        description: 'Read one note from the Obsidian vault by its vault-relative path.',
        inputSchema: {
            type: 'object',
            properties: { path: { type: 'string', description: "Vault-relative path, e.g. 'inbox/idea.md'." } },
            required: ['path'],
            additionalProperties: false,
        },
    },
    {
        name: 'obsidian_search',
        description:
            'Full-text search across the Obsidian vault. Returns matching vault-relative ' +
            'paths with the first matching line and its line number.',
        inputSchema: {
            type: 'object',
            properties: {
                pattern: { type: 'string', description: 'Regular expression (PowerShell/Select-String syntax).' },
                max: { type: 'integer', minimum: 1, maximum: 200, description: 'Maximum matching notes (default 30).' },
            },
            required: ['pattern'],
            additionalProperties: false,
        },
    },
    {
        name: 'obsidian_create',
        description:
            'Create a note in the Obsidian vault. Fails if the note already exists unless ' +
            'overwrite is true. Parent folders are created as needed.',
        inputSchema: {
            type: 'object',
            properties: {
                path: { type: 'string', description: 'Vault-relative path of the new note.' },
                content: { type: 'string', description: 'Initial content (Markdown).' },
                overwrite: { type: 'boolean', description: 'Replace an existing note (default false).' },
            },
            required: ['path'],
            additionalProperties: false,
        },
    },
    {
        name: 'obsidian_append',
        description: 'Append text to an existing note in the Obsidian vault.',
        inputSchema: {
            type: 'object',
            properties: {
                path: { type: 'string', description: 'Vault-relative path of an existing note.' },
                content: { type: 'string', description: 'Text to append.' },
                newline: { type: 'boolean', description: 'Insert a CRLF before the text (default false).' },
            },
            required: ['path', 'content'],
            additionalProperties: false,
        },
    },
    {
        name: 'matlab_run',
        description:
            'Run MATLAB headless (matlab -batch) and return its exit code, wall time and ' +
            'console output. Provide either inline code or the path of a .m script. ' +
            'MATLAB is slow to start; expect several seconds even for trivial code.',
        inputSchema: {
            type: 'object',
            properties: {
                code: { type: 'string', description: 'MATLAB statements, e.g. "disp(version)".' },
                file: { type: 'string', description: 'Absolute path to a .m script (run from its own folder).' },
            },
            additionalProperties: false,
        },
    },
    {
        name: 'office_read',
        description:
            'Read a Word (.docx), Excel (.xlsx) or PowerPoint (.pptx) file and return its ' +
            'text: paragraphs and table dimensions, sheets with a row preview, or the slide ' +
            'outline. The format is chosen from the file extension.',
        inputSchema: {
            type: 'object',
            properties: {
                path: { type: 'string', description: 'Absolute path to a .docx, .xlsx or .pptx file.' },
                rows: { type: 'integer', minimum: 1, maximum: 200, description: 'Excel rows to preview (default 10).' },
            },
            required: ['path'],
            additionalProperties: false,
        },
    },
    {
        name: 'office_write_docx',
        description: 'Create a Word document from a title and a list of paragraphs.',
        inputSchema: {
            type: 'object',
            properties: {
                path: { type: 'string', description: 'Absolute path of the .docx to create.' },
                title: { type: 'string', description: 'Heading text (level 1).' },
                lines: { type: 'array', items: { type: 'string' }, description: 'One paragraph per item.' },
            },
            required: ['path'],
            additionalProperties: false,
        },
    },
    {
        name: 'pdf_convert',
        description:
            'Convert a document to PDF with pandoc (requires a LaTeX engine such as xelatex ' +
            'on PATH). Fails loudly rather than silently producing a partial file.',
        inputSchema: {
            type: 'object',
            properties: {
                input: { type: 'string', description: 'Absolute path of the source document.' },
                output: { type: 'string', description: 'Absolute path of the PDF (defaults to input with .pdf).' },
            },
            required: ['input'],
            additionalProperties: false,
        },
    },
];

const TOOL_NAMES = new Set(TOOLS.map((t) => t.name));

// ---------------------------------------------------------------------------
// Bridge invocation
// ---------------------------------------------------------------------------
function callBridge(tool, args) {
    return new Promise((resolve) => {
        let child;
        try {
            child = spawn(
                POWERSHELL,
                ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', BRIDGE, '-Tool', tool],
                { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] }
            );
        } catch (err) {
            resolve({ ok: false, error: 'failed to start PowerShell: ' + (err && err.message ? err.message : String(err)) });
            return;
        }

        let stdout = '';
        let stderr = '';
        let settled = false;
        const finish = (value) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            resolve(value);
        };

        const timer = setTimeout(() => {
            try { child.kill(); } catch { /* already gone */ }
            finish({ ok: false, error: 'timed out after ' + TIMEOUT_MS + ' ms' });
        }, TIMEOUT_MS);

        child.stdout.setEncoding('utf8');
        child.stderr.setEncoding('utf8');
        child.stdout.on('data', (chunk) => { stdout += chunk; });
        child.stderr.on('data', (chunk) => { stderr += chunk; });

        child.on('error', (err) => finish({ ok: false, error: 'spawn error: ' + (err && err.message ? err.message : String(err)) }));

        child.on('close', (code) => {
            const text = stdout.trim();
            if (!text) {
                finish({
                    ok: false,
                    error: 'bridge produced no output',
                    exitCode: code,
                    stderr: stderr.trim().slice(0, 4000),
                });
                return;
            }
            try {
                finish(JSON.parse(text));
            } catch (err) {
                finish({
                    ok: false,
                    error: 'bridge output was not JSON: ' + (err && err.message ? err.message : String(err)),
                    exitCode: code,
                    stdout: text.slice(0, 4000),
                    stderr: stderr.trim().slice(0, 4000),
                });
            }
        });

        // Arguments go in on stdin, never through argv.
        child.stdin.on('error', () => { /* the child may exit before reading */ });
        child.stdin.end(JSON.stringify(args || {}));
    });
}

// ---------------------------------------------------------------------------
// JSON-RPC plumbing
// ---------------------------------------------------------------------------
function rpcError(code, message, data) {
    const err = new Error(message);
    err.rpc = { code, message, ...(data === undefined ? {} : { data }) };
    return err;
}

const handlers = {
    async initialize(params) {
        const asked = params && params.protocolVersion;
        const version = PROTOCOL_VERSIONS.includes(asked) ? asked : DEFAULT_PROTOCOL;
        log('initialize: client asked ' + (asked || '(none)') + ', answering ' + version);
        return {
            protocolVersion: version,
            capabilities: { tools: {} },
            serverInfo: SERVER_INFO,
        };
    },

    async 'tools/list'() {
        return { tools: TOOLS };
    },

    async 'tools/call'(params) {
        const name = params && params.name;
        const args = (params && params.arguments) || {};
        if (!name) throw rpcError(-32602, 'tools/call requires a tool name');
        if (!TOOL_NAMES.has(name)) throw rpcError(-32602, 'unknown tool: ' + name);

        const result = await callBridge(name, args);
        const text = JSON.stringify(result, null, 2);
        return { content: [{ type: 'text', text }], isError: result.ok === false };
    },

    async ping() {
        return {};
    },
};

function send(message) {
    process.stdout.write(JSON.stringify(message) + '\n');
}

async function handleLine(line) {
    const raw = line.trim();
    if (!raw) return;

    let message;
    try {
        message = JSON.parse(raw);
    } catch (err) {
        send({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'parse error: ' + err.message } });
        return;
    }

    const isNotification = message.id === undefined || message.id === null;
    const method = message.method;

    // Notifications (initialized, cancelled, ...) get no response by spec.
    if (isNotification && method) {
        if (method !== 'notifications/initialized') log('notification: ' + method);
        return;
    }

    const handler = handlers[method];
    if (!handler) {
        send({ jsonrpc: '2.0', id: message.id, error: { code: -32601, message: 'method not found: ' + method } });
        return;
    }

    try {
        const result = await handler(message.params || {});
        send({ jsonrpc: '2.0', id: message.id, result });
    } catch (err) {
        const rpc = err && err.rpc ? err.rpc : { code: -32603, message: err && err.message ? err.message : String(err) };
        log('error in ' + method + ': ' + rpc.message);
        send({ jsonrpc: '2.0', id: message.id, error: rpc });
    }
}

const rl = createInterface({ input: process.stdin, crlfDelay: Infinity, terminal: false });
let chain = Promise.resolve();
rl.on('line', (line) => {
    chain = chain.then(() => handleLine(line)).catch((err) => log('unhandled: ' + err));
});
rl.on('close', () => {
    chain.finally(() => process.exit(0));
});

process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));

log('ready (tools=' + TOOLS.length + ', bridge=' + BRIDGE + ', powershell=' + POWERSHELL + ')');
