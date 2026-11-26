#!/usr/bin/env node
/**
 * End-to-end smoke test for the dsh-connectors MCP server.
 *
 *   node mcp/smoke-test.mjs
 *
 * Speaks the real protocol over real stdio: it launches server.mjs exactly the
 * way a host does, then asserts the parts that must work on any machine.
 * Anything that depends on locally installed applications is NOT asserted here
 * (a CI runner has no Obsidian and no MATLAB); connector_status is the only
 * application-facing call, and it must succeed by reporting what is missing.
 *
 * Exit code = number of failed assertions.
 */

import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SERVER = path.join(HERE, 'server.mjs');

let passed = 0;
let failed = 0;

function ok(name, detail) {
    console.log('PASS  ' + name.padEnd(28) + ' ' + (detail ?? ''));
    passed++;
}
function bad(name, detail) {
    console.log('FAIL  ' + name.padEnd(28) + ' ' + (detail ?? ''));
    failed++;
}

function startServer() {
    const child = spawn(process.execPath, [SERVER], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');

    const pending = new Map();
    let buffer = '';
    let nextId = 1;
    let stderr = '';

    child.stderr.on('data', (c) => { stderr += c; });
    child.stdout.on('data', (chunk) => {
        buffer += chunk;
        let index;
        while ((index = buffer.indexOf('\n')) >= 0) {
            const line = buffer.slice(0, index).trim();
            buffer = buffer.slice(index + 1);
            if (!line) continue;
            let message;
            try { message = JSON.parse(line); } catch { continue; }
            const waiter = pending.get(message.id);
            if (waiter) { pending.delete(message.id); waiter(message); }
        }
    });

    function request(method, params) {
        const id = nextId++;
        const payload = { jsonrpc: '2.0', id, method, ...(params === undefined ? {} : { params }) };
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => { pending.delete(id); reject(new Error('timeout waiting for ' + method)); }, 120000);
            pending.set(id, (message) => { clearTimeout(timer); resolve(message); });
            child.stdin.write(JSON.stringify(payload) + '\n');
        });
    }

    function notify(method, params) {
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, ...(params === undefined ? {} : { params }) }) + '\n');
    }

    return { child, request, notify, getStderr: () => stderr };
}

const client = startServer();

try {
    // 1. initialize must answer with a protocol version we both understand.
    const init = await client.request('initialize', {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'smoke-test', version: '0.0.0' },
    });
    if (init.result && typeof init.result.protocolVersion === 'string') {
        ok('initialize', 'protocol=' + init.result.protocolVersion + ' server=' + init.result.serverInfo.name);
    } else {
        bad('initialize', JSON.stringify(init).slice(0, 300));
    }
    client.notify('notifications/initialized');

    // 2. tools/list must advertise the catalog, and every entry needs a schema.
    const list = await client.request('tools/list', {});
    const tools = (list.result && list.result.tools) || [];
    if (tools.length >= 10) ok('tools/list', 'count=' + tools.length);
    else bad('tools/list', 'count=' + tools.length);

    const missingSchema = tools.filter((t) => !t.inputSchema || t.inputSchema.type !== 'object');
    if (tools.length > 0 && missingSchema.length === 0) ok('schemas', 'all tools declare an object schema');
    else bad('schemas', missingSchema.map((t) => t.name).join(','));

    const names = tools.map((t) => t.name);
    const expected = ['connector_status', 'obsidian_list', 'obsidian_read', 'obsidian_search',
        'obsidian_create', 'obsidian_append', 'matlab_run', 'office_read',
        'office_write_docx', 'pdf_convert'];
    const absent = expected.filter((n) => !names.includes(n));
    if (absent.length === 0) ok('catalog', expected.length + ' expected tools present');
    else bad('catalog', 'missing: ' + absent.join(','));

    // 3. An unknown tool must be a protocol error, not a silent success.
    const unknown = await client.request('tools/call', { name: 'no_such_tool', arguments: {} });
    if (unknown.error && unknown.error.code === -32602) ok('unknown-tool', 'code=' + unknown.error.code);
    else bad('unknown-tool', JSON.stringify(unknown).slice(0, 300));

    // 4. The one tool that must work everywhere: it reports what is missing.
    const status = await client.request('tools/call', { name: 'connector_status', arguments: {} });
    const content = status.result && status.result.content && status.result.content[0];
    if (!content || content.type !== 'text') {
        bad('connector_status', 'no text content');
    } else {
        const payload = JSON.parse(content.text);
        const rows = Array.isArray(payload.data) ? payload.data : [];
        if (payload.ok === true && rows.length > 0) {
            const resolved = rows.filter((r) => r.resolved).length;
            ok('connector_status', rows.length + ' connectors, ' + resolved + ' resolved');
        } else {
            bad('connector_status', content.text.slice(0, 300));
        }
    }

    // 5. Bad arguments must surface as isError, but still be a valid response.
    const badArgs = await client.request('tools/call', { name: 'obsidian_read', arguments: {} });
    if (badArgs.result && badArgs.result.isError === true) ok('bad-arguments', 'isError=true');
    else bad('bad-arguments', JSON.stringify(badArgs).slice(0, 300));
} catch (err) {
    bad('harness', err.message);
} finally {
    client.child.stdin.end();
    client.child.kill();
}

const stderrTail = client.getStderr().trim();
if (stderrTail) {
    console.log('\n--- server stderr ---');
    console.log(stderrTail.split('\n').slice(-6).join('\n'));
}

console.log('\nRESULT: ' + passed + ' passed, ' + failed + ' failed');
process.exit(failed);
