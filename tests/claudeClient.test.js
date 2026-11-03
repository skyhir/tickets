// Unit tests for claudeClient (Claude Haiku 5.5 first, Gemini fallback).
// Same file in tickets / toll-processing / collections-script. Run: node tests/claudeClient.test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const clientFile = fs.existsSync(path.join(root, 'claudeClient.gs')) ? 'claudeClient.gs' : 'claudeClient.js';

function makeSandbox(props) {
  const calls = [];
  const queue = [];
  const ctx = {
    Logger: { log: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props || {})[k] || null }) },
    Utilities: { sleep: () => {} },
    UrlFetchApp: {
      fetch: (url, opts) => {
        calls.push({ url, opts, body: JSON.parse(opts.payload) });
        const next = queue.shift();
        if (!next) throw new Error('mock: no response queued');
        if (next.throws) throw new Error(next.throws);
        return {
          getResponseCode: () => next.status,
          getContentText: () => (typeof next.body === 'string' ? next.body : JSON.stringify(next.body)),
          getHeaders: () => ({ 'request-id': 'req_test' })
        };
      }
    }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(root, clientFile), 'utf8'), ctx, { filename: clientFile });
  ctx.__calls = calls;
  ctx.__queue = queue;
  return ctx;
}

const okBody = (text, extra) => Object.assign({
  model: 'claude-haiku-5-5',
  stop_reason: 'end_turn',
  content: [{ type: 'thinking', thinking: '' }, { type: 'text', text }],
  usage: { input_tokens: 1000, output_tokens: 200 }
}, extra || {});

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
}

// ---------------------------------------------------------------- request shape
test('buildClaudeRequest_: Haiku 5.5, no sampling params, PDF before text, schema + low effort', () => {
  const C = makeSandbox();
  const schema = { type: 'object', additionalProperties: false, required: ['a'], properties: { a: { type: 'string' } } };
  const body = C.buildClaudeRequest_({ prompt: 'hi', pdfBase64: 'QUJD\nREVG', jsonSchema: schema, maxTokens: 16000 });
  assert.strictEqual(body.model, 'claude-haiku-5-5');
  assert.strictEqual(body.max_tokens, 16000);
  for (const k of ['temperature', 'top_p', 'top_k', 'thinking']) assert.ok(!(k in body), k + ' must not be sent');
  const content = body.messages[0].content;
  assert.strictEqual(body.messages.length, 1);
  assert.strictEqual(body.messages[0].role, 'user');
  assert.strictEqual(content[0].type, 'document');
  assert.strictEqual(content[0].source.media_type, 'application/pdf');
  assert.strictEqual(content[0].source.data, 'QUJDREVG', 'base64 must have no newlines');
  assert.strictEqual(content[1].type, 'text');
  assert.strictEqual(body.output_config.effort, 'low');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(body.output_config.format)), { type: 'json_schema', schema });
});

test('buildClaudeRequest_: text-only, no format when no schema, default max_tokens >= 4096', () => {
  const C = makeSandbox();
  const body = C.buildClaudeRequest_({ prompt: 'csv please' });
  assert.strictEqual(body.messages[0].content.length, 1);
  assert.ok(!('format' in body.output_config));
  assert.ok(body.max_tokens >= 4096);
});

// ---------------------------------------------------------------- response parsing
test('parseClaudeResponse_: skips thinking blocks, parses JSON, prices tokens', () => {
  const C = makeSandbox();
  const r = C.parseClaudeResponse_(200, JSON.stringify(okBody('{"a":"x"}')), { expectJson: true });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.json.a, 'x');
  assert.strictEqual(r.text, '{"a":"x"}');
  assert.ok(Math.abs(r.costUsd - (1000 * 0.10 + 200 * 0.50) / 1e6) < 1e-12);
});

test('parseClaudeResponse_: fenced JSON still parses', () => {
  const C = makeSandbox();
  const r = C.parseClaudeResponse_(200, JSON.stringify(okBody('```json\n{"a":1}\n```')), { expectJson: true });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.json.a, 1);
});

test('parseClaudeResponse_: failure codes', () => {
  const C = makeSandbox();
  const cases = [
    [402, { type: 'error', error: { type: 'billing_error', message: 'credit' } }, 'http_402'],
    [429, { type: 'error', error: { type: 'rate_limit_error', message: 'slow down' } }, 'http_429'],
    [200, 'not json', 'unparseable_body'],
    [200, okBody('', { stop_reason: 'refusal', stop_details: { type: 'refusal', category: 'cyber', explanation: 'no' } }), 'refusal'],
    [200, okBody('{"a":', { stop_reason: 'max_tokens' }), 'max_tokens'],
    [200, okBody('x', { stop_reason: 'pause_turn' }), 'stop_pause_turn'],
    [200, okBody('   '), 'empty_text'],
    [200, okBody('{nope'), 'invalid_json']
  ];
  for (const [status, body, failure] of cases) {
    const r = C.parseClaudeResponse_(status, typeof body === 'string' ? body : JSON.stringify(body), { expectJson: true });
    assert.strictEqual(r.ok, false, failure);
    assert.strictEqual(r.failure, failure);
  }
  const refusal = C.parseClaudeResponse_(200, JSON.stringify(okBody('', { stop_reason: 'refusal', stop_details: { category: 'cyber' } })), {});
  assert.strictEqual(refusal.refusalCategory, 'cyber');
  const billing = C.parseClaudeResponse_(402, JSON.stringify({ error: { type: 'billing_error', message: 'credit' } }), {});
  assert.strictEqual(billing.errorType, 'billing_error');
});

test('parseClaudeResponse_: plain text (no expectJson) is returned as-is', () => {
  const C = makeSandbox();
  const r = C.parseClaudeResponse_(200, JSON.stringify(okBody('"a","b"\n"1","2"')), {});
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.json, undefined);
  assert.strictEqual(r.text, '"a","b"\n"1","2"');
});

// ---------------------------------------------------------------- HTTP call
test('callClaude_: headers + one retry on 529, then ok', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'sk-test' });
  C.__queue.push({ status: 529, body: { error: { type: 'overloaded_error', message: 'busy' } } }, { status: 200, body: okBody('hi') });
  const r = C.callClaude_({ prompt: 'p' });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.attempts, 2);
  assert.strictEqual(C.__calls.length, 2);
  const o = C.__calls[0].opts;
  assert.strictEqual(C.__calls[0].url, 'https://api.anthropic.com/v1/messages');
  assert.strictEqual(o.headers['x-api-key'], 'sk-test');
  assert.strictEqual(o.headers['anthropic-version'], '2023-06-01');
  assert.strictEqual(o.contentType, 'application/json');
  assert.strictEqual(o.muteHttpExceptions, true);
  assert.strictEqual(r.requestId, 'req_test');
});

test('callClaude_: no retry on 400; network throw becomes failure=network', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'sk-test' });
  C.__queue.push({ status: 400, body: { error: { type: 'invalid_request_error', message: 'bad' } } });
  const r = C.callClaude_({ prompt: 'p' });
  assert.strictEqual(r.failure, 'http_400');
  assert.strictEqual(C.__calls.length, 1);
  C.__queue.push({ throws: 'Address unavailable' }, { throws: 'Address unavailable' });
  const n = C.callClaude_({ prompt: 'p' });
  assert.strictEqual(n.failure, 'network');
  assert.ok(/Address unavailable/.test(n.errorMessage));
});

test('callClaude_: missing key never calls the network', () => {
  const C = makeSandbox({});
  const r = C.callClaude_({ prompt: 'p' });
  assert.strictEqual(r.failure, 'no_api_key');
  assert.strictEqual(C.__calls.length, 0);
});

test('claudeResultSummary_ carries the deciding state', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
  C.__queue.push({ status: 200, body: okBody('', { stop_reason: 'max_tokens' }) });
  const s = C.claudeResultSummary_(C.callClaude_({ prompt: 'p', pdfBase64: 'QUJD' }));
  for (const part of ['provider=anthropic', 'model=claude-haiku-5-5', 'failure=max_tokens', 'http=200', 'stop_reason=max_tokens', 'request_id=req_test', 'pdf_b64_chars=4', 'out_tokens=200']) {
    assert.ok(s.indexOf(part) !== -1, part + ' missing from: ' + s);
  }
});

// ---------------------------------------------------------------- fallback decision
function runFallback(C, over) {
  const logs = [];
  let geminiCalls = 0;
  const o = Object.assign({
    label: 'test',
    claude: { prompt: 'p', expectJson: true },
    geminiKey: () => 'g-key',
    gemini: (key) => { geminiCalls++; assert.strictEqual(key, 'g-key'); return { from: 'gemini' }; },
    geminiModel: 'gemini-2.5-flash',
    log: (level, msg) => logs.push(level + ' ' + msg)
  }, over || {});
  let out;
  let err;
  try { out = C.aiWithFallback_(o); } catch (e) { err = e; }
  return { out, err, logs, geminiCalls: () => geminiCalls };
}

test('aiWithFallback_: no Anthropic key -> existing Gemini path, no Claude call', () => {
  const C = makeSandbox({});
  const r = runFallback(C);
  assert.strictEqual(r.out.provider, 'google');
  assert.strictEqual(r.out.model, 'gemini-2.5-flash');
  assert.strictEqual(r.out.fallbackReason, 'no_anthropic_key');
  assert.strictEqual(r.out.value.from, 'gemini');
  assert.strictEqual(C.__calls.length, 0);
});

test('aiWithFallback_: Claude ok -> Gemini never called', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
  C.__queue.push({ status: 200, body: okBody('{"from":"claude"}') });
  let gemKeyAsked = false;
  const r = runFallback(C, { geminiKey: () => { gemKeyAsked = true; return 'g-key'; } });
  assert.strictEqual(r.out.provider, 'anthropic');
  assert.strictEqual(r.out.model, 'claude-haiku-5-5');
  assert.strictEqual(r.out.value.from, 'claude');
  assert.strictEqual(r.out.fallbackReason, null);
  assert.strictEqual(r.geminiCalls(), 0);
  assert.strictEqual(gemKeyAsked, false, 'Gemini key lookup (Secret Manager in collections) must stay lazy');
});

test('aiWithFallback_: refusal / max_tokens / 402 / bad JSON each fall back to Gemini', () => {
  const failures = [
    [{ status: 200, body: okBody('', { stop_reason: 'refusal', stop_details: { category: 'cyber' } }) }, 'claude_refusal'],
    [{ status: 200, body: okBody('{"a":', { stop_reason: 'max_tokens' }) }, 'claude_max_tokens'],
    [{ status: 402, body: { error: { type: 'billing_error', message: 'no credit' } } }, 'claude_http_402'],
    [{ status: 200, body: okBody('{oops') }, 'claude_invalid_json']
  ];
  for (const [resp, reason] of failures) {
    const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
    C.__queue.push(resp);
    const r = runFallback(C);
    assert.strictEqual(r.out.provider, 'google', reason);
    assert.strictEqual(r.out.fallbackReason, reason);
    assert.strictEqual(r.geminiCalls(), 1);
    assert.ok(r.logs.some((l) => l.indexOf('WARN') === 0 && l.indexOf('reason=' + reason) !== -1 && l.indexOf('provider=anthropic') !== -1), reason + ' fallback must be logged with state');
  }
});

test('aiWithFallback_: validator rejection falls back with failure=invalid_output', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
  C.__queue.push({ status: 200, body: okBody('{"wrong":true}') });
  const r = runFallback(C, { validate: (res) => (res.json.violations ? null : 'no violations array') });
  assert.strictEqual(r.out.provider, 'google');
  assert.strictEqual(r.out.fallbackReason, 'claude_invalid_output');
  assert.ok(r.logs.some((l) => l.indexOf('no violations array') !== -1));
});

test('aiWithFallback_: Claude fails and no Gemini key -> throws with the Claude state', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
  C.__queue.push({ status: 200, body: okBody('', { stop_reason: 'refusal' }) });
  const r = runFallback(C, { geminiKey: () => null });
  assert.ok(r.err);
  assert.ok(/failure=refusal/.test(r.err.message), r.err.message);
  assert.strictEqual(r.geminiCalls(), 0);
});

test('aiWithFallback_: neither key -> throws a clear message', () => {
  const C = makeSandbox({});
  const r = runFallback(C, { geminiKey: () => null });
  assert.ok(r.err && /Neither ANTHROPIC_API_KEY nor a Gemini API key/.test(r.err.message));
});

test('aiWithFallback_: Gemini error after a Claude failure propagates (call site keeps its message)', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: 'k' });
  C.__queue.push({ status: 500, body: { error: { type: 'api_error', message: 'x' } } }, { status: 500, body: { error: { type: 'api_error', message: 'x' } } });
  const r = runFallback(C, { gemini: () => { throw new Error('Gemini HTTP 402: prepay depleted'); } });
  assert.ok(r.err && /Gemini HTTP 402/.test(r.err.message));
  assert.ok(r.logs.some((l) => l.indexOf('ERROR') === 0 && l.indexOf('fallback_reason=claude_http_500') !== -1));
});

test('getAnthropicApiKey_: Script Property first, then getAnthropicKey() (collections Secret Manager)', () => {
  const C = makeSandbox({ ANTHROPIC_API_KEY: ' sk-prop ' });
  assert.strictEqual(C.getAnthropicApiKey_(), 'sk-prop');
  const D = makeSandbox({});
  assert.strictEqual(D.getAnthropicApiKey_(), null);
  D.getAnthropicKey = () => 'sk-gsm';
  assert.strictEqual(D.getAnthropicApiKey_(), 'sk-gsm');
});

// ---------------------------------------------------------------- compare helpers
test('claudeCompareDiff_: cents / case / whitespace tolerant, reports real differences', () => {
  const C = makeSandbox();
  const a = { v: [{ amt: 1.5, s: 'SFMTA ' }], x: null };
  const b = { v: [{ amt: 1.50, s: 'sfmta' }], x: '' };
  assert.deepStrictEqual(Array.from(C.claudeCompareDiff_(a, b)), []);
  const d = C.claudeCompareDiff_({ v: [{ amt: 1 }, { amt: 2 }] }, { v: [{ amt: 1.01 }] });
  assert.ok(d.some((x) => x.indexOf('v.length: 2 vs 1') === 0));
  assert.ok(d.some((x) => x.indexOf('v[0].amt') === 0));
  const md = C.claudeCompareMarkdown_('t', [{ name: 'a.pdf', claudeOk: true, geminiOk: true, diffs: [], costUsd: 0.001 }]);
  assert.ok(/1\/1 identical/.test(md));
});

console.log('claudeClient tests passed: ' + passed);
