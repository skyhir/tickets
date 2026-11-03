// Call-site tests: analyzeTicketWithGemini_ now asks Claude Haiku 5.5 first and
// falls back to the original Gemini call. Run: node tests/ticketExtraction.test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function sandbox(props) {
  const calls = [];
  const queue = [];
  const ctx = {
    console,
    Logger: { log: () => {} },
    Session: { getScriptTimeZone: () => 'America/Los_Angeles' },
    SpreadsheetApp: {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null }) },
    Utilities: { sleep: () => {}, base64Encode: (b) => Buffer.from(b).toString('base64'), formatDate: () => '' },
    DriveApp: {
      getFileById: () => ({
        getName: () => 'ticket.pdf',
        getBlob: () => ({ getContentType: () => 'application/pdf', getBytes: () => [37, 80, 68, 70] })
      })
    },
    UrlFetchApp: {
      fetch: (url, opts) => {
        calls.push({ url, body: JSON.parse(opts.payload) });
        const n = queue.shift();
        return { getResponseCode: () => n.status, getContentText: () => JSON.stringify(n.body), getHeaders: () => ({}) };
      }
    }
  };
  vm.createContext(ctx);
  for (const f of ['claudeClient.gs', 'TicketProcessing.gs']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), ctx, { filename: f });
  }
  ctx.logs = [];
  vm.runInContext('logToSheet_ = function (fn, msg, level) { logs.push(level + " " + msg); };', ctx);
  ctx.calls = calls;
  ctx.queue = queue;
  return ctx;
}

const violation = {
  isToll: false, licensePlate: 'FL 22DWCE', licenseState: null, violationDate: '02/14/2026', violationTime: '10:32 AM',
  violationId: 'SF-1', pinNumber: null, issuingAgency: 'SFMTA', agencyShortName: 'SFMTA', violationType: 'Expired Meter',
  violationLocation: '100 Market St', originalPenaltyAmount: '87.00', additionalPenaltyAmount: null, paymentWebsite: null, dueDate: null
};
const claudeOk = (obj) => ({ status: 200, body: { model: 'claude-haiku-5-5', stop_reason: 'end_turn', usage: { input_tokens: 3000, output_tokens: 400 }, content: [{ type: 'text', text: JSON.stringify(obj) }] } });
const geminiOk = (obj) => ({ status: 200, body: { candidates: [{ finishReason: 'STOP', content: { parts: [{ text: JSON.stringify(obj) }] } }] } });

// Claude answers → no Gemini call, post-processing unchanged.
{
  const C = sandbox({ ANTHROPIC_API_KEY: 'k', GEMINI_API_KEY: 'g' });
  C.queue.push(claudeOk({ isLateNotice: false, violations: [Object.assign({}, violation)] }));
  const r = C.analyzeTicketWithGemini_('file1');
  assert.strictEqual(r.success, true, r.message);
  assert.strictEqual(r.provider, 'anthropic');
  assert.strictEqual(r.modelUsed, 'claude-haiku-5-5');
  assert.strictEqual(C.calls.length, 1);
  assert.ok(C.calls[0].url.indexOf('api.anthropic.com') !== -1);
  const body = C.calls[0].body;
  assert.ok(!('temperature' in body));
  assert.strictEqual(body.messages[0].content[0].type, 'document');
  assert.strictEqual(body.output_config.format.type, 'json_schema');
  const v = r.data.violations[0];
  assert.strictEqual(v.originalPenaltyAmount, 87);
  assert.strictEqual(v.additionalPenaltyAmount, 0);
  assert.strictEqual(v.licensePlate, '22DWCE');
  assert.strictEqual(v.licenseState, 'FL');
  assert.ok(/Claude \(claude-haiku-5-5\)/.test(r.message));
}

// Claude truncated (max_tokens) → same request answered by Gemini.
{
  const C = sandbox({ ANTHROPIC_API_KEY: 'k', GEMINI_API_KEY: 'g' });
  C.queue.push({ status: 200, body: { stop_reason: 'max_tokens', content: [{ type: 'text', text: '{"isLate' }] } });
  C.queue.push(geminiOk({ isLateNotice: true, violations: [] }));
  const r = C.analyzeTicketWithGemini_('file2');
  assert.strictEqual(r.success, true, r.message);
  assert.strictEqual(r.provider, 'google');
  assert.strictEqual(r.modelUsed, 'gemini-2.5-flash');
  assert.strictEqual(r.fallbackReason, 'claude_max_tokens');
  assert.ok(C.calls[1].url.indexOf('gemini-2.5-flash:generateContent') !== -1);
  assert.strictEqual(C.calls[1].body.generationConfig.temperature, 0.1, 'Gemini request unchanged');
  assert.ok(C.logs.some((l) => /^WARN .*falling back to Gemini.*stop_reason=max_tokens/.test(l)));
}

// No Anthropic key → Gemini only, exactly as before.
{
  const C = sandbox({ GEMINI_API_KEY: 'g' });
  C.queue.push(geminiOk({ isLateNotice: false, violations: [Object.assign({}, violation)] }));
  const r = C.analyzeTicketWithGemini_('file3');
  assert.strictEqual(r.success, true);
  assert.strictEqual(r.provider, 'google');
  assert.strictEqual(C.calls.length, 1);
}

// Claude fails, Gemini key missing → clean failure with the Claude state.
{
  const C = sandbox({ ANTHROPIC_API_KEY: 'k' });
  C.queue.push({ status: 400, body: { error: { type: 'invalid_request_error', message: 'bad schema' } } });
  const r = C.analyzeTicketWithGemini_('file4');
  assert.strictEqual(r.success, false);
  assert.ok(/failure=http_400/.test(r.message) && /bad schema/.test(r.message), r.message);
}

// Schema keys == the prompt's documented keys, every object closed.
{
  const C = sandbox({});
  const s = vm.runInContext('TICKET_EXTRACTION_SCHEMA', C);
  assert.strictEqual(s.additionalProperties, false);
  const item = s.properties.violations.items;
  assert.strictEqual(item.additionalProperties, false);
  const prompt = C.buildTicketExtractionPrompt_();
  const promptKeys = Array.from(prompt.matchAll(/^\s+- "(\w+)":/gm)).map((m) => m[1]);
  assert.deepStrictEqual(Array.from(item.required).sort(), promptKeys.sort());
}

console.log('ticketExtraction tests passed');
