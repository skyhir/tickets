/**
 * claudeClient — minimal Claude Messages API client for Apps Script.
 *
 * Gemini → Claude Haiku 5.5 migration (Nov 2026). Google retires
 * gemini-2.5-flash on Dec 1 2026 and the Gemini prepay keeps running dry
 * (HTTP 402), so every AI call in this project now goes to Claude first and
 * falls back to the existing Gemini code for that one request when:
 *   - the Claude call fails (network error, non-200, refusal, max_tokens,
 *     empty text, unparseable/invalid output), AND a Gemini key is configured.
 * If ANTHROPIC_API_KEY is not configured at all, the Gemini path runs unchanged.
 *
 * Apps Script has no Anthropic SDK, so this is raw UrlFetchApp against
 * https://api.anthropic.com/v1/messages. Request-shape rules for Haiku 5.5:
 *   - no temperature/top_p/top_k (non-default values return 400),
 *   - no assistant prefill (400), no thinking budget_tokens (400),
 *   - thinking stays adaptive (on by default); extraction runs at effort "low",
 *   - thinking shares the max_tokens budget, so max_tokens is sized generously
 *     and stop_reason "max_tokens" is treated as a failure,
 *   - JSON comes back through structured outputs (output_config.format),
 *   - PDFs go in as a base64 `document` block placed BEFORE the text block.
 *
 * This file is identical in tickets / toll-processing / collections-script.
 * Keep it that way: fix it in all three.
 */

var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
var CLAUDE_API_VERSION = '2023-06-01';
var CLAUDE_MODEL = 'claude-haiku-5-5';
// Haiku 5.5 list price for prompts <= 100K tokens, USD per token. Cost logging only.
var CLAUDE_PRICE_IN_PER_TOKEN = 0.10 / 1000000;
var CLAUDE_PRICE_OUT_PER_TOKEN = 0.50 / 1000000;
// 408 timeout, 429 rate limit, 5xx server, 529 overloaded: worth one retry.
var CLAUDE_RETRYABLE_STATUS = { 408: true, 429: true, 500: true, 502: true, 503: true, 504: true, 529: true };
// UrlFetchApp has no per-request deadline option, so the 6-minute execution
// limit is protected by (a) max_tokens sizing at each call site and (b) only
// retrying when the failed attempt was quick. A slow failure goes straight to
// the Gemini fallback instead of spending the remaining budget on a retry.
var CLAUDE_RETRY_ONLY_IF_ATTEMPT_UNDER_MS = 20000;

/**
 * ANTHROPIC_API_KEY from Script Properties. collections-script also keeps it in
 * Secret Manager (getAnthropicKey() in collectionsEmails.js) — used when present.
 * @return {string|null}
 */
function getAnthropicApiKey_() {
  var key = null;
  try {
    key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  } catch (e) {
    key = null;
  }
  if (key && String(key).trim()) return String(key).trim();
  if (typeof getAnthropicKey === 'function') {
    try {
      var fromSecretManager = getAnthropicKey();
      if (fromSecretManager && String(fromSecretManager).trim()) return String(fromSecretManager).trim();
    } catch (e2) {
      // fall through: treated as "no Claude key", Gemini path runs
    }
  }
  return null;
}

/**
 * Builds the Messages API request body. Pure.
 * @param {{prompt:string, system?:string, pdfBase64?:string, jsonSchema?:Object,
 *          maxTokens?:number, effort?:string}} opts
 * @return {Object}
 */
function buildClaudeRequest_(opts) {
  var content = [];
  if (opts.pdfBase64) {
    content.push({
      type: 'document',
      source: {
        type: 'base64',
        media_type: 'application/pdf',
        // Utilities.base64Encode never wraps, but strip whitespace defensively.
        data: String(opts.pdfBase64).replace(/\s+/g, '')
      }
    });
  }
  content.push({ type: 'text', text: String(opts.prompt || '') });

  var body = {
    model: CLAUDE_MODEL,
    max_tokens: opts.maxTokens || 8192,
    messages: [{ role: 'user', content: content }],
    output_config: { effort: opts.effort || 'low' }
  };
  if (opts.system) body.system = opts.system;
  if (opts.jsonSchema) {
    body.output_config.format = { type: 'json_schema', schema: opts.jsonSchema };
  }
  return body;
}

/**
 * Turns an HTTP status + body into a result. Pure. Never throws.
 * `failure` is a short machine code (null when ok):
 *   http_<status> | unparseable_body | refusal | max_tokens | stop_<reason> |
 *   empty_text | invalid_json
 * @param {number} status
 * @param {string} bodyText
 * @param {{expectJson?:boolean}} opts
 * @return {Object}
 */
function parseClaudeResponse_(status, bodyText, opts) {
  opts = opts || {};
  var res = {
    ok: false,
    provider: 'anthropic',
    model: CLAUDE_MODEL,
    httpStatus: status,
    failure: null,
    stopReason: null,
    refusalCategory: null,
    errorType: null,
    errorMessage: null,
    text: null,
    json: undefined,
    usage: null,
    costUsd: null
  };

  var data;
  try {
    data = JSON.parse(bodyText);
  } catch (e) {
    data = null;
  }

  if (status !== 200) {
    res.failure = 'http_' + status;
    if (data && data.error) {
      res.errorType = data.error.type || null;
      res.errorMessage = data.error.message || null;
    } else {
      res.errorMessage = String(bodyText || '').substring(0, 300);
    }
    return res;
  }
  if (!data || typeof data !== 'object') {
    res.failure = 'unparseable_body';
    res.errorMessage = String(bodyText || '').substring(0, 300);
    return res;
  }

  if (data.model) res.model = data.model;
  if (data.usage) {
    res.usage = {
      input_tokens: data.usage.input_tokens || 0,
      output_tokens: data.usage.output_tokens || 0
    };
    res.costUsd = res.usage.input_tokens * CLAUDE_PRICE_IN_PER_TOKEN +
      res.usage.output_tokens * CLAUDE_PRICE_OUT_PER_TOKEN;
  }

  // Always check stop_reason before reading content.
  res.stopReason = data.stop_reason || null;
  if (res.stopReason === 'refusal') {
    res.failure = 'refusal';
    var sd = data.stop_details || {};
    res.refusalCategory = sd.category || null;
    res.errorMessage = sd.explanation || null;
    return res;
  }
  if (res.stopReason === 'max_tokens') {
    res.failure = 'max_tokens';
    return res;
  }
  if (res.stopReason !== 'end_turn' && res.stopReason !== 'stop_sequence') {
    res.failure = 'stop_' + res.stopReason;
    return res;
  }

  // Thinking blocks can come first: keep only text blocks.
  var text = '';
  var blocks = data.content || [];
  for (var i = 0; i < blocks.length; i++) {
    if (blocks[i] && blocks[i].type === 'text' && blocks[i].text) text += blocks[i].text;
  }
  res.text = text;
  if (!text.trim()) {
    res.failure = 'empty_text';
    return res;
  }

  if (opts.expectJson) {
    var cleaned = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
    try {
      res.json = JSON.parse(cleaned);
    } catch (e3) {
      res.failure = 'invalid_json';
      res.errorMessage = String(e3 && e3.message || e3) + ' | text: ' + cleaned.substring(0, 300);
      return res;
    }
  }

  res.ok = true;
  return res;
}

/**
 * Calls Claude once (plus at most one quick retry on 408/429/5xx/529/network).
 * Never throws; inspect `.ok` / `.failure`.
 * @param {{prompt:string, system?:string, pdfBase64?:string, jsonSchema?:Object,
 *          maxTokens?:number, effort?:string, apiKey?:string, maxAttempts?:number}} opts
 * @return {Object} parseClaudeResponse_ result + latencyMs, attempts, requestId,
 *                  requestBytes, pdfBase64Chars
 */
function callClaude_(opts) {
  var apiKey = opts.apiKey || getAnthropicApiKey_();
  var payload = JSON.stringify(buildClaudeRequest_(opts));
  var maxAttempts = opts.maxAttempts || 2;
  var started = Date.now();
  var res = null;

  if (!apiKey) {
    res = parseClaudeResponse_(0, '', opts);
    res.failure = 'no_api_key';
    res.errorMessage = 'ANTHROPIC_API_KEY is not configured';
    res.attempts = 0;
    res.latencyMs = 0;
    res.requestBytes = payload.length;
    res.pdfBase64Chars = opts.pdfBase64 ? String(opts.pdfBase64).length : 0;
    return res;
  }

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var t0 = Date.now();
    var requestId = null;
    try {
      var resp = UrlFetchApp.fetch(CLAUDE_API_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'x-api-key': apiKey,
          'anthropic-version': CLAUDE_API_VERSION
        },
        payload: payload,
        muteHttpExceptions: true
      });
      try {
        var headers = resp.getHeaders ? resp.getHeaders() : {};
        requestId = headers['request-id'] || headers['Request-Id'] || null;
      } catch (eh) {
        requestId = null;
      }
      res = parseClaudeResponse_(resp.getResponseCode(), resp.getContentText(), opts);
    } catch (e) {
      // UrlFetchApp throws on DNS/connection failures and timeouts even with muteHttpExceptions.
      res = parseClaudeResponse_(0, '', opts);
      res.failure = 'network';
      res.errorMessage = String(e && e.message || e);
    }
    var attemptMs = Date.now() - t0;
    res.requestId = requestId;
    res.attempts = attempt;
    res.latencyMs = Date.now() - started;
    res.requestBytes = payload.length;
    res.pdfBase64Chars = opts.pdfBase64 ? String(opts.pdfBase64).length : 0;

    var retryable = !res.ok &&
      (res.failure === 'network' || CLAUDE_RETRYABLE_STATUS[res.httpStatus] === true);
    if (!retryable || attempt >= maxAttempts || attemptMs >= CLAUDE_RETRY_ONLY_IF_ATTEMPT_UNDER_MS) break;
    Utilities.sleep(1500 * attempt);
  }
  return res;
}

/**
 * One-line, root-cause-ready description of a Claude result for logs.
 * @param {Object} res callClaude_ result
 * @return {string}
 */
function claudeResultSummary_(res) {
  if (!res) return 'provider=anthropic result=none';
  var u = res.usage || {};
  return [
    'provider=anthropic',
    'model=' + (res.model || CLAUDE_MODEL),
    'ok=' + !!res.ok,
    'failure=' + (res.failure || '-'),
    'http=' + (res.httpStatus == null ? '-' : res.httpStatus),
    'stop_reason=' + (res.stopReason || '-'),
    'refusal_category=' + (res.refusalCategory || '-'),
    'error_type=' + (res.errorType || '-'),
    'error=' + (res.errorMessage ? String(res.errorMessage).substring(0, 300) : '-'),
    'request_id=' + (res.requestId || '-'),
    'attempts=' + (res.attempts == null ? '-' : res.attempts),
    'latency_ms=' + (res.latencyMs == null ? '-' : res.latencyMs),
    'request_bytes=' + (res.requestBytes == null ? '-' : res.requestBytes),
    'pdf_b64_chars=' + (res.pdfBase64Chars == null ? '-' : res.pdfBase64Chars),
    'in_tokens=' + (u.input_tokens == null ? '-' : u.input_tokens),
    'out_tokens=' + (u.output_tokens == null ? '-' : u.output_tokens),
    'cost_usd=' + (res.costUsd == null ? '-' : res.costUsd.toFixed(5))
  ].join(' ');
}

/**
 * Claude first, Gemini fallback — the one place the fallback decision lives.
 *
 * @param {Object} o
 * @param {string} o.label             call-site name for logs
 * @param {Object} o.claude            callClaude_ options (prompt, pdfBase64, jsonSchema, maxTokens...)
 * @param {function(Object):?string} [o.validate]  returns an error string when a 200 Claude
 *                                     answer is still unusable (wrong shape); null when fine
 * @param {function():?string} o.geminiKey  lazily returns the Gemini key (null when not configured)
 * @param {function(string):*} o.gemini     runs the EXISTING Gemini path with that key; returns
 *                                     the same value shape as Claude's (parsed JSON / text) or throws
 * @param {string} o.geminiModel       for logs / model_used
 * @param {function(string,string)} [o.log]  (level, message) logger
 * @return {{provider:string, model:string, value:*, fallbackReason:?string, claude:?Object}}
 * @throws Error when no provider produced a usable answer (message carries the deciding state)
 */
function aiWithFallback_(o) {
  var log = o.log || function (level, msg) { Logger.log(level + ' ' + msg); };
  var claudeKey = getAnthropicApiKey_();
  var claudeRes = null;
  var fallbackReason;

  if (!claudeKey) {
    fallbackReason = 'no_anthropic_key';
    log('INFO', o.label + ': ANTHROPIC_API_KEY not configured, using Gemini (' + o.geminiModel + ').');
  } else {
    var claudeOpts = {};
    for (var k in o.claude) if (Object.prototype.hasOwnProperty.call(o.claude, k)) claudeOpts[k] = o.claude[k];
    claudeOpts.apiKey = claudeKey;
    claudeRes = callClaude_(claudeOpts);
    if (claudeRes.ok && o.validate) {
      var invalid = null;
      try {
        invalid = o.validate(claudeRes);
      } catch (ev) {
        invalid = 'validator threw: ' + String(ev && ev.message || ev);
      }
      if (invalid) {
        claudeRes.ok = false;
        claudeRes.failure = 'invalid_output';
        claudeRes.errorMessage = invalid;
      }
    }
    if (claudeRes.ok) {
      log('INFO', o.label + ': Claude OK. ' + claudeResultSummary_(claudeRes));
      return {
        provider: 'anthropic',
        model: claudeRes.model || CLAUDE_MODEL,
        value: claudeRes.json !== undefined ? claudeRes.json : claudeRes.text,
        fallbackReason: null,
        claude: claudeRes
      };
    }
    fallbackReason = 'claude_' + claudeRes.failure;
  }

  var geminiKey = null;
  try {
    geminiKey = o.geminiKey();
  } catch (eg) {
    geminiKey = null;
  }
  if (!geminiKey) {
    var why = claudeRes
      ? 'Claude failed and no Gemini key is configured, so there is no fallback. ' + claudeResultSummary_(claudeRes)
      : 'Neither ANTHROPIC_API_KEY nor a Gemini API key is configured.';
    log('ERROR', o.label + ': ' + why);
    throw new Error(why);
  }

  if (claudeRes) {
    log('WARN', o.label + ': falling back to Gemini (' + o.geminiModel + '), reason=' + fallbackReason +
      '. ' + claudeResultSummary_(claudeRes));
  }
  var g0 = Date.now();
  var value;
  try {
    value = o.gemini(geminiKey);
  } catch (e) {
    log('ERROR', o.label + ': Gemini (' + o.geminiModel + ') also failed after fallback_reason=' + fallbackReason +
      ' in ' + (Date.now() - g0) + 'ms: ' + String(e && e.message || e));
    throw e;
  }
  log('INFO', o.label + ': answered by provider=google model=' + o.geminiModel + ' fallback_reason=' + fallbackReason +
    ' gemini_latency_ms=' + (Date.now() - g0) + '.');
  return { provider: 'google', model: o.geminiModel, value: value, fallbackReason: fallbackReason, claude: claudeRes };
}

// --------------------------------------------------------------------------
// Side-by-side comparison helpers (used by the one-off compare function in
// each project; read-only apart from writing the report file to Drive).
// --------------------------------------------------------------------------

/**
 * Field-level differences between two JSON values. Numbers compare to the cent;
 * strings compare trimmed and case-insensitive. Pure.
 * @return {Array<string>} e.g. ['violations[0].dueDate: "04/07/2025" vs "2025-04-07"']
 */
function claudeCompareDiff_(a, b, path) {
  path = path || '';
  var out = [];
  var norm = function (v) {
    if (v === undefined || v === '') return null;
    if (typeof v === 'string') return v.trim().toLowerCase();
    if (typeof v === 'number') return Math.round(v * 100) / 100;
    return v;
  };
  if (Array.isArray(a) || Array.isArray(b)) {
    var aa = Array.isArray(a) ? a : [];
    var bb = Array.isArray(b) ? b : [];
    if (aa.length !== bb.length) out.push(path + '.length: ' + aa.length + ' vs ' + bb.length);
    for (var i = 0; i < Math.min(aa.length, bb.length); i++) {
      out = out.concat(claudeCompareDiff_(aa[i], bb[i], path + '[' + i + ']'));
    }
    return out;
  }
  if (a && b && typeof a === 'object' && typeof b === 'object') {
    var keys = {};
    var k;
    for (k in a) keys[k] = true;
    for (k in b) keys[k] = true;
    for (k in keys) out = out.concat(claudeCompareDiff_(a[k], b[k], path ? path + '.' + k : k));
    return out;
  }
  if (norm(a) !== norm(b)) out.push((path || '(value)') + ': ' + JSON.stringify(a) + ' vs ' + JSON.stringify(b));
  return out;
}

/**
 * Renders compare rows as Markdown. Pure.
 * @param {string} title
 * @param {Array<Object>} rows {name, claudeMs, geminiMs, claudeOk, geminiOk, inTok, outTok, costUsd, diffs[], note}
 * @return {string}
 */
function claudeCompareMarkdown_(title, rows) {
  var lines = ['# ' + title, '', 'Model: Claude `' + CLAUDE_MODEL + '` vs existing Gemini path. Generated ' + new Date().toISOString() + '.', ''];
  var totalCost = 0;
  var same = 0;
  lines.push('| # | input | Claude ms | Gemini ms | Claude in/out tok | Claude $ | fields that differ |');
  lines.push('|---|---|---|---|---|---|---|');
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    totalCost += r.costUsd || 0;
    if (r.claudeOk && r.geminiOk && (!r.diffs || !r.diffs.length)) same++;
    var diffText = !r.claudeOk ? 'CLAUDE FAILED: ' + (r.note || '') :
      !r.geminiOk ? 'GEMINI FAILED: ' + (r.note || '') :
      (r.diffs && r.diffs.length ? r.diffs.slice(0, 12).join('<br>') + (r.diffs.length > 12 ? '<br>…+' + (r.diffs.length - 12) : '') : 'identical');
    lines.push('| ' + (i + 1) + ' | ' + String(r.name).replace(/\|/g, '/') + ' | ' + (r.claudeMs == null ? '-' : r.claudeMs) +
      ' | ' + (r.geminiMs == null ? '-' : r.geminiMs) + ' | ' + (r.inTok == null ? '-' : r.inTok) + '/' + (r.outTok == null ? '-' : r.outTok) +
      ' | ' + (r.costUsd == null ? '-' : r.costUsd.toFixed(4)) + ' | ' + String(diffText).replace(/\|/g, '/').replace(/\n/g, ' ') + ' |');
  }
  lines.push('');
  lines.push('**Headline:** ' + same + '/' + rows.length + ' identical; Claude spend $' + totalCost.toFixed(4) + '.');
  return lines.join('\n');
}

/**
 * Writes the compare report to a Drive file (My Drive root) and returns its URL.
 * @return {string}
 */
function claudeCompareWriteReport_(fileName, markdown) {
  var file = DriveApp.createFile(fileName, markdown, 'text/markdown');
  Logger.log('Compare report: ' + file.getUrl());
  return file.getUrl();
}
