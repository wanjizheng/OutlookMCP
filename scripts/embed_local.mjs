#!/usr/bin/env node
/**
 * embed_local.mjs — Local GGUF embedding via node-llama-cpp.
 *
 * Two modes:
 *
 *   1) stdin/stdout JSON-line server (default; used by sync ingestion):
 *      Input:  one JSON line per request: { "texts": ["str1", "str2", ...] }
 *      Output: one JSON line per response: { "vectors": [[...], ...] }
 *              or on error: { "error": "message" }
 *      Pass --one-shot to exit after handling one request.
 *
 *   2) TCP server (--server PORT, used by email-rag search query embedding):
 *      Listens on 127.0.0.1:PORT. Accepts one connection at a time,
 *      reads one JSON-line request, sends one JSON-line response, closes.
 *      Survives across Python invocations until idle for IDLE_TIMEOUT_MS
 *      (default 30 min). Designed to amortise the ~26s model load
 *      across many search queries.
 *
 * Model path resolution (in order):
 *   1. --model-path CLI argument
 *   2. EMBED_LOCAL_MODEL_PATH env var
 *   3. memorySearch.local.modelPath in openclaw.json (reads hf: prefix)
 *   4. Default: hf:ggml-org/embeddinggemma-300m-qat-q8_0-GGUF/embeddinggemma-300m-qat-Q8_0.gguf
 *
 * GPU: auto-detected. NODE_LLAMA_CPP_GPU=cuda | metal | vulkan | false
 *   Set to 'false' to force CPU.
 */

import { createInterface } from 'readline';
import { fileURLToPath, pathToFileURL } from 'url';
import { existsSync, readFileSync, mkdirSync, writeFileSync, appendFileSync } from 'fs';
import { join, dirname } from 'path';
import { homedir } from 'os';
import { execFileSync } from 'child_process';
import { createRequire } from 'module';
import { createServer } from 'net';

// ── Locate node-llama-cpp ──────────────────────────────────────────────────

function findNodeLlamaCpp() {
  const candidates = [];

  // 1. npm global root / openclaw / node_modules
  try {
    const globalRoot = execFileSync('npm', ['root', '-g'], { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
    candidates.push(join(globalRoot, 'openclaw', 'node_modules', 'node-llama-cpp'));
  } catch (_) {}

  // 2. Relative to this script (skill can bundle its own copy)
  const scriptDir = fileURLToPath(new URL('.', import.meta.url));
  candidates.push(join(scriptDir, 'node_modules', 'node-llama-cpp'));

  // 3. System node_modules
  candidates.push('/usr/lib/node_modules/node-llama-cpp');
  candidates.push('/usr/local/lib/node_modules/node-llama-cpp');

  for (const p of candidates) {
    if (existsSync(join(p, 'package.json'))) return p;
  }
  return null;
}

// ── Optional explicit bootstrap of node-llama-cpp ──────────────────────────

function autoInstallNodeLlamaCpp() {
  const scriptDir = fileURLToPath(new URL('.', import.meta.url));
  process.stderr.write(JSON.stringify({ status: 'auto-install', message: 'node-llama-cpp not found, installing...' }) + '\n');

  // Ensure package.json exists (npm install needs it)
  const pkgPath = join(scriptDir, 'package.json');
  if (!existsSync(pkgPath)) {
    mkdirSync(scriptDir, { recursive: true });
    writeFileSync(pkgPath, JSON.stringify({ private: true, dependencies: { 'node-llama-cpp': '^3' } }));
  }

  try {
    const lockPath = join(scriptDir, 'package-lock.json');
    const installArgs = existsSync(lockPath)
      ? ['ci', '--omit=dev']
      : ['install', '--save-exact', 'node-llama-cpp@3.17.1'];
    execFileSync('npm', installArgs, {
      cwd: scriptDir,
      stdio: ['pipe', 'pipe', 'pipe'],
      timeout: 600_000, // 10 min (compilation can be slow)
    });
    process.stderr.write(JSON.stringify({ status: 'auto-install', message: 'node-llama-cpp installed successfully' }) + '\n');
  } catch (err) {
    const stderr = err.stderr ? err.stderr.toString().slice(-500) : err.message;
    process.stderr.write(JSON.stringify({ error: `auto-install failed: ${stderr}` }) + '\n');
    return null;
  }

  return findNodeLlamaCpp();
}

let llamaPath = findNodeLlamaCpp();
if (!llamaPath && process.env.EMAIL_RAG_AUTO_INSTALL === '1') {
  llamaPath = autoInstallNodeLlamaCpp();
}
if (!llamaPath) {
  process.stderr.write(JSON.stringify({ error: 'node-llama-cpp not found. Run `npm install` in this scripts directory, or set EMAIL_RAG_AUTO_INSTALL=1 for an explicit one-time bootstrap.' }) + '\n');
  process.exit(1);
}

// Dynamic import with resolved path
const { getLlama, LlamaEmbeddingContext, resolveModelFile } = await import(pathToFileURL(join(llamaPath, 'dist', 'index.js')).href);

// ── Resolve model path ─────────────────────────────────────────────────────

function resolveModelPath() {
  // 1. CLI --model-path
  const mpIdx = process.argv.indexOf('--model-path');
  if (mpIdx !== -1 && process.argv[mpIdx + 1]) return process.argv[mpIdx + 1];

  // 2. Env var
  if (process.env.EMBED_LOCAL_MODEL_PATH) return process.env.EMBED_LOCAL_MODEL_PATH;

  // 3. openclaw.json memorySearch.local.modelPath
  const configPath = join(homedir(), '.openclaw', 'openclaw.json');
  if (existsSync(configPath)) {
    try {
      const cfg = JSON.parse(readFileSync(configPath, 'utf8'));
      const mp = cfg?.agents?.defaults?.memorySearch?.local?.modelPath;
      if (mp) return mp;
    } catch (_) {}
  }

  // 4. Default
  return 'hf:ggml-org/embeddinggemma-300m-qat-q8_0-GGUF/embeddinggemma-300m-qat-Q8_0.gguf';
}

const modelPath = resolveModelPath();
const oneShot = process.argv.includes('--one-shot');

// Parse --server PORT and --idle-timeout-ms MS (only meaningful in --server mode).
const serverPortIdx = process.argv.indexOf('--server');
const serverPort = serverPortIdx !== -1 ? parseInt(process.argv[serverPortIdx + 1], 10) : null;
const idleTimeoutIdx = process.argv.indexOf('--idle-timeout-ms');
const idleTimeoutMs = idleTimeoutIdx !== -1 ? parseInt(process.argv[idleTimeoutIdx + 1], 10) : 30 * 60 * 1000; // 30 min default
const serverPidFile = process.env.EMBED_SERVER_PIDFILE || join(homedir(), '.openclaw', 'email-rag-embed-server.pid');
if (serverPort !== null && (!Number.isInteger(serverPort) || serverPort < 1 || serverPort > 65535)) {
  process.stderr.write(JSON.stringify({ error: 'Invalid --server port; expected 1..65535' }) + '\n');
  process.exit(2);
}
if (!Number.isFinite(idleTimeoutMs) || idleTimeoutMs < 1000) {
  process.stderr.write(JSON.stringify({ error: 'Invalid --idle-timeout-ms; expected at least 1000' }) + '\n');
  process.exit(2);
}

// ── Load model ─────────────────────────────────────────────────────────────

let llama, model, context;
let modelReady = false;

async function ensureModelLoaded() {
  if (modelReady) return;
  await initLlama();
  modelReady = true;
}

async function initLlama() {
  const envGpu = process.env.NODE_LLAMA_CPP_GPU;

  // NODE_LLAMA_CPP_GPU values:
  //   "false" | "cpu"  → force CPU only
  //   "cuda" | "vulkan" | "metal" → try that GPU, fallback to CPU
  //   unset → auto-detect (cuda → vulkan → metal → cpu)

  // NOTE: progressLogs:false is critical. node-llama-cpp's spawnCommand child-processes
  // the cmake/build subprocess and, with progressLogs:true, also (a) pipes the child's
  // stdout directly to our process.stdout (tens of lines of "Not searching for unused
  // variables...", "-- The C compiler identification...", etc.) which break our
  // line-delimited JSON protocol, AND (b) pipes our process.stdin into the child's
  // stdin (consuming embedding requests while cmake is still initializing).
  // Disabling progress logs keeps our stdin/stdout clean.
  const noProgress = { progressLogs: false };

  if (envGpu === 'false' || envGpu === 'cpu') {
    llama = await getLlama({ gpu: false, ...noProgress });
    process.stderr.write(JSON.stringify({ status: 'gpu-mode', mode: 'cpu', source: 'env' }) + '\n');
    return;
  }

  if (envGpu) {
    // Try the requested GPU backend, fallback to CPU if it fails
    try {
      llama = await getLlama({ gpu: envGpu, ...noProgress });
      process.stderr.write(JSON.stringify({ status: 'gpu-mode', mode: envGpu, source: 'env' }) + '\n');
      return;
    } catch (err) {
      process.stderr.write(JSON.stringify({ status: 'gpu-fallback', requested: envGpu, error: err.message, fallback: 'cpu' }) + '\n');
      llama = await getLlama({ gpu: false, ...noProgress });
      process.stderr.write(JSON.stringify({ status: 'gpu-mode', mode: 'cpu', source: 'env-fallback' }) + '\n');
      return;
    }
  }

  // Auto-detect: try GPU first (cuda → vulkan → metal), fall back to CPU
  for (const gpu of ['cuda', 'vulkan', 'metal']) {
    try {
      llama = await getLlama({ gpu, ...noProgress });
      process.stderr.write(JSON.stringify({ status: 'gpu-mode', mode: gpu, source: 'auto' }) + '\n');
      return;
    } catch (_) {
      // This GPU backend not available, try next
    }
  }

  // All GPU backends failed — use CPU
  llama = await getLlama({ gpu: false, ...noProgress });
  process.stderr.write(JSON.stringify({ status: 'gpu-mode', mode: 'cpu', source: 'auto-fallback' }) + '\n');
}

try {
  await initLlama();
  process.stderr.write(JSON.stringify({ status: 'loading', model: modelPath }) + '\n');

  // resolveModelFile handles hf: URIs → downloads to ~/.node-llama-cpp/models/ if needed
  // Falls back to treating modelPath as a direct file path
  let resolvedPath = modelPath;
  if (modelPath.startsWith('hf:')) {
    resolvedPath = await resolveModelFile(modelPath, {
      directory: join(homedir(), '.node-llama-cpp', 'models'),
    });
  }

  model = await llama.loadModel({ modelPath: resolvedPath });
  context = await model.createEmbeddingContext();
  modelReady = true;
  process.stderr.write(JSON.stringify({ status: 'ready', model: modelPath }) + '\n');
} catch (err) {
  process.stderr.write(JSON.stringify({ error: `Model load failed: ${err.message}` }) + '\n');
  process.exit(1);
}

// ── Embedding handler (shared by stdin and TCP modes) ─────────────────────

async function handleRequest(texts) {
  if (!Array.isArray(texts)) {
    return { error: 'texts must be an array' };
  }
  try {
    const vectors = [];
    for (const text of texts) {
      const emb = await context.getEmbeddingFor(text);
      vectors.push(Array.from(emb.vector));
    }
    return { vectors };
  } catch (err) {
    return { error: err.message };
  }
}

function readJsonLine(buf) {
  // Parses one JSON object from a buffered string. Returns { obj, rest } or
  // { obj: null, rest: buf } if no complete line yet.
  const nl = buf.indexOf('\n');
  if (nl === -1) return { obj: null, rest: buf };
  const line = buf.slice(0, nl).trim();
  const rest = buf.slice(nl + 1);
  if (!line) return { obj: null, rest };
  try {
    return { obj: JSON.parse(line), rest };
  } catch (e) {
    return { obj: { error: `JSON parse error: ${e.message}` }, rest };
  }
}

// ── TCP server mode (--server PORT) ────────────────────────────────────────

async function runTcpServer() {
  await ensureModelLoaded();
  const serverLogPath = join(homedir(), '.openclaw', 'workspace', 'data', 'email_rag', 'embed-server.log');
  try { mkdirSync(dirname(serverLogPath), { recursive: true }); } catch (_) {}

  // Write PID file so admin tooling can find us for graceful shutdown.
  try {
    writeFileSync(serverPidFile, String(process.pid));
  } catch (_) {}
  const cleanupPid = () => {
    try { writeFileSync(serverPidFile, ''); } catch (_) {}
  };
  process.on('exit', cleanupPid);
  process.on('SIGTERM', () => { cleanupPid(); process.exit(0); });
  process.on('SIGINT', () => { cleanupPid(); process.exit(0); });

  const server = createServer((socket) => {
    let buf = '';
    socket.on('data', (chunk) => {
      buf += chunk.toString('utf8');
      let nl;
      while ((nl = buf.indexOf('\n')) !== -1) {
        const line = buf.slice(0, nl).trim();
        buf = buf.slice(nl + 1);
        if (!line) continue;
        let req;
        try {
          req = JSON.parse(line);
        } catch (e) {
          socket.write(JSON.stringify({ error: `JSON parse error: ${e.message}` }) + '\n');
          continue;
        }
        lastActivity = Date.now();
        // Admin shutdown command: server listens for {"cmd":"shutdown"}.
        if (req && req.cmd === 'shutdown') {
          socket.write(JSON.stringify({ ok: true, msg: 'shutting down' }) + '\n');
          setTimeout(() => {
            cleanupPid();
            server.close(() => process.exit(0));
          }, 50);
          return;
        }
        if (req && req.cmd === 'ping') {
          socket.write(JSON.stringify({ ok: true, service: 'email-rag-embed' }) + '\n');
          socket.end();
          return;
        }
        const texts = req && req.texts;
        handleRequest(texts).then((resp) => {
          try { socket.write(JSON.stringify(resp) + '\n'); } catch (_) {}
          try { socket.end(); } catch (_) {}
        }).catch((err) => {
          try { socket.write(JSON.stringify({ error: String(err && err.message || err) }) + '\n'); } catch (_) {}
          try { socket.end(); } catch (_) {}
        });
      }
    });
    socket.on('error', () => { try { socket.destroy(); } catch (_) {} });
  });

  let lastActivity = Date.now();
  const idleCheck = setInterval(() => {
    if (Date.now() - lastActivity > idleTimeoutMs) {
      try { appendFileSync(serverLogPath, `[${new Date().toISOString()}] idle timeout ${idleTimeoutMs}ms reached; exiting\n`); } catch (_) {}
      clearInterval(idleCheck);
      try { server.close(); } catch (_) {}
      try { context.dispose(); } catch (_) {}
      try { model.dispose(); } catch (_) {}
      process.exit(0);
    }
  }, 30_000); // check every 30s

  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(serverPort, '127.0.0.1', () => {
      try { appendFileSync(serverLogPath, `[${new Date().toISOString()}] listening on 127.0.0.1:${serverPort} (idle=${idleTimeoutMs}ms)\n`); } catch (_) {}
      resolve();
    });
  });
}

// ── stdin/stdout server mode (default; used by sync ingestion) ────────────

async function runStdinServer() {
  await ensureModelLoaded();
  const rl = createInterface({ input: process.stdin, crlfDelay: Infinity });
  for await (const line of rl) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    let req;
    try {
      req = JSON.parse(trimmed);
    } catch (e) {
      process.stdout.write(JSON.stringify({ error: `JSON parse error: ${e.message}` }) + '\n');
      if (oneShot) break;
      continue;
    }
    const resp = await handleRequest(req.texts);
    process.stdout.write(JSON.stringify(resp) + '\n');
    if (oneShot) break;
  }
  try { rl.close(); } catch (_) {}
}

(async () => {
  try {
    if (serverPort !== null) {
      await runTcpServer();
    } else {
      await runStdinServer();
    }
  } catch (err) {
    process.stderr.write(JSON.stringify({ error: `embed_local.mjs fatal: ${err.message}` }) + '\n');
    process.exit(1);
  }
})();
