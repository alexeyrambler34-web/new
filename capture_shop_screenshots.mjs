import fs from "node:fs/promises";
import path from "node:path";

const [, , linksJson, outputDir, edgeExe] = process.argv;

if (!linksJson || !outputDir || !edgeExe) {
  console.error("Usage: node capture_shop_screenshots.mjs <links.json> <outputDir> <msedge.exe>");
  process.exit(1);
}

function sanitize(value) {
  return String(value || "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function jsonGet(url, attempts = 60) {
  let lastError;
  for (let i = 0; i < attempts; i += 1) {
    try {
      const response = await fetch(url);
      if (response.ok) return await response.json();
      lastError = new Error(`${response.status} ${response.statusText}`);
    } catch (error) {
      lastError = error;
    }
    await wait(500);
  }
  throw lastError || new Error(`Cannot GET ${url}`);
}

async function dismissOverlays(cdp) {
  await cdp.send("Input.dispatchKeyEvent", { type: "keyDown", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 }).catch(() => {});
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 }).catch(() => {});

  const expression = String.raw`(() => {
    const norm = (value) => String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
    const hit = (value) => /(закрыть|не сейчас|позже|нет спасибо|понятно|окей|продолжить без входа|без входа)/i.test(value || "");
    let changed = 0;

    for (const element of Array.from(document.querySelectorAll('button, [role="button"], a, [aria-label], [title]'))) {
      const label = [element.innerText, element.textContent, element.getAttribute("aria-label"), element.getAttribute("title")].map(norm).join(" ");
      if (hit(label)) {
        try {
          element.click();
          changed += 1;
        } catch {}
      }
    }

    const hasLoginText = (element) => {
      const text = norm(element.innerText || element.textContent);
      if (/войдите|станет дешевле|создадим аккаунт|номер телефона|получайте скидки|аккаунт/i.test(text)) return true;
      return Boolean(element.querySelector?.('input[type="tel"], input[placeholder*="телефон" i], input[name*="phone" i]'));
    };

    const isLargeOverlay = (element) => {
      const style = window.getComputedStyle(element);
      if (!["fixed", "sticky"].includes(style.position)) return false;
      const rect = element.getBoundingClientRect();
      const area = Math.max(0, rect.width) * Math.max(0, rect.height);
      const viewport = window.innerWidth * window.innerHeight;
      const z = Number.parseInt(style.zIndex, 10);
      return area > viewport * 0.18 || Number.isFinite(z) && z >= 100;
    };

    for (const element of Array.from(document.querySelectorAll('div, section, aside, dialog, form'))) {
      if (!document.body.contains(element)) continue;
      const role = norm(element.getAttribute("role"));
      const looksModal = role === "dialog" || role === "alertdialog" || element.matches?.('[data-auto*="login" i], [data-zone-name*="login" i], [class*="modal" i], [class*="popup" i]');
      if ((looksModal || isLargeOverlay(element)) && hasLoginText(element)) {
        let target = element;
        for (let i = 0; i < 4 && target.parentElement && target.parentElement !== document.body; i += 1) {
          const parent = target.parentElement;
          if (isLargeOverlay(parent) || norm(parent.getAttribute("role")) === "dialog") target = parent;
        }
        target.remove();
        changed += 1;
      }
    }

    for (const element of Array.from(document.querySelectorAll('div'))) {
      if (!document.body.contains(element)) continue;
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      const fullScreen = style.position === "fixed"
        && rect.width >= window.innerWidth * 0.9
        && rect.height >= window.innerHeight * 0.9
        && Number.parseFloat(style.opacity || "1") >= 0.2;
      if (fullScreen && !element.querySelector?.('img, article, [data-apiary-widget-name], [data-card-index]')) {
        const text = norm(element.innerText || element.textContent);
        if (!text || /войдите|станет дешевле|номер телефона|создадим аккаунт/i.test(text)) {
          element.remove();
          changed += 1;
        }
      }
    }

    document.documentElement.style.overflow = "auto";
    document.body.style.overflow = "auto";
    return changed;
  })()`;

  const result = await cdp.send("Runtime.evaluate", {
    expression,
    awaitPromise: true,
    returnByValue: true,
  }).catch(() => ({ result: { value: 0 } }));

  await wait(500);
  return result?.result?.value || 0;
}

class Cdp {
  constructor(wsUrl) {
    this.wsUrl = wsUrl;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.ws = new WebSocket(this.wsUrl);
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error("WebSocket timeout")), 30000);
      this.ws.addEventListener("open", () => {
        clearTimeout(timer);
        resolve();
      });
      this.ws.addEventListener("error", (event) => {
        clearTimeout(timer);
        reject(new Error(`WebSocket error: ${event.message || "unknown"}`));
      });
    });
    this.ws.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (message.id && this.pending.has(message.id)) {
        const { resolve, reject } = this.pending.get(message.id);
        this.pending.delete(message.id);
        if (message.error) reject(new Error(message.error.message || JSON.stringify(message.error)));
        else resolve(message.result);
      }
    });
  }

  async send(method, params = {}) {
    const id = this.nextId++;
    const payload = JSON.stringify({ id, method, params });
    const promise = new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      setTimeout(() => {
        if (this.pending.has(id)) {
          this.pending.delete(id);
          reject(new Error(`${method} timeout`));
        }
      }, 120000);
    });
    this.ws.send(payload);
    return promise;
  }

  close() {
    try {
      this.ws?.close();
    } catch {}
  }
}

async function captureOne(baseUrl, item, screenshotPath) {
  const response = await fetch(`${baseUrl}/json/new?${encodeURIComponent(item.url)}`, { method: "PUT" });
  if (!response.ok) {
    throw new Error(`Cannot open tab: ${response.status} ${response.statusText}`);
  }
  const create = await response.json();
  const cdp = new Cdp(create.webSocketDebuggerUrl);
    await cdp.connect();
  try {
    await cdp.send("Page.enable");
    await cdp.send("Runtime.enable");
    await cdp.send("Network.enable");
    await cdp.send("Network.setCacheDisabled", { cacheDisabled: true });
    await cdp.send("Emulation.setDeviceMetricsOverride", {
      width: 1440,
      height: 2200,
      deviceScaleFactor: 1,
      mobile: false,
    });
    await wait(9000);
    await dismissOverlays(cdp);
    for (let i = 0; i < 10; i += 1) {
      if (i === 3 || i === 7) await dismissOverlays(cdp);
      await cdp.send("Runtime.evaluate", { expression: "window.scrollBy(0, 1000)", awaitPromise: false });
      await wait(500);
    }
    await wait(2500);
    await dismissOverlays(cdp);
    await cdp.send("Runtime.evaluate", { expression: "window.scrollTo(0, 0)", awaitPromise: false });
    await wait(1500);
    await dismissOverlays(cdp);

    const metrics = await cdp.send("Page.getLayoutMetrics");
    const size = metrics.cssContentSize || metrics.contentSize;
    const width = Math.max(1440, Math.ceil(size.width || 1440));
    const height = Math.min(30000, Math.max(2200, Math.ceil(size.height || 2200)));

    const shot = await cdp.send("Page.captureScreenshot", {
      format: "jpeg",
      quality: 82,
      captureBeyondViewport: true,
      clip: { x: 0, y: 0, width, height, scale: 1 },
    });
    await fs.writeFile(screenshotPath, Buffer.from(shot.data, "base64"));
  } finally {
    cdp.close();
    await fetch(`${baseUrl}/json/close/${create.id}`, { method: "PUT" }).catch(() => {});
  }
}

const childProcess = await import("node:child_process");

const links = JSON.parse(await fs.readFile(linksJson, "utf8"));
await fs.mkdir(outputDir, { recursive: true });

const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
const runDir = path.join(outputDir, `screenshots_${timestamp}`);
const profileDir = path.join(outputDir, "edge-profile-for-monitoring");
await fs.mkdir(runDir, { recursive: true });
await fs.mkdir(profileDir, { recursive: true });

const port = 9222 + Math.floor(Math.random() * 300);
const baseUrl = `http://127.0.0.1:${port}`;
const edge = childProcess.spawn(edgeExe, [
  `--remote-debugging-port=${port}`,
  `--user-data-dir=${profileDir}`,
  "--remote-allow-origins=*",
  "--no-first-run",
  "--no-default-browser-check",
  "--disable-cache",
  "--disable-features=Translate",
  "about:blank",
], {
  detached: true,
  stdio: "ignore",
});
edge.unref();

await jsonGet(`${baseUrl}/json/version`);

const log = [];

for (let i = 0; i < links.length; i += 1) {
  const item = links[i];
  const baseName = `${String(i + 1).padStart(2, "0")}_row${item.row}_${sanitize(item.seller)}_${item.marketplace}`;
  const screenshotPath = path.join(runDir, `${baseName}.jpg`);
  const record = {
    index: i + 1,
    row: item.row,
    seller: item.seller,
    marketplace: item.marketplace,
    url: item.url,
    screenshot: screenshotPath,
    status: "started",
    startedAt: new Date().toISOString(),
  };

  try {
    console.log(`[${i + 1}/${links.length}] ${item.seller} / ${item.marketplace}`);
    await captureOne(baseUrl, item, screenshotPath);
    record.status = "ok";
  } catch (error) {
    record.status = "error";
    record.error = error?.message || String(error);
    console.log(`  ERROR: ${record.error}`);
  } finally {
    record.finishedAt = new Date().toISOString();
    log.push(record);
    await wait(1000);
  }
}

await fs.writeFile(path.join(runDir, "capture_log.json"), JSON.stringify(log, null, 2), "utf8");
await fs.writeFile(
  path.join(runDir, "capture_log.csv"),
  [
    "index,row,seller,marketplace,status,screenshot,url,error",
    ...log.map((item) =>
      [
        item.index,
        item.row,
        JSON.stringify(item.seller || ""),
        item.marketplace,
        item.status,
        JSON.stringify(item.screenshot || ""),
        JSON.stringify(item.url || ""),
        JSON.stringify(item.error || ""),
      ].join(","),
    ),
  ].join("\n"),
  "utf8",
);

await fetch(`${baseUrl}/json/close`).catch(() => {});
console.log(`Done: ${runDir}`);
