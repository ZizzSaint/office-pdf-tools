/** 子进程执行工具：超时、取消、进程树终止。 */
import { spawn } from "node:child_process";
import fs from "node:fs/promises";

function killTree(child) {
  if (!child || child.exitCode !== null || child.signalCode) return;
  try {
    if (process.platform === "win32") {
      spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore", windowsHide: true }).unref();
    } else {
      process.kill(-child.pid, "SIGKILL");
    }
  } catch {
    try { child.kill("SIGKILL"); } catch { /* ignore */ }
  }
}

/** 结束某个 PID（用于清理 Office COM 进程）。 */
export function killPid(pid) {
  if (!pid) return;
  try {
    if (process.platform === "win32") {
      spawn("taskkill", ["/PID", String(pid), "/T", "/F"], { stdio: "ignore", windowsHide: true }).unref();
    } else {
      process.kill(Number(pid), "SIGKILL");
    }
  } catch { /* ignore */ }
}

/**
 * @param {string} command
 * @param {string[]} args
 * @param {{timeoutMs?:number, signal?:AbortSignal, cwd?:string, env?:Record<string,string>, maxBuffer?:number, pidFile?:string}} [opts]
 * @returns {Promise<{code:number|null, stdout:string, stderr:string, timedOut:boolean, aborted:boolean, killedPid?:number}>}
 */
export function runProcess(command, args, opts = {}) {
  const {
    timeoutMs = 5 * 60 * 1000,
    signal,
    cwd,
    env,
    maxBuffer = 8 * 1024 * 1024,
    pidFile,
  } = opts;

  return new Promise((resolve) => {
    let settled = false;
    let stdout = "";
    let stderr = "";
    let timedOut = false;
    let aborted = false;

    const child = spawn(command, args, {
      cwd,
      env: { ...process.env, ...(env || {}) },
      windowsHide: true,
      detached: process.platform !== "win32",
      stdio: ["ignore", "pipe", "pipe"],
    });

    const timer = setTimeout(async () => {
      timedOut = true;
      killTree(child);
      // COM 场景下 Office 进程可能与 powershell 分离，需要单独清理。
      if (pidFile) {
        try {
          const pid = Number((await fs.readFile(pidFile, "utf8")).trim());
          if (pid) killPid(pid);
        } catch { /* ignore */ }
      }
    }, timeoutMs);

    const onAbort = () => {
      aborted = true;
      killTree(child);
      if (pidFile) fs.readFile(pidFile, "utf8").then((t) => killPid(Number(t.trim()))).catch(() => {});
    };
    if (signal) {
      if (signal.aborted) onAbort();
      else signal.addEventListener("abort", onAbort, { once: true });
    }

    child.stdout?.on("data", (chunk) => {
      if (stdout.length < maxBuffer) stdout += chunk.toString();
    });
    child.stderr?.on("data", (chunk) => {
      if (stderr.length < maxBuffer) stderr += chunk.toString();
    });

    const finish = async (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener?.("abort", onAbort);
      let killedPid;
      if (pidFile) {
        try {
          killedPid = Number((await fs.readFile(pidFile, "utf8")).trim()) || undefined;
        } catch { /* ignore */ }
      }
      resolve({ code, stdout, stderr, timedOut, aborted, killedPid });
    };

    child.on("error", (err) => {
      stderr += String(err?.message || err);
      finish(-1);
    });
    child.on("close", (code) => finish(code));
  });
}

export async function commandExists(command) {
  const probe = process.platform === "win32" ? "where" : "which";
  const { code } = await runProcess(probe, [command], { timeoutMs: 8000 });
  return code === 0;
}
