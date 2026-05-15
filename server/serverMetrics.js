import { promises as fsp } from "fs";
import os from "os";

let previousCpuSnapshot = null;
let previousNetworkSnapshot = null;

function takeCpuSnapshot() {
  const cpus = os.cpus();
  let idle = 0;
  let total = 0;
  cpus.forEach((cpu) => {
    idle += cpu.times.idle;
    total += Object.values(cpu.times).reduce((sum, value) => sum + value, 0);
  });
  return { idle, total };
}

function getCpuUsagePercent() {
  const current = takeCpuSnapshot();
  const previous = previousCpuSnapshot;
  previousCpuSnapshot = current;
  if (!previous) return 0;
  const idleDelta = current.idle - previous.idle;
  const totalDelta = current.total - previous.total;
  if (totalDelta <= 0) return 0;
  return Math.max(0, Math.min(100, ((totalDelta - idleDelta) / totalDelta) * 100));
}

async function getDiskStats(targetPath) {
  try {
    const stats = await fsp.statfs(targetPath);
    const blockSize = Number(stats.bsize) || 0;
    const total = Number(stats.blocks) * blockSize;
    const free = Number(stats.bfree) * blockSize;
    const available = Number(stats.bavail) * blockSize;
    const used = Math.max(0, total - free);
    const usedPercent = total > 0 ? (used / total) * 100 : 0;
    return { path: targetPath, total, used, free, available, usedPercent };
  } catch (_) {
    return { path: targetPath, total: 0, used: 0, free: 0, available: 0, usedPercent: 0 };
  }
}

function getNetworkInterfaceSummary() {
  return Object.entries(os.networkInterfaces())
    .map(([name, items]) => {
      const active = (items || []).filter((item) => !item.internal);
      if (!active.length) return null;
      return {
        name,
        addresses: active.map((item) => item.address).filter(Boolean),
        families: [...new Set(active.map((item) => item.family).filter(Boolean))]
      };
    })
    .filter(Boolean);
}

async function getNetworkStats() {
  const interfaces = getNetworkInterfaceSummary();
  try {
    const raw = await fsp.readFile("/proc/net/dev", "utf8");
    let rxBytes = 0;
    let txBytes = 0;
    raw.split(/\r?\n/).forEach((line) => {
      const parts = line.trim().split(/\s+/);
      if (parts.length < 17 || !parts[0].endsWith(":")) return;
      const name = parts[0].slice(0, -1);
      if (name === "lo") return;
      rxBytes += Number(parts[1]) || 0;
      txBytes += Number(parts[9]) || 0;
    });
    const now = Date.now();
    let rxPerSec = 0;
    let txPerSec = 0;
    if (previousNetworkSnapshot && now > previousNetworkSnapshot.at) {
      const seconds = (now - previousNetworkSnapshot.at) / 1000;
      rxPerSec = Math.max(0, (rxBytes - previousNetworkSnapshot.rxBytes) / seconds);
      txPerSec = Math.max(0, (txBytes - previousNetworkSnapshot.txBytes) / seconds);
    }
    previousNetworkSnapshot = { at: now, rxBytes, txBytes };
    return { interfaces, rxBytes, txBytes, rxPerSec, txPerSec, supported: true };
  } catch (_) {
    return { interfaces, rxBytes: 0, txBytes: 0, rxPerSec: 0, txPerSec: 0, supported: false };
  }
}

export async function buildServerMetricsPayload({
  nodeEnv,
  rootDir,
  mediaStoragePath,
  pool,
  getRealtimeConnectionCount
}) {
  const memoryTotal = os.totalmem();
  const memoryFree = os.freemem();
  const memoryUsed = Math.max(0, memoryTotal - memoryFree);
  const processMemory = process.memoryUsage();
  const disk = await getDiskStats(rootDir);
  const mediaDisk = mediaStoragePath === rootDir ? disk : await getDiskStats(mediaStoragePath);
  const network = await getNetworkStats();
  return {
    ok: true,
    at: Date.now(),
    server: {
      nodeEnv,
      platform: `${os.type()} ${os.release()}`,
      arch: os.arch(),
      hostname: os.hostname(),
      nodeVersion: process.version,
      uptimeSec: os.uptime(),
      processUptimeSec: process.uptime()
    },
    cpu: {
      usagePercent: getCpuUsagePercent(),
      cores: os.cpus().length,
      model: os.cpus()[0]?.model || ""
    },
    memory: {
      total: memoryTotal,
      used: memoryUsed,
      free: memoryFree,
      usedPercent: memoryTotal > 0 ? (memoryUsed / memoryTotal) * 100 : 0,
      processRss: processMemory.rss,
      processHeapUsed: processMemory.heapUsed,
      processHeapTotal: processMemory.heapTotal
    },
    disk,
    mediaDisk,
    network,
    process: {
      pid: process.pid,
      pool: {
        total: pool.totalCount,
        idle: pool.idleCount,
        waiting: pool.waitingCount
      }
    },
    realtime: {
      connections: getRealtimeConnectionCount()
    }
  };
}
