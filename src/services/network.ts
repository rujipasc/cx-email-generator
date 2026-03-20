import { spawnSync } from "bun";

export function getSystemProxy(): string | null {
  // รันบน Windows เท่านั้น
  if (process.platform !== "win32") return null;

  try {
    // ใช้ PowerShell ดึงค่าจาก Registry ตรงๆ เหมือนใน Go
    const command = `Get-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings' | Select-Object ProxyEnable, ProxyServer | ConvertTo-Json`;
    const proc = spawnSync(["powershell", "-Command", command]);
    
    if (proc.exitCode !== 0) return null;

    const output = JSON.parse(proc.stdout.toString());

    // ถ้า ProxyEnable ไม่เป็น 1 (ปิดอยู่) ให้คืนค่า null
    if (output.ProxyEnable !== 1 || !output.ProxyServer) {
      return null;
    }

    return parseProxyServer(output.ProxyServer);
  } catch (e) {
    return null;
  }
}

function parseProxyServer(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;

  // กรณีมีแยกโปรโตคอล เช่น "https=proxy:8080;http=proxy:8081"
  if (trimmed.includes("=")) {
    const segments = trimmed.split(";");
    
    // หา https ก่อนตามลำดับความสำคัญ เหมือนใน Go
    const httpsProxy = segments.find(s => s.toLowerCase().startsWith("https="));
    if (httpsProxy) return ensureScheme(httpsProxy.split("=")[1] || "");

    const httpProxy = segments.find(s => s.toLowerCase().startsWith("http="));
    if (httpProxy) return ensureScheme(httpProxy.split("=")[1] || "");
  }

  return ensureScheme(trimmed);
}

function ensureScheme(value: string): string {
  const v = value.trim();
  if (v.startsWith("http://") || v.startsWith("https://")) return v;
  return `http://${v}`;
}