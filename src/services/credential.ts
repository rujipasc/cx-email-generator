import { spawnSync } from "bun";
import { env } from "../config/env";

export async function getOrRequestSecret(): Promise<string> {
    // 1. ถ้าไม่ใช่ Windows (เช่น Mac ของคุณผึ้ง) ให้ดึงจาก .env
    if (process.platform !== "win32") {
        console.log("🍎 Running on Mac: Reading secret from .env");
        const secret = process.env.CLIENT_SECRET;
        if (!secret) {
            throw new Error("❌ ไม่พบ CLIENT_SECRET! โปรดสร้างไฟล์ .env และใส่ CLIENT_SECRET=your_key");
        }
        return secret;
    }

    // 2. ถ้าเป็น Windows ให้ดึงจาก Credential Manager
    const checkResult = spawnSync([
        "powershell", "-Command",
        `$p = (cmdkey /list:${env.credentialName} | Select-String "Password"); if ($p) { ($p.ToString() -split ':', 2)[1].Trim() }`
    ]);

    let secret = checkResult.stdout.toString().trim();

    // 3. ถ้าใน Windows ยังไม่มี Key (รันครั้งแรก) ให้เด้งหน้าต่างถาม
    if (!secret) {
        console.log("🔑 Windows: Credential not found, opening input box...");
        const promptScript = `
            Add-Type -AssemblyName Microsoft.VisualBasic;
            $input = [Microsoft.VisualBasic.Interaction]::InputBox("กรุณาใส่ Client Secret ของ CardX Graph API", "Setup Initial Credential", "");
            Write-Output $input
        `;
        const promptResult = spawnSync(["powershell", "-Command", promptScript]);
        const userInput = promptResult.stdout.toString().trim();

        if (!userInput) throw new Error("❌ Client Secret is required!");

        // บันทึกลง Windows Credential Manager ทันที
        spawnSync([
            "powershell", "-Command",
            `cmdkey /generic:${env.credentialName} /user:CardX_App /pass:"${userInput}"`
        ]);
        
        secret = userInput;
        console.log("✅ Secret saved to Windows successfully!");
    }

    return secret;
}