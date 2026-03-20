import { spawnSync } from "bun";
import { env } from "../config/env";

export async function getOrRequestSecret(): Promise<string> {
    // 1. ถ้าไม่ใช่ Windows (เช่น Mac) ให้ดึงจาก .env
    if (process.platform !== "win32") {
        console.log("🍎 Running on Mac: Reading secret from .env");
        const secret = process.env.CLIENT_SECRET;
        if (!secret) {
            throw new Error("❌ ไม่พบ CLIENT_SECRET! โปรดสร้างไฟล์ .env และใส่ CLIENT_SECRET=your_key");
        }
        return secret;
    }

    // 2. ถ้าเป็น Windows ให้ดึงจาก Credential Manager (ใช้ PasswordVault เพื่ออ่าน Password จริง)
    const readScript = `
        $vault = New-Object Windows.Security.Credentials.PasswordVault;
        try {
            $cred = $vault.Retrieve("${env.credentialName}", "CardX_App");
            Write-Output $cred.Password;
        } catch {
            Write-Output "";
        }
    `;
    
    const checkResult = spawnSync(["powershell", "-Command", readScript]);
    let secret = checkResult.stdout.toString().trim();

    // 3. ถ้ายังไม่มี Secret (หรือดึงไม่สำเร็จ) ให้เด้งหน้าต่างถาม
    if (!secret) {
        console.log("🔑 Windows: Credential not found or empty, opening input box...");
        const promptScript = `
            Add-Type -AssemblyName Microsoft.VisualBasic;
            $input = [Microsoft.VisualBasic.Interaction]::InputBox("กรุณาใส่ Client Secret ของ CardX Graph API", "Setup Initial Credential", "");
            Write-Output $input
        `;
        const promptResult = spawnSync(["powershell", "-Command", promptScript]);
        const userInput = promptResult.stdout.toString().trim();

        if (!userInput) throw new Error("❌ Client Secret is required!");

        // บันทึกลง Windows Credential Manager ทันที
        // เรายังใช้ cmdkey ในการบันทึกได้ เพราะมันใช้ง่ายและบันทึกลงที่เดียวกัน
        spawnSync([
            "powershell", "-Command",
            `cmdkey /generic:${env.credentialName} /user:CardX_App /pass:"${userInput}"`
        ]);
        
        secret = userInput;
        console.log("✅ Secret saved to Windows successfully!");
    } else {
        console.log("🔑 Windows: Credential found and loaded.");
    }

    return secret;
}