import { spawnSync } from "bun";
import { join } from "node:path";
import { existsSync, mkdirSync } from "node:fs";

// เก็บไฟล์ไว้ใน AppData ของ User (มาตรฐานโปรแกรม Windows)
const APPDATA = process.env.APPDATA || join(process.env.USERPROFILE || "", "AppData", "Roaming");
const CRED_DIR = join(APPDATA, "CardX-Email-Gen");
const CRED_PATH = join(CRED_DIR, ".vault_data");

export async function getOrRequestSecret(): Promise<string> {
    if (process.platform !== "win32") {
        return process.env.CLIENT_SECRET || "";
    }

    // 1. ลองดึงข้อมูลที่เข้ารหัสไว้จากไฟล์
    if (existsSync(CRED_PATH)) {
        const readScript = `
            try {
                $encrypted = Get-Content "${CRED_PATH}";
                $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect(
                    [System.Convert]::FromBase64String($encrypted),
                    $null,
                    [System.Security.Cryptography.DataProtectionScope]::CurrentUser
                );
                [System.Text.Encoding]::UTF8.GetString($unprotected);
            } catch { Write-Output "" }
        `;
        const result = spawnSync(["powershell", "-Command", readScript]);
        const savedSecret = result.stdout.toString().trim();
        
        if (savedSecret) {
            console.log("🔑 Windows: Credential loaded securely from vault.");
            return savedSecret;
        }
    }

    // 2. ถ้าไม่เจอ ให้ถามผู้ใช้ผ่าน InputBox
    console.log("🔑 Windows: No saved credential found, opening input box...");
    const promptScript = `
        Add-Type -AssemblyName Microsoft.VisualBasic;
        $input = [Microsoft.VisualBasic.Interaction]::InputBox("กรุณาใส่ Client Secret ของ CardX Graph API", "Initial Setup", "");
        $input
    `;
    const promptResult = spawnSync(["powershell", "-Command", promptScript]);
    const userInput = promptResult.stdout.toString().trim();

    if (!userInput) throw new Error("❌ Client Secret is required!");

    // 3. บันทึกและเข้ารหัสด้วย DPAPI (CurrentUser Scope)
    if (!existsSync(CRED_DIR)) mkdirSync(CRED_DIR, { recursive: true });

    const saveScript = `
        $data = [System.Text.Encoding]::UTF8.GetBytes("${userInput}");
        $encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
            $data,
            $null,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        );
        [System.Convert]::ToBase64String($encrypted) | Out-File "${CRED_PATH}" -Encoding ascii;
    `;
    spawnSync(["powershell", "-Command", saveScript]);
    
    console.log("✅ Secret encrypted and saved successfully!");
    return userInput;
}