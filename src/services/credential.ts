import { spawnSync } from "bun";
import { join } from "node:path";
import { existsSync, mkdirSync } from "node:fs";

const APPDATA = process.env.APPDATA || join(process.env.USERPROFILE || "", "AppData", "Roaming");
const CRED_DIR = join(APPDATA, "CardX-Email-Gen");
const CRED_PATH = join(CRED_DIR, "secure_vault.xml");

export async function getOrRequestSecret(): Promise<string> {
    if (process.platform !== "win32") {
        return process.env.CLIENT_SECRET || "";
    }

    // 1. ลองโหลดจากไฟล์ XML (ถ้ามี)
    if (existsSync(CRED_PATH)) {
        console.log(`🔍 Checking vault at: ${CRED_PATH}`);
        
        // ใช้ Import-CliXml ซึ่งจะถอดรหัสให้อัตโนมัติด้วย Windows Identity ของเรา
        const loadScript = `
            try {
                $cred = Import-CliXml -Path "${CRED_PATH}";
                $plainPassword = $cred.GetNetworkCredential().Password;
                Write-Output $plainPassword;
            } catch {
                Write-Output "";
            }
        `;
        
        const result = spawnSync(["powershell", "-Command", loadScript]);
        const savedSecret = result.stdout.toString().trim();

        if (savedSecret && savedSecret.length > 0) {
            console.log("✅ [DEBUG]: Saved secret found and loaded successfully.");
            return savedSecret;
        } else {
            console.log("⚠️ [DEBUG]: Vault file exists but returned empty secret.");
        }
    }

    // 2. ถ้าไม่เจอ หรือโหลดไม่ได้ ให้ถาม User
    console.log("🔑 Windows: Requesting Client Secret via InputBox...");
    const promptScript = `
        Add-Type -AssemblyName Microsoft.VisualBasic;
        $input = [Microsoft.VisualBasic.Interaction]::InputBox("กรุณาใส่ Client Secret ของ CardX Graph API", "CardX Setup", "");
        $input;
    `;
    const promptResult = spawnSync(["powershell", "-Command", promptScript]);
    const userInput = promptResult.stdout.toString().trim();

    if (!userInput) {
        throw new Error("❌ Client Secret is required to proceed.");
    }

    // 3. บันทึกแบบเข้ารหัสลงไฟล์ XML ทันที
    try {
        if (!existsSync(CRED_DIR)) mkdirSync(CRED_DIR, { recursive: true });

        // สร้าง PSCredential object แล้วเซฟเป็น XML (Windows จะเข้ารหัสให้เอง)
        const saveScript = `
            $secString = ConvertTo-SecureString "${userInput}" -AsPlainText -Force;
            $cred = New-Object System.Management.Automation.PSCredential("CardX_User", $secString);
            $cred | Export-CliXml -Path "${CRED_PATH}";
        `;
        
        const saveResult = spawnSync(["powershell", "-Command", saveScript]);
        
        if (saveResult.exitCode === 0) {
            console.log("💾 [DEBUG]: Secret has been encrypted and saved to vault.");
        } else {
            console.error("❌ [DEBUG]: Failed to save vault:", saveResult.stderr.toString());
        }
    } catch (err) {
        console.error("❌ Error during saving process:", err);
    }

    return userInput;
}