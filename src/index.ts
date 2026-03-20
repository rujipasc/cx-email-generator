import * as XLSX from "xlsx";
import { readdirSync, mkdirSync, existsSync, renameSync } from "node:fs";
import { join, parse } from "node:path";
import { getOrRequestSecret } from "./services/credential";
import { getAccessToken, isEmailReserved } from "./services/graph";
import { loadEmailHistory, addHistoryRecord } from "./services/history";
import { getSystemProxy } from "./services/network";

/**
 * Robust Pause Logic: ค้างหน้าจอไว้รอการกดปุ่ม
 */
async function waitAnyKey() {
  console.log("\n--------------------------------------------------");
  console.log("Press any key to exit...");
  process.stdin.setRawMode(true);
  process.stdin.resume();
  return new Promise<void>((resolve) => {
    process.stdin.once("data", (data) => {
      if (data[0] === 3) process.exit(1); // Ctrl+C
      process.stdin.setRawMode(false);
      process.stdin.pause();
      resolve();
    });
  });
}

async function run() {
  console.clear();
  console.log("==================================================");
  console.log("        CardX Email Generator Tool v1.0.0");
  console.log("==================================================");

  // 1. Network & Security Setup
  const proxyUrl = getSystemProxy();
  if (proxyUrl) {
    console.log(`[NETWORK] Proxy Detected: ${proxyUrl}`);
    process.env.HTTP_PROXY = proxyUrl;
    process.env.HTTPS_PROXY = proxyUrl;
  }
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"; // Bypass SSL สำหรับเน็ตองค์กร
  console.log("[NETWORK] SSL Verification: Disabled (Bypass mode)");

  const INPUT_DIR = "inputs";
  const OUTPUT_DIR = "processed_results";
  const ARCHIVE_DIR = "archives";

  try {
    // 2. Folder Automation
    [INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR].forEach(dir => {
      if (!existsSync(dir)) {
        console.log(`[SYSTEM] Creating directory: ${dir}`);
        mkdirSync(dir, { recursive: true });
      }
    });

    // 3. Authentication & History Sync
    const secret = await getOrRequestSecret();
    const token = await getAccessToken(secret);
    const historySet = await loadEmailHistory(token);

    // 4. File Scanning
    const files = readdirSync(INPUT_DIR).filter(
      (file) => file.endsWith(".xlsx") && !file.startsWith("~$")
    );

    if (files.length === 0) {
      console.log("\n[!] ไม่พบไฟล์ในโฟลเดอร์ 'inputs'");
      console.log("[!] กรุณานำไฟล์ Excel มาวางแล้วรันโปรแกรมใหม่อีกครั้ง");
    } else {
      console.log(`\n[INFO] พบไฟล์ทั้งหมด ${files.length} ไฟล์...`);

      for (const file of files) {
        const inputPath = join(INPUT_DIR, file);
        console.log(`\n[PROCESS] Processing: ${file}`);

        const workbook = XLSX.readFile(inputPath);
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) continue;
        const sheet = workbook.Sheets[firstSheetName];
        if (!sheet) continue;

        const data = XLSX.utils.sheet_to_json(sheet) as any[];
        const results = [];

        for (const row of data) {
          const rawFName = (row.name || row.Name || "").toString().trim();
          const rawLName = (row.surname || row.Surname || "").toString().trim();
          const fName = rawFName.toLowerCase().replace(/\s+/g, "");
          const lName = rawLName.toLowerCase().replace(/\s+/g, "");

          if (!fName || !lName) {
            results.push({ ...row, email: "ERROR: Missing Name/Surname" });
            continue;
          }

          let finalEmail = "MANUAL_REQUIRED";
          const domain = "@cardx.co.th";
          const candidates = [`${fName}${domain}`];
          for (let i = 1; i <= 7; i++) {
            candidates.push(`${fName}.${lName.substring(0, i)}${domain}`);
          }

          // ตรวจสอบ 3 ด่าน (AAD, Recycle Bin, SharePoint History)
          for (const email of candidates) {
            console.log(`   - Checking: ${email}`);
            const isTakenInAAD = await isEmailReserved(token, email);
            const isTakenInHistory = historySet.has(email.toLowerCase());

            if (!isTakenInAAD && !isTakenInHistory) {
              finalEmail = email;
              // บันทึกลง SharePoint ทันทีที่จองสำเร็จ
              await addHistoryRecord(token, {
                name: rawFName,
                surname: rawLName,
                email: finalEmail,
                empId: (row.empId || row.EmployeeID || row["Employee ID"] || "").toString(),
                fileName: file
              });
              historySet.add(finalEmail.toLowerCase());
              break;
            }
          }
          results.push({ ...row, email: finalEmail });
        }

        // 5. Save & Archive
        const outputName = `result_${parse(file).name}.xlsx`;
        const outputPath = join(OUTPUT_DIR, outputName);
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, XLSX.utils.json_to_sheet(results), "Result");
        XLSX.writeFile(newWb, outputPath);

        renameSync(inputPath, join(ARCHIVE_DIR, file));
        console.log(`[SUCCESS] Saved to: ${outputPath}`);
      }
      console.log("\n==================================================");
      console.log("         ✨ All processes completed! ✨");
      console.log("==================================================");
    }
  } catch (e: any) {
    console.error("\n❌ [FATAL ERROR]:", e.message);
  } finally {
    await waitAnyKey();
    process.exit(0);
  }
}

run();