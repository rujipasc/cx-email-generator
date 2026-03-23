import * as XLSX from "xlsx";
import fs from "node:fs";
import { readdirSync, mkdirSync, existsSync, renameSync } from "node:fs";
import { join, parse, resolve } from "node:path";
import { getOrRequestSecret } from "./services/credential";
import { getAccessToken, isEmailReserved } from "./services/graph";
import { loadEmailHistory, addHistoryRecord } from "./services/history";
import { getSystemProxy } from "./services/network";

/**
 * Robust Pause Logic: ค้างหน้าจอไว้รอการกดปุ่ม (ใช้ตอนจบโปรแกรม)
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
  console.log("        CardX Email Generator Tool v1.0.6"); // ขยับเวอร์ชันเล็กน้อย
  console.log("==================================================");

  const proxyUrl = getSystemProxy();
  if (proxyUrl) {
    console.log(`[NETWORK] Proxy Detected: ${proxyUrl}`);
    process.env.HTTP_PROXY = proxyUrl;
    process.env.HTTPS_PROXY = proxyUrl;
  }
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
  console.log("[NETWORK] SSL Verification: Disabled (Bypass mode)");

  const INPUT_DIR = "inputs";
  const OUTPUT_DIR = "processed_results";
  const ARCHIVE_DIR = "archives";

  try {
    [INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR].forEach(dir => {
      const p = resolve(process.cwd(), dir);
      if (!existsSync(p)) {
        console.log(`[SYSTEM] Creating directory: ${dir}`);
        mkdirSync(p, { recursive: true });
      }
    });

    const secret = await getOrRequestSecret();
    const token = await getAccessToken(secret);
    const historySet = await loadEmailHistory(token);

    const files = readdirSync(INPUT_DIR).filter(
      (file) => file.endsWith(".xlsx") && !file.startsWith("~$")
    );

    if (files.length === 0) {
      console.log("\n[!] ไม่พบไฟล์ในโฟลเดอร์ 'inputs'");
    } else {
      console.log(`\n[INFO] พบไฟล์ทั้งหมด ${files.length} ไฟล์...`);

      for (const file of files) {
        const inputPath = resolve(INPUT_DIR, file);
        console.log(`\n[PROCESS] Processing: ${file}`);

        try {
          // อ่านไฟล์แบบ Buffer (ท่าเดียวกับโปรเจกต์ Recruitment)
          const fileBuffer = fs.readFileSync(inputPath);
          const workbook = XLSX.read(fileBuffer, { type: "buffer", cellDates: true });

          const firstSheetName = workbook.SheetNames[0];
          if (!firstSheetName) continue;
          const sheet = workbook.Sheets[firstSheetName];
          const data = XLSX.utils.sheet_to_json(sheet!) as any[];
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

            for (const email of candidates) {
              console.log(`   - Checking: ${email}`);
              const isTakenInAAD = await isEmailReserved(token, email);
              const isTakenInHistory = historySet.has(email.toLowerCase());

              if (!isTakenInAAD && !isTakenInHistory) {
                finalEmail = email;
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

          // บันทึกไฟล์ผลลัพธ์
          const outputName = `result_${parse(file).name}.xlsx`;
          const outputDir = resolve(process.cwd(), OUTPUT_DIR);
          const outputPath = resolve(outputDir, outputName);

          const newWb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(newWb, XLSX.utils.json_to_sheet(results), "Result");

          const wbBuffer = XLSX.write(newWb, { bookType: 'xlsx', type: 'buffer'});
          fs.writeFileSync(outputPath, wbBuffer);
          console.log(`   ✅ [Success]: Saved to ${outputName}`);

          // ย้ายไป Archive
          const archiveDirFull = resolve(process.cwd(), ARCHIVE_DIR);
          renameSync(inputPath, join(archiveDirFull, file));
          console.log(`   📦 [ARCHIVED]: Original file moved to ${ARCHIVE_DIR}`);

        } catch (fileErr: any) {
          console.error(`   ❌ [ERROR] ในไฟล์ ${file}:`, fileErr.message);
        }
      } // ปิดลูป for (files)

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