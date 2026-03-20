import { getOrRequestSecret } from "./services/credential";
import { getAccessToken } from "./services/graph";

async function migrateCsvToSharePoint() {
  const SITE_ID =
    "cardxth.sharepoint.com,1f2215c9-8c61-4140-b562-035c7da1b884,691565de-2414-4c21-acca-9392bde3c942";
  const LIST_ID = "02184cc3-0684-4024-a83b-fc14b17f5aaa";
  const FILE_PATH = "EmployeeProfile.csv";

  try {
    const secret = await getOrRequestSecret();
    const token = await getAccessToken(secret);

    const file = Bun.file(FILE_PATH);
    const text = await file.text();
    const rows = text.split("\n").map((row) => {
      return row
        .split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/)
        .map((v) => v.replace(/^"|"$/g, "").trim());
    });

    const headers = rows[0];
    if (!headers) throw new Error("❌ CSV are empty!");

    // 2. หาตำแหน่ง Index และเช็คว่าเจอครบไหม (TypeScript จะรู้ว่าหลังจากจุดนี้เป็น number แน่นอน)
    const fNameIdx = headers.indexOf("First Name (Global)");
    const lNameIdx = headers.indexOf("Last Name (Global)");
    const emailIdx = headers.indexOf("Office Email");
    const empIdIdx = headers.indexOf("Employee ID");

    if (
      fNameIdx === -1 ||
      lNameIdx === -1 ||
      emailIdx === -1 ||
      empIdIdx === -1
    ) {
      console.error("Header ที่หาเจอ:", headers);
      throw new Error(
        "❌ ไม่พบคอลัมน์ที่ต้องการใน CSV (ตรวจสอบชื่อ Header ให้ตรงเป๊ะ)",
      );
    }
    const dataRows = rows.slice(1).filter((r) => r.length > 1);
    console.log(`🚀 Starting migration for ${dataRows.length} records...`);

    for (let i = 0; i < dataRows.length; i += 20) {
      const chunk = dataRows.slice(i, i + 20);

      const batchRequests = chunk.map((row, index) => ({
        id: (i + index).toString(),
        method: "POST",
        url: `/sites/${SITE_ID}/lists/${LIST_ID}/items`,
        headers: { "Content-Type": "application/json" },
        body: {
          fields: {
            // ใช้ Index ที่เราเช็คแล้วว่าไม่ใช่ -1 หรือ undefined
            Name: row[fNameIdx],
            Surname: row[lNameIdx],
            Email: (row[emailIdx] || "").toLowerCase().trim(),
            EmployeeId: row[empIdIdx],
            GeneratedDate: new Date().toISOString(),
            SourceFile: "Initial_Migration_CSV",
          },
        },
      }));

      const response = await fetch("https://graph.microsoft.com/v1.0/$batch", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ requests: batchRequests }),
      });
      if (response.ok) {
        console.log(`✅ Progress: ${i + chunk.length}/${dataRows.length}`)
      } else {
        console.error(`❌ Batch Error at index ${i}:`, await response.text());
      }
      await new Promise(r => setTimeout(r, 300));
    }

    console.log(`\n✨ Migration Completed Successfully!`)
  } catch (e: any) {
    console.error("\n❌ Error:", e.message);
  }
}

migrateCsvToSharePoint();