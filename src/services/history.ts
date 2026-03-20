import { env } from "../config/env";

  const SITE_ID =
    "cardxth.sharepoint.com,1f2215c9-8c61-4140-b562-035c7da1b884,691565de-2414-4c21-acca-9392bde3c942";
  const LIST_ID = "02184cc3-0684-4024-a83b-fc14b17f5aaa";

  export async function loadEmailHistory(token: string): Promise<Set<string>> {
    const historySet = new Set<string>();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?$expand=fields($select=Email)&$top=5000`;
    console.log("📥 Loading email history from SharePoint...")

    try {
        const res = await fetch(url, {
            headers: { Authorization: `Bearer ${token}` }
        });

        const json = await res.json() as any;

        if (json.value) {
            json.value.forEach((item: any) => {
                if (item.fields && item.fields.Email) {
                    historySet.add(item.fields.Email.toLowerCase().trim());
                }
            });
        }
        console.log(`✅ Loaded ${historySet.size} email addresses from SharePoint.`);
    } catch (error) {
        console.error("❌ Error loading email history:", error)
    }

    return historySet;

  }

export async function addHistoryRecord(token: string, data: {
    name: string,
    surname: string,
    email: string,
    empId?: string,
    fileName: string
}) {
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`;

    const body = {
        fields: {
            Name: data.name,
            Surname: data.surname,
            Email: data.email,
            EmployeeId: data.empId || "",
            GeneratedDate: new Date().toISOString(),
            SourceFile: data.fileName
        }
    };

    const res = await fetch(url, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    if (!res.ok) {
        console.error(`❌ Error adding record to SharePoint:`, await res.text());
    }
}