import { env } from "../config/env";

export async function getAccessToken(secret: string): Promise<string> {
    const url = `https://login.microsoftonline.com/${env.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
        client_id: env.clientId,
        client_secret: secret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
    })
    const res = await fetch(url, {method: "POST", body});
    const json = await res.json() as any;
    return json.access_token;
}

export async function isEmailReserved(token: string, email: string): Promise<boolean> {
    const activePath = `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${email}' or mail eq '${email}'&$select=id`;
    const deletedPath = `https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$filter=userPrincipalName eq '${email}'&$select=id`;

    const [activeRes, deletedRes] = await Promise.all([
        fetch(activePath, {headers: {Authorization: `Bearer ${token}`}}),
        fetch(deletedPath, {headers: {Authorization: `Bearer ${token}`}})
    ]);

    const activeDate = await activeRes.json() as any;
    const deletedDate = await deletedRes.json() as any;

    return activeDate.value?.length > 0 || deletedDate.value.length > 0;
}