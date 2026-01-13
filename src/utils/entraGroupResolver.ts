// Utility for expanding Entra ID (Azure AD) groups into user objects using Microsoft Graph API

export interface EntraUser {
    id: string;
    displayName: string;
    userPrincipalName: string;
    mail?: string;
}

/**
 * Expands all user members of an Entra ID (AAD) group via Microsoft Graph API.
 * @param groupId Object ID of the group in Azure AD
 * @param accessToken Microsoft Graph Bearer Token with GroupMember.Read.All and User.ReadBasic.All
 */
export async function expandEntraIdGroupMembers(groupId: string, accessToken: string): Promise<EntraUser[]> {
    const endpoint = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,mail`;
    const headers = { Authorization: `Bearer ${accessToken}` };
    const users: EntraUser[] = [];
    let url = endpoint;
    while (url) {
        const response = await fetch(url, { headers });
        if (!response.ok) throw new Error(`Graph API request failed: ${response.status} ${response.statusText}`);
        const data = await response.json();
        if (!data.value) break;
        for (const entry of data.value) {
            // Only process user objects, skip others (servicePrincipal, device, group)
            if (entry['@odata.type'] === '#microsoft.graph.user') {
                users.push({
                    id: entry.id,
                    displayName: entry.displayName,
                    userPrincipalName: entry.userPrincipalName,
                    mail: entry.mail
                });
            }
        }
        url = data['@odata.nextLink'];
    }
    return users;
}