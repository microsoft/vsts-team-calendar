import { getClient } from "azure-devops-extension-api";
import { GraphRestClient } from "azure-devops-extension-api/Graph";
import { Avatar, AvatarSize } from "azure-devops-extension-api/Profile";

import { generateColor } from "./Color";

// Resolves Azure DevOps member avatars via the Graph REST client and returns them as data URIs,
// with a locally generated initials avatar as the fallback.

const avatarCache = new Map<string, Promise<string>>();
const initialsCache = new Map<string, string>();

let graphClient: GraphRestClient | undefined;

function getGraphClient(): GraphRestClient {
    if (!graphClient) {
        graphClient = getClient(GraphRestClient);
    }
    return graphClient;
}

function avatarToDataUri(avatar: Avatar): string | null {
    const raw: any = avatar ? avatar.value : undefined;
    if (!raw) {
        return null;
    }

    let base64: string;
    if (typeof raw === "string") {
        base64 = raw;
    } else {
        // getAvatar can return the image as a byte array; encode it to base64.
        let binary = "";
        for (let i = 0; i < raw.length; i++) {
            binary += String.fromCharCode(raw[i] & 0xff);
        }
        base64 = btoa(binary);
    }

    return "data:image/png;base64," + base64;
}

function escapeXml(value: string): string {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function getInitials(displayName: string): string {
    const parts = (displayName || "").trim().split(/\s+/).filter(part => part.length > 0);
    if (parts.length === 0) {
        return "?";
    }
    if (parts.length === 1) {
        return parts[0].charAt(0).toUpperCase();
    }
    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

function buildInitialsSvg(displayName: string): string {
    const initials = escapeXml(getInitials(displayName));
    const color = generateColor(displayName || "?");
    const svg =
        '<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">' +
        '<circle cx="16" cy="16" r="16" fill="' + color + '"/>' +
        '<text x="16" y="21" font-family="Segoe UI, Arial, sans-serif" font-size="13" font-weight="600" ' +
        'fill="#ffffff" text-anchor="middle">' + initials + "</text>" +
        "</svg>";
    return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
}

// Locally generated, initials-based avatar as an SVG data URI, used as a placeholder and fallback.
export function getInitialsAvatar(displayName: string): string {
    const key = displayName || "";
    const existing = initialsCache.get(key);
    if (existing) {
        return existing;
    }
    const uri = buildInitialsSvg(key);
    initialsCache.set(key, uri);
    return uri;
}

// Resolves a member avatar to a data URI, falling back to an initials avatar when no descriptor is
// available or the lookup fails. Results are cached per identity.
export function getAvatarDataUri(descriptor: string | undefined, identityId: string, displayName: string): Promise<string> {
    if (!descriptor) {
        return Promise.resolve(getInitialsAvatar(displayName));
    }

    const cached = avatarCache.get(identityId);
    if (cached) {
        return cached;
    }

    const promise = getGraphClient()
        .getAvatar(descriptor, AvatarSize.Medium)
        .then(avatar => avatarToDataUri(avatar) || getInitialsAvatar(displayName))
        .catch(() => getInitialsAvatar(displayName));

    avatarCache.set(identityId, promise);
    return promise;
}
