import * as React from "react";

import { getAvatarDataUri, getInitialsAvatar } from "./AvatarService";

interface IAvatarIconProps {
    // Graph subject descriptor used to fetch the avatar; undefined falls back to initials.
    descriptor?: string;
    // Unique id used to cache the resolved avatar (member id or team id).
    identityId?: string;
    displayName: string;
    className?: string;
}

// Renders a member/team avatar, showing an initials placeholder until the real avatar resolves.
export const AvatarIcon: React.FC<IAvatarIconProps> = props => {
    const { descriptor, identityId, displayName, className } = props;
    const [src, setSrc] = React.useState<string>(() => getInitialsAvatar(displayName));

    React.useEffect(() => {
        let cancelled = false;
        const cacheKey = identityId || descriptor || displayName;
        setSrc(getInitialsAvatar(displayName));
        getAvatarDataUri(descriptor, cacheKey, displayName).then(resolved => {
            if (!cancelled) {
                setSrc(resolved);
            }
        });
        return () => {
            cancelled = true;
        };
    }, [descriptor, identityId, displayName]);

    return <img alt="" className={className} src={src} />;
};
