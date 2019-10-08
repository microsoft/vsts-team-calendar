export function getQueryVariable(variable: string): string | undefined {
    var query = window.location.search.substring(1);
    var vars = query.split("&");
    for (var i = 0; i < vars.length; i++) {
        var pair = vars[i].split("=");
        if (decodeURIComponent(pair[0]) == variable) {
            return decodeURIComponent(pair[1]);
        }
    }
}

export function setTeamQueryVariable(teamId: string) {
    window.history.replaceState({}, "", window.location.href + "?team=" + teamId);
}
