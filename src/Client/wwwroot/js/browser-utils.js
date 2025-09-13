export function openUrlInNewTab(url) {
    const tab = window.open("about:blank", "_blank");
    if (tab) {
        tab.location.href = url;
    } else {
        console.warn("Could not open tab. It may have been blocked.");
    }
}
