function(marker) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft"].includes(location.origin)) {
        throw new Error("Moff refresh requires an HTTPS Outlook document");
    }
    sessionStorage.setItem("__moff_auth_recovery", marker);
    // Let Cider retrieve the result before replacing the current document.
    setTimeout(() => location.reload(), 1500);
    return {scheduled: true};
}
