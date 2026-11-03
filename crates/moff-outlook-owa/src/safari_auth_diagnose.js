function() {
    const supportedOrigins = [
        "https://login.microsoftonline.com",
        "https://login.microsoft.com",
        "https://login.live.com",
        "https://login.w3.ibm.com",
        "https://login.ibm.com",
        "https://w3id.ibm.com",
    ];
    if (!supportedOrigins.includes(location.origin)) {
        return {step: "unsupported_sign_in_page"};
    }

    const visible = selector => Array.from(document.querySelectorAll(selector)).some(element => {
        if (element.hidden || element.getAttribute("aria-hidden") === "true") return false;
        const bounds = element.getBoundingClientRect();
        return bounds.width > 0 && bounds.height > 0;
    });

    if (visible("#acceptButton")) return {step: "consent_review"};
    if (visible("#idDiv_SAOTCS_Proofs")) return {step: "verification_choice"};
    if (visible('input[autocomplete="one-time-code"], #idTxtBx_SAOTCC_OTC')) {
        return {step: "verification_code"};
    }
    if (visible('input[type="password"]')) return {step: "password"};
    if (visible("#tilesHolder .tile-container, #otherTile, [data-testid='account-tile']")) {
        return {step: "account_selection"};
    }
    if (visible('#i0116, input[name="loginfmt"], input[autocomplete="username"], input[type="email"]')) {
        return {step: "username"};
    }
    return {step: "sign_in_page"};
}
