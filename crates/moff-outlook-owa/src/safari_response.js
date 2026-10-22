async function(response) {
    if (!response.ok) return {status: response.status};
    const reader = response.body.getReader(), decoder = new TextDecoder();
    let text = "", bytes = 0;
    try {
        while (true) {
            const part = await reader.read();
            if (part.done) break;
            bytes += part.value.length;
            if (bytes > 900000) {
                const error = new Error("Response exceeds Moff's browser limit");
                error.code = "moff_response_too_large";
                throw error;
            }
            text += decoder.decode(part.value, {stream: true});
        }
        text += decoder.decode();
    } finally { await reader.cancel(); }
    return {status: response.status, body: text};
}
