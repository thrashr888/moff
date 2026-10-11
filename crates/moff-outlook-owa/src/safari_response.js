async function(response) {
    if (!response.ok) return {status: response.status};
    const reader = response.body.getReader(), decoder = new TextDecoder();
    let text = "", bytes = 0;
    try {
        while (true) {
            const part = await reader.read();
            if (part.done) break;
            bytes += part.value.length;
            if (bytes > 900000) throw new Error("Response exceeds Moff's browser limit");
            text += decoder.decode(part.value, {stream: true});
        }
        text += decoder.decode();
    } finally { await reader.cancel(); }
    return {status: response.status, body: text};
}
