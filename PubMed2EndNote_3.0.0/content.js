// content.js

(function init() {
    const pmidMatch = window.location.href.match(/pubmed\.ncbi\.nlm\.nih\.gov\/(\d+)/);
    if (!pmidMatch) return;

    const pmid = pmidMatch[1];
    const container = document.createElement("div");
    container.id = "pubmed-endnote-shadow-container";
    Object.assign(container.style, { position: "fixed", top: "10px", right: "10px", zIndex: "999999" });
    document.body.appendChild(container);

    const shadowRoot = container.attachShadow({ mode: "open" });
    const icon = document.createElement("img");
    icon.src = chrome.runtime.getURL("icons/icon48.png");
    icon.alt = "EndNote Export";
    icon.title = `Export PMID: ${pmid} to EndNote format`;
    Object.assign(icon.style, {
        width: "48px", height: "48px", cursor: "pointer", borderRadius: "50%",
        boxShadow: "0 2px 6px rgba(0,0,0,0.2)", transition: "all 0.2s ease",
        background: "linear-gradient(135deg, #1e3a8a, #3b82f6)", padding: "8px"
    });

    icon.addEventListener("mouseenter", () => { Object.assign(icon.style, { transform: "scale(1.1)", boxShadow: "0 4px 12px rgba(0,0,0,0.3)" }); });
    icon.addEventListener("mouseleave", () => { Object.assign(icon.style, { transform: "scale(1)", boxShadow: "0 2px 6px rgba(0,0,0,0.2)" }); });
    shadowRoot.appendChild(icon);

    // Right-click on the icon shows a small menu
    const menu = document.createElement("div");
    Object.assign(menu.style, {
        display: "none", position: "absolute", top: "52px", right: "0",
        background: "white", borderRadius: "8px", boxShadow: "0 4px 16px rgba(0,0,0,0.25)",
        padding: "4px", minWidth: "160px", fontFamily: "sans-serif", fontSize: "14px"
    });
    const menuEntries = [
        { label: "Export to Clipboard", onClick: () => exportCitation() },
        { label: "Settings", onClick: () => chrome.runtime.sendMessage({ action: 'openSettings' }) },
    ];
    for (const entry of menuEntries) {
        const item = document.createElement("div");
        item.textContent = entry.label;
        Object.assign(item.style, {
            padding: "8px 12px", borderRadius: "6px", cursor: "pointer", color: "#1a202c", whiteSpace: "nowrap"
        });
        item.addEventListener("mouseenter", () => { item.style.background = "#eef2ff"; });
        item.addEventListener("mouseleave", () => { item.style.background = "transparent"; });
        item.addEventListener("click", () => {
            menu.style.display = "none";
            entry.onClick();
        });
        menu.appendChild(item);
    }
    shadowRoot.appendChild(menu);

    icon.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        menu.style.display = "block";
        document.addEventListener("click", () => { menu.style.display = "none"; }, { once: true });
    });

    const exportCitation = async () => {
        const originalTitle = icon.title;
        icon.style.cursor = "wait";

        try {
            const result = await chrome.storage.sync.get(['userEmail', 'fontFamily', 'fontSize', 'highlightEnabled']);
            if (!result.userEmail || result.userEmail.trim() === '') {
                icon.title = "Opening email setup...";
                icon.style.background = "linear-gradient(135deg, #f59e0b, #fbbf24)";
                chrome.runtime.sendMessage({ action: 'openSettings' });
                return;
            }

            icon.title = "Processing...";
            const xmlData = await fetchPubMedXml(pmid, result.userEmail.trim());

            // Conversion + clipboard write happen right here in the page (citation.js)
            await convertAndCopy(xmlData, {
                fontFamily: result.fontFamily || 'Times New Roman',
                fontSize: result.fontSize || '10.5',
                highlight: result.highlightEnabled !== false
            });

            icon.title = "✓ Successfully copied to clipboard!";
            icon.style.background = "linear-gradient(135deg, #059669, #10b981)";
            showNotification("Successfully copied to clipboard!", "success");
        } catch (error) {
            console.error('Error:', error);
            icon.title = `❌ Error: ${error.message}`;
            icon.style.background = "linear-gradient(135deg, #dc2626, #ef4444)";
            showNotification(`Error: ${error.message}`, "error");
        } finally {
            icon.style.cursor = "pointer";
            setTimeout(() => {
                icon.title = originalTitle;
                icon.style.background = "linear-gradient(135deg, #1e3a8a, #3b82f6)";
            }, 3000);
        }
    };

    icon.addEventListener("click", exportCitation);


    function showNotification(message, type) {
        const notification = document.createElement("div");
        notification.style.cssText = `
            position: fixed; top: 70px; right: 10px; z-index: 1000000; padding: 12px 20px;
            border-radius: 8px; color: white; font-family: sans-serif; font-size: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15); transform: translateX(120%);
            transition: transform 0.3s ease; max-width: 300px; word-wrap: break-word;
            background: ${type === 'success' ? 'linear-gradient(135deg, #059669, #10b981)' : 'linear-gradient(135deg, #dc2626, #ef4444)'};
        `;
        notification.textContent = message;
        document.body.appendChild(notification);
        setTimeout(() => { notification.style.transform = "translateX(0)"; }, 10);
        setTimeout(() => {
            notification.style.transform = "translateX(120%)";
            setTimeout(() => { notification.remove(); }, 300);
        }, 3000);
    }
})();
