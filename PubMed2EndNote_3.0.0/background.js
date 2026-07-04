// background.js
// Fetches PubMed XML (cross-origin, so it lives in the service worker).
// All conversion and clipboard work happens in the content script (citation.js).

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'fetchPubMedXML') {
        fetchXMLData(request.pmid, request.userEmail)
            .then(xmlData => sendResponse({ success: true, xmlData }))
            .catch(error => sendResponse({ success: false, error: error.message }));
        return true; // async
    } else if (request.action === 'openSettings') {
        chrome.runtime.openOptionsPage();
        return false;
    }
});

async function fetchXMLData(pmid, userEmail) {
    const url = `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=${pmid}&rettype=xml&retmode=xml&email=${encodeURIComponent(userEmail)}`;
    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`PubMed API error: ${response.status}`);
    }
    const xmlData = await response.text();
    if (!xmlData || xmlData.trim() === '') {
        throw new Error(`No data found for PMID: ${pmid}`);
    }
    return xmlData;
}
