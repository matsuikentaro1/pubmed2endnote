// background.js

class PubMedEndNoteService {
    constructor() {
        this.setupMessageListener();
    }

    setupMessageListener() {
        chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
            if (request.action === 'fetchAndFormatPMID') {
                this.handleSinglePMID(request.pmid, request.userEmail, sendResponse);
                return true; // async
            } else if (request.action === 'openSettings') {
                chrome.tabs.create({ url: chrome.runtime.getURL('options.html') });
                return false;
            }
        });
    }

    async handleSinglePMID(pmid, userEmail, sendResponse) {
        try {
            // 1. Fetch PubMed XML data
            const xmlData = await this.fetchXMLData(pmid, userEmail);

            // 2. Send the XML data to the native host for processing
            const hostName = "com.pubmed.endnote";
            chrome.runtime.sendNativeMessage(hostName, { xml_data: xmlData }, (response) => {
                // 3. Handle the response from the native host
                if (chrome.runtime.lastError) {
                    console.error("Native Messaging Error:", chrome.runtime.lastError.message);
                    sendResponse({ success: false, error: "Native host communication error: " + chrome.runtime.lastError.message });
                    return;
                }

                console.log("Received from native host:", response);
                if (response && response.status === 'success') {
                    sendResponse({ success: true, message: "Citation copied to clipboard." });
                } else {
                    const errorMessage = response ? response.message : "An unknown error occurred in the native host.";
                    console.error("Error response from native host:", errorMessage);
                    sendResponse({ success: false, error: errorMessage });
                }
            });
        } catch (error) {
            // This will catch errors from fetchXMLData
            console.error('Error processing PMID:', error);
            sendResponse({ success: false, error: error.message });
        }
    }

    async fetchXMLData(pmid, userEmail) {
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

}

new PubMedEndNoteService();