// background.js
// Only opens the options page on request from the content script
// (chrome.runtime.openOptionsPage is not available in content scripts).

chrome.runtime.onMessage.addListener((request) => {
    if (request.action === 'openSettings') {
        chrome.runtime.openOptionsPage();
    }
});
