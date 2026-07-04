// Options page JavaScript for PubMed2EndNote extension

class OptionsController {
    constructor() {
        this.emailInput = document.getElementById('email-input');
        this.fontSelect = document.getElementById('font-select');
        this.sizeSelect = document.getElementById('size-select');
        this.saveBtn = document.getElementById('save-btn');
        this.statusMessage = document.getElementById('status-message');

        this.init();
    }

    init() {
        this.setupEventListeners();
        this.loadSavedSettings();
    }

    setupEventListeners() {
        this.saveBtn.addEventListener('click', () => {
            this.saveSettings();
        });

        this.emailInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.saveSettings();
            }
        });

        // Any change resets the "Saved" state so the user notices unsaved edits
        for (const el of [this.emailInput, this.fontSelect, this.sizeSelect]) {
            el.addEventListener('input', () => this.markDirty());
            el.addEventListener('change', () => this.markDirty());
        }
    }

    markDirty() {
        this.saveBtn.textContent = 'Save Settings';
        this.saveBtn.classList.remove('saved');
        this.emailInput.classList.remove('saved');
        this.saveBtn.disabled = false;
    }

    async loadSavedSettings() {
        try {
            const result = await chrome.storage.sync.get(['userEmail', 'fontFamily', 'fontSize']);
            if (result.userEmail) {
                this.emailInput.value = result.userEmail;
                this.emailInput.classList.add('saved');
                this.saveBtn.textContent = 'Saved ✓';
                this.saveBtn.classList.add('saved');
            }
            this.fontSelect.value = result.fontFamily || '';
            this.sizeSelect.value = result.fontSize || '';
        } catch (error) {
            console.error('Error loading settings:', error);
            this.showStatus('Error loading settings', 'error');
        }
    }

    async saveSettings() {
        const email = this.emailInput.value.trim();

        if (!email) {
            this.showStatus('Please enter an email address', 'error');
            return;
        }

        if (!this.isValidEmail(email)) {
            this.showStatus('Please enter a valid email address', 'error');
            return;
        }

        try {
            this.saveBtn.textContent = 'Saving...';
            this.saveBtn.disabled = true;

            await chrome.storage.sync.set({
                userEmail: email,
                fontFamily: this.fontSelect.value,
                fontSize: this.sizeSelect.value
            });

            this.showStatus('Settings saved! You can now use the extension on PubMed pages.', 'success');
            this.saveBtn.textContent = 'Saved ✓';
            this.saveBtn.classList.add('saved');
            this.emailInput.classList.add('saved');

            setTimeout(() => {
                this.saveBtn.disabled = false;
            }, 1000);
        } catch (error) {
            console.error('Error saving settings:', error);
            this.showStatus('Error saving settings', 'error');
            this.markDirty();
        }
    }

    isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    }

    showStatus(message, type) {
        this.statusMessage.textContent = message;
        this.statusMessage.className = `status-message ${type}`;
        this.statusMessage.style.display = 'block';

        if (type === 'success' || type === 'error') {
            setTimeout(() => {
                this.statusMessage.style.display = 'none';
            }, 5000);
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new OptionsController();
});
