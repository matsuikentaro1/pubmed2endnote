// Options page JavaScript for PubMed2EndNote extension

class OptionsController {
    constructor() {
        this.emailInput = document.getElementById('email-input');
        this.saveBtn = document.getElementById('save-btn');
        this.statusMessage = document.getElementById('status-message');
        
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.loadSavedSettings();
    }

    setupEventListeners() {
        // Save button
        this.saveBtn.addEventListener('click', () => {
            this.saveSettings();
        });

        // Enter key to save
        this.emailInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.saveSettings();
            }
        });

        // Input validation
        this.emailInput.addEventListener('input', () => {
            this.validateInput();
        });
    }

    // Load saved settings from storage
    async loadSavedSettings() {
        try {
            const result = await chrome.storage.sync.get(['userEmail']);
            if (result.userEmail) {
                this.emailInput.value = result.userEmail;
                this.emailInput.classList.add('saved');
                
                // Keep the saved state permanently
                this.saveBtn.textContent = 'Saved ✓';
                this.saveBtn.classList.add('saved');
                
                this.showStatus('Email address is already configured. You can close this page and start using the extension.', 'info');
                
                // Don't automatically hide this message or reset the button
                // User can still change the email if they want
            }
        } catch (error) {
            console.error('Error loading settings:', error);
            this.showStatus('Error loading settings', 'error');
        }
    }

    // Save settings to storage
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

            await chrome.storage.sync.set({ userEmail: email });
            
            // Success state
            this.showStatus('Settings saved successfully! You can now use the extension on PubMed pages.', 'success');
            
            // Update UI to saved state
            this.saveBtn.textContent = 'Saved ✓';
            this.saveBtn.classList.add('saved');
            this.emailInput.classList.add('saved');
            
            // Re-enable the button after a short delay (but keep the "Saved" text)
            setTimeout(() => {
                this.saveBtn.disabled = false;
            }, 1000);
            
            // Don't reset the button text or remove the saved class
            // This way it stays as "Saved ✓" until the user changes the email
            
        } catch (error) {
            console.error('Error saving settings:', error);
            this.showStatus('Error saving settings', 'error');
            
            // Reset button state on error
            this.saveBtn.textContent = 'Save Settings';
            this.saveBtn.classList.remove('saved');
            this.emailInput.classList.remove('saved');
            this.saveBtn.disabled = false;
        }
    }

    // Validate email input
    validateInput() {
        const email = this.emailInput.value.trim();
        const isValid = email && this.isValidEmail(email);
        
        // If the email has changed from the saved value, reset the saved state
        chrome.storage.sync.get(['userEmail'], (result) => {
            if (result.userEmail && result.userEmail !== email) {
                this.saveBtn.textContent = 'Save Settings';
                this.saveBtn.classList.remove('saved');
                this.emailInput.classList.remove('saved');
            }
        });
        
        this.saveBtn.disabled = !isValid;
    }

    // Email validation helper
    isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    }

    // Show status message
    showStatus(message, type) {
        this.statusMessage.textContent = message;
        this.statusMessage.className = `status-message ${type}`;
        this.statusMessage.style.display = 'block';
        
        // For success and error messages, auto-hide after 5 seconds
        // For info messages about existing config, don't auto-hide
        if (type === 'success' || type === 'error') {
            setTimeout(() => {
                this.hideStatus();
            }, 5000);
        }
    }

    // Hide status message
    hideStatus() {
        this.statusMessage.style.display = 'none';
    }
}

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new OptionsController();
});