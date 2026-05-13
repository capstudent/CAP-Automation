// Version: 2026-01-17-v1 - Added column selection dropdowns, abort support
// Use the same origin as the page when served by Flask (production); fall back to
// localhost:5001 when opening index.html directly from disk during development.
const API_BASE_URL = (window.location.protocol === 'file:' || !window.location.origin)
    ? 'http://localhost:5001/api'
    : `${window.location.origin}/api`;

// Show/hide Start and Abort buttons during long-running operations
function setOperationRunning(tab, running) {
    const startBtn = document.getElementById(`${tab}-start-btn`);
    const abortBtn = document.getElementById(`${tab}-abort-btn`);
    if (startBtn) startBtn.disabled = running;
    if (abortBtn) {
        abortBtn.disabled = !running;
        abortBtn.style.opacity = running ? '1' : '0.5';
        abortBtn.title = running ? 'Click to stop the current operation' : 'Available when operation is running';
    }
}

async function abortAutomation(tab) {
    // Set the global JS abort flag so the chunk loop stops between chunks too.
    window._chunkAbortRequested = true;
    try {
        await apiCall('/automation/abort', 'POST');
        showStatus(`${tab}-status`, 'Abort requested. Stopping after current item...', 'info');
    } catch (error) {
        showStatus(`${tab}-status`, `Abort failed: ${error.message}`, 'error');
    }
}

// Process a list of items in chunks to avoid Render's HTTP request timeout.
// Each chunk is a separate API call; results are aggregated. Aborts between chunks.
//   items         - array of items to process
//   chunkSize     - max items per API call (default 25)
//   runChunkFn    - async (chunk, chunkIndex, totalChunks) => { results, successful, ... }
//   options       - { statusElementId, label, onChunkDone }
async function runChunked(items, chunkSize, runChunkFn, options = {}) {
    const total = items.length;
    const totalChunks = Math.max(1, Math.ceil(total / chunkSize));
    const allResults = [];
    let totalSuccessful = 0;
    window._chunkAbortRequested = false;

    for (let chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++) {
        if (window._chunkAbortRequested) {
            if (options.statusElementId) {
                showStatus(options.statusElementId,
                    `Aborted after chunk ${chunkIndex}/${totalChunks} (${allResults.length}/${total} processed)`,
                    'info');
            }
            break;
        }

        const start = chunkIndex * chunkSize;
        const chunk = items.slice(start, start + chunkSize);
        const label = options.label || 'items';

        if (options.statusElementId) {
            showStatus(options.statusElementId,
                `Processing chunk ${chunkIndex + 1}/${totalChunks} (${chunk.length} ${label}, ${allResults.length}/${total} done)...`,
                'info');
        }

        let chunkResult;
        try {
            chunkResult = await runChunkFn(chunk, chunkIndex, totalChunks);
        } catch (e) {
            const msg = `Failed at chunk ${chunkIndex + 1}/${totalChunks} (${allResults.length}/${total} done before failure): ${e.message}`;
            if (options.statusElementId) showStatus(options.statusElementId, msg, 'error');
            return { results: allResults, successful: totalSuccessful, total, error: e, chunksCompleted: chunkIndex };
        }

        if (chunkResult && Array.isArray(chunkResult.results)) {
            allResults.push(...chunkResult.results);
        }
        if (chunkResult && typeof chunkResult.successful === 'number') {
            totalSuccessful += chunkResult.successful;
        }
        if (options.onChunkDone) {
            try { options.onChunkDone(chunkIndex, totalChunks, chunkResult); } catch {}
        }
    }

    return { results: allResults, successful: totalSuccessful, total, chunksCompleted: totalChunks };
}

// Fetch a single sheet column as an array of trimmed, non-empty strings (skips header row).
async function fetchSheetColumn(sheetUrl, sheetName, columnIndex, oauthState) {
    const result = await apiCall('/sheets/connect', 'POST', {
        sheet_url: sheetUrl,
        sheet_name: sheetName,
        oauth_state: oauthState
    });
    if (!result || !result.success) {
        throw new Error(result?.error || 'Could not read sheet');
    }
    const columns = result.columns || [];
    if (columnIndex >= columns.length) {
        throw new Error(`Column index ${columnIndex} not found (sheet has ${columns.length} columns)`);
    }
    const col = columns[columnIndex] || [];
    return col.slice(1).map(v => (v || '').toString().trim()).filter(Boolean);
}

const CHUNK_SIZE = 25;

// Google OAuth token storage
let googleAccessToken = null;
let googleClientId = null;

// Global sheet connection info
let connectedSheetUrl = null;
let connectedSheetName = null;
let isSheetConnected = false;

// Helper function to get the connected sheet info
function getConnectedSheetInfo() {
    // Use global variables first, fall back to DOM elements (Connection tab)
    const url = connectedSheetUrl || document.getElementById('sheet-url')?.value || '';
    const name = connectedSheetName || document.getElementById('sheet-name')?.value || '';
    console.log('getConnectedSheetInfo - url:', url, 'name:', name);
    return { url, name };
}

// Helper function to check if sheet is connected and show error if not
function requireSheetConnection(statusElementId) {
    console.log('requireSheetConnection called, isSheetConnected:', isSheetConnected);
    console.log('connectedSheetUrl:', connectedSheetUrl, 'connectedSheetName:', connectedSheetName);
    
    if (!isSheetConnected || !connectedSheetUrl || !connectedSheetName) {
        if (statusElementId) {
            showStatus(statusElementId, 'Please connect to a Google Sheet first (in the Connection tab)', 'error');
        }
        return null;
    }
    return { url: connectedSheetUrl, name: connectedSheetName };
}

// Helper function to refresh sheet connection (re-read data)
async function refreshSheetConnection() {
    const sheetInfo = getConnectedSheetInfo();
    if (!sheetInfo.url || !sheetInfo.name) {
        return false;
    }
    
    try {
        const oauthState = getOAuthState();
        await apiCall('/sheets/connect', 'POST', {
            sheet_url: sheetInfo.url,
            sheet_name: sheetInfo.name,
            oauth_state: oauthState
        });
        return true;
    } catch (error) {
        console.error('Failed to refresh sheet connection:', error);
        return false;
    }
}

// Update connection indicators in all tabs
function updateConnectionIndicators() {
    const indicators = ['add', 'revoke', 'status', 'convert', 'compare'];
    
    console.log('Updating indicators, isSheetConnected:', isSheetConnected);
    console.log('connectedSheetUrl:', connectedSheetUrl);
    console.log('connectedSheetName:', connectedSheetName);
    
    indicators.forEach(tab => {
        const indicator = document.getElementById(`${tab}-sheet-indicator`);
        console.log(`Looking for ${tab}-sheet-indicator:`, indicator);
        if (indicator) {
            if (isSheetConnected && connectedSheetUrl && connectedSheetName) {
                indicator.innerHTML = `🟢 Connected to: <strong>${connectedSheetName}</strong>`;
                indicator.parentElement.style.background = '#e8f5e9';
                indicator.parentElement.style.borderLeft = '4px solid #4caf50';
            } else {
                indicator.innerHTML = '⚪ No sheet connected - <a href="#" onclick="showTab(\'connection\'); return false;">Connect first</a>';
                indicator.parentElement.style.background = '#fff3e0';
                indicator.parentElement.style.borderLeft = '4px solid #ff9800';
            }
        }
    });
}

// Initialize Google Identity Services when page loads
window.addEventListener('load', function() {
    // Initialize the auth indicator to "not signed in"
    updateAuthIndicator(false);
    
    // Update connection indicators
    updateConnectionIndicators();
    
    // Check if returning from OAuth redirect
    const urlParams = new URLSearchParams(window.location.search);
    const oauthSuccess = urlParams.get('oauth_success');
    const oauthError = urlParams.get('oauth_error');
    const state = urlParams.get('state');
    const error = urlParams.get('error');
    
    if (oauthSuccess === 'true' && state) {
        // Restore state from sessionStorage
        const storedState = sessionStorage.getItem('oauth_state');
        const flowType = sessionStorage.getItem('oauth_flow');
        
        if (storedState === state) {
            currentOAuthState = state;
            updateAuthIndicator(true, 'Successfully authenticated with Google!');
            showStatus('oauth-status', 'Successfully signed in with Google! You can now connect to sheets.', 'success');
            document.getElementById('connect-btn').disabled = false;
            
            // Clean up sessionStorage
            sessionStorage.removeItem('oauth_state');
            sessionStorage.removeItem('oauth_flow');
            
            // Clean up URL
            window.history.replaceState({}, document.title, window.location.pathname);
        }
    } else if (oauthError === 'true') {
        updateAuthIndicator(false, 'Authorization failed - please try again');
        showStatus('oauth-status', `Authorization failed: ${error || 'Unknown error'}`, 'error');
        
        // Clean up sessionStorage
        sessionStorage.removeItem('oauth_state');
        sessionStorage.removeItem('oauth_flow');
        
        // Clean up URL
        window.history.replaceState({}, document.title, window.location.pathname);
    }
    
    // Get client ID from backend
    fetch(`${API_BASE_URL}/sheets/oauth/client-id`)
        .then(response => response.json())
        .then(data => {
            if (data.client_id) {
                googleClientId = data.client_id;
                initializeGoogleSignIn(data.client_id);
            } else {
                const buttonContainer = document.getElementById('g_id_signin');
                if (buttonContainer) {
                    buttonContainer.innerHTML = `
                        <div style="padding: 16px; border: 2px dashed #f4b400; border-radius: 8px; background: #fffde7;">
                            <p style="margin: 0; color: #555;">Google Sign-In is temporarily unavailable. Please try again later.</p>
                        </div>
                    `;
                }
                showStatus('oauth-status', 'Google Sign-In is temporarily unavailable. Please try again later.', 'error');
            }
        })
        .catch(error => {
            console.error('Error fetching client ID:', error);
            const buttonContainer = document.getElementById('g_id_signin');
            if (buttonContainer) {
                buttonContainer.innerHTML = `
                    <div style="padding: 20px; border: 2px dashed #ccc; border-radius: 8px; background: #f9f9f9;">
                        <p style="margin: 0; color: #666;">Could not reach the server. Please try again later.</p>
                    </div>
                `;
            }
            showStatus('oauth-status', 'Could not reach the server. Please try again later.', 'error');
        });
});

// Initialize Google Sign-In
function initializeGoogleSignIn(clientId) {
    // Wait for Google Identity Services to load
    if (typeof google === 'undefined' || !google.accounts || !google.accounts.id) {
        setTimeout(() => initializeGoogleSignIn(clientId), 100);
        return;
    }
    
    try {
        google.accounts.id.initialize({
            client_id: clientId,
            callback: handleGoogleSignIn
        });
        
        google.accounts.id.renderButton(
            document.getElementById('g_id_signin'),
            {
                type: 'standard',
                size: 'large',
                theme: 'outline',
                text: 'sign_in_with',
                shape: 'rectangular',
                logo_alignment: 'left'
            }
        );
        
        // Optional: Show One Tap (can be annoying, so commented out)
        // google.accounts.id.prompt();
    } catch (error) {
        console.error('Error initializing Google Sign-In:', error);
        const buttonContainer = document.getElementById('g_id_signin');
        if (buttonContainer) {
            buttonContainer.innerHTML = `
                <div style="padding: 20px; border: 2px dashed #ccc; border-radius: 8px; background: #f9f9f9;">
                    <p style="margin: 0; color: #666;">Error initializing Google Sign-In. Please refresh the page.</p>
                </div>
            `;
        }
    }
}

// Handle Google Sign-In callback
function handleGoogleSignIn(response) {
    if (response.credential) {
        // This is a JWT token from Google
        googleAccessToken = response.credential;
        
        // Exchange JWT for access token via backend
        exchangeTokenForAccess(response.credential);
    }
}

// Exchange JWT credential for access token
async function exchangeTokenForAccess(credential) {
    updateAuthIndicator(false, 'Authenticating with Google...');
    showStatus('oauth-status', 'Authenticating with Google...', 'info');
    
    try {
        // Send current page URL so callback knows where to redirect
        // Normalize the URL - if pathname is just '/', use '/index.html'
        let pathname = window.location.pathname;
        if (pathname === '/' || pathname === '') {
            pathname = '/index.html';
        }
        const frontendUrl = window.location.origin + pathname;
        console.log('Sending frontend URL to backend:', frontendUrl);

        const result = await apiCall('/sheets/oauth/exchange-token', 'POST', {
            credential: credential,
            frontend_url: frontendUrl
        });
        
        if (result.success) {
            // Store the state before redirecting
            sessionStorage.setItem('oauth_state', result.state);
            sessionStorage.setItem('oauth_flow', 'google_signin');
            
            // Redirect to authorization URL (no popup needed!)
            updateAuthIndicator(false, 'Redirecting to Google for authorization...');
            showStatus('oauth-status', 'Redirecting to Google for authorization...', 'info');
            window.location.href = result.authorization_url;
        } else {
            updateAuthIndicator(false, 'Authentication failed - please try again');
            showStatus('oauth-status', `Error: ${result.error}`, 'error');
        }
    } catch (error) {
        updateAuthIndicator(false, 'Authentication failed - please try again');
        showStatus('oauth-status', `Error: ${error.message}`, 'error');
    }
}

// Connect to sheets using stored token
async function connectSheetsWithToken() {
    const sheetUrl = document.getElementById('sheet-url').value;
    const sheetName = document.getElementById('sheet-name').value;
    
    if (!sheetUrl || !sheetName) {
        showStatus('sheets-status', 'Please provide both Sheet URL and Worksheet Name', 'error');
        return;
    }
    
    if (!currentOAuthState) {
        showStatus('sheets-status', 'Please sign in with Google first', 'error');
        return;
    }
    
    clearStatus('sheets-status');
    showStatus('sheets-status', 'Connecting to Google Sheets...', 'info');
    
    try {
        const result = await apiCall('/sheets/connect', 'POST', {
            sheet_url: sheetUrl,
            sheet_name: sheetName,
            oauth_state: currentOAuthState
        });
        
        if (result.success) {
            // Store the connected sheet info globally
            connectedSheetUrl = sheetUrl;
            connectedSheetName = sheetName;
            isSheetConnected = true;
            
            // Update all connection indicators
            updateConnectionIndicators();
            
            showStatus('sheets-status', 
                `Connected successfully! Found ${result.columns_count} columns. This sheet will be used for all operations.`, 
                'success'
            );
        } else {
            isSheetConnected = false;
            updateConnectionIndicators();
            showStatus('sheets-status', `Error: ${result.error}`, 'error');
        }
    } catch (error) {
        isSheetConnected = false;
        updateConnectionIndicators();
        showStatus('sheets-status', `Error: ${error.message}`, 'error');
    }
}

// Tab management
function showTab(tabName, clickedButton) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Remove active class from all buttons
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show selected tab
    document.getElementById(tabName).classList.add('active');
    
    // Add active class to clicked button
    if (clickedButton) {
        clickedButton.classList.add('active');
    }
    
    // Update connection indicators whenever switching tabs
    updateConnectionIndicators();
}

// Utility functions
function showStatus(elementId, message, type = 'info') {
    const element = document.getElementById(elementId);
    element.textContent = message;
    element.className = `status-message ${type}`;
    element.style.display = 'block';
}

function clearStatus(elementId) {
    const element = document.getElementById(elementId);
    element.style.display = 'none';
}

function showResults(elementId, results, formatFunction) {
    const element = document.getElementById(elementId);
    element.innerHTML = formatFunction(results);
}

// API call wrapper
async function apiCall(endpoint, method = 'GET', data = null) {
    try {
        const options = {
            method,
            headers: {
                'Content-Type': 'application/json',
            }
        };
        
        if (data) {
            options.body = JSON.stringify(data);
        }
        
        const response = await fetch(`${API_BASE_URL}${endpoint}`, options);
        const result = await response.json();
        
        if (!response.ok) {
            throw new Error(result.error || 'Request failed');
        }
        
        return result;
    } catch (error) {
        throw error;
    }
}

// Helper function to get service account data
async function getServiceAccountData() {
    const fileInput = document.getElementById('service-account-file');
    const jsonInput = document.getElementById('service-account-json');
    
    // Check if elements exist (they're only on the connection tab)
    if (!fileInput || !jsonInput) {
        return {};
    }
    
    const jsonInputValue = jsonInput.value.trim();
    
    // If JSON text is provided, use it
    if (jsonInputValue) {
        try {
            const jsonData = JSON.parse(jsonInputValue);
            return { service_account_json: jsonData };
        } catch (e) {
            throw new Error('Invalid JSON in service account field');
        }
    }
    
    // If file is provided, read it
    if (fileInput.files && fileInput.files.length > 0) {
        return new Promise((resolve, reject) => {
            const file = fileInput.files[0];
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const jsonData = JSON.parse(e.target.result);
                    resolve({ service_account_json: jsonData });
                } catch (err) {
                    reject(new Error('Invalid JSON file'));
                }
            };
            reader.onerror = () => reject(new Error('Error reading file'));
            reader.readAsText(file);
        });
    }
    
    return {};
}

// OAuth state storage
let currentOAuthState = null;

// Helper function to get current OAuth state
function getOAuthState() {
    return currentOAuthState;
}

// Update Google Sign-In status indicator
function updateAuthIndicator(isSignedIn, details = '') {
    const indicator = document.getElementById('google-auth-indicator');
    const icon = document.getElementById('auth-status-icon');
    const text = document.getElementById('auth-status-text');
    const detailsElement = document.getElementById('auth-status-details');
    
    if (!indicator || !icon || !text || !detailsElement) return;
    
    if (isSignedIn) {
        // Signed in - green indicator
        indicator.style.borderLeftColor = '#48bb78';
        indicator.style.background = '#f0fff4';
        icon.textContent = '🟢';
        text.textContent = 'Signed in with Google';
        detailsElement.textContent = details || 'You can now connect to Google Sheets';
        document.getElementById('connect-btn').disabled = false;
    } else {
        // Not signed in - gray indicator
        indicator.style.borderLeftColor = '#cbd5e0';
        indicator.style.background = '#f5f5f5';
        icon.textContent = '⚪';
        text.textContent = 'Not signed in';
        detailsElement.textContent = details || 'Click the button below to sign in with Google';
        document.getElementById('connect-btn').disabled = true;
    }
}

// Listen for OAuth callback messages
window.addEventListener('message', function(event) {
    // Accept messages from any origin (for local development)
    if (event.data.type === 'oauth_success') {
        currentOAuthState = event.data.state;
        showStatus('oauth-status', 'Successfully signed in with Google! You can now connect to sheets.', 'success');
        document.getElementById('connect-btn').disabled = false;
    } else if (event.data.type === 'oauth_error') {
        showStatus('oauth-status', `Authorization failed: ${event.data.error}`, 'error');
        currentOAuthState = null;
    }
});

// Connect with Google OAuth
async function connectWithGoogle() {
    const clientId = document.getElementById('oauth-client-id').value.trim();
    const clientSecret = document.getElementById('oauth-client-secret').value.trim();
    
    clearStatus('oauth-status');
    showStatus('oauth-status', 'Starting Google authorization...', 'info');
    
    try {
        const result = await apiCall('/sheets/oauth/authorize', 'POST', {
            client_id: clientId || undefined,
            client_secret: clientSecret || undefined
        });
        
        if (result.success) {
            // Open OAuth URL in popup
            const width = 600;
            const height = 700;
            const left = (screen.width - width) / 2;
            const top = (screen.height - height) / 2;
            
            const popup = window.open(
                result.authorization_url,
                'Google Authorization',
                `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
            );
            
            currentOAuthState = result.state;
            showStatus('oauth-status', 'Please complete authorization in the popup window...', 'info');
            
            // Check if popup was closed
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    if (!currentOAuthState) {
                        showStatus('oauth-status', 'Authorization was cancelled or failed.', 'error');
                    }
                }
            }, 500);
        } else {
            showStatus('oauth-status', `Error: ${result.error}`, 'error');
        }
    } catch (error) {
        showStatus('oauth-status', `Error: ${error.message}`, 'error');
    }
}

// Connect to sheets using OAuth
async function connectSheetsWithOAuth() {
    const sheetUrl = document.getElementById('sheet-url').value;
    const sheetName = document.getElementById('sheet-name').value;
    
    if (!sheetUrl || !sheetName) {
        showStatus('oauth-status', 'Please enter Sheet URL and Worksheet Name first', 'error');
        return;
    }
    
    if (!currentOAuthState) {
        showStatus('oauth-status', 'Please authorize with Google first', 'error');
        return;
    }
    
    try {
        const result = await apiCall('/sheets/connect', 'POST', {
            sheet_url: sheetUrl,
            sheet_name: sheetName,
            oauth_state: currentOAuthState
        });
        
        if (result.success) {
            showStatus('oauth-status', 
                `Connected successfully! Found ${result.columns_count} columns.`, 
                'success'
            );
        } else {
            showStatus('oauth-status', `Error: ${result.error}`, 'error');
        }
    } catch (error) {
        showStatus('oauth-status', `Error: ${error.message}`, 'error');
    }
}

// Connection functions
async function connectSheets() {
    const sheetUrl = document.getElementById('sheet-url').value;
    const sheetName = document.getElementById('sheet-name').value;
    
    if (!sheetUrl || !sheetName) {
        showStatus('sheets-status', 'Please provide both Sheet URL and Worksheet Name', 'error');
        return;
    }
    
    clearStatus('sheets-status');
    showStatus('sheets-status', 'Connecting to Google Sheets...', 'info');
    
    try {
        const serviceAccountData = await getServiceAccountData();
        const result = await apiCall('/sheets/connect', 'POST', {
            sheet_url: sheetUrl,
            sheet_name: sheetName,
            ...serviceAccountData
        });
        
        showStatus('sheets-status', 
            `Connected successfully! Found ${result.columns_count} columns.`, 
            'success'
        );
    } catch (error) {
        showStatus('sheets-status', `Error: ${error.message}`, 'error');
    }
}

async function login() {
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    
    if (!username || !password) {
        showStatus('login-status', 'Please enter both username and password', 'error');
        return;
    }

    clearStatus('login-status');
    hideDuoCode();
    hideDuoPhonePicker();
    showStatus('login-status', 'Logging in... Please approve Duo push on your phone.', 'info');

    // Establish the Flask session cookie BEFORE starting the long-running /login.
    // Otherwise, the polling GETs below race the /login response and each lands
    // on its own brand-new session — so they never see the duo_push_options /
    // duo_verification_code that the /login request is setting on its session.
    try {
        await apiCall('/session/init', 'GET');
    } catch (_) {
        // Non-fatal: if this fails the login will still try, just without polling.
    }

    // While /login is blocking on the backend (waiting for Duo), poll for:
    //   1) the 6-digit "verified push" code (show big and bold when seen)
    //   2) the list of Duo Push options when the user has multiple phones
    //      (show a picker so they can choose by last 4 digits)
    let pollDone = false;
    let codeShown = false;
    let pickerShown = false;
    const pollLoop = async () => {
        while (!pollDone) {
            try {
                if (!codeShown) {
                    const r = await apiCall('/automation/duo-verification-code', 'GET');
                    if (r && r.code) {
                        showDuoCode(r.code);
                        codeShown = true;
                    }
                }
                if (!pickerShown) {
                    const o = await apiCall('/automation/duo-push-options', 'GET');
                    if (o && Array.isArray(o.options) && o.options.length > 0) {
                        showDuoPhonePicker(o.options);
                        pickerShown = true;
                    }
                }
            } catch (_) {
                // Endpoint may briefly fail during transitions — keep polling
            }
            await new Promise(res => setTimeout(res, 1500));
        }
    };
    pollLoop();

    try {
        const result = await apiCall('/automation/login', 'POST', {
            username,
            password
        });

        showStatus('login-status', result.message || 'Login successful!', 'success');
    } catch (error) {
        showStatus('login-status', `Error: ${error.message}`, 'error');
    } finally {
        pollDone = true;
        hideDuoCode();
        hideDuoPhonePicker();
    }
}

function showDuoPhonePicker(options) {
    let el = document.getElementById('duo-phone-picker');
    if (!el) {
        el = document.createElement('div');
        el.id = 'duo-phone-picker';
        el.style.cssText = 'margin:12px 0;padding:16px;border-radius:8px;background:#e3f2fd;border:2px solid #2196f3;';
        const statusEl = document.getElementById('login-status');
        if (statusEl && statusEl.parentNode) {
            statusEl.parentNode.insertBefore(el, statusEl);
        }
    }
    const list = options.map(o =>
        `<li style="margin:4px 0;color:#0d47a1;">${(o.label || '').replace(/</g, '&lt;')}</li>`
    ).join('');
    el.innerHTML = `
        <div style="font-weight:bold;color:#0d47a1;margin-bottom:8px;">
            You have multiple Duo Push devices. Pick one:
        </div>
        <ul style="margin:0 0 12px 18px;padding:0;font-family:monospace;font-size:0.95em;">${list}</ul>
        <div style="display:flex;gap:8px;align-items:center;">
            <input id="duo-phone-last4" type="text" inputmode="numeric" pattern="[0-9]{4}"
                   maxlength="4" placeholder="Last 4 digits"
                   style="padding:8px;border:1px solid #2196f3;border-radius:4px;font-family:monospace;font-size:1.1em;width:120px;">
            <button onclick="submitDuoPhoneChoice()" class="btn btn-primary"
                    style="padding:8px 16px;">Use this phone</button>
            <span id="duo-phone-picker-err" style="color:#c62828;font-size:0.9em;"></span>
        </div>
    `;
    el.style.display = 'block';
    setTimeout(() => {
        const input = document.getElementById('duo-phone-last4');
        if (input) input.focus();
    }, 50);
}

function hideDuoPhonePicker() {
    const el = document.getElementById('duo-phone-picker');
    if (el) el.style.display = 'none';
}

async function submitDuoPhoneChoice() {
    const input = document.getElementById('duo-phone-last4');
    const errEl = document.getElementById('duo-phone-picker-err');
    const last4 = (input?.value || '').trim();
    if (!/^\d{4}$/.test(last4)) {
        if (errEl) errEl.textContent = 'Enter exactly 4 digits.';
        return;
    }
    if (errEl) errEl.textContent = '';
    try {
        const r = await apiCall('/automation/duo-push-select', 'POST', { last4 });
        if (r && r.success) {
            hideDuoPhonePicker();
            showStatus('login-status', `Sending Duo Push to phone ending in ${last4}...`, 'info');
        } else {
            if (errEl) errEl.textContent = r?.error || 'Selection failed';
        }
    } catch (e) {
        if (errEl) errEl.textContent = e.message;
    }
}

function showDuoCode(code) {
    let el = document.getElementById('duo-verification-banner');
    if (!el) {
        el = document.createElement('div');
        el.id = 'duo-verification-banner';
        el.style.cssText = 'margin:12px 0;padding:16px;border-radius:8px;background:#fff3cd;border:2px solid #ffc107;text-align:center;';
        const statusEl = document.getElementById('login-status');
        if (statusEl && statusEl.parentNode) {
            statusEl.parentNode.insertBefore(el, statusEl);
        }
    }
    el.innerHTML = `
        <div style="font-size:0.95em;color:#856404;margin-bottom:6px;">
            Enter this code in your <strong>Duo Mobile</strong> app to approve:
        </div>
        <div style="font-size:2.5em;font-weight:bold;letter-spacing:0.15em;color:#000;font-family:monospace;">
            ${code}
        </div>
    `;
    el.style.display = 'block';
}

function hideDuoCode() {
    const el = document.getElementById('duo-verification-banner');
    if (el) el.style.display = 'none';
}

// Add privileges
async function addPrivileges() {
    const sheetInfo = requireSheetConnection('add-status');
    if (!sheetInfo) return;

    const appName = document.getElementById('add-app-name').value;
    const comment = document.getElementById('add-comment').value.trim();
    const performedByName = document.getElementById('add-performed-by-name').value;
    const columnIndex = parseInt(document.getElementById('add-read-column').value);

    if (!performedByName || performedByName.trim() === '') {
        showStatus('add-status', 'Please provide a Performed By Name', 'error');
        return;
    }

    clearStatus('add-status');
    clearResults('add-results');
    showStatus('add-status', 'Reading IDs from sheet...', 'info');
    setOperationRunning('add', true);

    try {
        const oauthState = getOAuthState();
        const ids = await fetchSheetColumn(sheetInfo.url, sheetInfo.name, columnIndex, oauthState);
        if (ids.length === 0) {
            showStatus('add-status', 'No IDs found in the selected column', 'error');
            return;
        }

        const summary = await runChunked(ids, CHUNK_SIZE, async (chunk) => {
            return await apiCall('/automation/add', 'POST', {
                sheet_url: sheetInfo.url,
                sheet_name: sheetInfo.name,
                app_name: appName,
                comment: comment,
                performed_by_name: performedByName.trim(),
                column_index: columnIndex,
                ids: chunk,
                oauth_state: oauthState
            });
        }, { statusElementId: 'add-status', label: 'IDs' });

        if (!summary.error) {
            showStatus('add-status',
                `Completed: ${summary.successful} successful out of ${summary.total}`,
                summary.successful === summary.total ? 'success' : 'info'
            );
        }
        showResults('add-results', summary.results, formatAddResults);
    } catch (error) {
        showStatus('add-status', `Error: ${error.message}`, 'error');
    } finally {
        setOperationRunning('add', false);
    }
}

function formatAddResults(results) {
    if (!results || results.length === 0) return '<p>No results</p>';
    
    let html = '<table><thead><tr><th>ID</th><th>Status</th><th>Error</th></tr></thead><tbody>';
    
    results.forEach(result => {
        const statusClass = result.success ? 'success' : 'error';
        const statusText = result.success ? 'Success' : 'Failed';
        html += `
            <tr class="result-item ${statusClass}">
                <td>${result.id}</td>
                <td>${statusText}</td>
                <td>${result.error || '-'}</td>
            </tr>
        `;
    });
    
    html += '</tbody></table>';
    return html;
}

// Revoke privileges
async function revokePrivileges() {
    const sheetInfo = requireSheetConnection('revoke-status');
    if (!sheetInfo) return;

    const appName = document.getElementById('revoke-app-name').value;
    const comment = document.getElementById('revoke-comment').value.trim();
    const columnIndex = parseInt(document.getElementById('revoke-read-column').value);

    clearStatus('revoke-status');
    clearResults('revoke-results');
    showStatus('revoke-status', 'Reading IDs from sheet...', 'info');
    setOperationRunning('revoke', true);

    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();

        const ids = await fetchSheetColumn(sheetInfo.url, sheetInfo.name, columnIndex, oauthState);
        if (ids.length === 0) {
            showStatus('revoke-status', 'No IDs found in the selected column', 'error');
            return;
        }

        const summary = await runChunked(ids, CHUNK_SIZE, async (chunk) => {
            const requestData = {
                sheet_url: sheetInfo.url,
                sheet_name: sheetInfo.name,
                app_name: appName,
                comment: comment,
                column_index: columnIndex,
                ids: chunk
            };
            if (oauthState) requestData.oauth_state = oauthState;
            else Object.assign(requestData, serviceAccountData);
            return await apiCall('/automation/revoke', 'POST', requestData);
        }, { statusElementId: 'revoke-status', label: 'IDs' });

        if (!summary.error) {
            showStatus('revoke-status',
                `Completed: ${summary.successful} successful out of ${summary.total}`,
                summary.successful === summary.total ? 'success' : 'info'
            );
        }
        showResults('revoke-results', summary.results, formatRevokeResults);
    } catch (error) {
        showStatus('revoke-status', `Error: ${error.message}`, 'error');
    } finally {
        setOperationRunning('revoke', false);
    }
}

function formatRevokeResults(results) {
    if (!results || results.length === 0) return '<p>No results</p>';
    let html = '<table><thead><tr><th>ID</th><th>Status</th><th>Error</th></tr></thead><tbody>';
    results.forEach(result => {
        const isAlreadySet = result.error && result.error.startsWith('Privilege already ');
        const statusClass = result.success || isAlreadySet ? 'success' : 'error';
        const statusText = result.success ? 'Success' : (isAlreadySet ? 'Success' : 'Failed');
        html += `
            <tr class="result-item ${statusClass}">
                <td>${result.id}</td>
                <td>${statusText}</td>
                <td>${result.error || '-'}</td>
            </tr>
        `;
    });
    html += '</tbody></table>';
    return html;
}

// Get employment status
async function getEmploymentStatus() {
    const sheetInfo = requireSheetConnection('status-status');
    if (!sheetInfo) return;
    setOperationRunning('status', true);

    const idsText = document.getElementById('status-ids').value;
    const idType = document.getElementById('status-id-type').value;
    const columnIndex = parseInt(document.getElementById('status-read-column').value);
    const toFields = [];
    const writeColumns = {};
    document.querySelectorAll('#status-writes-inline .convert-write-chip').forEach(chip => {
        const type = chip.dataset.type;
        const sel = chip.querySelector('select.convert-chip-col');
        if (type && sel) {
            toFields.push(type);
            writeColumns[type] = sel.value;
        }
    });
    if (toFields.length === 0) {
        showStatus('status-status', 'Select at least one field to get (click "Add field to get").', 'error');
        setOperationRunning('status', false);
        return;
    }

    clearStatus('status-status');
    clearResults('status-results');
    showStatus('status-status', 'Reading IDs...', 'info');
    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();

        // Pull IDs either from the textarea or from the selected sheet column.
        let ids = idsText.trim() ? idsText.split('\n').map(id => id.trim()).filter(id => id) : [];
        if (ids.length === 0) {
            ids = await fetchSheetColumn(sheetInfo.url, sheetInfo.name, columnIndex, oauthState);
        }
        if (ids.length === 0) {
            showStatus('status-status', 'No IDs found (textarea empty and column has no values)', 'error');
            return;
        }

        const summary = await runChunked(ids, CHUNK_SIZE, async (chunk, chunkIndex) => {
            const requestData = {
                sheet_url: sheetInfo.url,
                sheet_name: sheetInfo.name,
                ids: chunk,
                id_type: idType,
                column_index: columnIndex,
                to_fields: toFields,
                write_columns: writeColumns,
                row_offset: chunkIndex * CHUNK_SIZE,
                write_headers: chunkIndex === 0
            };
            if (oauthState) requestData.oauth_state = oauthState;
            else Object.assign(requestData, serviceAccountData);
            return await apiCall('/automation/get-employment-status', 'POST', requestData);
        }, { statusElementId: 'status-status', label: 'IDs' });

        if (!summary.error) {
            const colList = toFields.map(f => `${STATUS_FIELD_LABELS[f] || f}→${writeColumns[f]}`).join(', ');
            showStatus('status-status', `Done (${summary.results.length}/${summary.total}). Written to: ${colList}`, 'success');
        }
        showResults('status-results', summary.results, (r) => formatStatusResults(r, toFields));
    } catch (error) {
        showStatus('status-status', `Error: ${error.message}`, 'error');
    } finally {
        setOperationRunning('status', false);
    }
}

function formatStatusResults(results, toFields) {
    if (!results || results.length === 0) return '<p>No results</p>';
    const keys = { SOURCE_SYSTEM: 'source', EMPLOYMENT_STATUS: 'employment_status', STUDENT_STATUS_CODE: 'student_status_code' };
    let html = '<table><thead><tr><th>ID</th>';
    (toFields || []).forEach(f => { html += `<th>${STATUS_FIELD_LABELS[f] || f}</th>`; });
    html += '<th>Status</th></tr></thead><tbody>';
    results.forEach(result => {
        const statusClass = result.success ? 'success' : 'error';
        html += '<tr class="result-item ' + statusClass + '"><td>' + result.id + '</td>';
        (toFields || []).forEach(f => {
            const key = keys[f];
            html += '<td>' + (result[key] != null ? result[key] : '-') + '</td>';
        });
        html += '<td>' + (result.success ? 'Success' : result.error) + '</td></tr>';
    });
    html += '</tbody></table>';
    return html;
}

// Statuses — multi-select (Source System, Employment Status, Student Status Code)
const STATUS_FIELD_LABELS = { SOURCE_SYSTEM: 'Source System', EMPLOYMENT_STATUS: 'Employment Status', STUDENT_STATUS_CODE: 'Student Status Code' };
const STATUS_DEFAULT_COLUMNS = ['G', 'H', 'I'];
let statusToSelections = [];

function getNextDefaultStatusColumn() {
    const used = new Set(statusToSelections.map(s => s.column));
    for (const c of STATUS_DEFAULT_COLUMNS) {
        if (!used.has(c)) return c;
    }
    for (const c of COLUMN_LETTERS) {
        if (!used.has(c)) return c;
    }
    return 'Z';
}

function renderStatusWritesList() {
    const el = document.getElementById('status-writes-inline');
    const prefix = document.getElementById('status-writes-prefix');
    const sep = document.getElementById('status-writes-sep');
    if (!el) return;
    if (prefix) prefix.textContent = statusToSelections.length > 0 ? 'and writes ' : '';
    if (sep) sep.textContent = statusToSelections.length > 0 ? ', ' : '';
    if (statusToSelections.length === 0) {
        el.innerHTML = '';
        return;
    }
    const parts = statusToSelections.map((s, i) => {
        const and = i > 0 ? '<span class="convert-writes-and">and</span>' : '';
        return and + `<span class="convert-write-chip" data-type="${s.type}"><strong>${s.label}</strong> in column <select class="convert-chip-col status-chip-col" data-type="${s.type}">${getConvertColumnOptionsHtml(s.column)}</select><button type="button" class="convert-chip-remove" aria-label="Remove">×</button></span>`;
    });
    el.innerHTML = parts.join('');
    el.querySelectorAll('.status-chip-col').forEach(sel => {
        sel.addEventListener('change', function() {
            const s = statusToSelections.find(x => x.type === this.dataset.type);
            if (s) s.column = this.value;
        });
    });
    el.querySelectorAll('.convert-chip-remove').forEach(btn => {
        btn.addEventListener('click', function() {
            const chip = this.closest('.convert-write-chip');
            if (chip) {
                const t = chip.dataset.type;
                statusToSelections = statusToSelections.filter(x => x.type !== t);
                renderStatusWritesList();
            }
        });
    });
}

function initStatusTypeDropdown() {
    // Only close on outside click (same as Convert). Opening is handled by onclick in index.html head (toggleStatusAddType).
    document.body.addEventListener('click', function(e) {
        const addBtn = document.getElementById('status-add-type-btn');
        const dropdown = document.getElementById('status-type-dropdown');
        const wrapper = document.getElementById('status-type-wrapper');
        if (!addBtn || !dropdown || !wrapper) return;
        if (e.target.id === 'status-add-type-btn' || e.target.closest('#status-add-type-btn')) return;
        const option = e.target.closest('#status-type-dropdown button[role="option"]');
        if (option && dropdown.classList.contains('open')) {
            e.preventDefault();
            e.stopPropagation();
            const type = option.dataset.type;
            if (type && typeof statusToSelections !== 'undefined' && !statusToSelections.some(s => s.type === type)) {
                statusToSelections.push({
                    type,
                    label: STATUS_FIELD_LABELS[type] || type,
                    column: getNextDefaultStatusColumn()
                });
                if (typeof renderStatusWritesList === 'function') renderStatusWritesList();
            }
            dropdown.classList.remove('open');
            dropdown.style.display = 'none';
            addBtn.setAttribute('aria-expanded', 'false');
            dropdown.setAttribute('aria-hidden', 'true');
            return;
        }
        if (dropdown.classList.contains('open') && !wrapper.contains(e.target)) {
            dropdown.classList.remove('open');
            dropdown.style.display = 'none';
            addBtn.setAttribute('aria-expanded', 'false');
            dropdown.setAttribute('aria-hidden', 'true');
        }
    });
}

// Convert IDs — multi-select UI (sentence + dropdown)
const CONVERT_TYPE_LABELS = { SID: 'Short ID (SID)', NETID: 'Net ID', BID: 'Brown ID', BROWN_EMAIL: 'Brown Email' };
const CONVERT_DEFAULT_COLUMNS = ['B', 'C', 'D', 'E'];
const COLUMN_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
let convertToSelections = [];

function getConvertColumnOptionsHtml(selected) {
    return COLUMN_LETTERS.map(c => `<option value="${c}"${c === selected ? ' selected' : ''}>${c}</option>`).join('');
}

function getNextDefaultConvertColumn() {
    const used = new Set(convertToSelections.map(s => s.column));
    for (const c of CONVERT_DEFAULT_COLUMNS) {
        if (!used.has(c)) return c;
    }
    for (const c of COLUMN_LETTERS) {
        if (!used.has(c)) return c;
    }
    return 'Z';
}

function renderConvertWritesList() {
    const el = document.getElementById('convert-writes-inline');
    const prefix = document.getElementById('convert-writes-prefix');
    const sep = document.getElementById('convert-writes-sep');
    if (!el) return;
    if (prefix) prefix.textContent = convertToSelections.length > 0 ? 'and writes ' : '';
    if (sep) sep.textContent = convertToSelections.length > 0 ? ', ' : '';
    if (convertToSelections.length === 0) {
        el.innerHTML = '';
        return;
    }
    const parts = convertToSelections.map((s, i) => {
        const and = i > 0 ? '<span class="convert-writes-and">and</span>' : '';
        return and + `<span class="convert-write-chip" data-type="${s.type}"><strong>${s.label}</strong> in column <select class="convert-chip-col" data-type="${s.type}">${getConvertColumnOptionsHtml(s.column)}</select><button type="button" class="convert-chip-remove" aria-label="Remove">×</button></span>`;
    });
    el.innerHTML = parts.join('');
    el.querySelectorAll('.convert-chip-col').forEach(sel => {
        sel.addEventListener('change', function() {
            const s = convertToSelections.find(x => x.type === this.dataset.type);
            if (s) s.column = this.value;
        });
    });
    el.querySelectorAll('.convert-chip-remove').forEach(btn => {
        btn.addEventListener('click', function() {
            const chip = this.closest('.convert-write-chip');
            if (chip) {
                const t = chip.dataset.type;
                convertToSelections = convertToSelections.filter(x => x.type !== t);
                renderConvertWritesList();
            }
        });
    });
}

function initConvertTypeDropdown() {
    // Use event delegation so it works regardless of when the convert tab is in the DOM
    document.body.addEventListener('click', function(e) {
        const addBtn = document.getElementById('convert-add-type-btn');
        const dropdown = document.getElementById('convert-type-dropdown');
        if (!addBtn || !dropdown) return;
        // Click on "Add type to convert to" button -> toggle dropdown
        if (e.target.id === 'convert-add-type-btn' || e.target.closest('#convert-add-type-btn')) {
            e.preventDefault();
            e.stopPropagation();
            const open = dropdown.classList.toggle('open');
            addBtn.setAttribute('aria-expanded', open);
            dropdown.setAttribute('aria-hidden', !open);
            return;
        }
        // Click on an option inside the dropdown -> add type
        const option = e.target.closest('#convert-type-dropdown button[role="option"]');
        if (option && dropdown.classList.contains('open')) {
            e.preventDefault();
            e.stopPropagation();
            const type = option.dataset.type;
            if (type && !convertToSelections.some(s => s.type === type)) {
                convertToSelections.push({
                    type,
                    label: CONVERT_TYPE_LABELS[type] || type,
                    column: getNextDefaultConvertColumn()
                });
                renderConvertWritesList();
            }
            dropdown.classList.remove('open');
            addBtn.setAttribute('aria-expanded', 'false');
            dropdown.setAttribute('aria-hidden', 'true');
            return;
        }
        // Click outside -> close dropdown
        if (dropdown.classList.contains('open') && !dropdown.contains(e.target) && e.target !== addBtn) {
            dropdown.classList.remove('open');
            addBtn.setAttribute('aria-expanded', 'false');
            dropdown.setAttribute('aria-hidden', 'true');
        }
    });
}
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() {
        initConvertTypeDropdown();
        initStatusTypeDropdown();
        initMoveConditionBuilder();
    });
} else {
    initConvertTypeDropdown();
    initStatusTypeDropdown();
    initMoveConditionBuilder();
}

// Convert IDs
async function convertIds() {
    const sheetInfo = requireSheetConnection('convert-status');
    if (!sheetInfo) return;

    const idsText = document.getElementById('convert-ids').value;
    const fromType = document.getElementById('convert-from-type').value;
    const columnIndex = parseInt(document.getElementById('convert-read-column').value);
    const toTypes = [];
    const writeColumns = {};
    document.querySelectorAll('#convert-writes-inline .convert-write-chip').forEach(chip => {
        const type = chip.dataset.type;
        const sel = chip.querySelector('select.convert-chip-col');
        if (type && sel) {
            toTypes.push(type);
            writeColumns[type] = sel.value;
        }
    });
    if (toTypes.length === 0) {
        showStatus('convert-status', 'Select at least one type to convert to (click "Add type to convert to").', 'error');
        return;
    }

    clearStatus('convert-status');
    clearResults('convert-results');
    showStatus('convert-status', 'Reading IDs...', 'info');
    setOperationRunning('convert', true);

    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();

        let ids = idsText.trim() ? idsText.split('\n').map(id => id.trim()).filter(id => id) : [];
        if (ids.length === 0) {
            ids = await fetchSheetColumn(sheetInfo.url, sheetInfo.name, columnIndex, oauthState);
        }
        if (ids.length === 0) {
            showStatus('convert-status', 'No IDs found', 'error');
            return;
        }

        const summary = await runChunked(ids, CHUNK_SIZE, async (chunk, chunkIndex) => {
            const requestData = {
                sheet_url: sheetInfo.url,
                sheet_name: sheetInfo.name,
                ids: chunk,
                from_type: fromType,
                to_types: toTypes,
                write_columns: writeColumns,
                column_index: columnIndex,
                row_offset: chunkIndex * CHUNK_SIZE,
                write_headers: chunkIndex === 0
            };
            if (oauthState) requestData.oauth_state = oauthState;
            else Object.assign(requestData, serviceAccountData);
            return await apiCall('/automation/convert-id', 'POST', requestData);
        }, { statusElementId: 'convert-status', label: 'IDs' });

        if (!summary.error) {
            const colList = toTypes.map(t => `${t}→${writeColumns[t]}`).join(', ');
            showStatus('convert-status', `Done (${summary.results.length}/${summary.total}). Written to: ${colList}`, 'success');
        }
        showResults('convert-results', summary.results, (r) => formatConvertResults(r, toTypes));
    } catch (error) {
        showStatus('convert-status', `Error: ${error.message}`, 'error');
    } finally {
        setOperationRunning('convert', false);
    }
}

async function runConvertValidation() {
    const sheetInfo = requireSheetConnection('convert-status');
    if (!sheetInfo) return;

    const checkColumn = (document.getElementById('convert-validate-check-column')?.value || 'X').trim().toUpperCase();
    const dataStartRow = parseInt(document.getElementById('convert-validate-data-start-row')?.value || '2', 10);
    const mappingInputs = [
        { field: 'SID', elementId: 'convert-map-sid-column' },
        { field: 'NETID', elementId: 'convert-map-netid-column' },
        { field: 'BID', elementId: 'convert-map-bid-column' },
        { field: 'FIRST_NAME', elementId: 'convert-map-firstname-column' },
        { field: 'LAST_NAME', elementId: 'convert-map-lastname-column' }
    ];

    const searchMappings = mappingInputs
        .map(item => {
            const val = (document.getElementById(item.elementId)?.value || '').trim().toUpperCase();
            return val ? { search_field: item.field, column: val } : null;
        })
        .filter(Boolean);

    if (searchMappings.length === 0) {
        showStatus('convert-status', 'Select at least one search mapping column.', 'error');
        return;
    }

    if (!Number.isInteger(dataStartRow) || dataStartRow < 1) {
        showStatus('convert-status', 'Data Start Row must be an integer >= 1.', 'error');
        return;
    }

    const validateBtn = document.getElementById('convert-validate-start-btn');
    if (validateBtn) validateBtn.disabled = true;

    clearStatus('convert-status');
    clearResults('convert-results');
    showStatus('convert-status', 'Running validation check...', 'info');
    window._chunkAbortRequested = false;

    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();

        // Validation chunks by row range. We don't know the total up front, so we
        // keep advancing the window until a chunk returns fewer rows than requested.
        const allResults = [];
        let totals = { green_source: 0, green_check: 0, red_source: 0, matched: 0 };
        let chunkIndex = 0;
        let nextStart = dataStartRow;

        while (true) {
            if (window._chunkAbortRequested) {
                showStatus('convert-status',
                    `Aborted after chunk ${chunkIndex} (${allResults.length} rows processed).`,
                    'info');
                break;
            }
            const endRow = nextStart + CHUNK_SIZE - 1;

            showStatus('convert-status',
                `Processing chunk ${chunkIndex + 1} (rows ${nextStart}-${endRow}, ${allResults.length} done)...`,
                'info');

            const requestData = {
                sheet_url: sheetInfo.url,
                sheet_name: sheetInfo.name,
                search_mappings: searchMappings,
                check_column: checkColumn,
                data_start_row: nextStart,
                data_end_row: endRow
            };
            if (oauthState) requestData.oauth_state = oauthState;
            else Object.assign(requestData, serviceAccountData);

            let result;
            try {
                result = await apiCall('/automation/convert-validation', 'POST', requestData);
            } catch (e) {
                showStatus('convert-status',
                    `Failed at chunk ${chunkIndex + 1} (${allResults.length} rows done before failure): ${e.message}`,
                    'error');
                throw e;
            }

            const chunkResults = (result && result.results) || [];
            allResults.push(...chunkResults);
            totals.green_source += result?.green_source_cells || 0;
            totals.green_check += result?.green_check_cells || 0;
            totals.red_source += result?.red_source_cells || 0;
            totals.matched += result?.matched || 0;

            chunkIndex++;
            // If this chunk returned fewer rows than its window allowed, we're done.
            if (chunkResults.length < CHUNK_SIZE) break;
            nextStart = endRow + 1;
        }

        showStatus('convert-status',
            `Validation complete: ${totals.matched} matched, ${allResults.length - totals.matched} unmatched. ` +
            `Colored ${totals.green_source} source + ${totals.green_check} check cells green, ${totals.red_source} source cells red.`,
            'success');
        showResults('convert-results', allResults, formatConvertValidationResults);
    } catch (error) {
        showStatus('convert-status', `Error: ${error.message}`, 'error');
    } finally {
        if (validateBtn) validateBtn.disabled = false;
    }
}

function formatConvertResults(results, toTypes) {
    if (!results || results.length === 0) return '<p>No results</p>';
    const types = Array.isArray(toTypes) && toTypes.length ? toTypes : ['SID'];
    const headers = ['Original ID', ...types.map(t => t.replace('_', ' ')), 'Status'].join('</th><th>');
    let html = `<table><thead><tr><th>${headers}</th></tr></thead><tbody>`;
    results.forEach(result => {
        const statusClass = result.success ? 'success' : 'error';
        const cells = types.map(t => result.converted_ids && result.converted_ids[t] != null ? result.converted_ids[t] : (result.converted_id != null && types.length === 1 ? result.converted_id : '-'));
        const cellStr = cells.map(c => `<td>${c}</td>`).join('');
        html += `<tr class="result-item ${statusClass}"><td>${result.id}</td>${cellStr}<td>${result.success ? 'Success' : (result.error || 'Error')}</td></tr>`;
    });
    html += '</tbody></table>';
    return html;
}

function formatConvertValidationResults(results) {
    if (!results || results.length === 0) return '<p>No results</p>';
    let html = '<table><thead><tr><th>Search Values</th><th>Source Cells</th><th>Extracted Values</th><th>Matched Cells</th><th>Status</th></tr></thead><tbody>';
    results.forEach(result => {
        const matchedCells = (result.matched_check_cells || []).join(', ');
        const extracted = (result.extracted_values || []).join(', ');
        const searchValuesObj = result.search_values || {};
        const searchValues = Object.entries(searchValuesObj).map(([k, v]) => `${k}: ${v}`).join(', ');
        const sourceCells = (result.source_cells || []).join(', ');
        const matched = Array.isArray(result.matched_check_cells) && result.matched_check_cells.length > 0;
        const statusClass = matched ? 'success' : 'error';
        const statusText = matched ? 'Match found' : (result.error ? result.error : 'No match');
        html += `<tr class="result-item ${statusClass}"><td>${searchValues || '-'}</td><td>${sourceCells || '-'}</td><td>${extracted || '-'}</td><td>${matchedCells || '-'}</td><td>${statusText}</td></tr>`;
    });
    html += '</tbody></table>';
    return html;
}

// Compare lists
async function compareLists() {
    // Get sheet info from global connection
    const sheetInfo = requireSheetConnection('compare-status');
    if (!sheetInfo) {
        return;
    }
    
    // Refresh sheet data before processing
    showStatus('compare-status', 'Refreshing sheet data...', 'info');
    const refreshed = await refreshSheetConnection();
    if (!refreshed) {
        showStatus('compare-status', 'Failed to refresh sheet connection', 'error');
        return;
    }
    
    const list1Text = document.getElementById('list1').value;
    const list2Text = document.getElementById('list2').value;
    const column1Index = parseInt(document.getElementById('compare-list1-column').value);
    const column2Index = parseInt(document.getElementById('compare-list2-column').value);
    const toAddColumn = document.getElementById('compare-toadd-column').value;
    const toRemoveColumn = document.getElementById('compare-toremove-column').value;
    
    const list1 = list1Text.trim() ? list1Text.split('\n').map(item => item.trim()).filter(item => item) : [];
    const list2 = list2Text.trim() ? list2Text.split('\n').map(item => item.trim()).filter(item => item) : [];
    
    clearStatus('compare-status');
    clearResults('compare-results');
    
    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();
        const requestData = {
            sheet_url: sheetInfo.url,
            sheet_name: sheetInfo.name,
            list1: list1.length > 0 ? list1 : undefined,
            list2: list2.length > 0 ? list2 : undefined,
            column1_index: column1Index,
            column2_index: column2Index,
            to_add_column: toAddColumn,
            to_remove_column: toRemoveColumn
        };
        
        if (oauthState) {
            requestData.oauth_state = oauthState;
        } else {
            Object.assign(requestData, serviceAccountData);
        }
        
        const result = await apiCall('/automation/compare-lists', 'POST', requestData);
        
        showStatus('compare-status', 
            `Found ${result.to_add_count} to add, ${result.to_remove_count} to remove. Results written to columns ${toAddColumn} and ${toRemoveColumn}.`, 
            'success'
        );
        showResults('compare-results', result, formatCompareResults);
    } catch (error) {
        showStatus('compare-status', `Error: ${error.message}`, 'error');
    }
}

// Move and compact selected columns
let moveConditions = [];

const MOVE_CONDITION_OPERATORS = [
    { value: 'equals', label: 'equals' },
    { value: 'not_equals', label: 'not equals' },
    { value: 'contains', label: 'contains' },
    { value: 'not_contains', label: 'not contains' },
    { value: 'starts_with', label: 'starts with' },
    { value: 'ends_with', label: 'ends with' },
    { value: 'is_empty', label: 'is empty' },
    { value: 'is_not_empty', label: 'is not empty' }
];

function operatorNeedsValue(operator) {
    return operator !== 'is_empty' && operator !== 'is_not_empty';
}

function initMoveConditionBuilder() {
    if (!document.getElementById('move-conditions-container')) return;
    if (moveConditions.length === 0) {
        moveConditions = [
            { column: 'C', operator: 'equals', value: '' },
            { column: 'D', operator: 'equals', value: '' }
        ];
    }
    renderMoveConditionRows();
}

function addMoveConditionRow() {
    moveConditions.push({ column: '', operator: 'equals', value: '' });
    renderMoveConditionRows();
}

function removeMoveConditionRow(index) {
    if (moveConditions.length <= 1) {
        moveConditions[0] = { column: '', operator: 'equals', value: '' };
    } else {
        moveConditions.splice(index, 1);
    }
    renderMoveConditionRows();
}

function updateMoveConditionRow(index, key, value) {
    if (!moveConditions[index]) return;
    moveConditions[index][key] = value;
    renderMoveConditionRows();
}

function renderMoveConditionRows() {
    const container = document.getElementById('move-conditions-container');
    if (!container) return;

    container.innerHTML = moveConditions.map((cond, index) => {
        const options = MOVE_CONDITION_OPERATORS
            .map(op => `<option value="${op.value}"${op.value === cond.operator ? ' selected' : ''}>${op.label}</option>`)
            .join('');
        const needsValue = operatorNeedsValue(cond.operator);

        return `
            <div style="display: grid; grid-template-columns: 90px 1fr 1fr auto; gap: 8px; margin-bottom: 8px; align-items: center;">
                <input type="text" value="${cond.column || ''}" placeholder="Column"
                    oninput="updateMoveConditionRow(${index}, 'column', this.value.toUpperCase().trim())">
                <select onchange="updateMoveConditionRow(${index}, 'operator', this.value)">${options}</select>
                <input type="text" value="${cond.value || ''}" placeholder="Value"
                    ${needsValue ? '' : 'disabled'}
                    oninput="updateMoveConditionRow(${index}, 'value', this.value)">
                <button type="button" class="btn btn-danger" onclick="removeMoveConditionRow(${index})">×</button>
            </div>
        `;
    }).join('');
}

async function moveAndShiftColumns() {
    const sheetInfo = requireSheetConnection('move-status');
    if (!sheetInfo) {
        return;
    }

    showStatus('move-status', 'Refreshing sheet data...', 'info');
    const refreshed = await refreshSheetConnection();
    if (!refreshed) {
        showStatus('move-status', 'Failed to refresh sheet connection', 'error');
        return;
    }

    const sourceColumnsRaw = document.getElementById('move-source-columns').value || '';
    const destinationColumnsRaw = document.getElementById('move-destination-columns').value || '';
    const destinationStartRow = parseInt(document.getElementById('move-destination-start-row').value, 10);
    const dataStartRow = parseInt(document.getElementById('move-data-start-row').value, 10);

    const sourceColumns = sourceColumnsRaw.split(',').map(c => c.trim().toUpperCase()).filter(Boolean);
    const destinationColumns = destinationColumnsRaw.split(',').map(c => c.trim().toUpperCase()).filter(Boolean);

    if (sourceColumns.length === 0 || destinationColumns.length === 0) {
        showStatus('move-status', 'Please provide source and destination columns.', 'error');
        return;
    }

    if (sourceColumns.length !== destinationColumns.length) {
        showStatus('move-status', 'Source and destination column counts must match.', 'error');
        return;
    }

    if (!Number.isInteger(destinationStartRow) || destinationStartRow < 1 || !Number.isInteger(dataStartRow) || dataStartRow < 1) {
        showStatus('move-status', 'Row numbers must be integers greater than or equal to 1.', 'error');
        return;
    }

    const conditions = moveConditions
        .map(c => ({
            column: (c.column || '').trim().toUpperCase(),
            operator: c.operator || 'equals',
            value: c.value ?? ''
        }))
        .filter(c => c.column);

    if (conditions.length === 0) {
        showStatus('move-status', 'Please provide at least one condition with a column.', 'error');
        return;
    }

    for (const c of conditions) {
        if (operatorNeedsValue(c.operator) && String(c.value).trim() === '') {
            showStatus('move-status', `Condition on column ${c.column} requires a value.`, 'error');
            return;
        }
    }

    clearStatus('move-status');
    clearResults('move-results');
    showStatus('move-status', 'Moving matching values and compacting source columns...', 'info');

    try {
        const serviceAccountData = await getServiceAccountData();
        const oauthState = getOAuthState();
        const requestData = {
            sheet_url: sheetInfo.url,
            sheet_name: sheetInfo.name,
            source_columns: sourceColumns,
            destination_columns: destinationColumns,
            conditions,
            destination_start_row: destinationStartRow,
            data_start_row: dataStartRow
        };

        if (oauthState) {
            requestData.oauth_state = oauthState;
        } else {
            Object.assign(requestData, serviceAccountData);
        }

        const result = await apiCall('/automation/move-and-shift-columns', 'POST', requestData);
        showStatus('move-status', result.message || 'Move and shift completed.', 'success');

        showResults('move-results', result, (r) => `
            <table>
                <thead>
                    <tr><th>Metric</th><th>Value</th></tr>
                </thead>
                <tbody>
                    <tr><td>Moved Rows</td><td>${r.moved_count ?? 0}</td></tr>
                    <tr><td>Remaining Rows</td><td>${r.remaining_count ?? 0}</td></tr>
                    <tr><td>Processed Rows</td><td>${r.processed_rows ?? 0}</td></tr>
                </tbody>
            </table>
        `);
    } catch (error) {
        showStatus('move-status', `Error: ${error.message}`, 'error');
    }
}

function formatCompareResults(result) {
    let html = '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">';
    
    html += '<div><h3>To Add (' + result.to_add_count + ')</h3><ul>';
    if (result.to_add.length > 0) {
        result.to_add.forEach(item => {
            html += `<li>${item}</li>`;
        });
    } else {
        html += '<li>None</li>';
    }
    html += '</ul></div>';
    
    html += '<div><h3>To Remove (' + result.to_remove_count + ')</h3><ul>';
    if (result.to_remove.length > 0) {
        result.to_remove.forEach(item => {
            html += `<li>${item}</li>`;
        });
    } else {
        html += '<li>None</li>';
    }
    html += '</ul></div>';
    
    html += '</div>';
    return html;
}

// Logout
async function logout() {
    if (!confirm('Are you sure you want to logout and close the browser?')) {
        return;
    }
    
    try {
        await apiCall('/automation/logout', 'POST');
        alert('Logged out successfully');
    } catch (error) {
        alert(`Error: ${error.message}`);
    }
}

function clearResults(elementId) {
    const element = document.getElementById(elementId);
    element.innerHTML = '';
}

// Health check on load
window.addEventListener('load', async () => {
    try {
        await apiCall('/health');
        console.log('API connection successful');
    } catch (error) {
        showStatus('sheets-status', 
            'Warning: Could not connect to API. Make sure the server is running on http://localhost:5001', 
            'error'
        );
    }
});

