;(function() {
    "use strict";
    const GRAPH_BASE = "https://graph.microsoft.com/beta";

    // ADD THE PASSWORD SECURITY CODE HERE - BEFORE ANYTHING ELSE
    // Secure Password Protection with SHA-256 Hashing
    (function() {
        // Check if user has already authenticated in this session
        const isAuthenticated = sessionStorage.getItem('intunePolicyManagerAuth');
        
        if (!isAuthenticated) {
            authenticateUser();
        }
    })();

    async function authenticateUser() {
        const PASSWORD_HASH = '9e5c3ce7906db8e6f333df0e10deed2263b0a8738a9abe3f3612b8f52846689e';
        
        const maxAttempts = 3;
        let attempts = 0;
        let authorized = false;
        
        while (attempts < maxAttempts && !authorized) {
            const password = prompt(
                `🔒 Intune Policy Manager - Secure Access\n\n` +
                `Enter password (${maxAttempts - attempts} attempt${attempts === 2 ? '' : 's'} remaining):`
            );
            
            // User clicked cancel
            if (password === null) {
                showAccessDenied('Authentication cancelled');
                throw new Error('Access cancelled by user');
            }
            
            // Hash the entered password and compare
            const enteredHash = await hashPassword(password);
            
            if (enteredHash === PASSWORD_HASH) {
                authorized = true;
                sessionStorage.setItem('intunePolicyManagerAuth', 'true');
                
                // Add session timeout (2 hours)
                sessionStorage.setItem('authExpiry', Date.now() + (2 * 60 * 60 * 1000));
            } else {
                attempts++;
                if (attempts < maxAttempts) {
                    alert(`❌ Incorrect password.\n\n${maxAttempts - attempts} attempt${attempts === 2 ? '' : 's'} remaining.`);
                }
            }
        }
        
        if (!authorized) {
            showAccessDenied('Maximum password attempts exceeded');
            throw new Error('Unauthorized - Maximum attempts exceeded');
        }
    }

    // SHA-256 hashing function
    async function hashPassword(password) {
        const msgUint8 = new TextEncoder().encode(password);
        const hashBuffer = await crypto.subtle.digest('SHA-256', msgUint8);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        return hashHex;
    }

    function showAccessDenied(message) {
        document.body.innerHTML = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Access Denied</title>
            </head>
            <body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
                <div style="display: flex; justify-content: center; align-items: center; min-height: 100vh; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);">
                    <div style="text-align: center; background: white; padding: 60px 40px; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); max-width: 450px; margin: 20px;">
                        <h1 style="color: #1a1a1a; margin-bottom: 25px; font-size: 36px; font-weight: 700; font-style: italic;">No HAM for you...</h1>
                        
                        <div style="margin: 30px 0;">
                            <img src="https://luiserodz.github.io/theHAM/Ham.jpg" 
                                 alt="No Ham for You" 
                                 style="max-width: 300px; width: 100%; height: auto; border-radius: 10px; box-shadow: 0 5px 20px rgba(0,0,0,0.2);">
                        </div>
                        
                        <p style="color: #666; margin-bottom: 10px; font-size: 18px; line-height: 1.5;">${message}</p>
                        
                        <div style="margin-top: 35px; padding-top: 25px; border-top: 1px solid #eee;">
                            <p style="color: #999; font-size: 15px;">Please contact support for access credentials.</p>
                            <p style="color: #bbb; font-size: 13px; margin-top: 15px;">Intune Policy Manager</p>
                        </div>
                    </div>
                </div>
            </body>
            </html>
        `;
    }

    // Check session expiry on each load
    (function checkSessionExpiry() {
        const authExpiry = sessionStorage.getItem('authExpiry');
        if (authExpiry && Date.now() > parseInt(authExpiry)) {
            sessionStorage.removeItem('intunePolicyManagerAuth');
            sessionStorage.removeItem('authExpiry');
            location.reload();
        }
    })();

    // Auto-logout on inactivity (30 minutes)
    let inactivityTimer;
    const INACTIVITY_TIMEOUT = 30 * 60 * 1000; // 30 minutes

    function resetInactivityTimer() {
        clearTimeout(inactivityTimer);
        inactivityTimer = setTimeout(() => {
            sessionStorage.removeItem('intunePolicyManagerAuth');
            sessionStorage.removeItem('authExpiry');
            alert('⏱️ Session expired due to inactivity.\n\nPlease refresh the page to log in again.');
            location.reload();
        }, INACTIVITY_TIMEOUT);
    }

    // Start inactivity timer if authenticated
    if (sessionStorage.getItem('intunePolicyManagerAuth')) {
        document.addEventListener('click', resetInactivityTimer);
        document.addEventListener('keypress', resetInactivityTimer);
        document.addEventListener('scroll', resetInactivityTimer);
        document.addEventListener('mousemove', resetInactivityTimer);
        resetInactivityTimer();
    }

    console.log('Intune Policy Manager v2 initializing...');
        
        // Global variables
        let msalInstance = null;
        let accessToken = null;
        let tokenExpiresAt = null;
        let currentAction = null;
        let lastFocusedElement = null;
        let selectedPolicies = {
            upload: new Set(),
            manage: new Set()
        };
        let allPolicies = [];
        let policiesForUpload = [];
        let rateLimitDelay = 100;
        const MAX_RETRY_ATTEMPTS = 3;
        let selectedGroupId = null;
        let selectedGroupName = null;
        let manageSelectedGroupId = null;
        let manageSelectedGroupName = null;

        // Policy type endpoints mapping
        const policyEndpoints = {
            configurationPolicies: {
                list: '/deviceManagement/configurationPolicies',
                delete: '/deviceManagement/configurationPolicies/',
                assignments: '/deviceManagement/configurationPolicies/{id}/assignments',
                assign: '/deviceManagement/configurationPolicies/{id}/assign'
            },
            deviceConfigurations: {
                list: '/deviceManagement/deviceConfigurations',
                delete: '/deviceManagement/deviceConfigurations/',
                assignments: '/deviceManagement/deviceConfigurations/{id}/assignments',
                assign: '/deviceManagement/deviceConfigurations/{id}/assign'
            },
            compliancePolicies: {
                list: '/deviceManagement/deviceCompliancePolicies',
                delete: '/deviceManagement/deviceCompliancePolicies/',
                assignments: '/deviceManagement/deviceCompliancePolicies/{id}/assignments',
                assign: '/deviceManagement/deviceCompliancePolicies/{id}/assign'
            },
            groupPolicyConfigurations: {
                list: '/deviceManagement/groupPolicyConfigurations',
                delete: '/deviceManagement/groupPolicyConfigurations/',
                assignments: '/deviceManagement/groupPolicyConfigurations/{id}/assignments',
                assign: '/deviceManagement/groupPolicyConfigurations/{id}/assign'
            },
            intentPolicies: {
                list: '/deviceManagement/intents',
                delete: '/deviceManagement/intents/',
                assignments: '/deviceManagement/intents/{id}/assignments',
                assign: '/deviceManagement/intents/{id}/assign'
            }
        };

        // Tab switching
        function switchTab(tabName) {
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
                content.setAttribute('aria-hidden', 'true');
            });
            
            // Remove active class from all tabs
            document.querySelectorAll('.nav-tab').forEach(tab => {
                tab.classList.remove('active');
                tab.setAttribute('aria-selected', 'false');
            });
            
            // Show selected tab content
            const selectedContent = document.getElementById(tabName + 'Tab');
            if (selectedContent) {
                selectedContent.classList.add('active');
                selectedContent.setAttribute('aria-hidden', 'false');
            }
            
            // Find and activate the corresponding tab button
            document.querySelectorAll('.nav-tab').forEach(tab => {
                if (tab.getAttribute('aria-controls') === tabName + 'Tab') {
                    tab.classList.add('active');
                    tab.setAttribute('aria-selected', 'true');
                }
            });
        }

        // Toast notification system
        function showToast(message, type = 'info', duration = 3000) {
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            toast.textContent = message;
            document.body.appendChild(toast);
            
            setTimeout(() => {
                toast.style.animation = 'slideOutRight 0.3s ease';
                setTimeout(() => {
                    if (document.body.contains(toast)) {
                        document.body.removeChild(toast);
                    }
                }, 300);
            }, duration);
        }
        function handleError(context, error) {
            console.error(context, error);
            showToast(`${context}${error && error.message ? ": " + error.message : ""}`, "error", 5000);
        }


        // Copy redirect URI
        function copyRedirectUri() {
            const uri = document.getElementById('redirectUriDisplay').textContent;
            navigator.clipboard.writeText(uri).then(() => {
                showToast('Redirect URI copied to clipboard', 'success');
            }).catch(() => {
                // Fallback
                const textArea = document.createElement('textarea');
                textArea.value = uri;
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
                showToast('Redirect URI copied to clipboard', 'success');
            });
        }

        // API request with retry logic
        async function makeApiRequest(url, options = {}, retryCount = 0) {
            try {
                const headers = {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                    ...options.headers
                };
                
                const requestOptions = {
                    ...options,
                    headers: headers
                };
                
                const response = await fetch(url, requestOptions);
                
                if (response.status === 429) {
                    const retryAfter = response.headers.get('Retry-After') || Math.pow(2, retryCount + 1);
                    rateLimitDelay = Math.min(rateLimitDelay * 2, 5000);
                    console.log(`Rate limited. Waiting ${retryAfter} seconds...`);
                    await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
                    
                    if (retryCount < MAX_RETRY_ATTEMPTS) {
                        return makeApiRequest(url, options, retryCount + 1);
                    }
                }
                
                if (response.status === 401) {
                    await refreshToken();
                    if (retryCount < 1) {
                        return makeApiRequest(url, options, retryCount + 1);
                    }
                }
                
                rateLimitDelay = Math.max(rateLimitDelay * 0.9, 100);
                
                return response;
            } catch (error) {
                if (retryCount < MAX_RETRY_ATTEMPTS) {
                    await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount + 1) * 1000));
                    return makeApiRequest(url, options, retryCount + 1);
                }
                throw error;
            }
        }

        // Initialize the application
        document.addEventListener('DOMContentLoaded', function() {
            console.log('DOM loaded, initializing...');
            
            // Set up collapsible
            const howToToggle = document.getElementById('howToToggle');
            if (howToToggle) {
                howToToggle.addEventListener('click', function() {
                    this.classList.toggle('active');
                    const content = document.getElementById('howToContent');
                    if (content) {
                        content.classList.toggle('active');
                        this.setAttribute('aria-expanded', this.classList.contains('active'));
                    }
                });
            }
            
            // Display current redirect URI
            const currentUrl = window.location.href.split('#')[0].split('?')[0];
            const redirectUriDisplay = document.getElementById('redirectUriDisplay');
            if (redirectUriDisplay) {
                redirectUriDisplay.textContent = currentUrl;
            }
            
            // Check for sandboxed environment
            if (window.location.hostname.includes('googleusercontent.com') || 
                window.location.hostname.includes('jsfiddle.net') ||
                window.location.hostname.includes('codepen.io')) {
                const sandboxWarning = document.getElementById('sandboxWarning');
                if (sandboxWarning) {
                    sandboxWarning.style.display = 'flex';
                }
            }
            
            // Set up file input
            const fileButton = document.querySelector('.file-input-wrapper button');
            const fileInput = document.getElementById('policyFiles');
            if (fileButton && fileInput) {
                fileButton.addEventListener('click', () => fileInput.click());
                fileInput.addEventListener('change', handleFileSelection);
            }
            
            // Setup authentication
            setupAuthentication();
            
            // Setup event listeners
            setupEventListeners();
        });

        // Setup authentication
        function setupAuthentication() {
            const clientIdInput = document.getElementById('clientId');
            const loginBtn = document.getElementById('loginBtn');
            
            if (!clientIdInput || !loginBtn) {
                handleError('Authentication elements not found');
                return;
            }
            
            // Load saved client ID if available
            const savedClientId = localStorage.getItem('intuneManagerClientId');
            if (savedClientId) {
                clientIdInput.value = savedClientId;
                clientIdInput.dispatchEvent(new Event('input'));
            }
            
            clientIdInput.addEventListener('input', (e) => {
                const clientId = e.target.value.trim();
                const guidPattern = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
                
                if (guidPattern.test(clientId)) {
                    loginBtn.disabled = false;
                    clientIdInput.style.borderColor = '#28a745';
                    localStorage.setItem('intuneManagerClientId', clientId);
                    initializeMsal(clientId);
                } else {
                    loginBtn.disabled = true;
                    clientIdInput.style.borderColor = clientId ? '#dc3545' : '#ddd';
                }
            });
            
            loginBtn.addEventListener('click', signIn);
            
            const signOutBtn = document.getElementById('signOutBtn');
            if (signOutBtn) {
                signOutBtn.addEventListener('click', signOut);
            }
        }

        // Initialize MSAL
        function initializeMsal(clientId) {
            const msalConfig = {
                auth: {
                    clientId: clientId,
                    authority: 'https://login.microsoftonline.com/common',
                    redirectUri: window.location.href.split('#')[0].split('?')[0]
                },
                cache: {
                    cacheLocation: 'sessionStorage',
                    storeAuthStateInCookie: false
                }
            };
            
            try {
                msalInstance = new msal.PublicClientApplication(msalConfig);
                console.log('MSAL initialized successfully');
                
                // Check for existing session
                checkExistingSession();
            } catch (error) {
                handleError('Error initializing MSAL:', error);
                showToast('Error initializing authentication', 'error');
            }
        }

        // Check for existing session
        async function checkExistingSession() {
            if (!msalInstance) return;
            
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                try {
                    const response = await msalInstance.acquireTokenSilent({
                        scopes: [
                            'User.Read',
                            'DeviceManagementConfiguration.ReadWrite.All',
                            'DeviceManagementApps.ReadWrite.All',
                            'DeviceManagementServiceConfig.ReadWrite.All',
                            'Group.Read.All'
                        ],
                        account: accounts[0]
                    });
                    
                    accessToken = response.accessToken;
                    tokenExpiresAt = new Date(response.expiresOn);
                    updateAuthUI(accounts[0]);
                    console.log('Existing session found and restored');
                } catch (error) {
                    console.log('Silent token acquisition failed:', error);
                }
            }
        }

        // Sign in
        async function signIn() {
            if (!msalInstance) {
                showToast('Please enter a valid Client ID first', 'error');
                return;
            }
            
            const loginRequest = {
                scopes: [
                    'User.Read',
                    'DeviceManagementConfiguration.ReadWrite.All',
                    'DeviceManagementApps.ReadWrite.All',
                    'DeviceManagementServiceConfig.ReadWrite.All',
                    'Group.Read.All'
                ]
            };
            
            try {
                const loginResponse = await msalInstance.loginPopup(loginRequest);
                console.log('Login successful:', loginResponse);
                
                accessToken = loginResponse.accessToken;
                tokenExpiresAt = new Date(loginResponse.expiresOn);
                
                updateAuthUI(loginResponse.account);
                showToast('Successfully signed in', 'success');
                
                // Auto-refresh policies after sign in
                setTimeout(() => {
                    refreshPolicies();
                }, 1000);
                
            } catch (error) {
                handleError('Login failed:', error);
                
                let errorMessage = 'Authentication failed';
                if (error.errorMessage) {
                    if (error.errorMessage.includes('AADSTS65001')) {
                        errorMessage = 'Admin consent required. Please have an admin grant consent for the app permissions.';
                    } else if (error.errorMessage.includes('AADSTS50011')) {
                        errorMessage = 'Invalid redirect URI. Please check your app registration.';
                    } else {
                        errorMessage = error.errorMessage;
                    }
                }
                
                showToast(errorMessage, 'error', 5000);
            }
        }

        // Sign out
        async function signOut() {
            if (!msalInstance) return;
            
            try {
                await msalInstance.logoutPopup();
                accessToken = null;
                tokenExpiresAt = null;
                updateAuthUI(null);
                showToast('Successfully signed out', 'success');
                
                // Clear data
                allPolicies = [];
                policiesForUpload = [];
                selectedPolicies = {
                    upload: new Set(),
                    manage: new Set()
                };
                
                // Reset UI
                document.getElementById('managePolicyList').innerHTML = `
                    <div class="empty-state">
                        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                        </svg>
                        <h3>No Policies Loaded</h3>
                        <p>Sign in and click "Refresh" to load existing policies</p>
                    </div>
                `;
                document.getElementById('policyStats').style.display = 'none';
                
            } catch (error) {
                handleError('Logout failed:', error);
                showToast('Sign out failed', 'error');
            }
        }

        // Update auth UI
        function updateAuthUI(account) {
            const authStatus = document.getElementById('authStatus');
            const userInfo = document.getElementById('userInfo');
            const userEmail = document.getElementById('userEmail');
            const loginBtn = document.getElementById('loginBtn');
            const signOutBtn = document.getElementById('signOutBtn');
            const statusDot = document.getElementById('statusDot');
            const connectionStatus = document.getElementById('connectionStatus');
            
            if (account) {
                authStatus.className = 'auth-status connected';
                authStatus.innerHTML = '<span>Connected to Microsoft Graph</span><span id="tokenExpiry" style="font-size: 0.875rem;"></span>';
                userInfo.style.display = 'block';
                userEmail.textContent = account.username;
                loginBtn.style.display = 'none';
                signOutBtn.style.display = 'inline-flex';
                statusDot.classList.add('connected');
                connectionStatus.textContent = 'Connected';
                
                updateTokenExpiry();
            } else {
                authStatus.className = 'auth-status disconnected';
                authStatus.innerHTML = '<span>Not connected to Microsoft Graph</span>';
                userInfo.style.display = 'none';
                loginBtn.style.display = 'inline-flex';
                signOutBtn.style.display = 'none';
                statusDot.classList.remove('connected');
                connectionStatus.textContent = 'Disconnected';
            }
        }

        // Update token expiry display
        function updateTokenExpiry() {
            if (!tokenExpiresAt) return;
            
            const tokenExpiry = document.getElementById('tokenExpiry');
            if (!tokenExpiry) return;
            
            const now = new Date();
            const expiresIn = Math.round((tokenExpiresAt - now) / 1000 / 60);
            
            if (expiresIn > 0) {
                tokenExpiry.textContent = `Token expires in ${expiresIn} min`;
            } else {
                tokenExpiry.textContent = 'Token expired';
            }
        }

        // Refresh token
        async function refreshToken() {
            if (!msalInstance) return;
            
            try {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length === 0) throw new Error('No accounts found');
                
                const silentRequest = {
                    scopes: [
                        'User.Read',
                        'DeviceManagementConfiguration.ReadWrite.All',
                        'DeviceManagementApps.ReadWrite.All',
                        'DeviceManagementServiceConfig.ReadWrite.All',
                        'Group.Read.All'
                    ],
                    account: accounts[0]
                };
                
                const response = await msalInstance.acquireTokenSilent(silentRequest);
                accessToken = response.accessToken;
                tokenExpiresAt = new Date(response.expiresOn);
                updateTokenExpiry();
                
                console.log('Token refreshed successfully');
            } catch (error) {
                handleError('Token refresh failed:', error);
                showToast('Session expired. Please sign in again.', 'error');
                await signOut();
            }
        }

        // Setup event listeners
        function setupEventListeners() {
            // Upload button
            const uploadSelectedBtn = document.getElementById('uploadSelectedBtn');
            if (uploadSelectedBtn) {
                uploadSelectedBtn.addEventListener('click', uploadSelectedPolicies);
            }
            
            // Modal button event listeners
            const confirmBtn = document.getElementById('confirmBtn');
            if (confirmBtn) {
                confirmBtn.addEventListener('click', confirmAction);
            }
            
            const cancelModalBtn = document.getElementById('cancelModalBtn');
            if (cancelModalBtn) {
                cancelModalBtn.addEventListener('click', closeModal);
            }
            
            // Assignment modal radio buttons
            document.querySelectorAll('input[name="manageAssignmentTarget"]').forEach(radio => {
                radio.addEventListener('change', (e) => {
                    const groupContainer = document.getElementById('manageGroupSelectionContainer');
                    if (groupContainer) {
                        if (e.target.value === 'group') {
                            groupContainer.style.display = 'block';
                        } else {
                            groupContainer.style.display = 'none';
                            manageSelectedGroupId = null;
                            manageSelectedGroupName = null;
                        }
                    }
                });
            });
            
            // Manage group select
            const manageGroupSelect = document.getElementById('manageGroupSelect');
            if (manageGroupSelect) {
                manageGroupSelect.addEventListener('change', (e) => {
                    if (e.target.value) {
                        const selectedOption = e.target.options[e.target.selectedIndex];
                        manageSelectedGroupId = e.target.value;
                        manageSelectedGroupName = selectedOption.textContent;
                        
                        const manageSelectedGroupInfo = document.getElementById('manageSelectedGroupInfo');
                        const manageSelectedGroupNameElement = document.getElementById('manageSelectedGroupName');
                        if (manageSelectedGroupInfo) {
                            manageSelectedGroupInfo.style.display = 'block';
                        }
                        if (manageSelectedGroupNameElement) {
                            manageSelectedGroupNameElement.textContent = selectedOption.textContent;
                        }
                    } else {
                        manageSelectedGroupId = null;
                        manageSelectedGroupName = null;
                        const manageSelectedGroupInfo = document.getElementById('manageSelectedGroupInfo');
                        if (manageSelectedGroupInfo) {
                            manageSelectedGroupInfo.style.display = 'none';
                        }
                    }
                });
            }
            
            // Upload assignment target radio buttons
            document.querySelectorAll('input[name="assignmentTarget"]').forEach(radio => {
                radio.addEventListener('change', (e) => {
                    const groupContainer = document.getElementById('groupSelectionContainer');
                    if (groupContainer) {
                        if (e.target.value === 'group') {
                            groupContainer.style.display = 'block';
                        } else {
                            groupContainer.style.display = 'none';
                            selectedGroupId = null;
                            selectedGroupName = null;
                        }
                    }
                });
            });
            
            // Upload group select
            const groupSelect = document.getElementById('groupSelect');
            if (groupSelect) {
                groupSelect.addEventListener('change', (e) => {
                    if (e.target.value) {
                        const selectedOption = e.target.options[e.target.selectedIndex];
                        selectedGroupId = e.target.value;
                        selectedGroupName = selectedOption.textContent;
                        
                        const selectedGroupInfo = document.getElementById('selectedGroupInfo');
                        const selectedGroupNameElement = document.getElementById('selectedGroupName');
                        if (selectedGroupInfo) {
                            selectedGroupInfo.style.display = 'block';
                        }
                        if (selectedGroupNameElement) {
                            selectedGroupNameElement.textContent = selectedOption.textContent;
                        }
                    } else {
                        selectedGroupId = null;
                        selectedGroupName = null;
                        const selectedGroupInfo = document.getElementById('selectedGroupInfo');
                        if (selectedGroupInfo) {
                            selectedGroupInfo.style.display = 'none';
                        }
                    }
                });
            }
            
            // Manage search and filter inputs
            const manageSearchInput = document.getElementById('manageSearchInput');
            if (manageSearchInput) {
                manageSearchInput.addEventListener('input', filterPolicies);
            }
            
            const policyTypeFilter = document.getElementById('policyTypeFilter');
            if (policyTypeFilter) {
                policyTypeFilter.addEventListener('change', filterPolicies);
            }
            
            const showAssignmentsOnly = document.getElementById('showAssignmentsOnly');
            if (showAssignmentsOnly) {
                showAssignmentsOnly.addEventListener('change', filterPolicies);
            }
            
            // Token expiry update interval
            setInterval(updateTokenExpiry, 60000); // Update every minute
        }

        // Handle file selection
        function handleFileSelection(event) {
            const files = event.target.files;
            if (files.length === 0) return;
            
            policiesForUpload = [];
            const uploadList = document.getElementById('uploadList');
            const uploadPolicyList = document.getElementById('uploadPolicyList');
            
            selectedPolicies.upload.clear();
            uploadPolicyList.innerHTML = '';
            
            Array.from(files).forEach((file, index) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const policy = JSON.parse(e.target.result);
                        policy._fileName = file.name;
                        policy._index = index;
                        policiesForUpload.push(policy);
                        
                        const policyItem = createPolicyItem(policy, 'upload');
                        uploadPolicyList.appendChild(policyItem);
                        
                        if (policiesForUpload.length === files.length) {
                            uploadList.style.display = 'block';
                            updateUploadCount();
                        }
                    } catch (error) {
                        handleError('Error parsing file:', file.name, error);
                        showToast(`Error parsing ${file.name}: Invalid JSON`, 'error');
                    }
                };
                reader.readAsText(file);
            });
        }

        // Create policy item element
        function createPolicyItem(policy, context) {
            const item = document.createElement('div');
            item.className = 'policy-item';
            
            const identifier = context === 'upload' ? policy._index : policy.id;
            item.dataset.policyId = identifier;
            
            const header = document.createElement('div');
            header.className = 'policy-header';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `${context}-${identifier}`;
            checkbox.addEventListener('change', (e) => {
                if (e.target.checked) {
                    if (context === 'upload') {
                        selectedPolicies[context].add(policy._index);
                    } else {
                        selectedPolicies[context].add(policy.id);
                    }
                    item.classList.add('selected');
                } else {
                    if (context === 'upload') {
                        selectedPolicies[context].delete(policy._index);
                    } else {
                        selectedPolicies[context].delete(policy.id);
                    }
                    item.classList.remove('selected');
                }
                updateButtonStates(context);
            });
            
            const name = document.createElement('label');
            name.className = 'policy-name';
            name.htmlFor = checkbox.id;
            name.textContent = policy.displayName || policy.name || 'Unnamed Policy';
            
            const actions = document.createElement('div');
            actions.className = 'action-buttons';
            
            if (context === 'manage') {
                const viewBtn = document.createElement('button');
                viewBtn.className = 'icon-btn view';
                viewBtn.textContent = 'View';
                viewBtn.onclick = () => viewPolicy(policy);
                actions.appendChild(viewBtn);
            }
            
            header.appendChild(checkbox);
            header.appendChild(name);
            header.appendChild(actions);
            
            const details = document.createElement('div');
            details.className = 'policy-details';
            
            if (policy['@odata.type']) {
                const type = document.createElement('span');
                type.className = 'policy-type';
                type.textContent = policy['@odata.type'].split('.').pop();
                details.appendChild(type);
            }
            
            if (policy.description) {
                const desc = document.createElement('span');
                desc.textContent = policy.description.substring(0, 100) + (policy.description.length > 100 ? '...' : '');
                details.appendChild(desc);
            }
            
            if (context === 'manage' && policy._assignments && policy._assignments.length > 0) {
                const assignmentList = document.createElement('div');
                assignmentList.className = 'assignment-list';
                assignmentList.style.marginTop = '8px';
                
                policy._assignments.forEach(assignment => {
                    const assignItem = document.createElement('div');
                    assignItem.className = 'assignment-item';
                    
                    let targetName = 'Unknown';
                    if (assignment.target['@odata.type'].includes('allDevices')) {
                        targetName = '🌐 All Devices';
                    } else if (assignment.target['@odata.type'].includes('allUsers')) {
                        targetName = '👥 All Users';
                    } else if (assignment.target.groupId) {
                        targetName = `👥 Group: ${assignment.target.groupId}`;
                    }
                    
                    assignItem.textContent = targetName;
                    assignmentList.appendChild(assignItem);
                });
                
                details.appendChild(assignmentList);
            }
            
            item.appendChild(header);
            item.appendChild(details);
            
            return item;
        }

        // View policy details
        function viewPolicy(policy) {
            const viewer = document.getElementById('policyViewer');
            const title = document.getElementById('policyViewerTitle');
            const content = document.getElementById('policyViewerContent');
            lastFocusedElement = document.activeElement;
            
            title.textContent = policy.displayName || policy.name || 'Policy Details';
            content.textContent = JSON.stringify(policy, null, 2);
            
            viewer.style.display = 'block';
        }

        // Close policy viewer
        function closePolicyViewer() {
            document.getElementById('policyViewer').style.display = 'none';
            if (lastFocusedElement) lastFocusedElement.focus();
        }

        // Update button states
        function updateButtonStates(context) {
            const count = selectedPolicies[context].size;
            
            if (context === 'upload') {
                document.getElementById('uploadSelectedBtn').disabled = count === 0;
                document.getElementById('uploadCount').textContent = count;
            } else if (context === 'manage') {
                document.getElementById('exportSelectedBtn').disabled = count === 0;
                document.getElementById('duplicateSelectedBtn').disabled = count === 0;
                document.getElementById('assignBtn').disabled = count === 0;
                document.getElementById('unassignBtn').disabled = count === 0;
                document.getElementById('deleteSelectedBtn').disabled = count === 0;
                document.getElementById('selectedPoliciesCount').textContent = count;
            }
        }

        // Select all functions
        function selectAllUpload() {
            selectedPolicies.upload.clear();
            policiesForUpload.forEach((policy, index) => {
                selectedPolicies.upload.add(index);
            });
            updateCheckboxes('upload', true);
            updateButtonStates('upload');
        }

        function deselectAllUpload() {
            deselectAll('upload');
        }

        function selectAllManage() {
            const filtered = filterPoliciesList(allPolicies);
            selectedPolicies.manage.clear();
            filtered.forEach(policy => {
                selectedPolicies.manage.add(policy.id);
            });
            updateCheckboxes('manage', true);
            updateButtonStates('manage');
        }

        function deselectAllManage() {
            deselectAll('manage');
        }

        function deselectAll(context) {
            selectedPolicies[context].clear();
            updateCheckboxes(context, false);
            updateButtonStates(context);
        }

        function updateCheckboxes(context, checked) {
            const container = context === 'upload' ? 'uploadPolicyList' : 'managePolicyList';
            
            document.querySelectorAll(`#${container} input[type="checkbox"]`).forEach(cb => {
                cb.checked = checked;
                cb.closest('.policy-item').classList.toggle('selected', checked);
            });
        }

        // Update counts
        function updateUploadCount() {
            document.getElementById('uploadCount').textContent = policiesForUpload.length;
        }

        // Search for groups
        async function searchGroups() {
            if (!accessToken) {
                showToast('Please sign in first', 'error');
                return;
            }
            
            const searchTerm = document.getElementById('groupSearch').value.trim();
            if (!searchTerm) {
                showToast('Please enter a group name to search', 'warning');
                return;
            }
            
            try {
                const response = await makeApiRequest(
                    `https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,'${searchTerm}')&$select=id,displayName,description&$top=50`
                );
                
                if (response.ok) {
                    const data = await response.json();
                    const groupSelect = document.getElementById('groupSelect');
                    
                    groupSelect.innerHTML = '<option value="">Select a group...</option>';
                    
                    if (data.value && data.value.length > 0) {
                        data.value.forEach(group => {
                            const option = document.createElement('option');
                            option.value = group.id;
                            option.textContent = group.displayName;
                            option.title = group.description || '';
                            groupSelect.appendChild(option);
                        });
                        
                        groupSelect.style.display = 'block';
                        showToast(`Found ${data.value.length} groups`, 'success');
                    } else {
                        showToast('No groups found matching your search', 'info');
                        groupSelect.style.display = 'none';
                    }
                } else {
                    showToast('Error searching for groups', 'error');
                }
            } catch (error) {
                handleError('Error searching groups:', error);
                showToast('Error searching for groups', 'error');
            }
        }

        // Search for groups in manage assignment
        async function searchManageGroups() {
            if (!accessToken) {
                showToast('Please sign in first', 'error');
                return;
            }
            
            const searchTerm = document.getElementById('manageGroupSearch').value.trim();
            if (!searchTerm) {
                showToast('Please enter a group name to search', 'warning');
                return;
            }
            
            try {
                const response = await makeApiRequest(
                    `https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,'${searchTerm}')&$select=id,displayName,description&$top=50`
                );
                
                if (response.ok) {
                    const data = await response.json();
                    const groupSelect = document.getElementById('manageGroupSelect');
                    
                    groupSelect.innerHTML = '<option value="">Select a group...</option>';
                    
                    if (data.value && data.value.length > 0) {
                        data.value.forEach(group => {
                            const option = document.createElement('option');
                            option.value = group.id;
                            option.textContent = group.displayName;
                            option.title = group.description || '';
                            groupSelect.appendChild(option);
                        });
                        
                        groupSelect.style.display = 'block';
                        showToast(`Found ${data.value.length} groups`, 'success');
                    } else {
                        showToast('No groups found matching your search', 'info');
                        groupSelect.style.display = 'none';
                    }
                } else {
                    showToast('Error searching for groups', 'error');
                }
            } catch (error) {
                handleError('Error searching groups:', error);
                showToast('Error searching for groups', 'error');
            }
        }

        // Refresh policies
        async function refreshPolicies() {
            if (!accessToken) {
                showToast('Please sign in first', 'error');
                return;
            }
            
            const managePolicyList = document.getElementById('managePolicyList');
            managePolicyList.innerHTML = '<div class="skeleton skeleton-item"></div>'.repeat(5);
            managePolicyList.classList.add('loading');
            
            try {
                allPolicies = [];
                
                // Fetch all policy types
                for (const [type, endpoints] of Object.entries(policyEndpoints)) {
                    try {
                        const response = await makeApiRequest(`${GRAPH_BASE}${endpoints.list}`);
                        const data = await response.json();
                        
                        if (data.value) {
                            data.value.forEach(policy => {
                                policy._type = type;
                                allPolicies.push(policy);
                            });
                        }
                    } catch (error) {
                        handleError(`Error fetching ${type}:`, error);
                    }
                }
                
                // Fetch assignments for each policy
                for (const policy of allPolicies) {
                    const endpoint = policyEndpoints[policy._type];
                    if (!endpoint || !endpoint.assignments) continue;
                    
                    try {
                        const url = `${GRAPH_BASE}${endpoint.assignments.replace('{id}', policy.id)}`;
                        const response = await makeApiRequest(url);
                        
                        if (response.ok) {
                            const data = await response.json();
                            if (data.value && data.value.length > 0) {
                                policy._assignments = data.value;
                            }
                        }
                    } catch (error) {
                        handleError(`Error fetching assignments for ${policy.displayName}:`, error);
                    }
                }
                
                showToast(`Loaded ${allPolicies.length} policies`, 'success');
                displayPolicies();
                updatePolicyStats();
                
            } catch (error) {
                handleError('Error refreshing policies:', error);
                showToast('Error loading policies', 'error');
            } finally {
                managePolicyList.classList.remove('loading');
            }
        }

        // Display policies
        function displayPolicies() {
            const managePolicyList = document.getElementById('managePolicyList');
            const filtered = filterPoliciesList(allPolicies);
            
            if (filtered.length === 0) {
                managePolicyList.innerHTML = `
                    <div class="empty-state">
                        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                        </svg>
                        <h3>No Policies Found</h3>
                        <p>No policies match your search criteria</p>
                    </div>
                `;
                return;
            }
            
            managePolicyList.innerHTML = '';
            filtered.forEach(policy => {
                const item = createPolicyItem(policy, 'manage');
                managePolicyList.appendChild(item);
            });
            
            document.getElementById('manageFilterStats').textContent = `Showing ${filtered.length} of ${allPolicies.length} policies`;
        }

        // Filter policies
        function filterPolicies() {
            displayPolicies();
        }

        function filterPoliciesList(policies) {
            const searchTerm = document.getElementById('manageSearchInput').value.toLowerCase();
            const typeFilter = document.getElementById('policyTypeFilter').value;
            const showAssignmentsOnly = document.getElementById('showAssignmentsOnly').checked;
            
            return policies.filter(policy => {
                const matchesSearch = !searchTerm || 
                    (policy.displayName || policy.name || '').toLowerCase().includes(searchTerm) ||
                    (policy.description || '').toLowerCase().includes(searchTerm);
                
                const matchesType = !typeFilter || policy._type === typeFilter;
                
                const matchesAssignment = !showAssignmentsOnly || (policy._assignments && policy._assignments.length > 0);
                
                return matchesSearch && matchesType && matchesAssignment;
            });
        }

        // Update policy stats
        function updatePolicyStats() {
            document.getElementById('policyStats').style.display = 'grid';
            document.getElementById('totalPoliciesCount').textContent = allPolicies.length;
            
            const configCount = allPolicies.filter(p => p._type === 'configurationPolicies').length;
            const complianceCount = allPolicies.filter(p => p._type === 'compliancePolicies').length;
            
            document.getElementById('configPoliciesCount').textContent = configCount;
            document.getElementById('compliancePoliciesCount').textContent = complianceCount;
        }

        // Upload selected policies
        async function uploadSelectedPolicies() {
            if (!accessToken) {
                showToast('Please sign in first', 'error');
                return;
            }
            
            const selected = Array.from(selectedPolicies.upload);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            const assignmentTarget = document.querySelector('input[name="assignmentTarget"]:checked').value;
            if (assignmentTarget === 'group' && !selectedGroupId) {
                showToast('Please select a group for assignment', 'error');
                return;
            }
            
            const uploadProgress = document.getElementById('uploadProgress');
            const uploadResults = document.getElementById('uploadResults');
            const progressFill = document.getElementById('uploadProgressFill');
            const progressText = document.getElementById('uploadProgressText');
            const progressDetails = document.getElementById('uploadProgressDetails');
            const resultsTableBody = document.getElementById('uploadResultsTableBody');
            
            uploadProgress.style.display = 'block';
            uploadResults.style.display = 'none';
            resultsTableBody.innerHTML = '';
            
            const prefix = document.getElementById('policyPrefix').value;
            const results = [];
            
            for (let i = 0; i < selected.length; i++) {
                const policyIndex = selected[i];
                const policy = policiesForUpload[policyIndex];
                
                if (!policy) {
                    handleError('Policy not found at index:', policyIndex);
                    results.push({
                        name: `Unknown policy at index ${policyIndex}`,
                        status: 'Error',
                        details: 'Policy data not found'
                    });
                    continue;
                }
                
                const progress = Math.round(((i + 1) / selected.length) * 100);
                
                progressFill.style.width = `${progress}%`;
                progressText.textContent = `${progress}%`;
                progressDetails.textContent = `Uploading ${policy.displayName || policy.name || 'policy'}...`;
                
                try {
                    const uploadPolicy = JSON.parse(JSON.stringify(policy));
                    
                    // Remove metadata properties
                    const propertiesToRemove = [
                        '_fileName', '_index', 'id', 'createdDateTime', 'lastModifiedDateTime',
                        'version', 'referencedConfigurationPolicyCount', 'roleScopeTagIds',
                        'supportsScopeTags', 'deviceManagementApplicabilityRuleOsEdition',
                        'deviceManagementApplicabilityRuleOsVersion', 'deviceManagementApplicabilityRuleDeviceMode',
                        'creationSource', 'isAssigned', 'settingCount', 'templateReference', 'platformType'
                    ];
                    
                    propertiesToRemove.forEach(prop => {
                        delete uploadPolicy[prop];
                    });
                    
                    // Apply prefix if specified
                    if (prefix) {
                        uploadPolicy.displayName = prefix + (uploadPolicy.displayName || uploadPolicy.name || '');
                        if (uploadPolicy.name) {
                            uploadPolicy.name = prefix + uploadPolicy.name;
                        }
                    }
                    
                    // Ensure required fields
                    if (!uploadPolicy.displayName && !uploadPolicy.name) {
                        uploadPolicy.displayName = 'Unnamed Policy';
                    }
                    
                    // Determine the correct @odata.type and endpoint
                    if (!uploadPolicy['@odata.type']) {
                        if (uploadPolicy.settings) {
                            uploadPolicy['@odata.type'] = '#microsoft.graph.deviceManagementConfigurationPolicy';
                        } else if (uploadPolicy.scheduledActionsForRule) {
                            uploadPolicy['@odata.type'] = '#microsoft.graph.deviceCompliancePolicy';
                        } else {
                            uploadPolicy['@odata.type'] = '#microsoft.graph.deviceConfiguration';
                        }
                    }
                    
                    let endpoint = '/deviceManagement/deviceConfigurations';
                    const odataType = uploadPolicy['@odata.type'] || '';
                    
                    if (odataType && typeof odataType === 'string') {
                        if (odataType.includes('deviceManagementConfigurationPolicy') || odataType.includes('configurationPolicy')) {
                            endpoint = '/deviceManagement/configurationPolicies';
                            // Ensure required fields for configuration policies
                            if (!uploadPolicy.platforms) {
                                uploadPolicy.platforms = 'windows10';
                            }
                            if (!uploadPolicy.technologies) {
                                uploadPolicy.technologies = 'mdm';
                            }
                            if (!uploadPolicy.settings) {
                                uploadPolicy.settings = [];
                            }
                        } else if (odataType.includes('deviceCompliancePolicy') || odataType.includes('compliancePolicy')) {
                            endpoint = '/deviceManagement/deviceCompliancePolicies';
                        } else if (odataType.includes('groupPolicyConfiguration')) {
                            endpoint = '/deviceManagement/groupPolicyConfigurations';
                        } else if (odataType.includes('deviceManagementIntent') || odataType.includes('intent')) {
                            endpoint = '/deviceManagement/intents';
                        }
                    }
                    
                    const response = await makeApiRequest(
                        `${GRAPH_BASE}${endpoint}`,
                        {
                            method: 'POST',
                            body: JSON.stringify(uploadPolicy),
                            headers: {
                                'Content-Type': 'application/json'
                            }
                        }
                    );
                    
                    if (response.ok) {
                        const createdPolicy = await response.json();
                        
                        // Handle assignment if specified
                        if (assignmentTarget !== 'none' && createdPolicy.id) {
                            progressDetails.textContent = `Assigning ${policy.displayName || policy.name || 'policy'}...`;
                            
                            try {
                                await assignPolicyDuringUpload(createdPolicy, endpoint, assignmentTarget);
                                
                                results.push({
                                    name: uploadPolicy.displayName || uploadPolicy.name,
                                    status: 'Success',
                                    details: `Policy uploaded and assigned to ${assignmentTarget === 'group' ? selectedGroupName : assignmentTarget}`
                                });
                            } catch (assignError) {
                                handleError('Assignment error:', assignError);
                                results.push({
                                    name: uploadPolicy.displayName || uploadPolicy.name,
                                    status: 'Warning',
                                    details: `Policy uploaded but assignment failed: ${assignError.message}`
                                });
                            }
                        } else {
                            results.push({
                                name: uploadPolicy.displayName || uploadPolicy.name,
                                status: 'Success',
                                details: 'Policy uploaded successfully'
                            });
                        }
                    } else {
                        const errorText = await response.text();
                        handleError('Upload failed:', errorText);
                        
                        let errorMessage = `HTTP ${response.status}`;
                        try {
                            const errorJson = JSON.parse(errorText);
                            errorMessage = errorJson.error?.message || errorJson.message || errorText;
                        } catch {
                            errorMessage = errorText || `HTTP ${response.status}`;
                        }
                        
                        results.push({
                            name: uploadPolicy.displayName || uploadPolicy.name,
                            status: 'Error',
                            details: errorMessage
                        });
                    }
                } catch (error) {
                    handleError('Upload error:', error);
                    results.push({
                        name: policy.displayName || policy.name,
                        status: 'Error',
                        details: error.message
                    });
                }
                
                // Rate limiting delay
                if (i < selected.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, rateLimitDelay));
                }
            }
            
            // Hide progress, show results
            uploadProgress.style.display = 'none';
            uploadResults.style.display = 'block';
            
            results.forEach(result => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${result.name}</td>
                    <td class="status-${result.status.toLowerCase()}">${result.status}</td>
                    <td>${result.details}</td>
                `;
                resultsTableBody.appendChild(row);
            });
            
            const successCount = results.filter(r => r.status === 'Success').length;
            showToast(`Upload complete: ${successCount}/${results.length} successful`, 
                successCount === results.length ? 'success' : 'warning');
        }

        // Assign policy during upload
        async function assignPolicyDuringUpload(policy, endpoint, assignmentTarget) {
            let assignEndpoint = '';
            
            if (endpoint.includes('configurationPolicies')) {
                assignEndpoint = `/deviceManagement/configurationPolicies/${policy.id}/assign`;
            } else if (endpoint.includes('deviceConfigurations')) {
                assignEndpoint = `/deviceManagement/deviceConfigurations/${policy.id}/assign`;
            } else if (endpoint.includes('deviceCompliancePolicies')) {
                assignEndpoint = `/deviceManagement/deviceCompliancePolicies/${policy.id}/assign`;
            } else if (endpoint.includes('groupPolicyConfigurations')) {
                assignEndpoint = `/deviceManagement/groupPolicyConfigurations/${policy.id}/assign`;
            }
            
            if (!assignEndpoint) {
                throw new Error('Cannot determine assignment endpoint for this policy type');
            }
            
            const assignments = [];
            
            if (assignmentTarget === 'allDevices') {
                assignments.push({
                    target: {
                        '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget'
                    }
                });
            } else if (assignmentTarget === 'allUsers') {
                assignments.push({
                    target: {
                        '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget'
                    }
                });
            } else if (assignmentTarget === 'group' && selectedGroupId) {
                assignments.push({
                    target: {
                        '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                        groupId: selectedGroupId
                    }
                });
            }
            
            const response = await makeApiRequest(
                `${GRAPH_BASE}${assignEndpoint}`,
                {
                    method: 'POST',
                    body: JSON.stringify({ assignments })
                }
            );
            
            if (!response.ok) {
                const error = await response.text();
                throw new Error(error || `Assignment failed with status ${response.status}`);
            }
        }

        // Export results
        function exportResults(context) {
            const table = document.getElementById(`${context}ResultsTableBody`);
            const rows = Array.from(table.querySelectorAll('tr'));
            
            const csv = [
                ['Policy Name', 'Status', 'Details'],
                ...rows.map(row => {
                    const cells = row.querySelectorAll('td');
                    return [cells[0].textContent, cells[1].textContent, cells[2].textContent];
                })
            ].map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
            
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `intune-${context}-results-${new Date().toISOString().split('T')[0]}.csv`;
            a.click();
            URL.revokeObjectURL(url);
        }

        // Clear results
        function clearResults(context) {
            document.getElementById(`${context}Results`).style.display = 'none';
            document.getElementById(`${context}ResultsTableBody`).innerHTML = '';
        }

        // Export selected policies
        async function exportSelected() {
            const selected = Array.from(selectedPolicies.manage);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            const policies = selected.map(id => allPolicies.find(p => p.id === id)).filter(Boolean);
            
            const zip = new JSZip();
            policies.forEach(policy => {
                const fileName = `${policy.displayName || policy.name || policy.id}.json`;
                zip.file(fileName, JSON.stringify(policy, null, 2));
            });
            
            const content = await zip.generateAsync({ type: 'blob' });
            const url = URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = `intune-policies-export-${new Date().toISOString().split('T')[0]}.zip`;
            a.click();
            URL.revokeObjectURL(url);
            
            showToast(`Exported ${policies.length} policies`, 'success');
        }

        // Duplicate selected policies
        async function duplicateSelected() {
            const selected = Array.from(selectedPolicies.manage);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            let successCount = 0;
            
            for (const policyId of selected) {
                const policy = allPolicies.find(p => p.id === policyId);
                if (!policy) continue;
                
                try {
                    const duplicate = JSON.parse(JSON.stringify(policy));
                    
                    // Remove metadata properties
                    const propertiesToRemove = [
                        'id', 'createdDateTime', 'lastModifiedDateTime', 'version',
                        'referencedConfigurationPolicyCount', 'roleScopeTagIds',
                        'supportsScopeTags', 'deviceManagementApplicabilityRuleOsEdition',
                        'deviceManagementApplicabilityRuleOsVersion', 'deviceManagementApplicabilityRuleDeviceMode',
                        'creationSource', 'isAssigned', 'settingCount', 'templateReference',
                        'platformType', '_type', '_assignments'
                    ];
                    
                    propertiesToRemove.forEach(prop => {
                        delete duplicate[prop];
                    });
                    
                    // Add "Copy of" prefix
                    duplicate.displayName = `Copy of ${duplicate.displayName || duplicate.name}`;
                    if (duplicate.name) {
                        duplicate.name = `Copy of ${duplicate.name}`;
                    }
                    
                    const endpoint = policyEndpoints[policy._type];
                    if (!endpoint || !endpoint.list) continue;
                    
                    const response = await makeApiRequest(
                        `${GRAPH_BASE}${endpoint.list}`,
                        {
                            method: 'POST',
                            body: JSON.stringify(duplicate)
                        }
                    );
                    
                    if (response.ok) {
                        successCount++;
                    }
                } catch (error) {
                    handleError('Error duplicating policy:', error);
                }
            }
            
            showToast(`Duplicated ${successCount}/${selected.length} policies`, 
                successCount === selected.length ? 'success' : 'warning');
            
            if (successCount > 0) {
                setTimeout(refreshPolicies, 1000);
            }
        }

        // Assign selected policies
        function assignSelected() {
            const selected = Array.from(selectedPolicies.manage);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            // Show assignment modal
            const modal = document.getElementById('assignmentModal');
            lastFocusedElement = document.activeElement;
            document.getElementById('assignmentPolicyCount').textContent = selected.length;
            
            modal.style.display = 'block';
            modal.setAttribute('aria-hidden', 'false');
            
            // Reset selection
            document.querySelector('input[name="manageAssignmentTarget"][value="allUsers"]').checked = true;
            document.getElementById('manageGroupSelectionContainer').style.display = 'none';
            document.getElementById('manageGroupSelect').style.display = 'none';
            document.getElementById('manageSelectedGroupInfo').style.display = 'none';
            manageSelectedGroupId = null;
            manageSelectedGroupName = null;
        }
        
        // Close assignment modal
        function closeAssignmentModal() {
            const modal = document.getElementById('assignmentModal');
            modal.style.display = 'none';
            modal.setAttribute('aria-hidden', 'true');
            if (lastFocusedElement) lastFocusedElement.focus();
        
        // Confirm assignment
        async function confirmAssignment() {
            const assignmentTarget = document.querySelector('input[name="manageAssignmentTarget"]:checked').value;
            
            if (assignmentTarget === 'group' && !manageSelectedGroupId) {
                showToast('Please select a group for assignment', 'error');
                return;
            }
            
            closeAssignmentModal();
            
            const selected = Array.from(selectedPolicies.manage);
            const deleteProgress = document.getElementById('deleteProgress');
            const deleteResults = document.getElementById('deleteResults');
            const progressFill = document.getElementById('deleteProgressFill');
            const progressText = document.getElementById('deleteProgressText');
            const progressDetails = document.getElementById('deleteProgressDetails');
            const resultsTableBody = document.getElementById('deleteResultsTableBody');
            
            // Show progress, hide results
            deleteProgress.style.display = 'block';
            deleteResults.style.display = 'none';
            resultsTableBody.innerHTML = '';
            
            // Reset progress
            progressFill.style.width = '0%';
            progressText.textContent = '0%';
            
            const results = [];
            let successCount = 0;
            
            for (let i = 0; i < selected.length; i++) {
                const policyId = selected[i];
                const policy = allPolicies.find(p => p.id === policyId);
                if (!policy) continue;
                
                const progress = Math.round(((i + 1) / selected.length) * 100);
                progressFill.style.width = `${progress}%`;
                progressText.textContent = `${progress}%`;
                progressDetails.textContent = `Assigning ${policy.displayName || policy.name}...`;
                
                const endpoint = policyEndpoints[policy._type];
                if (!endpoint || !endpoint.assign) {
                    results.push({
                        name: policy.displayName || policy.name,
                        status: 'Error',
                        details: 'Unknown policy type'
                    });
                    continue;
                }
                
                try {
                    // Build assignment object
                    const assignments = [];
                    
                    if (assignmentTarget === 'allDevices') {
                        assignments.push({
                            target: {
                                '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget'
                            }
                        });
                    } else if (assignmentTarget === 'allUsers') {
                        assignments.push({
                            target: {
                                '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget'
                            }
                        });
                    } else if (assignmentTarget === 'group' && manageSelectedGroupId) {
                        assignments.push({
                            target: {
                                '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                                groupId: manageSelectedGroupId
                            }
                        });
                    }
                    
                    const url = `${GRAPH_BASE}${endpoint.assign.replace('{id}', policy.id)}`;
                    const response = await makeApiRequest(url, {
                        method: 'POST',
                        body: JSON.stringify({ assignments })
                    });
                    
                    if (response.ok || response.status === 204) {
                        successCount++;
                        const targetName = assignmentTarget === 'group' ? manageSelectedGroupName : assignmentTarget;
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Success',
                            details: `Assigned to ${targetName}`
                        });
                    } else {
                        const error = await response.text();
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Error',
                            details: error || `HTTP ${response.status}`
                        });
                    }
                } catch (error) {
                    handleError(`Error assigning ${policy.displayName}:`, error);
                    results.push({
                        name: policy.displayName || policy.name,
                        status: 'Error',
                        details: error.message
                    });
                }
                
                if (i < selected.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, rateLimitDelay));
                }
            }
            
            // Hide progress, show results
            deleteProgress.style.display = 'none';
            deleteResults.style.display = 'block';
            
            results.forEach(result => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${result.name}</td>
                    <td class="status-${result.status.toLowerCase()}">${result.status}</td>
                    <td>${result.details}</td>
                `;
                resultsTableBody.appendChild(row);
            });
            
            showToast(`Assigned ${successCount}/${selected.length} policies`, 
                successCount === selected.length ? 'success' : 'warning');
            
            if (successCount > 0) {
                setTimeout(refreshPolicies, 1000);
            }
        }

        // Unassign selected policies
        async function unassignSelected() {
            const selected = Array.from(selectedPolicies.manage);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            const policiesWithAssignments = selected
                .map(id => allPolicies.find(p => p.id === id))
                .filter(p => p && p._assignments && p._assignments.length > 0);
            
            if (policiesWithAssignments.length === 0) {
                showToast('Selected policies have no assignments', 'info');
                return;
            }
            
            // Show confirmation modal
            const modal = document.getElementById('confirmModal');
            const modalHeader = document.getElementById('modalHeader');
            const modalBody = document.getElementById('modalBody');
            const confirmBtn = document.getElementById('confirmBtn');
            lastFocusedElement = document.activeElement;
            
            modalHeader.textContent = 'Remove All Assignments?';
            modalBody.textContent = `This will remove all assignments from ${policiesWithAssignments.length} policies. Devices will no longer receive these configurations.`;
            confirmBtn.className = 'btn btn-warning';
            
            modal.style.display = 'block';
            modal.setAttribute('aria-hidden', 'false');
            
            // Set the action to execute when confirmed
            currentAction = async () => {
                const deleteProgress = document.getElementById('deleteProgress');
                const deleteResults = document.getElementById('deleteResults');
                const progressFill = document.getElementById('deleteProgressFill');
                const progressText = document.getElementById('deleteProgressText');
                const progressDetails = document.getElementById('deleteProgressDetails');
                const resultsTableBody = document.getElementById('deleteResultsTableBody');
                
                // Show progress, hide results
                deleteProgress.style.display = 'block';
                deleteResults.style.display = 'none';
                resultsTableBody.innerHTML = '';
                
                // Reset progress
                progressFill.style.width = '0%';
                progressText.textContent = '0%';
                
                const results = [];
                let successCount = 0;
                
                for (let i = 0; i < policiesWithAssignments.length; i++) {
                    const policy = policiesWithAssignments[i];
                    const progress = Math.round(((i + 1) / policiesWithAssignments.length) * 100);
                    
                    progressFill.style.width = `${progress}%`;
                    progressText.textContent = `${progress}%`;
                    progressDetails.textContent = `Removing assignments from ${policy.displayName || policy.name}...`;
                    
                    const endpoint = policyEndpoints[policy._type];
                    if (!endpoint || !endpoint.assign) {
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Error',
                            details: 'Unknown policy type'
                        });
                        continue;
                    }
                    
                    try {
                        const url = `${GRAPH_BASE}${endpoint.assign.replace('{id}', policy.id)}`;
                        const response = await makeApiRequest(url, {
                            method: 'POST',
                            body: JSON.stringify({ assignments: [] })
                        });
                        
                        if (response.ok || response.status === 204) {
                            successCount++;
                            results.push({
                                name: policy.displayName || policy.name,
                                status: 'Success',
                                details: 'All assignments removed'
                            });
                        } else {
                            const error = await response.text();
                            results.push({
                                name: policy.displayName || policy.name,
                                status: 'Error',
                                details: error || `HTTP ${response.status}`
                            });
                        }
                    } catch (error) {
                        handleError(`Error unassigning ${policy.displayName}:`, error);
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Error',
                            details: error.message
                        });
                    }
                    
                    if (i < policiesWithAssignments.length - 1) {
                        await new Promise(resolve => setTimeout(resolve, rateLimitDelay));
                    }
                }
                
                // Hide progress, show results
                deleteProgress.style.display = 'none';
                deleteResults.style.display = 'block';
                
                results.forEach(result => {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${result.name}</td>
                        <td class="status-${result.status.toLowerCase()}">${result.status}</td>
                        <td>${result.details}</td>
                    `;
                    resultsTableBody.appendChild(row);
                });
                
                showToast(`Removed assignments from ${successCount}/${policiesWithAssignments.length} policies`, 
                    successCount === policiesWithAssignments.length ? 'success' : 'warning');
                
                if (successCount > 0) {
                    setTimeout(refreshPolicies, 1000);
                }
            };
        }

        // Delete selected policies
        function deleteSelected() {
            const selected = Array.from(selectedPolicies.manage);
            if (selected.length === 0) {
                showToast('No policies selected', 'warning');
                return;
            }
            
            // Show confirmation modal
            const modal = document.getElementById('confirmModal');
            lastFocusedElement = document.activeElement;
            const modalHeader = document.getElementById('modalHeader');
            const modalBody = document.getElementById('modalBody');
            const confirmBtn = document.getElementById('confirmBtn');
            
            modalHeader.textContent = 'Delete Selected Policies?';
            modalBody.textContent = `This will permanently delete ${selected.length} policies. This action cannot be undone!`;
            confirmBtn.className = 'btn btn-danger';
            
            modal.style.display = 'block';
            modal.setAttribute('aria-hidden', 'false');
            
            // Set the action to execute when confirmed
            currentAction = async () => {
                const deleteProgress = document.getElementById('deleteProgress');
                const deleteResults = document.getElementById('deleteResults');
                const progressFill = document.getElementById('deleteProgressFill');
                const progressText = document.getElementById('deleteProgressText');
                const progressDetails = document.getElementById('deleteProgressDetails');
                const resultsTableBody = document.getElementById('deleteResultsTableBody');
                
                // Show progress, hide results
                deleteProgress.style.display = 'block';
                deleteResults.style.display = 'none';
                resultsTableBody.innerHTML = '';
                
                // Reset progress
                progressFill.style.width = '0%';
                progressText.textContent = '0%';
                
                const results = [];
                
                for (let i = 0; i < selected.length; i++) {
                    const policyId = selected[i];
                    const policy = allPolicies.find(p => p.id === policyId);
                    if (!policy) continue;
                    
                    const progress = Math.round(((i + 1) / selected.length) * 100);
                    progressFill.style.width = `${progress}%`;
                    progressText.textContent = `${progress}%`;
                    progressDetails.textContent = `Deleting ${policy.displayName || policy.name}...`;
                    
                    const endpoint = policyEndpoints[policy._type];
                    if (!endpoint || !endpoint.delete) {
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Error',
                            details: 'Unknown policy type'
                        });
                        continue;
                    }
                    
                    try {
                        const url = `${GRAPH_BASE}${endpoint.delete}${policy.id}`;
                        const response = await makeApiRequest(url, { method: 'DELETE' });
                        
                        if (response.ok || response.status === 204) {
                            results.push({
                                name: policy.displayName || policy.name,
                                status: 'Success',
                                details: 'Policy deleted successfully'
                            });
                        } else {
                            const error = await response.text();
                            results.push({
                                name: policy.displayName || policy.name,
                                status: 'Error',
                                details: error || `HTTP ${response.status}`
                            });
                        }
                    } catch (error) {
                        handleError('Delete error:', error);
                        results.push({
                            name: policy.displayName || policy.name,
                            status: 'Error',
                            details: error.message
                        });
                    }
                    
                    if (i < selected.length - 1) {
                        await new Promise(resolve => setTimeout(resolve, rateLimitDelay));
                    }
                }
                
                // Hide progress, show results
                deleteProgress.style.display = 'none';
                deleteResults.style.display = 'block';
                
                results.forEach(result => {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${result.name}</td>
                        <td class="status-${result.status.toLowerCase()}">${result.status}</td>
                        <td>${result.details}</td>
                    `;
                    resultsTableBody.appendChild(row);
                });
                
                const successCount = results.filter(r => r.status === 'Success').length;
                showToast(`Delete complete: ${successCount}/${results.length} successful`, 
                    successCount === results.length ? 'success' : 'warning');
                
                if (successCount > 0) {
                    setTimeout(refreshPolicies, 1000);
                }
            };
        }

        // Close modal
        function closeModal() {
            const modal = document.getElementById('confirmModal');
            modal.style.display = 'none';
            if (lastFocusedElement) lastFocusedElement.focus();
            currentAction = null;
        }
        // Confirm action
        async function confirmAction() {
            const modal = document.getElementById('confirmModal');
            modal.style.display = 'none';
            modal.setAttribute('aria-hidden', 'true');
            if (lastFocusedElement) lastFocusedElement.focus();
            
            if (currentAction && typeof currentAction === 'function') {
                try {
                    await currentAction();
                } catch (error) {
                    handleError('Error executing action:', error);
                    showToast('Error executing action: ' + error.message, 'error');
                }
                currentAction = null;
            }
        }

        // Modal click outside to close
        window.onclick = function(event) {
            const confirmModal = document.getElementById('confirmModal');
            const assignmentModal = document.getElementById('assignmentModal');
            const viewer = document.getElementById('policyViewer');
            
            if (event.target === confirmModal) {
                closeModal();
            }
            if (event.target === assignmentModal) {
                closeAssignmentModal();
            }
            if (event.target === viewer) {
                closePolicyViewer();
            }
        };

        console.log('Intune Policy Manager ready');

        // Expose public methods
        window.IntunePolicyManager = {
            switchTab,
            searchGroups,
            searchManageGroups,
            selectAllUpload,
            deselectAllUpload,
            exportResults,
            clearResults,
            refreshPolicies,
            selectAllManage,
            deselectAllManage,
            exportSelected,
            duplicateSelected,
            assignSelected,
            unassignSelected,
            deleteSelected,
            closeAssignmentModal,
            confirmAssignment,
            closePolicyViewer,
            copyRedirectUri
        };
    })();
