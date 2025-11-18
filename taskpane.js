// --- 1. CONFIGURATION ---

// Client ID from your Azure App Registration
const CLIENT_ID = "41572571-24e6-44ba-be2c-e3c2b4a0d959"; 
const GRAPH_SCOPES = ["Mail.Read", "User.Read"];
const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0";

// MSAL configuration object
const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
        // The redirectUri will be updated later with your public URL (e.g., Azure SWA URL)
        redirectUri: window.location.origin + "/taskpane.html" 
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// --- 2. OFFICE.JS INITIALIZATION ---

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("scanButton").onclick = scanEmail;
        updateStatus("Office.js is ready. Click 'Scan Email'.");
    }
});

function updateStatus(message) {
    document.getElementById("status").innerText = message;
}

// --- 3. AUTHENTICATION & TOKEN ACQUISITION ---

async function getToken() {
    try {
        // Attempt silent token acquisition
        const silentRequest = {
            scopes: GRAPH_SCOPES,
            account: msalInstance.getAllAccounts()[0]
        };
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        // If silent fails, fall back to interactive pop-up
        updateStatus("Authentication required. Opening pop-up...");
        const loginRequest = {
            scopes: GRAPH_SCOPES
        };
        const response = await msalInstance.loginPopup(loginRequest);
        
        // After successful login, try silent acquisition again
        const silentRequest = {
            scopes: GRAPH_SCOPES,
            account: response.account
        };
        const tokenResponse = await msalInstance.acquireTokenSilent(silentRequest);
        return tokenResponse.accessToken;
    }
}

// --- 4. CORE SCANNING FUNCTION ---

async function scanEmail() {
    updateStatus("Starting scan...");
    
    // Get the ID of the currently selected message
    Office.context.mailbox.item.getItemIdAsync(async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const itemId = result.value;
            
            try {
                const accessToken = await getToken();
                updateStatus("Successfully authenticated. Fetching email content...");

                // Call Microsoft Graph API to get the email body
                const graphUrl = `${GRAPH_ENDPOINT}/me/messages/${itemId}?$select=body,subject`;
                const response = await fetch(graphUrl, {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`
                    }
                });

                if (!response.ok) {
                    throw new Error(`Graph API call failed: ${response.statusText}`);
                }

                const emailData = await response.json();
                const bodyContent = emailData.body.content || "";
                
                // Process the content
                runScanningLogic(bodyContent);

            } catch (error) {
                console.error("Error during authentication or Graph API call:", error);
                updateStatus(`ERROR: Could not scan email. ${error.message}`);
            }

        } else {
            updateStatus("ERROR: Could not retrieve email ID.");
        }
    });
}

// --- 5. BUSINESS LOGIC (YOUR BILLABLE TIME SCANNER) ---

function runScanningLogic(text) {
    const textLower = text.toLowerCase();
    let resultMessage = "Scan Complete: No billable time keywords found (6 or 18 minutes).";
    let found = false;

    // Define keywords for 6 minutes (0.1 billable hours) and 18 minutes (0.3 billable hours)
    const keywords = [
        "6 minutes",
        "0.1 hr",
        "0.1 hours",
        "18 minutes",
        "0.3 hr",
        "0.3 hours"
    ];

    for (const keyword of keywords) {
        if (textLower.includes(keyword)) {
            resultMessage = `Scan Complete: Found billable time keyword: "${keyword}"`;
            found = true;
            break; 
        }
    }

    if (!found) {
        resultMessage += "\n\nFeel free to forward this email for a time entry.";
    }

    updateStatus(resultMessage);
}