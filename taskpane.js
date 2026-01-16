/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, msal, console, Blob, URL, window, localStorage */

// 1. GLOBAL ERROR HANDLER
window.onerror = function(message, source, lineno, colno, error) {
    const status = document.getElementById("status");
    if (status) {
        status.innerText = "CRITICAL JS ERROR: " + message + "\nLine: " + lineno;
        status.style.color = "red";
    }
    console.error("Global Error:", message, error);
};

// --- CONFIGURATION ---
const CLIENT_ID = "41572571-24e6-44ba-be2c-e3c2b4a0d959"; 
const REDIRECT_URI = "https://mbjohnst35.github.io/taskpane.html"; 

let ACTIVE_GEMINI_URL = ""; 

// 2. IMMEDIATE VISUAL CHECK
setTimeout(() => {
    const status = document.getElementById("status");
    if (status && status.innerText.includes("Loading")) {
        status.innerText = "v39 Loaded. Waiting for Office...";
    }
}, 500);

// 3. MAIN INITIALIZATION
Office.onReady((info) => {
    // Set up date defaults if in Outlook
    if (info.host === Office.HostType.Outlook) {
        const start = document.getElementById("startDate");
        const end = document.getElementById("endDate");
        if (start && end) {
            start.valueAsDate = new Date();
            end.valueAsDate = new Date();
        }
    }

    // Attach Button Listeners
    const runBtn = document.getElementById("runButton");
    if (runBtn) runBtn.onclick = startProcess;

    const saveKeyBtn = document.getElementById("saveKeyButton");
    if (saveKeyBtn) saveKeyBtn.onclick = saveApiKey;

    const resetKeyBtn = document.getElementById("resetKeyLink");
    if (resetKeyBtn) resetKeyBtn.onclick = resetApiKey;

    // CHECK FOR KEY
    checkApiKey();
});

// --- KEY MANAGEMENT ---
function checkApiKey() {
    const status = document.getElementById("status");
    let storedKey = null;

    try {
        storedKey = localStorage.getItem("gemini_api_key");
    } catch (e) {
        console.warn("LocalStorage access denied:", e);
        // If storage fails, we just stay on the login screen
    }

    const loginSection = document.getElementById("key-section");
    const mainSection = document.getElementById("main-section");

    if (storedKey) {
        // We have a key! Hide login, Show Main.
        if (loginSection) loginSection.style.display = "none";
        if (mainSection) mainSection.style.display = "block";
        
        status.innerText = "Key found. System Ready (v39).";
        status.style.color = "green";
        
        // Pre-validate the key model
        discoverGeminiModel(storedKey);
    } else {
        // No key. Ensure Login is visible (it is default, but just in case)
        if (loginSection) loginSection.style.display = "block";
        if (mainSection) mainSection.style.display = "none";
        
        status.innerText = "Setup Required: Please enter API Key.";
        status.style.color = "blue";
    }
}

function saveApiKey() {
    const input = document.getElementById("apiKeyInput");
    const key = input.value.trim();
    if (!key) {
        alert("Please paste a valid API Key.");
        return;
    }
    try {
        localStorage.setItem("gemini_api_key", key);
        checkApiKey(); // Reload UI
    } catch (e) {
        alert("Error saving key: " + e.message);
    }
}

function resetApiKey() {
    try {
        localStorage.removeItem("gemini_api_key");
        location.reload();
    } catch (e) {
        console.error(e);
    }
}

// --- ROBUST MODEL DISCOVERY ---
async function discoverGeminiModel(apiKey) {
    const status = document.getElementById("status");
    const listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
    const fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=" + apiKey;

    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 3000);

        const response = await fetch(listUrl, { signal: controller.signal });
        clearTimeout(timeoutId);

        if (!response.ok) {
            if (response.status === 400 || response.status === 403) throw new Error("INVALID_KEY");
            const err = await response.json();
            throw new Error(err.error?.message || response.statusText);
        }

        const data = await response.json();

        if (data.models) {
            let chosenModel = data.models.find(m => m.name.includes("gemini-2.5-flash"));
            if (chosenModel) {
                ACTIVE_GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/" + chosenModel.name + ":generateContent?key=" + apiKey;
                status.innerText = "System Ready (v39 - Turbo).";
                return;
            }
        }
        throw new Error("Preferred model not found.");

    } catch (e) {
        if (e.message === "INVALID_KEY") {
            status.innerText = "Error: Invalid API Key. Please reset.";
            status.style.color = "red";
            return;
        }
        ACTIVE_GEMINI_URL = fallbackUrl;
        console.warn("Discovery failed, using fallback.");
    }
}

async function startProcess() {
    // Get the key again for execution
    const apiKey = localStorage.getItem("gemini_api_key");
    if (!apiKey) {
        checkApiKey(); // Kick back to login screen
        return;
    }

    updateStatus("Initializing...", false);
    const button = document.getElementById("runButton");
    button.disabled = true;

    try {
        const accessToken = await getAccessToken();
        const folder = document.getElementById("folderSelect").value;
        const startInput = document.getElementById("startDate").value;
        const endInput = document.getElementById("endDate").value;
        
        if (!startInput || !endInput) throw new Error("Please select both start and end dates.");

        const startDate = new Date(startInput);
        const endDate = new Date(endInput);
        const timeVal = document.getElementById("timeValue").value;
        endDate.setHours(23, 59, 59, 999);

        const emails = await fetchEmails(accessToken, folder, startDate, endDate);

        if (emails.length === 0) {
            updateStatus("No emails found in that date range.", false);
            button.disabled = false;
            return;
        }

        updateStatus("Found " + emails.length + " emails. Starting Processing...", false);

        const reportData = [];
        const BATCH_SIZE = 10; 
        
        for (let i = 0; i < emails.length; i += BATCH_SIZE) {
            const chunk = emails.slice(i, i + BATCH_SIZE);
            const currentCount = Math.min(i + BATCH_SIZE, emails.length);
            
            updateStatus("Processing batch " + currentCount + "/" + emails.length + "...", false);
            
            const summarizedChunk = await processBatchWithAI(chunk, timeVal);
            reportData.push(...summarizedChunk);

            if (i + BATCH_SIZE < emails.length) {
                await new Promise(resolve => setTimeout(resolve, 2000));
            }
        }

        generateCSV(reportData);
        updateStatus("Success! Report generated for " + emails.length + " emails.", true);

    } catch (error) {
        updateStatus("Error: " + error.message, true);
        console.error(error);
    } finally {
        button.disabled = false;
    }
}

// --- PAGINATION FIX ---
async function fetchEmails(token, folder, start, end) {
    const startStr = start.toISOString();
    const endStr = end.toISOString();
    
    let url = "https://graph.microsoft.com/v1.0/me/mailFolders/" + folder + "/messages" +
        "?$filter=receivedDateTime ge " + startStr + " and receivedDateTime le " + endStr +
        "&$select=receivedDateTime,sender,toRecipients,subject,bodyPreview" +
        "&$top=500&$orderby=receivedDateTime desc"; 

    let allMessages = [];
    
    while (url) {
        updateStatus("Fetching emails... (Count: " + allMessages.length + ")", false);
        
        const response = await fetch(url, { headers: { Authorization: "Bearer " + token } });
        if (!response.ok) throw new Error("Graph API Error: " + response.statusText);
        
        const data = await response.json();
        
        if (data.value) {
            allMessages = allMessages.concat(data.value);
        }
        
        url = data["@odata.nextLink"]; 
    }
    
    return allMessages;
}

// --- BATCH AI FUNCTION ---
async function processBatchWithAI(emailBatch, timeVal) {
    let prompt = "Summarize the action or content of each email below in one concise sentence for a legal billing time entry. Do not use phrases like 'This email discusses' or 'The sender'. Start directly with the verb (e.g., 'Reviewed', 'Discussed', 'Sent'). Return the result as a JSON object where the key is the EmailID and the value is the summary.\n\n";
    
    emailBatch.forEach((email, index) => {
        const subject = (email.subject || "No Subject").replace(/(\r\n|\n|\r)/gm, " ");
        const body = (email.bodyPreview || "No Content").replace(/(\r\n|\n|\r)/gm, " ");
        prompt += `EmailID "${index}":\nSubject: ${subject}\nBody: ${body}\n\n`;
    });

    let summaries = {};

    try {
        if (!ACTIVE_GEMINI_URL) throw new Error("AI not initialized");

        const payload = {
            contents: [{ parts: [{ text: prompt }] }]
        };

        const response = await fetch(ACTIVE_GEMINI_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`API Error ${response.status}: ${errText}`);
        }

        const data = await response.json();
        const textResponse = data.candidates?.[0]?.content?.parts?.[0]?.text;
        
        const cleanJson = textResponse.replace(/```json/g, "").replace(/```/g, "").trim();
        summaries = JSON.parse(cleanJson);

    } catch (e) {
        console.error("Batch Failed:", e);
        emailBatch.forEach((_, index) => {
            summaries[index] = "Review email regarding " + (emailBatch[index].subject || "subject");
        });
    }

    return emailBatch.map((email, index) => {
        const dateObj = new Date(email.receivedDateTime);
        const senderName = email.sender?.emailAddress?.name || "Unknown";
        const senderAddr = email.sender?.emailAddress?.address || "Unknown";
        const recipients = (email.toRecipients || []).map(r => r.emailAddress.name).join("; ");
        
        let summary = summaries[index.toString()] || "Error: Summary missing";
        summary = summary.replace(/"/g, "'"); 

        return {
            "Date": dateObj.toLocaleDateString(),
            "Time": dateObj.toLocaleTimeString(),
            "Sender Name": senderName,
            "Sender Email": senderAddr,
            "Recipient Name": recipients,
            "Subject": (email.subject || "").replace(/,/g, " "),
            "Summary": summary,
            "Time Value": timeVal
        };
    });
}

async function getAccessToken() {
    const msalConfig = {
        auth: {
            clientId: CLIENT_ID,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: REDIRECT_URI,
        },
        cache: { cacheLocation: "localStorage" }
    };

    if (typeof msal === 'undefined') throw new Error("MSAL not loaded");

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const tokenRequest = { scopes: ["Mail.Read"] };

    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            tokenRequest.account = accounts[0];
            const response = await msalInstance.acquireTokenSilent(tokenRequest);
            return response.accessToken;
        } else {
            throw new Error("No account");
        }
    } catch (err) {
        const response = await msalInstance.acquireTokenPopup(tokenRequest);
        return response.accessToken;
    }
}

function generateCSV(data) {
    if (data.length === 0) return;
    const headers = Object.keys(data[0]);
    const csvRows = [];
    csvRows.push(headers.join(","));
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const values = headers.map(function(header) {
            let val = row[header] || "";
            val = String(val).replace(/"/g, '""'); 
            return '"' + val + '"';
        });
        csvRows.push(values.join(","));
    }
    const csvString = csvRows.join("\n");
    const blob = new Blob([csvString], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.getElementById("downloadLink");
    a.href = url;
    a.download = "Billable_AI_Report_" + new Date().getTime() + ".csv";
    a.click();
}

function updateStatus(message, isError) {
    const el = document.getElementById("status");
    if (el) {
        el.innerText = "v39: " + message; 
        el.style.color = isError ? "red" : "black";
    }
}