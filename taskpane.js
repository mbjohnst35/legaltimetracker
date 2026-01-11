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
        status.innerText = "v38 Loaded. Checking for Key...";
    }
}, 500);

// Add event listener immediately
document.addEventListener("DOMContentLoaded", () => {
    // Button Event Listeners
    const runBtn = document.getElementById("runButton");
    if (runBtn) runBtn.onclick = startProcess;

    const saveKeyBtn = document.getElementById("saveKeyButton");
    if (saveKeyBtn) saveKeyBtn.onclick = saveApiKey;

    const resetKeyBtn = document.getElementById("resetKeyLink");
    if (resetKeyBtn) resetKeyBtn.onclick = resetApiKey;
    
    // Check if we have a key saved
    checkApiKey();
});

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("startDate").valueAsDate = new Date();
        document.getElementById("endDate").valueAsDate = new Date();
    }
});

// --- KEY MANAGEMENT ---
function checkApiKey() {
    const storedKey = localStorage.getItem("gemini_api_key");
    const loginSection = document.getElementById("key-section");
    const mainSection = document.getElementById("main-section");
    const status = document.getElementById("status");

    if (!storedKey) {
        // No key found, show login
        if (loginSection) loginSection.style.display = "block";
        if (mainSection) mainSection.style.display = "none";
        status.innerText = "Please enter your API Key to begin.";
        status.style.color = "blue";
    } else {
        // Key found, show main app
        if (loginSection) loginSection.style.display = "none";
        if (mainSection) mainSection.style.display = "block";
        status.innerText = "Key found. Initializing AI...";
        discoverGeminiModel(storedKey);
    }
}

function saveApiKey() {
    const input = document.getElementById("apiKeyInput");
    const key = input.value.trim();
    if (!key) {
        // simple alert fallback
        const status = document.getElementById("status");
        status.innerText = "Please paste a valid API Key.";
        status.style.color = "red";
        return;
    }
    // Save to browser storage
    localStorage.setItem("gemini_api_key", key);
    // Reload check
    checkApiKey();
}

function resetApiKey() {
    localStorage.removeItem("gemini_api_key");
    location.reload();
}

// --- ROBUST MODEL DISCOVERY ---
async function discoverGeminiModel(apiKey) {
    const status = document.getElementById("status");
    const listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
    
    // UPDATED: Use 2.5-flash-preview
    const fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=" + apiKey;

    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 3000);

        const response = await fetch(listUrl, { signal: controller.signal });
        clearTimeout(timeoutId);

        if (!response.ok) {
            // If 403/400, the key is probably bad
            if (response.status === 400 || response.status === 403) {
                throw new Error("INVALID_KEY");
            }
            const err = await response.json();
            throw new Error(err.error?.message || response.statusText);
        }

        const data = await response.json();

        if (data.models) {
            let chosenModel = data.models.find(m => m.name.includes("gemini-2.5-flash"));
            
            if (chosenModel) {
                ACTIVE_GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/" + chosenModel.name + ":generateContent?key=" + apiKey;
                status.innerText = "System Ready (v38).";
                status.style.color = "green";
                return;
            }
        }
        throw new Error("Preferred model not found.");

    } catch (e) {
        if (e.message === "INVALID_KEY" || e.message.includes("API key")) {
            status.innerText = "Error: Invalid API Key. Please click 'Change API Key' below.";
            status.style.color = "red";
            // We don't auto-clear here to give user a chance to read the error, 
            // but they can click the reset link.
            return;
        }

        console.warn("Discovery failed. Using Fallback.", e);
        ACTIVE_GEMINI_URL = fallbackUrl;
        status.innerText = "System Ready (v38 - Fallback).";
        status.style.color = "blue"; 
    }
}

async function startProcess() {
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

        // Fetching logic handles the status updates internally now
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

            // 2-second rate limit pause
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
        "&$top=500&$orderby=receivedDateTime