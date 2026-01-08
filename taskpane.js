/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, msal, console, Blob, URL, window */

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
const CLIENT_ID = "AIzaSyDmKO808anbshOud4t51XY2ueOxDtN9IQc"; 
const REDIRECT_URI = "https://mbjohnst35.github.io/taskpane.html"; 

// --- GEMINI AI CONFIGURATION ---
// YOUR EXISTING KEY (Now upgraded via billing)
const GEMINI_API_KEY = "AIzaSyAPcZ6DzH17KAO9JzC_yB4hDxr_UbTcS9Q"; 

let ACTIVE_GEMINI_URL = ""; 

// 2. IMMEDIATE VISUAL CHECK
setTimeout(() => {
    const status = document.getElementById("status");
    if (status && status.innerText.includes("Loading")) {
        status.innerText = "v25 Loaded. Starting Discovery...";
    }
}, 500);

// Add event listener immediately
document.addEventListener("DOMContentLoaded", () => {
    const runBtn = document.getElementById("runButton");
    if (runBtn) runBtn.onclick = startProcess;
    
    discoverGeminiModel();
});

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("startDate").valueAsDate = new Date();
        document.getElementById("endDate").valueAsDate = new Date();
        const btn = document.getElementById("runButton");
        if (btn) btn.onclick = startProcess; 
    }
});

// --- ROBUST MODEL DISCOVERY ---
async function discoverGeminiModel() {
    const status = document.getElementById("status");
    const listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + GEMINI_API_KEY;
    const fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY;

    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 3000);

        const response = await fetch(listUrl, { signal: controller.signal });
        clearTimeout(timeoutId);

        const data = await response.json();

        if (data.models) {
            let chosenModel = data.models.find(m => m.name.includes("gemini-1.5-flash"));
            
            if (chosenModel) {
                ACTIVE_GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/" + chosenModel.name + ":generateContent?key=" + GEMINI_API_KEY;
                status.innerText = "System Ready (v25 - Turbo Mode). Model: " + chosenModel.displayName;
                status.style.color = "green";
                return;
            }
        }
        throw new Error("No compatible models found.");

    } catch (e) {
        console.warn("Discovery failed. Using fallback.", e);
        ACTIVE_GEMINI_URL = fallbackUrl;
        status.innerText = "System Ready (v25 - Fallback Mode).";
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

        updateStatus("Fetching emails from " + folder + "...", false);
        const emails = await fetchEmails(accessToken, folder, startDate, endDate);

        if (emails.length === 0) {
            updateStatus("No emails found in that date range.", false);
            button.disabled = false;
            return;
        }

        updateStatus("Found " + emails.length + " emails. Starting Turbo Processing...", false);

        // --- BATCH LOGIC (NO DELAY) ---
        const reportData = [];
        const BATCH_SIZE = 10; // Increased batch size for speed
        
        // We will fire all batches almost simultaneously
        for (let i = 0; i < emails.length; i += BATCH_SIZE) {
            const chunk = emails.slice(i, i + BATCH_SIZE);
            const currentCount = Math.min(i + BATCH_SIZE, emails.length);
            
            updateStatus("Processing batch " + currentCount + "/" + emails.length + "...", false);
            
            // Send chunk to AI
            const summarizedChunk = await processBatchWithAI(chunk, timeVal);
            reportData.push(...summarizedChunk);

            // REMOVED: await delay(4000); -> No waiting needed for paid keys!
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

// --- BATCH AI FUNCTION ---
async function processBatchWithAI(emailBatch, timeVal) {
    let prompt = "Summarize each of the following emails in exactly one concise sentence for a legal billing report. Return the result as a JSON object where the key is the EmailID and the value is the summary.\n\n";
    
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
            throw new Error(`API Error ${response.status}`);
        }

        const data = await response.json();
        const textResponse = data.candidates?.[0]?.content?.parts?.[0]?.text;
        
        const cleanJson = textResponse.replace(/```json/g, "").replace(/```/g, "").trim();
        summaries = JSON.parse(cleanJson);

    } catch (e) {
        console.error("Batch Failed:", e);
        emailBatch.forEach((_, index) => {
            summaries[index] = "Error: " + e.message;
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

async function fetchEmails(token, folder, start, end) {
    const startStr = start.toISOString();
    const endStr = end.toISOString();
    const url = "https://graph.microsoft.com/v1.0/me/mailFolders/" + folder + "/messages" +
        "?$filter=receivedDateTime ge " + startStr + " and receivedDateTime le " + endStr +
        "&$select=receivedDateTime,sender,toRecipients,subject,bodyPreview" +
        "&$top=500&$orderby=receivedDateTime desc";

    const response = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!response.ok) throw new Error("Graph API Error: " + response.statusText);
    const data = await response.json();
    return data.value;
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
        el.innerText = "v25: " + message; 
        el.style.color = isError ? "red" : "black";
    }
}