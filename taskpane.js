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
const CLIENT_ID = "41572571-24e6-44ba-be2c-e3c2b4a0d959"; 
const REDIRECT_URI = "https://mbjohnst35.github.io/taskpane.html"; 

// --- GEMINI AI CONFIGURATION ---
const GEMINI_API_KEY = "AIzaSyBm0bT3uUpzSjh-Nq8QT8E_6ZSL8cbQ3c0"; 

// STRATEGY: Try the V1 stable endpoint.
const URL_PRIMARY = "https://generativelanguage.googleapis.com/v1/models/gemini-pro:generateContent?key=" + GEMINI_API_KEY;

// 2. IMMEDIATE VISUAL CHECK
setTimeout(() => {
    const status = document.getElementById("status");
    if (status) {
        status.innerText = "System Ready (v19 - Error Cleaner). Waiting for user...";
        status.style.color = "blue";
    }
}, 500);

// Add event listener immediately
document.addEventListener("DOMContentLoaded", () => {
    const runBtn = document.getElementById("runButton");
    if (runBtn) {
        runBtn.onclick = startProcess; 
        console.log("Event listener attached to runButton via DOMContentLoaded");
    }
});

Office.onReady((info) => {
    console.log("Office.onReady called. Host:", info.host);
    if (info.host === Office.HostType.Outlook) {
        const startEl = document.getElementById("startDate");
        const endEl = document.getElementById("endDate");
        if (startEl && endEl) {
            startEl.valueAsDate = new Date();
            endEl.valueAsDate = new Date();
        }
        
        const btn = document.getElementById("runButton");
        if (btn) {
             btn.onclick = startProcess; 
        }
    }
});

async function startProcess() {
    console.log("Button clicked! Starting process...");
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

        updateStatus("Processing " + emails.length + " emails with AI...", false);

        const reportData = [];
        for (const email of emails) {
            const sName = (email.sender && email.sender.emailAddress) ? email.sender.emailAddress.name : "Unknown";
            updateStatus("Summarizing email from " + sName + "...", false);
            
            const processedRow = await processEmailWithAI(email, timeVal);
            reportData.push(processedRow);
        }

        generateCSV(reportData);
        updateStatus("Success! AI Report generated for " + emails.length + " emails.", true);

    } catch (error) {
        updateStatus("Error: " + error.message, true);
        console.error(error);
    } finally {
        button.disabled = false;
    }
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

    if (typeof msal === 'undefined') {
        throw new Error("MSAL library not loaded. Check internet connection.");
    }

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

async function processEmailWithAI(email, timeVal) {
    const dateObj = new Date(email.receivedDateTime);
    
    let senderName = "Unknown";
    let senderAddr = "Unknown";
    if (email.sender && email.sender.emailAddress) {
        senderName = email.sender.emailAddress.name || "Unknown";
        senderAddr = email.sender.emailAddress.address || "Unknown";
    }

    const recipients = email.toRecipients || [];
    const recNames = recipients.map(function(r) { return r.emailAddress.name; }).join("; ");
    const recAddrs = recipients.map(function(r) { return r.emailAddress.address; }).join("; ");
    
    const subjectSafe = (email.subject || "").replace(/,/g, " ");
    const bodyPreviewSafe = (email.bodyPreview || "");

    const emailText = "Subject: " + subjectSafe + "\nBody: " + bodyPreviewSafe;
    let summary = "No content";

    try {
        // CALL THE ROBUST API FUNCTION
        summary = await callGeminiWithFallback(emailText);
    } catch (e) {
        console.error("AI Error:", e);
        summary = "AI Error: " + e.message;
    }

    // Clean summary quotes for CSV safety
    summary = summary.replace(/"/g, "'");

    return {
        "Date": dateObj.toLocaleDateString(),
        "Time": dateObj.toLocaleTimeString(),
        "Sender Name": senderName,
        "Sender Email": senderAddr,
        "Recipient Name": recNames,
        "Recipient Email": recAddrs,
        "Subject": subjectSafe,
        "Summary": summary,
        "Time Value": timeVal
    };
}

// --- ROBUST AI FUNCTION ---
async function callGeminiWithFallback(text) {
    return await tryGeminiEndpoint(URL_PRIMARY, text);
}

async function tryGeminiEndpoint(url, text) {
    const prompt = "Summarize the following email in exactly one concise sentence for a legal billing report:\n\n" + text;
    const payload = {
        contents: [{ parts: [{ text: prompt }] }]
    };

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            // DIAGNOSTIC CHANGE: Flatten the JSON error so it fits in a CSV cell
            const errorText = await response.text();
            // Remove newlines and quotes to prevent CSV breakage
            const cleanError = errorText.replace(/(\r\n|\n|\r)/gm, " ").replace(/"/g, "'");
            return "API FAIL: " + response.status + " MSG: " + cleanError.substring(0, 150);
        }

        const data = await response.json();
        
        if (data.candidates && data.candidates.length > 0 && 
            data.candidates[0].content && 
            data.candidates[0].content.parts && 
            data.candidates[0].content.parts.length > 0) {
             return data.candidates[0].content.parts[0].text;
        }
        return "Error: Empty AI Response";

    } catch (networkError) {
        return "Network Error: " + networkError.message;
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
        el.innerText = "v19: " + message; 
        el.style.color = isError ? "red" : "black";
    }
}