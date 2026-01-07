/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, msal, console, Blob, URL */

// --- CONFIGURATION ---
const CLIENT_ID = "41572571-24e6-44ba-be2c-e3c2b4a0d959"; 
const REDIRECT_URI = "https://mbjohnst35.github.io/taskpane.html"; 

// --- GEMINI AI CONFIGURATION ---
// PASTE YOUR KEY HERE
const GEMINI_API_KEY = "AIzaSyBm0bT3uUpzSjh-Nq8QT8E_6ZSL8cbQ3c0"; 
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("startDate").valueAsDate = new Date();
        document.getElementById("endDate").valueAsDate = new Date();
        document.getElementById("runButton").onclick = startProcess;
    }
});

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

        updateStatus(`Processing ${emails.length} emails with AI... (This may take a moment)`, false);

        // Process emails one by one to handle async AI calls
        const reportData = [];
        for (const email of emails) {
            updateStatus(`Summarizing email from ${email.sender?.emailAddress?.name || 'Unknown'}...`, false);
            const processedRow = await processEmailWithAI(email, timeVal);
            reportData.push(processedRow);
        }

        generateCSV(reportData);
        updateStatus(`Success! AI Report generated for ${emails.length} emails.`, true);

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
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folder}/messages` +
        `?$filter=receivedDateTime ge ${startStr} and receivedDateTime le ${endStr}` +
        `&$select=receivedDateTime,sender,toRecipients,subject,bodyPreview` +
        `&$top=500&$orderby=receivedDateTime desc`;

    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!response.ok) throw new Error(`Graph API Error: ${response.statusText}`);
    const data = await response.json();
    return data.value;
}

// --- NEW AI FUNCTION ---
async function processEmailWithAI(email, timeVal) {
    const dateObj = new Date(email.receivedDateTime);
    const senderName = email.sender?.emailAddress?.name || "Unknown";
    const senderAddr = email.sender?.emailAddress?.address || "Unknown";
    const recipients = email.toRecipients || [];
    const recNames = recipients.map(r => r.emailAddress.name).join("; ");
    const recAddrs = recipients.map(r => r.emailAddress.address).join("; ");
    
    // Get text to summarize (Subject + Body Preview)
    const emailText = `Subject: ${email.subject || ""}\nBody: ${email.bodyPreview || ""}`;
    let summary = "No content";

    try {
        // Call Gemini API
        summary = await callGeminiAPI(emailText);
    } catch (e) {
        console.error("AI Error:", e);
        summary = "AI Error: " + e.message;
    }

    return {
        "Date": dateObj.toLocaleDateString(),
        "Time": dateObj.toLocaleTimeString(),
        "Sender Name": senderName,
        "Sender Email": senderAddr,
        "Recipient Name": recNames,
        "Recipient Email": recAddrs,
        "Subject": (email.subject || "").replace(/,/g, " "),
        "Summary": summary.replace(/"/g, "'"), // Clean quotes for CSV
        "Time Value": timeVal
    };
}

async function callGeminiAPI(text) {
    const prompt = `Summarize the following email in exactly one concise sentence for a legal billing report:\n\n${text}`;
    
    const payload = {
        contents: [{
            parts: [{ text: prompt }]
        }]
    };

    const response = await fetch(GEMINI_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });

    if (!response.ok) return "Error summarizing";

    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || "No summary generated";
}

function generateCSV(data) {
    if (data.length === 0) return;
    const headers = Object.keys(data[0]);
    const csvRows = [];
    csvRows.push(headers.join(","));
    for (const row of data) {
        const values = headers.map(header => {
            let val = row[header] || "";
            val = String(val).replace(/"/g, '""'); 
            return `"${val}"`;
        });
        csvRows.push(values.join(","));
    }
    const csvString = csvRows.join("\n");
    const blob = new Blob([csvString], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.getElementById("downloadLink");
    a.href = url;
    a.download = `Billable_AI_Report_${new Date().getTime()}.csv`;
    a.click();
}

function updateStatus(message, isError) {
    const el = document.getElementById("status");
    el.innerText = message;
    el.style.color = isError ? "red" : "black";
}