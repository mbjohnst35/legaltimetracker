/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, msal, console, Blob, URL */

// --- CONFIGURATION ---
const CLIENT_ID = "41572571-24e6-44ba-be2c-e3c2b4a0d959"; // Your new ID
const REDIRECT_URI = "[https://mbjohnst35.github.io/taskpane.html](https://mbjohnst35.github.io/taskpane.html)"; // Must match Azure exactly

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Set default dates to today
        document.getElementById("startDate").valueAsDate = new Date();
        document.getElementById("endDate").valueAsDate = new Date();
        document.getElementById("runButton").onclick = startProcess;
    }
});

// Main entry point
async function startProcess() {
    updateStatus("Initializing...", false);
    const button = document.getElementById("runButton");
    button.disabled = true;

    try {
        // 1. Authentication
        const accessToken = await getAccessToken();
        
        // 2. Get User Inputs
        const folder = document.getElementById("folderSelect").value;
        const startDate = new Date(document.getElementById("startDate").value);
        const endDate = new Date(document.getElementById("endDate").value);
        const timeVal = document.getElementById("timeValue").value;

        // Adjust endDate to include the full day (set to 23:59:59)
        endDate.setHours(23, 59, 59, 999);

        // 3. Fetch Emails from Graph
        updateStatus("Fetching emails from " + folder + "...", false);
        const emails = await fetchEmails(accessToken, folder, startDate, endDate);

        if (emails.length === 0) {
            updateStatus("No emails found in that date range.", false);
            button.disabled = false;
            return;
        }

        updateStatus(`Processing ${emails.length} emails...`, false);

        // 4. Process Data
        const reportData = emails.map(email => processEmail(email, timeVal));

        // 5. Generate CSV
        generateCSV(reportData);
        
        updateStatus(`Success! Report generated for ${emails.length} emails.`, true);

    } catch (error) {
        updateStatus("Error: " + error.message, true);
        console.error(error);
    } finally {
        button.disabled = false;
    }
}

// --- AUTHENTICATION ---
async function getAccessToken() {
    const msalConfig = {
        auth: {
            clientId: CLIENT_ID,
            authority: "[https://login.microsoftonline.com/common](https://login.microsoftonline.com/common)",
            redirectUri: REDIRECT_URI,
        },
        cache: { cacheLocation: "localStorage" }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const tokenRequest = { scopes: ["Mail.Read"] };

    try {
        // Try silent first
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            tokenRequest.account = accounts[0];
            const response = await msalInstance.acquireTokenSilent(tokenRequest);
            return response.accessToken;
        } else {
            throw new Error("No account");
        }
    } catch (err) {
        // Fallback to popup
        const response = await msalInstance.acquireTokenPopup(tokenRequest);
        return response.accessToken;
    }
}

// --- GRAPH API ---
async function fetchEmails(token, folder, start, end) {
    // Format dates for OData filter (ISO 8601)
    const startStr = start.toISOString();
    const endStr = end.toISOString();

    // Build Graph Query
    // filter: receivedDateTime between start and end
    // select: only fields we need to save bandwidth
    // top: max 100 per page (we'll do one page for safety, extendable later)
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folder}/messages` +
        `?$filter=receivedDateTime ge ${startStr} and receivedDateTime le ${endStr}` +
        `&$select=receivedDateTime,sender,toRecipients,subject,bodyPreview` +
        `&$top=500` + 
        `&$orderby=receivedDateTime desc`;

    const response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) throw new Error(`Graph API Error: ${response.statusText}`);
    
    const data = await response.json();
    return data.value; // Array of email objects
}

// --- DATA PROCESSING ---
function processEmail(email, timeVal) {
    // Safe extraction of nested properties
    const dateObj = new Date(email.receivedDateTime);
    const dateStr = dateObj.toLocaleDateString();
    const timeStr = dateObj.toLocaleTimeString();
    
    const senderName = email.sender?.emailAddress?.name || "Unknown";
    const senderAddr = email.sender?.emailAddress?.address || "Unknown";
    
    // Map recipients to a single string (e.g., "John; Jane")
    const recipients = email.toRecipients || [];
    const recNames = recipients.map(r => r.emailAddress.name).join("; ");
    const recAddrs = recipients.map(r => r.emailAddress.address).join("; ");

    // "Smart Summary" - Using bodyPreview from Graph (first 255 chars of text)
    // This is free and built-in. 
    let summary = email.bodyPreview || "No content";
    summary = summary.replace(/(\r\n|\n|\r)/gm, " "); // Remove newlines for CSV safety
    if (summary.length > 100) summary = summary.substring(0, 100) + "..."; // Truncate

    return {
        "Date": dateStr,
        "Time": timeStr,
        "Sender Name": senderName,
        "Sender Email": senderAddr,
        "Recipient Name": recNames,
        "Recipient Email": recAddrs,
        "Subject": (email.subject || "").replace(/,/g, " "), // Remove commas
        "Summary": summary,
        "Time Value": timeVal
    };
}

// --- CSV GENERATION ---
function generateCSV(data) {
    if (data.length === 0) return;

    const headers = Object.keys(data[0]);
    const csvRows = [];

    // Add Header Row
    csvRows.push(headers.join(","));

    // Add Data Rows
    for (const row of data) {
        const values = headers.map(header => {
            let val = row[header] || "";
            // Escape double quotes by doubling them
            val = String(val).replace(/"/g, '""'); 
            // Wrap in quotes to handle commas in data
            return `"${val}"`;
        });
        csvRows.push(values.join(","));
    }

    const csvString = csvRows.join("\n");
    const blob = new Blob([csvString], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    
    // Trigger Download
    const a = document.getElementById("downloadLink");
    a.href = url;
    a.download = `Billable_Report_${new Date().getTime()}.csv`;
    a.click();
}

function updateStatus(message, isError) {
    const el = document.getElementById("status");
    el.innerText = message;
    el.style.color = isError ? "red" : "black";
}