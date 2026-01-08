<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Assistant</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f3f4f6;
        }
        .chat-container {
            height: calc(100vh - 180px);
        }
        .message-bubble {
            max-width: 85%;
            overflow-wrap: break-word;
        }
        .typing-indicator span {
            animation: blink 1.4s infinite both;
        }
        .typing-indicator span:nth-child(2) { animation-delay: 0.2s; }
        .typing-indicator span:nth-child(3) { animation-delay: 0.4s; }
        @keyframes blink {
            0% { opacity: 0.2; }
            20% { opacity: 1; }
            100% { opacity: 0.2; }
        }
        /* Custom scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f1f1;
        }
        ::-webkit-scrollbar-thumb {
            background: #c1c1c1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #a8a8a8;
        }
        pre {
            background-color: #1e293b;
            color: #e2e8f0;
            padding: 1rem;
            border-radius: 0.5rem;
            overflow-x: auto;
            margin-top: 0.5rem;
            margin-bottom: 0.5rem;
        }
    </style>
</head>
<body class="bg-gray-100 h-screen flex flex-col items-center">

    <!-- Header -->
    <header class="w-full bg-white shadow-sm py-4 px-6 flex justify-between items-center z-10">
        <div class="flex items-center gap-3">
            <div class="bg-blue-600 text-white p-2 rounded-lg">
                <i class="fa-solid fa-robot text-xl"></i>
            </div>
            <h1 class="text-xl font-bold text-gray-800">AI Assistant</h1>
        </div>
        <div class="text-sm text-gray-500">
            Model: gemini-2.5-flash
        </div>
    </header>

    <!-- Main Chat Area -->
    <main class="flex-1 w-full max-w-4xl p-4 flex flex-col relative">
        
        <!-- Chat History -->
        <div id="chat-history" class="chat-container flex-1 overflow-y-auto bg-white rounded-t-xl shadow-sm p-6 space-y-6 scroll-smooth">
            
            <!-- Welcome Message -->
            <div class="flex gap-4">
                <div class="w-8 h-8 rounded-full bg-blue-600 flex items-center justify-center flex-shrink-0 text-white">
                    <i class="fa-solid fa-robot text-xs"></i>
                </div>
                <div class="message-bubble bg-gray-100 text-gray-800 rounded-2xl rounded-tl-none p-4 shadow-sm">
                    <p>Hello! I'm your AI assistant. How can I help you today?</p>
                </div>
            </div>

        </div>

        <!-- Input Area -->
        <div class="bg-white p-4 rounded-b-xl shadow-lg border-t border-gray-100">
            <div class="relative flex items-end gap-2 bg-gray-50 rounded-xl border border-gray-200 p-2 focus-within:ring-2 focus-within:ring-blue-500 focus-within:border-transparent transition-all">
                <textarea 
                    id="user-input" 
                    rows="1" 
                    placeholder="Type your message..." 
                    class="w-full bg-transparent border-none focus:ring-0 resize-none py-3 px-2 max-h-32 text-gray-700"
                    onkeydown="handleEnter(event)"></textarea>
                
                <button 
                    onclick="sendMessage()" 
                    id="send-btn"
                    class="bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-lg transition-colors flex-shrink-0 disabled:opacity-50 disabled:cursor-not-allowed">
                    <i class="fa-solid fa-paper-plane"></i>
                </button>
            </div>
            <div class="text-xs text-center text-gray-400 mt-2">
                Press Enter to send, Shift+Enter for new line
            </div>
        </div>

    </main>

    <script>
        // --- API CONFIGURATION ---
        const apiKey = ""; // System provided key (if available in env)
        const manualKey = "AIzaSyDmKO808anbshOud4t51XY2ueOxDtN9IQc"; // User provided key
        
        // Use manual key if system key is empty
        const effectiveKey = apiKey || manualKey;

        // Correct Model Endpoint to prevent 404s
        const MODEL_NAME = "gemini-2.5-flash-preview-09-2025";
        const API_URL = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent`;

        // State
        let isGenerating = false;
        const chatHistory = document.getElementById('chat-history');
        const userInput = document.getElementById('user-input');
        const sendBtn = document.getElementById('send-btn');

        // Auto-resize textarea
        userInput.addEventListener('input', function() {
            this.style.height = 'auto';
            this.style.height = (this.scrollHeight) + 'px';
            if(this.value === '') this.style.height = 'auto';
        });

        function handleEnter(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        }

        function scrollToBottom() {
            chatHistory.scrollTop = chatHistory.scrollHeight;
        }

        function appendMessage(role, text, isError = false) {
            const div = document.createElement('div');
            div.className = `flex gap-4 ${role === 'user' ? 'flex-row-reverse' : ''}`;
            
            const avatar = document.createElement('div');
            avatar.className = `w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 text-white ${role === 'user' ? 'bg-gray-700' : (isError ? 'bg-red-500' : 'bg-blue-600')}`;
            avatar.innerHTML = role === 'user' ? '<i class="fa-solid fa-user text-xs"></i>' : (isError ? '<i class="fa-solid fa-exclamation text-xs"></i>' : '<i class="fa-solid fa-robot text-xs"></i>');

            const bubble = document.createElement('div');
            bubble.className = `message-bubble p-4 shadow-sm text-sm md:text-base ${
                role === 'user' 
                ? 'bg-blue-600 text-white rounded-2xl rounded-tr-none' 
                : (isError ? 'bg-red-50 text-red-800 border border-red-200 rounded-2xl rounded-tl-none' : 'bg-gray-100 text-gray-800 rounded-2xl rounded-tl-none')
            }`;
            
            // Parse Markdown for AI responses
            if (role === 'model' && !isError) {
                bubble.innerHTML = marked.parse(text);
            } else {
                bubble.textContent = text;
            }

            div.appendChild(avatar);
            div.appendChild(bubble);
            chatHistory.appendChild(div);
            scrollToBottom();
            return bubble;
        }

        function showTypingIndicator() {
            const div = document.createElement('div');
            div.id = 'typing-indicator';
            div.className = 'flex gap-4';
            div.innerHTML = `
                <div class="w-8 h-8 rounded-full bg-blue-600 flex items-center justify-center flex-shrink-0 text-white">
                    <i class="fa-solid fa-robot text-xs"></i>
                </div>
                <div class="message-bubble bg-gray-100 text-gray-800 rounded-2xl rounded-tl-none p-4 shadow-sm flex items-center gap-1 typing-indicator">
                    <span class="w-2 h-2 bg-gray-400 rounded-full"></span>
                    <span class="w-2 h-2 bg-gray-400 rounded-full"></span>
                    <span class="w-2 h-2 bg-gray-400 rounded-full"></span>
                </div>
            `;
            chatHistory.appendChild(div);
            scrollToBottom();
            return div;
        }

        function removeTypingIndicator() {
            const indicator = document.getElementById('typing-indicator');
            if (indicator) indicator.remove();
        }

        async function sendMessage() {
            const text = userInput.value.trim();
            if (!text || isGenerating) return;

            // UI Updates
            userInput.value = '';
            userInput.style.height = 'auto';
            isGenerating = true;
            sendBtn.disabled = true;
            appendMessage('user', text);
            showTypingIndicator();

            try {
                if (!effectiveKey) {
                    throw new Error("No API Key provided. Please check the code configuration.");
                }

                const response = await fetch(`${API_URL}?key=${effectiveKey}`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        contents: [{
                            parts: [{ text: text }]
                        }]
                    })
                });

                if (!response.ok) {
                    const errData = await response.json();
                    console.error("API Error Details:", errData);
                    
                    // Handle specifically the 404 which is common with model version issues
                    if (response.status === 404) {
                        throw new Error("Error 404: The AI Model endpoint was not found. This usually means the model version is deprecated or the URL is malformed.");
                    }
                    
                    throw new Error(`API Error ${response.status}: ${errData.error?.message || response.statusText}`);
                }

                const data = await response.json();
                const aiText = data.candidates?.[0]?.content?.parts?.[0]?.text;

                if (aiText) {
                    removeTypingIndicator();
                    appendMessage('model', aiText);
                } else {
                    throw new Error("Received an empty response from the AI.");
                }

            } catch (error) {
                console.error(error);
                removeTypingIndicator();
                appendMessage('model', `Error: ${error.message}`, true);
            } finally {
                isGenerating = false;
                sendBtn.disabled = false;
                userInput.focus();
            }
        }
    </script>
</body>
</html>