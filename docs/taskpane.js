// Word AI Assistant - Multi-Provider Support with Writing Tools
// Supports: OpenAI, Google Gemini, Anthropic Claude, Local models
// Privacy: All data goes directly to your chosen AI provider - no external storage

const STORAGE_KEYS = {
  provider: 'word-ai-provider',
  apiKey: 'word-ai-api-key',
  localUrl: 'word-ai-local-url',
  localModel: 'word-ai-local-model',
  customPrompt: 'word-ai-custom-prompt',
  writingStyle: 'word-ai-writing-style'
};

// Get Word API reference
const getWord = () => window.Word || window.Office?.Word;

// ============== Quick Action Prompts ==============

const QUICK_ACTIONS = {
  analyze: `Please analyze the writing style of this document. Look at:
1. Tone (formal, casual, professional, academic, etc.)
2. Sentence structure and complexity
3. Vocabulary level
4. Voice (active vs passive)
5. Any patterns or habits

After analyzing, remember this style so you can mimic it when making future edits. Give me a summary of the writing style you detected.`,

  grammar: `Please check the entire document for grammar errors. For each error found:
1. Highlight the problematic text in yellow
2. Add a comment explaining the grammar issue and the correction

After checking, give me a summary of how many issues were found.`,

  spelling: `Please check the entire document for spelling errors and typos. For each error found:
1. Highlight the misspelled word in red
2. Add a comment with the correct spelling

After checking, give me a summary of what was found.`,

  formal: `Please rewrite the document content to use a formal, professional tone. Make these changes:
- Replace casual language with formal alternatives
- Use complete sentences
- Avoid contractions
- Use professional vocabulary
- Maintain a respectful, businesslike tone

Make the edits directly to the document.`,

  casual: `Please rewrite the document content to use a casual, conversational tone. Make these changes:
- Use contractions where natural
- Simplify complex sentences
- Use everyday vocabulary
- Make it sound like a friendly conversation
- Keep it approachable and relaxed

Make the edits directly to the document.`,

  professional: `Please adjust the document to have a professional business tone. This means:
- Clear and direct communication
- Appropriate formality without being stiff
- Action-oriented language
- Confident but not arrogant
- Industry-appropriate terminology

Make the edits directly to the document.`,

  friendly: `Please rewrite the document to have a warm, friendly tone. This means:
- Approachable and personable language
- Showing empathy and understanding
- Using inclusive language (we, us)
- Being helpful and supportive
- Adding warmth without being unprofessional

Make the edits directly to the document.`,

  clarity: `Please improve the clarity of this document. Focus on:
- Breaking up long, complex sentences
- Removing ambiguous phrases
- Making the main points obvious
- Using clearer word choices
- Improving logical flow
- Adding transitions where needed

Highlight any sections you changed in cyan and add comments explaining your improvements.`,

  concise: `Please make this document more concise without losing meaning. Focus on:
- Removing redundant words and phrases
- Eliminating filler words
- Combining sentences where appropriate
- Getting to the point faster
- Removing unnecessary qualifiers

Make the edits directly to the document.`,

  shorter: `Please significantly shorten this document while keeping the key points. Aim to reduce the length by about 30-50%. Remove:
- Redundant information
- Excessive examples
- Unnecessary elaboration
- Filler content

Make the edits directly to the document.`,

  longer: `Please expand and elaborate on this document. For each main point:
- Add more detail and explanation
- Include examples where helpful
- Expand on implications
- Add supporting information

Make the additions directly to the document.`,

  suggestions: `Please read through this document and provide suggestions for improvement. Consider:
- Overall structure and organization
- Clarity and readability
- Tone consistency
- Missing information
- Areas that could be stronger

Give me your suggestions as a list. Don't make changes yet - just tell me what you'd recommend.`
};

// ============== AI Provider Configurations ==============

const PROVIDER_CONFIG = {
  openai: {
    url: 'https://api.openai.com/v1/chat/completions',
    model: 'gpt-4o-mini',
    formatRequest: (messages, tools) => ({
      model: 'gpt-4o-mini',
      messages,
      tools,
      tool_choice: 'auto'
    }),
    parseResponse: (data) => data.choices?.[0]?.message,
    getHeaders: (apiKey) => ({
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    })
  },
  gemini: {
    url: (apiKey) => `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
    formatRequest: (messages, tools) => {
      const contents = messages.filter(m => m.role !== 'system').map(m => ({
        role: m.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: m.content || '' }]
      }));
      
      const systemInstruction = messages.find(m => m.role === 'system');
      
      const geminiTools = tools ? [{
        functionDeclarations: tools.map(t => ({
          name: t.function.name,
          description: t.function.description,
          parameters: t.function.parameters
        }))
      }] : undefined;
      
      return {
        contents,
        systemInstruction: systemInstruction ? { parts: [{ text: systemInstruction.content }] } : undefined,
        tools: geminiTools
      };
    },
    parseResponse: (data) => {
      const candidate = data.candidates?.[0];
      if (!candidate) return null;
      
      const parts = candidate.content?.parts || [];
      const textPart = parts.find(p => p.text);
      const functionCall = parts.find(p => p.functionCall);
      
      if (functionCall) {
        return {
          content: textPart?.text || '',
          tool_calls: [{
            id: 'gemini-' + Date.now(),
            function: {
              name: functionCall.functionCall.name,
              arguments: JSON.stringify(functionCall.functionCall.args || {})
            }
          }]
        };
      }
      
      return { content: textPart?.text || '' };
    },
    getHeaders: () => ({ 'Content-Type': 'application/json' })
  },
  claude: {
    url: 'https://api.anthropic.com/v1/messages',
    model: 'claude-3-haiku-20240307',
    formatRequest: (messages, tools) => {
      const systemMsg = messages.find(m => m.role === 'system');
      const otherMsgs = messages.filter(m => m.role !== 'system');
      
      const claudeTools = tools ? tools.map(t => ({
        name: t.function.name,
        description: t.function.description,
        input_schema: t.function.parameters
      })) : undefined;
      
      return {
        model: 'claude-3-haiku-20240307',
        max_tokens: 2048,
        system: systemMsg?.content || '',
        messages: otherMsgs.map(m => ({
          role: m.role === 'assistant' ? 'assistant' : 'user',
          content: m.content || ''
        })),
        tools: claudeTools
      };
    },
    parseResponse: (data) => {
      const content = data.content || [];
      const textBlock = content.find(c => c.type === 'text');
      const toolUse = content.find(c => c.type === 'tool_use');
      
      if (toolUse) {
        return {
          content: textBlock?.text || '',
          tool_calls: [{
            id: toolUse.id,
            function: {
              name: toolUse.name,
              arguments: JSON.stringify(toolUse.input || {})
            }
          }]
        };
      }
      
      return { content: textBlock?.text || '' };
    },
    getHeaders: (apiKey) => ({
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    })
  },
  local: {
    formatRequest: (messages, tools, model) => ({
      model: model || 'local-model',
      messages,
      tools,
      tool_choice: 'auto'
    }),
    parseResponse: (data) => data.choices?.[0]?.message,
    getHeaders: (apiKey) => {
      const headers = { 'Content-Type': 'application/json' };
      if (apiKey) headers['Authorization'] = `Bearer ${apiKey}`;
      return headers;
    }
  }
};

// ============== Document Editing Tools ==============

const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'delete_all_instances_of_text',
      description: 'Delete every instance of a specific word or phrase from the document.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The exact word or phrase to find and delete' }
        },
        required: ['searchText']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'replace_all_text',
      description: 'Replace every instance of a search string with a replacement string.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find' },
          replaceText: { type: 'string', description: 'The text to replace it with' }
        },
        required: ['searchText', 'replaceText']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'highlight_text',
      description: 'Highlight all instances of specific text with a color. Use this when asked to highlight, mark, or visually emphasize text.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to highlight' },
          color: { 
            type: 'string', 
            description: 'Highlight color: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, gray25, gray50',
            enum: ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'gray25', 'gray50']
          }
        },
        required: ['searchText', 'color']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'remove_highlight',
      description: 'Remove highlighting from all instances of specific text.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to remove highlighting from. Use "*" to remove all highlights.' }
        },
        required: ['searchText']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_comment',
      description: 'Add a comment/annotation to specific text in the document. The comment will appear in the margin.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to attach the comment to' },
          comment: { type: 'string', description: 'The comment text to add' },
          matchIndex: { type: 'number', description: 'Which occurrence to comment on (0 = first, 1 = second, etc.). Default is 0.' }
        },
        required: ['searchText', 'comment']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'insert_text_at_end',
      description: 'Insert text at the end of the document.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The text to insert' }
        },
        required: ['text']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'insert_text_at_start',
      description: 'Insert text at the beginning of the document.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The text to insert' }
        },
        required: ['text']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_document_content',
      description: 'Get the current text content of the document to analyze it.',
      parameters: { type: 'object', properties: {} }
    }
  }
];

// ============== Tool Implementations ==============

const HIGHLIGHT_COLORS = {
  yellow: 'Yellow',
  green: 'BrightGreen', 
  cyan: 'Turquoise',
  magenta: 'Pink',
  blue: 'Blue',
  red: 'Red',
  darkBlue: 'DarkBlue',
  darkCyan: 'DarkCyan',
  darkGreen: 'DarkGreen',
  darkMagenta: 'DarkMagenta',
  darkRed: 'DarkRed',
  darkYellow: 'DarkYellow',
  gray25: 'Gray25',
  gray50: 'Gray50'
};

async function deleteAllText(searchText) {
  if (!searchText || typeof searchText !== 'string') {
    return { success: false, count: 0, message: 'searchText must be a non-empty string.' };
  }
  
  const Word = getWord();
  if (!Word) return { success: false, count: 0, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    const results = context.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    
    const items = results.items;
    for (let i = items.length - 1; i >= 0; i--) {
      items[i].insertText('', Word.InsertLocation.replace);
    }
    await context.sync();
    
    return { success: true, count: items.length, message: `Deleted ${items.length} instance(s) of "${searchText}".` };
  });
}

async function replaceAllText(searchText, replaceText) {
  if (!searchText || typeof searchText !== 'string') {
    return { success: false, count: 0, message: 'searchText must be a non-empty string.' };
  }
  if (replaceText == null) replaceText = '';
  
  const Word = getWord();
  if (!Word) return { success: false, count: 0, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    const results = context.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    
    const items = results.items;
    for (let i = items.length - 1; i >= 0; i--) {
      items[i].insertText(replaceText, Word.InsertLocation.replace);
    }
    await context.sync();
    
    return { success: true, count: items.length, message: `Replaced ${items.length} instance(s) of "${searchText}" with "${replaceText}".` };
  });
}

async function highlightText(searchText, color) {
  if (!searchText || typeof searchText !== 'string') {
    return { success: false, count: 0, message: 'searchText must be a non-empty string.' };
  }
  
  const Word = getWord();
  if (!Word) return { success: false, count: 0, message: 'Word is not ready.' };
  
  const wordColor = HIGHLIGHT_COLORS[color] || 'Yellow';
  
  return Word.run(async (context) => {
    const results = context.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    
    const items = results.items;
    for (const item of items) {
      item.font.highlightColor = wordColor;
    }
    await context.sync();
    
    return { success: true, count: items.length, message: `Highlighted ${items.length} instance(s) of "${searchText}" in ${color}.` };
  });
}

async function removeHighlight(searchText) {
  const Word = getWord();
  if (!Word) return { success: false, count: 0, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    if (searchText === '*') {
      const body = context.document.body;
      body.font.highlightColor = null;
      await context.sync();
      return { success: true, message: 'Removed all highlights from the document.' };
    }
    
    const results = context.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    
    const items = results.items;
    for (const item of items) {
      item.font.highlightColor = null;
    }
    await context.sync();
    
    return { success: true, count: items.length, message: `Removed highlighting from ${items.length} instance(s) of "${searchText}".` };
  });
}

async function addComment(searchText, commentText, matchIndex = 0) {
  if (!searchText || typeof searchText !== 'string') {
    return { success: false, message: 'searchText must be a non-empty string.' };
  }
  if (!commentText || typeof commentText !== 'string') {
    return { success: false, message: 'comment must be a non-empty string.' };
  }
  
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    const results = context.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    
    const items = results.items;
    if (items.length === 0) {
      return { success: false, message: `Could not find "${searchText}" in the document.` };
    }
    
    const index = Math.min(matchIndex || 0, items.length - 1);
    const targetRange = items[index];
    
    targetRange.insertComment(commentText);
    await context.sync();
    
    return { success: true, message: `Added comment to "${searchText}" (occurrence ${index + 1} of ${items.length}).` };
  });
}

async function insertTextAtEnd(text) {
  if (!text || typeof text !== 'string') {
    return { success: false, message: 'text must be a non-empty string.' };
  }
  
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    context.document.body.insertText(text, Word.InsertLocation.end);
    await context.sync();
    return { success: true, message: 'Inserted text at end of document.' };
  });
}

async function insertTextAtStart(text) {
  if (!text || typeof text !== 'string') {
    return { success: false, message: 'text must be a non-empty string.' };
  }
  
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word is not ready.' };
  
  return Word.run(async (context) => {
    context.document.body.insertText(text, Word.InsertLocation.start);
    await context.sync();
    return { success: true, message: 'Inserted text at start of document.' };
  });
}

async function getDocumentContent() {
  const Word = getWord();
  if (!Word) return { success: false, content: '' };
  
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return { success: true, content: (body.text || '').slice(0, 12000) };
  });
}

async function executeTool(toolName, args) {
  switch (toolName) {
    case 'delete_all_instances_of_text':
      return deleteAllText(args.searchText);
    case 'replace_all_text':
      return replaceAllText(args.searchText, args.replaceText);
    case 'highlight_text':
      return highlightText(args.searchText, args.color);
    case 'remove_highlight':
      return removeHighlight(args.searchText);
    case 'add_comment':
      return addComment(args.searchText, args.comment, args.matchIndex);
    case 'insert_text_at_end':
      return insertTextAtEnd(args.text);
    case 'insert_text_at_start':
      return insertTextAtStart(args.text);
    case 'get_document_content':
      return getDocumentContent();
    default:
      return { success: false, message: `Unknown tool: ${toolName}` };
  }
}

// ============== API Calls ==============

async function callAI(provider, apiKey, messages, localUrl, localModel) {
  const config = PROVIDER_CONFIG[provider];
  if (!config) throw new Error(`Unknown provider: ${provider}`);
  
  let url = config.url;
  if (typeof url === 'function') {
    url = url(apiKey);
  }
  if (provider === 'local') {
    url = localUrl || 'http://localhost:1234/v1/chat/completions';
  }
  
  const headers = config.getHeaders(apiKey);
  const body = config.formatRequest(messages, TOOLS, localModel);
  
  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body)
  });
  
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || `API error: ${response.status}`);
  }
  
  const data = await response.json();
  return config.parseResponse(data);
}

// ============== UI Functions ==============

function addMessage(type, content, label = '') {
  const messagesEl = document.getElementById('messages');
  const div = document.createElement('div');
  div.className = `message ${type}`;
  
  // Add label
  if (label) {
    const labelEl = document.createElement('div');
    labelEl.className = 'message-label';
    labelEl.textContent = label;
    div.appendChild(labelEl);
  }
  
  // Add content
  const contentEl = document.createElement('div');
  contentEl.textContent = content;
  div.appendChild(contentEl);
  
  messagesEl.appendChild(div);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function setStatus(text) {
  document.getElementById('status').textContent = text;
}

function setSendEnabled(enabled) {
  const btn = document.getElementById('send-btn');
  btn.disabled = !enabled;
}

function setQuickButtonsEnabled(enabled) {
  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.disabled = !enabled;
    btn.style.opacity = enabled ? '1' : '0.5';
    btn.style.pointerEvents = enabled ? 'auto' : 'none';
  });
}

// ============== Main App ==============

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = '<p style="padding:20px;">This add-in only works in Microsoft Word.</p>';
    return;
  }
  
  // Elements
  const settingsToggle = document.getElementById('settings-toggle');
  const settingsPanel = document.getElementById('settings-panel');
  const providerSelect = document.getElementById('provider');
  const apiKeyInput = document.getElementById('api-key');
  const localSettings = document.getElementById('local-settings');
  const localUrlInput = document.getElementById('local-url');
  const localModelInput = document.getElementById('local-model');
  const customPromptInput = document.getElementById('custom-prompt');
  const userInput = document.getElementById('user-input');
  const sendBtn = document.getElementById('send-btn');
  const quickToggle = document.getElementById('quick-toggle');
  const quickGrid = document.getElementById('quick-grid');
  
  // Load saved settings
  providerSelect.value = localStorage.getItem(STORAGE_KEYS.provider) || 'openai';
  apiKeyInput.value = localStorage.getItem(STORAGE_KEYS.apiKey) || '';
  localUrlInput.value = localStorage.getItem(STORAGE_KEYS.localUrl) || 'http://localhost:1234/v1/chat/completions';
  localModelInput.value = localStorage.getItem(STORAGE_KEYS.localModel) || '';
  customPromptInput.value = localStorage.getItem(STORAGE_KEYS.customPrompt) || '';
  
  // Show/hide local settings
  const updateLocalSettings = () => {
    localSettings.classList.toggle('visible', providerSelect.value === 'local');
  };
  updateLocalSettings();
  
  // Settings toggle
  settingsToggle.addEventListener('click', () => {
    settingsPanel.classList.toggle('open');
  });
  
  // Quick actions toggle
  quickToggle.addEventListener('click', () => {
    const isCollapsed = quickGrid.classList.toggle('collapsed');
    quickToggle.textContent = isCollapsed ? 'Show' : 'Hide';
  });
  
  // Save settings on change
  providerSelect.addEventListener('change', () => {
    localStorage.setItem(STORAGE_KEYS.provider, providerSelect.value);
    updateLocalSettings();
  });
  
  apiKeyInput.addEventListener('change', () => {
    localStorage.setItem(STORAGE_KEYS.apiKey, apiKeyInput.value);
  });
  
  localUrlInput.addEventListener('change', () => {
    localStorage.setItem(STORAGE_KEYS.localUrl, localUrlInput.value);
  });
  
  localModelInput.addEventListener('change', () => {
    localStorage.setItem(STORAGE_KEYS.localModel, localModelInput.value);
  });
  
  customPromptInput.addEventListener('change', () => {
    localStorage.setItem(STORAGE_KEYS.customPrompt, customPromptInput.value);
  });
  
  // Enable/disable send button
  const updateSendButton = () => {
    const provider = providerSelect.value;
    const hasKey = apiKeyInput.value.trim().length > 0 || provider === 'local';
    const hasInput = userInput.value.trim().length > 0;
    sendBtn.disabled = !(hasKey && hasInput);
  };
  
  apiKeyInput.addEventListener('input', updateSendButton);
  userInput.addEventListener('input', updateSendButton);
  providerSelect.addEventListener('change', updateSendButton);
  updateSendButton();
  
  // Get saved writing style
  let savedWritingStyle = localStorage.getItem(STORAGE_KEYS.writingStyle) || '';
  
  // Build system prompt
  const getSystemPrompt = () => {
    let prompt = `You are an AI writing assistant that helps users edit their Word document. You have access to these tools:

- delete_all_instances_of_text: Remove every occurrence of a word or phrase
- replace_all_text: Find and replace text throughout the document  
- highlight_text: Highlight text with a color (yellow, green, cyan, magenta, blue, red, etc.)
- remove_highlight: Remove highlighting from text (use "*" to remove all)
- add_comment: Add a comment/annotation to specific text in the margin
- insert_text_at_end: Add text at the end
- insert_text_at_start: Add text at the beginning
- get_document_content: Read the document's text

When the user asks you to make edits, use the appropriate tools. ALWAYS use get_document_content first to see what's in the document before making changes.

Be concise in your responses. After performing an action, briefly confirm what you did.

When highlighting, always specify a color. Default to yellow if none specified.
When adding comments, the comment appears in the document margin attached to that text.`;

    // Add saved writing style if available
    if (savedWritingStyle) {
      prompt += `\n\nThe user's writing style has been analyzed previously. When making edits or additions to their document, try to match this style:\n${savedWritingStyle}`;
    }
    
    // Add custom prompt if set
    const customPrompt = customPromptInput.value.trim();
    if (customPrompt) {
      prompt += `\n\nAdditional user instructions:\n${customPrompt}`;
    }
    
    return prompt;
  };
  
  // Conversation history
  let conversation = [];
  
  const resetConversation = () => {
    conversation = [{ role: 'system', content: getSystemPrompt() }];
  };
  resetConversation();
  
  // Send message function
  const sendMessage = async (text, isQuickAction = false) => {
    if (!text) return;
    
    const provider = providerSelect.value;
    const apiKey = apiKeyInput.value.trim();
    const localUrl = localUrlInput.value.trim();
    const localModel = localModelInput.value.trim();
    
    if (!apiKey && provider !== 'local') {
      addMessage('error', 'Please enter your API key in Settings.', 'Error');
      settingsPanel.classList.add('open');
      return;
    }
    
    // Save settings
    localStorage.setItem(STORAGE_KEYS.apiKey, apiKey);
    
    if (!isQuickAction) {
      userInput.value = '';
      updateSendButton();
    }
    
    addMessage('user', text, 'You');
    setSendEnabled(false);
    setQuickButtonsEnabled(false);
    setStatus('Thinking...');
    
    // Reset conversation with fresh system prompt
    resetConversation();
    conversation.push({ role: 'user', content: text });
    
    try {
      let processing = true;
      
      while (processing) {
        const response = await callAI(provider, apiKey, conversation, localUrl, localModel);
        
        if (!response) {
          throw new Error('No response from API');
        }
        
        if (response.tool_calls && response.tool_calls.length > 0) {
          conversation.push({
            role: 'assistant',
            content: response.content || '',
            tool_calls: response.tool_calls
          });
          
          for (const toolCall of response.tool_calls) {
            const toolName = toolCall.function?.name;
            let args = {};
            
            try {
              args = JSON.parse(toolCall.function?.arguments || '{}');
            } catch (e) {}
            
            setStatus(`Running: ${toolName}...`);
            addMessage('tool', `${toolName}(${Object.values(args).join(', ').slice(0, 50)}...)`, 'Tool');
            
            let result;
            try {
              result = await executeTool(toolName, args);
            } catch (e) {
              result = { success: false, message: e.message };
            }
            
            const resultStr = typeof result === 'object' ? JSON.stringify(result) : String(result);
            
            conversation.push({
              role: 'tool',
              tool_call_id: toolCall.id,
              content: resultStr
            });
            
            setStatus('Processing...');
          }
        } else {
          const aiResponse = response.content?.trim() || 'Done.';
          addMessage('assistant', aiResponse, 'AI Assistant');
          conversation.push({ role: 'assistant', content: response.content || '' });
          processing = false;
          setStatus('Ready');
          
          // Check if this was a style analysis - save it
          if (text.includes('analyze') && text.toLowerCase().includes('style')) {
            savedWritingStyle = aiResponse;
            localStorage.setItem(STORAGE_KEYS.writingStyle, savedWritingStyle);
          }
        }
      }
    } catch (e) {
      addMessage('error', `Error: ${e.message}`, 'Error');
      setStatus('Error');
    } finally {
      setSendEnabled(true);
      setQuickButtonsEnabled(true);
      updateSendButton();
    }
  };
  
  // Send button click
  sendBtn.addEventListener('click', () => {
    sendMessage(userInput.value.trim());
  });
  
  // Quick action buttons
  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      const prompt = QUICK_ACTIONS[action];
      if (prompt) {
        sendMessage(prompt, true);
      }
    });
  });
  
  // Enter to send
  userInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendBtn.click();
    }
  });
  
  // Initial message
  addMessage('system', 'Configure your AI provider in Settings, then use Quick Actions or type a request. Your API key and settings are saved locally in your browser.', 'Welcome');
});
