// Word AI Assistant - Multi-Provider Support with Writing Tools
// Supports: OpenAI, Google Gemini, Anthropic Claude, Local models
// Privacy: All data goes directly to your chosen AI provider - no external storage

const STORAGE_KEYS = {
  // Privacy settings
  privacySaveKey: 'word-ai-privacy-save-key',
  privacySavePrompt: 'word-ai-privacy-save-prompt',
  privacySaveStyle: 'word-ai-privacy-save-style',
  privacySaveProvider: 'word-ai-privacy-save-provider',
  privacySaveButtons: 'word-ai-privacy-save-buttons',
  // Data
  provider: 'word-ai-provider',
  apiKey: 'word-ai-api-key',
  localUrl: 'word-ai-local-url',
  localModel: 'word-ai-local-model',
  customPrompt: 'word-ai-custom-prompt',
  writingStyle: 'word-ai-writing-style',
  buttonCustomizations: 'word-ai-button-customizations'
};

// Get Word API reference
const getWord = () => window.Word || window.Office?.Word;

// ============== Privacy Helpers ==============

function getPrivacySetting(key) {
  return localStorage.getItem(key) === 'true';
}

function setPrivacySetting(key, value) {
  localStorage.setItem(key, value ? 'true' : 'false');
}

function saveIfAllowed(key, value, privacyKey) {
  if (getPrivacySetting(privacyKey)) {
    localStorage.setItem(key, value);
  } else {
    localStorage.removeItem(key);
  }
}

function loadIfAllowed(key, privacyKey, defaultValue = '') {
  if (getPrivacySetting(privacyKey)) {
    return localStorage.getItem(key) || defaultValue;
  }
  return defaultValue;
}

function clearAllData() {
  localStorage.removeItem(STORAGE_KEYS.provider);
  localStorage.removeItem(STORAGE_KEYS.apiKey);
  localStorage.removeItem(STORAGE_KEYS.localUrl);
  localStorage.removeItem(STORAGE_KEYS.localModel);
  localStorage.removeItem(STORAGE_KEYS.customPrompt);
  localStorage.removeItem(STORAGE_KEYS.writingStyle);
  localStorage.removeItem(STORAGE_KEYS.buttonCustomizations);
}

// ============== Default Quick Action Prompts ==============

const DEFAULT_PROMPTS = {
  analyze: {
    name: 'Analyze My Writing Style',
    icon: '🔍',
    prompt: `Please analyze the writing style of this document. Look at:
1. Tone (formal, casual, professional, academic, etc.)
2. Sentence structure and complexity
3. Vocabulary level
4. Voice (active vs passive)
5. Any patterns or habits

After analyzing, remember this style so you can mimic it when making future edits. Give me a summary of the writing style you detected.`
  },
  grammar: {
    name: 'Grammar Check',
    icon: '✓',
    prompt: `Please check the entire document for grammar errors. For each error found:
1. Highlight the problematic text in yellow
2. Add a comment explaining the grammar issue and the correction

After checking, give me a summary of how many issues were found.`
  },
  spelling: {
    name: 'Spelling Check',
    icon: '📝',
    prompt: `Please check the entire document for spelling errors and typos. For each error found:
1. Highlight the misspelled word in red
2. Add a comment with the correct spelling

After checking, give me a summary of what was found.`
  },
  formal: {
    name: 'Make Formal',
    icon: '👔',
    prompt: `Please rewrite the document content to use a formal, professional tone. Make these changes:
- Replace casual language with formal alternatives
- Use complete sentences
- Avoid contractions
- Use professional vocabulary
- Maintain a respectful, businesslike tone

Make the edits directly to the document.`
  },
  casual: {
    name: 'Make Casual',
    icon: '😊',
    prompt: `Please rewrite the document content to use a casual, conversational tone. Make these changes:
- Use contractions where natural
- Simplify complex sentences
- Use everyday vocabulary
- Make it sound like a friendly conversation
- Keep it approachable and relaxed

Make the edits directly to the document.`
  },
  professional: {
    name: 'Professional Tone',
    icon: '💼',
    prompt: `Please adjust the document to have a professional business tone. This means:
- Clear and direct communication
- Appropriate formality without being stiff
- Action-oriented language
- Confident but not arrogant
- Industry-appropriate terminology

Make the edits directly to the document.`
  },
  friendly: {
    name: 'Friendly Tone',
    icon: '🤝',
    prompt: `Please rewrite the document to have a warm, friendly tone. This means:
- Approachable and personable language
- Showing empathy and understanding
- Using inclusive language (we, us)
- Being helpful and supportive
- Adding warmth without being unprofessional

Make the edits directly to the document.`
  },
  clarity: {
    name: 'Improve Clarity',
    icon: '💡',
    prompt: `Please improve the clarity of this document. Focus on:
- Breaking up long, complex sentences
- Removing ambiguous phrases
- Making the main points obvious
- Using clearer word choices
- Improving logical flow
- Adding transitions where needed

Highlight any sections you changed in cyan and add comments explaining your improvements.`
  },
  concise: {
    name: 'Make Concise',
    icon: '✂️',
    prompt: `Please make this document more concise without losing meaning. Focus on:
- Removing redundant words and phrases
- Eliminating filler words
- Combining sentences where appropriate
- Getting to the point faster
- Removing unnecessary qualifiers

Make the edits directly to the document.`
  },
  shorter: {
    name: 'Shorten',
    icon: '📉',
    prompt: `Please significantly shorten this document while keeping the key points. Aim to reduce the length by about 30-50%. Remove:
- Redundant information
- Excessive examples
- Unnecessary elaboration
- Filler content

Make the edits directly to the document.`
  },
  longer: {
    name: 'Expand/Elaborate',
    icon: '📈',
    prompt: `Please expand and elaborate on this document. For each main point:
- Add more detail and explanation
- Include examples where helpful
- Expand on implications
- Add supporting information

Make the additions directly to the document.`
  },
  suggestions: {
    name: 'Get Suggestions',
    icon: '💬',
    prompt: `Please read through this document and provide suggestions for improvement. Consider:
- Overall structure and organization
- Clarity and readability
- Tone consistency
- Missing information
- Areas that could be stronger

Give me your suggestions as a list. Don't make changes yet - just tell me what you'd recommend.`
  }
};

// Button customizations (loaded from storage)
let buttonCustomizations = {};

function loadButtonCustomizations() {
  const saved = loadIfAllowed(STORAGE_KEYS.buttonCustomizations, STORAGE_KEYS.privacySaveButtons, '{}');
  try {
    buttonCustomizations = JSON.parse(saved);
  } catch (e) {
    buttonCustomizations = {};
  }
}

function saveButtonCustomizations() {
  saveIfAllowed(STORAGE_KEYS.buttonCustomizations, JSON.stringify(buttonCustomizations), STORAGE_KEYS.privacySaveButtons);
}

function getButtonConfig(action) {
  const defaults = DEFAULT_PROMPTS[action];
  const custom = buttonCustomizations[action] || {};
  return {
    name: custom.name || defaults.name,
    icon: defaults.icon,
    prompt: defaults.prompt + (custom.extra ? '\n\nAdditional instructions: ' + custom.extra : ''),
    extra: custom.extra || '',
    isCustomized: !!(custom.name || custom.extra)
  };
}

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
      description: 'Highlight all instances of specific text with a color.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to highlight' },
          color: { 
            type: 'string', 
            description: 'Highlight color',
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
      description: 'Remove highlighting from text. Use "*" to remove all highlights.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to remove highlighting from' }
        },
        required: ['searchText']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_comment',
      description: 'Add a comment to specific text in the document margin.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to attach the comment to' },
          comment: { type: 'string', description: 'The comment text' },
          matchIndex: { type: 'number', description: 'Which occurrence (0 = first)' }
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
      description: 'Get the current text content of the document.',
      parameters: { type: 'object', properties: {} }
    }
  }
];

// ============== Tool Implementations ==============

const HIGHLIGHT_COLORS = {
  yellow: 'Yellow', green: 'BrightGreen', cyan: 'Turquoise', magenta: 'Pink',
  blue: 'Blue', red: 'Red', darkBlue: 'DarkBlue', darkCyan: 'DarkCyan',
  darkGreen: 'DarkGreen', darkMagenta: 'DarkMagenta', darkRed: 'DarkRed',
  darkYellow: 'DarkYellow', gray25: 'Gray25', gray50: 'Gray50'
};

async function deleteAllText(searchText) {
  if (!searchText) return { success: false, message: 'searchText required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await ctx.sync();
    for (let i = results.items.length - 1; i >= 0; i--) {
      results.items[i].insertText('', Word.InsertLocation.replace);
    }
    await ctx.sync();
    return { success: true, count: results.items.length, message: `Deleted ${results.items.length} instance(s)` };
  });
}

async function replaceAllText(searchText, replaceText) {
  if (!searchText) return { success: false, message: 'searchText required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await ctx.sync();
    for (let i = results.items.length - 1; i >= 0; i--) {
      results.items[i].insertText(replaceText || '', Word.InsertLocation.replace);
    }
    await ctx.sync();
    return { success: true, count: results.items.length, message: `Replaced ${results.items.length} instance(s)` };
  });
}

async function highlightText(searchText, color) {
  if (!searchText) return { success: false, message: 'searchText required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await ctx.sync();
    for (const item of results.items) {
      item.font.highlightColor = HIGHLIGHT_COLORS[color] || 'Yellow';
    }
    await ctx.sync();
    return { success: true, count: results.items.length, message: `Highlighted ${results.items.length} instance(s)` };
  });
}

async function removeHighlight(searchText) {
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    if (searchText === '*') {
      ctx.document.body.font.highlightColor = null;
      await ctx.sync();
      return { success: true, message: 'Removed all highlights' };
    }
    const results = ctx.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await ctx.sync();
    for (const item of results.items) {
      item.font.highlightColor = null;
    }
    await ctx.sync();
    return { success: true, count: results.items.length, message: `Removed highlights from ${results.items.length} instance(s)` };
  });
}

async function addComment(searchText, commentText, matchIndex = 0) {
  if (!searchText || !commentText) return { success: false, message: 'searchText and comment required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(searchText, { matchCase: false });
    results.load('items');
    await ctx.sync();
    if (results.items.length === 0) return { success: false, message: `"${searchText}" not found` };
    const idx = Math.min(matchIndex || 0, results.items.length - 1);
    results.items[idx].insertComment(commentText);
    await ctx.sync();
    return { success: true, message: `Added comment to occurrence ${idx + 1}` };
  });
}

async function insertTextAtEnd(text) {
  if (!text) return { success: false, message: 'text required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    ctx.document.body.insertText(text, Word.InsertLocation.end);
    await ctx.sync();
    return { success: true, message: 'Inserted at end' };
  });
}

async function insertTextAtStart(text) {
  if (!text) return { success: false, message: 'text required' };
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word not ready' };
  return Word.run(async (ctx) => {
    ctx.document.body.insertText(text, Word.InsertLocation.start);
    await ctx.sync();
    return { success: true, message: 'Inserted at start' };
  });
}

async function getDocumentContent() {
  const Word = getWord();
  if (!Word) return { success: false, content: '' };
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load('text');
    await ctx.sync();
    return { success: true, content: (body.text || '').slice(0, 12000) };
  });
}

async function executeTool(toolName, args) {
  switch (toolName) {
    case 'delete_all_instances_of_text': return deleteAllText(args.searchText);
    case 'replace_all_text': return replaceAllText(args.searchText, args.replaceText);
    case 'highlight_text': return highlightText(args.searchText, args.color);
    case 'remove_highlight': return removeHighlight(args.searchText);
    case 'add_comment': return addComment(args.searchText, args.comment, args.matchIndex);
    case 'insert_text_at_end': return insertTextAtEnd(args.text);
    case 'insert_text_at_start': return insertTextAtStart(args.text);
    case 'get_document_content': return getDocumentContent();
    default: return { success: false, message: `Unknown tool: ${toolName}` };
  }
}

// ============== API Calls ==============

async function callAI(provider, apiKey, messages, localUrl, localModel) {
  const config = PROVIDER_CONFIG[provider];
  if (!config) throw new Error(`Unknown provider: ${provider}`);
  
  let url = config.url;
  if (typeof url === 'function') url = url(apiKey);
  if (provider === 'local') url = localUrl || 'http://localhost:1234/v1/chat/completions';
  
  const response = await fetch(url, {
    method: 'POST',
    headers: config.getHeaders(apiKey),
    body: JSON.stringify(config.formatRequest(messages, TOOLS, localModel))
  });
  
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || `API error: ${response.status}`);
  }
  
  return config.parseResponse(await response.json());
}

// ============== UI Functions ==============

function addMessage(type, content, label = '') {
  const messagesEl = document.getElementById('messages');
  const div = document.createElement('div');
  div.className = `message ${type}`;
  if (label) {
    const labelEl = document.createElement('div');
    labelEl.className = 'message-label';
    labelEl.textContent = label;
    div.appendChild(labelEl);
  }
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
  document.getElementById('send-btn').disabled = !enabled;
}

function setQuickButtonsEnabled(enabled) {
  document.querySelectorAll('.quick-btn').forEach(btn => btn.disabled = !enabled);
}

// ============== Main App ==============

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = '<p style="padding:20px;">This add-in only works in Microsoft Word.</p>';
    return;
  }
  
  // Load button customizations
  loadButtonCustomizations();
  
  // Elements
  const settingsToggle = document.getElementById('settings-toggle');
  const settingsPanel = document.getElementById('settings-panel');
  const privacyToggle = document.getElementById('privacy-toggle');
  const privacyPanel = document.getElementById('privacy-panel');
  const providerSelect = document.getElementById('provider');
  const apiKeyInput = document.getElementById('api-key');
  const apiKeyHint = document.getElementById('api-key-hint');
  const localSettings = document.getElementById('local-settings');
  const localUrlInput = document.getElementById('local-url');
  const localModelInput = document.getElementById('local-model');
  const customPromptInput = document.getElementById('custom-prompt');
  const userInput = document.getElementById('user-input');
  const sendBtn = document.getElementById('send-btn');
  const quickToggle = document.getElementById('quick-toggle');
  const quickGrid = document.getElementById('quick-grid');
  const editToggle = document.getElementById('edit-toggle');
  const clearDataBtn = document.getElementById('clear-all-data');
  const popoutBtn = document.getElementById('popout-btn');
  
  // Modal elements
  const editModal = document.getElementById('edit-modal');
  const modalIcon = document.getElementById('modal-icon');
  const modalName = document.getElementById('modal-name');
  const modalExtra = document.getElementById('modal-extra');
  const modalDefault = document.getElementById('modal-default');
  const modalClose = document.getElementById('modal-close');
  const modalCancel = document.getElementById('modal-cancel');
  const modalSave = document.getElementById('modal-save');
  const modalReset = document.getElementById('modal-reset');
  
  let currentEditAction = null;
  let isEditMode = false;
  
  // Privacy toggles
  const privacySaveKey = document.getElementById('privacy-save-key');
  const privacySavePrompt = document.getElementById('privacy-save-prompt');
  const privacySaveStyle = document.getElementById('privacy-save-style');
  const privacySaveProvider = document.getElementById('privacy-save-provider');
  const privacySaveButtons = document.getElementById('privacy-save-buttons');
  
  // Load privacy settings
  privacySaveKey.checked = getPrivacySetting(STORAGE_KEYS.privacySaveKey);
  privacySavePrompt.checked = getPrivacySetting(STORAGE_KEYS.privacySavePrompt);
  privacySaveStyle.checked = getPrivacySetting(STORAGE_KEYS.privacySaveStyle);
  privacySaveProvider.checked = getPrivacySetting(STORAGE_KEYS.privacySaveProvider);
  privacySaveButtons.checked = getPrivacySetting(STORAGE_KEYS.privacySaveButtons);
  
  // Load saved data
  providerSelect.value = loadIfAllowed(STORAGE_KEYS.provider, STORAGE_KEYS.privacySaveProvider, 'openai');
  apiKeyInput.value = loadIfAllowed(STORAGE_KEYS.apiKey, STORAGE_KEYS.privacySaveKey, '');
  localUrlInput.value = loadIfAllowed(STORAGE_KEYS.localUrl, STORAGE_KEYS.privacySaveProvider, 'http://localhost:1234/v1/chat/completions');
  localModelInput.value = loadIfAllowed(STORAGE_KEYS.localModel, STORAGE_KEYS.privacySaveProvider, '');
  customPromptInput.value = loadIfAllowed(STORAGE_KEYS.customPrompt, STORAGE_KEYS.privacySavePrompt, '');
  
  // Update button appearances
  const updateButtonAppearances = () => {
    document.querySelectorAll('.quick-btn[data-action]').forEach(btn => {
      const action = btn.dataset.action;
      const config = getButtonConfig(action);
      btn.querySelector('.quick-btn-text').textContent = config.name;
      btn.classList.toggle('customized', config.isCustomized);
      btn.classList.toggle('edit-mode', isEditMode);
    });
  };
  updateButtonAppearances();
  
  // API key hint
  const updateApiKeyHint = () => {
    apiKeyHint.textContent = privacySaveKey.checked 
      ? 'Your API key will be saved locally'
      : 'API key will NOT be saved (enable in Privacy)';
    apiKeyHint.style.color = privacySaveKey.checked ? '#81c784' : '';
  };
  updateApiKeyHint();
  
  // Local settings visibility
  const updateLocalSettings = () => {
    localSettings.classList.toggle('visible', providerSelect.value === 'local');
  };
  updateLocalSettings();
  
  // Pop-out functionality
  popoutBtn.addEventListener('click', () => {
    const width = 400;
    const height = 700;
    const left = window.screenX + window.outerWidth - width - 50;
    const top = window.screenY + 50;
    window.open(
      window.location.href,
      'WordAIAssistant',
      `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
    );
  });
  
  // Panel toggles
  settingsToggle.addEventListener('click', () => {
    privacyPanel.classList.remove('open');
    settingsPanel.classList.toggle('open');
  });
  
  privacyToggle.addEventListener('click', () => {
    settingsPanel.classList.remove('open');
    privacyPanel.classList.toggle('open');
  });
  
  // Quick actions toggle
  quickToggle.addEventListener('click', () => {
    const collapsed = quickGrid.classList.toggle('collapsed');
    quickToggle.textContent = collapsed ? 'Show' : 'Hide';
  });
  
  // Edit mode toggle
  editToggle.addEventListener('click', () => {
    isEditMode = !isEditMode;
    editToggle.classList.toggle('active', isEditMode);
    editToggle.textContent = isEditMode ? 'Done' : 'Edit';
    updateButtonAppearances();
  });
  
  // Open edit modal
  const openEditModal = (action) => {
    currentEditAction = action;
    const defaults = DEFAULT_PROMPTS[action];
    const custom = buttonCustomizations[action] || {};
    
    modalIcon.textContent = defaults.icon;
    modalName.value = custom.name || '';
    modalExtra.value = custom.extra || '';
    modalDefault.value = defaults.prompt;
    
    editModal.classList.add('open');
  };
  
  // Close edit modal
  const closeEditModal = () => {
    editModal.classList.remove('open');
    currentEditAction = null;
  };
  
  modalClose.addEventListener('click', closeEditModal);
  modalCancel.addEventListener('click', closeEditModal);
  editModal.addEventListener('click', (e) => {
    if (e.target === editModal) closeEditModal();
  });
  
  // Save customization
  modalSave.addEventListener('click', () => {
    if (!currentEditAction) return;
    
    const name = modalName.value.trim();
    const extra = modalExtra.value.trim();
    
    if (name || extra) {
      buttonCustomizations[currentEditAction] = { name, extra };
    } else {
      delete buttonCustomizations[currentEditAction];
    }
    
    saveButtonCustomizations();
    updateButtonAppearances();
    closeEditModal();
  });
  
  // Reset to default
  modalReset.addEventListener('click', () => {
    if (!currentEditAction) return;
    delete buttonCustomizations[currentEditAction];
    saveButtonCustomizations();
    updateButtonAppearances();
    closeEditModal();
  });
  
  // Privacy toggle handlers
  privacySaveKey.addEventListener('change', () => {
    setPrivacySetting(STORAGE_KEYS.privacySaveKey, privacySaveKey.checked);
    if (privacySaveKey.checked && apiKeyInput.value) {
      localStorage.setItem(STORAGE_KEYS.apiKey, apiKeyInput.value);
    } else if (!privacySaveKey.checked) {
      localStorage.removeItem(STORAGE_KEYS.apiKey);
    }
    updateApiKeyHint();
  });
  
  privacySavePrompt.addEventListener('change', () => {
    setPrivacySetting(STORAGE_KEYS.privacySavePrompt, privacySavePrompt.checked);
    if (privacySavePrompt.checked && customPromptInput.value) {
      localStorage.setItem(STORAGE_KEYS.customPrompt, customPromptInput.value);
    } else if (!privacySavePrompt.checked) {
      localStorage.removeItem(STORAGE_KEYS.customPrompt);
    }
  });
  
  privacySaveStyle.addEventListener('change', () => {
    setPrivacySetting(STORAGE_KEYS.privacySaveStyle, privacySaveStyle.checked);
    if (!privacySaveStyle.checked) localStorage.removeItem(STORAGE_KEYS.writingStyle);
  });
  
  privacySaveProvider.addEventListener('change', () => {
    setPrivacySetting(STORAGE_KEYS.privacySaveProvider, privacySaveProvider.checked);
    if (privacySaveProvider.checked) {
      localStorage.setItem(STORAGE_KEYS.provider, providerSelect.value);
      localStorage.setItem(STORAGE_KEYS.localUrl, localUrlInput.value);
      localStorage.setItem(STORAGE_KEYS.localModel, localModelInput.value);
    } else {
      localStorage.removeItem(STORAGE_KEYS.provider);
      localStorage.removeItem(STORAGE_KEYS.localUrl);
      localStorage.removeItem(STORAGE_KEYS.localModel);
    }
  });
  
  privacySaveButtons.addEventListener('change', () => {
    setPrivacySetting(STORAGE_KEYS.privacySaveButtons, privacySaveButtons.checked);
    if (privacySaveButtons.checked) {
      saveButtonCustomizations();
    } else {
      localStorage.removeItem(STORAGE_KEYS.buttonCustomizations);
    }
  });
  
  // Clear all data
  clearDataBtn.addEventListener('click', () => {
    if (confirm('Clear all saved data including API keys, prompts, and customizations?')) {
      clearAllData();
      apiKeyInput.value = '';
      customPromptInput.value = '';
      providerSelect.value = 'openai';
      localUrlInput.value = 'http://localhost:1234/v1/chat/completions';
      localModelInput.value = '';
      buttonCustomizations = {};
      updateLocalSettings();
      updateButtonAppearances();
      addMessage('system', 'All saved data cleared.', 'Privacy');
    }
  });
  
  // Save on change
  providerSelect.addEventListener('change', () => {
    saveIfAllowed(STORAGE_KEYS.provider, providerSelect.value, STORAGE_KEYS.privacySaveProvider);
    updateLocalSettings();
  });
  apiKeyInput.addEventListener('change', () => saveIfAllowed(STORAGE_KEYS.apiKey, apiKeyInput.value, STORAGE_KEYS.privacySaveKey));
  localUrlInput.addEventListener('change', () => saveIfAllowed(STORAGE_KEYS.localUrl, localUrlInput.value, STORAGE_KEYS.privacySaveProvider));
  localModelInput.addEventListener('change', () => saveIfAllowed(STORAGE_KEYS.localModel, localModelInput.value, STORAGE_KEYS.privacySaveProvider));
  customPromptInput.addEventListener('change', () => saveIfAllowed(STORAGE_KEYS.customPrompt, customPromptInput.value, STORAGE_KEYS.privacySavePrompt));
  
  // Send button state
  const updateSendButton = () => {
    const hasKey = apiKeyInput.value.trim().length > 0 || providerSelect.value === 'local';
    const hasInput = userInput.value.trim().length > 0;
    sendBtn.disabled = !(hasKey && hasInput);
  };
  apiKeyInput.addEventListener('input', updateSendButton);
  userInput.addEventListener('input', updateSendButton);
  providerSelect.addEventListener('change', updateSendButton);
  updateSendButton();
  
  // Writing style
  let savedWritingStyle = loadIfAllowed(STORAGE_KEYS.writingStyle, STORAGE_KEYS.privacySaveStyle, '');
  
  // System prompt builder
  const getSystemPrompt = () => {
    let prompt = `You are an AI writing assistant that helps edit Word documents. Available tools:
- delete_all_instances_of_text, replace_all_text, highlight_text, remove_highlight
- add_comment, insert_text_at_end, insert_text_at_start, get_document_content

ALWAYS use get_document_content first before making changes. Be concise. Default highlight color is yellow.`;
    if (savedWritingStyle) prompt += `\n\nUser's writing style:\n${savedWritingStyle}`;
    const custom = customPromptInput.value.trim();
    if (custom) prompt += `\n\nUser instructions:\n${custom}`;
    return prompt;
  };
  
  let conversation = [];
  const resetConversation = () => {
    conversation = [{ role: 'system', content: getSystemPrompt() }];
  };
  resetConversation();
  
  // Send message
  const sendMessage = async (text, isQuickAction = false) => {
    if (!text) return;
    
    const provider = providerSelect.value;
    const apiKey = apiKeyInput.value.trim();
    
    if (!apiKey && provider !== 'local') {
      addMessage('error', 'Please enter your API key in Settings.', 'Error');
      settingsPanel.classList.add('open');
      return;
    }
    
    saveIfAllowed(STORAGE_KEYS.apiKey, apiKey, STORAGE_KEYS.privacySaveKey);
    
    if (!isQuickAction) {
      userInput.value = '';
      updateSendButton();
    }
    
    addMessage('user', text.length > 200 ? text.slice(0, 200) + '...' : text, 'You');
    setSendEnabled(false);
    setQuickButtonsEnabled(false);
    setStatus('Thinking...');
    
    resetConversation();
    conversation.push({ role: 'user', content: text });
    
    try {
      let processing = true;
      while (processing) {
        const response = await callAI(provider, apiKey, conversation, localUrlInput.value, localModelInput.value);
        if (!response) throw new Error('No response from API');
        
        if (response.tool_calls?.length > 0) {
          conversation.push({ role: 'assistant', content: response.content || '', tool_calls: response.tool_calls });
          for (const tc of response.tool_calls) {
            const name = tc.function?.name;
            let args = {};
            try { args = JSON.parse(tc.function?.arguments || '{}'); } catch (e) {}
            setStatus(`Running: ${name}...`);
            addMessage('tool', `${name}(${Object.values(args).join(', ').slice(0, 40)}...)`, 'Tool');
            let result;
            try { result = await executeTool(name, args); } catch (e) { result = { success: false, message: e.message }; }
            conversation.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(result) });
            setStatus('Processing...');
          }
        } else {
          const aiResponse = response.content?.trim() || 'Done.';
          addMessage('assistant', aiResponse, 'AI Assistant');
          conversation.push({ role: 'assistant', content: response.content || '' });
          processing = false;
          setStatus('Ready');
          
          if (text.includes('analyze') && text.toLowerCase().includes('style')) {
            savedWritingStyle = aiResponse;
            saveIfAllowed(STORAGE_KEYS.writingStyle, savedWritingStyle, STORAGE_KEYS.privacySaveStyle);
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
  
  // Send button
  sendBtn.addEventListener('click', () => sendMessage(userInput.value.trim()));
  
  // Quick action buttons
  document.querySelectorAll('.quick-btn[data-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      if (isEditMode) {
        openEditModal(action);
      } else {
        const config = getButtonConfig(action);
        sendMessage(config.prompt, true);
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
  
  // Welcome message
  addMessage('system', 'Configure your AI in Settings. Use Edit to customize quick actions. Click the pop-out button to open in a new window.', 'Welcome');
});
