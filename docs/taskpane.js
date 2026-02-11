// Word AI Assistant - Professional Writing Tool
// Full-featured quick actions with comprehensive customization

const STORAGE = {
  apiKey: 'word-ai-key',
  provider: 'word-ai-provider',
  localUrl: 'word-ai-local-url',
  localModel: 'word-ai-local-model',
  customPrompt: 'word-ai-prompt',
  glossaryReplace: 'word-ai-glossary-replace',
  glossaryAvoid: 'word-ai-glossary-avoid',
  actionConfigs: 'word-ai-action-configs',
  userPresets: 'word-ai-user-presets',
  privacySaveKey: 'word-ai-priv-key',
  privacySavePresets: 'word-ai-priv-presets',
  contextAwareness: 'word-ai-context',
  versionHistory: 'word-ai-history'
};

const getWord = () => window.Word || window.Office?.Word;

// ==================== ACTIONS DEFINITION ====================
const ACTIONS = {
  analyze: { name: 'Analyze Writing Style', icon: '🔍', desc: 'Detect tone, voice, and patterns', base: 'Analyze the writing style: tone, sentence structure, vocabulary level, voice (active/passive), and patterns. Summarize the detected style.' },
  grammar: { name: 'Grammar Check', icon: '✓', desc: 'Find and fix grammar issues', base: 'Check for grammar errors. Highlight issues in yellow and add comments with corrections.' },
  spelling: { name: 'Spelling Check', icon: '📝', desc: 'Find spelling errors', base: 'Check for spelling errors. Highlight misspellings in red and add comments with corrections.' },
  formal: { name: 'Make Formal', icon: '👔', desc: 'Professional, formal tone', base: 'Rewrite in a formal, professional tone. Avoid contractions, use complete sentences.' },
  casual: { name: 'Make Casual', icon: '😊', desc: 'Friendly, conversational', base: 'Rewrite in a casual, conversational tone. Use contractions, everyday vocabulary.' },
  professional: { name: 'Professional Tone', icon: '💼', desc: 'Business appropriate', base: 'Adjust to professional business tone. Clear, direct, action-oriented.' },
  friendly: { name: 'Friendly Tone', icon: '🤝', desc: 'Warm and approachable', base: 'Rewrite with warm, friendly tone. Approachable and personable.' },
  clarity: { name: 'Improve Clarity', icon: '💡', desc: 'Clearer, easier to read', base: 'Improve clarity. Break up complex sentences, remove ambiguity, improve flow.' },
  concise: { name: 'Make Concise', icon: '✂️', desc: 'Remove redundancy', base: 'Make more concise. Remove redundant words, filler, unnecessary qualifiers.' },
  shorter: { name: 'Shorten', icon: '📉', desc: 'Reduce by 30-50%', base: 'Significantly shorten while keeping key points. Reduce by 30-50%.' },
  longer: { name: 'Expand', icon: '📈', desc: 'Add detail and examples', base: 'Expand and elaborate. Add detail, examples, and supporting information.' },
  suggestions: { name: 'Get Suggestions', icon: '💬', desc: 'Improvement ideas', base: 'Provide suggestions for improvement. Don\'t make changes, just advise.' },
  transform: { name: 'Transform Format', icon: '🔄', desc: 'Change structure', base: 'Transform the format as specified.' }
};

// Default config for each action
const DEFAULT_CONFIG = {
  tone: 'neutral',
  length: 'maintain',
  formality: 3,
  structure: 'paragraphs',
  complexity: 12,
  instructions: '',
  jargon: 'standard',
  voice: 'mixed',
  revision: 2,
  citation: 'none'
};

// Built-in presets
const PRESETS = {
  academic: { name: 'Academic Rigor', tone: 'authoritative', formality: 5, complexity: 16, voice: 'passive', jargon: 'high', citation: 'apa' },
  marketing: { name: 'Marketing Blast', tone: 'persuasive', formality: 2, structure: 'bullets', voice: 'active', jargon: 'none' },
  eli5: { name: 'ELI5 Explainer', tone: 'empathetic', formality: 1, complexity: 5, jargon: 'none', voice: 'active' },
  technical: { name: 'Technical Doc', tone: 'authoritative', formality: 4, complexity: 16, structure: 'numbered', jargon: 'high' }
};

// Live preview examples based on settings
const PREVIEW_EXAMPLES = {
  neutral: 'The system processes data efficiently.',
  persuasive: 'Experience lightning-fast data processing that transforms your workflow!',
  empathetic: 'We understand how important fast data processing is for your success.',
  authoritative: 'The system implements high-efficiency data processing protocols.',
  witty: 'This bad boy crunches numbers faster than you can say "spreadsheet."'
};

// ==================== STORAGE HELPERS ====================
const save = (key, val) => localStorage.setItem(key, typeof val === 'object' ? JSON.stringify(val) : val);
const load = (key, def = '') => {
  const v = localStorage.getItem(key);
  if (!v) return def;
  try { return JSON.parse(v); } catch { return v; }
};
const loadBool = (key, def = false) => localStorage.getItem(key) === 'true' || (localStorage.getItem(key) === null && def);

let actionConfigs = {};
let userPresets = {};
let versionHistory = [];

function loadAllData() {
  if (loadBool(STORAGE.privacySaveKey)) {
    document.getElementById('api-key').value = load(STORAGE.apiKey, '');
  }
  document.getElementById('provider').value = load(STORAGE.provider, 'openai');
  document.getElementById('local-url').value = load(STORAGE.localUrl, 'http://localhost:1234/v1/chat/completions');
  document.getElementById('local-model').value = load(STORAGE.localModel, '');
  document.getElementById('custom-prompt').value = load(STORAGE.customPrompt, '');
  document.getElementById('glossary-replace').value = load(STORAGE.glossaryReplace, '');
  document.getElementById('glossary-avoid').value = load(STORAGE.glossaryAvoid, '');
  document.getElementById('privacy-save-key').checked = loadBool(STORAGE.privacySaveKey);
  document.getElementById('privacy-save-presets').checked = loadBool(STORAGE.privacySavePresets, true);
  document.getElementById('context-awareness').checked = loadBool(STORAGE.contextAwareness, true);
  
  if (loadBool(STORAGE.privacySavePresets, true)) {
    actionConfigs = load(STORAGE.actionConfigs, {});
    userPresets = load(STORAGE.userPresets, {});
  }
  versionHistory = load(STORAGE.versionHistory, []);
}

function saveData(key, val) {
  if (key === STORAGE.apiKey && !loadBool(STORAGE.privacySaveKey)) return;
  if ((key === STORAGE.actionConfigs || key === STORAGE.userPresets) && !loadBool(STORAGE.privacySavePresets, true)) return;
  save(key, val);
}

// ==================== PROVIDER CONFIG ====================
const PROVIDERS = {
  openai: {
    url: 'https://api.openai.com/v1/chat/completions',
    format: (msgs, tools) => ({ model: 'gpt-4o-mini', messages: msgs, tools, tool_choice: 'auto' }),
    parse: d => d.choices?.[0]?.message,
    headers: k => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${k}` })
  },
  gemini: {
    url: k => `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${k}`,
    format: (msgs, tools) => {
      const contents = msgs.filter(m => m.role !== 'system').map(m => ({ role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content || '' }] }));
      const sys = msgs.find(m => m.role === 'system');
      return { contents, systemInstruction: sys ? { parts: [{ text: sys.content }] } : undefined, tools: tools ? [{ functionDeclarations: tools.map(t => ({ name: t.function.name, description: t.function.description, parameters: t.function.parameters })) }] : undefined };
    },
    parse: d => {
      const c = d.candidates?.[0]?.content?.parts || [];
      const txt = c.find(p => p.text);
      const fn = c.find(p => p.functionCall);
      if (fn) return { content: txt?.text || '', tool_calls: [{ id: 'g-' + Date.now(), function: { name: fn.functionCall.name, arguments: JSON.stringify(fn.functionCall.args || {}) } }] };
      return { content: txt?.text || '' };
    },
    headers: () => ({ 'Content-Type': 'application/json' })
  },
  claude: {
    url: 'https://api.anthropic.com/v1/messages',
    format: (msgs, tools) => {
      const sys = msgs.find(m => m.role === 'system');
      return { model: 'claude-3-haiku-20240307', max_tokens: 2048, system: sys?.content || '', messages: msgs.filter(m => m.role !== 'system').map(m => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: m.content || '' })), tools: tools?.map(t => ({ name: t.function.name, description: t.function.description, input_schema: t.function.parameters })) };
    },
    parse: d => {
      const txt = d.content?.find(c => c.type === 'text');
      const tool = d.content?.find(c => c.type === 'tool_use');
      if (tool) return { content: txt?.text || '', tool_calls: [{ id: tool.id, function: { name: tool.name, arguments: JSON.stringify(tool.input || {}) } }] };
      return { content: txt?.text || '' };
    },
    headers: k => ({ 'Content-Type': 'application/json', 'x-api-key': k, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' })
  },
  local: {
    format: (msgs, tools, model) => ({ model: model || 'local', messages: msgs, tools, tool_choice: 'auto' }),
    parse: d => d.choices?.[0]?.message,
    headers: k => ({ 'Content-Type': 'application/json', ...(k ? { 'Authorization': `Bearer ${k}` } : {}) })
  }
};

// ==================== TOOLS ====================
const TOOLS = [
  { 
    type: 'function', 
    function: { 
      name: 'get_document_content', 
      description: 'Read and return the current text content of the Word document. ALWAYS call this first before making any edits.', 
      parameters: { type: 'object', properties: {} } 
    } 
  },
  { 
    type: 'function', 
    function: { 
      name: 'replace_all_text', 
      description: 'Find all instances of searchText in the document and replace them with replaceText. Use this for making text changes, corrections, rewrites, etc.', 
      parameters: { 
        type: 'object', 
        properties: { 
          searchText: { type: 'string', description: 'The exact text to find in the document' }, 
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
      description: 'Highlight all instances of the specified text with a colored background. Use for marking errors, important text, or visual emphasis.', 
      parameters: { 
        type: 'object', 
        properties: { 
          searchText: { type: 'string', description: 'The text to highlight' }, 
          color: { type: 'string', enum: ['yellow', 'green', 'cyan', 'red', 'blue'], description: 'Highlight color: yellow (default), green, cyan, red, blue' } 
        }, 
        required: ['searchText', 'color'] 
      } 
    } 
  },
  { 
    type: 'function', 
    function: { 
      name: 'add_comment', 
      description: 'Add a comment/annotation to specific text in the document. The comment appears in the margin next to the text.', 
      parameters: { 
        type: 'object', 
        properties: { 
          searchText: { type: 'string', description: 'The text to attach the comment to' }, 
          comment: { type: 'string', description: 'The comment text to display in the margin' } 
        }, 
        required: ['searchText', 'comment'] 
      } 
    } 
  },
  { 
    type: 'function', 
    function: { 
      name: 'insert_text', 
      description: 'Insert new text at the beginning or end of the document.', 
      parameters: { 
        type: 'object', 
        properties: { 
          text: { type: 'string', description: 'The text to insert' }, 
          position: { type: 'string', enum: ['start', 'end'], description: 'Where to insert: start or end of document' } 
        }, 
        required: ['text', 'position'] 
      } 
    } 
  },
  { 
    type: 'function', 
    function: { 
      name: 'delete_text', 
      description: 'Delete all instances of the specified text from the document.', 
      parameters: { 
        type: 'object', 
        properties: { 
          searchText: { type: 'string', description: 'The exact text to delete' } 
        }, 
        required: ['searchText'] 
      } 
    } 
  }
];

const COLORS = { yellow: 'Yellow', green: 'BrightGreen', cyan: 'Turquoise', red: 'Red', blue: 'Blue' };

async function execTool(name, args) {
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word API not available. Make sure you are in Word.' };
  
  try {
    return await Word.run(async ctx => {
      const body = ctx.document.body;
      
      switch (name) {
        case 'get_document_content':
          body.load('text');
          await ctx.sync();
          const text = body.text || '';
          return { 
            success: true, 
            content: text.slice(0, 15000),
            message: `Read ${text.length} characters from document`
          };
          
        case 'replace_all_text': {
          if (!args.searchText) return { success: false, message: 'searchText is required' };
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          const count = results.items.length;
          if (count === 0) {
            return { success: true, count: 0, message: `"${args.searchText}" not found in document` };
          }
          for (let i = count - 1; i >= 0; i--) {
            results.items[i].insertText(args.replaceText || '', Word.InsertLocation.replace);
          }
          await ctx.sync();
          return { 
            success: true, 
            count: count, 
            message: `Replaced ${count} instance(s) of "${args.searchText}" with "${args.replaceText || '(deleted)'}"` 
          };
        }
        
        case 'highlight_text': {
          if (!args.searchText) return { success: false, message: 'searchText is required' };
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          const count = results.items.length;
          if (count === 0) {
            return { success: false, count: 0, message: `"${args.searchText}" not found in document` };
          }
          for (const item of results.items) {
            item.font.highlightColor = COLORS[args.color] || 'Yellow';
          }
          await ctx.sync();
          return { 
            success: true, 
            count: count, 
            message: `Highlighted ${count} instance(s) of "${args.searchText}" in ${args.color || 'yellow'}` 
          };
        }
        
        case 'add_comment': {
          if (!args.searchText) return { success: false, message: 'searchText is required' };
          if (!args.comment) return { success: false, message: 'comment text is required' };
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) {
            return { success: false, message: `"${args.searchText}" not found in document` };
          }
          results.items[0].insertComment(args.comment);
          await ctx.sync();
          return { 
            success: true, 
            message: `Added comment to "${args.searchText}": "${args.comment}"` 
          };
        }
        
        case 'insert_text': {
          if (!args.text) return { success: false, message: 'text is required' };
          const loc = args.position === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end;
          body.insertText(args.text, loc);
          await ctx.sync();
          return { 
            success: true, 
            message: `Inserted text at ${args.position || 'end'} of document` 
          };
        }
          
        case 'delete_text': {
          if (!args.searchText) return { success: false, message: 'searchText is required' };
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          const count = results.items.length;
          if (count === 0) {
            return { success: true, count: 0, message: `"${args.searchText}" not found in document` };
          }
          for (let i = count - 1; i >= 0; i--) {
            results.items[i].insertText('', Word.InsertLocation.replace);
          }
          await ctx.sync();
          return { 
            success: true, 
            count: count, 
            message: `Deleted ${count} instance(s) of "${args.searchText}"` 
          };
        }
        
        default:
          return { success: false, message: `Unknown tool: ${name}` };
      }
    });
  } catch (error) {
    console.error('Tool execution error:', error);
    return { success: false, message: `Error: ${error.message}` };
  }
}

// ==================== CONTEXT AWARENESS ====================
async function getDocumentContext() {
  if (!loadBool(STORAGE.contextAwareness, true)) return '';
  
  const Word = getWord();
  if (!Word) return '';
  
  try {
    return await Word.run(async ctx => {
      const selection = ctx.document.getSelection();
      const body = ctx.document.body;
      body.load('text');
      selection.load('text');
      await ctx.sync();
      
      const fullText = body.text || '';
      const selectedText = selection.text || '';
      
      // Find context around selection
      if (selectedText && fullText.includes(selectedText)) {
        const idx = fullText.indexOf(selectedText);
        const before = fullText.slice(Math.max(0, idx - 500), idx);
        const after = fullText.slice(idx + selectedText.length, idx + selectedText.length + 500);
        return `[CONTEXT BEFORE]: ${before}\n[SELECTED TEXT]: ${selectedText}\n[CONTEXT AFTER]: ${after}`;
      }
      
      return `[DOCUMENT EXCERPT]: ${fullText.slice(0, 2000)}`;
    });
  } catch {
    return '';
  }
}

// ==================== PROMPT BUILDER ====================
function buildPrompt(actionKey, config) {
  const action = ACTIONS[actionKey];
  const glossaryReplace = document.getElementById('glossary-replace').value.trim();
  const glossaryAvoid = document.getElementById('glossary-avoid').value.trim();
  const customPrompt = document.getElementById('custom-prompt').value.trim();
  
  let prompt = `You are an AI writing assistant. ${action.base}\n\n`;
  
  // Add parameters
  prompt += `PARAMETERS:\n`;
  prompt += `- Tone: ${config.tone}\n`;
  prompt += `- Target Length: ${config.length}\n`;
  prompt += `- Formality Level: ${config.formality}/5 (1=casual chat, 5=legal/academic)\n`;
  prompt += `- Output Structure: ${config.structure}\n`;
  prompt += `- Reading Level: Grade ${config.complexity}\n`;
  prompt += `- Jargon: ${config.jargon}\n`;
  prompt += `- Voice: ${config.voice}\n`;
  prompt += `- Revision Depth: ${['Surface (grammar only)', 'Structural (rewrite sentences)', 'Conceptual (reorganize ideas)'][config.revision - 1]}\n`;
  if (config.citation !== 'none') prompt += `- Citation Style: ${config.citation}\n`;
  
  // Glossary rules
  if (glossaryReplace) {
    prompt += `\nWORD REPLACEMENTS (always apply):\n`;
    glossaryReplace.split('\n').forEach(line => {
      const [old, neu] = line.split('→').map(s => s.trim());
      if (old && neu) prompt += `- Replace "${old}" with "${neu}"\n`;
    });
  }
  
  if (glossaryAvoid) {
    prompt += `\nWORDS TO AVOID (never use these):\n`;
    glossaryAvoid.split('\n').forEach(word => {
      if (word.trim()) prompt += `- ${word.trim()}\n`;
    });
  }
  
  // Custom instructions
  if (customPrompt) prompt += `\nGLOBAL INSTRUCTIONS:\n${customPrompt}\n`;
  if (config.instructions) prompt += `\nSPECIFIC INSTRUCTIONS:\n${config.instructions}\n`;
  
  prompt += `\nALWAYS use get_document_content first to read the document before making changes.`;
  
  return prompt;
}

// ==================== API CALL ====================
async function callAIWithMessages(messages) {
  const provider = document.getElementById('provider').value;
  const apiKey = document.getElementById('api-key').value.trim();
  const localUrl = document.getElementById('local-url').value.trim();
  const localModel = document.getElementById('local-model').value.trim();
  
  if (!apiKey && provider !== 'local') throw new Error('API key required');
  
  const cfg = PROVIDERS[provider];
  let url = typeof cfg.url === 'function' ? cfg.url(apiKey) : cfg.url;
  if (provider === 'local') url = localUrl;
  
  const resp = await fetch(url, {
    method: 'POST',
    headers: cfg.headers(apiKey),
    body: JSON.stringify(cfg.format(messages, TOOLS, localModel))
  });
  
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error?.message || `API error ${resp.status}`);
  }
  
  return cfg.parse(await resp.json());
}

// ==================== RUN ACTION ====================
async function runAction(actionKey, config) {
  setStatus('Reading document...');
  
  const context = await getDocumentContext();
  const systemPrompt = buildPrompt(actionKey, config);
  const userMessage = context 
    ? `Here is the document context:\n\n${context}\n\nPlease proceed with the action. Use the tools to make changes to the document.`
    : 'Please use get_document_content to read the document first, then proceed with the action using the appropriate tools.';
  
  let messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userMessage }
  ];
  
  try {
    let processing = true;
    let iterations = 0;
    const maxIterations = 15;
    
    while (processing && iterations < maxIterations) {
      iterations++;
      setStatus(`Processing (step ${iterations})...`);
      
      const response = await callAIWithMessages(messages);
      
      if (!response) {
        throw new Error('No response from AI');
      }
      
      if (response.tool_calls?.length > 0) {
        // Add assistant message with tool calls
        messages.push({ 
          role: 'assistant', 
          content: response.content || '', 
          tool_calls: response.tool_calls 
        });
        
        // Execute each tool call
        for (const tc of response.tool_calls) {
          const name = tc.function?.name;
          let args = {};
          try { args = JSON.parse(tc.function?.arguments || '{}'); } catch {}
          
          setStatus(`Executing: ${name}...`);
          addMessage('tool', `${name}(${JSON.stringify(args).slice(0, 80)}...)`, 'Tool');
          
          const result = await execTool(name, args);
          
          // Add tool result to messages
          messages.push({ 
            role: 'tool', 
            tool_call_id: tc.id, 
            content: JSON.stringify(result) 
          });
          
          // Log result
          if (result.success) {
            addMessage('system', `✓ ${result.message || name + ' completed'}`, 'Result');
          } else {
            addMessage('error', `✗ ${result.message || 'Failed'}`, 'Result');
          }
          
          // Save to history if it was a content change
          if (name === 'replace_all_text' && result.success && result.count > 0) {
            addToHistory(actionKey, args.searchText, args.replaceText);
          }
        }
        
        // Continue the loop to let AI process results and potentially make more calls
      } else {
        // No more tool calls - AI is done
        if (response.content?.trim()) {
          addMessage('assistant', response.content.trim(), 'AI');
        } else {
          addMessage('assistant', 'Done! Changes have been applied to your document.', 'AI');
        }
        processing = false;
      }
    }
    
    if (iterations >= maxIterations) {
      addMessage('system', 'Reached maximum iterations. Some changes may be incomplete.', 'Warning');
    }
    
    setStatus('Ready');
  } catch (e) {
    addMessage('error', `Error: ${e.message}`, 'Error');
    setStatus('Error');
    console.error(e);
  }
}

// ==================== VERSION HISTORY ====================
function addToHistory(action, original, replacement) {
  versionHistory.unshift({
    id: Date.now(),
    action,
    original,
    replacement,
    timestamp: new Date().toLocaleTimeString()
  });
  if (versionHistory.length > 50) versionHistory.pop();
  save(STORAGE.versionHistory, versionHistory);
}

function renderHistory() {
  const list = document.getElementById('history-list');
  if (versionHistory.length === 0) {
    list.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:20px;">No history yet.</p>';
    return;
  }
  
  list.innerHTML = versionHistory.map(h => `
    <div style="padding:10px;border-bottom:1px solid var(--border);">
      <div style="font-size:10px;color:var(--text-muted);">${h.timestamp} - ${ACTIONS[h.action]?.name || h.action}</div>
      <div style="font-size:11px;margin-top:4px;"><strong>Original:</strong> ${(h.original || '').slice(0, 100)}...</div>
      <div style="font-size:11px;"><strong>Changed to:</strong> ${(h.replacement || '').slice(0, 100)}...</div>
      <button onclick="revertHistory('${h.id}')" style="margin-top:6px;padding:4px 10px;background:var(--primary);color:white;border:none;border-radius:4px;cursor:pointer;font-size:10px;">Revert</button>
    </div>
  `).join('');
}

window.revertHistory = async function(id) {
  const item = versionHistory.find(h => h.id === parseInt(id));
  if (!item) return;
  
  try {
    await execTool('replace_all_text', { searchText: item.replacement, replaceText: item.original });
    addMessage('system', `Reverted: "${item.replacement}" → "${item.original}"`, 'History');
  } catch (e) {
    addMessage('error', e.message, 'Error');
  }
};

// ==================== UI ====================
function addMessage(type, content, label = '') {
  const msgs = document.getElementById('messages');
  const div = document.createElement('div');
  div.className = `message ${type}`;
  if (label) {
    const lbl = document.createElement('div');
    lbl.className = 'message-label';
    lbl.textContent = label;
    div.appendChild(lbl);
  }
  const txt = document.createElement('div');
  txt.textContent = content;
  div.appendChild(txt);
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
}

function setStatus(text) {
  document.getElementById('status').textContent = text;
}

function renderActionsList() {
  const list = document.getElementById('quick-actions-list');
  list.innerHTML = Object.entries(ACTIONS).map(([key, action]) => {
    const hasCustom = actionConfigs[key] && Object.keys(actionConfigs[key]).length > 0;
    return `
      <div class="quick-action-item ${hasCustom ? 'customized' : ''}" data-action="${key}">
        <span class="qa-icon">${action.icon}</span>
        <div class="qa-info">
          <div class="qa-name">${action.name}</div>
          <div class="qa-desc">${action.desc}</div>
        </div>
        <div class="qa-buttons">
          <button class="qa-btn settings" data-action="${key}" data-mode="settings" title="Configure">⚙️</button>
          <button class="qa-btn run" data-action="${key}" data-mode="run">Run</button>
        </div>
      </div>
    `;
  }).join('');
}

function updatePreview() {
  const tone = document.getElementById('modal-tone').value;
  document.getElementById('preview-result').textContent = PREVIEW_EXAMPLES[tone] || PREVIEW_EXAMPLES.neutral;
}

// ==================== MODAL HANDLING ====================
let currentAction = null;
let currentConfig = { ...DEFAULT_CONFIG };

function openModal(actionKey) {
  currentAction = actionKey;
  const action = ACTIONS[actionKey];
  currentConfig = { ...DEFAULT_CONFIG, ...(actionConfigs[actionKey] || {}) };
  
  document.getElementById('modal-icon').textContent = action.icon;
  document.getElementById('modal-action-name').textContent = action.name;
  
  // Set values
  document.getElementById('modal-tone').value = currentConfig.tone;
  document.getElementById('modal-formality').value = currentConfig.formality;
  document.getElementById('modal-instructions').value = currentConfig.instructions;
  document.getElementById('modal-citation').value = currentConfig.citation;
  document.getElementById('modal-revision').value = currentConfig.revision;
  
  // Set segmented controls
  setSegmented('modal-length', currentConfig.length);
  setSegmented('modal-complexity', currentConfig.complexity.toString());
  setSegmented('modal-jargon', currentConfig.jargon);
  setSegmented('modal-voice', currentConfig.voice);
  
  // Set icon buttons
  setIconBtn('modal-structure', currentConfig.structure);
  
  updateCharCount();
  updatePreview();
  
  document.getElementById('action-modal').classList.add('open');
}

function closeModal() {
  document.getElementById('action-modal').classList.remove('open');
  currentAction = null;
}

function setSegmented(id, value) {
  document.querySelectorAll(`#${id} .segmented-btn`).forEach(btn => {
    btn.classList.toggle('active', btn.dataset.value === value);
  });
}

function getSegmented(id) {
  return document.querySelector(`#${id} .segmented-btn.active`)?.dataset.value || '';
}

function setIconBtn(id, value) {
  document.querySelectorAll(`#${id} .icon-btn`).forEach(btn => {
    btn.classList.toggle('active', btn.dataset.value === value);
  });
}

function getIconBtn(id) {
  return document.querySelector(`#${id} .icon-btn.active`)?.dataset.value || 'paragraphs';
}

function updateCharCount() {
  const input = document.getElementById('modal-instructions');
  const counter = document.getElementById('modal-char-count');
  const len = input.value.length;
  counter.textContent = `${len} / 280`;
  counter.className = 'char-counter' + (len > 250 ? (len > 280 ? ' error' : ' warn') : '');
}

function getConfigFromModal() {
  return {
    tone: document.getElementById('modal-tone').value,
    length: getSegmented('modal-length'),
    formality: parseInt(document.getElementById('modal-formality').value),
    structure: getIconBtn('modal-structure'),
    complexity: parseInt(getSegmented('modal-complexity')),
    instructions: document.getElementById('modal-instructions').value.slice(0, 280),
    jargon: getSegmented('modal-jargon'),
    voice: getSegmented('modal-voice'),
    revision: parseInt(document.getElementById('modal-revision').value),
    citation: document.getElementById('modal-citation').value
  };
}

function applyPreset(presetKey) {
  const preset = PRESETS[presetKey] || userPresets[presetKey];
  if (!preset) return;
  
  Object.entries(preset).forEach(([k, v]) => {
    if (k === 'name') return;
    currentConfig[k] = v;
  });
  
  // Update UI
  document.getElementById('modal-tone').value = currentConfig.tone || 'neutral';
  document.getElementById('modal-formality').value = currentConfig.formality || 3;
  setSegmented('modal-length', currentConfig.length || 'maintain');
  setSegmented('modal-complexity', (currentConfig.complexity || 12).toString());
  setSegmented('modal-jargon', currentConfig.jargon || 'standard');
  setSegmented('modal-voice', currentConfig.voice || 'mixed');
  setIconBtn('modal-structure', currentConfig.structure || 'paragraphs');
  document.getElementById('modal-revision').value = currentConfig.revision || 2;
  document.getElementById('modal-citation').value = currentConfig.citation || 'none';
  
  updatePreview();
}

// ==================== INIT ====================
Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = '<p style="padding:20px;">This add-in requires Microsoft Word.</p>';
    return;
  }
  
  loadAllData();
  renderActionsList();
  
  // Tab switching
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      document.querySelector(`.tab-content[data-tab="${tab.dataset.tab}"]`).classList.add('active');
    });
  });
  
  // Quick action buttons
  document.getElementById('quick-actions-list').addEventListener('click', e => {
    const btn = e.target.closest('.qa-btn');
    if (!btn) return;
    
    const action = btn.dataset.action;
    if (btn.dataset.mode === 'settings') {
      openModal(action);
    } else {
      const config = { ...DEFAULT_CONFIG, ...(actionConfigs[action] || {}) };
      runAction(action, config);
      document.querySelector('.tab[data-tab="chat"]').click();
    }
  });
  
  // Modal controls
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-cancel').addEventListener('click', closeModal);
  document.getElementById('action-modal').addEventListener('click', e => {
    if (e.target.id === 'action-modal') closeModal();
  });
  
  document.getElementById('modal-apply').addEventListener('click', () => {
    const config = getConfigFromModal();
    actionConfigs[currentAction] = config;
    saveData(STORAGE.actionConfigs, actionConfigs);
    renderActionsList();
    closeModal();
    runAction(currentAction, config);
    document.querySelector('.tab[data-tab="chat"]').click();
  });
  
  document.getElementById('modal-reset').addEventListener('click', () => {
    delete actionConfigs[currentAction];
    saveData(STORAGE.actionConfigs, actionConfigs);
    currentConfig = { ...DEFAULT_CONFIG };
    openModal(currentAction); // Refresh modal
    renderActionsList();
  });
  
  // Preset handling
  document.getElementById('modal-preset').addEventListener('change', e => {
    if (e.target.value) applyPreset(e.target.value);
    e.target.value = '';
  });
  
  document.getElementById('modal-save-preset').addEventListener('click', () => {
    const name = prompt('Preset name:');
    if (!name) return;
    const config = getConfigFromModal();
    userPresets[name.toLowerCase().replace(/\s+/g, '_')] = { name, ...config };
    saveData(STORAGE.userPresets, userPresets);
    
    // Add to dropdown
    const opt = document.createElement('option');
    opt.value = name.toLowerCase().replace(/\s+/g, '_');
    opt.textContent = name + ' (Custom)';
    document.getElementById('modal-preset').appendChild(opt);
  });
  
  // Segmented controls
  document.querySelectorAll('.segmented').forEach(seg => {
    seg.addEventListener('click', e => {
      const btn = e.target.closest('.segmented-btn');
      if (!btn) return;
      seg.querySelectorAll('.segmented-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });
  
  // Icon buttons
  document.querySelectorAll('.icon-row').forEach(row => {
    row.addEventListener('click', e => {
      const btn = e.target.closest('.icon-btn');
      if (!btn) return;
      row.querySelectorAll('.icon-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });
  
  // Accordion
  document.querySelectorAll('.accordion-header').forEach(header => {
    header.addEventListener('click', () => {
      header.parentElement.classList.toggle('open');
    });
  });
  
  // Character counter
  document.getElementById('modal-instructions').addEventListener('input', updateCharCount);
  
  // Tone preview
  document.getElementById('modal-tone').addEventListener('change', updatePreview);
  
  // Provider change
  document.getElementById('provider').addEventListener('change', () => {
    const isLocal = document.getElementById('provider').value === 'local';
    document.getElementById('local-settings').style.display = isLocal ? 'block' : 'none';
    saveData(STORAGE.provider, document.getElementById('provider').value);
  });
  document.getElementById('local-settings').style.display = document.getElementById('provider').value === 'local' ? 'block' : 'none';
  
  // Save settings on change
  document.getElementById('api-key').addEventListener('change', () => saveData(STORAGE.apiKey, document.getElementById('api-key').value));
  document.getElementById('local-url').addEventListener('change', () => saveData(STORAGE.localUrl, document.getElementById('local-url').value));
  document.getElementById('local-model').addEventListener('change', () => saveData(STORAGE.localModel, document.getElementById('local-model').value));
  document.getElementById('custom-prompt').addEventListener('change', () => saveData(STORAGE.customPrompt, document.getElementById('custom-prompt').value));
  document.getElementById('glossary-replace').addEventListener('change', () => saveData(STORAGE.glossaryReplace, document.getElementById('glossary-replace').value));
  document.getElementById('glossary-avoid').addEventListener('change', () => saveData(STORAGE.glossaryAvoid, document.getElementById('glossary-avoid').value));
  
  // Privacy toggles
  document.getElementById('privacy-save-key').addEventListener('change', e => {
    save(STORAGE.privacySaveKey, e.target.checked);
    if (!e.target.checked) localStorage.removeItem(STORAGE.apiKey);
  });
  document.getElementById('privacy-save-presets').addEventListener('change', e => {
    save(STORAGE.privacySavePresets, e.target.checked);
    if (!e.target.checked) {
      localStorage.removeItem(STORAGE.actionConfigs);
      localStorage.removeItem(STORAGE.userPresets);
    }
  });
  document.getElementById('context-awareness').addEventListener('change', e => {
    save(STORAGE.contextAwareness, e.target.checked);
  });
  
  // Clear data
  document.getElementById('clear-data').addEventListener('click', () => {
    if (confirm('Clear all saved data?')) {
      Object.values(STORAGE).forEach(k => localStorage.removeItem(k));
      location.reload();
    }
  });
  
  // Export settings
  document.getElementById('export-settings').addEventListener('click', () => {
    const exportData = {
      version: '1.0',
      exportDate: new Date().toISOString(),
      settings: {
        apiKey: document.getElementById('api-key').value,
        provider: document.getElementById('provider').value,
        localUrl: document.getElementById('local-url').value,
        localModel: document.getElementById('local-model').value,
        customPrompt: document.getElementById('custom-prompt').value,
        glossaryReplace: document.getElementById('glossary-replace').value,
        glossaryAvoid: document.getElementById('glossary-avoid').value,
        contextAwareness: document.getElementById('context-awareness').checked
      },
      actionConfigs: actionConfigs,
      userPresets: userPresets
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `word-ai-settings-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    document.getElementById('import-status').innerHTML = '<span style="color:var(--success);">✓ Settings exported successfully!</span>';
    setTimeout(() => document.getElementById('import-status').innerHTML = '', 3000);
  });
  
  // Import settings
  document.getElementById('import-settings').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const data = JSON.parse(event.target.result);
        
        if (!data.settings) {
          throw new Error('Invalid settings file format');
        }
        
        // Apply settings
        if (data.settings.apiKey) {
          document.getElementById('api-key').value = data.settings.apiKey;
          save(STORAGE.apiKey, data.settings.apiKey);
        }
        if (data.settings.provider) {
          document.getElementById('provider').value = data.settings.provider;
          save(STORAGE.provider, data.settings.provider);
          document.getElementById('local-settings').style.display = data.settings.provider === 'local' ? 'block' : 'none';
        }
        if (data.settings.localUrl) {
          document.getElementById('local-url').value = data.settings.localUrl;
          save(STORAGE.localUrl, data.settings.localUrl);
        }
        if (data.settings.localModel) {
          document.getElementById('local-model').value = data.settings.localModel;
          save(STORAGE.localModel, data.settings.localModel);
        }
        if (data.settings.customPrompt !== undefined) {
          document.getElementById('custom-prompt').value = data.settings.customPrompt;
          save(STORAGE.customPrompt, data.settings.customPrompt);
        }
        if (data.settings.glossaryReplace !== undefined) {
          document.getElementById('glossary-replace').value = data.settings.glossaryReplace;
          save(STORAGE.glossaryReplace, data.settings.glossaryReplace);
        }
        if (data.settings.glossaryAvoid !== undefined) {
          document.getElementById('glossary-avoid').value = data.settings.glossaryAvoid;
          save(STORAGE.glossaryAvoid, data.settings.glossaryAvoid);
        }
        if (data.settings.contextAwareness !== undefined) {
          document.getElementById('context-awareness').checked = data.settings.contextAwareness;
          save(STORAGE.contextAwareness, data.settings.contextAwareness);
        }
        
        // Apply action configs
        if (data.actionConfigs) {
          actionConfigs = data.actionConfigs;
          save(STORAGE.actionConfigs, actionConfigs);
          renderActionsList();
        }
        
        // Apply user presets
        if (data.userPresets) {
          userPresets = data.userPresets;
          save(STORAGE.userPresets, userPresets);
          
          // Add presets to dropdown
          const presetSelect = document.getElementById('modal-preset');
          Object.entries(userPresets).forEach(([key, preset]) => {
            if (!presetSelect.querySelector(`option[value="${key}"]`)) {
              const opt = document.createElement('option');
              opt.value = key;
              opt.textContent = preset.name + ' (Custom)';
              presetSelect.appendChild(opt);
            }
          });
        }
        
        document.getElementById('import-status').innerHTML = '<span style="color:var(--success);">✓ Settings imported successfully!</span>';
        setTimeout(() => document.getElementById('import-status').innerHTML = '', 3000);
        
      } catch (err) {
        document.getElementById('import-status').innerHTML = `<span style="color:var(--danger);">✗ Error: ${err.message}</span>`;
      }
    };
    
    reader.readAsText(file);
    e.target.value = ''; // Reset file input
  });
  
  // History modal
  document.getElementById('history-btn').addEventListener('click', () => {
    renderHistory();
    document.getElementById('history-modal').classList.add('open');
  });
  document.getElementById('history-close').addEventListener('click', () => {
    document.getElementById('history-modal').classList.remove('open');
  });
  document.getElementById('history-done').addEventListener('click', () => {
    document.getElementById('history-modal').classList.remove('open');
  });
  document.getElementById('history-clear').addEventListener('click', () => {
    if (confirm('Clear all history?')) {
      versionHistory = [];
      save(STORAGE.versionHistory, []);
      renderHistory();
    }
  });
  
  // Pop-out
  document.getElementById('popout-btn').addEventListener('click', () => {
    window.open(location.href, 'WordAI', 'width=450,height=700,resizable=yes');
  });
  
  // Chat send
  const sendBtn = document.getElementById('send-btn');
  const userInput = document.getElementById('user-input');
  
  const updateSendBtn = () => {
    const hasKey = document.getElementById('api-key').value.trim() || document.getElementById('provider').value === 'local';
    sendBtn.disabled = !hasKey || !userInput.value.trim();
  };
  
  userInput.addEventListener('input', updateSendBtn);
  document.getElementById('api-key').addEventListener('input', updateSendBtn);
  
  sendBtn.addEventListener('click', async () => {
    const text = userInput.value.trim();
    if (!text) return;
    
    userInput.value = '';
    updateSendBtn();
    addMessage('user', text, 'You');
    
    const customPrompt = document.getElementById('custom-prompt').value;
    const glossaryReplace = document.getElementById('glossary-replace').value;
    const glossaryAvoid = document.getElementById('glossary-avoid').value;
    
    let systemPrompt = `You are an AI writing assistant that helps edit Word documents. You have access to these tools:
- get_document_content: Read the document text
- replace_all_text: Find and replace text
- highlight_text: Highlight text with colors (yellow, green, cyan, red, blue)
- add_comment: Add a comment to specific text
- insert_text: Insert text at start or end of document
- delete_text: Delete all instances of text

IMPORTANT: You MUST use these tools to make changes to the document. The user cannot see your text responses as document changes - you must use the tools.

When the user asks you to edit, highlight, comment, or modify the document:
1. First use get_document_content to read what's in the document
2. Then use the appropriate tools to make the changes
3. Confirm what you did

`;
    
    if (glossaryReplace) {
      systemPrompt += `\nWORD REPLACEMENTS (always apply):\n`;
      glossaryReplace.split('\n').forEach(line => {
        const [old, neu] = line.split('→').map(s => s.trim());
        if (old && neu) systemPrompt += `- Replace "${old}" with "${neu}"\n`;
      });
    }
    
    if (glossaryAvoid) {
      systemPrompt += `\nWORDS TO AVOID:\n`;
      glossaryAvoid.split('\n').forEach(word => {
        if (word.trim()) systemPrompt += `- ${word.trim()}\n`;
      });
    }
    
    if (customPrompt) {
      systemPrompt += `\nADDITIONAL INSTRUCTIONS:\n${customPrompt}\n`;
    }
    
    let messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: text }
    ];
    
    try {
      let processing = true;
      let iterations = 0;
      const maxIterations = 15;
      
      while (processing && iterations < maxIterations) {
        iterations++;
        setStatus(`Processing (step ${iterations})...`);
        
        const response = await callAIWithMessages(messages);
        
        if (!response) {
          throw new Error('No response from AI');
        }
        
        if (response.tool_calls?.length > 0) {
          messages.push({ 
            role: 'assistant', 
            content: response.content || '', 
            tool_calls: response.tool_calls 
          });
          
          for (const tc of response.tool_calls) {
            const name = tc.function?.name;
            let args = {};
            try { args = JSON.parse(tc.function?.arguments || '{}'); } catch {}
            
            setStatus(`Executing: ${name}...`);
            addMessage('tool', `${name}(${JSON.stringify(args).slice(0, 80)}...)`, 'Tool');
            
            const result = await execTool(name, args);
            
            messages.push({ 
              role: 'tool', 
              tool_call_id: tc.id, 
              content: JSON.stringify(result) 
            });
            
            if (result.success) {
              addMessage('system', `✓ ${result.message || name + ' completed'}${result.count !== undefined ? ` (${result.count} items)` : ''}`, 'Result');
            } else {
              addMessage('error', `✗ ${result.message || 'Failed'}`, 'Result');
            }
            
            if (name === 'replace_all_text' && result.success && result.count > 0) {
              addToHistory('chat', args.searchText, args.replaceText);
            }
          }
        } else {
          if (response.content?.trim()) {
            addMessage('assistant', response.content.trim(), 'AI');
          } else {
            addMessage('assistant', 'Done!', 'AI');
          }
          processing = false;
        }
      }
      
      setStatus('Ready');
    } catch (e) {
      addMessage('error', `Error: ${e.message}`, 'Error');
      setStatus('Error');
      console.error(e);
    }
  });
  
  userInput.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendBtn.click();
    }
  });
  
  addMessage('system', 'Select a Quick Action or use Chat. Configure each action with the ⚙️ button.', 'Welcome');
});
