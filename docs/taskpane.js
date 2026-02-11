// Word AI Assistant - Full Word Functionality
// Complete control over Word documents via AI

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
  analyze: { name: 'Analyze Writing Style', icon: '🔍', desc: 'Detect tone, voice, and patterns', base: 'Analyze the writing style and summarize findings.' },
  grammar: { name: 'Grammar Check', icon: '✓', desc: 'Find and fix grammar issues', base: 'Check for grammar errors. Highlight issues and add comments with corrections.' },
  spelling: { name: 'Spelling Check', icon: '📝', desc: 'Find spelling errors', base: 'Check for spelling errors. Highlight and comment with corrections.' },
  formal: { name: 'Make Formal', icon: '👔', desc: 'Professional tone', base: 'Rewrite in formal, professional tone.' },
  casual: { name: 'Make Casual', icon: '😊', desc: 'Conversational', base: 'Rewrite in casual, conversational tone.' },
  clarity: { name: 'Improve Clarity', icon: '💡', desc: 'Easier to read', base: 'Improve clarity and readability.' },
  concise: { name: 'Make Concise', icon: '✂️', desc: 'Remove fluff', base: 'Make more concise without losing meaning.' },
  format: { name: 'Auto Format', icon: '🎨', desc: 'Apply formatting', base: 'Apply appropriate formatting, headings, and structure.' },
  
  // === STYLE GUIDE BUTTONS ===
  occStyle: { 
    name: 'OCC Style Format', 
    icon: '🏥', 
    desc: 'Apply VA OCC style guide', 
    base: `Apply Office of Connected Care (OCC) Writing Style Guide formatting to this document.

KEY OCC STYLE RULES TO APPLY:

TERMINOLOGY CORRECTIONS (replace these):
- "healthcare" → "health care" (two words, except in official names)
- "the VA" → "VA" (never use "the" before VA)
- "Veterans Administration" → "U.S. Department of Veterans Affairs" or "VA"
- "patient" → "Veteran" (in Veteran-facing content)
- "encounter" → "episode of care" (for telehealth)
- "click" → "select" (preferred action verb)
- "log in" → "sign in" (preferred)
- "login" → "sign-in" (as modifier)
- "doctor" → "provider" or "health care professional"
- "clinical team" or "health care team" → "VA care team"
- "teleprovider" → DO NOT USE (use "provider" or "health care professional")

CAPITALIZATION:
- "Veteran" - always capitalize the V
- "Nation's Veterans" - capitalize when used together
- "Service member" - capital S
- Telespecialties: TeleDermatology, TeleAudiology, TeleMental Health, TeleWound Care, etc.
- VA Video Connect (never abbreviate as VVC except for "VVC Now")
- My HealtheVet (two words, never abbreviate as MHV)
- Connected Care Academy (not "learning management system")

ABBREVIATIONS:
- Spell out on first reference, abbreviate in follow-up
- Do NOT put abbreviation in parentheses after first spelling: "the U.S. Department of Veterans Affairs" then "VA" later
- VISN does not need to be spelled out
- VistA does not need to be spelled out

FORMATTING:
- Phone numbers: xxx-xxx-xxxx format with hyphens
- "fiscal year" (spell out, FY acceptable in tables)
- "internet" (lowercase)
- "email" (no hyphen)
- "website" (one word, lowercase)
- "app" not "application" (except "WebVRAM application")

Make all necessary replacements throughout the document.` 
  },
  
  occCheck: { 
    name: 'OCC Compliance Check', 
    icon: '✅', 
    desc: 'Check OCC style compliance', 
    base: `Review this document for Office of Connected Care (OCC) Style Guide compliance. For each violation found, highlight the text in yellow and add a comment explaining the issue and correction.

CHECK FOR THESE COMMON VIOLATIONS:

1. TERMINOLOGY ERRORS:
   - "healthcare" should be "health care"
   - "the VA" should be just "VA"
   - "Veterans Administration" is incorrect
   - "click" should be "select"
   - "log in/login" should be "sign in/sign-in"
   - "doctor" should be "provider" or "health care professional"
   - "encounter" should be "episode of care" (for telehealth)
   - "clinical team" should be "VA care team"

2. CAPITALIZATION ERRORS:
   - "veteran" should be "Veteran"
   - Telespecialties must be capitalized: TeleDermatology, TeleAudiology, etc.
   - "My healthevet" should be "My HealtheVet"

3. ABBREVIATION ERRORS:
   - VVC used (should spell out "VA Video Connect")
   - MHV used (should spell out "My HealtheVet")
   - Abbreviation in parentheses after first use (remove parentheses)

4. FORMATTING ERRORS:
   - Phone numbers not in xxx-xxx-xxxx format
   - "Internet" capitalized (should be lowercase)
   - "e-mail" hyphenated (should be "email")
   - "web site" as two words (should be "website")

5. PROHIBITED TERMS:
   - "teleprovider" (use "provider" instead)
   - "our Veterans" (say "Veterans" or "Nation's Veterans")
   - "Clinical Video Telehealth" (now "Synchronous Telehealth")
   - "Store-and-Forward Telehealth" (now "Asynchronous Telehealth")

Highlight each violation and add a comment with the correct usage.` 
  },
  
  apStyle: { 
    name: 'AP Style Format', 
    icon: '📰', 
    desc: 'Apply AP Stylebook rules', 
    base: `Apply Associated Press (AP) Stylebook formatting rules to this document.

KEY AP STYLE RULES TO APPLY:

NUMBERS:
- Spell out one through nine
- Use numerals for 10 and above
- Always use numerals for ages: "a 5-year-old child"
- Always use numerals for percentages: "5 percent" (spell out "percent")
- Money: $1 million (not $1,000,000), $500,000
- Dimensions: 5 feet 6 inches tall

DATES AND TIMES:
- Abbreviate months with dates: Jan. 15, Feb. 20, March 5 (spell out March, April, May, June, July)
- No "th" or "nd": "Jan. 15" not "Jan. 15th"
- Times: 9 a.m., 4:30 p.m. (lowercase with periods)
- Use "noon" and "midnight" not "12 p.m." or "12 a.m."
- Years: 2024 (no apostrophe for decades: 1990s not 1990's)

TITLES:
- Capitalize formal titles before names: President Biden, Dr. Smith
- Lowercase after names or standing alone: Joe Biden, president of the United States
- Abbreviate certain titles before names: Dr., Gov., Lt. Gov., Rep., Sen., Rev.

PUNCTUATION:
- NO Oxford comma (serial comma) unless needed for clarity
- Periods and commas go INSIDE quotation marks
- Colons and semicolons go OUTSIDE quotation marks
- Single space after periods

WORDS AND PHRASES:
- "health care" (two words)
- "email" (no hyphen)
- "internet" (lowercase)
- "website" (one word, lowercase)
- "cellphone" (one word)
- "fundraising, fundraiser" (one word)
- "percent" (spell out, not %)
- "their" as singular (for gender-neutral)

ABBREVIATIONS:
- Spell out on first reference
- U.S. as adjective: "U.S. government"
- United States as noun: "in the United States"
- State abbreviations after city names (use AP abbreviations, not postal): Ala., Ariz., Ark., Calif., Colo., Conn., Del., Fla., Ga., Ill., Ind., Kan., Ky., La., Md., Mass., Mich., Minn., Miss., Mo., Mont., Neb., Nev., N.H., N.J., N.M., N.Y., N.C., N.D., Okla., Ore., Pa., R.I., S.C., S.D., Tenn., Vt., Va., Wash., W.Va., Wis., Wyo.
- Spell out Alaska, Hawaii, Idaho, Iowa, Maine, Ohio, Texas, Utah

ACADEMIC DEGREES:
- Use periods: Ph.D., M.A., B.A., M.D.
- Offset with commas: "John Smith, Ph.D., spoke at the event."

HYPHENS:
- Compound modifiers before nouns: "full-time employee" but "works full time"
- No hyphen with -ly adverbs: "newly elected official"

Make all necessary corrections throughout the document.` 
  }
};

const DEFAULT_CONFIG = {
  tone: 'neutral', length: 'maintain', formality: 3, structure: 'paragraphs',
  complexity: 12, instructions: '', jargon: 'standard', voice: 'mixed', revision: 2, citation: 'none'
};

const PRESETS = {
  academic: { name: 'Academic', tone: 'authoritative', formality: 5, complexity: 16, voice: 'passive', jargon: 'high' },
  marketing: { name: 'Marketing', tone: 'persuasive', formality: 2, structure: 'bullets', voice: 'active', jargon: 'none' },
  eli5: { name: 'ELI5', tone: 'empathetic', formality: 1, complexity: 5, jargon: 'none', voice: 'active' }
};

const PREVIEW_EXAMPLES = {
  neutral: 'The system processes data efficiently.',
  persuasive: 'Experience lightning-fast data processing!',
  empathetic: 'We understand how important fast processing is.',
  authoritative: 'The system implements high-efficiency protocols.',
  witty: 'This bad boy crunches numbers faster than you can blink.'
};

// ==================== STORAGE ====================
const save = (key, val) => localStorage.setItem(key, typeof val === 'object' ? JSON.stringify(val) : val);
const load = (key, def = '') => { const v = localStorage.getItem(key); if (!v) return def; try { return JSON.parse(v); } catch { return v; } };
const loadBool = (key, def = false) => localStorage.getItem(key) === 'true' || (localStorage.getItem(key) === null && def);

let actionConfigs = {};
let userPresets = {};
let versionHistory = [];

function loadAllData() {
  if (loadBool(STORAGE.privacySaveKey)) document.getElementById('api-key').value = load(STORAGE.apiKey, '');
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

// ==================== PROVIDERS ====================
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
      return { model: 'claude-3-haiku-20240307', max_tokens: 4096, system: sys?.content || '', messages: msgs.filter(m => m.role !== 'system').map(m => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: m.content || '' })), tools: tools?.map(t => ({ name: t.function.name, description: t.function.description, input_schema: t.function.parameters })) };
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

// ==================== COMPREHENSIVE WORD TOOLS ====================
const TOOLS = [
  // === DOCUMENT READING ===
  { type: 'function', function: { name: 'get_document_content', description: 'Read the entire document text. ALWAYS call this first before making edits.', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'get_selection', description: 'Get the currently selected text in the document.', parameters: { type: 'object', properties: {} } } },
  
  // === TEXT EDITING ===
  { type: 'function', function: { name: 'replace_text', description: 'Find and replace text throughout the document.', parameters: { type: 'object', properties: { find: { type: 'string', description: 'Text to find' }, replace: { type: 'string', description: 'Replacement text' }, matchCase: { type: 'boolean', description: 'Case sensitive match' } }, required: ['find', 'replace'] } } },
  { type: 'function', function: { name: 'insert_text', description: 'Insert text at a position (start, end, or replace selection).', parameters: { type: 'object', properties: { text: { type: 'string' }, position: { type: 'string', enum: ['start', 'end', 'replace_selection'] } }, required: ['text', 'position'] } } },
  { type: 'function', function: { name: 'delete_text', description: 'Delete all instances of specific text.', parameters: { type: 'object', properties: { text: { type: 'string' } }, required: ['text'] } } },
  
  // === FONT FORMATTING (Home Tab - Font Group) ===
  { type: 'function', function: { name: 'format_text', description: 'Apply font formatting to specific text. Can set bold, italic, underline, strikethrough, subscript, superscript, font name, size, color, highlight.', parameters: { type: 'object', properties: { 
    searchText: { type: 'string', description: 'Text to format (required)' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'string', enum: ['none', 'single', 'double', 'dotted', 'dashed', 'wave'] },
    strikethrough: { type: 'boolean' },
    subscript: { type: 'boolean' },
    superscript: { type: 'boolean' },
    fontName: { type: 'string', description: 'Font family name like Arial, Times New Roman, Calibri' },
    fontSize: { type: 'number', description: 'Font size in points' },
    fontColor: { type: 'string', description: 'Color name or hex like red, blue, #FF0000' },
    highlightColor: { type: 'string', enum: ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'gray25', 'gray50', 'black', 'none'] },
    allCaps: { type: 'boolean' },
    smallCaps: { type: 'boolean' }
  }, required: ['searchText'] } } },
  
  { type: 'function', function: { name: 'clear_formatting', description: 'Remove all formatting from specific text, returning it to default.', parameters: { type: 'object', properties: { searchText: { type: 'string' } }, required: ['searchText'] } } },
  
  // === PARAGRAPH FORMATTING (Home Tab - Paragraph Group) ===
  { type: 'function', function: { name: 'format_paragraph', description: 'Format paragraphs containing specific text. Set alignment, line spacing, indentation, spacing before/after.', parameters: { type: 'object', properties: {
    searchText: { type: 'string', description: 'Text within the paragraph to format' },
    alignment: { type: 'string', enum: ['left', 'center', 'right', 'justified'] },
    lineSpacing: { type: 'number', description: 'Line spacing (1, 1.5, 2, etc.)' },
    spaceBefore: { type: 'number', description: 'Space before paragraph in points' },
    spaceAfter: { type: 'number', description: 'Space after paragraph in points' },
    firstLineIndent: { type: 'number', description: 'First line indent in points' },
    leftIndent: { type: 'number', description: 'Left indent in points' },
    rightIndent: { type: 'number', description: 'Right indent in points' }
  }, required: ['searchText'] } } },
  
  // === LISTS (Home Tab - Paragraph Group) ===
  { type: 'function', function: { name: 'create_list', description: 'Convert text/paragraphs into a bulleted or numbered list.', parameters: { type: 'object', properties: {
    searchText: { type: 'string', description: 'Text to convert to list' },
    listType: { type: 'string', enum: ['bullet', 'number'], description: 'Bullet or numbered list' }
  }, required: ['searchText', 'listType'] } } },
  
  { type: 'function', function: { name: 'remove_list', description: 'Remove list formatting from text.', parameters: { type: 'object', properties: { searchText: { type: 'string' } }, required: ['searchText'] } } },
  
  // === STYLES (Home Tab - Styles Group) ===
  { type: 'function', function: { name: 'apply_style', description: 'Apply a built-in Word style to text/paragraph.', parameters: { type: 'object', properties: {
    searchText: { type: 'string' },
    styleName: { type: 'string', enum: ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Title', 'Subtitle', 'Quote', 'IntenseQuote', 'ListParagraph', 'NoSpacing'], description: 'Style to apply' }
  }, required: ['searchText', 'styleName'] } } },
  
  // === COMMENTS ===
  { type: 'function', function: { name: 'add_comment', description: 'Add a comment to specific text in the document margin.', parameters: { type: 'object', properties: { searchText: { type: 'string' }, comment: { type: 'string' } }, required: ['searchText', 'comment'] } } },
  
  // === TABLES ===
  { type: 'function', function: { name: 'insert_table', description: 'Insert a table at the end of the document.', parameters: { type: 'object', properties: {
    rows: { type: 'number', description: 'Number of rows' },
    columns: { type: 'number', description: 'Number of columns' },
    data: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: '2D array of cell contents, e.g. [["A1","B1"],["A2","B2"]]' }
  }, required: ['rows', 'columns'] } } },
  
  { type: 'function', function: { name: 'format_table', description: 'Format an existing table (by index, 0 = first table).', parameters: { type: 'object', properties: {
    tableIndex: { type: 'number', description: 'Table index (0 = first)' },
    style: { type: 'string', enum: ['TableGrid', 'TableGridLight', 'PlainTable1', 'PlainTable2', 'GridTable1Light', 'GridTable4'] },
    headerRow: { type: 'boolean', description: 'Format first row as header' }
  }, required: ['tableIndex'] } } },
  
  { type: 'function', function: { name: 'update_table_cell', description: 'Update content of a specific table cell.', parameters: { type: 'object', properties: {
    tableIndex: { type: 'number' },
    rowIndex: { type: 'number' },
    columnIndex: { type: 'number' },
    content: { type: 'string' }
  }, required: ['tableIndex', 'rowIndex', 'columnIndex', 'content'] } } },
  
  { type: 'function', function: { name: 'add_table_row', description: 'Add a row to a table.', parameters: { type: 'object', properties: {
    tableIndex: { type: 'number' },
    position: { type: 'string', enum: ['start', 'end'] },
    values: { type: 'array', items: { type: 'string' }, description: 'Cell values for new row' }
  }, required: ['tableIndex', 'position'] } } },
  
  // === PAGE LAYOUT ===
  { type: 'function', function: { name: 'insert_break', description: 'Insert a page break or section break.', parameters: { type: 'object', properties: {
    breakType: { type: 'string', enum: ['page', 'line', 'sectionNext', 'sectionContinuous'] }
  }, required: ['breakType'] } } },
  
  // === FIND OPERATIONS ===
  { type: 'function', function: { name: 'find_text', description: 'Find all instances of text and return their count and context.', parameters: { type: 'object', properties: { searchText: { type: 'string' }, matchCase: { type: 'boolean' } }, required: ['searchText'] } } },
  
  // === UNDO/DOCUMENT STATE ===
  { type: 'function', function: { name: 'get_document_info', description: 'Get document metadata and statistics.', parameters: { type: 'object', properties: {} } } }
];

// Color mappings
const HIGHLIGHT_COLORS = { yellow: 'Yellow', green: 'BrightGreen', cyan: 'Turquoise', magenta: 'Pink', blue: 'Blue', red: 'Red', darkBlue: 'DarkBlue', darkCyan: 'DarkCyan', darkGreen: 'DarkGreen', darkMagenta: 'DarkMagenta', darkRed: 'DarkRed', darkYellow: 'DarkYellow', gray25: 'Gray25', gray50: 'Gray50', black: 'Black', none: null };
const UNDERLINE_TYPES = { none: 'None', single: 'Single', double: 'Double', dotted: 'Dotted', dashed: 'Dash', wave: 'Wave' };
const ALIGNMENTS = { left: 'Left', center: 'Centered', right: 'Right', justified: 'Justified' };

// ==================== TOOL EXECUTION ====================
async function execTool(name, args) {
  const Word = getWord();
  if (!Word) return { success: false, message: 'Word API not available' };
  
  try {
    return await Word.run(async ctx => {
      const body = ctx.document.body;
      
      switch (name) {
        // === DOCUMENT READING ===
        case 'get_document_content': {
          body.load('text');
          await ctx.sync();
          return { success: true, content: body.text?.slice(0, 20000) || '', message: `Read ${body.text?.length || 0} chars` };
        }
        
        case 'get_selection': {
          const sel = ctx.document.getSelection();
          sel.load('text');
          await ctx.sync();
          return { success: true, content: sel.text || '', message: sel.text ? `Selected: "${sel.text.slice(0, 100)}..."` : 'No selection' };
        }
        
        case 'get_document_info': {
          body.load('text');
          const tables = body.tables;
          const paragraphs = body.paragraphs;
          tables.load('count');
          paragraphs.load('count');
          await ctx.sync();
          const words = (body.text || '').split(/\s+/).filter(w => w).length;
          return { success: true, info: { characters: body.text?.length || 0, words, paragraphs: paragraphs.count, tables: tables.count } };
        }
        
        // === TEXT EDITING ===
        case 'replace_text': {
          const results = body.search(args.find, { matchCase: args.matchCase || false });
          results.load('items');
          await ctx.sync();
          for (let i = results.items.length - 1; i >= 0; i--) {
            results.items[i].insertText(args.replace, Word.InsertLocation.replace);
          }
          await ctx.sync();
          return { success: true, count: results.items.length, message: `Replaced ${results.items.length} instance(s)` };
        }
        
        case 'insert_text': {
          if (args.position === 'replace_selection') {
            const sel = ctx.document.getSelection();
            sel.insertText(args.text, Word.InsertLocation.replace);
          } else {
            body.insertText(args.text, args.position === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end);
          }
          await ctx.sync();
          return { success: true, message: `Inserted text at ${args.position}` };
        }
        
        case 'delete_text': {
          const results = body.search(args.text, { matchCase: false });
          results.load('items');
          await ctx.sync();
          for (let i = results.items.length - 1; i >= 0; i--) {
            results.items[i].insertText('', Word.InsertLocation.replace);
          }
          await ctx.sync();
          return { success: true, count: results.items.length, message: `Deleted ${results.items.length} instance(s)` };
        }
        
        case 'find_text': {
          const results = body.search(args.searchText, { matchCase: args.matchCase || false });
          results.load('items');
          await ctx.sync();
          return { success: true, count: results.items.length, message: `Found ${results.items.length} instance(s) of "${args.searchText}"` };
        }
        
        // === FONT FORMATTING ===
        case 'format_text': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) return { success: false, message: `"${args.searchText}" not found` };
          
          for (const range of results.items) {
            const font = range.font;
            if (args.bold !== undefined) font.bold = args.bold;
            if (args.italic !== undefined) font.italic = args.italic;
            if (args.underline) font.underline = UNDERLINE_TYPES[args.underline] || args.underline;
            if (args.strikethrough !== undefined) font.strikeThrough = args.strikethrough;
            if (args.subscript !== undefined) font.subscript = args.subscript;
            if (args.superscript !== undefined) font.superscript = args.superscript;
            if (args.fontName) font.name = args.fontName;
            if (args.fontSize) font.size = args.fontSize;
            if (args.fontColor) font.color = args.fontColor;
            if (args.highlightColor !== undefined) font.highlightColor = HIGHLIGHT_COLORS[args.highlightColor] ?? args.highlightColor;
            if (args.allCaps !== undefined) font.allCaps = args.allCaps;
            if (args.smallCaps !== undefined) font.smallCaps = args.smallCaps;
          }
          await ctx.sync();
          return { success: true, count: results.items.length, message: `Formatted ${results.items.length} instance(s)` };
        }
        
        case 'clear_formatting': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          for (const range of results.items) {
            range.font.bold = false;
            range.font.italic = false;
            range.font.underline = 'None';
            range.font.strikeThrough = false;
            range.font.highlightColor = null;
            range.font.color = 'black';
          }
          await ctx.sync();
          return { success: true, count: results.items.length, message: `Cleared formatting on ${results.items.length} instance(s)` };
        }
        
        // === PARAGRAPH FORMATTING ===
        case 'format_paragraph': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) return { success: false, message: `"${args.searchText}" not found` };
          
          for (const range of results.items) {
            const para = range.paragraphs.getFirst();
            if (args.alignment) para.alignment = ALIGNMENTS[args.alignment] || args.alignment;
            if (args.lineSpacing) para.lineSpacing = args.lineSpacing * 12; // Convert to points
            if (args.spaceBefore !== undefined) para.spaceBefore = args.spaceBefore;
            if (args.spaceAfter !== undefined) para.spaceAfter = args.spaceAfter;
            if (args.firstLineIndent !== undefined) para.firstLineIndent = args.firstLineIndent;
            if (args.leftIndent !== undefined) para.leftIndent = args.leftIndent;
            if (args.rightIndent !== undefined) para.rightIndent = args.rightIndent;
          }
          await ctx.sync();
          return { success: true, message: `Formatted paragraph(s) containing "${args.searchText}"` };
        }
        
        // === LISTS ===
        case 'create_list': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) return { success: false, message: `"${args.searchText}" not found` };
          
          for (const range of results.items) {
            const para = range.paragraphs.getFirst();
            if (args.listType === 'bullet') {
              para.listItem.level = 0;
            } else {
              para.startNewList();
            }
          }
          await ctx.sync();
          return { success: true, message: `Applied ${args.listType} list formatting` };
        }
        
        case 'remove_list': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          for (const range of results.items) {
            const para = range.paragraphs.getFirst();
            para.detachFromList();
          }
          await ctx.sync();
          return { success: true, message: 'Removed list formatting' };
        }
        
        // === STYLES ===
        case 'apply_style': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) return { success: false, message: `"${args.searchText}" not found` };
          
          for (const range of results.items) {
            const para = range.paragraphs.getFirst();
            para.styleBuiltIn = args.styleName;
          }
          await ctx.sync();
          return { success: true, message: `Applied ${args.styleName} style` };
        }
        
        // === COMMENTS ===
        case 'add_comment': {
          const results = body.search(args.searchText, { matchCase: false });
          results.load('items');
          await ctx.sync();
          if (results.items.length === 0) return { success: false, message: `"${args.searchText}" not found` };
          results.items[0].insertComment(args.comment);
          await ctx.sync();
          return { success: true, message: `Added comment to "${args.searchText}"` };
        }
        
        // === TABLES ===
        case 'insert_table': {
          const table = body.insertTable(args.rows, args.columns, Word.InsertLocation.end);
          if (args.data) {
            for (let r = 0; r < args.data.length && r < args.rows; r++) {
              for (let c = 0; c < args.data[r].length && c < args.columns; c++) {
                table.getCell(r, c).body.insertText(args.data[r][c], Word.InsertLocation.replace);
              }
            }
          }
          await ctx.sync();
          return { success: true, message: `Inserted ${args.rows}x${args.columns} table` };
        }
        
        case 'format_table': {
          const tables = body.tables;
          tables.load('items');
          await ctx.sync();
          if (args.tableIndex >= tables.items.length) return { success: false, message: 'Table not found' };
          const table = tables.items[args.tableIndex];
          if (args.style) table.styleBuiltIn = args.style;
          if (args.headerRow) table.headerRowCount = 1;
          await ctx.sync();
          return { success: true, message: 'Table formatted' };
        }
        
        case 'update_table_cell': {
          const tables = body.tables;
          tables.load('items');
          await ctx.sync();
          if (args.tableIndex >= tables.items.length) return { success: false, message: 'Table not found' };
          const cell = tables.items[args.tableIndex].getCell(args.rowIndex, args.columnIndex);
          cell.body.insertText(args.content, Word.InsertLocation.replace);
          await ctx.sync();
          return { success: true, message: `Updated cell [${args.rowIndex},${args.columnIndex}]` };
        }
        
        case 'add_table_row': {
          const tables = body.tables;
          tables.load('items');
          await ctx.sync();
          if (args.tableIndex >= tables.items.length) return { success: false, message: 'Table not found' };
          const table = tables.items[args.tableIndex];
          const row = table.addRows(args.position === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end, 1);
          if (args.values) {
            row.load('cells');
            await ctx.sync();
            for (let i = 0; i < args.values.length; i++) {
              row.cells.items[i]?.body.insertText(args.values[i], Word.InsertLocation.replace);
            }
          }
          await ctx.sync();
          return { success: true, message: 'Added row to table' };
        }
        
        // === PAGE LAYOUT ===
        case 'insert_break': {
          const breakTypes = {
            page: Word.BreakType.page,
            line: Word.BreakType.line,
            sectionNext: Word.BreakType.sectionNext,
            sectionContinuous: Word.BreakType.sectionContinuous
          };
          body.insertBreak(breakTypes[args.breakType] || Word.BreakType.page, Word.InsertLocation.end);
          await ctx.sync();
          return { success: true, message: `Inserted ${args.breakType} break` };
        }
        
        default:
          return { success: false, message: `Unknown tool: ${name}` };
      }
    });
  } catch (err) {
    console.error('Tool error:', err);
    return { success: false, message: err.message };
  }
}

// ==================== CONTEXT AWARENESS ====================
async function getDocumentContext() {
  if (!loadBool(STORAGE.contextAwareness, true)) return '';
  const Word = getWord();
  if (!Word) return '';
  try {
    return await Word.run(async ctx => {
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return body.text?.slice(0, 5000) || '';
    });
  } catch { return ''; }
}

// ==================== PROMPT BUILDER ====================
function buildPrompt(actionKey, config) {
  const action = ACTIONS[actionKey];
  const glossaryReplace = document.getElementById('glossary-replace')?.value?.trim() || '';
  const glossaryAvoid = document.getElementById('glossary-avoid')?.value?.trim() || '';
  const customPrompt = document.getElementById('custom-prompt')?.value?.trim() || '';
  
  let prompt = `You are an AI assistant that DIRECTLY EDITS Word documents. You MUST use the provided tools to make changes - the user will see changes appear in their document in real-time.

TASK: ${action.base}

AVAILABLE TOOLS (use these to edit the document):
- get_document_content: Read the document first
- replace_text: Find and replace text
- format_text: Apply font formatting (bold, italic, underline, color, highlight, font size/name, etc.)
- format_paragraph: Set alignment, spacing, indentation
- apply_style: Apply Heading1, Heading2, Title, etc.
- add_comment: Add margin comments
- insert_table: Create tables
- update_table_cell: Edit table cells
- insert_break: Add page/section breaks
- create_list: Make bullet/numbered lists

CRITICAL RULES:
1. ALWAYS call get_document_content FIRST to see what's in the document
2. Make ALL edits using the tools - your text responses are just confirmations
3. For formatting: use format_text with the exact text to format
4. For rewrites: use replace_text with old text and new text
5. Be thorough - process the entire document as requested

`;
  
  if (config?.instructions) prompt += `\nSPECIFIC INSTRUCTIONS: ${config.instructions}\n`;
  if (glossaryReplace) {
    prompt += `\nWORD REPLACEMENTS:\n`;
    glossaryReplace.split('\n').forEach(l => {
      const [o, n] = l.split('→').map(s => s.trim());
      if (o && n) prompt += `- "${o}" → "${n}"\n`;
    });
  }
  if (glossaryAvoid) prompt += `\nWORDS TO AVOID: ${glossaryAvoid.split('\n').map(w => w.trim()).filter(w => w).join(', ')}\n`;
  if (customPrompt) prompt += `\nUSER INSTRUCTIONS: ${customPrompt}\n`;
  
  return prompt;
}

// ==================== API CALL ====================
async function callAI(messages) {
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
  setStatus('Starting...');
  const context = await getDocumentContext();
  const systemPrompt = buildPrompt(actionKey, config);
  
  let messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: context ? `Document content:\n\n${context}\n\nPlease proceed with the task.` : 'Please read the document and proceed.' }
  ];
  
  try {
    let iterations = 0;
    while (iterations < 20) {
      iterations++;
      setStatus(`Processing (${iterations})...`);
      
      const response = await callAI(messages);
      if (!response) throw new Error('No response');
      
      if (response.tool_calls?.length > 0) {
        messages.push({ role: 'assistant', content: response.content || '', tool_calls: response.tool_calls });
        
        for (const tc of response.tool_calls) {
          const name = tc.function?.name;
          let args = {};
          try { args = JSON.parse(tc.function?.arguments || '{}'); } catch {}
          
          setStatus(`${name}...`);
          addMessage('tool', `${name}(${JSON.stringify(args).slice(0, 60)}...)`, 'Tool');
          
          const result = await execTool(name, args);
          messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(result) });
          
          if (result.success) {
            addMessage('system', `✓ ${result.message}`, 'Done');
          } else {
            addMessage('error', `✗ ${result.message}`, 'Error');
          }
          
          if (name === 'replace_text' && result.count > 0) {
            addToHistory(actionKey, args.find, args.replace);
          }
        }
      } else {
        if (response.content?.trim()) addMessage('assistant', response.content.trim(), 'AI');
        break;
      }
    }
    setStatus('Ready');
  } catch (e) {
    addMessage('error', e.message, 'Error');
    setStatus('Error');
  }
}

// ==================== HISTORY ====================
function addToHistory(action, original, replacement) {
  versionHistory.unshift({ id: Date.now(), action, original, replacement, time: new Date().toLocaleTimeString() });
  if (versionHistory.length > 50) versionHistory.pop();
  save(STORAGE.versionHistory, versionHistory);
}

function renderHistory() {
  const list = document.getElementById('history-list');
  if (!versionHistory.length) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);">No history</p>'; return; }
  list.innerHTML = versionHistory.map(h => `
    <div style="padding:8px;border-bottom:1px solid var(--border);font-size:11px;">
      <div style="color:var(--text-muted);">${h.time}</div>
      <div><b>From:</b> ${(h.original||'').slice(0,50)}...</div>
      <div><b>To:</b> ${(h.replacement||'').slice(0,50)}...</div>
      <button onclick="revertHistory(${h.id})" style="margin-top:4px;padding:3px 8px;background:var(--primary);color:white;border:none;border-radius:4px;cursor:pointer;">Revert</button>
    </div>
  `).join('');
}

window.revertHistory = async function(id) {
  const h = versionHistory.find(x => x.id === id);
  if (h) {
    await execTool('replace_text', { find: h.replacement, replace: h.original });
    addMessage('system', 'Reverted', 'History');
  }
};

// ==================== UI ====================
function addMessage(type, content, label = '') {
  const msgs = document.getElementById('messages');
  const div = document.createElement('div');
  div.className = `message ${type}`;
  if (label) { const l = document.createElement('div'); l.className = 'message-label'; l.textContent = label; div.appendChild(l); }
  const t = document.createElement('div'); t.textContent = content; div.appendChild(t);
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
}

function setStatus(text) { document.getElementById('status').textContent = text; }

function renderActionsList() {
  const list = document.getElementById('quick-actions-list');
  list.innerHTML = Object.entries(ACTIONS).map(([key, action]) => `
    <div class="quick-action-item ${actionConfigs[key] ? 'customized' : ''}" data-action="${key}">
      <span class="qa-icon">${action.icon}</span>
      <div class="qa-info"><div class="qa-name">${action.name}</div><div class="qa-desc">${action.desc}</div></div>
      <div class="qa-buttons">
        <button class="qa-btn settings" data-action="${key}" data-mode="settings">⚙️</button>
        <button class="qa-btn run" data-action="${key}" data-mode="run">Run</button>
      </div>
    </div>
  `).join('');
}

// ==================== MODAL ====================
let currentAction = null;
let currentConfig = { ...DEFAULT_CONFIG };

function openModal(key) {
  currentAction = key;
  currentConfig = { ...DEFAULT_CONFIG, ...(actionConfigs[key] || {}) };
  const action = ACTIONS[key];
  document.getElementById('modal-icon').textContent = action.icon;
  document.getElementById('modal-action-name').textContent = action.name;
  document.getElementById('modal-tone').value = currentConfig.tone;
  document.getElementById('modal-formality').value = currentConfig.formality;
  document.getElementById('modal-instructions').value = currentConfig.instructions || '';
  document.getElementById('modal-revision').value = currentConfig.revision;
  document.getElementById('modal-citation').value = currentConfig.citation;
  setSegmented('modal-length', currentConfig.length);
  setSegmented('modal-complexity', String(currentConfig.complexity));
  setSegmented('modal-jargon', currentConfig.jargon);
  setSegmented('modal-voice', currentConfig.voice);
  setIconBtn('modal-structure', currentConfig.structure);
  updateCharCount();
  document.getElementById('action-modal').classList.add('open');
}

function closeModal() { document.getElementById('action-modal').classList.remove('open'); }
function setSegmented(id, val) { document.querySelectorAll(`#${id} .segmented-btn`).forEach(b => b.classList.toggle('active', b.dataset.value === val)); }
function getSegmented(id) { return document.querySelector(`#${id} .segmented-btn.active`)?.dataset.value || ''; }
function setIconBtn(id, val) { document.querySelectorAll(`#${id} .icon-btn`).forEach(b => b.classList.toggle('active', b.dataset.value === val)); }
function getIconBtn(id) { return document.querySelector(`#${id} .icon-btn.active`)?.dataset.value || 'paragraphs'; }
function updateCharCount() { const i = document.getElementById('modal-instructions'); const c = document.getElementById('modal-char-count'); c.textContent = `${i.value.length}/280`; }

function getConfigFromModal() {
  return {
    tone: document.getElementById('modal-tone').value,
    length: getSegmented('modal-length'),
    formality: +document.getElementById('modal-formality').value,
    structure: getIconBtn('modal-structure'),
    complexity: +getSegmented('modal-complexity'),
    instructions: document.getElementById('modal-instructions').value.slice(0, 280),
    jargon: getSegmented('modal-jargon'),
    voice: getSegmented('modal-voice'),
    revision: +document.getElementById('modal-revision').value,
    citation: document.getElementById('modal-citation').value
  };
}

// ==================== INIT ====================
Office.onReady(info => {
  if (info.host !== Office.HostType.Word) { document.body.innerHTML = '<p>Requires Word</p>'; return; }
  
  loadAllData();
  renderActionsList();
  
  // Tabs
  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(x => x.classList.remove('active'));
    t.classList.add('active');
    document.querySelector(`.tab-content[data-tab="${t.dataset.tab}"]`).classList.add('active');
  }));
  
  // Quick actions
  document.getElementById('quick-actions-list').addEventListener('click', e => {
    const btn = e.target.closest('.qa-btn');
    if (!btn) return;
    if (btn.dataset.mode === 'settings') openModal(btn.dataset.action);
    else { runAction(btn.dataset.action, { ...DEFAULT_CONFIG, ...(actionConfigs[btn.dataset.action] || {}) }); document.querySelector('.tab[data-tab="chat"]').click(); }
  });
  
  // Modal
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-cancel').addEventListener('click', closeModal);
  document.getElementById('action-modal').addEventListener('click', e => { if (e.target.id === 'action-modal') closeModal(); });
  document.getElementById('modal-apply').addEventListener('click', () => {
    const cfg = getConfigFromModal();
    actionConfigs[currentAction] = cfg;
    saveData(STORAGE.actionConfigs, actionConfigs);
    closeModal();
    runAction(currentAction, cfg);
    document.querySelector('.tab[data-tab="chat"]').click();
    renderActionsList();
  });
  document.getElementById('modal-reset').addEventListener('click', () => { delete actionConfigs[currentAction]; saveData(STORAGE.actionConfigs, actionConfigs); closeModal(); renderActionsList(); });
  document.getElementById('modal-preset').addEventListener('change', e => { if (e.target.value && PRESETS[e.target.value]) { Object.assign(currentConfig, PRESETS[e.target.value]); openModal(currentAction); } e.target.value = ''; });
  
  // Segmented/Icon controls
  document.querySelectorAll('.segmented').forEach(s => s.addEventListener('click', e => { const b = e.target.closest('.segmented-btn'); if (b) { s.querySelectorAll('.segmented-btn').forEach(x => x.classList.remove('active')); b.classList.add('active'); } }));
  document.querySelectorAll('.icon-row').forEach(r => r.addEventListener('click', e => { const b = e.target.closest('.icon-btn'); if (b) { r.querySelectorAll('.icon-btn').forEach(x => x.classList.remove('active')); b.classList.add('active'); } }));
  document.querySelectorAll('.accordion-header').forEach(h => h.addEventListener('click', () => h.parentElement.classList.toggle('open')));
  document.getElementById('modal-instructions').addEventListener('input', updateCharCount);
  
  // Settings
  document.getElementById('provider').addEventListener('change', () => {
    document.getElementById('local-settings').style.display = document.getElementById('provider').value === 'local' ? 'block' : 'none';
    saveData(STORAGE.provider, document.getElementById('provider').value);
  });
  document.getElementById('local-settings').style.display = document.getElementById('provider').value === 'local' ? 'block' : 'none';
  
  ['api-key', 'local-url', 'local-model', 'custom-prompt', 'glossary-replace', 'glossary-avoid'].forEach(id => {
    document.getElementById(id)?.addEventListener('change', () => saveData(STORAGE[id.replace(/-/g, '')] || STORAGE.apiKey, document.getElementById(id).value));
  });
  
  document.getElementById('privacy-save-key').addEventListener('change', e => { save(STORAGE.privacySaveKey, e.target.checked); if (!e.target.checked) localStorage.removeItem(STORAGE.apiKey); });
  document.getElementById('privacy-save-presets').addEventListener('change', e => { save(STORAGE.privacySavePresets, e.target.checked); });
  document.getElementById('context-awareness').addEventListener('change', e => save(STORAGE.contextAwareness, e.target.checked));
  document.getElementById('clear-data').addEventListener('click', () => { if (confirm('Clear all?')) { Object.values(STORAGE).forEach(k => localStorage.removeItem(k)); location.reload(); } });
  
  // History
  document.getElementById('history-btn').addEventListener('click', () => { renderHistory(); document.getElementById('history-modal').classList.add('open'); });
  document.getElementById('history-close').addEventListener('click', () => document.getElementById('history-modal').classList.remove('open'));
  document.getElementById('history-done').addEventListener('click', () => document.getElementById('history-modal').classList.remove('open'));
  document.getElementById('history-clear').addEventListener('click', () => { versionHistory = []; save(STORAGE.versionHistory, []); renderHistory(); });
  
  // Pop-out
  document.getElementById('popout-btn').addEventListener('click', () => window.open(location.href, 'WordAI', 'width=450,height=700'));
  
  // Export/Import
  document.getElementById('export-settings')?.addEventListener('click', () => {
    const data = { version: '1.0', date: new Date().toISOString(), settings: {
      apiKey: document.getElementById('api-key').value,
      provider: document.getElementById('provider').value,
      localUrl: document.getElementById('local-url').value,
      localModel: document.getElementById('local-model').value,
      customPrompt: document.getElementById('custom-prompt').value,
      glossaryReplace: document.getElementById('glossary-replace').value,
      glossaryAvoid: document.getElementById('glossary-avoid').value
    }, actionConfigs, userPresets };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `word-ai-settings.json`; a.click();
  });
  
  document.getElementById('import-settings')?.addEventListener('change', e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.settings) {
          Object.entries(data.settings).forEach(([k, v]) => { const el = document.getElementById(k.replace(/([A-Z])/g, '-$1').toLowerCase()); if (el) el.value = v; });
          Object.entries(data.settings).forEach(([k, v]) => save(STORAGE[k] || k, v));
        }
        if (data.actionConfigs) { actionConfigs = data.actionConfigs; save(STORAGE.actionConfigs, actionConfigs); renderActionsList(); }
        if (data.userPresets) { userPresets = data.userPresets; save(STORAGE.userPresets, userPresets); }
        document.getElementById('import-status').innerHTML = '<span style="color:var(--success)">✓ Imported</span>';
      } catch (err) { document.getElementById('import-status').innerHTML = `<span style="color:var(--danger)">✗ ${err.message}</span>`; }
    };
    reader.readAsText(file);
  });
  
  // Chat
  const sendBtn = document.getElementById('send-btn');
  const userInput = document.getElementById('user-input');
  const updateSendBtn = () => { sendBtn.disabled = !(document.getElementById('api-key').value || document.getElementById('provider').value === 'local') || !userInput.value.trim(); };
  userInput.addEventListener('input', updateSendBtn);
  document.getElementById('api-key').addEventListener('input', updateSendBtn);
  
  sendBtn.addEventListener('click', async () => {
    const text = userInput.value.trim();
    if (!text) return;
    userInput.value = '';
    addMessage('user', text, 'You');
    
    const systemPrompt = `You are an AI that DIRECTLY EDITS Word documents using tools. Changes you make appear IMMEDIATELY in the user's document.

AVAILABLE TOOLS:
- get_document_content: Read the document
- replace_text: Find & replace text  
- format_text: Apply formatting (bold, italic, underline, colors, fonts, highlight, etc.)
- format_paragraph: Alignment, spacing, indentation
- apply_style: Heading1, Heading2, Title, Normal, etc.
- add_comment: Add margin comments
- insert_table: Create tables with data
- update_table_cell: Edit table cells
- add_table_row: Add rows to tables
- create_list: Bullet or numbered lists
- insert_break: Page/section breaks
- insert_text: Add text at start/end
- delete_text: Remove text

RULES:
1. ALWAYS use get_document_content FIRST
2. ALL changes must use tools - don't just describe what to do
3. Make changes directly to the document
4. Be thorough and complete

${document.getElementById('custom-prompt').value || ''}`;

    let messages = [{ role: 'system', content: systemPrompt }, { role: 'user', content: text }];
    
    try {
      let iterations = 0;
      while (iterations < 20) {
        iterations++;
        setStatus(`Working (${iterations})...`);
        const response = await callAI(messages);
        if (!response) throw new Error('No response');
        
        if (response.tool_calls?.length > 0) {
          messages.push({ role: 'assistant', content: response.content || '', tool_calls: response.tool_calls });
          for (const tc of response.tool_calls) {
            const name = tc.function?.name;
            let args = {}; try { args = JSON.parse(tc.function?.arguments || '{}'); } catch {}
            setStatus(`${name}...`);
            addMessage('tool', `${name}(${JSON.stringify(args).slice(0, 60)}...)`, 'Tool');
            const result = await execTool(name, args);
            messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(result) });
            addMessage(result.success ? 'system' : 'error', `${result.success ? '✓' : '✗'} ${result.message}`, result.success ? 'Done' : 'Error');
          }
        } else {
          if (response.content?.trim()) addMessage('assistant', response.content.trim(), 'AI');
          break;
        }
      }
      setStatus('Ready');
    } catch (e) { addMessage('error', e.message, 'Error'); setStatus('Error'); }
  });
  
  userInput.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendBtn.click(); } });
  
  addMessage('system', 'Tell me what to do and I\'ll edit your document directly. Try: "Make all headings bold" or "Add a table with 3 columns"', 'Ready');
});
