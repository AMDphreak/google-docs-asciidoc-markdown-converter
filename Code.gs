/**
 * Google Apps Script to convert Markdown/AsciiDoc to Google Docs formatting
 * Adds menu items to right-click context menu and Extensions menu
 */

/**
 * Runs when the document is opened
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Markdown/AsciiDoc Converter')
    .addItem('Convert Selected Text', 'convertSelectedText')
    .addSeparator()
    .addItem('Convert Markdown', 'convertMarkdown')
    .addItem('Convert AsciiDoc', 'convertAsciiDoc')
    .addToUi();
}

/**
 * Creates a custom context menu item
 */
function onInstall() {
  onOpen();
}

/**
 * Converts the selected text from Markdown/AsciiDoc to Google Docs formatting
 */
function convertSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection || selection.getRangeElements().length === 0) {
    DocumentApp.getUi().alert('Please select some text to convert.');
    return;
  }
  
  const rangeElements = selection.getRangeElements();
  let text = '';
  let startOffset = 0;
  let endOffset = 0;
  let element = null;
  
  // Get the selected text
  for (let i = 0; i < rangeElements.length; i++) {
    const rangeElement = rangeElements[i];
    if (rangeElement.getElement().getType() === DocumentApp.ElementType.TEXT) {
      element = rangeElement.getElement().asText();
      const start = rangeElement.getStartOffset();
      const end = rangeElement.getEndOffsetInclusive();
      text += element.getText().substring(start, end + 1);
      
      if (i === 0) {
        startOffset = start;
      }
      if (i === rangeElements.length - 1) {
        endOffset = end;
      }
    }
  }
  
  if (!text) {
    DocumentApp.getUi().alert('No text found in selection.');
    return;
  }
  
  // Detect format (try AsciiDoc first, then Markdown)
  const format = detectFormat(text);
  
  // Convert based on detected format
  if (format === 'asciidoc') {
    convertAsciiDocToGoogleDocs(rangeElements, text);
  } else {
    convertMarkdownToGoogleDocs(rangeElements, text);
  }
}

/**
 * Detects if text is AsciiDoc or Markdown
 */
function detectFormat(text) {
  // AsciiDoc indicators
  const asciidocPatterns = [
    /^=+\s/,           // Headers with =
    /^\[.*\]$/,        // Attribute lists
    /^\.{3,}/,         // Block titles
    /^\[source/        // Source blocks
  ];
  
  const lines = text.split('\n');
  let asciidocScore = 0;
  let markdownScore = 0;
  
  for (const line of lines) {
    for (const pattern of asciidocPatterns) {
      if (pattern.test(line.trim())) {
        asciidocScore++;
        break;
      }
    }
    
    // Markdown indicators
    if (/^#{1,6}\s/.test(line.trim())) {
      markdownScore++;
    }
  }
  
  return asciidocScore > markdownScore ? 'asciidoc' : 'markdown';
}

/**
 * Converts Markdown text to Google Docs formatting
 */
function convertMarkdownToGoogleDocs(rangeElements, text) {
  const doc = DocumentApp.getActiveDocument();
  
  const firstElement = rangeElements[0].getElement();
  const lastElement = rangeElements[rangeElements.length - 1].getElement();
  const startOffset = rangeElements[0].getStartOffset();
  const endOffset = rangeElements[rangeElements.length - 1].getEndOffsetInclusive();
  
  // Process the text and create formatted version
  const lines = text.split('\n');
  const processedLines = [];
  const formattingInstructions = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const processed = processMarkdownLine(line, i);
    processedLines.push(processed.text);
    formattingInstructions.push(processed.formatting);
  }
  
  // Combine processed text
  const newText = processedLines.join('\n');
  
  // Replace the selected text
  if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
    const textElement = firstElement.asText();
    textElement.deleteText(startOffset, endOffset);
    textElement.insertText(startOffset, newText);
    
    // Apply formatting
    let currentPos = startOffset;
    for (let i = 0; i < processedLines.length; i++) {
      const line = processedLines[i];
      const formatting = formattingInstructions[i];
      
      applyFormattingToRange(textElement, currentPos, currentPos + line.length - 1, formatting);
      
      currentPos += line.length + 1; // +1 for newline
    }
  }
}

/**
 * Processes a single Markdown line and returns text with formatting info
 */
function processMarkdownLine(line, lineIndex) {
  let text = line;
  const formatting = {
    headerLevel: null,
    boldRanges: [],
    italicRanges: [],
    codeRanges: [],
    linkRanges: []
  };
  
  // Headers (# Header)
  const headerMatch = text.match(/^(#{1,6})\s+(.+)$/);
  if (headerMatch) {
    formatting.headerLevel = headerMatch[1].length;
    text = headerMatch[2];
    return { text: text, formatting: formatting };
  }
  
  // Links [text](url) - process first to avoid conflicts
  const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;
  let match;
  const linkReplacements = [];
  while ((match = linkRegex.exec(line)) !== null) {
    linkReplacements.push({
      original: match[0],
      replacement: match[1],
      url: match[2],
      index: match.index
    });
  }
  
  // Replace links in reverse order to maintain indices
  for (let i = linkReplacements.length - 1; i >= 0; i--) {
    const link = linkReplacements[i];
    text = text.substring(0, link.index) + link.replacement + text.substring(link.index + link.original.length);
    formatting.linkRanges.push({
      start: link.index,
      end: link.index + link.replacement.length - 1,
      url: link.url
    });
  }
  
  // Bold **text** or __text__ (process before italic to avoid conflicts)
  const boldRegex = /(\*\*|__)(.+?)\1/g;
  const boldReplacements = [];
  while ((match = boldRegex.exec(text)) !== null) {
    boldReplacements.push({
      original: match[0],
      replacement: match[2],
      index: match.index
    });
  }
  
  for (let i = boldReplacements.length - 1; i >= 0; i--) {
    const bold = boldReplacements[i];
    text = text.substring(0, bold.index) + bold.replacement + text.substring(bold.index + bold.original.length);
    formatting.boldRanges.push({
      start: bold.index,
      end: bold.index + bold.replacement.length - 1
    });
  }
  
  // Code `text` (process before italic)
  const codeRegex = /`([^`]+)`/g;
  const codeReplacements = [];
  while ((match = codeRegex.exec(text)) !== null) {
    codeReplacements.push({
      original: match[0],
      replacement: match[1],
      index: match.index
    });
  }
  
  for (let i = codeReplacements.length - 1; i >= 0; i--) {
    const code = codeReplacements[i];
    text = text.substring(0, code.index) + code.replacement + text.substring(code.index + code.original.length);
    formatting.codeRanges.push({
      start: code.index,
      end: code.index + code.replacement.length - 1
    });
  }
  
  // Italic *text* or _text_ (single asterisk/underscore, not already bold)
  // Use a simpler approach: match single * or _ that aren't part of ** or __
  const italicRegex = /(?:^|[^*_])\*([^*\n]+?)\*(?![*_])|(?:^|[^*_])_([^_\n]+?)_(?![*_])/g;
  const italicReplacements = [];
  while ((match = italicRegex.exec(text)) !== null) {
    const italicText = match[1] || match[2];
    // Adjust index if we matched a preceding character
    const actualIndex = match.index + (match[0][0] === '*' || match[0][0] === '_' ? 0 : 1);
    italicReplacements.push({
      original: match[0],
      replacement: italicText,
      index: actualIndex
    });
  }
  
  for (let i = italicReplacements.length - 1; i >= 0; i--) {
    const italic = italicReplacements[i];
    text = text.substring(0, italic.index) + italic.replacement + text.substring(italic.index + italic.original.length);
    formatting.italicRanges.push({
      start: italic.index,
      end: italic.index + italic.replacement.length - 1
    });
  }
  
  // Remove list markers
  if (/^[-*+]\s/.test(text.trim())) {
    text = text.replace(/^[-*+]\s+/, '');
  }
  
  return { text: text, formatting: formatting };
}

/**
 * Applies formatting to a text range
 */
function applyFormattingToRange(textElement, startPos, endPos, formatting) {
  // Apply header formatting
  if (formatting.headerLevel) {
    textElement.setFontSize(startPos, endPos, getHeaderSize(formatting.headerLevel));
    textElement.setBold(startPos, endPos, true);
    return;
  }
  
  // Apply bold
  for (const range of formatting.boldRanges) {
    const actualStart = startPos + range.start;
    const actualEnd = startPos + range.end;
    if (actualStart <= actualEnd && actualEnd < textElement.getText().length) {
      textElement.setBold(actualStart, actualEnd, true);
    }
  }
  
  // Apply italic
  for (const range of formatting.italicRanges) {
    const actualStart = startPos + range.start;
    const actualEnd = startPos + range.end;
    if (actualStart <= actualEnd && actualEnd < textElement.getText().length) {
      textElement.setItalic(actualStart, actualEnd, true);
    }
  }
  
  // Apply code formatting
  for (const range of formatting.codeRanges) {
    const actualStart = startPos + range.start;
    const actualEnd = startPos + range.end;
    if (actualStart <= actualEnd && actualEnd < textElement.getText().length) {
      textElement.setFontFamily(actualStart, actualEnd, 'Courier New');
      textElement.setBackgroundColor(actualStart, actualEnd, '#f4f4f4');
    }
  }
  
  // Apply links
  for (const range of formatting.linkRanges) {
    const actualStart = startPos + range.start;
    const actualEnd = startPos + range.end;
    if (actualStart <= actualEnd && actualEnd < textElement.getText().length) {
      try {
        textElement.setLinkUrl(actualStart, actualEnd, range.url);
      } catch (e) {
        // URL might be invalid, skip
      }
    }
  }
}

/**
 * Converts AsciiDoc text to Google Docs formatting
 */
function convertAsciiDocToGoogleDocs(rangeElements, text) {
  const doc = DocumentApp.getActiveDocument();
  
  const firstElement = rangeElements[0].getElement();
  const lastElement = rangeElements[rangeElements.length - 1].getElement();
  const startOffset = rangeElements[0].getStartOffset();
  const endOffset = rangeElements[rangeElements.length - 1].getEndOffsetInclusive();
  
  const lines = text.split('\n');
  const processedLines = [];
  const formattingInstructions = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const processed = processAsciiDocLine(line);
    processedLines.push(processed.text);
    formattingInstructions.push(processed.formatting);
  }
  
  // Combine processed text
  const newText = processedLines.join('\n');
  
  // Replace the selected text
  if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
    const textElement = firstElement.asText();
    textElement.deleteText(startOffset, endOffset);
    textElement.insertText(startOffset, newText);
    
    // Apply formatting
    let currentPos = startOffset;
    for (let i = 0; i < processedLines.length; i++) {
      const line = processedLines[i];
      const formatting = formattingInstructions[i];
      
      applyFormattingToRange(textElement, currentPos, currentPos + line.length - 1, formatting);
      
      currentPos += line.length + 1; // +1 for newline
    }
  }
}

/**
 * Processes a single AsciiDoc line and returns text with formatting info
 */
function processAsciiDocLine(line) {
  let text = line;
  const formatting = {
    headerLevel: null,
    boldRanges: [],
    italicRanges: [],
    codeRanges: [],
    linkRanges: []
  };
  
  // Headers (= Header, == Header, etc.)
  const headerMatch = text.match(/^(=+)\s+(.+)$/);
  if (headerMatch) {
    formatting.headerLevel = headerMatch[1].length;
    text = headerMatch[2];
    return { text: text, formatting: formatting };
  }
  
  // Bold **text** or *text* (when not part of italic)
  const boldRegex = /\*\*([^*]+)\*\*/g;
  const boldReplacements = [];
  let match;
  while ((match = boldRegex.exec(text)) !== null) {
    boldReplacements.push({
      original: match[0],
      replacement: match[1],
      index: match.index
    });
  }
  
  for (let i = boldReplacements.length - 1; i >= 0; i--) {
    const bold = boldReplacements[i];
    text = text.substring(0, bold.index) + bold.replacement + text.substring(bold.index + bold.original.length);
    formatting.boldRanges.push({
      start: bold.index,
      end: bold.index + bold.replacement.length - 1
    });
  }
  
  // Code ``text`` or `text`
  const codeRegex = /``([^`]+)``|`([^`]+)`/g;
  const codeReplacements = [];
  while ((match = codeRegex.exec(text)) !== null) {
    const codeText = match[1] || match[2];
    codeReplacements.push({
      original: match[0],
      replacement: codeText,
      index: match.index
    });
  }
  
  for (let i = codeReplacements.length - 1; i >= 0; i--) {
    const code = codeReplacements[i];
    text = text.substring(0, code.index) + code.replacement + text.substring(code.index + code.original.length);
    formatting.codeRanges.push({
      start: code.index,
      end: code.index + code.replacement.length - 1
    });
  }
  
  // Italic *text* (single asterisk, not already bold)
  // Match single * that aren't part of **
  const italicRegex = /(?:^|[^*])\*([^*\n]+?)\*(?![*])/g;
  const italicReplacements = [];
  while ((match = italicRegex.exec(text)) !== null) {
    // Adjust index if we matched a preceding character
    const actualIndex = match.index + (match[0][0] === '*' ? 0 : 1);
    italicReplacements.push({
      original: match[0],
      replacement: match[1],
      index: actualIndex
    });
  }
  
  for (let i = italicReplacements.length - 1; i >= 0; i--) {
    const italic = italicReplacements[i];
    text = text.substring(0, italic.index) + italic.replacement + text.substring(italic.index + italic.original.length);
    formatting.italicRanges.push({
      start: italic.index,
      end: italic.index + italic.replacement.length - 1
    });
  }
  
  // Links link:url[text] or url[text]
  const linkRegex = /(?:link:)?([^\s\[\]]+)\[([^\]]+)\]/g;
  const linkReplacements = [];
  while ((match = linkRegex.exec(text)) !== null) {
    linkReplacements.push({
      original: match[0],
      replacement: match[2],
      url: match[1],
      index: match.index
    });
  }
  
  for (let i = linkReplacements.length - 1; i >= 0; i--) {
    const link = linkReplacements[i];
    text = text.substring(0, link.index) + link.replacement + text.substring(link.index + link.original.length);
    formatting.linkRanges.push({
      start: link.index,
      end: link.index + link.replacement.length - 1,
      url: link.url
    });
  }
  
  return { text: text, formatting: formatting };
}

/**
 * Gets font size for header level
 */
function getHeaderSize(level) {
  const sizes = {
    1: 24,
    2: 20,
    3: 16,
    4: 14,
    5: 12,
    6: 11
  };
  return sizes[level] || 11;
}

/**
 * Menu item: Convert Markdown
 */
function convertMarkdown() {
  convertSelectedText();
}

/**
 * Menu item: Convert AsciiDoc
 */
function convertAsciiDoc() {
  convertSelectedText();
}

