/**
 * Google Apps Script to convert Markdown/AsciiDoc to Google Docs formatting
 * Adds menu items to right-click context menu and Extensions menu
 */

/**
 * Runs when the document is opened
 */
function onOpen() {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu('Markdown/AsciiDoc Converter')
      .addItem('Convert Selected Text', 'convertSelectedText')
      .addSeparator()
      .addItem('Convert Markdown', 'convertMarkdown')
      .addItem('Convert AsciiDoc', 'convertAsciiDoc')
      .addToUi();
  } catch (error) {
    Logger.log('Error in onOpen: ' + error.toString());
  }
}

/**
 * Creates a custom context menu item
 */
function onInstall() {
  try {
    onOpen();
  } catch (error) {
    Logger.log('Error in onInstall: ' + error.toString());
    showError('Installation Error', 'Failed to install the converter. Please refresh the page and try again.');
  }
}

/**
 * Shows a loading indicator message (non-blocking)
 * Note: In Google Apps Script, we log the status since modal dialogs block execution
 */
function showLoadingIndicator(message) {
  try {
    Logger.log('Loading: ' + (message || 'Processing...'));
    // Note: Modal dialogs block execution in Apps Script, so we just log
    // The success/error messages will provide user feedback
  } catch (error) {
    Logger.log('Error showing loading indicator: ' + error.toString());
  }
}

/**
 * Shows a success message
 */
function showSuccess(message, format) {
  try {
    const ui = DocumentApp.getUi();
    const formatText = format ? ` (${format})` : '';
    ui.alert('Success', `Conversion completed successfully${formatText}!\n\n${message}`, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('Error showing success message: ' + error.toString());
  }
}

/**
 * Shows an error message with details
 */
function showError(title, message, details) {
  try {
    const ui = DocumentApp.getUi();
    const fullMessage = details ? `${message}\n\nDetails: ${details}` : message;
    ui.alert(title || 'Error', fullMessage, ui.ButtonSet.OK);
    Logger.log(`Error: ${title} - ${message}${details ? ' - ' + details : ''}`);
  } catch (error) {
    Logger.log('Error showing error message: ' + error.toString());
  }
}

/**
 * Converts the selected text from Markdown/AsciiDoc to Google Docs formatting
 */
function convertSelectedText() {
  try {
    const doc = DocumentApp.getActiveDocument();
    
    if (!doc) {
      showError('Document Error', 'Unable to access the document. Please ensure you have a document open.');
      return;
    }
    
    const selection = doc.getSelection();
    
    if (!selection || selection.getRangeElements().length === 0) {
      showError('Selection Required', 'Please select some text to convert.');
      return;
    }
    
    const rangeElements = selection.getRangeElements();
    
    if (!rangeElements || rangeElements.length === 0) {
      showError('Selection Error', 'Unable to read the selected text. Please try selecting the text again.');
      return;
    }
    
    let text = '';
    let startOffset = 0;
    let endOffset = 0;
    let element = null;
    
    // Get the selected text
    try {
      for (let i = 0; i < rangeElements.length; i++) {
        const rangeElement = rangeElements[i];
        if (!rangeElement || !rangeElement.getElement()) {
          continue;
        }
        
        if (rangeElement.getElement().getType() === DocumentApp.ElementType.TEXT) {
          element = rangeElement.getElement().asText();
          const start = rangeElement.getStartOffset();
          const end = rangeElement.getEndOffsetInclusive();
          
          if (start < 0 || end < 0 || start > end) {
            Logger.log(`Invalid range: start=${start}, end=${end}`);
            continue;
          }
          
          const elementText = element.getText();
          if (elementText && elementText.length > 0) {
            const substring = elementText.substring(start, end + 1);
            text += substring;
            
            if (i === 0) {
              startOffset = start;
            }
            if (i === rangeElements.length - 1) {
              endOffset = end;
            }
          }
        }
      }
    } catch (error) {
      showError('Text Extraction Error', 'Failed to extract text from selection.', error.toString());
      return;
    }
    
    if (!text || text.trim().length === 0) {
      showError('Empty Selection', 'No text found in selection. Please select text that contains content.');
      return;
    }
    
    // Detect format (try AsciiDoc first, then Markdown)
    let format;
    try {
      format = detectFormat(text);
    } catch (error) {
      showError('Format Detection Error', 'Failed to detect the format of the selected text.', error.toString());
      return;
    }
    
    // Show loading indicator
    try {
      showLoadingIndicator('Converting ' + format + '...');
    } catch (error) {
      Logger.log('Could not show loading indicator: ' + error.toString());
    }
    
    // Convert based on detected format
    let success = false;
    try {
      if (format === 'asciidoc') {
        convertAsciiDocToGoogleDocs(rangeElements, text);
        success = true;
      } else {
        convertMarkdownToGoogleDocs(rangeElements, text);
        success = true;
      }
    } catch (error) {
      showError('Conversion Error', `Failed to convert ${format} text.`, error.toString());
      Logger.log('Conversion error: ' + error.toString());
      return;
    }
    
    // Show success message
    if (success) {
      const charCount = text.length;
      const lineCount = text.split('\n').length;
      showSuccess(`Converted ${charCount} characters across ${lineCount} line(s)`, format);
    }
    
  } catch (error) {
    showError('Unexpected Error', 'An unexpected error occurred during conversion.', error.toString());
    Logger.log('Unexpected error in convertSelectedText: ' + error.toString());
  }
}

/**
 * Detects if text is AsciiDoc or Markdown
 */
function detectFormat(text) {
  try {
    if (!text || typeof text !== 'string') {
      throw new Error('Invalid text input for format detection');
    }
    
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
      if (!line || typeof line !== 'string') {
        continue;
      }
      
      const trimmedLine = line.trim();
      
      for (const pattern of asciidocPatterns) {
        try {
          if (pattern.test(trimmedLine)) {
            asciidocScore++;
            break;
          }
        } catch (error) {
          Logger.log('Error testing AsciiDoc pattern: ' + error.toString());
        }
      }
      
      // Markdown indicators
      try {
        if (/^#{1,6}\s/.test(trimmedLine)) {
          markdownScore++;
        }
      } catch (error) {
        Logger.log('Error testing Markdown pattern: ' + error.toString());
      }
    }
    
    return asciidocScore > markdownScore ? 'asciidoc' : 'markdown';
  } catch (error) {
    Logger.log('Error in detectFormat: ' + error.toString());
    // Default to markdown if detection fails
    return 'markdown';
  }
}

/**
 * Converts Markdown text to Google Docs formatting
 */
function convertMarkdownToGoogleDocs(rangeElements, text) {
  try {
    if (!rangeElements || rangeElements.length === 0) {
      throw new Error('No range elements provided');
    }
    
    if (!text || typeof text !== 'string') {
      throw new Error('Invalid text input');
    }
    
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      throw new Error('Unable to access document');
    }
    
    const firstElement = rangeElements[0].getElement();
    if (!firstElement) {
      throw new Error('Unable to access first element');
    }
    
    const startOffset = rangeElements[0].getStartOffset();
    const endOffset = rangeElements[rangeElements.length - 1].getEndOffsetInclusive();
    
    if (startOffset < 0 || endOffset < 0 || startOffset > endOffset) {
      throw new Error(`Invalid text range: start=${startOffset}, end=${endOffset}`);
    }
    
    // Process the text and create formatted version
    const lines = text.split('\n');
    const processedLines = [];
    const formattingInstructions = [];
    
    for (let i = 0; i < lines.length; i++) {
      try {
        const line = lines[i];
        const processed = processMarkdownLine(line, i);
        processedLines.push(processed.text);
        formattingInstructions.push(processed.formatting);
      } catch (error) {
        Logger.log(`Error processing line ${i}: ${error.toString()}`);
        // Continue with unprocessed line
        processedLines.push(lines[i]);
        formattingInstructions.push({
          headerLevel: null,
          boldRanges: [],
          italicRanges: [],
          codeRanges: [],
          linkRanges: []
        });
      }
    }
    
    // Combine processed text
    const newText = processedLines.join('\n');
    
    // Replace the selected text
    if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
      const textElement = firstElement.asText();
      
      try {
        textElement.deleteText(startOffset, endOffset);
        textElement.insertText(startOffset, newText);
      } catch (error) {
        throw new Error(`Failed to replace text: ${error.toString()}`);
      }
      
      // Apply formatting
      try {
        let currentPos = startOffset;
        for (let i = 0; i < processedLines.length; i++) {
          const line = processedLines[i];
          const formatting = formattingInstructions[i];
          
          if (line.length > 0) {
            applyFormattingToRange(textElement, currentPos, currentPos + line.length - 1, formatting);
          }
          
          currentPos += line.length + 1; // +1 for newline
        }
      } catch (error) {
        Logger.log(`Error applying formatting: ${error.toString()}`);
        // Text was replaced, but formatting failed - this is not critical
      }
    } else {
      throw new Error('Selected element is not a text element');
    }
  } catch (error) {
    Logger.log('Error in convertMarkdownToGoogleDocs: ' + error.toString());
    throw error;
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
  try {
    if (!textElement || !formatting) {
      return;
    }
    
    const textLength = textElement.getText().length;
    
    // Validate positions
    if (startPos < 0 || endPos < 0 || startPos >= textLength || endPos >= textLength) {
      Logger.log(`Invalid formatting range: startPos=${startPos}, endPos=${endPos}, textLength=${textLength}`);
      return;
    }
    
    // Apply header formatting
    if (formatting.headerLevel) {
      try {
        const headerSize = getHeaderSize(formatting.headerLevel);
        textElement.setFontSize(startPos, endPos, headerSize);
        textElement.setBold(startPos, endPos, true);
      } catch (error) {
        Logger.log(`Error applying header formatting: ${error.toString()}`);
      }
      return;
    }
    
    // Apply bold
    if (formatting.boldRanges && Array.isArray(formatting.boldRanges)) {
      for (const range of formatting.boldRanges) {
        try {
          const actualStart = startPos + range.start;
          const actualEnd = startPos + range.end;
          if (actualStart >= 0 && actualEnd >= 0 && actualStart <= actualEnd && actualEnd < textLength) {
            textElement.setBold(actualStart, actualEnd, true);
          }
        } catch (error) {
          Logger.log(`Error applying bold formatting: ${error.toString()}`);
        }
      }
    }
    
    // Apply italic
    if (formatting.italicRanges && Array.isArray(formatting.italicRanges)) {
      for (const range of formatting.italicRanges) {
        try {
          const actualStart = startPos + range.start;
          const actualEnd = startPos + range.end;
          if (actualStart >= 0 && actualEnd >= 0 && actualStart <= actualEnd && actualEnd < textLength) {
            textElement.setItalic(actualStart, actualEnd, true);
          }
        } catch (error) {
          Logger.log(`Error applying italic formatting: ${error.toString()}`);
        }
      }
    }
    
    // Apply code formatting
    if (formatting.codeRanges && Array.isArray(formatting.codeRanges)) {
      for (const range of formatting.codeRanges) {
        try {
          const actualStart = startPos + range.start;
          const actualEnd = startPos + range.end;
          if (actualStart >= 0 && actualEnd >= 0 && actualStart <= actualEnd && actualEnd < textLength) {
            textElement.setFontFamily(actualStart, actualEnd, 'Courier New');
            textElement.setBackgroundColor(actualStart, actualEnd, '#f4f4f4');
          }
        } catch (error) {
          Logger.log(`Error applying code formatting: ${error.toString()}`);
        }
      }
    }
    
    // Apply links
    if (formatting.linkRanges && Array.isArray(formatting.linkRanges)) {
      for (const range of formatting.linkRanges) {
        try {
          const actualStart = startPos + range.start;
          const actualEnd = startPos + range.end;
          if (actualStart >= 0 && actualEnd >= 0 && actualStart <= actualEnd && actualEnd < textLength && range.url) {
            // Validate URL format
            if (typeof range.url === 'string' && range.url.trim().length > 0) {
              textElement.setLinkUrl(actualStart, actualEnd, range.url);
            }
          }
        } catch (error) {
          // URL might be invalid, skip silently but log
          Logger.log(`Error applying link (URL: ${range.url}): ${error.toString()}`);
        }
      }
    }
  } catch (error) {
    Logger.log(`Error in applyFormattingToRange: ${error.toString()}`);
    // Don't throw - formatting errors shouldn't break the conversion
  }
}

/**
 * Converts AsciiDoc text to Google Docs formatting
 */
function convertAsciiDocToGoogleDocs(rangeElements, text) {
  try {
    if (!rangeElements || rangeElements.length === 0) {
      throw new Error('No range elements provided');
    }
    
    if (!text || typeof text !== 'string') {
      throw new Error('Invalid text input');
    }
    
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      throw new Error('Unable to access document');
    }
    
    const firstElement = rangeElements[0].getElement();
    if (!firstElement) {
      throw new Error('Unable to access first element');
    }
    
    const startOffset = rangeElements[0].getStartOffset();
    const endOffset = rangeElements[rangeElements.length - 1].getEndOffsetInclusive();
    
    if (startOffset < 0 || endOffset < 0 || startOffset > endOffset) {
      throw new Error(`Invalid text range: start=${startOffset}, end=${endOffset}`);
    }
    
    const lines = text.split('\n');
    const processedLines = [];
    const formattingInstructions = [];
    
    for (let i = 0; i < lines.length; i++) {
      try {
        const line = lines[i];
        const processed = processAsciiDocLine(line);
        processedLines.push(processed.text);
        formattingInstructions.push(processed.formatting);
      } catch (error) {
        Logger.log(`Error processing AsciiDoc line ${i}: ${error.toString()}`);
        // Continue with unprocessed line
        processedLines.push(lines[i]);
        formattingInstructions.push({
          headerLevel: null,
          boldRanges: [],
          italicRanges: [],
          codeRanges: [],
          linkRanges: []
        });
      }
    }
    
    // Combine processed text
    const newText = processedLines.join('\n');
    
    // Replace the selected text
    if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
      const textElement = firstElement.asText();
      
      try {
        textElement.deleteText(startOffset, endOffset);
        textElement.insertText(startOffset, newText);
      } catch (error) {
        throw new Error(`Failed to replace text: ${error.toString()}`);
      }
      
      // Apply formatting
      try {
        let currentPos = startOffset;
        for (let i = 0; i < processedLines.length; i++) {
          const line = processedLines[i];
          const formatting = formattingInstructions[i];
          
          if (line.length > 0) {
            applyFormattingToRange(textElement, currentPos, currentPos + line.length - 1, formatting);
          }
          
          currentPos += line.length + 1; // +1 for newline
        }
      } catch (error) {
        Logger.log(`Error applying AsciiDoc formatting: ${error.toString()}`);
        // Text was replaced, but formatting failed - this is not critical
      }
    } else {
      throw new Error('Selected element is not a text element');
    }
  } catch (error) {
    Logger.log('Error in convertAsciiDocToGoogleDocs: ' + error.toString());
    throw error;
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
  try {
    if (!level || typeof level !== 'number' || level < 1 || level > 6) {
      Logger.log(`Invalid header level: ${level}, defaulting to 11`);
      return 11;
    }
    
    const sizes = {
      1: 24,
      2: 20,
      3: 16,
      4: 14,
      5: 12,
      6: 11
    };
    return sizes[level] || 11;
  } catch (error) {
    Logger.log(`Error in getHeaderSize: ${error.toString()}`);
    return 11;
  }
}

/**
 * Menu item: Convert Markdown
 */
function convertMarkdown() {
  try {
    convertSelectedText();
  } catch (error) {
    showError('Conversion Error', 'Failed to convert Markdown text.', error.toString());
    Logger.log('Error in convertMarkdown: ' + error.toString());
  }
}

/**
 * Menu item: Convert AsciiDoc
 */
function convertAsciiDoc() {
  try {
    convertSelectedText();
  } catch (error) {
    showError('Conversion Error', 'Failed to convert AsciiDoc text.', error.toString());
    Logger.log('Error in convertAsciiDoc: ' + error.toString());
  }
}

