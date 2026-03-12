// DOM Utility Functions
const DOM = {
    elements: {
        inputArea: document.getElementById('input-area'),
        processButton: document.getElementById('process-button'),
        excelButton: document.getElementById('excel-button'), // New button
        clearButton: document.getElementById('clear-button'),
        resultContainer: document.getElementById('result-container'),
        resultContent: document.getElementById('result-content'),
        errorMessage: document.getElementById('error-message'),
        successMessage: document.getElementById('success-message'),
        copyButton: document.getElementById('copy-button')
    },

    showElement(element, display = 'block') {
        element.style.display = display;
        element.classList.add('fade-in');
    },

    hideElement(element) {
        element.style.display = 'none';
    }
};

// JSON Processing Utilities
const JsonUtils = {
    isValidJSON(text) {
        try {
            JSON.parse(text);
            return true;
        } catch {
            return false;
        }
    },

    splitAtFirstColon(str) {
        const colonIndex = str.indexOf(':');
        return [str.substring(0, colonIndex), str.substring(colonIndex + 1)];
    },

    splitRespectingBrackets(input) {
        const result = [];
        let currentPart = '';
        let bracketCount = 0;
        let braceCount = 0;
        let inQuotes = false;
        let escape = false;

        for (const char of input) {
            if (escape) {
                currentPart += char;
                escape = false;
                continue;
            }

            if (char === '\\') {
                currentPart += char;
                escape = true;
                continue;
            }

            if (char === '"' && !escape) {
                inQuotes = !inQuotes;
                currentPart += char;
                continue;
            }

            if (inQuotes) {
                currentPart += char;
                continue;
            }

            if (char === '[') bracketCount++;
            else if (char === ']') bracketCount--;
            else if (char === '{') braceCount++;
            else if (char === '}') braceCount--;
            else if (char === ',' && bracketCount === 0 && braceCount === 0) {
                result.push(currentPart.trim());
                currentPart = '';
                continue;
            }
            currentPart += char;
        }

        if (currentPart.trim()) result.push(currentPart.trim());
        return result;
    },

    parseValue(value) {
        const trimmed = value.trim();
        const lower = trimmed.toLowerCase();

        if (lower === 'null') return null;
        if (lower === 'true') return true;
        if (lower === 'false') return false;
        if (!isNaN(trimmed) && trimmed !== '') return Number(trimmed);

        if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
            try {
                return JSON.parse(trimmed);
            } catch {
                const content = trimmed.slice(1, -1).trim();
                return content ? this.splitRespectingBrackets(content).map(item => this.parseValue(item.trim())) : [];
            }
        }

        if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
            try {
                return JSON.parse(trimmed);
            } catch {
                const content = trimmed.slice(1, -1).trim();
                if (!content) return {};
                
                const result = {};
                this.splitRespectingBrackets(content).forEach(part => {
                    if (part.includes(':')) {
                        const [key, val] = this.splitAtFirstColon(part);
                        result[key.trim()] = this.parseValue(val.trim());
                    }
                });
                return result;
            }
        }

        if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || 
            (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
            return trimmed.slice(1, -1);
        }

        return trimmed;
    }
};

// JSON Auto-Fixer
const JsonFixer = {
    tryParse(s) {
        try { return JSON.stringify(JSON.parse(s), null, 2); } catch { return null; }
    },

    applyBaseFixes(s) {
        // Remove trailing commas before } or ]
        s = s.replace(/,(\s*[}\]])/g, '$1');
        // Quote unquoted object keys
        s = s.replace(/([{,]\s*)([a-zA-Z_$][a-zA-Z0-9_$]*)(\s*:)/g, '$1"$2"$3');
        // Add missing commas: a value ending (string/number/bool/null/}/]) followed by newline then a new key/value
        s = s.replace(/(["\d\]}\w])(\s*\n\s*)(")/g, '$1,$2$3');
        return s;
    },

    attemptFix(input) {
        const base = this.applyBaseFixes(input.trim());

        // Strategy 1: base fixes only
        let result = this.tryParse(base);
        if (result) return result;

        // Strategy 2: replace ALL single quotes with double quotes
        result = this.tryParse(base.replace(/'/g, '"'));
        if (result) return result;

        // Strategy 3: close unterminated strings at line ends ("value\n → "value"\n)
        const lineClosed = base.replace(/"([^"\n]*)\n/g, '"$1"\n');
        result = this.tryParse(lineClosed);
        if (result) return result;

        // Strategy 4: line-closed + single-quote swap
        result = this.tryParse(lineClosed.replace(/'/g, '"'));
        if (result) return result;

        // Strategy 5: mismatched quote regex on base
        const mismatch = base
            .replace(/"([^"'\n]*?)'/g, '"$1"')
            .replace(/'([^"'\n]*?)"/g, '"$1"');
        result = this.tryParse(mismatch);
        if (result) return result;

        // Strategy 6: mismatch + line-close combined
        result = this.tryParse(mismatch.replace(/"([^"\n]*)\n/g, '"$1"\n'));
        if (result) return result;

        return null;
    }
};

// Main Application Logic
const JsonConverter = {
    processInput(input) {
        const trimmedInput = input.trim();
        if (!trimmedInput) throw new Error('Please enter JSON or a key-value phrase');

        const isJsonLike = (trimmedInput.startsWith('{') && trimmedInput.endsWith('}')) || 
                          (trimmedInput.startsWith('[') && trimmedInput.endsWith(']'));

        if (!this.hasBalancedBrackets(trimmedInput)) {
            throw this.createDetailedJsonError(trimmedInput);
        }

        if (isJsonLike) {
            try {
                const parsed = JSON.parse(trimmedInput);
                return {
                    result: JSON.stringify(parsed, null, 2),
                    isValidJson: true,
                    message: 'Valid JSON detected and formatted'
                };
            } catch (e) {
                throw this.createDetailedJsonError(trimmedInput);
            }
        }

        const result = {};
        JsonUtils.splitRespectingBrackets(trimmedInput).forEach(part => {
            if (part.includes(':')) {
                const [key, value] = JsonUtils.splitAtFirstColon(part);
                result[key.trim()] = JsonUtils.parseValue(value);
            }
        });

        return {
            result: JSON.stringify(result, null, 2),
            isValidJson: false,
            message: 'Successfully converted phrase to JSON'
        };
    },

    convertToExcel(input) {
        const trimmedInput = input.trim();
        if (!trimmedInput) throw new Error('Please enter JSON to convert to Excel');

        let jsonData;
        if (JsonUtils.isValidJSON(trimmedInput)) {
            jsonData = JSON.parse(trimmedInput);
        } else {
            const result = {};
            JsonUtils.splitRespectingBrackets(trimmedInput).forEach(part => {
                if (part.includes(':')) {
                    const [key, value] = JsonUtils.splitAtFirstColon(part);
                    result[key.trim()] = JsonUtils.parseValue(value);
                }
            });
            jsonData = result;
        }

        // Tab 1: Non-Array Values
        const nonArrayData = [];
        for (const [key, value] of Object.entries(jsonData)) {
            if (!Array.isArray(value) && typeof value !== 'object') {
                nonArrayData.push({ Key: key, Value: value });
            }
        }

        // Tab 2: Server Topology
        const serverTopologyData = [];
        const topologyString = jsonData.serversTopology || '[]';
        const topologyEntries = topologyString.slice(1, -1).split(',');
        topologyEntries.forEach(entry => {
            if (entry.trim()) {
                const [mainPart, updated] = entry.split(':');
                const [serverName, provider, status, type] = mainPart.split('-');
                serverTopologyData.push({
                    'Server Name': serverName,
                    'Provider': provider,
                    'Status': status,
                    'Type': type,
                    'Updated?': updated === '1' ? 'Yes' : 'No'
                });
            }
        });

        // Create Workbook
        const wb = XLSX.utils.book_new();
        const ws1 = XLSX.utils.json_to_sheet(nonArrayData);
        XLSX.utils.book_append_sheet(wb, ws1, 'Non-Array Values');
        const ws2 = XLSX.utils.json_to_sheet(serverTopologyData);
        XLSX.utils.book_append_sheet(wb, ws2, 'Server Topology');

        // Download Excel File
        XLSX.writeFile(wb, 'converted_data.xlsx');
    },

    hasBalancedBrackets(str) {
        let stack = [];
        for (let char of str) {
            if (char === '{') stack.push('}');
            else if (char === '[') stack.push(']');
            else if (char === '(') stack.push(')');
            else if (char === '}' || char === ']' || char === ')') {
                if (stack.length === 0 || stack.pop() !== char) {
                    return false;
                }
            }
        }
        return stack.length === 0;
    },

    createDetailedJsonError(input) {
        try {
            JSON.parse(input);
        } catch (e) {
            const positionMatch = e.message.match(/position (\d+)/);
            const position = positionMatch ? parseInt(positionMatch[1]) : -1;
            
            if (position < 0) return new Error(`Invalid JSON: ${e.message}`);

            const lines = input.substring(0, position).split('\n');
            const lineNumber = lines.length;
            const column = position - input.lastIndexOf('\n', position);
            const lineStart = input.lastIndexOf('\n', position) + 1;
            const lineEnd = input.indexOf('\n', position);
            const line = input.substring(lineStart, lineEnd > -1 ? lineEnd : input.length);
            
            let suggestion = '';
            const char = input.charAt(position);
            if (e.message.includes('Unexpected token')) {
                suggestion = char === ':' ? 'Missing quotation marks around a property name?' :
                           char === ',' ? 'Extra comma or missing property?' :
                           (char === '}' || char === ']') ? 'Missing comma between properties?' : '';
            } else if (e.message.includes('control character')) {
                suggestion = 'Missing closing quotation mark?';
            } else if (e.message.includes('Expected property name')) {
                suggestion = 'Property name must be in double quotes.';
            }

            return new Error(`Invalid JSON at line ${lineNumber}, column ${column}: ${e.message}\n\n${line}\n${' '.repeat(Math.max(0, column - 1))}^ Error is here\n\nSuggestion: ${suggestion || 'Check syntax near this position.'}`);
        }
    },

    syntaxHighlight(json) {
        return json
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, match => {
                const cls = /^"/.test(match) ? (/:$/.test(match) ? 'json-key' : 'json-string') :
                           /true|false/.test(match) ? 'json-boolean' :
                           /null/.test(match) ? 'json-null' : 'json-number';
                return `<span class="${cls}">${match}</span>`;
            })
            .replace(/([{}[\]])/g, '<span class="json-braces">$1</span>')
            .split('\n')
            .map(line => `<span class="line">${line}</span>`)
            .join('\n');
    },

    createInteractiveJson(data) {
        // First, format JSON as string with proper indentation
        const jsonString = JSON.stringify(data, null, 2);
        const lines = jsonString.split('\n');
        
        // Track which lines are collapsible (start of objects/arrays)
        const collapsibleLines = new Map(); // line index -> {id, type, openLine, closeLine}
        const lineGroups = new Map(); // Track which lines belong to which group (excluding nested groups)
        const idCounter = { count: 0 };
        const stack = [];
        
        // First pass: find all object/array boundaries
        lines.forEach((line, index) => {
            const trimmed = line.trim();
            // Check if line ends with { or [ (could be on same line as key, or standalone)
            if (trimmed.endsWith('{') || trimmed.endsWith('[')) {
                const id = `json-${idCounter.count++}`;
                const isObject = trimmed.endsWith('{');
                collapsibleLines.set(index, { id, type: isObject ? 'object' : 'array', openLine: index });
                stack.push({ lineIndex: index, id, startIndex: index });
            } 
            // Check if line starts with } or ] (could be standalone or have content before)
            // Also check if line contains } or ] at the start (after trimming)
            else if (trimmed.startsWith('}') || trimmed.startsWith(']') || 
                     (trimmed.length > 0 && (trimmed[0] === '}' || trimmed[0] === ']'))) {
                if (stack.length > 0) {
                    const open = stack.pop();
                    const lineInfo = collapsibleLines.get(open.lineIndex);
                    if (lineInfo) {
                        lineInfo.closeLine = index;
                    }
                }
            }
        });
        
        // Second pass: assign lines to groups, excluding nested sections
        for (let [openIndex, lineInfo] of collapsibleLines.entries()) {
            const closeIndex = lineInfo.closeLine;
            if (closeIndex === undefined) continue;
            
            // Find all nested collapsible sections within this one
            const nestedRanges = [];
            for (let [nestedIndex, nestedInfo] of collapsibleLines.entries()) {
                if (nestedIndex > openIndex && nestedIndex < closeIndex && nestedInfo.closeLine) {
                    nestedRanges.push({ start: nestedIndex, end: nestedInfo.closeLine });
                }
            }
            // Sort by start position
            nestedRanges.sort((a, b) => a.start - b.start);
            
            // Mark lines that are not inside nested ranges
            for (let i = openIndex + 1; i < closeIndex; i++) {
                let isInNestedRange = false;
                for (const range of nestedRanges) {
                    if (i >= range.start && i <= range.end) {
                        isInNestedRange = true;
                        break;
                    }
                }
                if (!isInNestedRange) {
                    lineGroups.set(i, lineInfo.id);
                }
            }
        }
        
        // Build HTML with line numbers and chevrons
        const html = lines.map((line, index) => {
            const lineInfo = collapsibleLines.get(index);
            // Only mark as collapsible if it has a valid closeLine (was properly matched)
            const isCollapsible = lineInfo && lineInfo.openLine === index && lineInfo.closeLine !== undefined;
            const lineNum = index + 1;
            const groupId = lineGroups.get(index);
            
            // Apply syntax highlighting
            let highlightedLine = this.syntaxHighlightLine(line);
            
            if (isCollapsible) {
                return `<div class="json-line json-collapsible-line" data-line="${lineNum}" data-group="${lineInfo.id}" data-open-line="${lineInfo.openLine}" data-close-line="${lineInfo.closeLine}">
                    <span class="line-number">${lineNum}</span>
                    <span class="line-toggle" data-target="${lineInfo.id}">
                        <i class="fas fa-chevron-down"></i>
                    </span>
                    <span class="line-content">${highlightedLine}</span>
                </div>`;
            } else if (groupId) {
                return `<div class="json-line json-grouped-line" data-line="${lineNum}" data-group="${groupId}">
                    <span class="line-number">${lineNum}</span>
                    <span class="line-toggle"></span>
                    <span class="line-content">${highlightedLine}</span>
                </div>`;
            } else {
                return `<div class="json-line" data-line="${lineNum}">
                    <span class="line-number">${lineNum}</span>
                    <span class="line-toggle"></span>
                    <span class="line-content">${highlightedLine}</span>
                </div>`;
            }
        }).join('');
        
        return html;
    },

    syntaxHighlightLine(line) {
        return line
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, match => {
                const cls = /^"/.test(match) ? (/:$/.test(match) ? 'json-key' : 'json-string') :
                           /true|false/.test(match) ? 'json-boolean' :
                           /null/.test(match) ? 'json-null' : 'json-number';
                return `<span class="${cls}">${match}</span>`;
            })
            .replace(/([{}[\]])/g, '<span class="json-braces">$1</span>');
    },

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
};

// UI Controller
const UIController = {
    clear() {
        DOM.elements.inputArea.value = '';
        DOM.hideElement(DOM.elements.resultContainer);
        DOM.hideElement(DOM.elements.errorMessage);
        DOM.hideElement(DOM.elements.successMessage);
    },

    showError(message, input = null) {
        const lines = message.split('\n');
        const firstLine = lines[0];
        const rest = lines.slice(1).join('\n');
        const detailHtml = rest
            ? `<pre class="error-detail">${rest.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</pre>`
            : '';
        const fixPopup = input
            ? `<div class="fix-popup"><button class="fix-btn" id="fix-json-btn"><i class="fas fa-wrench"></i> Fix it for me?</button></div>`
            : '';
        DOM.elements.errorMessage.innerHTML =
            `<div class="error-main"><i class="fas fa-exclamation-circle"></i> ${firstLine}${detailHtml}</div>${fixPopup}`;
        DOM.elements.errorMessage._fixInput = input;
        DOM.showElement(DOM.elements.errorMessage);
        DOM.hideElement(DOM.elements.successMessage);
        DOM.hideElement(DOM.elements.resultContainer);
    },

    showResult(result, isValid, message) {
        try {
            const jsonData = JSON.parse(result);
            DOM.elements.resultContent.innerHTML = JsonConverter.createInteractiveJson(jsonData);
            DOM.elements.resultContent.classList.remove('has-line-numbers');
            this.attachCollapseHandlers();
        } catch (e) {
            // Fallback to syntax highlighting if parsing fails
            DOM.elements.resultContent.innerHTML = JsonConverter.syntaxHighlight(result);
            DOM.elements.resultContent.classList.add('has-line-numbers');
        }
        
        DOM.showElement(DOM.elements.resultContainer);
        DOM.hideElement(DOM.elements.errorMessage);
        
        if (isValid || result) {
            DOM.elements.successMessage.innerHTML = `<i class="fas fa-check-circle"></i> ${message}`;
            DOM.showElement(DOM.elements.successMessage);
        }
    },

    attachCollapseHandlers() {
        const toggleElement = (targetId) => {
            const startLine = DOM.elements.resultContent.querySelector(`.json-collapsible-line[data-group="${targetId}"]`);
            
            if (!startLine) return;
            
            const openLineNum = parseInt(startLine.getAttribute('data-open-line'));
            const closeLineNum = parseInt(startLine.getAttribute('data-close-line'));
            const isCollapsed = startLine.classList.contains('collapsed');
            
            // Get all lines in the result container
            const allLines = Array.from(DOM.elements.resultContent.querySelectorAll('.json-line'));
            
            // Build a map of nested collapsed sections for quick lookup
            // CRITICAL: We need to find ALL nested collapsed objects, not just when expanding
            // This ensures that when we collapse a parent, nested collapsed objects stay collapsed
            const nestedCollapsedRanges = [];
            allLines.forEach(checkLine => {
                if (checkLine.classList.contains('json-collapsible-line') && 
                    checkLine !== startLine &&
                    checkLine.classList.contains('collapsed')) {
                    const checkOpen = parseInt(checkLine.getAttribute('data-open-line'));
                    const checkClose = parseInt(checkLine.getAttribute('data-close-line'));
                    // Only include if it's within our range
                    if (checkOpen > openLineNum && checkClose < closeLineNum) {
                        nestedCollapsedRanges.push({ start: checkOpen, end: checkClose });
                    }
                }
            });
            
            // Hide/show all lines between open and close
            // CRITICAL FIX: We must hide ALL lines between openLineNum and closeLineNum,
            // including nested collapsible sections (which have their own group IDs)
            
            // Build a set of line numbers that should stay hidden (nested collapsed ranges)
            // CRITICAL FIX: Always preserve nested collapsed objects, whether we're collapsing or expanding parent
            const nestedCollapsedLineSet = new Set();
            // When expanding parent, we need to keep nested collapsed objects hidden
            // When collapsing parent, we hide everything anyway, but we still track nested ranges
            // to ensure they stay collapsed when parent is expanded again
            for (const range of nestedCollapsedRanges) {
                // Include only nested CONTENT (exclude the opening line) so when we expand the parent
                // we SHOW the nested object's row (e.g. "contract") but keep its content hidden
                for (let lineIdx = range.start + 1; lineIdx <= range.end; lineIdx++) {
                    nestedCollapsedLineSet.add(lineIdx);
                }
            }
            
            // More robust approach: iterate through all lines and check if they're in range
            let linesProcessed = 0;
            let linesHidden = 0;
            let linesShown = 0;
            let linesSkippedNested = 0;
            let linesWithMissingData = 0;
            
            for (let i = 0; i < allLines.length; i++) {
                const line = allLines[i];
                const lineNumAttr = line.getAttribute('data-line');
                
                if (!lineNumAttr) {
                    linesWithMissingData++;
                    continue; // Skip if no line number
                }
                
                const lineNum = parseInt(lineNumAttr, 10);
                if (isNaN(lineNum)) {
                    linesWithMissingData++;
                    continue; // Skip if invalid line number
                }
                
                // Convert to 0-based index for comparison with openLineNum/closeLineNum
                const lineNum0Based = lineNum - 1;
                
                // Only process lines between open and close (excluding the closing brace line itself)
                // openLineNum and closeLineNum are already 0-based
                if (lineNum0Based > openLineNum && lineNum0Based < closeLineNum) {
                    linesProcessed++;
                    
                    // Check if this line should stay hidden (inside a nested collapsed section)
                    const shouldStayHidden = nestedCollapsedLineSet.has(lineNum0Based);
                    
                    if (shouldStayHidden) {
                        // This line is inside a nested collapsed object - keep it hidden
                        line.style.display = 'none';
                        linesSkippedNested++;
                    } else {
                        // Toggle visibility for this line
                        // isCollapsed=true means currently collapsed, so we're expanding (show lines)
                        // isCollapsed=false means currently expanded, so we're collapsing (hide lines)
                        const newDisplay = isCollapsed ? '' : 'none';
                        line.style.display = newDisplay;
                        if (newDisplay === 'none') {
                            linesHidden++;
                        } else {
                            linesShown++;
                        }
                    }
                }
            }
            
            // Update toggle icon
            const toggle = startLine.querySelector('.line-toggle[data-target]');
            if (toggle) {
                const icon = toggle.querySelector('i');
                if (icon) {
                    icon.classList.toggle('fa-chevron-down', isCollapsed);
                    icon.classList.toggle('fa-chevron-right', !isCollapsed);
                }
                startLine.classList.toggle('collapsed', !isCollapsed);
            }
        };
        
        // Attach handlers to toggle icons
        const toggles = DOM.elements.resultContent.querySelectorAll('.line-toggle[data-target]');
        toggles.forEach(toggle => {
            toggle.addEventListener('click', (e) => {
                e.stopPropagation();
                const targetId = toggle.getAttribute('data-target');
                toggleElement(targetId);
            });
        });
    },

    showExcelSuccess() {
        DOM.elements.successMessage.innerHTML = `<i class="fas fa-check-circle"></i> Excel file generated successfully!`;
        DOM.showElement(DOM.elements.successMessage);
        DOM.hideElement(DOM.elements.errorMessage);
    },

    async copyToClipboard() {
        // Get JSON text without line numbers: use .line-content only (interactive view)
        const lineContents = DOM.elements.resultContent.querySelectorAll('.line-content');
        const text = lineContents.length > 0
            ? Array.from(lineContents).map(el => el.textContent || '').join('\n')
            : (DOM.elements.resultContent.textContent || DOM.elements.resultContent.innerText || '');
        const button = DOM.elements.copyButton;
        const originalHTML = button.innerHTML;

        try {
            await navigator.clipboard.writeText(text);
            this.showCopySuccess(button, originalHTML);
        } catch {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            document.body.appendChild(textarea);
            textarea.select();

            try {
                document.execCommand('copy');
                this.showCopySuccess(button, originalHTML);
            } catch (err) {
                console.error('Copy failed:', err);
            }
            document.body.removeChild(textarea);
        }
    },

    showCopySuccess(button, originalHTML) {
        button.innerHTML = '<i class="fas fa-check"></i> Copied!';
        button.style.backgroundColor = '#28a745';
        setTimeout(() => {
            button.innerHTML = originalHTML;
            button.style.backgroundColor = '';
        }, 2000);
    }
};

// Event Handlers
const EventHandlers = {
    init() {
        DOM.elements.clearButton.addEventListener('click', () => UIController.clear());
        
        DOM.elements.processButton.addEventListener('click', () => {
            try {
                const { result, isValidJson, message } = JsonConverter.processInput(DOM.elements.inputArea.value);
                UIController.showResult(result, isValidJson, message);
            } catch (err) {
                UIController.showError(err.message, DOM.elements.inputArea.value);
            }
        });

        DOM.elements.excelButton.addEventListener('click', () => {
            try {
                JsonConverter.convertToExcel(DOM.elements.inputArea.value);
                UIController.showExcelSuccess();
            } catch (err) {
                UIController.showError(err.message);
            }
        });

        DOM.elements.copyButton.addEventListener('click', () => UIController.copyToClipboard());
        
        DOM.elements.inputArea.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
                DOM.elements.processButton.click();
            }
        });

        DOM.elements.errorMessage.addEventListener('click', (e) => {
            if (!e.target.closest('#fix-json-btn')) return;
            const input = DOM.elements.errorMessage._fixInput;
            if (!input) return;
            const fixed = JsonFixer.attemptFix(input);
            if (fixed) {
                DOM.elements.inputArea.value = fixed;
                DOM.elements.processButton.click();
            } else {
                UIController.showError('Could not auto-fix the JSON. Please correct it manually.');
            }
        });
    }
};

// Initialize application
EventHandlers.init();