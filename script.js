// DOM Utility Functions
const DOM = {
    elements: {
        inputArea: document.getElementById('input-area'),
        processButton: document.getElementById('process-button'),
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

// Main Application Logic
const JsonConverter = {
    processInput(input) {
        const trimmedInput = input.trim();
        if (!trimmedInput) throw new Error('Please enter JSON or a key-value phrase');

        const isJsonLike = (trimmedInput.startsWith('{') && trimmedInput.endsWith('}')) || 
                          (trimmedInput.startsWith('[') && trimmedInput.endsWith(']'));

        if (JsonUtils.isValidJSON(trimmedInput)) {
            const parsed = JSON.parse(trimmedInput);
            return {
                result: JSON.stringify(parsed, null, 2),
                isValidJson: true,
                message: 'Valid JSON detected and formatted'
            };
        }

        if (isJsonLike) throw this.createDetailedJsonError(trimmedInput);

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

            return new Error(`Invalid JSON at line ${lineNumber}, column ${column}: ${e.message}\n\n${line}\n${' '.repeat(column - 1)}^ Error is here\n\nSuggestion: ${suggestion || 'Check syntax near this position.'}`);
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

    showError(message) {
        DOM.elements.errorMessage.innerHTML = `<i class="fas fa-exclamation-circle"></i> ${message}`;
        DOM.showElement(DOM.elements.errorMessage);
        DOM.hideElement(DOM.elements.successMessage);
        DOM.hideElement(DOM.elements.resultContainer);
    },

    showResult(result, isValid, message) {
        DOM.elements.resultContent.innerHTML = JsonConverter.syntaxHighlight(result);
        DOM.showElement(DOM.elements.resultContainer);
        DOM.hideElement(DOM.elements.errorMessage);
        
        if (isValid || result) {
            DOM.elements.successMessage.innerHTML = `<i class="fas fa-check-circle"></i> ${message}`;
            DOM.showElement(DOM.elements.successMessage);
        }
    },

    async copyToClipboard() {
        const text = DOM.elements.resultContent.textContent;
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
        button.innerHTML = '<i class="fas fa-check"></i>Copied!';
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
                UIController.showError(err.message);
            }
        });

        DOM.elements.copyButton.addEventListener('click', () => UIController.copyToClipboard());
        
        DOM.elements.inputArea.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
                DOM.elements.processButton.click();
            }
        });
    }
};

// Initialize application
EventHandlers.init();