// #region agent log
const __dbg = (location, message, data, hypothesisId) => {
    try {
        fetch('http://127.0.0.1:7407/ingest/5e83c94c-89c8-47b3-956a-93e1eab26603', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'ebd0e3' },
            body: JSON.stringify({ sessionId: 'ebd0e3', hypothesisId, location, message, data, timestamp: Date.now() })
        }).catch(() => {});
    } catch (e) {}
};
__dbg('script.js:load', 'SCRIPT LOADED build=missing-brace-v2', { ts: Date.now() }, 'H1');
window.addEventListener('error', (e) => {
    __dbg('window:onerror', 'Uncaught error', { message: e.message, source: e.filename, line: e.lineno, stack: e.error && e.error.stack }, 'H2');
});
// #endregion

// DOM Utility Functions
const DOM = {
    elements: {
        inputArea: document.getElementById('input-area'),
        processButton: document.getElementById('process-button'),
        excelButton: document.getElementById('excel-button'),
        clearButton: document.getElementById('clear-button'),
        resultContainer: document.getElementById('result-container'),
        resultContent: document.getElementById('result-content'),
        errorMessage: document.getElementById('error-message'),
        successMessage: document.getElementById('success-message'),
        copyButton: document.getElementById('copy-button'),
        viewToggleButton: document.getElementById('view-toggle-button'),
        // Schema tab
        schemaJsonArea: document.getElementById('schema-json-area'),
        schemaArea: document.getElementById('schema-area'),
        schemaValidateBtn: document.getElementById('schema-validate-btn'),
        schemaClearBtn: document.getElementById('schema-clear-btn'),
        schemaResultArea: document.getElementById('schema-result-area'),
        loadSchemaExample: document.getElementById('load-schema-example'),
        tabBtns: document.querySelectorAll('.tab-btn'),
        tabFormat: document.getElementById('tab-format'),
        tabSchema: document.getElementById('tab-schema'),
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
//
// Robust, parser-based JSON repair. Instead of guessing with a stack of
// regexes, it walks the input character-by-character and rebuilds a valid JSON
// document, repairing virtually every common class of error:
//   - missing, extra, leading, and trailing commas
//   - missing colons between keys and values
//   - missing values (filled with null) and missing closing }/]
//   - unquoted object keys and unquoted string values
//   - single quotes and smart/curly quotes -> double quotes
//   - unclosed / unterminated strings
//   - mismatched or unescaped quotes inside strings
//   - // line comments and /* block comments
//   - Markdown ``` code fences and JSONP / MongoDB function call wrappers
//   - Python literals (None / True / False) and `undefined`
//   - concatenated strings ("a" + "b"), ellipsis (1, 2, ...), regex literals
//   - Newline-Delimited JSON (NDJSON) -> a single array
//   - invalid escape sequences, unescaped control characters, bad numbers
//
// Ported from the MIT-licensed `jsonrepair` algorithm by Jos de Jong.

const JsonRepairError = class extends Error {
    constructor(message, position) {
        super(`${message} at position ${position}`);
        this.position = position;
    }
};

// ---- character classification helpers ----
const _isHex = (char) => /^[0-9A-Fa-f]$/.test(char);
const _isDigit = (char) => char >= '0' && char <= '9';
const _isValidStringCharacter = (char) => char >= '\u0020';
const _isDelimiter = (char) => ',:[]/{}()\n+'.includes(char);
const _isFunctionNameCharStart = (char) =>
    (char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z') || char === '_' || char === '$';
const _isFunctionNameChar = (char) => _isFunctionNameCharStart(char) || (char >= '0' && char <= '9');
const _regexUrlStart = /^(http|https|ftp|mailto|file|data|irc):\/\/$/;
const _regexUrlChar = /^[A-Za-z0-9-._~:/?#@!$&'()*+;=]$/;
const _isUnquotedStringDelimiter = (char) => ',[]/{}\n+'.includes(char);
const _regexStartOfValue = /^[[{\w-]$/;
const _isStartOfValue = (char) => _isQuote(char) || _regexStartOfValue.test(char);
const _isControlCharacter = (char) =>
    char === '\n' || char === '\r' || char === '\t' || char === '\b' || char === '\f';

const _codeSpace = 0x20, _codeNewline = 0xa, _codeTab = 0x9, _codeReturn = 0xd;
const _codeNonBreakingSpace = 0xa0, _codeEnQuad = 0x2000, _codeHairSpace = 0x200a;
const _codeNarrowNoBreakSpace = 0x202f, _codeMediumMathematicalSpace = 0x205f, _codeIdeographicSpace = 0x3000;

const _isWhitespace = (text, index) => {
    const code = text.charCodeAt(index);
    return code === _codeSpace || code === _codeNewline || code === _codeTab || code === _codeReturn;
};
const _isWhitespaceExceptNewline = (text, index) => {
    const code = text.charCodeAt(index);
    return code === _codeSpace || code === _codeTab || code === _codeReturn;
};
const _isSpecialWhitespace = (text, index) => {
    const code = text.charCodeAt(index);
    return code === _codeNonBreakingSpace ||
        (code >= _codeEnQuad && code <= _codeHairSpace) ||
        code === _codeNarrowNoBreakSpace ||
        code === _codeMediumMathematicalSpace ||
        code === _codeIdeographicSpace;
};

const _isQuote = (char) => _isDoubleQuoteLike(char) || _isSingleQuoteLike(char);
const _isDoubleQuoteLike = (char) => char === '"' || char === '\u201c' || char === '\u201d';
const _isDoubleQuote = (char) => char === '"';
const _isSingleQuoteLike = (char) =>
    char === "'" || char === '\u2018' || char === '\u2019' || char === '\u0060' || char === '\u00b4';
const _isSingleQuote = (char) => char === "'";

const _stripLastOccurrence = (text, textToStrip, stripRemainingText = false) => {
    const index = text.lastIndexOf(textToStrip);
    return index !== -1
        ? text.substring(0, index) + (stripRemainingText ? '' : text.substring(index + 1))
        : text;
};
const _insertBeforeLastWhitespace = (text, textToInsert) => {
    let index = text.length;
    if (!_isWhitespace(text, index - 1)) return text + textToInsert;
    while (_isWhitespace(text, index - 1)) index--;
    return text.substring(0, index) + textToInsert + text.substring(index);
};
const _removeAtIndex = (text, start, count) => text.substring(0, start) + text.substring(start + count);
const _endsWithCommaOrNewline = (text) => /[,\n][ \t\r]*$/.test(text);
const _atEndOfBlockComment = (text, i) => text[i] === '*' && text[i + 1] === '/';

const _controlCharacters = { '\b': '\\b', '\f': '\\f', '\n': '\\n', '\r': '\\r', '\t': '\\t' };
const _escapeCharacters = { '"': '"', '\\': '\\', '/': '/', b: '\b', f: '\f', n: '\n', r: '\r', t: '\t' };

/**
 * Repair a string containing an invalid JSON document and return valid JSON.
 * Throws JsonRepairError when the input cannot be repaired.
 */
function jsonRepair(text) {
    let i = 0;
    let output = '';

    parseMarkdownCodeBlock();
    const processed = parseValue();
    if (!processed) throwUnexpectedEnd();
    parseMarkdownCodeBlock();

    const processedComma = parseCharacter(',');
    if (processedComma) parseWhitespaceAndSkipComments();

    if (_isStartOfValue(text[i]) && _endsWithCommaOrNewline(output)) {
        if (!processedComma) output = _insertBeforeLastWhitespace(output, ',');
        parseNewlineDelimitedJSON();
    } else if (processedComma) {
        output = _stripLastOccurrence(output, ',');
    }

    // repair redundant end quotes/brackets
    while (text[i] === '}' || text[i] === ']') {
        i++;
        parseWhitespaceAndSkipComments();
    }

    if (i >= text.length) return output;
    throwUnexpectedCharacter();

    function parseValue() {
        parseWhitespaceAndSkipComments();
        const processed = parseObject() || parseArray() || parseString() || parseNumber() ||
            parseKeywords() || parseUnquotedString(false) || parseRegex();
        parseWhitespaceAndSkipComments();
        return processed;
    }

    function parseWhitespaceAndSkipComments(skipNewline = true) {
        const start = i;
        let changed = parseWhitespace(skipNewline);
        do {
            changed = parseComment();
            if (changed) changed = parseWhitespace(skipNewline);
        } while (changed);
        return i > start;
    }

    function parseWhitespace(skipNewline) {
        const isWhite = skipNewline ? _isWhitespace : _isWhitespaceExceptNewline;
        let whitespace = '';
        while (true) {
            if (isWhite(text, i)) {
                whitespace += text[i];
                i++;
            } else if (_isSpecialWhitespace(text, i)) {
                whitespace += ' ';
                i++;
            } else {
                break;
            }
        }
        if (whitespace.length > 0) {
            output += whitespace;
            return true;
        }
        return false;
    }

    function parseComment() {
        if (text[i] === '/' && text[i + 1] === '*') {
            while (i < text.length && !_atEndOfBlockComment(text, i)) i++;
            i += 2;
            return true;
        }
        if (text[i] === '/' && text[i + 1] === '/') {
            while (i < text.length && text[i] !== '\n') i++;
            return true;
        }
        return false;
    }

    function parseMarkdownCodeBlock() {
        if (text.slice(i, i + 3) === '```') {
            i += 3;
            if (_isFunctionNameCharStart(text[i])) {
                while (i < text.length && _isFunctionNameChar(text[i])) i++;
            }
            parseWhitespaceAndSkipComments();
            return true;
        }
        return false;
    }

    function parseCharacter(char) {
        if (text[i] === char) {
            output += text[i];
            i++;
            return true;
        }
        return false;
    }

    function skipCharacter(char) {
        if (text[i] === char) {
            i++;
            return true;
        }
        return false;
    }

    function skipEscapeCharacter() {
        return skipCharacter('\\');
    }

    function skipEllipsis() {
        parseWhitespaceAndSkipComments();
        if (text[i] === '.' && text[i + 1] === '.' && text[i + 2] === '.') {
            i += 3;
            parseWhitespaceAndSkipComments();
            skipCharacter(',');
            return true;
        }
        return false;
    }

    function parseObject() {
        if (text[i] === '{') {
            output += '{';
            i++;
            return parseObjectBody();
        }
        return false;
    }

    // Parse the body of an object (everything after the opening brace) up to and
    // including the closing brace. Assumes the opening '{' has already been
    // emitted to the output and consumed/inferred from the input.
    function parseObjectBody() {
        parseWhitespaceAndSkipComments();

        if (skipCharacter(',')) parseWhitespaceAndSkipComments();

        let initial = true;
        while (i < text.length && text[i] !== '}') {
            let processedComma;
            if (!initial) {
                processedComma = parseCharacter(',');
                if (!processedComma) output = _insertBeforeLastWhitespace(output, ',');
                parseWhitespaceAndSkipComments();
            } else {
                processedComma = true;
                initial = false;
            }

            skipEllipsis();

            const processedKey = parseString() || parseUnquotedString(true);
            if (!processedKey) {
                if (text[i] === '}' || text[i] === '{' || text[i] === ']' ||
                    text[i] === '[' || text[i] === undefined) {
                    output = _stripLastOccurrence(output, ',');
                } else {
                    throwObjectKeyExpected();
                }
                break;
            }

            parseWhitespaceAndSkipComments();
            const processedColon = parseCharacter(':');
            const truncatedText = i >= text.length;
            if (!processedColon) {
                if (_isStartOfValue(text[i]) || truncatedText) {
                    output = _insertBeforeLastWhitespace(output, ':');
                } else {
                    throwColonExpected();
                }
            }
            // Repair a missing opening brace: when a key's value is itself a
            // naked object body (e.g. `"key": "subKey": value ... }`), insert
            // the '{' and parse it as an object.
            let processedValue;
            if (isStartOfMissingBraceObject()) {
                output += '{';
                processedValue = parseObjectBody();
            } else {
                processedValue = parseValue();
            }
            if (!processedValue) {
                if (processedColon || truncatedText) {
                    output += 'null';
                } else {
                    throwColonExpected();
                }
            }
        }

        if (text[i] === '}') {
            output += '}';
            i++;
        } else {
            output = _insertBeforeLastWhitespace(output, '}');
        }
        return true;
    }

    // Look-ahead (no input consumed): true when the upcoming value is actually
    // the start of an object that is missing its opening brace, i.e. a quoted
    // key immediately followed by a colon (`"subKey":`).
    function isStartOfMissingBraceObject() {
        let j = i;
        while (j < text.length && _isWhitespace(text, j)) j++;
        if (!_isQuote(text[j])) return false;
        j++;
        while (j < text.length) {
            const c = text[j];
            if (c === '\\') { j += 2; continue; }
            if (c === '\n') return false; // unterminated string — not this case
            if (_isQuote(c)) { j++; break; }
            j++;
        }
        while (j < text.length && _isWhitespace(text, j)) j++;
        return text[j] === ':';
    }

    function parseArray() {
        if (text[i] === '[') {
            output += '[';
            i++;
            parseWhitespaceAndSkipComments();

            if (skipCharacter(',')) parseWhitespaceAndSkipComments();

            let initial = true;
            while (i < text.length && text[i] !== ']') {
                if (!initial) {
                    const processedComma = parseCharacter(',');
                    if (!processedComma) output = _insertBeforeLastWhitespace(output, ',');
                } else {
                    initial = false;
                }

                skipEllipsis();

                const processedValue = parseValue();
                if (!processedValue) {
                    output = _stripLastOccurrence(output, ',');
                    break;
                }
            }

            if (text[i] === ']') {
                output += ']';
                i++;
            } else {
                output = _insertBeforeLastWhitespace(output, ']');
            }
            return true;
        }
        return false;
    }

    function parseNewlineDelimitedJSON() {
        let initial = true;
        let processedValue = true;
        while (processedValue) {
            if (!initial) {
                const processedComma = parseCharacter(',');
                if (!processedComma) output = _insertBeforeLastWhitespace(output, ',');
            } else {
                initial = false;
            }
            processedValue = parseValue();
        }
        if (!processedValue) output = _stripLastOccurrence(output, ',');
        output = `[\n${output}\n]`;
    }

    function parseString(stopAtDelimiter = false, stopAtIndex = -1) {
        let skipEscapeChars = text[i] === '\\';
        if (skipEscapeChars) {
            i++;
            skipEscapeChars = true;
        }
        if (_isQuote(text[i])) {
            const isEndQuote = _isDoubleQuote(text[i]) ? _isDoubleQuote
                : _isSingleQuote(text[i]) ? _isSingleQuote
                : _isSingleQuoteLike(text[i]) ? _isSingleQuoteLike
                : _isDoubleQuoteLike;

            const iBefore = i;
            const oBefore = output.length;
            let str = '"';
            i++;

            while (true) {
                if (i >= text.length) {
                    const iPrev = prevNonWhitespaceIndex(i - 1);
                    if (!stopAtDelimiter && _isDelimiter(text.charAt(iPrev))) {
                        i = iBefore;
                        output = output.substring(0, oBefore);
                        return parseString(true);
                    }
                    str = _insertBeforeLastWhitespace(str, '"');
                    output += str;
                    return true;
                } else if (i === stopAtIndex) {
                    str = _insertBeforeLastWhitespace(str, '"');
                    output += str;
                    return true;
                } else if (text[i] === '\n' || text[i] === '\r') {
                    // Unterminated string: a raw newline inside a string almost
                    // always means a missing closing quote. Close the string at
                    // the end of the current line rather than swallowing the
                    // following structural characters (e.g. a closing brace).
                    str = _insertBeforeLastWhitespace(str, '"');
                    output += str;
                    return true;
                } else if (isEndQuote(text[i])) {
                    const iQuote = i;
                    const oQuote = str.length;
                    str += '"';
                    i++;
                    output += str;
                    parseWhitespaceAndSkipComments(false);
                    if (stopAtDelimiter || i >= text.length || _isDelimiter(text[i]) ||
                        _isQuote(text[i]) || _isDigit(text[i])) {
                        parseConcatenatedString();
                        return true;
                    }
                    const iPrevChar = prevNonWhitespaceIndex(iQuote - 1);
                    const prevChar = text.charAt(iPrevChar);
                    if (prevChar === ',') {
                        i = iBefore;
                        output = output.substring(0, oBefore);
                        return parseString(false, iPrevChar);
                    }
                    if (_isDelimiter(prevChar)) {
                        i = iBefore;
                        output = output.substring(0, oBefore);
                        return parseString(true);
                    }
                    output = output.substring(0, oBefore);
                    i = iQuote + 1;
                    str = `${str.substring(0, oQuote)}\\${str.substring(oQuote)}`;
                } else if (stopAtDelimiter && _isUnquotedStringDelimiter(text[i])) {
                    if (text[i - 1] === ':' && _regexUrlStart.test(text.substring(iBefore + 1, i + 2))) {
                        while (i < text.length && _regexUrlChar.test(text[i])) {
                            str += text[i];
                            i++;
                        }
                    }
                    str = _insertBeforeLastWhitespace(str, '"');
                    output += str;
                    parseConcatenatedString();
                    return true;
                } else if (text[i] === '\\') {
                    const char = text.charAt(i + 1);
                    const escapeChar = _escapeCharacters[char];
                    if (escapeChar !== undefined) {
                        str += text.slice(i, i + 2);
                        i += 2;
                    } else if (char === 'u') {
                        let j = 2;
                        while (j < 6 && _isHex(text[i + j])) j++;
                        if (j === 6) {
                            str += text.slice(i, i + 6);
                            i += 6;
                        } else if (i + j >= text.length) {
                            i = text.length;
                        } else {
                            throwInvalidUnicodeCharacter();
                        }
                    } else {
                        str += char;
                        i += 2;
                    }
                } else {
                    const char = text.charAt(i);
                    if (char === '"' && text[i - 1] !== '\\') {
                        str += `\\${char}`;
                        i++;
                    } else if (_isControlCharacter(char)) {
                        str += _controlCharacters[char];
                        i++;
                    } else {
                        if (!_isValidStringCharacter(char)) throwInvalidCharacter(char);
                        str += char;
                        i++;
                    }
                }
                if (skipEscapeChars) skipEscapeCharacter();
            }
        }
        return false;
    }

    function parseConcatenatedString() {
        let processed = false;
        parseWhitespaceAndSkipComments();
        while (text[i] === '+') {
            processed = true;
            i++;
            parseWhitespaceAndSkipComments();
            output = _stripLastOccurrence(output, '"', true);
            const start = output.length;
            const parsedStr = parseString();
            if (parsedStr) {
                output = _removeAtIndex(output, start, 1);
            } else {
                output = _insertBeforeLastWhitespace(output, '"');
            }
        }
        return processed;
    }

    function parseNumber() {
        const start = i;
        if (text[i] === '-') {
            i++;
            if (atEndOfNumber()) { repairNumberEndingWithNumericSymbol(start); return true; }
            if (!_isDigit(text[i])) { i = start; return false; }
        }
        while (_isDigit(text[i])) i++;
        if (text[i] === '.') {
            i++;
            if (atEndOfNumber()) { repairNumberEndingWithNumericSymbol(start); return true; }
            if (!_isDigit(text[i])) { i = start; return false; }
            while (_isDigit(text[i])) i++;
        }
        if (text[i] === 'e' || text[i] === 'E') {
            i++;
            if (text[i] === '-' || text[i] === '+') i++;
            if (atEndOfNumber()) { repairNumberEndingWithNumericSymbol(start); return true; }
            if (!_isDigit(text[i])) { i = start; return false; }
            while (_isDigit(text[i])) i++;
        }
        if (!atEndOfNumber()) { i = start; return false; }
        if (i > start) {
            const num = text.slice(start, i);
            const hasInvalidLeadingZero = /^0\d/.test(num);
            output += hasInvalidLeadingZero ? `"${num}"` : num;
            return true;
        }
        return false;
    }

    function parseKeywords() {
        return parseKeyword('true', 'true') || parseKeyword('false', 'false') ||
            parseKeyword('null', 'null') ||
            parseKeyword('True', 'true') || parseKeyword('False', 'false') ||
            parseKeyword('None', 'null');
    }

    function parseKeyword(name, value) {
        if (text.slice(i, i + name.length) === name) {
            output += value;
            i += name.length;
            return true;
        }
        return false;
    }

    function parseUnquotedString(isKey) {
        const start = i;
        if (_isFunctionNameCharStart(text[i])) {
            while (i < text.length && _isFunctionNameChar(text[i])) i++;
            let j = i;
            while (_isWhitespace(text, j)) j++;
            if (text[j] === '(') {
                // MongoDB function call NumberLong("2") or JSONP callback({...});
                i = j + 1;
                parseValue();
                if (text[i] === ')') {
                    i++;
                    if (text[i] === ';') i++;
                }
                return true;
            }
        }
        while (i < text.length && !_isUnquotedStringDelimiter(text[i]) && !_isQuote(text[i]) &&
            (!isKey || text[i] !== ':')) {
            i++;
        }
        if (text[i - 1] === ':' && _regexUrlStart.test(text.substring(start, i + 2))) {
            while (i < text.length && _regexUrlChar.test(text[i])) i++;
        }
        if (i > start) {
            while (_isWhitespace(text, i - 1) && i > 0) i--;
            const symbol = text.slice(start, i);
            output += symbol === 'undefined' ? 'null' : JSON.stringify(symbol);
            if (text[i] === '"') i++;
            return true;
        }
        return false;
    }

    function parseRegex() {
        if (text[i] === '/') {
            const start = i;
            i++;
            while (i < text.length && (text[i] !== '/' || text[i - 1] === '\\')) i++;
            i++;
            output += `"${text.substring(start, i)}"`;
            return true;
        }
        return false;
    }

    function prevNonWhitespaceIndex(start) {
        let prev = start;
        while (prev > 0 && _isWhitespace(text, prev)) prev--;
        return prev;
    }

    function atEndOfNumber() {
        return i >= text.length || _isDelimiter(text[i]) || _isWhitespace(text, i);
    }

    function repairNumberEndingWithNumericSymbol(start) {
        output += `${text.slice(start, i)}0`;
    }

    function throwInvalidCharacter(char) {
        throw new JsonRepairError(`Invalid character ${JSON.stringify(char)}`, i);
    }
    function throwUnexpectedCharacter() {
        throw new JsonRepairError(`Unexpected character ${JSON.stringify(text[i])}`, i);
    }
    function throwUnexpectedEnd() {
        throw new JsonRepairError('Unexpected end of json string', text.length);
    }
    function throwObjectKeyExpected() {
        throw new JsonRepairError('Object key expected', i);
    }
    function throwColonExpected() {
        throw new JsonRepairError('Colon expected', i);
    }
    function throwInvalidUnicodeCharacter() {
        const chars = text.slice(i, i + 6);
        throw new JsonRepairError(`Invalid unicode character "${chars}"`, i);
    }
}

const JsonFixer = {
    // Returns pretty-printed valid JSON, or null when the input can't be repaired.
    attemptFix(input) {
        // #region agent log
        __dbg('script.js:attemptFix', 'entry', { inputType: typeof input, inputLen: typeof input === 'string' ? input.length : null }, 'H2');
        // #endregion
        if (typeof input !== 'string') return null;
        const trimmed = input.trim();
        if (!trimmed) return null;

        const candidates = [trimmed];
        // Fully escaped JSON pasted from logs, e.g. {\"key\":\"value\"}
        if (trimmed.includes('\\"')) candidates.push(trimmed.replace(/\\"/g, '"'));

        for (const candidate of candidates) {
            try {
                const repaired = jsonRepair(candidate);
                const out = JSON.stringify(JSON.parse(repaired), null, 2);
                // #region agent log
                __dbg('script.js:attemptFix', 'jsonRepair candidate SUCCEEDED', { outLen: out.length }, 'H2');
                // #endregion
                return out;
            } catch (e) {
                // #region agent log
                __dbg('script.js:attemptFix', 'jsonRepair candidate FAILED', { error: String(e && e.message) }, 'H2');
                // #endregion
            }
        }

        // Last resort: maybe it was already valid JSON.
        try {
            return JSON.stringify(JSON.parse(trimmed), null, 2);
        } catch {
            return null;
        }
    }
};

// Main Application Logic
const JsonConverter = {
    processInput(input) {
        const trimmedInput = input.trim();
        if (!trimmedInput) throw new Error('Please enter JSON or a key-value phrase');

        // Auto-detect escaped JSON (\"key\":\"value\") and unescape silently
        if (trimmedInput.includes('\\"')) {
            const unescaped = trimmedInput.replace(/\\"/g, '"');
            try {
                const parsed = JSON.parse(unescaped);
                return {
                    result: JSON.stringify(parsed, null, 2),
                    isValidJson: true,
                    message: 'Escaped JSON detected — unescaped and formatted'
                };
            } catch {}
        }

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

            if (position < 0) return new Error('Invalid JSON — check the syntax and try again.');

            const lines = input.substring(0, position).split('\n');
            const lineNumber = lines.length;
            const column = position - input.lastIndexOf('\n', position);
            const char = input.charAt(position);

            let what = '';
            if (e.message.includes('Unexpected token') || e.message.includes('Unexpected non-whitespace')) {
                what = char === ':' ? 'missing quotes around a property name' :
                       char === ',' ? 'extra comma or missing value' :
                       (char === '}' || char === ']') ? 'missing comma between properties' :
                       char ? `unexpected character "${char}"` : 'unexpected end of input';
            } else if (e.message.includes('control character') || e.message.includes('Bad string')) {
                what = 'unclosed string (missing closing quote)';
            } else if (e.message.includes('Expected property name')) {
                what = 'property name must be in double quotes';
            } else if (e.message.includes('Unexpected end')) {
                what = 'JSON is incomplete — missing closing brackets or braces';
            } else {
                what = 'syntax error';
            }

            return new Error(`Problem at line ${lineNumber}, column ${column} — ${what}.`);
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
    },

    createFriendlyView(data, depth = 0) {
        if (data === null) return `<span class="fv-null">null</span>`;

        if (typeof data !== 'object') {
            if (typeof data === 'boolean')
                return `<span class="fv-badge ${data ? 'fv-badge-true' : 'fv-badge-false'}">${data ? 'Yes' : 'No'}</span>`;
            if (typeof data === 'number')
                return `<span class="fv-number">${data.toLocaleString()}</span>`;
            return `<span class="fv-string">${this.escapeHtml(String(data))}</span>`;
        }

        if (Array.isArray(data)) {
            if (data.length === 0) return `<em class="fv-empty">empty list</em>`;
            const allPrimitive = data.every(item => item === null || typeof item !== 'object');
            if (allPrimitive) {
                return `<div class="fv-tags">${data.map(v =>
                    `<span class="fv-tag">${this.escapeHtml(v === null ? 'null' : String(v))}</span>`
                ).join('')}</div>`;
            }
            return `<div class="fv-array-list">${data.map((item, i) => `
                <details class="fv-item-card" open>
                    <summary class="fv-item-summary"><i class="fas fa-layer-group"></i> Item ${i + 1} <span style="font-weight:400;color:#718096;margin-left:4px">of ${data.length}</span></summary>
                    <div class="fv-item-body">${this.createFriendlyView(item, depth + 1)}</div>
                </details>`).join('')}
            </div>`;
        }

        const entries = Object.entries(data);
        if (entries.length === 0) return `<em class="fv-empty">empty</em>`;

        return `<div class="fv-table${depth > 0 ? ' fv-table-nested' : ''}">
            ${entries.map(([key, value]) => {
                const isComplex = value !== null && typeof value === 'object';
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());
                const summaryIcon = isComplex
                    ? (Array.isArray(value)
                        ? `<i class="fas fa-list" style="margin-right:5px;opacity:0.6"></i>`
                        : `<i class="fas fa-folder-open" style="margin-right:5px;opacity:0.6"></i>`)
                    : '';
                if (isComplex) {
                    return `<div class="fv-row fv-row-complex">
                        <span class="fv-key">${this.escapeHtml(label)}</span>
                        <details class="fv-nested-details" ${depth < 2 ? 'open' : ''}>
                            <summary class="fv-nested-summary">${summaryIcon}</summary>
                            <div class="fv-nested-body">${this.createFriendlyView(value, depth + 1)}</div>
                        </details>
                    </div>`;
                }
                return `<div class="fv-row">
                    <span class="fv-key">${this.escapeHtml(label)}</span>
                    <span class="fv-val">${this.createFriendlyView(value, depth + 1)}</span>
                </div>`;
            }).join('')}
        </div>`;
    }
};

// JSON Schema Validator
const SchemaValidator = {
    validate(data, schema) {
        const errors = [];
        const defs = Object.assign({}, schema.$defs, schema.definitions);
        this._validate(data, schema, '', errors, defs);
        return errors;
    },

    _validate(data, schema, path, errors, defs) {
        if (!schema || typeof schema !== 'object') return;

        // Merge in any new $defs / definitions
        if (schema.$defs) Object.assign(defs, schema.$defs);
        if (schema.definitions) Object.assign(defs, schema.definitions);

        // $ref resolution
        if (schema.$ref) {
            const ref = schema.$ref;
            let refSchema = null;
            if (ref.startsWith('#/$defs/')) refSchema = defs[ref.slice('#/$defs/'.length)];
            else if (ref.startsWith('#/definitions/')) refSchema = defs[ref.slice('#/definitions/'.length)];
            if (refSchema) this._validate(data, refSchema, path, errors, defs);
            return;
        }

        const loc = path || '(root)';

        // type
        if (schema.type !== undefined) {
            const types = Array.isArray(schema.type) ? schema.type : [schema.type];
            if (!types.some(t => this._checkType(data, t))) {
                errors.push({ path: loc, message: `must be of type ${types.join(' or ')}` });
                return;
            }
        }

        // enum / const
        if (schema.enum !== undefined) {
            if (!schema.enum.some(v => JSON.stringify(v) === JSON.stringify(data)))
                errors.push({ path: loc, message: `must be one of: ${schema.enum.map(v => JSON.stringify(v)).join(', ')}` });
        }
        if (schema.const !== undefined) {
            if (JSON.stringify(data) !== JSON.stringify(schema.const))
                errors.push({ path: loc, message: `must equal ${JSON.stringify(schema.const)}` });
        }

        // String keywords
        if (typeof data === 'string') {
            if (schema.minLength !== undefined && data.length < schema.minLength)
                errors.push({ path: loc, message: `must be at least ${schema.minLength} characters long` });
            if (schema.maxLength !== undefined && data.length > schema.maxLength)
                errors.push({ path: loc, message: `must be at most ${schema.maxLength} characters long` });
            if (schema.pattern !== undefined) {
                try { if (!new RegExp(schema.pattern).test(data)) errors.push({ path: loc, message: `must match pattern "${schema.pattern}"` }); }
                catch {}
            }
            if (schema.format !== undefined) {
                const msg = this._checkFormat(data, schema.format);
                if (msg) errors.push({ path: loc, message: msg });
            }
        }

        // Number keywords
        if (typeof data === 'number') {
            if (schema.minimum !== undefined && data < schema.minimum)
                errors.push({ path: loc, message: `must be >= ${schema.minimum}` });
            if (schema.maximum !== undefined && data > schema.maximum)
                errors.push({ path: loc, message: `must be <= ${schema.maximum}` });
            if (schema.exclusiveMinimum !== undefined && data <= schema.exclusiveMinimum)
                errors.push({ path: loc, message: `must be > ${schema.exclusiveMinimum}` });
            if (schema.exclusiveMaximum !== undefined && data >= schema.exclusiveMaximum)
                errors.push({ path: loc, message: `must be < ${schema.exclusiveMaximum}` });
            if (schema.multipleOf !== undefined && data % schema.multipleOf !== 0)
                errors.push({ path: loc, message: `must be a multiple of ${schema.multipleOf}` });
        }

        // Object keywords
        if (data !== null && typeof data === 'object' && !Array.isArray(data)) {
            if (Array.isArray(schema.required)) {
                for (const key of schema.required) {
                    if (!(key in data))
                        errors.push({ path: path ? `${path}/${key}` : key, message: `required property is missing` });
                }
            }
            if (schema.properties) {
                for (const [key, sub] of Object.entries(schema.properties)) {
                    if (key in data) this._validate(data[key], sub, path ? `${path}/${key}` : key, errors, defs);
                }
            }
            if (schema.additionalProperties === false && schema.properties) {
                const allowed = new Set(Object.keys(schema.properties));
                for (const key of Object.keys(data)) {
                    if (!allowed.has(key))
                        errors.push({ path: path ? `${path}/${key}` : key, message: `additional property is not allowed` });
                }
            } else if (schema.additionalProperties && typeof schema.additionalProperties === 'object' && schema.properties) {
                const allowed = new Set(Object.keys(schema.properties));
                for (const key of Object.keys(data)) {
                    if (!allowed.has(key))
                        this._validate(data[key], schema.additionalProperties, path ? `${path}/${key}` : key, errors, defs);
                }
            }
            const numProps = Object.keys(data).length;
            if (schema.minProperties !== undefined && numProps < schema.minProperties)
                errors.push({ path: loc, message: `must have at least ${schema.minProperties} properties` });
            if (schema.maxProperties !== undefined && numProps > schema.maxProperties)
                errors.push({ path: loc, message: `must have at most ${schema.maxProperties} properties` });
        }

        // Array keywords
        if (Array.isArray(data)) {
            if (schema.minItems !== undefined && data.length < schema.minItems)
                errors.push({ path: loc, message: `must have at least ${schema.minItems} items` });
            if (schema.maxItems !== undefined && data.length > schema.maxItems)
                errors.push({ path: loc, message: `must have at most ${schema.maxItems} items` });
            if (schema.uniqueItems && data.length !== new Set(data.map(v => JSON.stringify(v))).size)
                errors.push({ path: loc, message: `must have unique items` });
            if (schema.items) {
                if (Array.isArray(schema.items)) {
                    schema.items.forEach((sub, i) => { if (i < data.length) this._validate(data[i], sub, `${path}/${i}`, errors, defs); });
                } else {
                    data.forEach((item, i) => this._validate(item, schema.items, `${path}/${i}`, errors, defs));
                }
            }
        }

        // Logical combinators
        if (schema.allOf) {
            for (const sub of schema.allOf) this._validate(data, sub, path, errors, defs);
        }
        if (schema.anyOf) {
            const valid = schema.anyOf.some(sub => { const e = []; this._validate(data, sub, path, e, defs); return e.length === 0; });
            if (!valid) errors.push({ path: loc, message: `must match at least one of the allowed schemas` });
        }
        if (schema.oneOf) {
            const count = schema.oneOf.filter(sub => { const e = []; this._validate(data, sub, path, e, defs); return e.length === 0; }).length;
            if (count !== 1) errors.push({ path: loc, message: `must match exactly one of the allowed schemas (matched ${count})` });
        }
        if (schema.not) {
            const e = []; this._validate(data, schema.not, path, e, defs);
            if (e.length === 0) errors.push({ path: loc, message: `must not match the "not" schema` });
        }
        if (schema.if) {
            const e = []; this._validate(data, schema.if, path, e, defs);
            if (e.length === 0 && schema.then) this._validate(data, schema.then, path, errors, defs);
            else if (e.length > 0 && schema.else) this._validate(data, schema.else, path, errors, defs);
        }
    },

    _checkType(data, type) {
        switch (type) {
            case 'null':    return data === null;
            case 'boolean': return typeof data === 'boolean';
            case 'integer': return typeof data === 'number' && Number.isInteger(data);
            case 'number':  return typeof data === 'number';
            case 'string':  return typeof data === 'string';
            case 'array':   return Array.isArray(data);
            case 'object':  return data !== null && typeof data === 'object' && !Array.isArray(data);
            default:        return true;
        }
    },

    _checkFormat(data, format) {
        const checks = {
            'date-time': [/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})?$/, 'must be a valid date-time (ISO 8601)'],
            'date':      [/^\d{4}-\d{2}-\d{2}$/, 'must be a valid date (YYYY-MM-DD)'],
            'time':      [/^\d{2}:\d{2}:\d{2}$/, 'must be a valid time (HH:MM:SS)'],
            'email':     [/^[^\s@]+@[^\s@]+\.[^\s@]+$/, 'must be a valid email address'],
            'uri':       [/^[a-zA-Z][a-zA-Z0-9+\-.]*:/, 'must be a valid URI'],
            'ipv4':      [/^(\d{1,3}\.){3}\d{1,3}$/, 'must be a valid IPv4 address'],
            'uuid':      [/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i, 'must be a valid UUID'],
        };
        const check = checks[format];
        if (!check) return null;
        return check[0].test(data) ? null : check[1];
    }
};

// App State
const AppState = {
    currentJsonData: null,
    viewMode: 'json' // 'json' | 'friendly'
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
            AppState.currentJsonData = JSON.parse(result);
            this.renderCurrentView();
        } catch (e) {
            AppState.currentJsonData = null;
            DOM.elements.resultContent.innerHTML = JsonConverter.syntaxHighlight(result);
            DOM.elements.resultContent.classList.add('has-line-numbers');
            DOM.elements.resultContent.classList.remove('fv-container');
        }

        DOM.showElement(DOM.elements.resultContainer);
        DOM.hideElement(DOM.elements.errorMessage);

        if (isValid || result) {
            DOM.elements.successMessage.innerHTML = `<i class="fas fa-check-circle"></i> ${message}`;
            DOM.showElement(DOM.elements.successMessage);
        }
    },

    renderCurrentView() {
        if (!AppState.currentJsonData) return;
        const el = DOM.elements.resultContent;
        el.classList.remove('has-line-numbers');
        if (AppState.viewMode === 'friendly') {
            el.classList.add('fv-container');
            el.innerHTML = JsonConverter.createFriendlyView(AppState.currentJsonData);
        } else {
            el.classList.remove('fv-container');
            el.innerHTML = JsonConverter.createInteractiveJson(AppState.currentJsonData);
            this.attachCollapseHandlers();
        }
    },

    setViewMode(mode) {
        AppState.viewMode = mode;
        const btn = DOM.elements.viewToggleButton;
        if (mode === 'friendly') {
            btn.innerHTML = '<i class="fas fa-code"></i> JSON View';
            btn.classList.add('active');
        } else {
            btn.innerHTML = '<i class="fas fa-table"></i> Friendly View';
            btn.classList.remove('active');
        }
        this.renderCurrentView();
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
        // In friendly view, copy the formatted JSON from AppState; in JSON view strip line numbers
        let text;
        if (AppState.viewMode === 'friendly' && AppState.currentJsonData) {
            text = JSON.stringify(AppState.currentJsonData, null, 2);
        } else {
            const lineContents = DOM.elements.resultContent.querySelectorAll('.line-content');
            text = lineContents.length > 0
                ? Array.from(lineContents).map(el => el.textContent || '').join('\n')
                : (DOM.elements.resultContent.textContent || DOM.elements.resultContent.innerText || '');
        }
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

// Schema UI Controller
const SchemaUI = {
    EXAMPLE_JSON: JSON.stringify({
        name: "Alice",
        age: 28,
        email: "alice@example.com",
        role: "admin",
        scores: [95, 87, 100]
    }, null, 2),

    EXAMPLE_SCHEMA: JSON.stringify({
        type: "object",
        required: ["name", "age", "email"],
        properties: {
            name:   { type: "string", minLength: 1 },
            age:    { type: "integer", minimum: 0, maximum: 120 },
            email:  { type: "string", format: "email" },
            role:   { type: "string", enum: ["admin", "user", "guest"] },
            scores: { type: "array", items: { type: "number" }, minItems: 1 }
        },
        additionalProperties: false
    }, null, 2),

    run() {
        const jsonRaw = DOM.elements.schemaJsonArea.value.trim();
        const schemaRaw = DOM.elements.schemaArea.value.trim();
        const area = DOM.elements.schemaResultArea;

        if (!jsonRaw || !schemaRaw) {
            area.style.display = 'block';
            area.innerHTML = `<div class="schema-parse-error"><strong>Missing input.</strong> Please provide both JSON data and a JSON Schema.</div>`;
            return;
        }

        let data, schema;
        try { data = JSON.parse(jsonRaw); }
        catch (e) {
            area.style.display = 'block';
            area.innerHTML = `<div class="schema-parse-error"><strong>Invalid JSON data:</strong> ${this._esc(e.message)}</div>`;
            return;
        }
        try { schema = JSON.parse(schemaRaw); }
        catch (e) {
            area.style.display = 'block';
            area.innerHTML = `<div class="schema-parse-error"><strong>Invalid JSON Schema:</strong> ${this._esc(e.message)}</div>`;
            return;
        }

        const errors = SchemaValidator.validate(data, schema);
        area.style.display = 'block';

        if (errors.length === 0) {
            area.innerHTML = `<div class="schema-valid-banner"><i class="fas fa-check-circle"></i> Valid — JSON matches the schema perfectly.</div>`;
        } else {
            const items = errors.map(err =>
                `<li class="schema-error-item">
                    <span class="schema-error-path">${this._esc(err.path)}</span>
                    <span class="schema-error-msg">${this._esc(err.message)}</span>
                </li>`
            ).join('');
            area.innerHTML = `
                <div class="schema-error-banner"><i class="fas fa-times-circle"></i> ${errors.length} validation error${errors.length > 1 ? 's' : ''} found</div>
                <ul class="schema-error-list">${items}</ul>`;
        }
    },

    clear() {
        DOM.elements.schemaJsonArea.value = '';
        DOM.elements.schemaArea.value = '';
        DOM.elements.schemaResultArea.style.display = 'none';
        DOM.elements.schemaResultArea.innerHTML = '';
    },

    _esc(str) {
        return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }
};

// Event Handlers
const EventHandlers = {
    init() {
        // Tab switching
        DOM.elements.tabBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                DOM.elements.tabBtns.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                const tab = btn.dataset.tab;
                DOM.elements.tabFormat.style.display = tab === 'format' ? '' : 'none';
                DOM.elements.tabSchema.style.display = tab === 'schema' ? '' : 'none';
            });
        });

        // Schema tab actions
        DOM.elements.schemaValidateBtn.addEventListener('click', () => SchemaUI.run());
        DOM.elements.schemaClearBtn.addEventListener('click', () => SchemaUI.clear());
        DOM.elements.loadSchemaExample.addEventListener('click', () => {
            DOM.elements.schemaJsonArea.value = SchemaUI.EXAMPLE_JSON;
            DOM.elements.schemaArea.value = SchemaUI.EXAMPLE_SCHEMA;
            DOM.elements.schemaResultArea.style.display = 'none';
        });

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

        DOM.elements.viewToggleButton.addEventListener('click', () => {
            UIController.setViewMode(AppState.viewMode === 'json' ? 'friendly' : 'json');
        });
        
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
            // #region agent log
            __dbg('script.js:fixButton', 'attemptFix returned', { fixedIsNull: fixed === null, fixedLen: fixed && fixed.length }, 'H2');
            // #endregion
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