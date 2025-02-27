// DOM Elements
const inputArea = document.getElementById('input-area');
const processButton = document.getElementById('process-button');
const clearButton = document.getElementById('clear-button');
const resultContainer = document.getElementById('result-container');
const resultContent = document.getElementById('result-content');
const errorMessage = document.getElementById('error-message');
const successMessage = document.getElementById('success-message');
const copyButton = document.getElementById('copy-button');

// Clear button handler
clearButton.addEventListener('click', () => {
    inputArea.value = '';
    clearResults();
});

// Function to detect if input is already valid JSON
function isValidJSON(text) {
    try {
        JSON.parse(text);
        return true;
    } catch (e) {
        return false;
    }
}

// Helper function to split string at first colon
function splitAtFirstColon(str) {
    const colonIndex = str.indexOf(':');
    return [str.substring(0, colonIndex), str.substring(colonIndex + 1)];
}

// Helper function to split by commas but respect brackets and braces
function splitRespectingBrackets(input) {
    const result = [];
    let currentPart = '';
    let bracketCount = 0;
    let braceCount = 0;
    let inQuotes = false;
    let escape = false;
    
    for (let i = 0; i < input.length; i++) {
        const char = input[i];
        
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
        
        if (char === '[') {
            bracketCount++;
            currentPart += char;
        } else if (char === ']') {
            bracketCount--;
            currentPart += char;
        } else if (char === '{') {
            braceCount++;
            currentPart += char;
        } else if (char === '}') {
            braceCount--;
            currentPart += char;
        } else if (char === ',' && bracketCount === 0 && braceCount === 0) {
            result.push(currentPart.trim());
            currentPart = '';
        } else {
            currentPart += char;
        }
    }
    
    if (currentPart.trim()) {
        result.push(currentPart.trim());
    }
    
    return result;
}

// Helper function to parse different value types
function parseValue(value) {
    // Handle null, booleans, and numbers
    if (value.toLowerCase() === 'null') {
        return null;
    } else if (value.toLowerCase() === 'true') {
        return true;
    } else if (value.toLowerCase() === 'false') {
        return false;
    } else if (!isNaN(value) && value !== '') {
        return Number(value);
    }
    
    // Handle arrays
    if (value.startsWith('[') && value.endsWith(']')) {
        try {
            // Try to parse it as a JSON array
            return JSON.parse(value);
        } catch (e) {
            // If that fails, do manual array parsing
            const arrayContent = value.substring(1, value.length - 1).trim();
            if (!arrayContent) return [];
            
            const items = splitRespectingBrackets(arrayContent);
            return items.map(item => parseValue(item.trim()));
        }
    }
    
    // Handle nested objects
    if (value.startsWith('{') && value.endsWith('}')) {
        try {
            // Try to parse it as a JSON object
            return JSON.parse(value);
        } catch (e) {
            // If that fails, recursively process it
            const objectContent = value.substring(1, value.length - 1).trim();
            if (!objectContent) return {};
            
            const nestedResult = {};
            const nestedParts = splitRespectingBrackets(objectContent);
            
            nestedParts.forEach(part => {
                if (part.includes(':')) {
                    let [nestedKey, nestedValue] = splitAtFirstColon(part);
                    nestedKey = nestedKey.trim();
                    nestedValue = nestedValue.trim();
                    
                    nestedResult[nestedKey] = parseValue(nestedValue);
                }
            });
            
            return nestedResult;
        }
    }
    
    // Handle quoted strings
    if ((value.startsWith('"') && value.endsWith('"')) || 
        (value.startsWith("'") && value.endsWith("'"))) {
        return value.substring(1, value.length - 1);
    }
    
    // Default case: return as string
    return value;
}

// Unified process function that handles both conversion and validation
function processJsonInput(input) {
    input = input.trim();
    
    if (!input) {
        throw new Error('Please enter JSON or a key-value phrase');
    }
    
    // Check if input looks like JSON (starts with { or [)
    const looksLikeJson = (input.startsWith('{') && input.endsWith('}')) || 
                          (input.startsWith('[') && input.endsWith(']'));
    
    // First check if the input is already valid JSON
    try {
        // Try to parse as JSON
        const parsed = JSON.parse(input);
        
        // Return the original input, just with proper formatting
        return {
            result: JSON.stringify(parsed, null, 2),
            isValidJson: true,
            message: 'Valid JSON detected and formatted'
        };
    } catch (e) {
        // If it looks like JSON but failed to parse, it's invalid JSON
        if (looksLikeJson) {
            throw new Error('Invalid JSON: ' + e.message);
        }
        
        // Not valid JSON, try to convert from phrase
        try {
            // Process as comma-separated key-values
            const result = {};
            
            // Split by commas but respect brackets and parentheses
            const parts = splitRespectingBrackets(input);
            
            parts.forEach(part => {
                // Check for key-value pairs (contains a colon)
                if (part.includes(':')) {
                    let [key, value] = splitAtFirstColon(part);
                    key = key.trim();
                    value = value.trim();
                    
                    // Handle different value types
                    result[key] = parseValue(value);
                }
            });
            
            // Format the JSON with proper indentation
            return {
                result: JSON.stringify(result, null, 2),
                isValidJson: false,
                message: 'Successfully converted phrase to JSON'
            };
        } catch (err) {
            throw new Error('Could not process input: ' + err.message);
        }
    }
}

// Function to syntax highlight JSON
function syntaxHighlight(json) {
    let result = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    result = result.replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, function (match) {
        let cls = 'json-number';
        if (/^"/.test(match)) {
            if (/:$/.test(match)) {
                cls = 'json-key';
            } else {
                cls = 'json-string';
            }
        } else if (/true|false/.test(match)) {
            cls = 'json-boolean';
        } else if (/null/.test(match)) {
            cls = 'json-null';
        }
        return '<span class="' + cls + '">' + match + '</span>';
    });
    
    result = result.replace(/([{}[\]])/g, '<span class="json-braces">$1</span>');
    
    return result;
}

// Clear results and errors
function clearResults() {
    resultContainer.style.display = 'none';
    errorMessage.style.display = 'none';
    successMessage.style.display = 'none';
}

// Show error
function showError(message) {
    errorMessage.innerHTML = '<i class="fas fa-exclamation-circle"></i> ' + message;
    errorMessage.style.display = 'block';
    errorMessage.classList.add('fade-in');
    successMessage.style.display = 'none';
    resultContainer.style.display = 'none';
}

// Show result
function showResult(resultText, isValid = false) {
    // First, apply syntax highlighting
    const highlightedText = syntaxHighlight(resultText);
    
    // Split the text by newlines and wrap each line in a span
    const lines = highlightedText.split('\n');
    const linesWithNumbers = lines.map(line => 
        `<span class="line">${line}</span>`
    ).join('\n');
    
    // Add the formatted text to the result container
    resultContent.innerHTML = linesWithNumbers;
    resultContainer.style.display = 'block';
    errorMessage.style.display = 'none';
    
    // Show success message for valid JSON
    if (isValid) {
        successMessage.style.display = 'block';
        successMessage.classList.add('fade-in');
    } else {
        successMessage.style.display = 'none';
    }
}

// Process button click handler
processButton.addEventListener('click', () => {
    const input = inputArea.value;
    
    try {
        const { result, isValidJson, message } = processJsonInput(input);
        showResult(result, isValidJson);
        
        if (isValidJson || result) {
            successMessage.innerHTML = '<i class="fas fa-check-circle"></i> ' + message;
            successMessage.style.display = 'block';
            successMessage.classList.add('fade-in');
        }
    } catch (err) {
        showError(err.message);
    }
});

// Copy to clipboard functionality
copyButton.addEventListener('click', function() {
    const textToCopy = resultContent.textContent;
    
    navigator.clipboard.writeText(textToCopy)
        .then(() => {
            const originalHTML = copyButton.innerHTML;
            copyButton.innerHTML = '<i class="fas fa-check"></i>Copied!';
            copyButton.style.backgroundColor = '#28a745';
            
            setTimeout(() => {
                copyButton.innerHTML = originalHTML;
                copyButton.style.backgroundColor = '';
            }, 2000);
        })
        .catch(() => {
            // Fallback for browsers that don't support clipboard API
            const textarea = document.createElement('textarea');
            textarea.value = textToCopy;
            textarea.style.position = 'fixed';
            document.body.appendChild(textarea);
            textarea.select();
            
            try {
                document.execCommand('copy');
                const originalHTML = copyButton.innerHTML;
                copyButton.innerHTML = '<i class="fas fa-check"></i>Copied!';
                copyButton.style.backgroundColor = '#28a745';
                
                setTimeout(() => {
                    copyButton.innerHTML = originalHTML;
                    copyButton.style.backgroundColor = '';
                }, 2000);
            } catch (err) {
                console.error('Failed to copy', err);
            }
            
            document.body.removeChild(textarea);
        });
});

// Handle Enter key in textarea
inputArea.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && e.ctrlKey) {
        processButton.click();
    }
});