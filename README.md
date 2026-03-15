# JSON Toolbox

A simple web-based tool for converting, validating, and formatting JSON. Convert key-value phrases to JSON, validate syntax, format output, and export to Excel—all in your browser.

## Features

- **Validate & Format JSON** — Paste or type JSON and get formatted, syntax-highlighted output with helpful error messages when invalid
- **Phrase to JSON** — Convert natural key-value phrases (e.g. `name: John, age: 30, hobbies: [reading, gaming]`) into valid JSON
- **Collapsible View** — Expand and collapse objects and arrays to navigate large JSON structures easily (similar to Postman)
- **Friendly View** — Switch between raw JSON and a table-friendly view for easier reading
- **Copy to Clipboard** — Copy the formatted JSON without line numbers for clean paste into other tools
- **Export to Excel** — Convert JSON data to an Excel file with multiple sheets (e.g. Non-Array Values, Server Topology)

## How to Use

1. **Process JSON or Phrase**
   - Paste valid JSON or type key-value pairs in the input area
   - Click **Process** (or press Ctrl+Enter) to validate and format
   - Invalid JSON shows detailed error messages with line and column hints

2. **Navigate Large JSON**
   - Use the chevron icons next to objects and arrays to collapse or expand sections
   - Collapsed nested objects stay collapsed when you expand their parent
   - Line numbers appear on the left for reference

3. **Copy**
   - Click **Copy to Clipboard** to copy the JSON without line numbers
   - Paste into editors, APIs, or other tools

4. **Export to Excel**
   - Enter JSON in the input area
   - Click **Convert to Excel** to download an `.xlsx` file

5. **Clear**
   - Click **Clear** to reset the input and results

## Getting Started

1. go to https://ofirka2.github.io/Json-Validator/
2. No build step or server required—runs entirely in the browser

## Tech Stack

- Plain HTML, CSS, and JavaScript
- [Font Awesome](https://fontawesome.com/) for icons
- [SheetJS (xlsx)](https://sheetjs.com/) for Excel export

## License

Ofirka JSON Toolbox © 2026
