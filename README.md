# n8n-nodes-pdf4me-excel

This is an n8n community node that enables you to process Excel documents with PDF4ME's powerful Excel processing capabilities. Add customizable headers and footers to your Excel worksheets with full control over styling and positioning.

n8n is a fair-code licensed workflow automation platform.

## Table of Contents

- [Installation](#installation)
- [Operations](#operations)
- [Credentials](#credentials)
- [Usage](#usage)
- [Resources](#resources)
- [Version History](#version-history)

## Installation

### Community Nodes (Recommended)

For users on n8n v0.187+, you can install this node directly from the n8n Community Nodes panel in the n8n editor:

1. Open your n8n editor
2. Go to **Settings > Community Nodes**
3. Search for "n8n-nodes-pdf4me-excel"
4. Click **Install**
5. Reload the editor

### Manual Installation

You can also install this node manually in a specific n8n project:

1. Navigate to your n8n installation directory
2. Run the following command:
   ```bash
   npm install n8n-nodes-pdf4me-excel
   ```
3. Restart your n8n server

For Docker-based deployments, add the package to your package.json and rebuild the image:

```json
{
  "name": "n8n-custom",
  "version": "1.0.0",
  "description": "",
  "main": "index.js",
  "scripts": {
    "start": "n8n"
  },
  "dependencies": {
    "n8n": "^1.0.0",
    "n8n-nodes-pdf4me-excel": "^0.8.1"
  }
}
```

## Operations

### Add Text Header Footer To Excel

Add customizable text headers and footers to Excel worksheets with full control over alignment, font size, and color.

**Features:**
- **Worksheet Selection**: Target specific worksheets or apply to all worksheets
- **Header Text**: Add customizable header text with alignment options (left, center, right)
- **Footer Text**: Add customizable footer text with page numbering support (&P for current page, &N for total pages)
- **Font Styling**: Configure font size (6-72pt) and color (hex format)
- **Flexible Input**: Support for binary data and base64 encoded Excel files

**Use Cases:**
- Add company branding to automated reports
- Include page numbers on all worksheets
- Add confidential markings to sensitive documents
- Customize financial reports with headers and footers
- Brand template documents before distribution

## Credentials

To use this node, you need a PDF4ME API key:

1. Sign up for a free account at [PDF4ME](https://portal.pdf4me.com/register)
2. Navigate to [API Keys](https://portal.pdf4me.com/api-keys) in your account
3. Generate a new API key
4. Add the API key to your n8n credentials:
   - In n8n, go to **Credentials > New**
   - Select **PDF4me API**
   - Enter your API key
   - Save the credentials

## Usage

### Basic Example: Add Header and Footer to Excel

This example shows how to add a header and footer to an Excel file:

1. **Input Node**: Use a node that provides an Excel file (e.g., HTTP Request, Google Drive, etc.)
2. **PDF4me Excel Node**: Configure with:
   - Input Data Type: Binary Data
   - Worksheet Name: Sheet1
   - Header Text: Company Confidential
   - Footer Text: Page &P of &N
   - Header/Footer Alignment: Center
   - Font Size: 10
   - Font Color: #000000
   - Apply To All Worksheets: false
3. **Output**: The modified Excel file with headers and footers

### Advanced Example: Process Multiple Files

Process multiple Excel files with custom headers:

1. **Loop Node**: Iterate over multiple files
2. **PDF4me Excel Node**: Add headers/footers to each file
3. **Save/Send**: Save the processed files or send them via email

## Input Options

### Input Data Type
- **Binary Data**: Use Excel files from previous nodes (most common)
- **Base64 String**: Provide Excel content as base64 encoded string

### Header/Footer Configuration
- **Worksheet Name**: The worksheet to apply headers/footers to (default: "Sheet1")
- **Header Text**: Text to display in the header
- **Footer Text**: Text to display in the footer (supports &P for page number, &N for total pages)
- **Header Alignment**: left, center, or right
- **Footer Alignment**: left, center, or right
- **Font Size**: 6-72 points
- **Font Color**: Hex color code (e.g., #000000 for black)
- **Apply To All Worksheets**: Apply to all worksheets in the workbook (true/false)

### Output Options
- **Output File Name**: Name for the processed Excel file
- **Binary Data Output Name**: Custom name for the binary data in n8n output

## Resources

- [PDF4ME API Documentation](https://dev.pdf4me.com/apiv2/documentation/)
- [PDF4ME Portal](https://portal.pdf4me.com/)
- [n8n Documentation](https://docs.n8n.io/)
- [n8n Community](https://community.n8n.io/)

## Compatibility

- Minimum n8n version: 0.187.0
- Supported Excel formats: .xlsx
- API: PDF4ME API v2

## Support

For issues and feature requests:
- GitHub Issues: [n8n-nodes-pdf4me-excel](https://github.com/pdf4me/n8n-nodes-pdf4me-excel/issues)
- PDF4ME Support: support@pdf4me.com
- n8n Community: [community.n8n.io](https://community.n8n.io/)

## Version History

### 1.0.0 (Current)
- Initial release
- Add Text Header Footer To Excel operation
- Support for custom font styling
- Worksheet-specific or workbook-wide application
- Binary data and base64 input support

## License

[MIT License](LICENSE.md)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Keywords

n8n, n8n-community-node-package, excel, spreadsheet, headers, footers, pdf4me, office, documents, automation

## Author

PDF4me - https://pdf4me.com

## Acknowledgments

Built with [n8n](https://n8n.io/) and powered by [PDF4ME API](https://pdf4me.com/)
