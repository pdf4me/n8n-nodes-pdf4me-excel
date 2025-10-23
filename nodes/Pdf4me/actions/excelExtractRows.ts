import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';


export const description: INodeProperties[] = [
	// === INPUT FILE SETTINGS ===
	{
		displayName: 'Excel File Input Method',
		name: 'inputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the Excel file for processing',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
		options: [
			{
				name: 'From Previous Node (Binary Data)',
				value: 'binaryData',
				description: 'Use Excel file passed from a previous n8n node',
			},
			{
				name: 'Base64 Encoded String',
				value: 'base64',
				description: 'Provide Excel file content as base64 encoded string',
			},
			{
				name: 'Download from URL',
				value: 'url',
				description: 'Download Excel file directly from a web URL',
			},
		],
	},
	{
		displayName: 'Binary Data Property Name',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the Excel file (usually \'data\')',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
				inputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Base64 Encoded Excel Content',
		name: 'base64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the Excel file data',
		placeholder: 'UEsDBBQABgAIAAAAIQBi7p1oXgEAAJAEAAATAAgCW0NvbnRlbnRfVHlwZXNdLnht...',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
				inputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'Excel File URL',
		name: 'url',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the Excel file from (must be publicly accessible)',
		placeholder: 'https://example.com/file.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
				inputDataType: ['url'],
			},
		},
	},
	// === EXTRACTION SETTINGS ===
	{
		displayName: 'Worksheet Name',
		name: 'worksheetName',
		type: 'string',
		required: false,
		default: '',
		description: 'Name of the worksheet to extract from (leave empty for first worksheet)',
		placeholder: 'Sheet1',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Has Header Row',
		name: 'hasHeaderRow',
		type: 'boolean',
		default: true,
		description: 'Whether the first row contains column headers',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	// === ROW RANGE SETTINGS ===
	{
		displayName: 'First Row (Start)',
		name: 'firstRow',
		type: 'number',
		default: 0,
		description: 'Starting row for extraction (0-based, 0 = first row)',
		typeOptions: {
			minValue: 0,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Last Row (End)',
		name: 'lastRow',
		type: 'number',
		default: -1,
		description: 'Ending row for extraction (-1 for last row with data)',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'First Column (Start)',
		name: 'firstColumn',
		type: 'number',
		default: 0,
		description: 'Starting column for extraction (0-based, 0 = first column)',
		typeOptions: {
			minValue: 0,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Last Column (End)',
		name: 'lastColumn',
		type: 'number',
		default: -1,
		description: 'Ending column for extraction (-1 for last column with data)',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	// === ADVANCED SETTINGS ===
	{
		displayName: 'Exclude Empty Rows',
		name: 'excludeEmptyRows',
		type: 'boolean',
		default: true,
		description: 'Skip rows that are completely empty',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Exclude Hidden Rows',
		name: 'excludeHiddenRows',
		type: 'boolean',
		default: true,
		description: 'Skip hidden rows in extraction',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Exclude Hidden Columns',
		name: 'excludeHiddenColumns',
		type: 'boolean',
		default: true,
		description: 'Skip hidden columns in extraction',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Export Values As Text',
		name: 'exportValuesAsText',
		type: 'boolean',
		default: false,
		description: 'Export all cell values as text strings',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Export Empty Cells',
		name: 'exportEmptyCells',
		type: 'boolean',
		default: false,
		description: 'Include empty cells in the output',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Export as Object Structure',
		name: 'exportAsObject',
		type: 'boolean',
		default: false,
		description: 'Export data as structured object instead of array',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Hyperlink Format',
		name: 'hyperlinkFormat',
		type: 'options',
		default: 'Text',
		description: 'How to handle hyperlinks in cells',
		options: [
			{
				name: 'Text',
				value: 'Text',
				description: 'Show only the display text of hyperlinks',
			},
			{
				name: 'URL',
				value: 'Url',
				description: 'Show only the URL of hyperlinks',
			},
			{
				name: 'Both',
				value: 'Both',
				description: 'Show both text and URL of hyperlinks',
			},
		],
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	{
		displayName: 'Culture & Language Settings',
		name: 'culture',
		type: 'string',
		default: 'en-US',
		description: 'Culture name for data formatting (e.g., en-US, de-DE, fr-FR)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Source Document Name',
		name: 'docName',
		type: 'string',
		default: 'myExcelFile.xlsx',
		description: 'Name of the original Excel file (for reference and processing)',
		placeholder: 'myExcelFile.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExcelExtractRows],
			},
		},
	},
];

/**
 * Extract rows from Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Return JSON data
 * Extracts specified rows from a worksheet and returns them as JSON data
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;

		// Get extraction parameters
		const worksheetName = this.getNodeParameter('worksheetName', index, '') as string;
		const hasHeaderRow = this.getNodeParameter('hasHeaderRow', index, true) as boolean;
		const firstRow = this.getNodeParameter('firstRow', index, 0) as number;
		const lastRow = this.getNodeParameter('lastRow', index, -1) as number;
		const firstColumn = this.getNodeParameter('firstColumn', index, 0) as number;
		const lastColumn = this.getNodeParameter('lastColumn', index, -1) as number;
		const excludeEmptyRows = this.getNodeParameter('excludeEmptyRows', index, true) as boolean;
		const excludeHiddenRows = this.getNodeParameter('excludeHiddenRows', index, true) as boolean;
		const excludeHiddenColumns = this.getNodeParameter('excludeHiddenColumns', index, true) as boolean;
		const exportValuesAsText = this.getNodeParameter('exportValuesAsText', index, false) as boolean;
		const exportEmptyCells = this.getNodeParameter('exportEmptyCells', index, false) as boolean;
		const exportAsObject = this.getNodeParameter('exportAsObject', index, false) as boolean;
		const hyperlinkFormat = this.getNodeParameter('hyperlinkFormat', index, 'Text') as string;
		const culture = this.getNodeParameter('culture', index, 'en-US') as string;

		let docContent: string;
		let originalFileName = docName;

		// Handle different input data types
		if (inputDataType === 'binaryData') {
			// Get Excel content from binary data
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index) as string;
			const item = this.getInputData(index);

			if (!item[0].binary || !item[0].binary[binaryPropertyName]) {
				throw new Error(`No binary data found in property '${binaryPropertyName}'`);
			}

			const binaryData = item[0].binary[binaryPropertyName];
			const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
			docContent = buffer.toString('base64');

			if (binaryData.fileName) {
				originalFileName = binaryData.fileName;
			}
		} else if (inputDataType === 'base64') {
			// Use base64 content directly
			docContent = this.getNodeParameter('base64Content', index) as string;

			// Remove data URL prefix if present
			if (docContent.includes(',')) {
				docContent = docContent.split(',')[1];
			}
		} else if (inputDataType === 'url') {
			// Download Excel file from URL
			const url = this.getNodeParameter('url', index) as string;

			if (!url || url.trim() === '') {
				throw new Error('URL is required when using URL input type');
			}

			try {
				// Download the file using n8n's helpers
				const response = await this.helpers.httpRequest({
					method: 'GET',
					url,
					encoding: 'arraybuffer',
					returnFullResponse: true,
				});

				// Convert to base64
				const buffer = Buffer.from(response.body as ArrayBuffer);
				docContent = buffer.toString('base64');

				// Try to extract filename from URL or Content-Disposition header
				const contentDisposition = response.headers['content-disposition'];
				if (contentDisposition) {
					const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
					if (filenameMatch && filenameMatch[1]) {
						originalFileName = filenameMatch[1].replace(/['"]/g, '');
					}
				}

				// Fallback: extract filename from URL
				if (originalFileName === docName) {
					const urlParts = url.split('/');
					const urlFilename = urlParts[urlParts.length - 1].split('?')[0];
					if (urlFilename) {
						originalFileName = decodeURIComponent(urlFilename);
					}
				}
			} catch (error) {
				const errorMessage = error instanceof Error ? error.message : 'Unknown error';
				throw new Error(`Failed to download file from URL: ${errorMessage}`);
			}
		} else {
			throw new Error(`Unsupported input data type: ${inputDataType}`);
		}

		// Validate content
		if (!docContent || docContent.trim() === '') {
			throw new Error('Excel content is required');
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			extractRowsToExcelAction: {
				WorksheetName: worksheetName,
				HasHeaderRow: hasHeaderRow,
				FirstRow: firstRow,
				LastRow: lastRow,
				FirstColumn: firstColumn,
				LastColumn: lastColumn,
				ExcludeEmptyRows: excludeEmptyRows,
				ExcludeHiddenRows: excludeHiddenRows,
				ExcludeHiddenColumns: excludeHiddenColumns,
				ExportValuesAsText: exportValuesAsText,
				ExportEmptyCells: exportEmptyCells,
				ExportAsObject: exportAsObject,
				HyperlinkFormat: hyperlinkFormat,
				Culture: culture,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelExtractRows',
			body,
		);

		if (responseData) {
			// Handle the response - Excel Extract Rows API returns JSON data
			let extractedData: unknown;

			if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;

				// The API returns JSON in format: { data: [...], ... } or similar structures
				// Check for different possible response structures
				extractedData =
					response.data ||
					response.rows ||
					response.extractedData ||
					response.result ||
					response;

				// If no specific data field found, use the entire response
				if (!extractedData) {
					extractedData = response;
				}
			} else if (typeof responseData === 'string') {
				// Try to parse as JSON if it's a string
				try {
					extractedData = JSON.parse(responseData);
				} catch {
					extractedData = responseData;
				}
			} else {
				extractedData = responseData;
			}

			// Count rows for summary
			let rowCount = 0;
			if (Array.isArray(extractedData)) {
				rowCount = extractedData.length;
			} else if (typeof extractedData === 'object' && extractedData !== null) {
				// If it's an object, try to count rows in common structures
				const dataObj = extractedData as IDataObject;
				if (dataObj.rows && Array.isArray(dataObj.rows)) {
					rowCount = dataObj.rows.length;
				} else if (dataObj.data && Array.isArray(dataObj.data)) {
					rowCount = dataObj.data.length;
				} else {
					rowCount = 1; // Single object
				}
			}

			return [
				{
					json: {
						success: true,
						originalFileName,
						worksheetName: worksheetName || 'First Worksheet',
						hasHeaderRow,
						firstRow,
						lastRow,
						firstColumn,
						lastColumn,
						excludeEmptyRows,
						excludeHiddenRows,
						excludeHiddenColumns,
						exportValuesAsText,
						exportEmptyCells,
						exportAsObject,
						hyperlinkFormat,
						culture,
						rowsExtracted: rowCount,
						extractedData: extractedData as IDataObject,
						message: `Successfully extracted ${rowCount} row(s) from worksheet '${worksheetName || 'First Worksheet'}'`,
					},
				},
			];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Extract rows from Excel failed: ${errorMessage}`);
	}
}
