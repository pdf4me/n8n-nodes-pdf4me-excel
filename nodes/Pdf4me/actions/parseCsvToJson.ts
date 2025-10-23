import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';


export const description: INodeProperties[] = [
	// === INPUT FILE SETTINGS ===
	{
		displayName: 'CSV File Input Method',
		name: 'inputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the CSV file for processing',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
		options: [
			{
				name: 'From Previous Node (Binary Data)',
				value: 'binaryData',
				description: 'Use CSV file passed from a previous n8n node',
			},
			{
				name: 'Base64 Encoded String',
				value: 'base64',
				description: 'Provide CSV file content as base64 encoded string',
			},
			{
				name: 'Download from URL',
				value: 'url',
				description: 'Download CSV file directly from a web URL',
			},
		],
	},
	{
		displayName: 'Binary Data Property Name',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the CSV file (usually \'data\')',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
				inputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Base64 Encoded CSV Content',
		name: 'base64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the CSV file data',
		placeholder: 'UEsDBBQABgAIAAAAIQBi7p1oXgEAAJAEAAATAAgCW0NvbnRlbnRfVHlwZXNdLnht...',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
				inputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'CSV File URL',
		name: 'url',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the CSV file from (must be publicly accessible)',
		placeholder: 'https://example.com/data.csv',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === CSV PARSING SETTINGS ===
	{
		displayName: 'CSV Delimiter',
		name: 'delimiter',
		type: 'options',
		default: ',',
		description: 'Character used to separate values in the CSV file',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
		options: [
			{
				name: 'Comma (,)',
				value: ',',
				description: 'Standard CSV delimiter',
			},
			{
				name: 'Semicolon (;)',
				value: ';',
				description: 'Common in European locales',
			},
			{
				name: 'Tab (\\t)',
				value: '\t',
				description: 'Tab-separated values (TSV)',
			},
			{
				name: 'Pipe (|)',
				value: '|',
				description: 'Pipe-delimited files',
			},
			{
				name: 'Custom',
				value: 'custom',
				description: 'Specify a custom delimiter',
			},
		],
	},
	{
		displayName: 'Custom Delimiter',
		name: 'customDelimiter',
		type: 'string',
		default: '',
		description: 'Custom character to use as delimiter',
		placeholder: ':',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
				delimiter: ['custom'],
			},
		},
	},
	{
		displayName: 'Skip First Line',
		name: 'skipFirstLine',
		type: 'boolean',
		default: false,
		description: 'Whether to skip the first line when parsing (useful if first line is not a header or contains metadata)',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
	{
		displayName: 'Custom Column Headers',
		name: 'columnHeaders',
		type: 'string',
		default: '',
		description: 'Comma-separated list of custom column headers to use instead of the ones in the file. Leave empty to use headers from CSV.',
		placeholder: 'Name,Email,Phone,Address',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
	{
		displayName: 'Culture Name',
		name: 'cultureName',
		type: 'string',
		default: 'en-US',
		description: 'Culture name for data parsing (affects number and date formats)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'parsed_data.json',
		description: 'Name for the output JSON file',
		placeholder: 'output.json',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
	{
		displayName: 'Source Document Name',
		name: 'docName',
		type: 'string',
		default: 'data.csv',
		description: 'Name of the original CSV file (for reference and processing)',
		placeholder: 'data.csv',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
	{
		displayName: 'Output Binary Data Name',
		name: 'binaryDataName',
		type: 'string',
		default: 'data',
		description: 'Name for the binary data in the n8n output (used to access the processed file)',
		placeholder: 'excel-file',
		displayOptions: {
			show: {
				operation: [ActionConstants.ParseCsvToExcel],
			},
		},
	},
];

/**
 * Parse CSV file and convert to JSON using PDF4Me API
 * Process: Read CSV file → Encode to base64 → Send API request → Poll for completion → Save JSON file
 * Converts CSV files to JSON format with customizable parsing options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get CSV parsing parameters
		const delimiterOption = this.getNodeParameter('delimiter', index, ',') as string;
		const customDelimiter = this.getNodeParameter('customDelimiter', index, '') as string;
		const skipFirstLine = this.getNodeParameter('skipFirstLine', index, false) as boolean;
		const columnHeaders = this.getNodeParameter('columnHeaders', index, '') as string;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		// Determine actual delimiter to use
		let delimiter = delimiterOption;
		if (delimiterOption === 'custom') {
			if (!customDelimiter) {
				throw new Error('Custom delimiter is required when "Custom" is selected');
			}
			delimiter = customDelimiter;
		}

		let docContent: string;
		let originalFileName = docName;

		// Handle different input data types
		if (inputDataType === 'binaryData') {
			// Get CSV content from binary data
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
			// Download CSV file from URL
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
			throw new Error('CSV content is required');
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			CsvParseToExcelAction: {
				Delimiter: delimiter,
				ColumnHeaders: columnHeaders,
				SkipFirstLine: skipFirstLine,
				CultureName: cultureName,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelParseCsv',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'parsed_data';
				fileName = `${baseName}.txt`;
			}

			// Ensure .txt extension
			if (!fileName.toLowerCase().endsWith('.txt')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.txt`;
			}

			// Handle the response - API returns base64 encoded text content
			let decodedContent: string;

			if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;

				// Check for errors first
				if (response.Success === false) {
					const errorMessage = response.ErrorMessage as string || 'Unknown error occurred';
					const errors = response.Errors as string[] || [];
					const errorDetails = errors.length > 0 ? ` Details: ${errors.join(', ')}` : '';
					throw new Error(`CSV parsing failed: ${errorMessage}${errorDetails}`);
				}

				// Extract base64 content from response
				const base64Content =
					(response.document as string) ||
					(response.docData as string) ||
					(response.content as string) ||
					(response.data as string) ||
					(response.fileContent as string);

				if (!base64Content) {
					const keys = Object.keys(responseData).join(', ');
					throw new Error(`CSV API returned unexpected structure. Available keys: ${keys}`);
				}

				// Decode base64 to get text content
				try {
					const buffer = Buffer.from(base64Content, 'base64');
					decodedContent = buffer.toString('utf8');
				} catch (error) {
					// If base64 decoding fails, use the string directly
					decodedContent = base64Content;
				}
			} else if (typeof responseData === 'string') {
				// Direct string response - try to decode as base64
				try {
					const buffer = Buffer.from(responseData, 'base64');
					decodedContent = buffer.toString('utf8');
				} catch (error) {
					// If base64 decoding fails, use the string directly
					decodedContent = responseData;
				}
			} else {
				throw new Error(`Unexpected response format: ${typeof responseData}`);
			}

			// Create TXT file content
			const txtBuffer = Buffer.from(decodedContent, 'utf8');

			// Create JSON file content (same content, different extension)
			const jsonBuffer = Buffer.from(decodedContent, 'utf8');

			// Create TXT file
			const txtFileName = fileName.replace(/\.[^.]*$/, '.txt');
			const txtBinaryData = await this.helpers.prepareBinaryData(
				txtBuffer,
				txtFileName,
				'text/plain',
			);

			// Create JSON file (same content, different extension)
			const jsonFileName = fileName.replace(/\.[^.]*$/, '.json');
			const jsonBinaryData = await this.helpers.prepareBinaryData(
				jsonBuffer,
				jsonFileName,
				'application/json',
			);

			// Determine the binary data names
			const txtBinaryDataKey = `${binaryDataName || 'data'}_txt`;
			const jsonBinaryDataKey = `${binaryDataName || 'data'}_json`;

			return [
				{
					json: {
						txtFileName,
						jsonFileName,
						txtFileSize: txtBuffer.length,
						jsonFileSize: jsonBuffer.length,
						success: true,
						originalFileName,
						delimiter,
						skipFirstLine,
						columnHeaders: columnHeaders || 'from CSV',
						cultureName,
						message: 'Successfully converted CSV and created TXT and JSON outputs',
					},
					binary: {
						[txtBinaryDataKey]: txtBinaryData,
						[jsonBinaryDataKey]: jsonBinaryData,
					},
				},
			];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Parse CSV to Excel failed: ${errorMessage}`);
	}
}


