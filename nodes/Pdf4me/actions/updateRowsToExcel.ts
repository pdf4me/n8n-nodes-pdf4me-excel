import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import { NodeOperationError, NodeApiError } from 'n8n-workflow';
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
				operation: [ActionConstants.UpdateRowsToExcel],
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
				operation: [ActionConstants.UpdateRowsToExcel],
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
				operation: [ActionConstants.UpdateRowsToExcel],
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
				operation: [ActionConstants.UpdateRowsToExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === UPDATE SETTINGS ===
	{
		displayName: 'Worksheet Name',
		name: 'worksheetName',
		type: 'string',
		required: true,
		default: 'Sheet1',
		description: 'Name of the worksheet to update',
		placeholder: 'Sheet1',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'JSON Data',
		name: 'jsonData',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
			rows: 10,
		},
		required: true,
		default: '[{"Product":"Widget","Price":19.99,"Stock":100}]',
		description: 'JSON data containing the row values to update. Can be a JSON array of objects or a JSON string.',
		placeholder: '[{"Product":"Widget","Price":19.99,"Stock":100}]',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Start Row',
		name: 'startRow',
		type: 'number',
		default: 1,
		required: true,
		description: 'Starting row number for updates (1-based indexing)',
		typeOptions: {
			minValue: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Start Column',
		name: 'startColumn',
		type: 'number',
		default: 1,
		required: true,
		description: 'Starting column number for updates (1-based indexing)',
		typeOptions: {
			minValue: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	// === ADVANCED SETTINGS ===
	{
		displayName: 'Convert Numeric and Date Values',
		name: 'convertNumericAndDate',
		type: 'boolean',
		default: true,
		description: 'Whether to automatically convert numeric and date values to their proper types',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Date Format',
		name: 'dateFormat',
		type: 'string',
		default: 'yyyy-MM-dd',
		description: 'Format pattern for date values (e.g., yyyy-MM-dd, MM/dd/yyyy, dd-MMM-yyyy)',
		placeholder: 'yyyy-MM-dd',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Numeric Format',
		name: 'numericFormat',
		type: 'string',
		default: 'N2',
		description: 'Format pattern for numeric values (e.g., N2 for 2 decimals, N0 for no decimals, C for currency)',
		placeholder: 'N2',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Ignore Null Values',
		name: 'ignoreNullValues',
		type: 'boolean',
		default: false,
		description: 'Whether to skip null values when updating (keeps existing cell values)',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Ignore Attribute Titles',
		name: 'ignoreAttributeTitles',
		type: 'boolean',
		default: false,
		description: 'Whether to skip JSON property names when updating (only updates values)',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Culture Name',
		name: 'cultureName',
		type: 'string',
		default: 'en-US',
		description: 'Culture name for data formatting (e.g., en-US, de-DE, fr-FR)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_updated.xlsx',
		description: 'Name for the processed Excel file (will have rows updated)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
	{
		displayName: 'Source Document Name',
		name: 'docName',
		type: 'string',
		default: 'myExcelFile.xlsx',
		description: 'Name of the original Excel file (for reference and processing)',
		placeholder: 'myExcelFile.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateRowsToExcel],
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
				operation: [ActionConstants.UpdateRowsToExcel],
			},
		},
	},
];

/**
 * Update rows in Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request with JSON data → Poll for completion → Save Excel file
 * Updates specified rows in a worksheet with JSON data starting from a specific position
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get update parameters
		const worksheetName = this.getNodeParameter('worksheetName', index) as string;
		const jsonDataInput = this.getNodeParameter('jsonData', index) as string;
		const startRow = this.getNodeParameter('startRow', index, 1) as number;
		const startColumn = this.getNodeParameter('startColumn', index, 1) as number;
		const convertNumericAndDate = this.getNodeParameter('convertNumericAndDate', index, true) as boolean;
		const dateFormat = this.getNodeParameter('dateFormat', index, 'yyyy-MM-dd') as string;
		const numericFormat = this.getNodeParameter('numericFormat', index, 'N2') as string;
		const ignoreNullValues = this.getNodeParameter('ignoreNullValues', index, false) as boolean;
		const ignoreAttributeTitles = this.getNodeParameter('ignoreAttributeTitles', index, false) as boolean;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		let docContent: string;
		let originalFileName = docName;

		// Handle different input data types
		if (inputDataType === 'binaryData') {
			// Get Excel content from binary data
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index) as string;
			const item = this.getInputData(index);

			if (!item[0].binary || !item[0].binary[binaryPropertyName]) {
				throw new NodeOperationError(this.getNode(), `No binary data found in property '${binaryPropertyName}'`);
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
				throw new NodeOperationError(this.getNode(), 'URL is required when using URL input type');
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
				throw new NodeOperationError(this.getNode(), `Failed to download file from URL: ${errorMessage}`);
			}
		} else {
			throw new NodeOperationError(this.getNode(), `Unsupported input data type: ${inputDataType}`);
		}

		// Validate content
		if (!docContent || docContent.trim() === '') {
			throw new NodeOperationError(this.getNode(), 'Excel content is required');
		}

		// Validate and parse JSON data
		let jsonData: string;
		if (!jsonDataInput || jsonDataInput.trim() === '') {
			throw new NodeOperationError(this.getNode(), 'JSON data is required');
		}

		// Check if the input is already a valid JSON string or needs to be stringified
		try {
			// Try to parse to validate it's valid JSON
			JSON.parse(jsonDataInput);
			// If successful, use it as-is
			jsonData = jsonDataInput;
		} catch (parseError) {
			// If parsing fails, it might be a JavaScript object that needs stringifying
			throw new NodeOperationError(this.getNode(), `Invalid JSON data: ${parseError instanceof Error ? parseError.message : 'Unknown error'}`);
		}

		// Validate worksheet name
		if (!worksheetName || worksheetName.trim() === '') {
			throw new NodeOperationError(this.getNode(), 'Worksheet name is required');
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			UpdateRowsToExcelAction: {
				WorksheetName: worksheetName,
				JsonData: jsonData,
				StartRow: startRow,
				StartColumn: startColumn,
				ConvertNumericAndDate: convertNumericAndDate,
				DateFormat: dateFormat,
				NumericFormat: numericFormat,
				IgnoreNullValues: ignoreNullValues,
				IgnoreAttributeTitles: ignoreAttributeTitles,
				CultureName: cultureName,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelUpdateRows',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_updated';
				fileName = `${baseName}.xlsx`;
			}

			// Ensure .xlsx extension
			if (!fileName.toLowerCase().endsWith('.xlsx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.xlsx`;
			}

			// Handle the response - Excel API returns JSON with embedded base64 file
			let excelBuffer: Buffer;

			// The API returns JSON in format: { document: { docData: "base64..." }, ... }
			// or { docData: "base64..." } or similar structures
			// Check for Buffer first to properly narrow TypeScript types
			if (Buffer.isBuffer(responseData)) {
				// Direct binary response
				excelBuffer = responseData;
			} else if (typeof responseData === 'string') {
				// Base64 string response
				excelBuffer = Buffer.from(responseData, 'base64');
			} else if (typeof responseData === 'object' && responseData !== null) {
				// Try different possible response structures from IDataObject
				const response = responseData as IDataObject;

				// Check if the response has a document field
				if (response.document) {
					const document = response.document;

					// The document could be a string (base64) or an object with nested fields
					if (typeof document === 'string') {
						// Document itself is the base64 content
						excelBuffer = Buffer.from(document, 'base64');
					} else if (typeof document === 'object' && document !== null) {
						// Document is an object, extract base64 from possible fields
						const docObj = document as IDataObject;
						const docContent =
							(docObj.docData as string) ||
							(docObj.content as string) ||
							(docObj.docContent as string) ||
							(docObj.data as string) ||
							(docObj.file as string);

						if (!docContent) {
							const docKeys = Object.keys(docObj).join(', ');
							throw new Error(`Document object has unexpected structure. Available keys: ${docKeys}`);
						}

						excelBuffer = Buffer.from(docContent, 'base64');
					} else {
						throw new Error(`Document field is neither string nor object: ${typeof document}`);
					}
				} else {
					// No document field, try other possible locations
					const docContent =
						(response.docData as string) ||
						(response.content as string) ||
						(response.fileContent as string) ||
						(response.data as string);

					if (!docContent) {
						// If no known field found, log the structure for debugging
						const keys = Object.keys(responseData).join(', ');
						throw new Error(`Excel API returned unexpected JSON structure. Available keys: ${keys}`);
					}

					excelBuffer = Buffer.from(docContent, 'base64');
				}
			} else {
				throw new Error(`Unexpected response format: ${typeof responseData}`);
			}

			// Validate the response contains Excel data
			if (!excelBuffer || excelBuffer.length < 1000) {
				throw new Error(
					'Invalid Excel response from API. The file appears to be too small or corrupted.',
				);
			}

			// Validate Excel file format (XLSX files start with PK signature - ZIP format)
			const magicBytes = excelBuffer.toString('hex', 0, 4);
			if (magicBytes !== '504b0304') {
				throw new Error(
					`Invalid Excel file format. Expected XLSX file but got unexpected data. Magic bytes: ${magicBytes}`,
				);
			}

			// Create binary data for output
			const binaryData = await this.helpers.prepareBinaryData(
				excelBuffer,
				fileName,
				'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			);

			// Determine the binary data name
			const binaryDataKey = binaryDataName || 'data';

			// Parse JSON data for output summary
			let rowCount = 0;
			try {
				const parsedData = JSON.parse(jsonData);
				rowCount = Array.isArray(parsedData) ? parsedData.length : 1;
			} catch {
				rowCount = 1;
			}

			return [
				{
					json: {
						fileName,
						fileSize: excelBuffer.length,
						success: true,
						originalFileName,
						worksheetName,
						startRow,
						startColumn,
						rowsUpdated: rowCount,
						convertNumericAndDate,
						dateFormat,
						numericFormat,
						ignoreNullValues,
						ignoreAttributeTitles,
						cultureName,
						message: `Successfully updated ${rowCount} row(s) in worksheet '${worksheetName}'`,
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				},
			];
		}

		throw new NodeOperationError(this.getNode(), 'No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		// If it's already a NodeOperationError or NodeApiError, re-throw as-is
		if (error instanceof NodeOperationError || error instanceof NodeApiError) {
			throw error;
		}
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new NodeOperationError(this.getNode(), `Update rows in Excel failed: ${errorMessage}`);
	}
}


