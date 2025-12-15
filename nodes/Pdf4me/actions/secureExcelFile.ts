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
				operation: [ActionConstants.SecureExcelFile],
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
				operation: [ActionConstants.SecureExcelFile],
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
				operation: [ActionConstants.SecureExcelFile],
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
				operation: [ActionConstants.SecureExcelFile],
				inputDataType: ['url'],
			},
		},
	},
	// === FILE PROTECTION ===
	{
		displayName: 'File Protection Password',
		name: 'password',
		type: 'string',
		typeOptions: {
			password: true,
		},
		default: '',
		description: 'Password to encrypt and protect the entire Excel file. Leave empty for no file-level encryption.',
		placeholder: 'Enter password',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
			},
		},
	},
	// === WORKBOOK PROTECTION ===
	{
		displayName: 'Protect Workbook Structure',
		name: 'protectWorkbook',
		type: 'boolean',
		default: true,
		description: 'Whether to protect workbook structure (prevents adding/deleting/renaming sheets)',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
			},
		},
	},
	{
		displayName: 'Workbook Protection Password',
		name: 'protectWorkbookPassword',
		type: 'string',
		typeOptions: {
			password: true,
		},
		default: '',
		description: 'Separate password for workbook structure protection. Leave empty to use no password.',
		placeholder: 'Enter workbook password',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorkbook: [true],
			},
		},
	},
	// === WORKSHEET PROTECTION ===
	{
		displayName: 'Protect Worksheets',
		name: 'protectWorksheets',
		type: 'boolean',
		default: true,
		description: 'Whether to protect worksheet content (prevents editing cells, formatting, etc.)',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
			},
		},
	},
	{
		displayName: 'Worksheet Protection Type',
		name: 'worksheetProtectionType',
		type: 'options',
		default: 'All',
		description: 'Type of worksheet protection to apply',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorksheets: [true],
			},
		},
		options: [
			{
				name: 'All',
				value: 'All',
				description: 'Protect all aspects (contents, objects, and scenarios)',
			},
			{
				name: 'Contents',
				value: 'Contents',
				description: 'Protect cell contents only',
			},
			{
				name: 'Objects',
				value: 'Objects',
				description: 'Protect embedded objects only',
			},
			{
				name: 'Scenarios',
				value: 'Scenarios',
				description: 'Protect scenarios only',
			},
		],
	},
	{
		displayName: 'Worksheet Protection Password',
		name: 'worksheetProtectionPassword',
		type: 'string',
		typeOptions: {
			password: true,
		},
		default: '',
		description: 'Separate password for worksheet protection. Leave empty to use no password.',
		placeholder: 'Enter worksheet password',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorksheets: [true],
			},
		},
	},
	{
		displayName: 'Worksheet Selection',
		name: 'worksheetSelection',
		type: 'options',
		default: 'all',
		description: 'Which worksheets to protect',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorksheets: [true],
			},
		},
		options: [
			{
				name: 'All Worksheets',
				value: 'all',
				description: 'Protect all worksheets in the workbook',
			},
			{
				name: 'Specific Worksheets by Name',
				value: 'name',
				description: 'Protect specific worksheets by their names',
			},
			{
				name: 'Specific Worksheets by Index',
				value: 'index',
				description: 'Protect specific worksheets by their index positions',
			},
		],
	},
	{
		displayName: 'Worksheet Names',
		name: 'worksheetNames',
		type: 'string',
		default: '',
		description: 'Comma-separated list of worksheet names to protect (e.g., "Sheet1, Sheet2, Data")',
		placeholder: 'Sheet1, Sheet2',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorksheets: [true],
				worksheetSelection: ['name'],
			},
		},
	},
	{
		displayName: 'Worksheet Indexes',
		name: 'worksheetIndexes',
		type: 'string',
		default: '',
		description: 'Comma-separated list of worksheet indexes to protect (1-based, e.g., "1, 2, 3")',
		placeholder: '1, 2, 3',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
				protectWorksheets: [true],
				worksheetSelection: ['index'],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_secured.xlsx',
		description: 'Name for the processed Excel file (will be password protected)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.SecureExcelFile],
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
				operation: [ActionConstants.SecureExcelFile],
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
				operation: [ActionConstants.SecureExcelFile],
			},
		},
	},
];

/**
 * Secure Excel files with password protection using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save secured Excel file
 * Supports file-level encryption, workbook structure protection, and worksheet content protection
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get security parameters
		const password = this.getNodeParameter('password', index, '') as string;
		const protectWorkbook = this.getNodeParameter('protectWorkbook', index, true) as boolean;
		const protectWorkbookPassword = this.getNodeParameter('protectWorkbookPassword', index, '') as string;
		const protectWorksheets = this.getNodeParameter('protectWorksheets', index, true) as boolean;
		const worksheetProtectionType = this.getNodeParameter('worksheetProtectionType', index, 'All') as string;
		const worksheetProtectionPassword = this.getNodeParameter('worksheetProtectionPassword', index, '') as string;
		const worksheetSelection = this.getNodeParameter('worksheetSelection', index, 'all') as string;
		const worksheetNames = this.getNodeParameter('worksheetNames', index, '') as string;
		const worksheetIndexes = this.getNodeParameter('worksheetIndexes', index, '') as string;

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

		// Validate that at least one protection is enabled
		if (!password && !protectWorkbook && !protectWorksheets) {
			throw new NodeOperationError(this.getNode(), 'At least one protection option must be enabled (file password, workbook protection, or worksheet protection)');
		}

		// Determine worksheet names and indexes based on selection
		let finalWorksheetNames = '';
		let finalWorksheetIndexes = '';

		if (protectWorksheets) {
			if (worksheetSelection === 'name') {
				if (!worksheetNames || worksheetNames.trim() === '') {
					throw new NodeOperationError(this.getNode(), 'Worksheet names are required when protecting specific worksheets by name');
				}
				finalWorksheetNames = worksheetNames;
			} else if (worksheetSelection === 'index') {
				if (!worksheetIndexes || worksheetIndexes.trim() === '') {
					throw new NodeOperationError(this.getNode(), 'Worksheet indexes are required when protecting specific worksheets by index');
				}
				finalWorksheetIndexes = worksheetIndexes;
			}
			// If worksheetSelection === 'all', leave both empty
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			SecureExcelAction: {
				Password: password,
				ProtectWorkbook: protectWorkbook,
				ProtectWorkbookPassword: protectWorkbookPassword,
				ProtectWorksheets: protectWorksheets,
				WorksheetProtectionType: worksheetProtectionType,
				WorksheetProtectionPassword: worksheetProtectionPassword,
				WorksheetNames: finalWorksheetNames,
				WorksheetIndexes: finalWorksheetIndexes,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelSecure',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_secured';
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

			// Validate Excel file format (support both .xls and .xlsx formats)
			const magicBytes = excelBuffer.toString('hex', 0, 4);
			const isXlsx = magicBytes === '504b0304'; // ZIP signature for .xlsx files
			const isXls = magicBytes === 'd0cf11e0'; // OLE signature for .xls files

			if (!isXlsx && !isXls) {
				throw new Error(
					`Invalid Excel file format. Expected .xls or .xlsx file but got unexpected data. Magic bytes: ${magicBytes}`,
				);
			}

			// Create binary data for output
			const mimeType = isXlsx
				? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' // .xlsx
				: 'application/vnd.ms-excel'; // .xls

			const binaryData = await this.helpers.prepareBinaryData(
				excelBuffer,
				fileName,
				mimeType,
			);

			// Determine the binary data name
			const binaryDataKey = binaryDataName || 'data';

			// Build protection summary
			const protectionSummary: string[] = [];
			if (password) {
				protectionSummary.push('File encryption enabled');
			}
			if (protectWorkbook) {
				protectionSummary.push('Workbook structure protected');
			}
			if (protectWorksheets) {
				let worksheetInfo = 'Worksheets protected';
				if (worksheetSelection === 'name') {
					worksheetInfo += ` (${worksheetNames})`;
				} else if (worksheetSelection === 'index') {
					worksheetInfo += ` (indexes: ${worksheetIndexes})`;
				} else {
					worksheetInfo += ' (all)';
				}
				protectionSummary.push(worksheetInfo);
			}

			return [
				{
					json: {
						fileName,
						fileSize: excelBuffer.length,
						success: true,
						originalFileName,
						fileFormat: isXlsx ? 'xlsx' : 'xls',
						filePasswordProtected: !!password,
						workbookProtected: protectWorkbook,
						worksheetsProtected: protectWorksheets,
						worksheetProtectionType: protectWorksheets ? worksheetProtectionType : undefined,
						worksheetSelection: protectWorksheets ? worksheetSelection : undefined,
						protectionSummary,
						message: `Successfully secured Excel file (${isXlsx ? '.xlsx' : '.xls'}): ${protectionSummary.join(', ')}`,
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
		throw new NodeOperationError(this.getNode(), `Secure Excel file failed: ${errorMessage}`);
	}
}


