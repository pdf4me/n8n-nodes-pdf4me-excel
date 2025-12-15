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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === EXTRACT SETTINGS ===
	{
		displayName: 'Extract By',
		name: 'extractBy',
		type: 'options',
		default: 'name',
		description: 'Choose whether to extract worksheets by name or by index',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExtractWorksheetFromExcel],
			},
		},
		options: [
			{
				name: 'Worksheet Names',
				value: 'name',
				description: 'Extract worksheets by their names',
			},
			{
				name: 'Worksheet Indexes',
				value: 'index',
				description: 'Extract worksheets by their index positions',
			},
		],
	},
	{
		displayName: 'Worksheet Names',
		name: 'worksheetNames',
		type: 'string',
		default: '',
		description: 'Comma-separated list of worksheet names to extract (e.g., "Sheet1, Sheet2, Data"). All other worksheets will be removed.',
		placeholder: 'Sheet1, Sheet2',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExtractWorksheetFromExcel],
				extractBy: ['name'],
			},
		},
	},
	{
		displayName: 'Worksheet Indexes',
		name: 'worksheetIndexes',
		type: 'string',
		default: '',
		description: 'Comma-separated list of worksheet indexes to extract (1-based, e.g., "1, 2, 3"). All other worksheets will be removed.',
		placeholder: '1, 2, 3',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExtractWorksheetFromExcel],
				extractBy: ['index'],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_extracted.xlsx',
		description: 'Name for the processed Excel file (will contain only extracted worksheet(s))',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.ExtractWorksheetFromExcel],
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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
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
				operation: [ActionConstants.ExtractWorksheetFromExcel],
			},
		},
	},
];

/**
 * Extract worksheets from Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save Excel file
 * Extracts (keeps) specified worksheets by name or index, removing all others
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get extract parameters
		const extractBy = this.getNodeParameter('extractBy', index, 'name') as string;
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

		// Validate extraction criteria
		if (extractBy === 'name') {
			if (!worksheetNames || worksheetNames.trim() === '') {
				throw new NodeOperationError(this.getNode(), 'Worksheet names are required when extracting by name');
			}
		} else if (extractBy === 'index') {
			if (!worksheetIndexes || worksheetIndexes.trim() === '') {
				throw new NodeOperationError(this.getNode(), 'Worksheet indexes are required when extracting by index');
			}
		}

		// Determine which field to send based on extractBy
		const finalWorksheetNames = extractBy === 'name' ? worksheetNames : '';
		const finalWorksheetIndexes = extractBy === 'index' ? worksheetIndexes : '';

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			ExtractWorksheetToExcelAction: {
				WorksheetNames: finalWorksheetNames,
				WorksheetIndexes: finalWorksheetIndexes,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelExtractWorksheet',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_extracted';
				fileName = `${baseName}.xlsx`;
			}

			// Ensure .xlsx extension
			if (!fileName.toLowerCase().endsWith('.xlsx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.xlsx`;
			}

			// Handle the response - API returns JSON data, not binary Excel
			let decodedContent: string;
			let parsedData: unknown;

			// The API returns JSON data in the document field
			if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;

				// Check if the response has a document field
				if (response.document) {
					const document = response.document;

					// The document could be a string (base64) or an object with nested fields
					if (typeof document === 'string') {
						// Document itself is the base64 content
						try {
							const buffer = Buffer.from(document, 'base64');
							decodedContent = buffer.toString('utf8');
						} catch (error) {
							// If base64 decoding fails, use the string directly
							decodedContent = document;
						}
					} else if (typeof document === 'object' && document !== null) {
						// Document is an object, extract content from possible fields
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

						try {
							const buffer = Buffer.from(docContent, 'base64');
							decodedContent = buffer.toString('utf8');
						} catch (error) {
							// If base64 decoding fails, use the string directly
							decodedContent = docContent;
						}
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

					try {
						const buffer = Buffer.from(docContent, 'base64');
						decodedContent = buffer.toString('utf8');
					} catch (error) {
						// If base64 decoding fails, use the string directly
						decodedContent = docContent;
					}
				}
			} else if (typeof responseData === 'string') {
				// Direct string response
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

			// Try to parse the content as JSON
			try {
				parsedData = JSON.parse(decodedContent);
			} catch (jsonError) {
				// If JSON parsing fails, create a structured response
				parsedData = {
					fileName: fileName.replace(/\.[^.]*$/, '.txt'),
					fileSize: decodedContent.length,
					fileType: 'Text',
					originalFileName,
					extractBy,
					worksheetNames: extractBy === 'name' ? worksheetNames : undefined,
					worksheetIndexes: extractBy === 'index' ? worksheetIndexes : undefined,
					message: 'Successfully processed Excel data',
					rawContent: decodedContent,
					note: 'Content processed as text data',
				};
			}

			// Create JSON content
			const jsonContent = JSON.stringify(parsedData, null, 2);
			const jsonBuffer = Buffer.from(jsonContent, 'utf8');

			// Create TXT content
			const txtContent = decodedContent;
			const txtBuffer = Buffer.from(txtContent, 'utf8');

			// Create JSON file
			const jsonFileName = fileName.replace(/\.[^.]*$/, '.json');
			const jsonBinaryData = await this.helpers.prepareBinaryData(
				jsonBuffer,
				jsonFileName,
				'application/json',
			);

			// Create TXT file
			const txtFileName = fileName.replace(/\.[^.]*$/, '.txt');
			const txtBinaryData = await this.helpers.prepareBinaryData(
				txtBuffer,
				txtFileName,
				'text/plain',
			);

			// Determine the binary data names
			const jsonBinaryDataKey = `${binaryDataName || 'data'}_json`;
			const txtBinaryDataKey = `${binaryDataName || 'data'}_txt`;

			// Build success message
			let extractionInfo: string;
			if (extractBy === 'name') {
				extractionInfo = `worksheet(s) '${worksheetNames}' by name`;
			} else {
				extractionInfo = `worksheet(s) at index(es) '${worksheetIndexes}'`;
			}

			return [
				{
					json: {
						jsonFileName,
						txtFileName,
						jsonFileSize: jsonBuffer.length,
						txtFileSize: txtBuffer.length,
						success: true,
						originalFileName,
						extractBy,
						worksheetNames: extractBy === 'name' ? worksheetNames : undefined,
						worksheetIndexes: extractBy === 'index' ? worksheetIndexes : undefined,
						parsedData: parsedData as IDataObject,
						message: `Successfully extracted ${extractionInfo} and created JSON and TXT outputs`,
					},
					binary: {
						[jsonBinaryDataKey]: jsonBinaryData,
						[txtBinaryDataKey]: txtBinaryData,
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
		throw new NodeOperationError(this.getNode(), `Extract worksheet from Excel failed: ${errorMessage}`);
	}
}


