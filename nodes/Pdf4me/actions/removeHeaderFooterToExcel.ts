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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === HEADER REMOVAL SETTINGS ===
	{
		displayName: 'Remove Header Left',
		name: 'removeHeaderLeft',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the left header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Remove Header Center',
		name: 'removeHeaderCenter',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the center header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Remove Header Right',
		name: 'removeHeaderRight',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the right header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	// === FOOTER REMOVAL SETTINGS ===
	{
		displayName: 'Remove Footer Left',
		name: 'removeFooterLeft',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the left footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Remove Footer Center',
		name: 'removeFooterCenter',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the center footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Remove Footer Right',
		name: 'removeFooterRight',
		type: 'boolean',
		default: false,
		description: 'Whether to remove the right footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_without_header_footer.xlsx',
		description: 'Name for the processed Excel file (will have headers/footers removed)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
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
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	// === WORKSHEET SETTINGS ===
	{
		displayName: 'Target Worksheet Name',
		name: 'worksheetName',
		type: 'string',
		default: 'Sheet1',
		description: 'Name of the specific worksheet to remove header/footer from (e.g., Sheet1, Data, Summary)',
		placeholder: 'Sheet1',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Apply to All Worksheets',
		name: 'applyToAllWorksheets',
		type: 'boolean',
		default: false,
		description: 'Apply the removal to all worksheets in the workbook (ignores Target Worksheet Name)',
		displayOptions: {
			show: {
				operation: [ActionConstants.RemoveHeaderFooterToExcel],
			},
		},
	},
];

/**
 * Remove text header and footer from Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save Excel file
 * Removes headers and footers from Excel worksheets with selective removal options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get removal parameters
		const removeHeaderLeft = this.getNodeParameter('removeHeaderLeft', index, false) as boolean;
		const removeHeaderCenter = this.getNodeParameter('removeHeaderCenter', index, false) as boolean;
		const removeHeaderRight = this.getNodeParameter('removeHeaderRight', index, false) as boolean;
		const removeFooterLeft = this.getNodeParameter('removeFooterLeft', index, false) as boolean;
		const removeFooterCenter = this.getNodeParameter('removeFooterCenter', index, false) as boolean;
		const removeFooterRight = this.getNodeParameter('removeFooterRight', index, false) as boolean;

		// Get additional parameters
		const worksheetName = this.getNodeParameter('worksheetName', index, 'Sheet1') as string;
		const applyToAllWorksheets = this.getNodeParameter('applyToAllWorksheets', index, false) as boolean;

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
			removeHeaderFooterToExcelAction: {
				worksheetName: worksheetName,
				RemoveHeaderLeft: removeHeaderLeft,
				RemoveHeaderCenter: removeHeaderCenter,
				RemoveHeaderRight: removeHeaderRight,
				RemoveFooterLeft: removeFooterLeft,
				RemoveFooterCenter: removeFooterCenter,
				RemoveFooterRight: removeFooterRight,
				applyToAllWorksheets: applyToAllWorksheets,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelRemoveHeaderFooter',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_without_header_footer';
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

			return [
				{
					json: {
						fileName,
						fileSize: excelBuffer.length,
						success: true,
						originalFileName,
						// Removal settings
						removeHeaderLeft,
						removeHeaderCenter,
						removeHeaderRight,
						removeFooterLeft,
						removeFooterCenter,
						removeFooterRight,
						// Additional parameters
						worksheetName,
						applyToAllWorksheets,
						message: 'Header and footer removed from Excel file successfully',
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				},
			];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Remove header/footer from Excel failed: ${errorMessage}`);
	}
}

