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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === IMAGE INPUT SETTINGS ===
	{
		displayName: 'Image Input Method',
		name: 'imageInputType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the image for the header/footer',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
		options: [
			{
				name: 'From Previous Node (Binary Data)',
				value: 'binaryData',
				description: 'Use image file passed from a previous n8n node',
			},
			{
				name: 'Base64 Encoded String',
				value: 'base64',
				description: 'Provide image content as base64 encoded string',
			},
			{
				name: 'Download from URL',
				value: 'url',
				description: 'Download image directly from a web URL',
			},
		],
	},
	{
		displayName: 'Image Binary Data Property Name',
		name: 'imageBinaryPropertyName',
		type: 'string',
		required: true,
		default: 'image',
		description: 'Name of the binary property containing the image file',
		placeholder: 'image',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
				imageInputType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Base64 Encoded Image Content',
		name: 'imageBase64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the image data (PNG, JPEG, etc.)',
		placeholder: 'iVBORw0KGgoAAAANSUhEUgAAAMcAAAAwCAIAAAAq8PkwAAAAAXNSR0IArs4c6Q...',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
				imageInputType: ['base64'],
			},
		},
	},
	{
		displayName: 'Image URL',
		name: 'imageUrl',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the image from (must be publicly accessible)',
		placeholder: 'https://example.com/logo.png',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
				imageInputType: ['url'],
			},
		},
	},
	// === HEADER/FOOTER SETTINGS ===
	{
		displayName: 'Apply to Header or Footer',
		name: 'isHeader',
		type: 'boolean',
		default: true,
		description: 'Whether to apply the image to the header (true) or footer (false)',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Image Position',
		name: 'position',
		type: 'options',
		default: 'Center',
		description: 'Position of the image in the header/footer',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
		options: [
			{ name: 'Left', value: 'Left' },
			{ name: 'Center', value: 'Center' },
			{ name: 'Right', value: 'Right' },
		],
	},
	// === WORKSHEET SETTINGS ===
	{
		displayName: 'Target Worksheets',
		name: 'worksheetNames',
		type: 'string',
		default: 'Sheet1',
		description: 'Comma-separated list of worksheet names to apply the image header/footer to (e.g., Sheet1, Sheet2, Data)',
		placeholder: 'Sheet1, Sheet2',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	// === MARGIN SETTINGS ===
	{
		displayName: 'Top Margin',
		name: 'topMargin',
		type: 'number',
		default: 1.9,
		description: 'Top margin in centimeters',
		typeOptions: {
			minValue: 0,
			maxValue: 10,
			numberPrecision: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Bottom Margin',
		name: 'bottomMargin',
		type: 'number',
		default: 1.9,
		description: 'Bottom margin in centimeters',
		typeOptions: {
			minValue: 0,
			maxValue: 10,
			numberPrecision: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Left Margin',
		name: 'leftMargin',
		type: 'number',
		default: 1.9,
		description: 'Left margin in centimeters',
		typeOptions: {
			minValue: 0,
			maxValue: 10,
			numberPrecision: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Right Margin',
		name: 'rightMargin',
		type: 'number',
		default: 1.9,
		description: 'Right margin in centimeters',
		typeOptions: {
			minValue: 0,
			maxValue: 10,
			numberPrecision: 1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_with_image_header_footer.xlsx',
		description: 'Name for the processed Excel file (will have image header/footer added)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
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
				operation: [ActionConstants.AddImageHeaderFooterToExcel],
			},
		},
	},
];

/**
 * Add image header and footer to Excel files using PDF4Me API
 * Process: Read Excel file → Read Image → Encode both to base64 → Send API request → Poll for completion → Save Excel file
 * Adds image-based headers and footers to Excel worksheets with position and margin options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const imageInputType = this.getNodeParameter('imageInputType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get header/footer parameters
		const isHeader = this.getNodeParameter('isHeader', index, true) as boolean;
		const position = this.getNodeParameter('position', index, 'Center') as string;
		const worksheetNamesStr = this.getNodeParameter('worksheetNames', index, 'Sheet1') as string;
		const topMargin = this.getNodeParameter('topMargin', index, 1.9) as number;
		const bottomMargin = this.getNodeParameter('bottomMargin', index, 1.9) as number;
		const leftMargin = this.getNodeParameter('leftMargin', index, 1.9) as number;
		const rightMargin = this.getNodeParameter('rightMargin', index, 1.9) as number;

		// Parse worksheet names
		const worksheetNames = worksheetNamesStr
			.split(',')
			.map((name) => name.trim())
			.filter((name) => name.length > 0);

		let docContent: string;
		let originalFileName = docName;

		// Handle different Excel input data types
		if (inputDataType === 'binaryData') {
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
			docContent = this.getNodeParameter('base64Content', index) as string;

			// Remove data URL prefix if present
			if (docContent.includes(',')) {
				docContent = docContent.split(',')[1];
			}
		} else if (inputDataType === 'url') {
			const url = this.getNodeParameter('url', index) as string;

			if (!url || url.trim() === '') {
				throw new NodeOperationError(this.getNode(), 'URL is required when using URL input type');
			}

			try {
				const response = await this.helpers.httpRequest({
					method: 'GET',
					url,
					encoding: 'arraybuffer',
					returnFullResponse: true,
				});

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
				throw new NodeOperationError(this.getNode(), `Failed to download Excel file from URL: ${errorMessage}`);
			}
		} else {
			throw new NodeOperationError(this.getNode(), `Unsupported input data type: ${inputDataType}`);
		}

		// Validate Excel content
		if (!docContent || docContent.trim() === '') {
			throw new NodeOperationError(this.getNode(), 'Excel content is required');
		}

		// Handle different image input types
		let imageContent: string;

		if (imageInputType === 'binaryData') {
			const imageBinaryPropertyName = this.getNodeParameter('imageBinaryPropertyName', index) as string;
			const item = this.getInputData(index);

			if (!item[0].binary || !item[0].binary[imageBinaryPropertyName]) {
				throw new NodeOperationError(this.getNode(), `No binary data found in property '${imageBinaryPropertyName}'`);
			}

			const buffer = await this.helpers.getBinaryDataBuffer(index, imageBinaryPropertyName);
			imageContent = buffer.toString('base64');
		} else if (imageInputType === 'base64') {
			imageContent = this.getNodeParameter('imageBase64Content', index) as string;

			// Remove data URL prefix if present
			if (imageContent.includes(',')) {
				imageContent = imageContent.split(',')[1];
			}
		} else if (imageInputType === 'url') {
			const imageUrl = this.getNodeParameter('imageUrl', index) as string;

			if (!imageUrl || imageUrl.trim() === '') {
				throw new NodeOperationError(this.getNode(), 'Image URL is required when using URL input type');
			}

			try {
				const response = await this.helpers.httpRequest({
					method: 'GET',
					url: imageUrl,
					encoding: 'arraybuffer',
					returnFullResponse: true,
				});

				const buffer = Buffer.from(response.body as ArrayBuffer);
				imageContent = buffer.toString('base64');
			} catch (error) {
				const errorMessage = error instanceof Error ? error.message : 'Unknown error';
				throw new NodeOperationError(this.getNode(), `Failed to download image from URL: ${errorMessage}`);
			}
		} else {
			throw new NodeOperationError(this.getNode(), `Unsupported image input type: ${imageInputType}`);
		}

		// Validate image content
		if (!imageContent || imageContent.trim() === '') {
			throw new NodeOperationError(this.getNode(), 'Image content is required');
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			imageContent,
			addImageHeaderFooterToExcelAction: {
				IsHeader: isHeader,
				Position: position,
				WorksheetNames: worksheetNames,
				TopMargin: topMargin,
				bottomMargin: bottomMargin,
				LeftMargin: leftMargin,
				RightMargin: rightMargin,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelAddImageHeaderFooter',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_with_image_header_footer';
				fileName = `${baseName}.xlsx`;
			}

			// Ensure .xlsx extension
			if (!fileName.toLowerCase().endsWith('.xlsx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.xlsx`;
			}

			// Handle the response - Excel API returns JSON with embedded base64 file
			let excelBuffer: Buffer;

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
						isHeader,
						position,
						worksheetNames,
						topMargin,
						bottomMargin,
						leftMargin,
						rightMargin,
						message: 'Image header/footer added to Excel file successfully',
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
		throw new NodeOperationError(this.getNode(), `Add image header/footer to Excel failed: ${errorMessage}`);
	}
}

