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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === DETAILED HEADER/FOOTER SETTINGS ===
	{
		displayName: 'Header Left Text',
		name: 'headerLeft',
		type: 'string',
		default: '',
		description: 'Text to appear on the left side of the header (e.g., Company Name, Date)',
		placeholder: 'Company Name',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Left Text Color',
		name: 'headerLeftColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the left header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Left Text Size',
		name: 'headerLeftFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the left header text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Center Text',
		name: 'headerCenter',
		type: 'string',
		default: '',
		description: 'Text to appear in the center of the header (e.g., Document Title, Confidential)',
		placeholder: 'Confidential',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Center Text Color',
		name: 'headerCenterColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the center header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Center Text Size',
		name: 'headerCenterFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the center header text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Right Text',
		name: 'headerRight',
		type: 'string',
		default: '',
		description: 'Text to appear on the right side of the header (e.g., Date, Page Number)',
		placeholder: 'Date',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Right Text Color',
		name: 'headerRightColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the right header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Header Right Text Size',
		name: 'headerRightFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the right header text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	// === FOOTER SETTINGS ===
	{
		displayName: 'Footer Left Text',
		name: 'footerLeft',
		type: 'string',
		default: '',
		description: 'Text to appear on the left side of the footer (e.g., Author, File Path)',
		placeholder: 'Author',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Left Text Color',
		name: 'footerLeftColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the left footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Left Text Size',
		name: 'footerLeftFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the left footer text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Center Text',
		name: 'footerCenter',
		type: 'string',
		default: '',
		description: 'Text to appear in the center of the footer. Use &P for current page, &N for total pages',
		placeholder: 'Page &P of &N',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Center Text Color',
		name: 'footerCenterColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the center footer text. Use the color picker to choose any color.',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Center Text Size',
		name: 'footerCenterFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the center footer text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Right Text',
		name: 'footerRight',
		type: 'string',
		default: '',
		description: 'Text to appear on the right side of the footer (e.g., Document Title, Copyright)',
		placeholder: 'Document Title',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Right Text Color',
		name: 'footerRightColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the right footer text. Use the color picker to choose any color.',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Footer Right Text Size',
		name: 'footerRightFontSize',
		type: 'number',
		default: 11,
		description: 'Font size for the right footer text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_with_header_footer.xlsx',
		description: 'Name for the processed Excel file (will have headers/footers added)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
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
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	// === WORKSHEET SETTINGS ===
	{
		displayName: 'Target Worksheet Name',
		name: 'worksheetName',
		type: 'string',
		default: 'Sheet1',
		description: 'Name of the specific worksheet to add header/footer to (e.g., Sheet1, Data, Summary)',
		placeholder: 'Sheet1',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	// === SIMPLIFIED HEADER/FOOTER SETTINGS ===
	{
		displayName: 'Simple Header Text',
		name: 'headerText',
		type: 'string',
		default: '',
		description: 'Single header text (alternative to detailed left/center/right settings above)',
		placeholder: 'Company Name - Confidential',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Simple Footer Text',
		name: 'footerText',
		type: 'string',
		default: '',
		description: 'Single footer text (alternative to detailed left/center/right settings above)',
		placeholder: 'Page &P of &N',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Simple Header Alignment',
		name: 'headerAlignment',
		type: 'options',
		default: 'center',
		description: 'Alignment for the simple header text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
		options: [
			{ name: 'Left Aligned', value: 'left' },
			{ name: 'Center Aligned', value: 'center' },
			{ name: 'Right Aligned', value: 'right' },
		],
	},
	{
		displayName: 'Simple Footer Alignment',
		name: 'footerAlignment',
		type: 'options',
		default: 'center',
		description: 'Alignment for the simple footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
		options: [
			{ name: 'Left Aligned', value: 'left' },
			{ name: 'Center Aligned', value: 'center' },
			{ name: 'Right Aligned', value: 'right' },
		],
	},
	{
		displayName: 'Simple Header/Footer Font Size',
		name: 'fontSize',
		type: 'number',
		default: 10,
		description: 'Font size for the simple header and footer text (6-72)',
		typeOptions: {
			minValue: 6,
			maxValue: 72,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Simple Header/Footer Text Color',
		name: 'fontColor',
		type: 'color',
		default: '#000000',
		description: 'Color for the simple header and footer text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
	{
		displayName: 'Apply to All Worksheets',
		name: 'applyToAllWorksheets',
		type: 'boolean',
		default: false,
		description: 'Apply the same header/footer to all worksheets in the workbook (ignores Target Worksheet Name)',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextHeaderFooterToExcel],
			},
		},
	},
];

/**
 * Add text header and footer to Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save Excel file
 * Adds customizable headers and footers to Excel worksheets with alignment, font, and color options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get all header/footer parameters
		const headerLeft = this.getNodeParameter('headerLeft', index, '') as string;
		const headerCenter = this.getNodeParameter('headerCenter', index, '') as string;
		const headerRight = this.getNodeParameter('headerRight', index, '') as string;
		const footerLeft = this.getNodeParameter('footerLeft', index, '') as string;
		const footerCenter = this.getNodeParameter('footerCenter', index, '') as string;
		const footerRight = this.getNodeParameter('footerRight', index, '') as string;

		// Get hex color values directly
		const headerLeftColor = this.getNodeParameter('headerLeftColor', index, '#000000') as string;
		const headerCenterColor = this.getNodeParameter('headerCenterColor', index, '#000000') as string;
		const headerRightColor = this.getNodeParameter('headerRightColor', index, '#000000') as string;
		const footerLeftColor = this.getNodeParameter('footerLeftColor', index, '#000000') as string;
		const footerCenterColor = this.getNodeParameter('footerCenterColor', index, '#000000') as string;
		const footerRightColor = this.getNodeParameter('footerRightColor', index, '#000000') as string;

		const headerLeftFontSize = this.getNodeParameter('headerLeftFontSize', index, 11) as number;
		const headerCenterFontSize = this.getNodeParameter('headerCenterFontSize', index, 11) as number;
		const headerRightFontSize = this.getNodeParameter('headerRightFontSize', index, 11) as number;
		const footerLeftFontSize = this.getNodeParameter('footerLeftFontSize', index, 11) as number;
		const footerCenterFontSize = this.getNodeParameter('footerCenterFontSize', index, 11) as number;
		const footerRightFontSize = this.getNodeParameter('footerRightFontSize', index, 11) as number;

		// Get additional parameters
		const worksheetName = this.getNodeParameter('worksheetName', index, 'Sheet1') as string;
		const headerText = this.getNodeParameter('headerText', index, '') as string;
		const footerText = this.getNodeParameter('footerText', index, '') as string;
		const headerAlignment = this.getNodeParameter('headerAlignment', index, 'center') as string;
		const footerAlignment = this.getNodeParameter('footerAlignment', index, 'center') as string;
		const fontSize = this.getNodeParameter('fontSize', index, 10) as number;
		const fontColor = this.getNodeParameter('fontColor', index, '#000000') as string;
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
			addTextHeaderFooterToExcelAction: {
				// Individual header/footer sections
				HeaderLeft: headerLeft,
				HeaderCenter: headerCenter,
				HeaderRight: headerRight,
				FooterLeft: footerLeft,
				FooterCenter: footerCenter,
				FooterRight: footerRight,
				HeaderLeftColor: headerLeftColor,
				HeaderLeftFontSize: headerLeftFontSize,
				HeaderCenterColor: headerCenterColor,
				HeaderCenterFontSize: headerCenterFontSize,
				HeaderRightColor: headerRightColor,
				HeaderRightFontSize: headerRightFontSize,
				FooterLeftColor: footerLeftColor,
				FooterLeftFontSize: footerLeftFontSize,
				FooterCenterColor: footerCenterColor,
				FooterCenterFontSize: footerCenterFontSize,
				FooterRightColor: footerRightColor,
				FooterRightFontSize: footerRightFontSize,
				// Additional parameters
				worksheetName: worksheetName,
				headerText: headerText,
				footerText: footerText,
				headerAlignment: headerAlignment,
				footerAlignment: footerAlignment,
				fontSize: fontSize,
				fontColor: fontColor,
				applyToAllWorksheets: applyToAllWorksheets,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelAddTextHeaderFooter',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_with_header_footer';
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
						// Individual header/footer sections
						headerLeft,
						headerCenter,
						headerRight,
						footerLeft,
						footerCenter,
						footerRight,
						// Additional parameters
						worksheetName,
						headerText,
						footerText,
						headerAlignment,
						footerAlignment,
						fontSize,
						fontColor,
						applyToAllWorksheets,
						message: 'Header and footer added to Excel file successfully',
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
		throw new Error(`Add text header/footer to Excel failed: ${errorMessage}`);
	}
}

