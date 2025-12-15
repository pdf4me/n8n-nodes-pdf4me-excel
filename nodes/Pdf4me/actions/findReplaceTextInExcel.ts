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
				operation: [ActionConstants.FindReplaceTextInExcel],
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
				operation: [ActionConstants.FindReplaceTextInExcel],
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
				operation: [ActionConstants.FindReplaceTextInExcel],
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
				operation: [ActionConstants.FindReplaceTextInExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === FIND AND REPLACE SETTINGS ===
	{
		displayName: 'Replace Operations',
		name: 'phrases',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		placeholder: 'Add Replace Operation',
		description: 'Define search and replace operations to perform on the Excel file',
		displayOptions: {
			show: {
				operation: [ActionConstants.FindReplaceTextInExcel],
			},
		},
		options: [
			{
				name: 'phraseValues',
				displayName: 'Replace Operation',
				values: [
					{
						displayName: 'Search Text',
						name: 'searchText',
						type: 'string',
						default: '',
						required: true,
						description: 'Text to search for in the Excel file',
						placeholder: 'Dulce',
					},
					{
						displayName: 'Replacement Text',
						name: 'replacementText',
						type: 'string',
						default: '',
						required: true,
						description: 'Text to replace with',
						placeholder: 'PDF4me',
					},
					{
						displayName: 'Use Regular Expression',
						name: 'isExpression',
						type: 'boolean',
						default: false,
						description: 'Whether to use regular expressions for search',
					},
					{
						displayName: 'Match Entire Cell',
						name: 'matchEntireCell',
						type: 'boolean',
						default: false,
						description: 'Whether to match entire cell content only',
					},
					{
						displayName: 'Case Sensitive',
						name: 'caseSensitive',
						type: 'boolean',
						default: false,
						description: 'Whether to perform case-sensitive search',
					},
					{
						displayName: 'Apply Formatting',
						name: 'applyFormatting',
						type: 'boolean',
						default: false,
						description: 'Whether to apply custom formatting to replacement text',
					},
					{
						displayName: 'Font Name',
						name: 'fontName',
						type: 'options',
						default: '',
						description: 'Font name for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: '' },
							{ name: 'Arial', value: 'Arial' },
							{ name: 'Times New Roman', value: 'Times New Roman' },
							{ name: 'Courier New', value: 'Courier New' },
							{ name: 'Verdana', value: 'Verdana' },
							{ name: 'Calibri', value: 'Calibri' },
							{ name: 'Helvetica', value: 'Helvetica' },
							{ name: 'Georgia', value: 'Georgia' },
							{ name: 'Tahoma', value: 'Tahoma' },
						],
					},
					{
						displayName: 'Font Color',
						name: 'fontColor',
						type: 'color',
						default: '',
						description: 'Font color for replacement text (hex code)',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
					},
					{
						displayName: 'Font Size',
						name: 'fontSize',
						type: 'number',
						default: 0,
						description: 'Font size for replacement text (0 to inherit)',
						typeOptions: {
							minValue: 0,
							maxValue: 72,
						},
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
					},
					{
						displayName: 'Bold',
						name: 'bold',
						type: 'options',
						default: 'Inherit',
						description: 'Bold formatting for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'Apply Bold', value: 'Apply' },
							{ name: 'Remove Bold', value: 'Remove' },
						],
					},
					{
						displayName: 'Italic',
						name: 'italic',
						type: 'options',
						default: 'Inherit',
						description: 'Italic formatting for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'Apply Italic', value: 'Apply' },
							{ name: 'Remove Italic', value: 'Remove' },
						],
					},
					{
						displayName: 'Strikethrough Type',
						name: 'strikethroughType',
						type: 'options',
						default: 'Inherit',
						description: 'Strikethrough formatting for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'None', value: 'None' },
							{ name: 'Single', value: 'Single' },
							{ name: 'Double', value: 'Double' },
						],
					},
					{
						displayName: 'Underline Type',
						name: 'underlineType',
						type: 'options',
						default: 'Inherit',
						description: 'Underline formatting for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'None', value: 'None' },
							{ name: 'Single', value: 'Single' },
							{ name: 'Double', value: 'Double' },
							{ name: 'Dash', value: 'Dash' },
							{ name: 'Dotted', value: 'Dotted' },
						],
					},
					{
						displayName: 'Script Type',
						name: 'scriptType',
						type: 'options',
						default: 'Inherit',
						description: 'Script type (superscript/subscript) for replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'None', value: 'None' },
							{ name: 'Superscript', value: 'Superscript' },
							{ name: 'Subscript', value: 'Subscript' },
						],
					},
					{
						displayName: 'Font Scheme Type',
						name: 'fontSchemeType',
						type: 'options',
						default: 'Inherit',
						description: 'Font scheme type to apply to the text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'None', value: 'None' },
							{ name: 'Major', value: 'Major' },
							{ name: 'Minor', value: 'Minor' },
						],
					},
					{
						displayName: 'Theme Colour',
						name: 'themeColour',
						type: 'options',
						default: 'Inherit',
						description: 'Theme colour to apply to the replacement text',
						displayOptions: {
							show: {
								applyFormatting: [true],
							},
						},
						options: [
							{ name: 'Inherit (No Change)', value: 'Inherit' },
							{ name: 'None', value: 'None' },
							{ name: 'Accent1', value: 'Accent1' },
							{ name: 'Accent2', value: 'Accent2' },
							{ name: 'Accent3', value: 'Accent3' },
							{ name: 'Accent4', value: 'Accent4' },
							{ name: 'Accent5', value: 'Accent5' },
							{ name: 'Accent6', value: 'Accent6' },
							{ name: 'Dark1', value: 'Dark1' },
							{ name: 'Dark2', value: 'Dark2' },
							{ name: 'Light1', value: 'Light1' },
							{ name: 'Light2', value: 'Light2' },
						],
					},
				],
			},
		],
	},
	{
		displayName: 'Culture Name',
		name: 'cultureName',
		type: 'string',
		default: 'en-US',
		description: 'Culture name for text processing (e.g., en-US, de-DE, fr-FR)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.FindReplaceTextInExcel],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_replaced.xlsx',
		description: 'Name for the processed Excel file (will have text replaced)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.FindReplaceTextInExcel],
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
				operation: [ActionConstants.FindReplaceTextInExcel],
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
				operation: [ActionConstants.FindReplaceTextInExcel],
			},
		},
	},
];

/**
 * Find and replace text in Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save Excel file
 * Supports multiple search/replace operations with custom formatting options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get find and replace parameters
		const phrasesData = this.getNodeParameter('phrases', index, {}) as IDataObject;
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

		// Parse and build phrases array
		const phrases: IDataObject[] = [];
		const phraseValues = (phrasesData.phraseValues as IDataObject[]) || [];

		if (phraseValues.length === 0) {
			throw new NodeOperationError(this.getNode(), 'At least one replace operation is required');
		}

		for (const phrase of phraseValues) {
			const searchText = phrase.searchText as string;
			const replacementText = phrase.replacementText as string;

			if (!searchText) {
				throw new NodeOperationError(this.getNode(), 'Search text is required for all replace operations');
			}
			if (replacementText === undefined || replacementText === null) {
				throw new NodeOperationError(this.getNode(), 'Replacement text is required for all replace operations');
			}

			const isExpression = phrase.isExpression as boolean || false;
			const matchEntireCell = phrase.matchEntireCell as boolean || false;
			const caseSensitive = phrase.caseSensitive as boolean || false;
			const applyFormatting = phrase.applyFormatting as boolean || false;

			const phraseObj: IDataObject = {
				SearchText: searchText,
				ReplacementText: replacementText,
				IsExpression: isExpression,
				MatchEntireCell: matchEntireCell,
				CaseSensitive: caseSensitive,
			};

			// Add formatting if enabled
			if (applyFormatting) {
				const formatting: IDataObject = {};

				const fontName = phrase.fontName as string;
				const fontColor = phrase.fontColor as string;
				const fontSize = phrase.fontSize as number;
				const bold = phrase.bold as string || 'Inherit';
				const italic = phrase.italic as string || 'Inherit';
				const strikethroughType = phrase.strikethroughType as string || 'Inherit';
				const underlineType = phrase.underlineType as string || 'Inherit';
				const scriptType = phrase.scriptType as string || 'Inherit';
				const fontSchemeType = phrase.fontSchemeType as string || 'Inherit';
				const themeColour = phrase.themeColour as string || 'Inherit';

				if (fontName) {
					formatting.FontName = fontName;
				}
				if (fontColor) {
					formatting.FontColor = fontColor;
				}
				if (fontSize && fontSize > 0) {
					formatting.FontSize = fontSize;
				}

				formatting.Bold = bold;
				formatting.Italic = italic;
				formatting.StrikethroughType = strikethroughType;
				formatting.UnderlineType = underlineType;
				formatting.ScriptType = scriptType;
				formatting.FontSchemeType = fontSchemeType;
				formatting.ThemeColour = themeColour;

				phraseObj.Formatting = formatting;
			}

			phrases.push(phraseObj);
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			ReplaceTextToExcelAction: {
				Phrases: phrases,
				CultureName: cultureName,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelFindAndReplaceTextInExcel',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_replaced';
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

			// Build summary of operations for JSON output
			const operations = phraseValues.map((phrase) => ({
				searchText: phrase.searchText,
				replacementText: phrase.replacementText,
				isExpression: phrase.isExpression || false,
				matchEntireCell: phrase.matchEntireCell || false,
				caseSensitive: phrase.caseSensitive || false,
				formattingApplied: phrase.applyFormatting || false,
			}));

			return [
				{
					json: {
						fileName,
						fileSize: excelBuffer.length,
						success: true,
						originalFileName,
						operationsCount: phrases.length,
						operations,
						cultureName,
						message: `Text replacement completed successfully with ${phrases.length} operation(s)`,
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
		throw new NodeOperationError(this.getNode(), `Find and replace text in Excel failed: ${errorMessage}`);
	}
}


