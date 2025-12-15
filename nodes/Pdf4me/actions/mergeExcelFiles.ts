import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import { NodeOperationError, NodeApiError } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';


export const description: INodeProperties[] = [
	// === FILES TO MERGE ===
	{
		displayName: 'Files to Merge',
		name: 'filesToMerge',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		placeholder: 'Add File',
		description: 'Add Excel files to merge together',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeExcelFiles],
			},
		},
		options: [
			{
				name: 'fileValues',
				displayName: 'File',
				values: [
					{
						displayName: 'File Input Method',
						name: 'inputMethod',
						type: 'options',
						default: 'binaryData',
						description: 'How to provide the Excel file',
						options: [
							{
								name: 'From Previous Node (Binary Data)',
								value: 'binaryData',
								description: 'Use Excel file from binary data',
							},
							{
								name: 'Base64 Encoded String',
								value: 'base64',
								description: 'Provide base64 encoded file content',
							},
							{
								name: 'Download from URL',
								value: 'url',
								description: 'Download file from a URL',
							},
						],
					},
					{
						displayName: 'Binary Data Property Name',
						name: 'binaryPropertyName',
						type: 'string',
						default: 'data',
						description: 'Name of the binary property containing the file',
						placeholder: 'data',
						displayOptions: {
							show: {
								inputMethod: ['binaryData'],
							},
						},
					},
					{
						displayName: 'Base64 Content',
						name: 'base64Content',
						type: 'string',
						typeOptions: {
							alwaysOpenEditWindow: true,
						},
						default: '',
						description: 'Base64 encoded file content',
						placeholder: 'UEsDBBQAAAAIAAAAIQA08b3+ywUAALMZAAATACQA...',
						displayOptions: {
							show: {
								inputMethod: ['base64'],
							},
						},
					},
					{
						displayName: 'File URL',
						name: 'fileUrl',
						type: 'string',
						default: '',
						description: 'URL to download the file from',
						placeholder: 'https://example.com/file.xlsx',
						displayOptions: {
							show: {
								inputMethod: ['url'],
							},
						},
					},
					{
						displayName: 'Filename',
						name: 'filename',
						type: 'string',
						default: '',
						description: 'Filename with extension (e.g., file1.xlsx)',
						placeholder: 'file1.xlsx',
					},
					{
						displayName: 'Sort Position',
						name: 'sortPosition',
						type: 'number',
						default: 0,
						description: 'Position in merge order (lowest first, 0 = first)',
						typeOptions: {
							minValue: 0,
						},
					},
					{
						displayName: 'Specific Worksheets',
						name: 'worksheetsToMerge',
						type: 'string',
						default: '',
						description: 'Comma-separated list of specific worksheet names to merge from this file (leave empty to merge all)',
						placeholder: 'Sheet1, Sheet2',
					},
				],
			},
		],
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output Format',
		name: 'outputFormat',
		type: 'options',
		default: 'XLSX',
		description: 'Format for the merged output file',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeExcelFiles],
			},
		},
		options: [
			{
				name: 'XLSX (Excel 2007+)',
				value: 'XLSX',
				description: 'Modern Excel format (.xlsx)',
			},
			{
				name: 'XLS (Excel 97-2003)',
				value: 'XLS',
				description: 'Legacy Excel format (.xls)',
			},
			{
				name: 'PDF',
				value: 'PDF',
				description: 'Portable Document Format (.pdf)',
			},
			{
				name: 'CSV',
				value: 'CSV',
				description: 'Comma-Separated Values (.csv)',
			},
		],
	},
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'merged-workbook',
		description: 'Name for the merged output file (without extension, will be added automatically)',
		placeholder: 'merged-workbook',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeExcelFiles],
			},
		},
	},
	{
		displayName: 'Output Binary Data Name',
		name: 'binaryDataName',
		type: 'string',
		default: 'data',
		description: 'Name for the binary data in the n8n output',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeExcelFiles],
			},
		},
	},
];

/**
 * Merge multiple Excel files into one using PDF4Me API
 * Process: Read multiple Excel files → Encode to base64 → Send API request → Poll for completion → Save merged file
 * Supports merging multiple Excel files with customizable worksheet selection and output format
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const filesToMergeData = this.getNodeParameter('filesToMerge', index, {}) as IDataObject;
		const outputFormat = this.getNodeParameter('outputFormat', index, 'XLSX') as string;
		const outputFileName = this.getNodeParameter('outputFileName', index, 'merged-workbook') as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index, 'data') as string;

		const fileValues = (filesToMergeData.fileValues as IDataObject[]) || [];

		if (fileValues.length === 0) {
			throw new NodeOperationError(this.getNode(), 'At least one file is required to merge');
		}

		if (fileValues.length === 1) {
			throw new NodeOperationError(this.getNode(), 'At least two files are required to merge');
		}

		// Process each file
		const documents: IDataObject[] = [];

		for (let i = 0; i < fileValues.length; i++) {
			const file = fileValues[i];
			const inputMethod = file.inputMethod as string;
			const filename = (file.filename as string) || `file${i + 1}.xlsx`;
			const sortPosition = (file.sortPosition as number) || i;
			const worksheetsToMerge = (file.worksheetsToMerge as string) || '';

			let fileContent: string;

			// Handle different input methods
			if (inputMethod === 'binaryData') {
				const binaryPropertyName = file.binaryPropertyName as string;
				const item = this.getInputData(index);

				if (!item[0].binary || !item[0].binary[binaryPropertyName]) {
					throw new NodeOperationError(this.getNode(), `File ${i + 1}: No binary data found in property '${binaryPropertyName}'`);
				}

				const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
				fileContent = buffer.toString('base64');
			} else if (inputMethod === 'base64') {
				fileContent = file.base64Content as string;

				// Remove data URL prefix if present
				if (fileContent.includes(',')) {
					fileContent = fileContent.split(',')[1];
				}

				if (!fileContent || fileContent.trim() === '') {
					throw new NodeOperationError(this.getNode(), `File ${i + 1}: Base64 content is required`);
				}
			} else if (inputMethod === 'url') {
				const url = file.fileUrl as string;

				if (!url || url.trim() === '') {
					throw new NodeOperationError(this.getNode(), `File ${i + 1}: URL is required when using URL input type`);
				}

				try {
					const response = await this.helpers.httpRequest({
						method: 'GET',
						url,
						encoding: 'arraybuffer',
						returnFullResponse: true,
					});

					const buffer = Buffer.from(response.body as ArrayBuffer);
					fileContent = buffer.toString('base64');
				} catch (error) {
					const errorMessage = error instanceof Error ? error.message : 'Unknown error';
					throw new NodeOperationError(this.getNode(), `File ${i + 1}: Failed to download from URL: ${errorMessage}`);
				}
			} else {
				throw new NodeOperationError(this.getNode(), `File ${i + 1}: Unsupported input method: ${inputMethod}`);
			}

			// Build document object
			const documentObj: IDataObject = {
				Filename: filename,
				FileContent: fileContent,
				SortPosition: sortPosition,
			};

			// Add worksheets to merge if specified
			if (worksheetsToMerge && worksheetsToMerge.trim() !== '') {
				const worksheetList = worksheetsToMerge
					.split(',')
					.map(name => name.trim())
					.filter(name => name !== '');
				if (worksheetList.length > 0) {
					documentObj.WorksheetsToMerge = worksheetList;
				}
			}

			documents.push(documentObj);
		}

		// Sort documents by sortPosition
		documents.sort((a, b) => (a.SortPosition as number) - (b.SortPosition as number));

		// Build the request body according to the API specification
		const body: IDataObject = {
			MergeFilesToExcelAction: {
				OutputFileName: outputFileName,
				OutputFormat: outputFormat,
				Documents: documents,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelMergeFiles',
			body,
		);

		if (responseData) {
			// Determine file extension based on output format
			let fileExtension = '.xlsx';
			let mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

			switch (outputFormat.toUpperCase()) {
			case 'XLSX':
				fileExtension = '.xlsx';
				mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
				break;
			case 'XLS':
				fileExtension = '.xls';
				mimeType = 'application/vnd.ms-excel';
				break;
			case 'PDF':
				fileExtension = '.pdf';
				mimeType = 'application/pdf';
				break;
			case 'CSV':
				fileExtension = '.csv';
				mimeType = 'text/csv';
				break;
			}

			// Generate filename
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				fileName = 'merged-workbook';
			}

			// Remove existing extension if present
			fileName = fileName.replace(/\.(xlsx|xls|pdf|csv)$/i, '');
			// Add correct extension
			fileName = `${fileName}${fileExtension}`;

			// Handle the response - Excel API returns JSON with embedded base64 file
			let fileBuffer: Buffer;

			// Check for Buffer first to properly narrow TypeScript types
			if (Buffer.isBuffer(responseData)) {
				fileBuffer = responseData;
			} else if (typeof responseData === 'string') {
				fileBuffer = Buffer.from(responseData, 'base64');
			} else if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;

				if (response.document) {
					const document = response.document;

					if (typeof document === 'string') {
						fileBuffer = Buffer.from(document, 'base64');
					} else if (typeof document === 'object' && document !== null) {
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

						fileBuffer = Buffer.from(docContent, 'base64');
					} else {
						throw new Error(`Document field is neither string nor object: ${typeof document}`);
					}
				} else {
					const docContent =
						(response.docData as string) ||
						(response.content as string) ||
						(response.fileContent as string) ||
						(response.data as string);

					if (!docContent) {
						const keys = Object.keys(responseData).join(', ');
						throw new Error(`Excel API returned unexpected JSON structure. Available keys: ${keys}`);
					}

					fileBuffer = Buffer.from(docContent, 'base64');
				}
			} else {
				throw new Error(`Unexpected response format: ${typeof responseData}`);
			}

			// Validate the response contains data
			if (!fileBuffer || fileBuffer.length < 100) {
				throw new Error('Invalid response from API. The file appears to be too small or corrupted.');
			}

			// Create binary data for output
			const binaryData = await this.helpers.prepareBinaryData(
				fileBuffer,
				fileName,
				mimeType,
			);

			// Determine the binary data name
			const binaryDataKey = binaryDataName || 'data';

			// Build summary
			const filesSummary = documents.map((doc) => ({
				filename: doc.Filename,
				sortPosition: doc.SortPosition,
				worksheets: doc.WorksheetsToMerge || 'all',
			}));

			return [
				{
					json: {
						fileName,
						fileSize: fileBuffer.length,
						success: true,
						outputFormat,
						filesCount: documents.length,
						filesMerged: filesSummary,
						message: `Successfully merged ${documents.length} Excel file(s) into ${outputFormat} format`,
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				},
			];
		}

		throw new NodeOperationError(this.getNode(), 'No response data received from PDF4ME API');
	} catch (error) {
		// If it's already a NodeOperationError or NodeApiError, re-throw as-is
		if (error instanceof NodeOperationError || error instanceof NodeApiError) {
			throw error;
		}
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new NodeOperationError(this.getNode(), `Merge Excel files failed: ${errorMessage}`);
	}
}


