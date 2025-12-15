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
				operation: [ActionConstants.MergeRowsInExcel],
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
				operation: [ActionConstants.MergeRowsInExcel],
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
				operation: [ActionConstants.MergeRowsInExcel],
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
				operation: [ActionConstants.MergeRowsInExcel],
				inputDataType: ['url'],
			},
		},
	},
	// === MERGE SETTINGS ===
	{
		displayName: 'Worksheet Numbers',
		name: 'worksheetNumbers',
		type: 'string',
		default: '',
		description: 'Comma-separated worksheet numbers (1-based) to merge rows from. Leave empty to merge all worksheets. Example: "1,2,3" merges first three worksheets.',
		placeholder: '1,2,3',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeRowsInExcel],
			},
		},
	},
	{
		displayName: 'Merge Key Columns',
		name: 'mergeKeyColumns',
		type: 'string',
		default: '',
		description: 'Comma-separated column letters or numbers (1-based) to use as merge keys. Rows with matching values in these columns will be merged. Example: "A,B" or "1,2".',
		placeholder: 'A,B / 1,2',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeRowsInExcel],
			},
		},
	},
	{
		displayName: 'Output Format',
		name: 'outputFormat',
		type: 'options',
		default: 'XLSX',
		description: 'Format for the merged output file',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeRowsInExcel],
			},
		},
		options: [
			{
				name: 'XLSX (Excel 2007+)',
				value: 'XLSX',
				description: 'Modern Excel format (.xlsx)',
			},
			{
				name: 'XLSB (Excel Binary)',
				value: 'XLSB',
				description: 'Excel Binary Workbook format (.xlsb)',
			},
			{
				name: 'XLS (Excel 97-2003)',
				value: 'XLS',
				description: 'Legacy Excel format (.xls)',
			},
			{
				name: 'CSV',
				value: 'CSV',
				description: 'Comma-Separated Values (.csv)',
			},
		],
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'excel_merged_rows.xlsx',
		description: 'Name for the processed Excel file (will have rows merged)',
		placeholder: 'output.xlsx',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeRowsInExcel],
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
				operation: [ActionConstants.MergeRowsInExcel],
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
				operation: [ActionConstants.MergeRowsInExcel],
			},
		},
	},
];

/**
 * Merge rows in Excel files using PDF4Me API
 * Process: Read Excel file → Encode to base64 → Send API request → Poll for completion → Save Excel file
 * Merges rows across specified worksheets based on key columns
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get merge parameters
		const worksheetNumbers = this.getNodeParameter('worksheetNumbers', index, '') as string;
		const mergeKeyColumns = this.getNodeParameter('mergeKeyColumns', index, '') as string;
		const outputFormat = this.getNodeParameter('outputFormat', index, 'XLSX') as string;

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

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
			MergeRowsToExcelAction: {
				WorksheetNumbers: worksheetNumbers,
				MergeKeyColumns: mergeKeyColumns,
				OutputFormat: outputFormat,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Excel/ExcelMergeRows',
			body,
		);

		if (responseData) {
			// Determine file extension and MIME type based on output format
			let fileExtension = '.xlsx';
			let mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

			switch (outputFormat.toUpperCase()) {
			case 'XLSX':
				fileExtension = '.xlsx';
				mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
				break;
			case 'XLSB':
				fileExtension = '.xlsb';
				mimeType = 'application/vnd.ms-excel.sheet.binary.macroEnabled.12';
				break;
			case 'XLS':
				fileExtension = '.xls';
				mimeType = 'application/vnd.ms-excel';
				break;
			case 'CSV':
				fileExtension = '.csv';
				mimeType = 'text/csv';
				break;
			}

			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'excel_merged_rows';
				fileName = `${baseName}${fileExtension}`;
			}

			// Ensure correct extension
			fileName = fileName.replace(/\.(xlsx|xlsb|xls|csv)$/i, '');
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

			return [
				{
					json: {
						fileName,
						fileSize: fileBuffer.length,
						success: true,
						originalFileName,
						worksheetNumbers: worksheetNumbers || 'all worksheets',
						mergeKeyColumns: mergeKeyColumns || 'none specified',
						outputFormat,
						message: `Successfully merged rows in Excel file${worksheetNumbers ? ` from worksheets ${worksheetNumbers}` : ''}${mergeKeyColumns ? ` using key columns ${mergeKeyColumns}` : ''}`,
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
		throw new NodeOperationError(this.getNode(), `Merge rows in Excel failed: ${errorMessage}`);
	}
}


