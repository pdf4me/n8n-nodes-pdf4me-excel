import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	ILoadOptionsFunctions,
	JsonObject,
	IHttpRequestMethods,
	IHttpRequestOptions,
	INode,
} from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';

export async function pdf4meApiRequest(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions,
	url: string,
	body: IDataObject = {},
	method: IHttpRequestMethods = 'POST',
	qs: IDataObject = {},
	option: IDataObject = {},
): Promise<Buffer | IDataObject> {
	// Determine if this is a JSON response operation (AI processing endpoints)
	const isJsonResponse = url.includes('/ProcessInvoice') || url.includes('/ProcessHealthCard') ||
		url.includes('/ProcessContract') || url.includes('/ParseDocument') ||
		url.includes('/ClassifyDocument');

	let options: IHttpRequestOptions = {
		baseURL: 'https://api.pdf4me.com',
		url: url,
		headers: {
			'Content-Type': 'application/json',
		},
		method,
		qs,
		body,
		json: isJsonResponse, // Parse as JSON for AI processing operations
		encoding: isJsonResponse ? undefined : 'arraybuffer' as const, // Use default encoding for JSON, arraybuffer for binary
		returnFullResponse: true, // Need full response to check status
		ignoreHttpStatusErrors: true, // Don't throw on non-2xx status codes
	};
	options = Object.assign({}, options, option);
	if (Object.keys(options.body as IDataObject).length === 0) {
		delete options.body;
	}

	try {
		const response = await this.helpers.httpRequestWithAuthentication.call(this, 'pdf4meExcelApi', {
			url: `${options.baseURL}${options.url}`,
			method: options.method,
			headers: options.headers,
			body: options.body,
			qs: options.qs,
			encoding: isJsonResponse ? undefined : 'arraybuffer' as const,
			// SSL validation is handled by n8n's httpRequestWithAuthentication
			returnFullResponse: options.returnFullResponse,
			json: options.json,
		});

		// Check if response is successful
		if (response.statusCode === 200) {
			// For JSON responses (AI processing), return the parsed JSON directly
			if (isJsonResponse) {
				return response.body; // Already parsed when json: true is set
			}

			// For binary responses, return binary content
			if (response.body instanceof Buffer) {
				return response.body;
			} else if (typeof response.body === 'string') {
				// If it's a string, it might be an error message
				if (response.body.length < 100) {
					throw new Error(`API returned error message: ${response.body}`);
				}
				// Try to convert from base64 if it's a long string
				try {
					return Buffer.from(response.body, 'base64');
				} catch (error) {
					throw new Error(`API returned unexpected string response: ${response.body.substring(0, 100)}...`);
				}
			} else {
				return Buffer.from(response.body, 'binary');
			}
		} else {
			// Error response - try to parse as JSON for error details
			let errorMessage = `HTTP ${response.statusCode}`;
			try {
				const errorJson = JSON.parse(response.body);
				errorMessage = errorJson.message || errorJson.error || errorJson.detail || errorMessage;
			} catch {
				errorMessage = `${errorMessage}: ${response.body}`;
			}
			throw new Error(errorMessage);
		}
	} catch (error) {
		throw new NodeApiError(this.getNode(), error as JsonObject);
	}
}

// Removed n8nSleep and all artificial delay logic to comply with n8n community guidelines.

// Delay function using PDF4ME's DelayAsync endpoint
async function delayAsync(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions,
): Promise<void> {
	await this.helpers.httpRequestWithAuthentication.call(this, 'pdf4meExcelApi', {
		url: 'https://api.pdf4me.com/api/v2/AddDelay',
		method: 'GET',
		returnFullResponse: true,
		ignoreHttpStatusErrors: true,
	});
}

export async function pdf4meAsyncRequest(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions,
	url: string,
	body: IDataObject = {},
	method: IHttpRequestMethods = 'POST',
	qs: IDataObject = {},
	option: IDataObject = {},
): Promise<Buffer | IDataObject> {
	// Use the body as-is without modifying it
	const asyncBody = body;

	// Determine if this is a JSON response operation
	// These operations return JSON with embedded data instead of raw binary
	const isJsonResponse =
		// Image operations
		url.includes('/CreateImages') || url.includes('/CreateImagesFromPdf') ||
		url.includes('/GetImageMetadata') ||
		// AI processing operations
		url.includes('/ProcessInvoice') || url.includes('/ProcessHealthCard') ||
		url.includes('/ProcessContract') || url.includes('/ParseDocument') ||
		url.includes('/ClassifyDocument') ||
		// Document operations
		url.includes('/GetTrackingChangesInWord') ||
		// Extraction operations
		url.includes('/ExtractResources') || url.includes('/ExtractPdfFormData') ||
		url.includes('/GetPdfMetadata') || url.includes('/ExtractTextByExpression') ||
		url.includes('/ExtractAttachmentFromPdf') || url.includes('/ExtractTableFromPdf') ||
		// Excel operations - return JSON with base64 encoded file
		url.includes('/ApiV2Excel/');

	let options: IHttpRequestOptions = {
		baseURL: 'https://api.pdf4me.com',
		url: url,
		headers: {
			'Content-Type': 'application/json',
		},
		method,
		qs,
		body: asyncBody,
		json: isJsonResponse, // Parse as JSON for operations that return structured data
		returnFullResponse: true, // Need full response to get headers
		ignoreHttpStatusErrors: true, // Don't throw on non-2xx status codes
		timeout: 1000023, // 60 second timeout for initial request (increased from 30s)
	};
	options = Object.assign({}, options, option);

	try {
		// Make initial request
		const response = await this.helpers.httpRequestWithAuthentication.call(this, 'pdf4meExcelApi', {
			url: `${options.baseURL}${options.url}`,
			method: options.method,
			headers: options.headers,
			body: options.body,
			qs: options.qs,
			encoding: isJsonResponse ? undefined : 'arraybuffer' as const,
			// SSL validation is handled by n8n's httpRequestWithAuthentication
			returnFullResponse: options.returnFullResponse,
			json: options.json,
			timeout: options.timeout,
		});

		if (response.statusCode === 200) {
			// Immediate success
			if (isJsonResponse) {
				return response.body; // Already parsed when json: true is set
			} else {
				// Handle binary response
				if (response.body instanceof Buffer) {
					return response.body;
				} else if (typeof response.body === 'string') {
					if (response.body.length < 100) {
						throw new Error(`API returned error message: ${response.body}`);
					}
					try {
						return Buffer.from(response.body, 'base64');
					} catch {
						throw new Error(`API returned unexpected string response: ${response.body.substring(0, 100)}...`);
					}
				} else {
					return Buffer.from(response.body, 'binary');
				}
			}
		} else if (response.statusCode === 202) {
			// Async processing - always start polling when API returns 202
			const locationUrl = response.headers.headers?.location || response.headers.location;
			if (!locationUrl) {
				throw new Error('No polling URL found in response');
			}

			// Start polling immediately when API returns 202
			// Poll the location URL until completion
			return await pollForCompletion.call(this, locationUrl, isJsonResponse);
		} else {
			let errorMessage = `API Error: ${response.statusCode}`;
			try {
				if (typeof response.body === 'string') {
					const errorJson = JSON.parse(response.body);
					errorMessage = errorJson.message || errorJson.error || errorJson.detail || errorMessage;
				} else {
					errorMessage = `${errorMessage}: ${response.body}`;
				}
			} catch {
				errorMessage = `${errorMessage}: ${response.body}`;
			}
			throw new Error(errorMessage);
		}
	} catch (error) {
		throw new NodeApiError(this.getNode(), error as JsonObject);
	}
}

export function sanitizeProfiles(data: IDataObject, node?: INode): void {
	// Convert profiles to a trimmed string (or empty string if not provided)
	const profilesValue = data.profiles ? String(data.profiles).trim() : '';

	// If the profiles field is empty, remove it from the payload
	if (!profilesValue) {
		delete data.profiles;
		return;
	}

	try {
		// Wrap profiles in curly braces if they are not already
		let sanitized = profilesValue;
		if (!sanitized.startsWith('{')) {
			sanitized = `{ ${sanitized}`;
		}
		if (!sanitized.endsWith('}')) {
			sanitized = `${sanitized} }`;
		}
		data.profiles = sanitized;
	} catch (error) {
		const errorMessage = 'Invalid JSON in Profiles. Check https://dev.pdf4me.com/ or contact support@pdf4me.com for help. ' +
			(error as Error).message;
		if (node) {
			throw new NodeOperationError(node, errorMessage);
		}
		throw new Error(errorMessage);
	}
}

/**
 * ActionConstants provides a mapping of supported Excel operations for the PDF4me Excel node.
 */
export const ActionConstants = {
	AddTextHeaderFooterToExcel: 'Add Text Header Footer To Excel',
	AddImageHeaderFooterToExcel: 'Add Image Header Footer To Excel',
	RemoveHeaderFooterToExcel: 'Remove Header Footer To Excel',
	AddTextWatermarkToExcel: 'Add Text Watermark To Excel',
	RemoveWatermarkFromExcel: 'Remove Watermark From Excel',
	FindReplaceTextInExcel: 'Find Replace Text In Excel',
	UpdateRowsToExcel: 'Update Rows To Excel',
	AddRowsToExcel: 'Add Rows To Excel',
	ExcelExtractRows: 'Excel Extract Rows',
	DeleteRowsFromExcel: 'Delete Rows From Excel',
	DeleteWorksheetFromExcel: 'Delete Worksheet From Excel',
	ExtractWorksheetFromExcel: 'Extract Worksheet From Excel',
	SecureExcelFile: 'Secure Excel File',
	UnlockExcelFile: 'Unlock Excel File',
	MergeExcelFiles: 'Merge Excel Files',
	MergeRowsInExcel: 'Merge Rows In Excel',
	ParseCsvToExcel: 'Parse CSV To JSON',
};

/**
 * Poll the PDF4me API for async operation completion
 * Used for operations that return 202 (Accepted) status initially
 *
 * @param locationUrl - The polling URL returned in the Location header
 * @param isJsonResponse - Whether the operation returns JSON (true) or binary data (false)
 * @param maxRetries - Maximum number of polling attempts (default: 9000)
 * @returns The final result (JSON object or Buffer)
 */
async function pollForCompletion(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions,
	locationUrl: string,
	isJsonResponse: boolean,
	maxRetries: number = 9000,
): Promise<Buffer | IDataObject> {
	let retryCount = 0;

	while (retryCount < maxRetries) {
		try {
			// Make polling request with appropriate encoding based on response type
			const pollResponse = await this.helpers.httpRequestWithAuthentication.call(this, 'pdf4meExcelApi', {
				url: locationUrl,
				method: 'GET',
				encoding: isJsonResponse ? undefined : 'arraybuffer' as const,
				returnFullResponse: true,
				json: isJsonResponse, // Parse JSON for operations like Excel, AI processing, metadata extraction
				ignoreHttpStatusErrors: true,
			});

			if (pollResponse.statusCode === 200) {
				// Success - return the final result
				if (isJsonResponse) {
					// For JSON responses (Excel, AI operations, etc.), return the parsed object
					return pollResponse.body; // Already parsed when json: true is set
				} else {
					// For binary responses (PDFs, images, etc.), return as Buffer
					// Handle binary response
					if (pollResponse.body instanceof Buffer) {
						return pollResponse.body;
					} else if (typeof pollResponse.body === 'string') {
						if (pollResponse.body.length < 100) {
							throw new Error(`API returned error message: ${pollResponse.body}`);
						}
						try {
							return Buffer.from(pollResponse.body, 'base64');
						} catch {
							throw new Error(`API returned unexpected string response: ${pollResponse.body.substring(0, 100)}...`);
						}
					} else {
						return Buffer.from(pollResponse.body, 'binary');
					}
				}
			} else if (pollResponse.statusCode === 202) {
				// Still processing, continue polling with 10 second backoff
				retryCount++;
				// Use PDF4ME's DelayAsync endpoint for 10 second delay
				await delayAsync.call(this);
				continue;
			} else if (pollResponse.statusCode === 404) {
				// Job not found or expired
				throw new Error('Processing job not found or expired. The document processing may have timed out.');
			} else {
				// Other error
				let errorMessage = `Polling failed with status ${pollResponse.statusCode}`;
				try {
					if (typeof pollResponse.body === 'string') {
						const errorJson = JSON.parse(pollResponse.body);
						errorMessage = errorJson.message || errorJson.error || errorJson.detail || errorMessage;
					} else {
						errorMessage = `${errorMessage}: ${pollResponse.body}`;
					}
				} catch {
					errorMessage = `${errorMessage}: ${pollResponse.body}`;
				}
				throw new Error(errorMessage);
			}
		} catch (error) {
			// If it's a network error, retry with minimal backoff
			if (error.message.includes('ENOTFOUND') || error.message.includes('ECONNRESET') || error.message.includes('timeout')) {
				retryCount++;
				if (retryCount >= maxRetries) {
					throw new Error(`Network error during polling after ${maxRetries} attempts: ${error.message}`);
				}
				// Use PDF4ME's DelayAsync endpoint for 10 second delay on network errors
				await delayAsync.call(this);
				continue;
			}
			// For other errors, throw immediately
			throw error;
		}
	}

	throw new Error(`Document processing timed out after ${maxRetries} polling attempts. The operation may still be processing on the server.`);
}
