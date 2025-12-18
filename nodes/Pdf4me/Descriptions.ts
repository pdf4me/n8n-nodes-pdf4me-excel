/* eslint-disable n8n-nodes-base/node-filename-against-convention, n8n-nodes-base/node-param-default-missing */
import { INodeTypeDescription, NodeConnectionType } from 'n8n-workflow';
import * as addTextHeaderFooterToExcel from './actions/addTextHeaderFooterToExcel';
import * as addImageHeaderFooterToExcel from './actions/addImageHeaderFooterToExcel';
import * as removeHeaderFooterToExcel from './actions/removeHeaderFooterToExcel';
import * as addTextWatermarkToExcel from './actions/addTextWatermarkToExcel';
import * as removeWatermarkFromExcel from './actions/removeWatermarkFromExcel';
import * as findReplaceTextInExcel from './actions/findReplaceTextInExcel';
import * as updateRowsToExcel from './actions/updateRowsToExcel';
import * as addRowsToExcel from './actions/addRowsToExcel';
import * as excelExtractRows from './actions/excelExtractRows';
import * as deleteRowsFromExcel from './actions/deleteRowsFromExcel';
import * as deleteWorksheetFromExcel from './actions/deleteWorksheetFromExcel';
import * as extractWorksheetFromExcel from './actions/extractWorksheetFromExcel';
import * as secureExcelFile from './actions/secureExcelFile';
import * as unlockExcelFile from './actions/unlockExcelFile';
import * as mergeExcelFiles from './actions/mergeExcelFiles';
import * as mergeRowsInExcel from './actions/mergeRowsInExcel';
import * as parseCsvToExcel from './actions/parseCsvToJson';
import { ActionConstants } from './GenericFunctions';

export const descriptions: INodeTypeDescription = {
	displayName: 'PDF4me Excel',
	name: 'pdf4meExcel',
	description: ' Revolutionize Excel workflows with intelligent automation: dynamically manipulate spreadsheets, merge data seamlessly, secure files with enterprise-grade encryption, and extract insights through advanced row operations - all powered by cutting-edge PDF4me API technology',
	defaults: {
		name: 'PDF4me Excel',
	},
	group: ['transform'],
	icon: 'file:300.svg',
	inputs: [NodeConnectionType.Main],
	outputs: [NodeConnectionType.Main],
	credentials: [
		{
			name: 'pdf4meExcelApi',
			required: true,
		},
	], // eslint-disable-line n8n-nodes-base/node-param-default-missing
	properties: [
		{
			displayName: 'Operation',
			name: 'operation',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: 'Add Text Header/Footer',
					description: 'Add customizable text headers and footers to Excel worksheets with alignment, font size, and color options',
					value: ActionConstants.AddTextHeaderFooterToExcel,
					action: 'Add Text Header/Footer ',
				},
				{
					name: 'Add Image Header/Footer',
					description: 'Add image-based headers and footers to Excel worksheets with position and margin options',
					value: ActionConstants.AddImageHeaderFooterToExcel,
					action: 'Add Image Header/Footer ',
				},
				{
					name: 'Remove Header/Footer',
					description: 'Remove headers and footers from Excel worksheets with selective removal options',
					value: ActionConstants.RemoveHeaderFooterToExcel,
					action: 'Remove Header/Footer ',
				},
				{
					name: 'Add Text Watermark',
					description: 'Add customizable text watermark to Excel worksheets with font, color, rotation, and transparency options',
					value: ActionConstants.AddTextWatermarkToExcel,
					action: 'Add text watermark ',
				},
				{
					name: 'Remove Watermark',
					description: 'Remove watermarks from Excel worksheets with selective removal options',
					value: ActionConstants.RemoveWatermarkFromExcel,
					action: 'Remove watermark',
				},
				{
					name: 'Find and Replace Text',
					description: 'Find and replace text in Excel worksheets with custom formatting and multiple operations',
					value: ActionConstants.FindReplaceTextInExcel,
					action: 'Find and replace text ',
				},
				{
					name: 'Update Rows',
					description: 'Update rows in Excel worksheets with JSON data and custom formatting options',
					value: ActionConstants.UpdateRowsToExcel,
					action: 'Update rows ',
				},
				{
					name: 'Add Rows',
					description: 'Add new rows to Excel worksheets with JSON data and custom formatting options',
					value: ActionConstants.AddRowsToExcel,
					action: 'Add rows to Excel worksheet',
				},
				{
					name: 'Extract Rows',
					description: 'Extract rows from Excel worksheets and return as JSON data with customizable extraction options',
					value: ActionConstants.ExcelExtractRows,
					action: 'Extract rows from Excel worksheet',
				},
				{
					name: 'Delete Rows',
					description: 'Delete specified row ranges from Excel worksheets',
					value: ActionConstants.DeleteRowsFromExcel,
					action: 'Delete rows from Excel worksheet',
				},
				{
					name: 'Delete Worksheet',
					description: 'Delete entire worksheets from Excel files by name or index',
					value: ActionConstants.DeleteWorksheetFromExcel,
					action: 'Delete worksheet from Excel file',
				},
				{
					name: 'Extract Worksheet',
					description: 'Extract (keep only) specific worksheets from Excel files by name or index',
					value: ActionConstants.ExtractWorksheetFromExcel,
					action: 'Extract worksheet from Excel file',
				},
				{
					name: 'Secure File',
					description: 'Password protect and secure Excel files with file encryption, workbook, and worksheet protection',
					value: ActionConstants.SecureExcelFile,
					action: 'Secure Excel file ',
				},
				{
					name: 'Unlock File',
					description: 'Remove password protection and unlock secured Excel files',
					value: ActionConstants.UnlockExcelFile,
					action: 'Unlock Excel file and remove protection',
				},
				{
					name: 'Merge Files',
					description: 'Merge multiple Excel files into a single workbook with customizable output format',
					value: ActionConstants.MergeExcelFiles,
					action: 'Merge multiple Excel files',
				},
				{
					name: 'Merge Rows',
					description: 'Merge rows within an Excel file based on key columns across specified worksheets',
					value: ActionConstants.MergeRowsInExcel,
					action: 'Merge rows in Excel file',
				},
				{
					name: 'Parse CSV to JSON',
					description: 'Convert CSV files to JSON format with customizable delimiter and column headers',
					value: ActionConstants.ParseCsvToExcel,
					action: 'Parse CSV and convert to JSON',
				},
			],
			default: ActionConstants.AddTextHeaderFooterToExcel,
		},
		...addTextHeaderFooterToExcel.description,
		...addImageHeaderFooterToExcel.description,
		...removeHeaderFooterToExcel.description,
		...addTextWatermarkToExcel.description,
		...removeWatermarkFromExcel.description,
		...findReplaceTextInExcel.description,
		...updateRowsToExcel.description,
		...addRowsToExcel.description,
		...excelExtractRows.description,
		...deleteRowsFromExcel.description,
		...deleteWorksheetFromExcel.description,
		...extractWorksheetFromExcel.description,
		...secureExcelFile.description,
		...unlockExcelFile.description,
		...mergeExcelFiles.description,
		...mergeRowsInExcel.description,
		...parseCsvToExcel.description,
	],
	subtitle: '={{$parameter["operation"]}}',
	version: 1,
};
