import {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	INodeTypeBaseDescription,
	INodeExecutionData,
} from 'n8n-workflow';

import { descriptions } from './Descriptions';
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

export class Pdf4meExcel implements INodeType {
	description: INodeTypeDescription;

	constructor(baseDescription: INodeTypeBaseDescription) {
		this.description = {
			...baseDescription,
			...descriptions,
		};
	}

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const operationResult: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const action = this.getNodeParameter('operation', i);

			try {
				let results: INodeExecutionData[] = [];
				if (action === ActionConstants.AddTextHeaderFooterToExcel) {
					results = await addTextHeaderFooterToExcel.execute.call(this, i);
				} else if (action === ActionConstants.AddImageHeaderFooterToExcel) {
					results = await addImageHeaderFooterToExcel.execute.call(this, i);
				} else if (action === ActionConstants.RemoveHeaderFooterToExcel) {
					results = await removeHeaderFooterToExcel.execute.call(this, i);
				} else if (action === ActionConstants.AddTextWatermarkToExcel) {
					results = await addTextWatermarkToExcel.execute.call(this, i);
				} else if (action === ActionConstants.RemoveWatermarkFromExcel) {
					results = await removeWatermarkFromExcel.execute.call(this, i);
				} else if (action === ActionConstants.FindReplaceTextInExcel) {
					results = await findReplaceTextInExcel.execute.call(this, i);
				} else if (action === ActionConstants.UpdateRowsToExcel) {
					results = await updateRowsToExcel.execute.call(this, i);
				} else if (action === ActionConstants.AddRowsToExcel) {
					results = await addRowsToExcel.execute.call(this, i);
				} else if (action === ActionConstants.ExcelExtractRows) {
					results = await excelExtractRows.execute.call(this, i);
				} else if (action === ActionConstants.DeleteRowsFromExcel) {
					results = await deleteRowsFromExcel.execute.call(this, i);
				} else if (action === ActionConstants.DeleteWorksheetFromExcel) {
					results = await deleteWorksheetFromExcel.execute.call(this, i);
				} else if (action === ActionConstants.ExtractWorksheetFromExcel) {
					results = await extractWorksheetFromExcel.execute.call(this, i);
				} else if (action === ActionConstants.SecureExcelFile) {
					results = await secureExcelFile.execute.call(this, i);
				} else if (action === ActionConstants.UnlockExcelFile) {
					results = await unlockExcelFile.execute.call(this, i);
				} else if (action === ActionConstants.MergeExcelFiles) {
					results = await mergeExcelFiles.execute.call(this, i);
				} else if (action === ActionConstants.MergeRowsInExcel) {
					results = await mergeRowsInExcel.execute.call(this, i);
				} else if (action === ActionConstants.ParseCsvToExcel) {
					results = await parseCsvToExcel.execute.call(this, i);
				}
				if (!results || results.length === 0) {
					operationResult.push({
						json: {},
						pairedItem: { item: i },
					});
				} else {
					// Add pairedItem to each result to maintain data lineage
					for (const result of results) {
						operationResult.push({
							...result,
							pairedItem: { item: i },
						});
					}
				}
			} catch (err) {
				if (this.continueOnFail()) {
					operationResult.push({
						json: this.getInputData(i)[0].json,
						error: err,
						pairedItem: { item: i },
					});
				} else {
					throw err;
				}
			}
		}

		return [operationResult];
	}
}
