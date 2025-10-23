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

export class Pdf4me implements INodeType {
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
				if (action === ActionConstants.AddTextHeaderFooterToExcel) {
					operationResult.push(...(await addTextHeaderFooterToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.AddImageHeaderFooterToExcel) {
					operationResult.push(...(await addImageHeaderFooterToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.RemoveHeaderFooterToExcel) {
					operationResult.push(...(await removeHeaderFooterToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.AddTextWatermarkToExcel) {
					operationResult.push(...(await addTextWatermarkToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.RemoveWatermarkFromExcel) {
					operationResult.push(...(await removeWatermarkFromExcel.execute.call(this, i)));
				} else if (action === ActionConstants.FindReplaceTextInExcel) {
					operationResult.push(...(await findReplaceTextInExcel.execute.call(this, i)));
				} else if (action === ActionConstants.UpdateRowsToExcel) {
					operationResult.push(...(await updateRowsToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.AddRowsToExcel) {
					operationResult.push(...(await addRowsToExcel.execute.call(this, i)));
				} else if (action === ActionConstants.ExcelExtractRows) {
					operationResult.push(...(await excelExtractRows.execute.call(this, i)));
				} else if (action === ActionConstants.DeleteRowsFromExcel) {
					operationResult.push(...(await deleteRowsFromExcel.execute.call(this, i)));
				} else if (action === ActionConstants.DeleteWorksheetFromExcel) {
					operationResult.push(...(await deleteWorksheetFromExcel.execute.call(this, i)));
				} else if (action === ActionConstants.ExtractWorksheetFromExcel) {
					operationResult.push(...(await extractWorksheetFromExcel.execute.call(this, i)));
				} else if (action === ActionConstants.SecureExcelFile) {
					operationResult.push(...(await secureExcelFile.execute.call(this, i)));
				} else if (action === ActionConstants.UnlockExcelFile) {
					operationResult.push(...(await unlockExcelFile.execute.call(this, i)));
				} else if (action === ActionConstants.MergeExcelFiles) {
					operationResult.push(...(await mergeExcelFiles.execute.call(this, i)));
				} else if (action === ActionConstants.MergeRowsInExcel) {
					operationResult.push(...(await mergeRowsInExcel.execute.call(this, i)));
				} else if (action === ActionConstants.ParseCsvToExcel) {
					operationResult.push(...(await parseCsvToExcel.execute.call(this, i)));
				}
			} catch (err) {
				if (this.continueOnFail()) {
					operationResult.push({ json: this.getInputData(i)[0].json, error: err });
				} else {
					throw err;
				}
			}
		}

		return [operationResult];
	}
}
