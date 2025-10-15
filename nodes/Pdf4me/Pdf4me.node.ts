import {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	INodeTypeBaseDescription,
	INodeExecutionData,
} from 'n8n-workflow';

import { descriptions } from './Descriptions';
import * as addTextHeaderFooterToExcel from './actions/addTextHeaderFooterToExcel';
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
