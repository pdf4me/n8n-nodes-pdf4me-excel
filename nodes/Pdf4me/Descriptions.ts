/* eslint-disable n8n-nodes-base/node-filename-against-convention, n8n-nodes-base/node-param-default-missing */
import { INodeTypeDescription, NodeConnectionType } from 'n8n-workflow';
import * as addTextHeaderFooterToExcel from './actions/addTextHeaderFooterToExcel';
import { ActionConstants } from './GenericFunctions';

export const descriptions: INodeTypeDescription = {
	displayName: 'PDF4me Excel',
	name: 'pdf4meExcel',
	description: 'Process Excel files: add customizable headers and footers to worksheets with professional styling and formatting using PDF4me API',
	defaults: {
		name: 'PDF4me Excel',
	},
	group: ['transform'],
	icon: 'file:300.svg',
	inputs: [NodeConnectionType.Main],
	outputs: [NodeConnectionType.Main],
	credentials: [
		{
			name: 'pdf4meApi',
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
					action: 'Add text header and footer to Excel worksheet',
				},
			],
			default: ActionConstants.AddTextHeaderFooterToExcel,
		},
		...addTextHeaderFooterToExcel.description,
	],
	subtitle: '={{$parameter["operation"]}}',
	version: 1,
};
