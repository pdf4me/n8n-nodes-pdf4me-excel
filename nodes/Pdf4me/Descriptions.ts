/* eslint-disable n8n-nodes-base/node-filename-against-convention, n8n-nodes-base/node-param-default-missing */
import { INodeTypeDescription, NodeConnectionType } from 'n8n-workflow';
import * as addTextHeaderFooterToExcel from './actions/addTextHeaderFooterToExcel';
import * as addImageHeaderFooterToExcel from './actions/addImageHeaderFooterToExcel';
import * as removeHeaderFooterToExcel from './actions/removeHeaderFooterToExcel';
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
				{
					name: 'Add Image Header/Footer',
					description: 'Add image-based headers and footers to Excel worksheets with position and margin options',
					value: ActionConstants.AddImageHeaderFooterToExcel,
					action: 'Add image header and footer to Excel worksheet',
				},
				{
					name: 'Remove Header/Footer',
					description: 'Remove headers and footers from Excel worksheets with selective removal options',
					value: ActionConstants.RemoveHeaderFooterToExcel,
					action: 'Remove header and footer from Excel worksheet',
				},
			],
			default: ActionConstants.AddTextHeaderFooterToExcel,
		},
		...addTextHeaderFooterToExcel.description,
		...addImageHeaderFooterToExcel.description,
		...removeHeaderFooterToExcel.description,
	],
	subtitle: '={{$parameter["operation"]}}',
	version: 1,
};
