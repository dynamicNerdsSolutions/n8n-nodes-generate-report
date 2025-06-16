import {
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	NodeConnectionType,
	NodeOperationError,
	IExecuteFunctions,
	IDataObject
} from 'n8n-workflow';

import * as iconv from 'iconv-lite';
iconv.encodingExists('utf8');

import { TemplateData, TemplateHandler } from 'easy-template-x';

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);



export class GenerateReport implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Generate Report',
		name: 'generateReport',
		icon: 'file:report_template.svg',
		group: ['transform'],
		version: 1,
		description: 'Generate a report from a DocX Template and JSON data.',
		defaults: {
			name: 'Generate Report',
		},
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		properties: [
			{
				displayName: 'Template Key',
				name: 'sourceKey',
				type: 'string',
				default: 'template',
				required: true,
				placeholder: 'template',
				description: 'The name of the binary key to get the template from. It is also possible to define deep keys by using dot-notation like for example: "level1.level2.currentKey".',
			},
			{
				displayName: 'Output Key',
				name: 'destinationKey',
				type: 'string',
				default: 'report',
				required: true,
				placeholder: 'report',
				description: 'The name of the binary key to copy data to. It is also possible to define deep keys by using dot-notation like for example: "level1.level2.newKey".',
			},
			{
				displayName: 'Input Data',
				name: 'data',
				type: 'string',
				default: '',
				required: true,
				placeholder: 'data',
				description: 'Data to use to fill the report. Insert as string, so please use JSON.stringify(data) if needed.',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				default: 'Report',
				required: true,
				placeholder: 'Report',
				description: 'File name of the output document, enter without extension',
			},
			{
				displayName: 'Tag Delimiters',
				name: 'tagDelimiters',
				type: 'collection',
				default: {
					tagStart: '{{',
					tagEnd: '}}',
					containerTagOpen: '#',
					containerTagClose: '/'
				},
				options: [
					{
						displayName: 'Tag Start Delimiters',
						name: 'tagStart',
						type: 'string',
						default: '{{',
					},
					{
						displayName: 'Tag End Delimiters',
						name: 'tagEnd',
						type: 'string',
						default: '}}',
					},
					{
						displayName: 'Container Tag Open',
						name: 'containerTagOpen',
						type: 'string',
						default: '#',
					},
					{
						displayName: 'Container Tag Close',
						name: 'containerTagClose',
						type: 'string',
						default: '/',
					}
				]
			}
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {

		const items = this.getInputData();

		const returnData: INodeExecutionData[] = [];

		let item: INodeExecutionData;
		let newItem: INodeExecutionData;

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			item = items[itemIndex];
			newItem = {
				json: {},
				binary: {},
			};

			const sourceKey = this.getNodeParameter('sourceKey', itemIndex) as string;
			const destinationKey = this.getNodeParameter('destinationKey', itemIndex) as string;
			const data = this.getNodeParameter('data', itemIndex) as string;
			const outputFileName = this.getNodeParameter('outputFileName', itemIndex) as string;
			const tagDelimiters = this.getNodeParameter('tagDelimiters', itemIndex) as IDataObject;
			this.logger.debug(JSON.stringify(tagDelimiters));
			const tagStart = tagDelimiters.tagStart as string;
			const tagEnd = tagDelimiters.tagEnd as string;
			// const containerTagOpen = tagDelimiters.containerTagOpen as string;
			// const containerTagClose = tagDelimiters.containerTagClose as string;
			let templateData = {} as TemplateData;
			try{
				templateData = JSON.parse(data) as TemplateData;
			}
			catch(err){
				throw new NodeOperationError(this.getNode(), 'Something went wrong while parsing the template data.' + err as string);
			}

			if (item.binary === undefined) {
				throw new NodeOperationError(this.getNode(), 'No binary data exists on item!');
			}

			const binaryDataBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, sourceKey);

			const handler = new TemplateHandler({
				delimiters: {
						tagStart: tagStart,
						tagEnd: tagEnd,
						containerTagOpen: "#",
						containerTagClose: "/"
				},
			});

			try{

				const doc = await handler.process(binaryDataBuffer, templateData);


				newItem.binary![destinationKey] = await this.helpers.prepareBinaryData(doc, `${outputFileName}.docx`);


			}
			catch(err){
				throw new NodeOperationError(this.getNode(), 'Something went wrong creating the report. ' + err as string);
			}





			returnData.push(newItem);
		}

		return [returnData];
	}
}
