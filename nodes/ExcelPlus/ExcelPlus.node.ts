import type {
	IBinaryData,
	IDataObject,
	IExecuteFunctions,
	INode,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IPairedItemData,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';
// exceljs is inlined into the dist bundle by esbuild; n8n Cloud's lint rule against runtime deps does not apply to bundled code.
// eslint-disable-next-line @n8n/community-nodes/no-restricted-imports
import ExcelJS from 'exceljs';

const MAX_INPUTS = 20;

const numberOfInputsOptions = Array.from({ length: MAX_INPUTS }, (_, i) => {
	const value = i + 1;
	return { name: String(value), value };
});

const INVALID_SHEET_NAME_CHARS = /[[\]/\\*?:]/g;

function sanitizeSheetName(name: string): string {
	return name.replace(INVALID_SHEET_NAME_CHARS, '_').slice(0, 31) || 'Sheet';
}

function uniqueSheetName(node: INode, desired: string, used: Set<string>): string {
	const base = sanitizeSheetName(desired);
	if (!used.has(base)) return base;
	for (let counter = 1; counter < 1000; counter++) {
		const suffix = ` (${counter})`;
		const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
		if (!used.has(candidate)) return candidate;
	}
	throw new NodeOperationError(node, `Could not allocate a unique sheet name for "${desired}".`);
}

function stripExtension(fileName: string): string {
	const lastDot = fileName.lastIndexOf('.');
	return lastDot > 0 ? fileName.slice(0, lastDot) : fileName;
}

const INVALID_FILE_NAME_CHARS = /[\\/:*?"<>|]/g;

function sanitizeOutputFileName(raw: string): string {
	const cleaned = raw.replace(INVALID_FILE_NAME_CHARS, '_').replace(/^\.+/, '').trim();
	const base = cleaned === '' ? 'merged' : stripExtension(cleaned) || 'merged';
	return `${base}.xlsx`;
}

function buildPrefixedName(
	fileName: string | undefined,
	sheetName: string,
	fallbackPrefix: string,
): string {
	const prefixSource = fileName ? stripExtension(fileName) : fallbackPrefix;
	const prefix = prefixSource.trim() || fallbackPrefix;
	return `${prefix} - ${sheetName}`;
}

function parsePerInputNames(
	node: INode,
	raw: unknown,
	expectedCount: number,
): Record<string, string> {
	let parsed: unknown = raw;
	if (typeof raw === 'string') {
		const trimmed = raw.trim();
		if (trimmed === '') return {};
		try {
			parsed = JSON.parse(trimmed);
		} catch (error) {
			throw new NodeOperationError(
				node,
				`"Binary Field Names per Input" must be a valid JSON object: ${(error as Error).message}`,
			);
		}
	}

	if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
		throw new NodeOperationError(
			node,
			'"Binary Field Names per Input" must be a JSON object keyed by input number.',
		);
	}

	const result: Record<string, string> = {};
	for (const [key, value] of Object.entries(parsed as Record<string, unknown>)) {
		const n = Number(key);
		if (!Number.isInteger(n) || n < 1 || n > expectedCount) {
			throw new NodeOperationError(
				node,
				`Key "${key}" in "Binary Field Names per Input" is not a valid input number (expected 1-${expectedCount}).`,
			);
		}
		if (typeof value !== 'string' || value === '') {
			throw new NodeOperationError(node, `Binary field name for input ${n} must be a non-empty string.`);
		}
		result[String(n)] = value;
	}
	return result;
}

function copySheet(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet): void {
	source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
		const newRow = target.getRow(rowNumber);
		if (row.height !== undefined) newRow.height = row.height;
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const newCell = newRow.getCell(colNumber);
			newCell.value = cell.value;
			if (cell.style) newCell.style = { ...cell.style };
			if (cell.numFmt) newCell.numFmt = cell.numFmt;
		});
		newRow.commit();
	});

	source.columns?.forEach((col, idx) => {
		if (col?.width !== undefined) {
			target.getColumn(idx + 1).width = col.width;
		}
	});

	const merges = (source.model as { merges?: string[] }).merges;
	if (merges) {
		for (const range of merges) {
			try {
				target.mergeCells(range);
			} catch {
				// Skip invalid ranges so a single bad merge doesn't abort the whole copy.
			}
		}
	}

	if (source.views?.length) {
		target.views = source.views.map((view) => ({ ...view }));
	}
}

export class ExcelPlus implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel Plus',
		name: 'excelPlus',
		icon: { light: 'file:excelPlus.svg', dark: 'file:excelPlus.dark.svg' },
		group: ['transform'],
		version: 1,
		subtitle: '={{ $parameter["operation"] }}',
		description: 'Work with Excel files. Currently supports merging sheets from multiple workbooks.',
		defaults: {
			name: 'Excel Plus',
		},
		usableAsTool: true,
		inputs: `={{ ((p) => { const n = Math.min(Math.max(parseInt(p.numberOfInputs, 10) || 1, 1), ${MAX_INPUTS}); return Array.from({ length: n }, (_, i) => ({ type: 'main', displayName: 'Input ' + (i + 1) })); })($parameter) }}`,
		outputs: [NodeConnectionTypes.Main],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Merge Sheets',
						value: 'mergeSheets',
						description: 'Combine all sheets from multiple Excel files into a single workbook',
						action: 'Merge sheets from multiple workbooks',
					},
				],
				default: 'mergeSheets',
			},
			{
				displayName: 'Number of Inputs',
				name: 'numberOfInputs',
				type: 'options',
				noDataExpression: true,
				options: numberOfInputsOptions,
				default: 2,
				description:
					'How many input connections this node accepts. Reconnect upstream nodes after changing this value.',
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
					},
				},
			},
			{
				displayName: 'Binary Field Mode',
				name: 'binaryFieldMode',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Different per Input',
						value: 'different',
						description: 'Specify a binary property name for each input individually',
					},
					{
						name: 'Same Across All Inputs',
						value: 'same',
						description: 'Use the same binary property name on every input',
					},
				],
				default: 'same',
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
					},
				},
			},
			{
				displayName: 'Binary Field Name',
				name: 'binaryFieldName',
				type: 'string',
				default: 'data',
				required: true,
				placeholder: 'data',
				description: 'Name of the binary property on incoming items that holds the Excel file',
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
						binaryFieldMode: ['same'],
					},
				},
			},
			{
				displayName: 'Binary Field Names per Input',
				name: 'binaryFieldNamesPerInput',
				type: 'json',
				default: '{\n  "1": "data",\n  "2": "data"\n}',
				required: true,
				description:
					'JSON object mapping each input number (as a string key) to the binary property name to read on that input. Keys must be "1", "2", ... up to the number of inputs above.',
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
						binaryFieldMode: ['different'],
					},
				},
			},
			{
				displayName: 'Sheet Name Conflict',
				name: 'sheetNameConflict',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Auto-Suffix With Counter',
						value: 'suffix',
						description: 'Append " (1)", " (2)", ... to duplicate sheet names.',
					},
					{
						name: 'Prefix With File Name',
						value: 'prefix',
						description: 'Prefix every sheet with its source file name (no extension)',
					},
				],
				default: 'suffix',
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
					},
				},
			},
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				displayOptions: {
					show: {
						operation: ['mergeSheets'],
					},
				},
				options: [
					{
						displayName: 'Output File Name',
						name: 'outputFileName',
						type: 'string',
						default: 'merged.xlsx',
						placeholder: 'merged.xlsx',
						description:
							'File name for the merged workbook on the output binary. A ".xlsx" extension is added automatically if missing.',
					},
				],
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const operation = this.getNodeParameter('operation', 0) as string;

		if (operation !== 'mergeSheets') {
			throw new NodeOperationError(this.getNode(), `Unsupported operation: ${operation}.`);
		}

		const numberOfInputs = Math.min(
			Math.max(this.getNodeParameter('numberOfInputs', 0) as number, 1),
			MAX_INPUTS,
		);
		const binaryFieldMode = this.getNodeParameter('binaryFieldMode', 0) as 'same' | 'different';
		const sheetNameConflict = this.getNodeParameter('sheetNameConflict', 0) as 'suffix' | 'prefix';
		const options = this.getNodeParameter('options', 0, {}) as {
			outputFileName?: string;
		};
		const outputFileName = sanitizeOutputFileName(options.outputFileName ?? 'merged.xlsx');

		const node = this.getNode();
		let resolveBinaryName: (inputIndex: number) => string;
		if (binaryFieldMode === 'same') {
			const sharedName = (this.getNodeParameter('binaryFieldName', 0) as string).trim();
			if (!sharedName) {
				throw new NodeOperationError(node, '"Binary Field Name" cannot be empty.');
			}
			resolveBinaryName = () => sharedName;
		} else {
			const raw = this.getNodeParameter('binaryFieldNamesPerInput', 0);
			const mapping = parsePerInputNames(node, raw, numberOfInputs);
			resolveBinaryName = (inputIndex) => {
				const name = mapping[String(inputIndex + 1)];
				if (!name) {
					throw new NodeOperationError(
						node,
						`No binary field name configured for input ${inputIndex + 1}.`,
					);
				}
				return name;
			};
		}

		const outputWorkbook = new ExcelJS.Workbook();
		outputWorkbook.creator = 'n8n-nodes-excel-plus';
		outputWorkbook.created = new Date();

		const usedNames = new Set<string>();
		const pairedItems: IPairedItemData[] = [];
		const skipped: Array<{ input: number; item: number; error: string }> = [];
		let mergedFileCount = 0;
		let mergedSheetCount = 0;

		for (let inputIndex = 0; inputIndex < numberOfInputs; inputIndex++) {
			const items = this.getInputData(inputIndex);

			for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
				try {
					const binaryFieldName = resolveBinaryName(inputIndex);
					const item = items[itemIndex];
					const binary: IBinaryData | undefined = item.binary?.[binaryFieldName];
					if (!binary) {
						throw new NodeOperationError(
							node,
							`Input ${inputIndex + 1}, item ${itemIndex} has no binary property "${binaryFieldName}".`,
							{ itemIndex },
						);
					}

					let buffer: Buffer;
					if (binary.id) {
						const stream = await this.helpers.getBinaryStream(binary.id);
						buffer = await this.helpers.binaryToBuffer(stream);
					} else {
						buffer = Buffer.from(binary.data, 'base64') as unknown as Buffer;
					}

					const sourceWorkbook = new ExcelJS.Workbook();
					try {
						// ExcelJS's Buffer typing predates Node 22's Buffer<ArrayBufferLike> generic.
						await sourceWorkbook.xlsx.load(buffer as never);
					} catch (error) {
						throw new NodeOperationError(
							node,
							`Failed to read Excel file from input ${inputIndex + 1}, item ${itemIndex} (binary "${binaryFieldName}"): ${(error as Error).message}`,
							{ itemIndex },
						);
					}

					const fallbackPrefix = `Input ${inputIndex + 1}`;
					for (const sourceSheet of sourceWorkbook.worksheets) {
						const desiredName =
							sheetNameConflict === 'prefix'
								? buildPrefixedName(binary.fileName, sourceSheet.name, fallbackPrefix)
								: sourceSheet.name;

						const finalName = uniqueSheetName(node, desiredName, usedNames);
						usedNames.add(finalName);

						const targetSheet = outputWorkbook.addWorksheet(finalName);
						copySheet(sourceSheet, targetSheet);
						mergedSheetCount++;
					}

					pairedItems.push({ item: itemIndex, input: inputIndex });
					mergedFileCount++;
				} catch (error) {
					if (this.continueOnFail()) {
						skipped.push({
							input: inputIndex + 1,
							item: itemIndex,
							error: (error as Error).message,
						});
						continue;
					}
					throw new NodeOperationError(node, error as Error, { itemIndex });
				}
			}
		}

		if (mergedSheetCount === 0) {
			throw new NodeOperationError(
				node,
				'No sheets were merged. Make sure each input is wired and carries an Excel binary, then try again.',
			);
		}

		const writeResult = await outputWorkbook.xlsx.writeBuffer();
		// Bridge Node 22's Buffer<ArrayBufferLike> back to the legacy Buffer type expected by n8n-workflow.
		const outputBuffer = Buffer.from(writeResult as unknown as ArrayBuffer) as unknown as Buffer;
		const outputBinary = await this.helpers.prepareBinaryData(
			outputBuffer,
			outputFileName,
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		);

		const jsonPayload: IDataObject = {
			operation,
			mergedFiles: mergedFileCount,
			mergedSheets: mergedSheetCount,
			sheetNames: Array.from(usedNames),
			fileName: outputFileName,
		};
		if (skipped.length > 0) jsonPayload.skipped = skipped;

		return [
			[
				{
					json: jsonPayload,
					binary: {
						data: outputBinary,
					},
					pairedItem: pairedItems,
				},
			],
		];
	}
}
