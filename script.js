'use strict';

const {tokenize} = require('excel-formula-tokenizer');
const {buildTree, visit} = require('excel-formula-ast');

const formula = 'SUM(1, 2)';

window.t = (f) => {
	let tokens = tokenize(String(f));
	let tree = buildTree(tokens);
	console.log(tree);
}

const functions = {
	SUM: (...vals) => vals.reduce((sum, v) => v + sum, 0)
};

const getVal = (spreadsheet, tree) => {
	if (tree.type === 'number' || tree.type === 'text') {
		return tree.value;
	}
	if (tree.type === 'function') {
		return functions[tree.name](...tree.arguments.map(arg => getVal(spreadsheet, arg)))
	}
	if (tree.type === 'cell') {
		let cell = spreadsheet.cells[tree.key] || { value: '' };
		return cellContents(spreadsheet, cell);
	}
	if (tree.type === 'binary-expression') {
		if (tree.operator === '+') {
			return getVal(spreadsheet, tree.left) + getVal(spreadsheet, tree.right);
		}
		if (tree.operator === '-') {
			return getVal(spreadsheet, tree.left) - getVal(spreadsheet, tree.right);
		}
		if (tree.operator === '*') {
			return getVal(spreadsheet, tree.left) * getVal(spreadsheet, tree.right);
		}
		if (tree.operator === '/') {
			return getVal(spreadsheet, tree.left) / getVal(spreadsheet, tree.right);
		}
	}
	console.log(tree);
};

const cellContents = (spreadsheet, cell) => {
	let tokens = tokenize(String(cell.value));
	let tree = buildTree(tokens);
	return getVal(spreadsheet, tree);
};

fetch('example.json')
	.then(result => result.json())
	.then(spreadsheet => {
		Object.entries(spreadsheet.cells).forEach(([address, cell]) => {
			address = address.match(/^([A-Z]+)([0-9]+)$/);
			cell.col = linum2int(address[1]) - 1;
			cell.row = Number(address[2]) - 1;
		});
		renderSpreadsheet(spreadsheet);
	});

const $ = selector => document.querySelector(selector);

const $canvas = $('canvas');
const ctx = $canvas.getContext('2d');

ctx.textAlign = 'left';
ctx.textBaseline = 'top';
ctx.translate(0.5, 0.5); // prevents aliasing

const scrollX = 0;
const scrollY = 0;

const renderSpreadsheet = spreadsheet => {
	Object.values(spreadsheet.cells).forEach(cell => {
		renderCell(spreadsheet, cell);
	});
};

const CELL_HEIGHT = 20;
const CELL_WIDTH = 100;
const STYLE_DEFAULTS = {
	'background-color': '#FFFFFF',
	'color': '#000000',
	'font-weight': 'normal',
	'font-size': '12px',
	'font-family': 'Arial',
	'border-top-width': '1px',
	'border-top-color': '#ccc',
	'border-right-width': '1px',
	'border-right-color': '#ccc',
	'border-bottom-width': '1px',
	'border-bottom-color': '#ccc',
	'border-left-width': '1px',
	'border-left-color': '#ccc',
};

const getStyle = (cell, attr) => cell.style && cell.style[attr] || STYLE_DEFAULTS[attr];

const renderCell = (spreadsheet, cell) => {
	let x = cell.col * CELL_WIDTH + scrollX;
	let y = cell.row * CELL_HEIGHT + scrollY;
	
	ctx.fillStyle = getStyle(cell, 'background-color');
	ctx.fillRect(x, y, CELL_WIDTH, CELL_HEIGHT);


	ctx.strokeStyle = getStyle(cell, 'border-top-color');
	ctx.lineWidth = parseInt(getStyle(cell, 'border-top-width'), 10);
	ctx.beginPath();
	ctx.moveTo(x, y);
	ctx.lineTo(x + CELL_WIDTH, y);
	ctx.stroke();

	ctx.strokeStyle = getStyle(cell, 'border-right-color');
	ctx.lineWidth = parseInt(getStyle(cell, 'border-right-width'), 10);
	ctx.beginPath();
	ctx.moveTo(x + CELL_WIDTH, y);
	ctx.lineTo(x + CELL_WIDTH, y + CELL_HEIGHT);
	ctx.stroke();

	ctx.strokeStyle = getStyle(cell, 'border-bottom-color');
	ctx.lineWidth = parseInt(getStyle(cell, 'border-bottom-width'), 10);
	ctx.beginPath();
	ctx.moveTo(x + CELL_WIDTH, y + CELL_HEIGHT);
	ctx.lineTo(x, y + CELL_HEIGHT);
	ctx.stroke();

	ctx.strokeStyle = getStyle(cell, 'border-left-color');
	ctx.lineWidth = parseInt(getStyle(cell, 'border-left-width'), 10);
	ctx.beginPath();
	ctx.moveTo(x, y + CELL_HEIGHT);
	ctx.lineTo(x, y);
	ctx.stroke();


	ctx.font = [getStyle(cell, 'font-weight'), getStyle(cell, 'font-size'), getStyle(cell, 'font-family')].join(' ');
	ctx.fillStyle = getStyle(cell, 'color');
	ctx.fillText(cellContents(spreadsheet, cell), x, y)
};

const linum2int = input => {
	input = input.replace(/[^A-Za-z]/, '');
	let output = 0;
	for (let i = 0; i < input.length; i++) {
		output = output * 26 + parseInt(input.substr(i, 1), 36) - 9;
	}
	return output;
}

function int2linum(input) {
	let zeros = 0;
	let next = input;
	let generation = 0;
	while (next >= 27) {
		next = (next - 1) / 26 - (next - 1) % 26 / 26;
		zeros += next * Math.pow(27, generation);
		generation++;
	}
	return (input + zeros).toString(27).replace(/./g, ($0) => {
		return '_abcdefghijklmnopqrstuvwxyz'.charAt(parseInt($0, 27));
	});
}