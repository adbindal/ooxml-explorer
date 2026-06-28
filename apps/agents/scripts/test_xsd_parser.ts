import { getXSDGrounding } from '../mastra/xsdParser';

console.log('--- cantSplit (docx) ---');
console.log(JSON.stringify(getXSDGrounding('cantSplit', 'docx'), null, 2));

console.log('--- tblHeader (docx) ---');
console.log(JSON.stringify(getXSDGrounding('tblHeader', 'docx'), null, 2));

console.log('--- row (xlsx) ---');
console.log(JSON.stringify(getXSDGrounding('row', 'xlsx'), null, 2));
