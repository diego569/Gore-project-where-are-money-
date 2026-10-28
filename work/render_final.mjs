import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx'));
for(const [name,range] of [['Resumen','A1:I23'],['Proyectos','A1:J9'],['Proyectos','K5:AG8'],['PMI','A1:M9'],['Control','A44:F62'],['Clasificacion Alex','A1:F9']]) {
 const b=await wb.render({sheetName:name,range,scale:1});await fs.writeFile(`work/final_${name.replaceAll(' ','_')}_${range.replace(':','_')}.png`,new Uint8Array(await b.arrayBuffer()));
}
console.log((await wb.inspect({kind:'region',sheetId:'Resumen',range:'A6:I15',maxChars:1000,tableMaxCols:9})).ndjson);
