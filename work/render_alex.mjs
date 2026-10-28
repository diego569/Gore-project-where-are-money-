import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load('ref docs/0001917176_1_SINVERSION (1).xlsm'));
for(const [sheetName,range] of [['Gerencia de Infraestructura','A1:L18'],['Gerencias y Oficinas','A1:L18'],['ABS','A1:L15']]){
 const im=await w.render({sheetName,range,scale:1});await fs.writeFile('work/alex_'+sheetName.replaceAll(' ','_')+'.png',new Uint8Array(await im.arrayBuffer()));
}
