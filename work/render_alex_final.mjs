import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load('outputs/alex_central/Alex_Sede_Central_2026.xlsx'));
for(const [sheetName,range] of [['Gerencia de Infraestructura','A1:O14'],['Gerencias y Oficinas','A1:O11'],['Gerencias y Oficinas','A314:O323'],['GRDE','A1:O11'],['GRRNyDA','A1:O10'],['PETT','A1:O11'],['PECSA','A1:O13'],['DS','A1:O10'],['DEPI','A1:O10'],['ABS','A1:L22']]){
 const im=await w.render({sheetName,range,scale:1});await fs.writeFile('work/new_alex_'+sheetName.replaceAll(' ','_')+'_'+range.replace(':','_')+'.png',new Uint8Array(await im.arrayBuffer()));
}
