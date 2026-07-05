const xlsx = require('xlsx');
const fs = require('fs');

try {
    const wb = xlsx.readFile('ฐานข้อมูลวิชา(2).xlsx');
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    const output = {
        sheetName: sheetName,
        rows: json.slice(0, 10)
    };
    
    fs.writeFileSync('excel_structure.json', JSON.stringify(output, null, 2));
    console.log('Saved to excel_structure.json');
} catch (e) {
    console.error(e);
}
