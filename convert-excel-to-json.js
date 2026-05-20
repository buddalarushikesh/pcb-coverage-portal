const XLSX = require('xlsx');
const fs = require('fs');

// Convert your updated Excel file to JSON
function convertExcelToJSON(excelPath, outputJsonPath) {
  try {
    const workbook = XLSX.readFile(excelPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    
    // Convert to the format expected by the system
    const converted = {};
    
    if (excelPath.includes('stock')) {
      // Convert stock data
      for (const row of data) {
        const key = String(row.Plant) + String(row.Material).padStart(8, '0');
        converted[key] = {
          unrestricted: row.Unrestricted || 0,
          transit: row.Transit || 0,
          snapshot_date: new Date().toISOString().split('T')[0],
          plant: String(row.Plant),
          material: String(row.Material)
        };
      }
    } else if (excelPath.includes('master')) {
      // Convert master data
      for (const row of data) {
        const key = String(row['Plant Code']) + String(row.Item_PCB).padStart(8, '0');
        const perday = (row.ST || 0) / 300;
        converted[key] = {
          item_pcb: row.Item_PCB,
          plant_code: String(row['Plant Code']),
          plant_name: getPlantName(String(row['Plant Code'])),
          description: row.Description,
          model: row.Model || '',
          customer: row.Cust || '',
          lead_time: row['Lead Time'] || 0,
          annual_plan: row.ST || 0,
          perday: perday
        };
      }
    }
    
    fs.writeFileSync(outputJsonPath, JSON.stringify(converted, null, 2));
    console.log(`✅ Converted ${Object.keys(converted).length} records to ${outputJsonPath}`);
    return true;
  } catch (error) {
    console.error('Error converting Excel:', error);
    return false;
  }
}

function getPlantName(plantCode) {
  const plants = {
    "1300": "DHR", "1600": "BWL", "3200": "BLR", "4100": "SAND I",
    "2400": "SAND II", "1800": "HDR", "1700": "PANT", "9400": "LATL PANT",
    "2100": "CHINWAD", "2300": "CHAKN 2", "2200": "CHAKN 3", "7400": "LATL-CHAKAN"
  };
  return plants[plantCode] || plantCode;
}

// Convert all updated files
console.log('🔄 Converting Excel files to JSON...\n');

// Update these paths to your actual Excel file locations
const filesToConvert = [
  { excel: './samples/stock_data_sample.xlsx', json: './data/stockData.json', type: 'stock' },
  { excel: './samples/pcb_master_sample.xlsx', json: './data/pcbMaster.json', type: 'master' },
  { excel: './samples/transit_data_sample.xlsx', json: './data/stockData.json', type: 'transit' }
];

for (const file of filesToConvert) {
  if (fs.existsSync(file.excel)) {
    convertExcelToJSON(file.excel, file.json);
  }
}

console.log('\n✅ Conversion complete! Restart server or recalculate status.');