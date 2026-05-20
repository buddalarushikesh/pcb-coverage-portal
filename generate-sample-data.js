const XLSX = require('xlsx');
const fs = require('fs');

// Ensure samples directory exists
if (!fs.existsSync('./samples')) {
  fs.mkdirSync('./samples');
}

// Sample PCB Master Data
const masterData = [
  { Item_PCB: 'PCB001', 'Plant Code': '1300', Description: 'Engine Control Unit', Model: 'ECU-X1', Cust: 'Maruti', 'Lead Time': 2, ST: 15000 },
  { Item_PCB: 'PCB002', 'Plant Code': '1600', Description: 'Battery Management System', Model: 'BMS-V2', Cust: 'Tata', 'Lead Time': 2, ST: 12000 },
  { Item_PCB: 'PCB003', 'Plant Code': '3200', Description: 'Infotainment System', Model: 'INFO-S3', Cust: 'Hyundai', 'Lead Time': 10, ST: 8000 },
  { Item_PCB: 'PCB004', 'Plant Code': '4100', Description: 'Dashboard Cluster', Model: 'DASH-M4', Cust: 'Mahindra', 'Lead Time': 7, ST: 10000 },
  { Item_PCB: 'PCB005', 'Plant Code': '2400', Description: 'Lighting Control', Model: 'LIGHT-R5', Cust: 'Honda', 'Lead Time': 7, ST: 9000 },
  { Item_PCB: 'PCB006', 'Plant Code': '1800', Description: 'Power Window Module', Model: 'PWR-W6', Cust: 'Ford', 'Lead Time': 4, ST: 11000 },
  { Item_PCB: 'PCB007', 'Plant Code': '1700', Description: 'Airbag Controller', Model: 'AIR-X7', Cust: 'BMW', 'Lead Time': 4, ST: 5000 },
  { Item_PCB: 'PCB008', 'Plant Code': '9400', Description: 'GPS Tracker', Model: 'GPS-T8', Cust: 'Toyota', 'Lead Time': 4, ST: 7000 },
  { Item_PCB: 'PCB009', 'Plant Code': '2100', Description: 'Climate Control', Model: 'CLIM-C9', Cust: 'Nissan', 'Lead Time': 7, ST: 6000 },
  { Item_PCB: 'PCB010', 'Plant Code': '2300', Description: 'Parking Sensor', Model: 'PARK-P0', Cust: 'Volkswagen', 'Lead Time': 7, ST: 8000 }
];

// Sample Stock Data (MB52)
const stockData = [
  { Plant: '1300', Material: 'PCB001', 'Storage Location': '4000', Unrestricted: 500 },
  { Plant: '1300', Material: 'PCB001', 'Storage Location': '2700', Unrestricted: 200 },
  { Plant: '1600', Material: 'PCB002', 'Storage Location': '4000', Unrestricted: 300 },
  { Plant: '3200', Material: 'PCB003', 'Storage Location': '4000', Unrestricted: 150 },
  { Plant: '4100', Material: 'PCB004', 'Storage Location': '4000', Unrestricted: 80 },
  { Plant: '2400', Material: 'PCB005', 'Storage Location': '4000', Unrestricted: 450 },
  { Plant: '1800', Material: 'PCB006', 'Storage Location': '4000', Unrestricted: 600 },
  { Plant: '1700', Material: 'PCB007', 'Storage Location': '4000', Unrestricted: 25 },
  { Plant: '9400', Material: 'PCB008', 'Storage Location': '4000', Unrestricted: 120 },
  { Plant: '2100', Material: 'PCB009', 'Storage Location': '4000', Unrestricted: 90 },
  { Plant: '2300', Material: 'PCB010', 'Storage Location': '4000', Unrestricted: 200 }
];

// Sample Transit Data
const transitData = [
  { 'Row Labels': '1300PCB001', 'Sum of Quantity': 100 },
  { 'Row Labels': '1600PCB002', 'Sum of Quantity': 50 },
  { 'Row Labels': '3200PCB003', 'Sum of Quantity': 75 },
  { 'Row Labels': '4100PCB004', 'Sum of Quantity': 30 },
  { 'Row Labels': '2400PCB005', 'Sum of Quantity': 60 }
];

// Calculate Excel serial date for current date
function dateToExcelSerial(date) {
  const excelEpoch = new Date(1899, 11, 30);
  const diffDays = Math.floor((date - excelEpoch) / (24 * 60 * 60 * 1000));
  return diffDays;
}

// Sample Sale Data (last 30 days)
const saleData = [];
const today = new Date();
for (let i = 0; i < 30; i++) {
  const date = new Date(today);
  date.setDate(today.getDate() - i);
  
  // Random sales for different plants
  const sales = [
    { 'Material No.': 'PCB001', 'Plant Name': 'DHR', Qty: Math.floor(Math.random() * 50) + 10, 'Date of Invoice': dateToExcelSerial(date) },
    { 'Material No.': 'PCB002', 'Plant Name': 'BWL', Qty: Math.floor(Math.random() * 40) + 10, 'Date of Invoice': dateToExcelSerial(date) },
    { 'Material No.': 'PCB003', 'Plant Name': 'BLR', Qty: Math.floor(Math.random() * 30) + 5, 'Date of Invoice': dateToExcelSerial(date) },
    { 'Material No.': 'PCB004', 'Plant Name': 'SAND I', Qty: Math.floor(Math.random() * 35) + 5, 'Date of Invoice': dateToExcelSerial(date) }
  ];
  saleData.push(...sales);
}

// Generate Excel files
function generateSampleFiles() {
  console.log('📁 Generating sample Excel files...');
  
  // 1. Master File
  const masterWs = XLSX.utils.json_to_sheet(masterData);
  const masterWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(masterWb, masterWs, 'PCB_Lumax');
  XLSX.writeFile(masterWb, './samples/pcb_master_sample.xlsx');
  console.log('✅ Created: samples/pcb_master_sample.xlsx');
  
  // 2. Stock File
  const stockWs = XLSX.utils.json_to_sheet(stockData);
  const stockWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(stockWb, stockWs, 'Sheet1');
  XLSX.writeFile(stockWb, './samples/stock_data_sample.xlsx');
  console.log('✅ Created: samples/stock_data_sample.xlsx');
  
  // 3. Transit File
  const transitWs = XLSX.utils.json_to_sheet(transitData);
  const transitWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(transitWb, transitWs, 'Sheet1');
  XLSX.writeFile(transitWb, './samples/transit_data_sample.xlsx');
  console.log('✅ Created: samples/transit_data_sample.xlsx');
  
  // 4. Sale File
  const saleWs = XLSX.utils.json_to_sheet(saleData);
  const saleWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(saleWb, saleWs, 'Sheet1');
  XLSX.writeFile(saleWb, './samples/sale_data_sample.xlsx');
  console.log('✅ Created: samples/sale_data_sample.xlsx');
  
  console.log('\n📋 Sample files generated successfully!');
  console.log('📍 Location: ./samples/ folder');
  console.log('\n🔧 Next steps:');
  console.log('1. Go to http://localhost:3000/upload.html');
  console.log('2. Upload files in this order:');
  console.log('   - PCB Master List (pcb_master_sample.xlsx)');
  console.log('   - Stock Data (stock_data_sample.xlsx)');
  console.log('   - Transit Stock (transit_data_sample.xlsx)');
  console.log('   - Sale Data (sale_data_sample.xlsx)');
  console.log('3. Click "Recalculate Coverage Status"');
  console.log('4. View dashboard and status table');
}

generateSampleFiles();