const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// ============ CONFIGURE YOUR EXCEL FILE PATHS HERE ============
const EXCEL_FILES = {
  master: './samples/pcb_master_sample.xlsx',    // Change to your master file path
  stock: './samples/stock_data_sample.xlsx',     // Change to your stock file path
  transit: './samples/transit_data_sample.xlsx', // Change to your transit file path
  sale: './samples/sale_data_sample.xlsx'        // Change to your sale file path
};

// Plants data (fixed - from requirements)
const plants = {
  "1300": { name: "DHR", lead_time: 2 },
  "1600": { name: "BWL", lead_time: 2 },
  "3200": { name: "BLR", lead_time: 10 },
  "4100": { name: "SAND I", lead_time: 7 },
  "2400": { name: "SAND II", lead_time: 7 },
  "1800": { name: "HDR", lead_time: 4 },
  "1700": { name: "PANT", lead_time: 4 },
  "9400": { name: "LATL PANT", lead_time: 4 },
  "2100": { name: "CHINWAD", lead_time: 7 },
  "2300": { name: "CHAKN 2", lead_time: 7 },
  "2200": { name: "CHAKN 3", lead_time: 7 },
  "7400": { name: "LATL-CHAKAN", lead_time: 7 }
};

// Helper functions
function makeKey(plantCode, materialCode) {
  return String(plantCode) + String(materialCode).trim().padStart(8, '0');
}

function excelDateToJSDate(serial) {
  if (!serial || typeof serial !== 'number') return new Date();
  return new Date((serial - 25569) * 86400 * 1000);
}

function classifyStatus(coverageDays, leadTime) {
  if (coverageDays === null || isNaN(coverageDays) || coverageDays === undefined) {
    return 'No Plan';
  }
  
  if (leadTime <= 2) {
    if (coverageDays < 2) return 'Very Critical';
    if (coverageDays < 3) return 'Potential Critical';
    if (coverageDays <= 6) return 'Sufficient';
    return 'High';
  } else if (leadTime <= 4) {
    if (coverageDays < 2) return 'Very Critical';
    if (coverageDays < 4) return 'Potential Critical';
    if (coverageDays <= 10) return 'Sufficient';
    return 'High';
  } else if (leadTime <= 7) {
    if (coverageDays < 4) return 'Very Critical';
    if (coverageDays < 7) return 'Potential Critical';
    if (coverageDays <= 16) return 'Sufficient';
    return 'High';
  } else {
    if (coverageDays < 5) return 'Very Critical';
    if (coverageDays < 10) return 'Potential Critical';
    if (coverageDays <= 16) return 'Sufficient';
    return 'High';
  }
}

// ============ READ DATA DIRECTLY FROM EXCEL ============

// Read PCB Master from Excel
function getPCBMaster() {
  if (!fs.existsSync(EXCEL_FILES.master)) {
    console.error(`Master file not found: ${EXCEL_FILES.master}`);
    return [];
  }
  
  const workbook = XLSX.readFile(EXCEL_FILES.master);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const master = [];
  for (const row of rows) {
    const itemPcb = String(row.Item_PCB || row['Item PCB'] || '');
    const plantCode = String(row['Plant Code'] || row.Plant_Code || '');
    
    if (itemPcb && plantCode && plants[plantCode]) {
      const annualPlan = parseFloat(row.ST || row['annual plan'] || 0);
      const perday = annualPlan > 0 ? annualPlan / 300 : 0;
      
      master.push({
        item_pcb: itemPcb,
        plant_code: plantCode,
        plant_name: plants[plantCode].name,
        description: row.Description || '',
        model: row.Model || '',
        customer: row.Cust || row.Customer || '',
        lead_time: parseInt(row['Lead Time']) || plants[plantCode].lead_time,
        annual_plan: annualPlan,
        perday: perday
      });
    }
  }
  
  console.log(`📋 Read ${master.length} PCB records from Excel`);
  return master;
}

// Read Stock data from Excel
function getStockData() {
  if (!fs.existsSync(EXCEL_FILES.stock)) {
    console.error(`Stock file not found: ${EXCEL_FILES.stock}`);
    return {};
  }
  
  const workbook = XLSX.readFile(EXCEL_FILES.stock);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const validSlocs = ['4000', '2700', '2799', '1700', '2400', '1600'];
  const stockMap = {};
  
  for (const row of rows) {
    const plant = String(row.Plant || '');
    const material = String(row.Material || '');
    const unrestricted = parseFloat(row.Unrestricted) || 0;
    const storageLoc = String(row['Storage Location'] || '');
    
    if (plant && material && validSlocs.includes(storageLoc)) {
      const key = makeKey(plant, material);
      if (!stockMap[key]) {
        stockMap[key] = { stock: 0, transit: 0 };
      }
      stockMap[key].stock += unrestricted;
    }
  }
  
  console.log(`📦 Read ${Object.keys(stockMap).length} stock records from Excel`);
  return stockMap;
}

// Read Transit data from Excel
function getTransitData() {
  if (!fs.existsSync(EXCEL_FILES.transit)) {
    console.log(`Transit file not found, skipping...`);
    return {};
  }
  
  const workbook = XLSX.readFile(EXCEL_FILES.transit);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const transitMap = {};
  for (const row of rows) {
    const key = String(row['Row Labels'] || row['Key'] || '');
    const quantity = parseFloat(row['Sum of Quantity'] || row['Quantity'] || 0);
    
    if (key && !isNaN(quantity)) {
      transitMap[key] = quantity;
    }
  }
  
  console.log(`🚚 Read ${Object.keys(transitMap).length} transit records from Excel`);
  return transitMap;
}

// Read Sale data from Excel
function getSaleData() {
  if (!fs.existsSync(EXCEL_FILES.sale)) {
    console.log(`Sale file not found, skipping...`);
    return {};
  }
  
  const workbook = XLSX.readFile(EXCEL_FILES.sale);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const currentDate = new Date();
  const currentMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
  const saleMap = {};
  
  for (const row of rows) {
    const materialNo = String(row['Material No.'] || row['Material'] || '');
    const plantName = String(row['Plant Name'] || row['Plant'] || '');
    const quantity = parseFloat(row['Qty'] || row['Quantity'] || 0);
    const dateOfInvoice = row['Date of Invoice'];
    
    // Find plant code
    let plantCode = null;
    for (const [code, plant] of Object.entries(plants)) {
      if (plant.name === plantName || plantName.includes(plant.name)) {
        plantCode = code;
        break;
      }
    }
    
    if (materialNo && plantCode && quantity > 0) {
      let invoiceDate;
      if (typeof dateOfInvoice === 'number') {
        invoiceDate = excelDateToJSDate(dateOfInvoice);
      } else if (dateOfInvoice) {
        invoiceDate = new Date(dateOfInvoice);
      } else {
        continue;
      }
      
      const invoiceMonth = `${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}`;
      
      if (invoiceMonth === currentMonth) {
        const key = makeKey(plantCode, materialNo);
        saleMap[key] = (saleMap[key] || 0) + quantity;
      }
    }
  }
  
  console.log(`💰 Read ${Object.keys(saleMap).length} sale records from Excel for current month`);
  return saleMap;
}

// Calculate perday from actual consumption (last 30 days)
function calculateActualPerday(itemPcb, plantCode) {
  if (!fs.existsSync(EXCEL_FILES.sale)) {
    return 0;
  }
  
  const workbook = XLSX.readFile(EXCEL_FILES.sale);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);
  
  let totalQty = 0;
  let daysWithData = 0;
  
  for (const row of rows) {
    const materialNo = String(row['Material No.'] || row['Material'] || '');
    const plantName = String(row['Plant Name'] || row['Plant'] || '');
    const quantity = parseFloat(row['Qty'] || row['Quantity'] || 0);
    const dateOfInvoice = row['Date of Invoice'];
    
    let plantCode_row = null;
    for (const [code, plant] of Object.entries(plants)) {
      if (plant.name === plantName || plantName.includes(plant.name)) {
        plantCode_row = code;
        break;
      }
    }
    
    if (materialNo === itemPcb && plantCode_row === plantCode && quantity > 0) {
      let invoiceDate;
      if (typeof dateOfInvoice === 'number') {
        invoiceDate = excelDateToJSDate(dateOfInvoice);
      } else if (dateOfInvoice) {
        invoiceDate = new Date(dateOfInvoice);
      } else {
        continue;
      }
      
      if (invoiceDate >= cutoffDate) {
        totalQty += quantity;
        daysWithData++;
      }
    }
  }
  
  // Average over last 30 days (assuming 25 working days)
  return daysWithData > 0 ? totalQty / 25 : 0;
}

// Main function to get all status data directly from Excel
function getAllStatus() {
  console.log('\n🔄 Reading fresh data from Excel files...');
  
  const masterList = getPCBMaster();
  const stockMap = getStockData();
  const transitMap = getTransitData();
  
  // Merge transit into stock
  for (const [key, transit] of Object.entries(transitMap)) {
    if (!stockMap[key]) {
      stockMap[key] = { stock: 0, transit: 0 };
    }
    stockMap[key].transit = transit;
  }
  
  const statusList = [];
  const snapshotDate = new Date().toISOString().split('T')[0];
  
  for (const master of masterList) {
    const key = makeKey(master.plant_code, master.item_pcb);
    const stockInfo = stockMap[key] || { stock: 0, transit: 0 };
    
    const finalStock = (stockInfo.stock || 0) + (stockInfo.transit || 0);
    
    // Calculate perday: actual consumption first, then fallback to master perday
    const actualPerday = calculateActualPerday(master.item_pcb, master.plant_code);
    const perday = actualPerday > 0 ? actualPerday : master.perday;
    
    let coverageDays = null;
    if (perday > 0) {
      coverageDays = finalStock / perday;
    }
    
    const remarks = classifyStatus(coverageDays, master.lead_time);
    
    statusList.push({
      snapshot_date: snapshotDate,
      item_pcb: master.item_pcb,
      plant_code: master.plant_code,
      plant_name: master.plant_name,
      description: master.description,
      model: master.model,
      customer: master.customer,
      lead_time: master.lead_time,
      perday: perday,
      annual_plan: master.annual_plan,
      stock: stockInfo.stock || 0,
      transit: stockInfo.transit || 0,
      final_stock: finalStock,
      coverage_days: coverageDays,
      remarks: remarks
    });
  }
  
  console.log(`✅ Calculated ${statusList.length} status records\n`);
  return statusList;
}

// ============ API ROUTES ============

app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'PCB Portal - Excel Only Mode' });
});

app.get('/api/plants', (req, res) => {
  res.json(plants);
});

app.get('/api/status', (req, res) => {
  try {
    const { plant, status, search, page = 1, limit = 50 } = req.query;
    
    // Always read fresh from Excel
    let results = getAllStatus();
    
    // Apply filters
    if (plant) {
      results = results.filter(r => r.plant_code === plant);
    }
    if (status) {
      results = results.filter(r => r.remarks === status);
    }
    if (search) {
      results = results.filter(r => 
        r.item_pcb.includes(search) || 
        (r.description || '').includes(search)
      );
    }
    
    // Sort by priority
    const priority = { 'Very Critical': 1, 'Potential Critical': 2, 'Sufficient': 3, 'High': 4, 'No Plan': 5 };
    results.sort((a, b) => (priority[a.remarks] || 99) - (priority[b.remarks] || 99));
    
    const total = results.length;
    const start = (parseInt(page) - 1) * parseInt(limit);
    const paginated = results.slice(start, start + parseInt(limit));
    
    const summary = {
      veryCritical: results.filter(r => r.remarks === 'Very Critical').length,
      potentialCritical: results.filter(r => r.remarks === 'Potential Critical').length,
      sufficient: results.filter(r => r.remarks === 'Sufficient').length,
      high: results.filter(r => r.remarks === 'High').length,
      noPlan: results.filter(r => r.remarks === 'No Plan').length
    };
    
    res.json({
      success: true,
      data: paginated,
      total: total,
      totalPages: Math.ceil(total / parseInt(limit)),
      currentPage: parseInt(page),
      summary: summary
    });
  } catch (error) {
    console.error('Status error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/status/summary', (req, res) => {
  try {
    const results = getAllStatus();
    
    const plantSummary = {};
    let totalPCBs = results.length;
    let summary = {
      veryCritical: 0,
      potentialCritical: 0,
      sufficient: 0,
      high: 0,
      noPlan: 0
    };
    
    for (const record of results) {
      if (!plantSummary[record.plant_code]) {
        plantSummary[record.plant_code] = {
          plant_code: record.plant_code,
          plant_name: record.plant_name,
          total: 0,
          veryCritical: 0,
          potentialCritical: 0,
          sufficient: 0,
          high: 0,
          noPlan: 0
        };
      }
      
      plantSummary[record.plant_code].total++;
      
      switch(record.remarks) {
        case 'Very Critical':
          plantSummary[record.plant_code].veryCritical++;
          summary.veryCritical++;
          break;
        case 'Potential Critical':
          plantSummary[record.plant_code].potentialCritical++;
          summary.potentialCritical++;
          break;
        case 'Sufficient':
          plantSummary[record.plant_code].sufficient++;
          summary.sufficient++;
          break;
        case 'High':
          plantSummary[record.plant_code].high++;
          summary.high++;
          break;
        default:
          plantSummary[record.plant_code].noPlan++;
          summary.noPlan++;
      }
    }
    
    res.json({
      success: true,
      totalPCBs: totalPCBs,
      summary: summary,
      plants: Object.values(plantSummary)
    });
  } catch (error) {
    console.error('Summary error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/export/excel', (req, res) => {
  try {
    const results = getAllStatus();
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    const excelData = results.map((r, index) => ({
      'S.No': index + 1,
      'Item_PCB': r.item_pcb,
      'Description': r.description,
      'Model': r.model,
      'Customer': r.customer,
      'Plant': `${r.plant_code} - ${r.plant_name}`,
      'Lead Time': r.lead_time,
      'Perday': r.perday?.toFixed(2),
      'Stock': r.stock?.toFixed(0),
      'Transit': r.transit?.toFixed(0),
      'Final Stock': r.final_stock?.toFixed(0),
      'Coverage Days': r.coverage_days?.toFixed(1),
      'Status': r.remarks
    }));
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PCB_Status');
    
    const filename = `pcb_status_${snapshotDate}.xlsx`;
    const filepath = path.join(__dirname, filename);
    XLSX.writeFile(wb, filepath);
    
    res.download(filepath, filename, (err) => {
      if (err) console.error('Download error:', err);
      setTimeout(() => {
        if (fs.existsSync(filepath)) fs.unlinkSync(filepath);
      }, 60000);
    });
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Route to get file status
app.get('/api/file-status', (req, res) => {
  const fileStatus = {};
  for (const [type, filePath] of Object.entries(EXCEL_FILES)) {
    fileStatus[type] = fs.existsSync(filePath);
  }
  res.json(fileStatus);
});

// Start server
app.listen(PORT, () => {
  console.log(`\n✅ PCB Portal running on http://localhost:${PORT}`);
  console.log(`📁 Excel Files Mode - NO JSON Storage!\n`);
  console.log(`📋 Reading from Excel files:`);
  for (const [type, filePath] of Object.entries(EXCEL_FILES)) {
    const exists = fs.existsSync(filePath);
    console.log(`   ${type}: ${filePath} ${exists ? '✅' : '❌'}`);
  }
  console.log(`\n💡 Simply update your Excel files and refresh the browser!`);
});