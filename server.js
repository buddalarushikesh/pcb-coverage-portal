const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Ensure directories exist
if (!fs.existsSync('./data')) fs.mkdirSync('./data');
if (!fs.existsSync('./uploads')) fs.mkdirSync('./uploads');

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/')
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + '.xlsx');
  }
});
const upload = multer({ storage: storage });

// Plants data
const plants = {
  "1300": { name: "DHR", lead_time: 2, group: "lead_time_2" },
  "1600": { name: "BWL", lead_time: 2, group: "lead_time_2" },
  "3200": { name: "BLR", lead_time: 10, group: "lead_time_10" },
  "4100": { name: "SAND I", lead_time: 7, group: "lead_time_7" },
  "2400": { name: "SAND II", lead_time: 7, group: "lead_time_7" },
  "1800": { name: "HDR", lead_time: 4, group: "lead_time_4" },
  "1700": { name: "PANT", lead_time: 4, group: "lead_time_4" },
  "9400": { name: "LATL PANT", lead_time: 4, group: "lead_time_4" },
  "2100": { name: "CHINWAD", lead_time: 7, group: "lead_time_7" },
  "2300": { name: "CHAKN 2", lead_time: 7, group: "lead_time_7" },
  "2200": { name: "CHAKN 3", lead_time: 7, group: "lead_time_7" },
  "7400": { name: "LATL-CHAKAN", lead_time: 7, group: "lead_time_7" }
};

// Data storage (in-memory with JSON files for persistence)
let stockData = {};      // { key: { unrestricted, transit, snapshot_date } }
let pcbMaster = {};      // { key: { item_pcb, plant_code, description, perday, etc } }
let dispatchData = {};   // { key: { period, qty } }
let pcbStatus = {};      // { key: { snapshot_date, coverage_days, remarks, etc } }
let uploadLogs = [];

// Helper: Make key for plant+material
function makeKey(plantCode, materialCode) {
  return String(plantCode) + String(materialCode).padStart(8, '0');
}

// Helper: Excel date to JS Date
function excelDateToJSDate(serial) {
  if (!serial || typeof serial !== 'number') return new Date();
  return new Date((serial - 25569) * 86400 * 1000);
}

// Helper: Classify status based on coverage days and lead time
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

// Helper: Classify asking rate
function classifyAskingRate(adherencePct) {
  if (!adherencePct || adherencePct === 0) return 'Falling';
  if (adherencePct < 0.85) return 'Behind';
  if (adherencePct >= 0.85 && adherencePct <= 1.05) return 'Meet';
  if (adherencePct > 1.05) return 'Advance';
  return 'Behind';
}

// Load data from JSON files on startup
function loadDataFromFiles() {
  try {
    if (fs.existsSync('./data/stockData.json')) {
      stockData = JSON.parse(fs.readFileSync('./data/stockData.json', 'utf8'));
      console.log(`Loaded stock data: ${Object.keys(stockData).length} records`);
    }
    if (fs.existsSync('./data/pcbMaster.json')) {
      pcbMaster = JSON.parse(fs.readFileSync('./data/pcbMaster.json', 'utf8'));
      console.log(`Loaded PCB master: ${Object.keys(pcbMaster).length} records`);
    }
    if (fs.existsSync('./data/dispatchData.json')) {
      dispatchData = JSON.parse(fs.readFileSync('./data/dispatchData.json', 'utf8'));
      console.log(`Loaded dispatch data: ${Object.keys(dispatchData).length} records`);
    }
    if (fs.existsSync('./data/pcbStatus.json')) {
      pcbStatus = JSON.parse(fs.readFileSync('./data/pcbStatus.json', 'utf8'));
      console.log(`Loaded PCB status: ${Object.keys(pcbStatus).length} records`);
    }
    if (fs.existsSync('./data/uploadLogs.json')) {
      uploadLogs = JSON.parse(fs.readFileSync('./data/uploadLogs.json', 'utf8'));
    }
  } catch (error) {
    console.error('Error loading data:', error);
  }
}

// Save data to JSON files
function saveDataToFiles() {
  try {
    fs.writeFileSync('./data/stockData.json', JSON.stringify(stockData, null, 2));
    fs.writeFileSync('./data/pcbMaster.json', JSON.stringify(pcbMaster, null, 2));
    fs.writeFileSync('./data/dispatchData.json', JSON.stringify(dispatchData, null, 2));
    fs.writeFileSync('./data/pcbStatus.json', JSON.stringify(pcbStatus, null, 2));
    fs.writeFileSync('./data/uploadLogs.json', JSON.stringify(uploadLogs, null, 2));
  } catch (error) {
    console.error('Error saving data:', error);
  }
}

// Log upload activity
function logUpload(fileType, filename, records, status, errorMsg = null) {
  const log = {
    id: uploadLogs.length + 1,
    upload_date: new Date().toISOString(),
    file_type: fileType,
    filename: filename,
    records_processed: records,
    status: status,
    error_msg: errorMsg
  };
  uploadLogs.unshift(log);
  if (uploadLogs.length > 50) uploadLogs.pop();
  saveDataToFiles();
}

// ============ API ROUTES ============

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'PCB Portal Server Running' });
});

// Get plants list
app.get('/api/plants', (req, res) => {
  res.json(plants);
});

// Upload Stock Data (MB52)
app.post('/api/upload/stock', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }
    
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    const validSlocs = ['4000', '2700', '2799', '1700', '2400', '1600'];
    let processed = 0;
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    // Group by plant and material
    const grouped = {};
    
    for (const row of rows) {
      const plant = String(row.Plant || '');
      const material = String(row.Material || '');
      const unrestricted = parseFloat(row.Unrestricted) || 0;
      const storageLoc = String(row['Storage Location'] || '');
      
      if (plant && material && validSlocs.includes(storageLoc)) {
        const key = makeKey(plant, material);
        if (!grouped[key]) {
          grouped[key] = { plant, material, unrestricted: 0 };
        }
        grouped[key].unrestricted += unrestricted;
        processed++;
      }
    }
    
    // Save to stockData
    for (const [key, data] of Object.entries(grouped)) {
      if (!stockData[key]) {
        stockData[key] = {};
      }
      stockData[key].unrestricted = data.unrestricted;
      stockData[key].snapshot_date = snapshotDate;
      stockData[key].plant = data.plant;
      stockData[key].material = data.material;
    }
    
    saveDataToFiles();
    logUpload('stock', req.file.originalname, processed, 'success');
    
    res.json({ success: true, records: processed, message: 'Stock data uploaded successfully' });
  } catch (error) {
    console.error('Stock upload error:', error);
    logUpload('stock', req.file?.originalname, 0, 'error', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Upload Transit Stock
app.post('/api/upload/transit', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }
    
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    let processed = 0;
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    for (const row of rows) {
      const key = String(row['Row Labels'] || row['Key'] || '');
      const quantity = parseFloat(row['Sum of Quantity'] || row['Quantity'] || 0);
      
      if (key && !isNaN(quantity)) {
        if (!stockData[key]) {
          stockData[key] = {};
        }
        stockData[key].transit = quantity;
        stockData[key].snapshot_date = snapshotDate;
        processed++;
      }
    }
    
    saveDataToFiles();
    logUpload('transit', req.file.originalname, processed, 'success');
    
    res.json({ success: true, records: processed, message: 'Transit data uploaded successfully' });
  } catch (error) {
    console.error('Transit upload error:', error);
    logUpload('transit', req.file?.originalname, 0, 'error', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Upload Sale/Dispatch Data
app.post('/api/upload/sale', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }
    
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    const currentDate = new Date();
    const currentMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
    let processed = 0;
    
    const grouped = {};
    
    for (const row of rows) {
      const materialNo = String(row['Material No.'] || row['Material'] || '');
      const plantName = String(row['Plant Name'] || row['Plant'] || '');
      const quantity = parseFloat(row['Qty'] || row['Quantity'] || 0);
      const dateOfInvoice = row['Date of Invoice'];
      
      // Find plant code from name
      let plantCode = null;
      for (const [code, plant] of Object.entries(plants)) {
        if (plant.name === plantName || plantName.includes(plant.name)) {
          plantCode = code;
          break;
        }
      }
      
      if (materialNo && plantCode && quantity > 0) {
        // Filter to current month
        let invoiceDate;
        if (typeof dateOfInvoice === 'number') {
          invoiceDate = excelDateToJSDate(dateOfInvoice);
        } else if (dateOfInvoice) {
          invoiceDate = new Date(dateOfInvoice);
        } else {
          invoiceDate = currentDate;
        }
        
        const invoiceMonth = `${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}`;
        
        if (invoiceMonth === currentMonth) {
          const key = makeKey(plantCode, materialNo);
          if (!grouped[key]) {
            grouped[key] = 0;
          }
          grouped[key] += quantity;
          processed++;
        }
      }
    }
    
    // Save to dispatchData
    for (const [key, qty] of Object.entries(grouped)) {
      const dispatchKey = `${key}_${currentMonth}`;
      dispatchData[dispatchKey] = {
        key: key,
        period: currentMonth,
        dispatched_qty: qty
      };
    }
    
    saveDataToFiles();
    logUpload('sale', req.file.originalname, processed, 'success');
    
    res.json({ success: true, records: processed, message: 'Sale data uploaded successfully' });
  } catch (error) {
    console.error('Sale upload error:', error);
    logUpload('sale', req.file?.originalname, 0, 'error', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Upload PCB Master List
app.post('/api/upload/master', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }
    
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    let processed = 0;
    
    for (const row of rows) {
      const itemPcb = String(row['Item_PCB'] || row['Item PCB'] || '');
      const plantCode = String(row['Plant Code'] || row['Plant_Code'] || '');
      const description = row['Description'] || '';
      const model = row['Model'] || '';
      const customer = row['Cust'] || row['Customer'] || '';
      const annualPlan = parseFloat(row['ST'] || row['annual plan'] || 0);
      const leadTime = parseInt(row['Lead Time'] || plants[plantCode]?.lead_time || 0);
      
      if (itemPcb && plantCode) {
        const key = makeKey(plantCode, itemPcb);
        const plantInfo = plants[plantCode] || {};
        const perday = annualPlan > 0 ? annualPlan / 300 : 0;
        
        pcbMaster[key] = {
          item_pcb: itemPcb,
          plant_code: plantCode,
          plant_name: plantInfo.name || '',
          description: description,
          model: model,
          customer: customer,
          lead_time: leadTime || plantInfo.lead_time || 0,
          annual_plan: annualPlan,
          perday: perday,
          notes: 'SOP'
        };
        processed++;
      }
    }
    
    saveDataToFiles();
    logUpload('master', req.file.originalname, processed, 'success');
    
    res.json({ success: true, records: processed, message: 'PCB Master uploaded successfully' });
  } catch (error) {
    console.error('Master upload error:', error);
    logUpload('master', req.file?.originalname, 0, 'error', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Calculate Status
app.post('/api/calculate', async (req, res) => {
  try {
    const snapshotDate = new Date().toISOString().split('T')[0];
    let calculated = 0;
    
    for (const [key, master] of Object.entries(pcbMaster)) {
      const stock = stockData[key] || {};
      const unrestricted = stock.unrestricted || 0;
      const transit = stock.transit || 0;
      const finalStock = unrestricted + transit;
      const perday = master.perday || 0;
      
      let coverageDays = null;
      if (perday > 0) {
        coverageDays = finalStock / perday;
      }
      
      const leadTime = master.lead_time || 0;
      const remarks = classifyStatus(coverageDays, leadTime);
      
      // Get dispatch data
      const currentMonth = new Date().toISOString().slice(0, 7);
      const dispatchKey = `${key}_${currentMonth}`;
      const dispatchedQty = dispatchData[dispatchKey]?.dispatched_qty || 0;
      const scheduledQty = master.perday * 25; // Approximate monthly plan
      const adherencePct = scheduledQty > 0 ? dispatchedQty / scheduledQty : 0;
      const askingRate = classifyAskingRate(adherencePct);
      
      const statusKey = `${key}_${snapshotDate}`;
      pcbStatus[statusKey] = {
        snapshot_date: snapshotDate,
        item_pcb: master.item_pcb,
        plant_code: master.plant_code,
        plant_name: master.plant_name,
        description: master.description,
        model: master.model,
        customer: master.customer,
        lead_time: leadTime,
        perday: perday,
        annual_plan: master.annual_plan,
        dispatched_till_date: dispatchedQty,
        scheduled_till_date: scheduledQty,
        adherence_pct: adherencePct,
        asking_rate: askingRate,
        stock: unrestricted,
        transit: transit,
        final_stock: finalStock,
        coverage_days: coverageDays,
        remarks: remarks
      };
      calculated++;
    }
    
    saveDataToFiles();
    res.json({ success: true, calculated: calculated, message: 'Status calculated successfully' });
  } catch (error) {
    console.error('Calculate error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get Status with filters
app.get('/api/status', (req, res) => {
  try {
    const { date, plant, status, search, page = 1, limit = 50 } = req.query;
    const snapshotDate = date || new Date().toISOString().split('T')[0];
    
    let results = [];
    
    for (const [key, record] of Object.entries(pcbStatus)) {
      if (record.snapshot_date === snapshotDate) {
        if (plant && record.plant_code !== plant) continue;
        if (status && record.remarks !== status) continue;
        if (search && !record.item_pcb.includes(search) && !(record.description || '').includes(search)) continue;
        results.push(record);
      }
    }
    
    // Sort by remarks priority
    const priority = { 'Very Critical': 1, 'Potential Critical': 2, 'Sufficient': 3, 'High': 4, 'No Plan': 5 };
    results.sort((a, b) => (priority[a.remarks] || 99) - (priority[b.remarks] || 99));
    
    const total = results.length;
    const start = (parseInt(page) - 1) * parseInt(limit);
    const paginated = results.slice(start, start + parseInt(limit));
    
    // Summary
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

// Get Status Summary for Dashboard
app.get('/api/status/summary', (req, res) => {
  try {
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    const plantSummary = {};
    let totalPCBs = 0;
    let summary = {
      veryCritical: 0,
      potentialCritical: 0,
      sufficient: 0,
      high: 0,
      noPlan: 0
    };
    
    for (const [key, record] of Object.entries(pcbStatus)) {
      if (record.snapshot_date === snapshotDate) {
        totalPCBs++;
        
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

// Get History for item
app.get('/api/history/:item_pcb/:plant_code', (req, res) => {
  try {
    const { item_pcb, plant_code } = req.params;
    const key = makeKey(plant_code, item_pcb);
    
    const history = [];
    for (const [statusKey, record] of Object.entries(pcbStatus)) {
      if (record.item_pcb === item_pcb && record.plant_code === plant_code) {
        history.push({
          date: record.snapshot_date,
          coverage_days: record.coverage_days,
          remarks: record.remarks,
          stock: record.stock,
          transit: record.transit,
          final_stock: record.final_stock
        });
      }
    }
    
    history.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    const current = history[history.length - 1] || {};
    const info = pcbMaster[key] || {};
    
    res.json({
      success: true,
      history: history.slice(-30),
      current: current,
      info: info
    });
  } catch (error) {
    console.error('History error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get Upload Logs
app.get('/api/uploads/log', (req, res) => {
  res.json(uploadLogs.slice(0, 20));
});

// Export to Excel
app.get('/api/export/excel', (req, res) => {
  try {
    const { plant, status, search } = req.query;
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    let results = [];
    for (const [key, record] of Object.entries(pcbStatus)) {
      if (record.snapshot_date === snapshotDate) {
        if (plant && record.plant_code !== plant) continue;
        if (status && record.remarks !== status) continue;
        if (search && !record.item_pcb.includes(search) && !(record.description || '').includes(search)) continue;
        results.push(record);
      }
    }
    
    // Prepare data for Excel
    const excelData = results.map((r, index) => ({
      'S.No': index + 1,
      'Item': r.item_pcb,
      'Item_PCB': r.item_pcb,
      'Description': r.description,
      'Model': r.model,
      'Cust': r.customer,
      'Plant Code': r.plant_code,
      'Plant': r.plant_name,
      'Lead time': r.lead_time,
      'Perday': r.perday?.toFixed(2),
      'Stock': r.stock?.toFixed(0),
      'Transit': r.transit?.toFixed(0),
      'Final stock': r.final_stock?.toFixed(0),
      'Cov./Day': r.coverage_days?.toFixed(1),
      'Remarks': r.remarks,
      'Adherence %': r.adherence_pct ? (r.adherence_pct * 100).toFixed(1) + '%' : '-',
      'Asking Rate': r.asking_rate
    }));
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PCB_Detail');
    
    const filename = `pcb_coverage_${snapshotDate}.xlsx`;
    const filepath = path.join(__dirname, 'uploads', filename);
    XLSX.writeFile(wb, filepath);
    
    res.download(filepath, filename, (err) => {
      if (err) console.error('Download error:', err);
      setTimeout(() => fs.unlinkSync(filepath), 60000);
    });
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Load data and start server
loadDataFromFiles();
// Bulk import from any Excel file
app.post('/api/bulk-import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded' });
    }
    
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    let imported = 0;
    const snapshotDate = new Date().toISOString().split('T')[0];
    
    for (const row of rows) {
      const itemPcb = String(row['Item_PCB'] || row['Item PCB'] || '');
      const plantCode = String(row['Plant Code'] || row['Plant_Code'] || '');
      const stock = parseFloat(row['Stock'] || row['Current Stock'] || 0);
      const transit = parseFloat(row['Transit'] || row['In Transit'] || 0);
      const perday = parseFloat(row['Perday'] || row['Daily Consumption'] || 0);
      
      if (itemPcb && plantCode) {
        const key = makeKey(plantCode, itemPcb);
        
        // Update stock
        if (!stockData[key]) stockData[key] = {};
        stockData[key].unrestricted = stock;
        stockData[key].transit = transit;
        stockData[key].snapshot_date = snapshotDate;
        
        // Update master if perday provided
        if (perday > 0 && pcbMaster[key]) {
          pcbMaster[key].perday = perday;
        }
        
        imported++;
      }
    }
    
    saveDataToFiles();
    
    // Auto-recalculate
    // Trigger calculation for imported items
    let calculated = 0;
    for (const [key, master] of Object.entries(pcbMaster)) {
      const stock = stockData[key] || {};
      const finalStock = (stock.unrestricted || 0) + (stock.transit || 0);
      const perday = master.perday || 0;
      let coverageDays = perday > 0 ? finalStock / perday : null;
      const remarks = classifyStatus(coverageDays, master.lead_time);
      
      const statusKey = `${key}_${snapshotDate}`;
      pcbStatus[statusKey] = {
        ...pcbStatus[statusKey],
        ...master,
        snapshot_date: snapshotDate,
        stock: stock.unrestricted || 0,
        transit: stock.transit || 0,
        final_stock: finalStock,
        coverage_days: coverageDays,
        remarks: remarks
      };
      calculated++;
    }
    
    saveDataToFiles();
    
    res.json({ 
      success: true, 
      imported: imported,
      calculated: calculated,
      message: `Imported ${imported} records and recalculated ${calculated} items` 
    });
  } catch (error) {
    console.error('Bulk import error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});
// Force refresh from uploaded Excel files
app.post('/api/refresh-from-excel', async (req, res) => {
  try {
    const { fileType, filePath } = req.body;
    
    if (!fs.existsSync(filePath)) {
      return res.status(400).json({ success: false, error: 'File not found' });
    }
    
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    
    let processed = 0;
    
    if (fileType === 'stock') {
      // Clear existing stock data for this snapshot
      const snapshotDate = new Date().toISOString().split('T')[0];
      
      for (const row of rows) {
        const plant = String(row.Plant || '');
        const material = String(row.Material || '');
        const unrestricted = parseFloat(row.Unrestricted) || 0;
        
        if (plant && material) {
          const key = makeKey(plant, material);
          if (!stockData[key]) stockData[key] = {};
          stockData[key].unrestricted = unrestricted;
          stockData[key].snapshot_date = snapshotDate;
          stockData[key].plant = plant;
          stockData[key].material = material;
          processed++;
        }
      }
    } else if (fileType === 'master') {
      for (const row of rows) {
        const itemPcb = String(row.Item_PCB || row['Item PCB'] || '');
        const plantCode = String(row['Plant Code'] || row.Plant_Code || '');
        
        if (itemPcb && plantCode) {
          const key = makeKey(plantCode, itemPcb);
          const annualPlan = parseFloat(row.ST || row['annual plan'] || 0);
          const perday = annualPlan > 0 ? annualPlan / 300 : 0;
          
          pcbMaster[key] = {
            item_pcb: itemPcb,
            plant_code: plantCode,
            plant_name: plants[plantCode]?.name || '',
            description: row.Description || '',
            model: row.Model || '',
            customer: row.Cust || row.Customer || '',
            lead_time: parseInt(row['Lead Time']) || plants[plantCode]?.lead_time || 0,
            annual_plan: annualPlan,
            perday: perday
          };
          processed++;
        }
      }
    }
    
    saveDataToFiles();
    res.json({ success: true, records: processed, message: `Refreshed ${processed} records` });
  } catch (error) {
    console.error('Refresh error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});
app.listen(PORT, () => {
  console.log(`✅ PCB Portal running on http://localhost:${PORT}`);
  console.log(`📁 Plants loaded: ${Object.keys(plants).length} plants`);
  console.log(`📊 Stock records: ${Object.keys(stockData).length}`);
  console.log(`📋 PCB Master records: ${Object.keys(pcbMaster).length}`);
  console.log(`🌐 Open your browser to view the dashboard`);
});

module.exports = app;