document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const fileUploadComp = document.getElementById('fileUpload');
    const statusText = document.getElementById('statusText');
    const generateBtn = document.getElementById('generateBtn');
    const outputSection = document.getElementById('outputSection');
    const summaryTableBody = document.querySelector('#summaryTable tbody');
    const totalShiftDisplay = document.getElementById('totalShift');
    const totalShiftChargesDisplay = document.getElementById('totalShiftCharges');
    const totalOtDisplay = document.getElementById('totalOt');
    const totalOtChargesDisplay = document.getElementById('totalOtCharges');
    const totalBdDisplay = document.getElementById('totalBd');
    const totalBdChargesDisplay = document.getElementById('totalBdCharges');
    const grandTotalDisplay = document.getElementById('grandTotal');
    const errorContainer = document.getElementById('errorContainer');
    const errorMessage = document.getElementById('errorMessage');
    const exportBtn = document.getElementById('exportBtn');

    // Report Section DOM Elements
    const generateReportBtn = document.getElementById('generateReportBtn');
    const reportSection = document.getElementById('reportSection');
    const reportTableBody = document.querySelector('#reportTable tbody');
    const reportGrandTotalDisplay = document.getElementById('reportGrandTotal');
    const exportReportBtn = document.getElementById('exportReportBtn');

    let uploadedFile = null;
    let summaryData = [];
    let reportData = [];

    // ========================================
    // Rate Card Data (Machine-based, no vendor dependency)
    // ========================================
    const RATE_CARD = [
        { crane_name: "400 T", hourly_charges: 61538.46, ot_amount_per_hr: 5128.21, penalty_per_hr: 1282 },
        { crane_name: "300 T", hourly_charges: 4246.80, ot_amount_per_hr: 3000, penalty_per_hr: 1061.7 },
        { crane_name: "160 T", hourly_charges: 2564.10, ot_amount_per_hr: 1500, penalty_per_hr: 641 },
        { crane_name: "100 T", hourly_charges: 1602.56, ot_amount_per_hr: 993, penalty_per_hr: 400.64 },
        { crane_name: "80 T", hourly_charges: 1362.18, ot_amount_per_hr: 813, penalty_per_hr: 340.54 },
        { crane_name: "55 T", hourly_charges: 1442.31, ot_amount_per_hr: 1442, penalty_per_hr: 360.58 },
        { crane_name: "40 T", hourly_charges: 1400.00, ot_amount_per_hr: 780, penalty_per_hr: 350 },
        { crane_name: "F-15", hourly_charges: 362.50, ot_amount_per_hr: 340, penalty_per_hr: 90.63 },
        { crane_name: "TRX-23", hourly_charges: 487.50, ot_amount_per_hr: 390, penalty_per_hr: 121.88 },
        { crane_name: "F-20", hourly_charges: 762.50, ot_amount_per_hr: 460, penalty_per_hr: 190.63 }
    ];

    // ========================================
    // Cost Center Data
    // ========================================
    const CC_DATA = {
        "25225": { "name": "A-F BF", "division": "IRON MAKING" },
        "20300": { "name": "A-F BF, IEM-1", "division": "IRON MAKING" },
        "25242": { "name": "A-F BF, IEM-2", "division": "IRON MAKING" },
        "25223": { "name": "C-BF", "division": "IRON MAKING" },
        "23222": { "name": "E-BF", "division": "IRON MAKING" },
        "25226": { "name": "F-BF", "division": "IRON MAKING" },
        "25320": { "name": "G-BF", "division": "IRON MAKING" },
        "20420": { "name": "G-BF, IEM", "division": "IRON MAKING" },
        "25230": { "name": "G-BF, Belt Group", "division": "IRON MAKING" },
        "40420": { "name": "G-BF-SSP", "division": "IRON MAKING" },
        "25325": { "name": "H-BF (Mechanical)", "division": "IRON MAKING" },
        "25345": { "name": "H-BF (Operation)-1", "division": "IRON MAKING" },
        "20520": { "name": "H-BF (Operation)-2", "division": "IRON MAKING" },
        "25981": { "name": "I-BF", "division": "IRON MAKING" },
        "25021": { "name": "Coke Plant, Battery 5,6,7", "division": "IRON MAKING" },
        "25032": { "name": "Coke Plant, BAT 10/11", "division": "IRON MAKING" },
        "22203": { "name": "Coke Plant, CDQ 10/11", "division": "IRON MAKING" },
        "25022": { "name": "Coke Plant Bat 8,9", "division": "IRON MAKING" },
        "25023": { "name": "Coke Plant Old BPP", "division": "IRON MAKING" },
        "25033": { "name": "Coke Plant New BPP", "division": "IRON MAKING" },
        "20150": { "name": "Coke Plant- 1", "division": "IRON MAKING" },
        "20112": { "name": "Coke Plant- 2", "division": "IRON MAKING" },
        "20113": { "name": "Coke Plant- 3", "division": "IRON MAKING" },
        "20140": { "name": "Coke Plant- 4", "division": "IRON MAKING" },
        "25042": { "name": "Coke Plant- 5", "division": "IRON MAKING" },
        "25035": { "name": "Coke Plant- 6", "division": "IRON MAKING" },
        "25036": { "name": "Coke Plant- 7", "division": "IRON MAKING" },
        "25043": { "name": "Coke Plant, IEM", "division": "IRON MAKING" },
        "25121": { "name": "Sinter Plant-1 (Mech.)", "division": "IRON MAKING" },
        "20230": { "name": "Sinter Plant-1 (Operation-1)", "division": "IRON MAKING" },
        "23020": { "name": "Sinter Plant-1 (Operation-2)", "division": "IRON MAKING" },
        "25141": { "name": "Sinter Plant-1 (IEM)", "division": "IRON MAKING" },
        "25122": { "name": "Sinter Plant-2", "division": "IRON MAKING" },
        "25142": { "name": "Sinter Plant-2 (IEM)", "division": "IRON MAKING" },
        "25123": { "name": "Sinter Plant-3", "division": "IRON MAKING" },
        "25143": { "name": "Sinter Plant-3 (IEM-1)", "division": "IRON MAKING" },
        "23325": { "name": "Sinter Plant-3 (IEM-2)", "division": "IRON MAKING" },
        "25124": { "name": "Sinter Plant-4", "division": "IRON MAKING" },
        "20270": { "name": "Sinter Plant-4 (IEM-1)", "division": "IRON MAKING" },
        "25144": { "name": "Sinter Plant-4 (IEM-2)", "division": "IRON MAKING" },
        "25120": { "name": "RMBB-1 (Mech)", "division": "IRON MAKING" },
        "20220": { "name": "RMBB-1 (Operation)", "division": "IRON MAKING" },
        "25125": { "name": "RMBB#2-1", "division": "IRON MAKING" },
        "25145": { "name": "RMBB#2-2", "division": "IRON MAKING" },
        "20250": { "name": "RMBB#2-3", "division": "IRON MAKING" },
        "25964": { "name": "Pellet Plant (Mech)-1", "division": "IRON MAKING" },
        "25965": { "name": "Pellet Plant (Mech)-2", "division": "IRON MAKING" },
        "20651": { "name": "Pellet Plant Operation-1", "division": "IRON MAKING" },
        "20661": { "name": "Pellet Plant Operation-2", "division": "IRON MAKING" },
        "25974": { "name": "Pellet Plant-IEM", "division": "IRON MAKING" },
        "25975": { "name": "Pellet Plant-FME", "division": "IRON MAKING" },
        "25222": { "name": "FMM-IM", "division": "IRON MAKING" },
        "25421": { "name": "LD#1- 1", "division": "STEEL MAKING" },
        "25420": { "name": "LD#1- 2", "division": "STEEL MAKING" },
        "25422": { "name": "LD#1- 3", "division": "STEEL MAKING" },
        "23221": { "name": "LD#1- 4", "division": "STEEL MAKING" },
        "21030": { "name": "LD#1-Operation", "division": "STEEL MAKING" },
        "25441": { "name": "LD#1-IEM", "division": "STEEL MAKING" },
        "25443": { "name": "LD#1-FME", "division": "STEEL MAKING" },
        "25521": { "name": "LD#2- 1", "division": "STEEL MAKING" },
        "25522": { "name": "LD#2- 2", "division": "STEEL MAKING" },
        "25221": { "name": "LD#2- 3", "division": "STEEL MAKING" },
        "25952": { "name": "LD#2- 4", "division": "STEEL MAKING" },
        "22020": { "name": "LD#2- 5", "division": "STEEL MAKING" },
        "22022": { "name": "LD#2- 6", "division": "STEEL MAKING" },
        "25541": { "name": "LD#2, IEM", "division": "STEEL MAKING" },
        "25543": { "name": "LD#2, FME", "division": "STEEL MAKING" },
        "22021": { "name": "LD#2, BIM", "division": "STEEL MAKING" },
        "25966": { "name": "LD#3 (Mechanical) -1", "division": "STEEL MAKING" },
        "25967": { "name": "LD#3 (Mechanical) -2", "division": "STEEL MAKING" },
        "22412": { "name": "LD#3 (Operation)", "division": "STEEL MAKING" },
        "25231": { "name": "RMM Mech-1", "division": "STEEL MAKING" },
        "20161": { "name": "RMM Mech-2", "division": "STEEL MAKING" },
        "20660": { "name": "RMM Mech-3", "division": "STEEL MAKING" },
        "20004": { "name": "RMM OPN", "division": "STEEL MAKING" },
        "20061": { "name": "RMM Coke", "division": "STEEL MAKING" },
        "25935": { "name": "Lime Plant", "division": "STEEL MAKING" },
        "25695": { "name": "Belt Group", "division": "STEEL MAKING" },
        "25355": { "name": "SPP", "division": "STEEL MAKING" },
        "25335": { "name": "SGDP", "division": "STEEL MAKING" },
        "25620": { "name": "HSM", "division": "MILLS" },
        "23332": { "name": "HSM, FME-1", "division": "MILLS" },
        "25641": { "name": "HSM, FME-2", "division": "MILLS" },
        "25728": { "name": "CRM-1", "division": "MILLS" },
        "25721": { "name": "CRM-2", "division": "MILLS" },
        "25720": { "name": "CRM-3", "division": "MILLS" },
        "23223": { "name": "CRM-4", "division": "MILLS" },
        "25746": { "name": "CRM-5", "division": "MILLS" },
        "22211": { "name": "CRM ARP", "division": "MILLS" },
        "22230": { "name": "CRM BAF", "division": "MILLS" },
        "22250": { "name": "CRM F&D", "division": "MILLS" },
        "25969": { "name": "TSCR", "division": "MILLS" },
        "25050": { "name": "TSCR-FME", "division": "MILLS" },
        "25031": { "name": "MM", "division": "MILLS" },
        "25051": { "name": "MM, IEM", "division": "MILLS" },
        "25820": { "name": "WRM", "division": "MILLS" },
        "25920": { "name": "NBM-1", "division": "MILLS" },
        "21330": { "name": "NBM-2", "division": "MILLS" },
        "25431": { "name": "PH#3", "division": "POWER SYSTEM" },
        "23031": { "name": "PH#3 -2", "division": "POWER SYSTEM" },
        "25531": { "name": "PH#4", "division": "POWER SYSTEM" },
        "25631": { "name": "PH#5 -1", "division": "POWER SYSTEM" },
        "25634": { "name": "PH#5 -2", "division": "POWER SYSTEM" },
        "23065": { "name": "BH#2", "division": "POWER SYSTEM" },
        "25731": { "name": "BH#3", "division": "POWER SYSTEM" },
        "23069": { "name": "BH#4", "division": "POWER SYSTEM" },
        "25732": { "name": "BPH-1", "division": "POWER SYSTEM" },
        "23054": { "name": "BPH-2", "division": "POWER SYSTEM" },
        "23079": { "name": "BPH-3", "division": "POWER SYSTEM" },
        "23114": { "name": "BPH-4", "division": "POWER SYSTEM" },
        "23220": { "name": "FMM", "division": "OTHER" },
        "23224": { "name": "Belt Group-1", "division": "OTHER" },
        "20130": { "name": "Belt Group-2", "division": "OTHER" },
        "29450": { "name": "GORC", "division": "OTHER" },
        "23433": { "name": "SMD", "division": "OTHER" },
        "20006": { "name": "RMM Logistics", "division": "OTHER" },
        "26302": { "name": "Central Hub", "division": "OTHER" },
        "29443": { "name": "Fire Brigade", "division": "OTHER" },
        "23130": { "name": "Water Manag", "division": "OTHER" },
        "22306": { "name": "CRM BARA", "division": "OTHER" },
        "41120": { "name": "IBMD-1", "division": "OTHER" },
        "41130": { "name": "IBMD-2", "division": "OTHER" },
        "41140": { "name": "IBMD-3", "division": "OTHER" },
        "23321": { "name": "FME-1", "division": "OTHER" },
        "23322": { "name": "FME-2", "division": "OTHER" },
        "25140": { "name": "FME-3", "division": "OTHER" },
        "25251": { "name": "FME-4", "division": "OTHER" },
        "20316": { "name": "FME-5", "division": "OTHER" },
        "25934": { "name": "FMD MECHANICAL", "division": "OTHER" },
        "23017": { "name": "FMD-1", "division": "OTHER" },
        "25018": { "name": "FMD-2", "division": "OTHER" },
        "23016": { "name": "FMD-3", "division": "OTHER" },
        "23015": { "name": "FMD GCP-1", "division": "OTHER" },
        "25931": { "name": "FMD GCP-2", "division": "OTHER" },
        "23018": { "name": "FMD GAS LINE", "division": "OTHER" },
        "25932": { "name": "FMD GMS", "division": "OTHER" },
        "23012": { "name": "FMD GHB", "division": "OTHER" },
        "25635": { "name": "40 MW-1", "division": "OTHER" },
        "25655": { "name": "40 MW-2", "division": "OTHER" },
        "29162": { "name": "E&P", "division": "OTHER" },
        "29200": { "name": "E&P-HSM", "division": "OTHER" },
        "23086": { "name": "T-30", "division": "OTHER" },
        "23084": { "name": "MPDS", "division": "OTHER" },
        "25351": { "name": "MRP, IEM", "division": "OTHER" },
        "29010": { "name": "R&D-1", "division": "OTHER" },
        "29016": { "name": "R&D-2", "division": "OTHER" },
        "29011": { "name": "TWSGS", "division": "OTHER" },
        "21120": { "name": "Engg. Services (LP)", "division": "OTHER" },
        "23615": { "name": "Engg. Services (AC)", "division": "OTHER" },
        "22207": { "name": "Engg. Services (AC-CRM)", "division": "OTHER" },
        "20260": { "name": "Engg. Services (PP)", "division": "OTHER" },
        "22220": { "name": "Engg. Services (BIM-1)", "division": "OTHER" },
        "23711": { "name": "Engg. Services (BIM-2)", "division": "OTHER" },
        "22142": { "name": "Engg. Services (BIM-3)", "division": "OTHER" },
        "20240": { "name": "Engg. Services (BIM-4)", "division": "OTHER" },
        "22130": { "name": "Engg. Services (BIM-5)", "division": "OTHER" },
        "23151": { "name": "Engg. Services (BIM-6)", "division": "OTHER" },
        "23420": { "name": "Engg. Services (BIM SMD-1)", "division": "OTHER" },
        "25990": { "name": "Engg. Services (BIM SMD-2)", "division": "OTHER" },
        "21021": { "name": "Engg. Services (BIM LD#1)", "division": "OTHER" },
        "20710": { "name": "Engg. Services (BIM I-BF)", "division": "OTHER" },
        "23431": { "name": "Segment Shop", "division": "OTHER" },
        "21003": { "name": "EQMS, SME", "division": "OTHER" },
        "23716": { "name": "EQMS-1", "division": "OTHER" },
        "23712": { "name": "EQMS-2", "division": "OTHER" },
        "23718": { "name": "EQMS-3", "division": "OTHER" },
        "26202": { "name": "Delivery Management", "division": "OTHER" }
    };

    // ========================================
    // File Upload Handler
    // ========================================
    fileUploadComp.addEventListener('change', (e) => {
        const file = e.target.files[0];
        handleFileSelect(file);
    });

    function handleFileSelect(file) {
        if (!file) return;
        uploadedFile = file;
        statusText.textContent = `Selected: ${file.name}`;
        generateBtn.disabled = false;
        errorContainer.classList.add('hidden');
        outputSection.classList.add('hidden');
        reportSection.classList.add('hidden');
    }

    function resetFileUpload() {
        uploadedFile = null;
        fileUploadComp.value = '';
        statusText.textContent = '';
        generateBtn.disabled = true;
        outputSection.classList.add('hidden');
        reportSection.classList.add('hidden');
        errorContainer.classList.add('hidden');
    }

    // ========================================
    // Generate Summary
    // ========================================
    generateBtn.addEventListener('click', async () => {
        if (!uploadedFile) return;

        try {
            const rawRecords = await parseExcel(uploadedFile);
            console.log("Extracted Records:", rawRecords); // Debug log

            // For now, just display raw extracted data (no aggregation/calculation)
            renderRawData(rawRecords);
            outputSection.classList.remove('hidden');
        } catch (err) {
            showError("Error processing file: " + err.message);
            console.error(err);
        }
    });

    function showError(msg) {
        errorMessage.textContent = msg;
        errorContainer.classList.remove('hidden');
    }

    // ========================================
    // Parse Excel - Extracts data from all sheets
    // ========================================
    function parseExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const allRecords = [];

                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

                        // Extract Machine Name and Vendor Name from header area (first 5 rows)
                        const sheetMeta = extractSheetMetadata(jsonData);

                        // Find data table and extract rows
                        const dataRows = extractDataRows(jsonData, sheetMeta);
                        allRecords.push(...dataRows);
                    });

                    resolve(allRecords);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Extracts Machine Name from cell A3 of the sheet.
     * Cell format: "MACHINE NAME:- TRX-23    VENDOR NAME:-    SWASTIK    WORK ORDER No.:-    123S214S    BILLING MONTH:-DEC"
     * We extract just "TRX-23" (text between "MACHINE NAME:-" and "VENDOR NAME")
     */
    function extractSheetMetadata(rows) {
        let machineName = "Unknown";

        // Machine Name is in Row 3 (index 2), cell A3
        if (rows.length > 2) {
            const row3 = rows[2]; // Row 3 (0-indexed = 2)

            // Check cell A3
            if (row3 && row3[0]) {
                const cellA3 = String(row3[0]).trim();

                // Extract machine name: text after "MACHINE NAME:-" and before "VENDOR NAME"
                const machineMatch = cellA3.match(/MACHINE\s*NAME[:\-\s]+([A-Za-z0-9\-\(\)]+)/i);
                if (machineMatch) {
                    machineName = machineMatch[1].trim();
                }
            }
        }

        return { machineName };
    }

    /**
     * Finds the data table header row and extracts data rows.
     * Looks for columns: SHIFT QTY, OT QTY, BREAKDOWN HOUR, Cost Center
     */
    function extractDataRows(rows, sheetMeta) {
        const records = [];
        let headerRowIndex = -1;
        let colIndices = {
            shiftQty: -1,
            otQty: -1,
            bdHour: -1,
            costCenter: -1
        };

        // Find header row
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowStrings = row.map(cell => String(cell).toUpperCase().trim());

            // Look for key columns
            const shiftIdx = rowStrings.findIndex(s => s.includes("SHIFT") && s.includes("QTY"));
            const otIdx = rowStrings.findIndex(s => s.includes("OT") && s.includes("QTY"));
            const bdIdx = rowStrings.findIndex(s => s.includes("BREAKDOWN") && s.includes("HOUR"));
            const ccIdx = rowStrings.findIndex(s => s === "COST CENTER" || s === "CC");

            if (shiftIdx !== -1 && otIdx !== -1 && ccIdx !== -1) {
                headerRowIndex = i;
                colIndices = {
                    shiftQty: shiftIdx,
                    otQty: otIdx,
                    bdHour: bdIdx,
                    costCenter: ccIdx
                };
                break;
            }
        }

        if (headerRowIndex === -1) {
            console.warn("Could not find header row in sheet. Metadata:", sheetMeta);
            return records;
        }

        // Extract data rows (rows after header)
        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            const cc = String(row[colIndices.costCenter] || "").trim();

            // Skip empty rows or rows without a valid cost center
            if (!cc || cc === "" || cc.toUpperCase() === "COST CENTER" || cc.toUpperCase() === "CC") continue;

            const shiftQty = parseFloat(row[colIndices.shiftQty]) || 0;
            const otQty = parseFloat(row[colIndices.otQty]) || 0;
            const bdHour = colIndices.bdHour !== -1 ? (parseFloat(row[colIndices.bdHour]) || 0) : 0;

            records.push({
                machine: sheetMeta.machineName,
                cc: cc,
                shiftQty: shiftQty,
                otQty: otQty,
                bdHour: bdHour
            });
        }

        return records;
    }

    // ========================================
    // Process Data - Calculate costs and aggregate
    // ========================================
    function processData(rawRecords) {
        const aggregationMap = {};

        rawRecords.forEach(record => {
            // Find rate for this machine (no vendor dependency now)
            const rate = findRate(record.machine);

            // Calculate cost
            const shiftCost = record.shiftQty * rate.hourly_charges;
            const otCost = record.otQty * rate.ot_amount_per_hr;
            const penalty = record.bdHour * rate.penalty_per_hr;
            const totalCost = shiftCost + otCost - penalty;

            // Aggregation key: machine + cc
            const key = `${record.machine}|||${record.cc}`;

            if (!aggregationMap[key]) {
                aggregationMap[key] = {
                    machine: record.machine,
                    cc: record.cc,
                    shiftQty: 0,
                    otQty: 0,
                    bdHour: 0,
                    totalAmt: 0
                };
            }

            aggregationMap[key].shiftQty += record.shiftQty;
            aggregationMap[key].otQty += record.otQty;
            aggregationMap[key].bdHour += record.bdHour;
            aggregationMap[key].totalAmt += totalCost;
        });

        // Convert to array
        summaryData = Object.values(aggregationMap);
        return summaryData;
    }

    /**
     * Finds the rate card entry for a given machine name.
     * Uses fuzzy matching to handle slight variations in naming (e.g., "F-15(1)" matches "F-15").
     */
    function findRate(machineName) {
        const defaultRate = { hourly_charges: 0, ot_amount_per_hr: 0, penalty_per_hr: 0 };

        // Normalize machine name for matching (remove parentheses, spaces, numbers in parens)
        const normMachine = machineName.toLowerCase().replace(/\([^)]*\)/g, '').replace(/[\s]/g, '').trim();

        for (const rateEntry of RATE_CARD) {
            const entryCrane = rateEntry.crane_name.toLowerCase().replace(/[\s]/g, '').trim();

            // Match machine name (handle variations like "F-15(1)" matching "F-15")
            if (normMachine.includes(entryCrane) || entryCrane.includes(normMachine)) {
                return rateEntry;
            }
        }

        console.warn(`Rate not found for Machine: "${machineName}". Using default rates.`);
        return defaultRate;
    }

    // ========================================
    // Render Raw Extracted Data (with calculation breakdown)
    // ========================================
    function renderRawData(records) {
        summaryTableBody.innerHTML = '';
        let grandShift = 0, grandOt = 0, grandBd = 0;
        let grandShiftCharges = 0, grandOtCharges = 0, grandBdCharges = 0, grandTotal = 0;

        records.forEach(item => {
            const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };

            // Get rate card for this machine
            const rate = findRate(item.machine);

            // Calculate charges
            const shiftCharges = item.shiftQty * rate.hourly_charges;
            const otCharges = item.otQty * rate.ot_amount_per_hr;
            const bdCharges = item.bdHour * rate.penalty_per_hr;
            const netTotal = shiftCharges + otCharges - bdCharges;

            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${item.machine}</td>
                <td>${item.cc}</td>
                <td>${ccInfo.name}</td>
                <td>${ccInfo.division}</td>
                <td>${item.shiftQty.toFixed(2)}</td>
                <td>₹ ${rate.hourly_charges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>₹ ${shiftCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>${item.otQty.toFixed(2)}</td>
                <td>₹ ${rate.ot_amount_per_hr.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>₹ ${otCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>${item.bdHour.toFixed(2)}</td>
                <td>₹ ${rate.penalty_per_hr.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>₹ ${bdCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td>₹ ${netTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            `;
            summaryTableBody.appendChild(row);

            grandShift += item.shiftQty;
            grandOt += item.otQty;
            grandBd += item.bdHour;
            grandShiftCharges += shiftCharges;
            grandOtCharges += otCharges;
            grandBdCharges += bdCharges;
            grandTotal += netTotal;
        });

        totalShiftDisplay.textContent = grandShift.toFixed(2);
        totalShiftChargesDisplay.textContent = `₹ ${grandShiftCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
        totalOtDisplay.textContent = grandOt.toFixed(2);
        totalOtChargesDisplay.textContent = `₹ ${grandOtCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
        totalBdDisplay.textContent = grandBd.toFixed(2);
        totalBdChargesDisplay.textContent = `₹ ${grandBdCharges.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
        grandTotalDisplay.textContent = `₹ ${grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

        // Store for later use
        summaryData = records;
    }

    // ========================================
    // Render Table (with calculations)
    // ========================================
    function renderTable(data) {
        summaryTableBody.innerHTML = '';
        let grandShift = 0, grandOt = 0, grandBd = 0, grandTotal = 0;

        data.forEach(item => {
            const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${item.machine}</td>
                <td>${item.cc}</td>
                <td>${ccInfo.name}</td>
                <td>${ccInfo.division}</td>
                <td>${item.shiftQty.toFixed(2)}</td>
                <td>${item.otQty.toFixed(2)}</td>
                <td>${item.bdHour.toFixed(2)}</td>
                <td>₹ ${item.totalAmt.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            `;
            summaryTableBody.appendChild(row);

            grandShift += item.shiftQty;
            grandOt += item.otQty;
            grandBd += item.bdHour;
            grandTotal += item.totalAmt;
        });

        totalShiftDisplay.textContent = grandShift.toFixed(2);
        totalOtDisplay.textContent = grandOt.toFixed(2);
        totalBdDisplay.textContent = grandBd.toFixed(2);
        grandTotalDisplay.textContent = `₹ ${grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }

    // ========================================
    // Export to Excel
    // ========================================
    exportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) return;

        const exportData = summaryData.map(item => {
            const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
            return {
                "Machine": item.machine,
                "CC": item.cc,
                "CC Name": ccInfo.name,
                "Division": ccInfo.division,
                "Shift Qty": item.shiftQty,
                "OT Qty": item.otQty,
                "BD Hour": item.bdHour,
                "Total Amt": item.totalAmt
            };
        });

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Crane Billing Summary");
        XLSX.writeFile(wb, "crane_billing_summary.xlsx");
    });

    // ========================================
    // Generate Report - Aggregate by Cost Center
    // ========================================
    generateReportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) {
            showError("No data available. Please generate summary first.");
            return;
        }

        // Aggregate by Cost Center
        const ccAggregation = {};

        summaryData.forEach(item => {
            // Get rate for this machine to calculate total
            const rate = findRate(item.machine);
            const shiftCharges = item.shiftQty * rate.hourly_charges;
            const otCharges = item.otQty * rate.ot_amount_per_hr;
            const bdCharges = item.bdHour * rate.penalty_per_hr;
            const netTotal = shiftCharges + otCharges - bdCharges;

            if (!ccAggregation[item.cc]) {
                const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
                ccAggregation[item.cc] = {
                    cc: item.cc,
                    ccName: ccInfo.name,
                    division: ccInfo.division,
                    totalAmount: 0
                };
            }

            ccAggregation[item.cc].totalAmount += netTotal;
        });

        // Convert to array and sort by division, then CC
        reportData = Object.values(ccAggregation).sort((a, b) => {
            if (a.division !== b.division) {
                return a.division.localeCompare(b.division);
            }
            return a.cc.localeCompare(b.cc);
        });

        // Render Report Table
        renderReport(reportData);
        reportSection.classList.remove('hidden');
    });

    // ========================================
    // Render Report Table
    // ========================================
    function renderReport(data) {
        reportTableBody.innerHTML = '';
        let grandTotal = 0;

        data.forEach(item => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${item.division}</td>
                <td>${item.cc}</td>
                <td>${item.ccName}</td>
                <td>₹ ${item.totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            `;
            reportTableBody.appendChild(row);
            grandTotal += item.totalAmount;
        });

        reportGrandTotalDisplay.textContent = `₹ ${grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }

    // ========================================
    // Export Report to Excel
    // ========================================
    exportReportBtn.addEventListener('click', () => {
        if (reportData.length === 0) return;

        const exportData = reportData.map(item => ({
            "Division": item.division,
            "CC Number": item.cc,
            "CC Name": item.ccName,
            "Total Amount": item.totalAmount
        }));

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Cost Center Report");
        XLSX.writeFile(wb, "cost_center_report.xlsx");
    });
});
