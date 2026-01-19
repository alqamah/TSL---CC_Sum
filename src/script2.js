document.addEventListener('DOMContentLoaded', () => {
    const fileUploadComp = document.getElementById('fileUpload');
    const fileNameDisplay = document.getElementById('fileName');
    const fileInfoDiv = document.getElementById('fileInfo');
    const removeFileBtn = document.getElementById('removeFile');
    const generateBtn = document.getElementById('generateBtn');
    const outputSection = document.getElementById('outputSection');
    const summaryTableBody = document.querySelector('#summaryTable tbody');
    const grandTotalDisplay = document.getElementById('grandTotal');
    const errorContainer = document.getElementById('errorContainer');
    const errorMessage = document.getElementById('errorMessage');
    const exportBtn = document.getElementById('exportBtn');

    let uploadedFile = null;
    let summaryData = [];

    const CC_DATA = {
        "25225": {
            "name": "A-F BF",
            "division": "IRON MAKING"
        },
        "20300": {
            "name": "A-F BF, IEM-1",
            "division": "IRON MAKING"
        },
        "25242": {
            "name": "A-F BF, IEM-2",
            "division": "IRON MAKING"
        },
        "25223": {
            "name": "C-BF",
            "division": "IRON MAKING"
        },
        "23222": {
            "name": "E-BF",
            "division": "IRON MAKING"
        },
        "25226": {
            "name": "F-BF",
            "division": "IRON MAKING"
        },
        "25320": {
            "name": "G-BF",
            "division": "IRON MAKING"
        },
        "20420": {
            "name": "G-BF, IEM",
            "division": "IRON MAKING"
        },
        "25230": {
            "name": "G-BF, Belt Group",
            "division": "IRON MAKING"
        },
        "40420": {
            "name": "G-BF-SSP",
            "division": "IRON MAKING"
        },
        "25325": {
            "name": "H-BF (Mechanical)",
            "division": "IRON MAKING"
        },
        "25345": {
            "name": "H-BF (Operation)-1",
            "division": "IRON MAKING"
        },
        "20520": {
            "name": "H-BF (Operation)-2",
            "division": "IRON MAKING"
        },
        "25981": {
            "name": "I-BF",
            "division": "IRON MAKING"
        },
        "25021": {
            "name": "Coke Plant, Battery 5,6,7",
            "division": "IRON MAKING"
        },
        "25032": {
            "name": "Coke Plant, BAT 10/11",
            "division": "IRON MAKING"
        },
        "22203": {
            "name": "Coke Plant, CDQ 10/11",
            "division": "IRON MAKING"
        },
        "25022": {
            "name": "Coke Plant Bat 8,9",
            "division": "IRON MAKING"
        },
        "25023": {
            "name": "Coke Plant Old BPP",
            "division": "IRON MAKING"
        },
        "25033": {
            "name": "Coke Plant New BPP",
            "division": "IRON MAKING"
        },
        "20150": {
            "name": "Coke Plant- 1",
            "division": "IRON MAKING"
        },
        "20112": {
            "name": "Coke Plant- 2",
            "division": "IRON MAKING"
        },
        "20113": {
            "name": "Coke Plant- 3",
            "division": "IRON MAKING"
        },
        "20140": {
            "name": "Coke Plant- 4",
            "division": "IRON MAKING"
        },
        "25042": {
            "name": "Coke Plant- 5",
            "division": "IRON MAKING"
        },
        "25035": {
            "name": "Coke Plant- 6",
            "division": "IRON MAKING"
        },
        "25036": {
            "name": "Coke Plant- 7",
            "division": "IRON MAKING"
        },
        "25043": {
            "name": "Coke Plant, IEM",
            "division": "IRON MAKING"
        },
        "25121": {
            "name": "Sinter Plant-1 (Mech.)",
            "division": "IRON MAKING"
        },
        "20230": {
            "name": "Sinter Plant-1 (Operation-1)",
            "division": "IRON MAKING"
        },
        "23020": {
            "name": "Sinter Plant-1 (Operation-2)",
            "division": "IRON MAKING"
        },
        "25141": {
            "name": "Sinter Plant-1 (IEM)",
            "division": "IRON MAKING"
        },
        "25122": {
            "name": "Sinter Plant-2",
            "division": "IRON MAKING"
        },
        "25142": {
            "name": "Sinter Plant-2 (IEM)",
            "division": "IRON MAKING"
        },
        "25123": {
            "name": "Sinter Plant-3",
            "division": "IRON MAKING"
        },
        "25143": {
            "name": "Sinter Plant-3 (IEM-1)",
            "division": "IRON MAKING"
        },
        "23325": {
            "name": "Sinter Plant-3 (IEM-2)",
            "division": "IRON MAKING"
        },
        "25124": {
            "name": "Sinter Plant-4",
            "division": "IRON MAKING"
        },
        "20270": {
            "name": "Sinter Plant-4 (IEM-1)",
            "division": "IRON MAKING"
        },
        "25144": {
            "name": "Sinter Plant-4 (IEM-2)",
            "division": "IRON MAKING"
        },
        "25120": {
            "name": "RMBB-1 (Mech)",
            "division": "IRON MAKING"
        },
        "20220": {
            "name": "RMBB-1 (Operation)",
            "division": "IRON MAKING"
        },
        "25125": {
            "name": "RMBB#2-1",
            "division": "IRON MAKING"
        },
        "25145": {
            "name": "RMBB#2-2",
            "division": "IRON MAKING"
        },
        "20250": {
            "name": "RMBB#2-3",
            "division": "IRON MAKING"
        },
        "25964": {
            "name": "Pellet Plant (Mech)-1",
            "division": "IRON MAKING"
        },
        "25965": {
            "name": "Pellet Plant (Mech)-2",
            "division": "IRON MAKING"
        },
        "20651": {
            "name": "Pellet Plant Operation-1",
            "division": "IRON MAKING"
        },
        "20661": {
            "name": "Pellet Plant Operation-2",
            "division": "IRON MAKING"
        },
        "25974": {
            "name": "Pellet Plant-IEM",
            "division": "IRON MAKING"
        },
        "25975": {
            "name": "Pellet Plant-FME",
            "division": "IRON MAKING"
        },
        "25222": {
            "name": "FMM-IM",
            "division": "IRON MAKING"
        },
        "25421": {
            "name": "LD#1- 1",
            "division": "STEEL MAKING"
        },
        "25420": {
            "name": "LD#1- 2",
            "division": "STEEL MAKING"
        },
        "25422": {
            "name": "LD#1- 3",
            "division": "STEEL MAKING"
        },
        "23221": {
            "name": "LD#1- 4",
            "division": "STEEL MAKING"
        },
        "21030": {
            "name": "LD#1-Operation",
            "division": "STEEL MAKING"
        },
        "25441": {
            "name": "LD#1-IEM",
            "division": "STEEL MAKING"
        },
        "25443": {
            "name": "LD#1-FME",
            "division": "STEEL MAKING"
        },
        "25521": {
            "name": "LD#2- 1",
            "division": "STEEL MAKING"
        },
        "25522": {
            "name": "LD#2- 2",
            "division": "STEEL MAKING"
        },
        "25221": {
            "name": "LD#2- 3",
            "division": "STEEL MAKING"
        },
        "25952": {
            "name": "LD#2- 4",
            "division": "STEEL MAKING"
        },
        "22020": {
            "name": "LD#2- 5",
            "division": "STEEL MAKING"
        },
        "22022": {
            "name": "LD#2- 6",
            "division": "STEEL MAKING"
        },
        "25541": {
            "name": "LD#2, IEM",
            "division": "STEEL MAKING"
        },
        "25543": {
            "name": "LD#2, FME",
            "division": "STEEL MAKING"
        },
        "22021": {
            "name": "LD#2, BIM",
            "division": "STEEL MAKING"
        },
        "25966": {
            "name": "LD#3 (Mechanical) -1",
            "division": "STEEL MAKING"
        },
        "25967": {
            "name": "LD#3 (Mechanical) -2",
            "division": "STEEL MAKING"
        },
        "22412": {
            "name": "LD#3 (Operation)",
            "division": "STEEL MAKING"
        },
        "25231": {
            "name": "RMM Mech-1",
            "division": "STEEL MAKING"
        },
        "20161": {
            "name": "RMM Mech-2",
            "division": "STEEL MAKING"
        },
        "20660": {
            "name": "RMM Mech-3",
            "division": "STEEL MAKING"
        },
        "20004": {
            "name": "RMM OPN",
            "division": "STEEL MAKING"
        },
        "20061": {
            "name": "RMM Coke",
            "division": "STEEL MAKING"
        },
        "25935": {
            "name": "Lime Plant",
            "division": "STEEL MAKING"
        },
        "25695": {
            "name": "Belt Group",
            "division": "STEEL MAKING"
        },
        "25355": {
            "name": "SPP",
            "division": "STEEL MAKING"
        },
        "25335": {
            "name": "SGDP",
            "division": "STEEL MAKING"
        },
        "25620": {
            "name": "HSM",
            "division": "MILLS"
        },
        "23332": {
            "name": "HSM, FME-1",
            "division": "MILLS"
        },
        "25641": {
            "name": "HSM, FME-2",
            "division": "MILLS"
        },
        "25728": {
            "name": "CRM-1",
            "division": "MILLS"
        },
        "25721": {
            "name": "CRM-2",
            "division": "MILLS"
        },
        "25720": {
            "name": "CRM-3",
            "division": "MILLS"
        },
        "23223": {
            "name": "CRM-4",
            "division": "MILLS"
        },
        "25746": {
            "name": "CRM-5",
            "division": "MILLS"
        },
        "22211": {
            "name": "CRM ARP",
            "division": "MILLS"
        },
        "22230": {
            "name": "CRM BAF",
            "division": "MILLS"
        },
        "22250": {
            "name": "CRM F&D",
            "division": "MILLS"
        },
        "25969": {
            "name": "TSCR",
            "division": "MILLS"
        },
        "25050": {
            "name": "TSCR-FME",
            "division": "MILLS"
        },
        "25031": {
            "name": "MM",
            "division": "MILLS"
        },
        "25051": {
            "name": "MM, IEM",
            "division": "MILLS"
        },
        "25820": {
            "name": "WRM",
            "division": "MILLS"
        },
        "25920": {
            "name": "NBM-1",
            "division": "MILLS"
        },
        "21330": {
            "name": "NBM-2",
            "division": "MILLS"
        },
        "25431": {
            "name": "PH#3",
            "division": "POWER SYSTEM"
        },
        "23031": {
            "name": "PH#3 -2",
            "division": "POWER SYSTEM"
        },
        "25531": {
            "name": "PH#4",
            "division": "POWER SYSTEM"
        },
        "25631": {
            "name": "PH#5 -1",
            "division": "POWER SYSTEM"
        },
        "25634": {
            "name": "PH#5 -2",
            "division": "POWER SYSTEM"
        },
        "23065": {
            "name": "BH#2",
            "division": "POWER SYSTEM"
        },
        "25731": {
            "name": "BH#3",
            "division": "POWER SYSTEM"
        },
        "23069": {
            "name": "BH#4",
            "division": "POWER SYSTEM"
        },
        "25732": {
            "name": "BPH-1",
            "division": "POWER SYSTEM"
        },
        "23054": {
            "name": "BPH-2",
            "division": "POWER SYSTEM"
        },
        "23079": {
            "name": "BPH-3",
            "division": "POWER SYSTEM"
        },
        "23114": {
            "name": "BPH-4",
            "division": "POWER SYSTEM"
        },
        "23220": {
            "name": "FMM",
            "division": "OTHER"
        },
        "23224": {
            "name": "Belt Group-1",
            "division": "OTHER"
        },
        "20130": {
            "name": "Belt Group-2",
            "division": "OTHER"
        },
        "29450": {
            "name": "GORC",
            "division": "OTHER"
        },
        "23433": {
            "name": "SMD",
            "division": "OTHER"
        },
        "20006": {
            "name": "RMM Logistics",
            "division": "OTHER"
        },
        "26302": {
            "name": "Central Hub",
            "division": "OTHER"
        },
        "29443": {
            "name": "Fire Brigade",
            "division": "OTHER"
        },
        "23130": {
            "name": "Water Manag",
            "division": "OTHER"
        },
        "22306": {
            "name": "CRM BARA",
            "division": "OTHER"
        },
        "41120": {
            "name": "IBMD-1",
            "division": "OTHER"
        },
        "41130": {
            "name": "IBMD-2",
            "division": "OTHER"
        },
        "41140": {
            "name": "IBMD-3",
            "division": "OTHER"
        },
        "23321": {
            "name": "FME-1",
            "division": "OTHER"
        },
        "23322": {
            "name": "FME-2",
            "division": "OTHER"
        },
        "25140": {
            "name": "FME-3",
            "division": "OTHER"
        },
        "25251": {
            "name": "FME-4",
            "division": "OTHER"
        },
        "20316": {
            "name": "FME-5",
            "division": "OTHER"
        },
        "25934": {
            "name": "FMD MECHANICAL",
            "division": "OTHER"
        },
        "23017": {
            "name": "FMD-1",
            "division": "OTHER"
        },
        "25018": {
            "name": "FMD-2",
            "division": "OTHER"
        },
        "23016": {
            "name": "FMD-3",
            "division": "OTHER"
        },
        "23015": {
            "name": "FMD GCP-1",
            "division": "OTHER"
        },
        "25931": {
            "name": "FMD GCP-2",
            "division": "OTHER"
        },
        "23018": {
            "name": "FMD GAS LINE",
            "division": "OTHER"
        },
        "25932": {
            "name": "FMD GMS",
            "division": "OTHER"
        },
        "23012": {
            "name": "FMD GHB",
            "division": "OTHER"
        },
        "25635": {
            "name": "40 MW-1",
            "division": "OTHER"
        },
        "25655": {
            "name": "40 MW-2",
            "division": "OTHER"
        },
        "29162": {
            "name": "E&P",
            "division": "OTHER"
        },
        "29200": {
            "name": "E&P-HSM",
            "division": "OTHER"
        },
        "23086": {
            "name": "T-30",
            "division": "OTHER"
        },
        "23084": {
            "name": "MPDS",
            "division": "OTHER"
        },
        "25351": {
            "name": "MRP, IEM",
            "division": "OTHER"
        },
        "29010": {
            "name": "R&D-1",
            "division": "OTHER"
        },
        "29016": {
            "name": "R&D-2",
            "division": "OTHER"
        },
        "29011": {
            "name": "TWSGS",
            "division": "OTHER"
        },
        "21120": {
            "name": "Engg. Services (LP)",
            "division": "OTHER"
        },
        "23615": {
            "name": "Engg. Services (AC)",
            "division": "OTHER"
        },
        "22207": {
            "name": "Engg. Services (AC-CRM)",
            "division": "OTHER"
        },
        "20260": {
            "name": "Engg. Services (PP)",
            "division": "OTHER"
        },
        "22220": {
            "name": "Engg. Services (BIM-1)",
            "division": "OTHER"
        },
        "23711": {
            "name": "Engg. Services (BIM-2)",
            "division": "OTHER"
        },
        "22142": {
            "name": "Engg. Services (BIM-3)",
            "division": "OTHER"
        },
        "20240": {
            "name": "Engg. Services (BIM-4)",
            "division": "OTHER"
        },
        "22130": {
            "name": "Engg. Services (BIM-5)",
            "division": "OTHER"
        },
        "23151": {
            "name": "Engg. Services (BIM-6)",
            "division": "OTHER"
        },
        "23420": {
            "name": "Engg. Services (BIM SMD-1)",
            "division": "OTHER"
        },
        "25990": {
            "name": "Engg. Services (BIM SMD-2)",
            "division": "OTHER"
        },
        "21021": {
            "name": "Engg. Services (BIM LD#1)",
            "division": "OTHER"
        },
        "20710": {
            "name": "Engg. Services (BIM I-BF)",
            "division": "OTHER"
        },
        "23431": {
            "name": "Segment Shop",
            "division": "OTHER"
        },
        "21003": {
            "name": "EQMS, SME",
            "division": "OTHER"
        },
        "23716": {
            "name": "EQMS-1",
            "division": "OTHER"
        },
        "23712": {
            "name": "EQMS-2",
            "division": "OTHER"
        },
        "23718": {
            "name": "EQMS-3",
            "division": "OTHER"
        },
        "26202": {
            "name": "Delivery Management",
            "division": "OTHER"
        }
    };

    // File Upload Handler
    fileUploadComp.addEventListener('change', (e) => {
        const file = e.target.files[0];
        handleFileSelect(file);
    });

    function handleFileSelect(file) {
        if (!file) return;
        uploadedFile = file;
        fileNameDisplay.textContent = file.name;
        fileInfoDiv.classList.remove('hidden');
        document.querySelector('.upload-text').textContent = "File uploaded";
        generateBtn.disabled = false;
        errorContainer.classList.add('hidden');
        outputSection.classList.add('hidden');
        exportBtn.disabled = true;
    }

    removeFileBtn.addEventListener('click', () => {
        uploadedFile = null;
        fileUploadComp.value = '';
        fileInfoDiv.classList.add('hidden');
        document.querySelector('.upload-text').textContent = "Click to upload or drag and drop Excel/CSV file";
        generateBtn.disabled = true;
        outputSection.classList.add('hidden');
        errorContainer.classList.add('hidden');
        summaryData = []; // Clear summary data
        exportBtn.disabled = true;
    });

    // Generate Summary
    generateBtn.addEventListener('click', async () => {
        if (!uploadedFile) return;

        try {
            const data = await parseExcel(uploadedFile);
            const aggregated = processData(data);
            renderTable(aggregated);
            outputSection.classList.remove('hidden');
        } catch (err) {
            showError("Error processing file: " + err.message);
            console.error(err);
        }
    });

    function showError(msg) {
        errorMessage.textContent = msg;
        errorContainer.classList.remove('hidden');
        outputSection.classList.add('hidden');
    }

    // --- Core Logic Functions ---

    // Helper to find column indices based on header names (case-insensitive, partial match)
    function findHeaderIndices(headers) {
        // Common header names for Cost Center and Amount in finance reports (adapt as needed)
        const ccHeaderOptions = ['cost center', 'cc code', 'co_area']; 
        const amountHeaderOptions = ['amount', 'total amount', 'value', 'amount in co. code currency'];

        let ccIndex = -1;
        let amountIndex = -1;

        for (let i = 0; i < headers.length; i++) {
            const header = String(headers[i]).toLowerCase().trim();
            if (ccIndex === -1 && ccHeaderOptions.some(opt => header.includes(opt))) {
                ccIndex = i;
            }
            if (amountIndex === -1 && amountHeaderOptions.some(opt => header.includes(opt))) {
                amountIndex = i;
            }
            if (ccIndex !== -1 && amountIndex !== -1) break;
        }

        if (ccIndex === -1 || amountIndex === -1) {
             throw new Error("Could not find required columns: 'Cost Center' and 'Amount'. Please ensure your Excel file contains these headers.");
        }

        return { ccIndex, amountIndex };
    }

    function parseExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    // Assume first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    // Read as array of arrays (AOA) to handle complex layouts
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                    resolve(jsonData);
                } catch (err) {
                    reject(new Error("Failed to read Excel data. Make sure it's a valid XLSX/XLS file."));
                }
            };

            reader.onerror = (err) => {
                reject(new Error("File read operation failed."));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    function processData(data) {
        if (!data || data.length < 2) {
            showError("The file appears to be empty or has no data rows.");
            return [];
        }

        const headers = data[0];
        const rows = data.slice(1);

        const { ccIndex, amountIndex } = findHeaderIndices(headers);

        const aggregationMap = {}; // { 'CC_CODE': { total: 1234.56, ... } }

        for (const row of rows) {
            let ccCode = String(row[ccIndex]).trim();
            let amount = parseFloat(row[amountIndex]);

            // Clean up CC Code (e.g., remove non-digits, take first 5 digits)
            ccCode = ccCode.replace(/[^0-9]/g, ''); 
            if (ccCode.length > 5) {
                ccCode = ccCode.substring(0, 5);
            }

            if (!ccCode || isNaN(amount) || amount === 0) {
                continue; // Skip invalid or zero-amount rows
            }
            
            const ccInfo = CC_DATA[ccCode];
            if (!ccInfo) {
                // Skip Cost Centers not in the master CC_DATA list
                continue; 
            }

            if (!aggregationMap[ccCode]) {
                aggregationMap[ccCode] = {
                    cc: ccCode,
                    name: ccInfo.name,
                    division: ccInfo.division,
                    total: 0
                };
            }

            aggregationMap[ccCode].total += amount;
        }

        const finalData = Object.values(aggregationMap);
        
        // Store for export
        summaryData = finalData; 
        
        return finalData;
    }

    function renderTable(aggregatedData) {
        summaryTableBody.innerHTML = ''; // Clear previous data
        let grandTotal = 0;

        // Sort by division and then by CC code for better readability
        aggregatedData.sort((a, b) => {
            if (a.division < b.division) return -1;
            if (a.division > b.division) return 1;
            return a.cc.localeCompare(b.cc);
        });

        aggregatedData.forEach(item => {
            const row = summaryTableBody.insertRow();
            
            // Format number to 2 decimal places with Indian locale for commas
            const formattedTotal = item.total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            
            row.innerHTML = `
                <td>${item.cc}</td>
                <td>${item.division}</td>
                <td>${item.name}</td>
                <td style="text-align: right;">${formattedTotal}</td>
            `;
            
            grandTotal += item.total;
        });
        
        // Update Grand Total in the footer
        grandTotalDisplay.textContent = grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        
        // Enable export button
        exportBtn.disabled = aggregatedData.length === 0;
    }

    exportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) return;
        exportToExcel(summaryData);
    });

    function exportToExcel(data) {
        const exportableData = data.map(item => ({
            'CC': item.cc,
            'Division': item.division,
            'CC Name': item.name,
            'Total Amount (₹)': item.total
        }));

        // Calculate and add the Grand Total row
        const grandTotal = exportableData.reduce((sum, item) => sum + item['Total Amount (₹)'], 0);
        
        exportableData.push({
            'CC': 'Grand Total',
            'Division': '',
            'CC Name': '',
            'Total Amount (₹)': grandTotal
        });

        const worksheet = XLSX.utils.json_to_sheet(exportableData);

        // Styling (optional: to make the last row bold/different)
        const totalRowIndex = exportableData.length; // 1-based index
        if (worksheet[`A${totalRowIndex}`]) {
            worksheet[`A${totalRowIndex}`].s = { font: { bold: true } };
            worksheet[`D${totalRowIndex}`].s = { font: { bold: true } };
        }

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "CostCenterSummary");

        XLSX.writeFile(workbook, "Cost_Center_Summary.xlsx");
    }
});
