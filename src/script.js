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
    }

    removeFileBtn.addEventListener('click', () => {
        uploadedFile = null;
        fileUploadComp.value = '';
        fileInfoDiv.classList.add('hidden');
        document.querySelector('.upload-text').textContent = "Click to upload or drag and drop Excel/CSV file";
        generateBtn.disabled = true;
        outputSection.classList.add('hidden');
        errorContainer.classList.add('hidden');
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
                    reject(err);
                }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    function processData(rows) {
        if (!rows || rows.length === 0) throw new Error("File is empty");

        const totals = {};
        let ccIndex = -1;
        let totalIndex = -1;
        let foundAnyHeader = false;

        rows.forEach((row, rowIndex) => {
            // Check if this row is a header row
            // We look for "CC" and "TOTAL" (case-insensitive)
            // convert row components to string for safe checking
            const rowStrings = row.map(cell => String(cell).toUpperCase().trim());

            const currentCcIndex = rowStrings.indexOf("CC");
            const currentTotalIndex = rowStrings.indexOf("TOTAL");

            if (currentCcIndex !== -1 && currentTotalIndex !== -1) {
                // Found a header row, update indices for subsequent rows
                ccIndex = currentCcIndex;
                totalIndex = currentTotalIndex;
                foundAnyHeader = true;
                return; // processing next row
            }

            // If we have established indices, try to read data
            if (ccIndex !== -1 && totalIndex !== -1) {
                const ccVal = row[ccIndex];
                const totalVal = row[totalIndex];

                // Validate CC: must be present and not "CC" (in case of duplicate headers processed incorrectly, though unlikely with logic above)
                // Also ignore if it looks like a "Total" label in the CC column (though usually that's empty)
                if (ccVal === undefined || ccVal === null || String(ccVal).trim() === "") return;

                // If CC is literally "CC", it's a header, skip (should be caught by header detection but safety check)
                if (String(ccVal).toUpperCase().trim() === "CC") return;

                // Validate Amount
                let amount = 0;
                if (typeof totalVal === 'number') {
                    amount = totalVal;
                } else if (typeof totalVal === 'string') {
                    // Remove commas, currency symbols
                    const clean = totalVal.replace(/[,]/g, '');
                    amount = parseFloat(clean);
                }

                if (isNaN(amount)) amount = 0;

                // If CC is numeric-ish, treat as valid. 
                // Some "Total" rows might have text in some columns, ensure we don't pick them up.
                // Usually "Block Total" rows have empty CC in the provided screenshot.
                // If the user puts "Total" in the CC column, we should skip it.
                if (String(ccVal).toUpperCase().includes("TOTAL")) return;

                const key = String(ccVal).trim();

                if (!totals[key]) totals[key] = 0;
                totals[key] += amount;
            }
        });

        if (!foundAnyHeader) {
            throw new Error("FILE FORMAT ERROR: Could not find a header row containing both 'CC' and 'TOTAL', call +91-8050509173 for help.");
        }

        // Convert back to array
        const result = Object.keys(totals)
            .map(key => {
                const info = CC_DATA[key] || {};
                return {
                    cc: key,
                    division: info.division || "Unknown",
                    name: info.name || "Unknown",
                    total: totals[key]
                };
            })
            .filter(item => item.total !== 0 && item.total !== null); // Filter 0 or null

        // Sort by CC (numeric if possible)
        result.sort((a, b) => {
            const numA = parseFloat(a.cc);
            const numB = parseFloat(b.cc);
            if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
            return a.cc.localeCompare(b.cc);
        });

        summaryData = result; // save for export
        return result;
    }

    function renderTable(data) {
        summaryTableBody.innerHTML = '';
        let grandTotal = 0;

        data.forEach(item => {
            const tr = document.createElement('tr');
            grandTotal += item.total;
            tr.innerHTML = `
                <td>${item.cc}</td>
                <td>${item.division}</td>
                <td>${item.name}</td>
                <td>${item.total.toFixed(2)}</td>
            `;
            summaryTableBody.appendChild(tr);
        });

        grandTotalDisplay.textContent = grandTotal.toFixed(2);
        // Update footer colspan
        document.querySelector('.total-row td:first-child').setAttribute('colspan', '3');
    }

    // Export Functionality
    exportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) return;

        const ws = XLSX.utils.json_to_sheet(summaryData.map(item => ({
            "Cost Center": item.cc,
            "Division": item.division,
            "Description": item.name,
            "Total Amount": item.total
        })));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Summary");
        XLSX.writeFile(wb, "Cost_Center_Summary.xlsx");
    });
});
