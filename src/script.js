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
    let ccMap = {};

    const ccInfoRaw = `
IRON MAKING:
    A-F BF	25225		
	A-F BF, IEM-1	20300		
	A-F BF, IEM-2	25242		
	C-BF	25223		
	E-BF	23222		
	F-BF	25226		
	G-BF	25320		
	G-BF, IEM	20420		
	G-BF, Belt Group	25230		
	G-BF-SSP	40420		
	H-BF (Mechanical)	25325		
	H-BF (Operation)-1	25345		
	H-BF (Operation)-2	20520		
	I-BF	25981		
	Coke Plant, Battery 5,6,7	25021		
	Coke Plant, BAT 10/11	25032		
	Coke Plant, CDQ 10/11	22203		
	Coke Plant Bat 8,9	25022		
	Coke Plant Old BPP	25023		
	Coke Plant New BPP	25033		
	Coke Plant- 1	20150		
	Coke Plant- 2	20112		
	Coke Plant- 3	20113		
	Coke Plant- 4	20140		
	Coke Plant- 5	25042		
	Coke Plant- 6	25035		
	Coke Plant- 7	25036		
	Coke Plant, IEM	25043		
	Sinter Plant-1 (Mech.)	25121		
	Sinter Plant-1 (Operation-1)	20230		
	Sinter Plant-1 (Operation-2)	23020		
	Sinter Plant-1 (IEM)	25141		
	Sinter Plant-2	25122		
	Sinter Plant-2 (IEM)	25142		
	Sinter Plant-3	25123		
	Sinter Plant-3 (IEM-1)	25143		
	Sinter Plant-3 (IEM-2)	23325		
	Sinter Plant-4	25124		
	Sinter Plant-4 (IEM-1)	20270		
	Sinter Plant-4 (IEM-2)	25144		
	RMBB-1 (Mech)	25120		
	RMBB-1 (Operation)	20220		
	RMBB#2-1	25125		
	RMBB#2-2	25145		
	RMBB#2-3	20250		
	Pellet Plant (Mech)-1	25964		
	Pellet Plant (Mech)-2	25965		
	Pellet Plant Operation-1	20651		
	Pellet Plant Operation-2	20661		
	Pellet Plant-IEM	25974		
	Pellet Plant-FME	25975		
	FMM-IM	25222		
STEEL MAKING:	
    LD#1- 1	25421		
	LD#1- 2	25420		
	LD#1- 3	25422		
	LD#1- 4	23221		
	LD#1-Operation	21030		
	LD#1-IEM	25441		
	LD#1-FME	25443		
	LD#2- 1	25521		
	LD#2- 2	25522		
	LD#2- 3	25221		
	LD#2- 4	25952		
	LD#2- 5	22020		
	LD#2- 6	22022		
	LD#2, IEM	25541		
	LD#2, FME	25543		
	LD#2, BIM	22021		
	LD#3 (Mechanical) -1	25966		
	LD#3 (Mechanical) -2	25967		
	LD#3 (Operation)	22412		
	RMM Mech-1	25231		
	RMM Mech-2	20161		
	RMM Mech-3	20660		
	RMM OPN	20004		
	RMM Coke	20061		
	Lime Plant	25935		
	Belt Group	25695		
	SPP	25355		
	SGDP	25335		
MILLS:	
    HSM	25620		
	HSM, FME-1	23332		
	HSM, FME-2	25641		
	CRM-1	25728		
	CRM-2	25721		
	CRM-3	25720		
	CRM-4	23223		
	CRM-5	25746		
	CRM ARP	22211		
	CRM BAF	22230		
	CRM F&D	22250		
	TSCR	25969		
	TSCR-FME	25050		
	MM	25031		
	MM, IEM	25051		
	WRM	25820		
	NBM-1	25920		
	NBM-2	21330		
POWER SYSTEM:	
    PH#3	25431		
	PH#3 -2	23031		
	PH#4	25531		
	PH#5 -1	25631		
	PH#5 -2	25634		
	BH#2	23065		
	BH#3	25731		
	BH#4	23069		
	BPH-1	25732		
	BPH-2	23054		
	BPH-3	23079		
	BPH-4	23114		
OTHER:	
    FMM	23220		
	Belt Group-1	23224		
	Belt Group-2	20130		
	GORC	29450		
	SMD	23433		
	RMM Logistics	20006		
	Central Hub	26302		
	Fire Brigade	29443		
	Water Manag	23130		
	CRM BARA	22306		
	IBMD-1	41120		
	IBMD-2	41130		
	IBMD-3	41140		
	FME-1	23321		
	FME-2	23322		
	FME-3	25140		
	FME-4	25251		
	FME-5	20316		
	FMD MECHANICAL	25934		
	FMD-1	23017		
	FMD-2	25018		
	FMD-3	23016		
	FMD GCP-1	23015		
	FMD GCP-2	25931		
	FMD GAS LINE	23018		
	FMD GMS	25932		
	FMD GHB	23012		
	40 MW-1	25635		
	40 MW-2	25655		
	E&P	29162		
	E&P-HSM	29200		
	T-30	23086		
	"HML	"	25131		
	MPDS	23084		
	MRP, IEM	25351		
	R&D-1	29010		
	R&D-2	29016		
	TWSGS	29011		
	Engg. Services (LP)	21120		
	Engg. Services (AC)	23615		
	Engg. Services (AC-CRM)	22207		
	Engg. Services (PP)	20260		
	Engg. Services (BIM-1)	22220		
	Engg. Services (BIM-2)	23711		
	Engg. Services (BIM-3)	22142		
	Engg. Services (BIM-4)	20240		
	Engg. Services (BIM-5)	22130		
	Engg. Services (BIM-6)	23151		
	Engg. Services (BIM SMD-1)	23420		
	Engg. Services (BIM SMD-2)	25990		
	Engg. Services (BIM LD#1)	21021		
	Engg. Services (BIM I-BF)	20710		
	Segment Shop	23431		
	EQMS, SME	21003		
	EQMS-1	23716		
	EQMS-2	23712		
	EQMS-3	23718		
	Delivery Management	26202		
`;

    // Parse CC Info
    function parseCCInfo() {
        const lines = ccInfoRaw.split('\n');
        lines.forEach(line => {
            // Split by tab or multiple spaces
            // Look for the last number in the line which is likely the CC
            // Structure is usually: Name <whitespace> CC <whitespace>
            const parts = line.trim().split(/\t+/);
            if (parts.length >= 2) {
                // Try to identify the CC part. It's usually the second column or the last numeric column.
                // In the provided file, it looks like: Name [tab] CC [tab] [empty]
                // So parts[1] is likely the CC.
                const potentialCC = parts[1].trim();
                if (/^\d+$/.test(potentialCC)) {
                    const name = parts[0].trim().replace(/^"|"$/g, '').trim(); // Remove surrounding quotes like "HML "
                    ccMap[potentialCC] = name;
                }
            }
        });
        console.log("Loaded CC Map:", Object.keys(ccMap).length, "entries");
    }

    // Initialize CC Map
    parseCCInfo();

    // File Upload Handler
    fileUploadComp.addEventListener('change', (e) => {
        const file = e.target.files[0];
        handleFileSelect(file);
    });

    // Drag and Drop
    const uploadArea = document.querySelector('.upload-area');
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = 'var(--primary-color)';
        uploadArea.style.backgroundColor = '#eff6ff';
    });
    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = 'var(--border-color)';
        uploadArea.style.backgroundColor = '#f8fafc';
    });
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = 'var(--border-color)';
        uploadArea.style.backgroundColor = '#f8fafc';
        if (e.dataTransfer.files.length) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
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
            throw new Error("Could not find a header row containing both 'CC' and 'TOTAL'. Please check your file format.");
        }

        // Convert back to array
        const result = Object.keys(totals)
            .map(key => ({
                cc: key,
                name: ccMap[key] || "Unknown",
                total: totals[key]
            }))
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
                <td>${item.name}</td>
                <td>${item.total.toFixed(2)}</td>
            `;
            summaryTableBody.appendChild(tr);
        });

        grandTotalDisplay.textContent = grandTotal.toFixed(2);
        // Update footer colspan
        document.querySelector('.total-row td:first-child').setAttribute('colspan', '2');
    }

    // Export Functionality
    exportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) return;

        const ws = XLSX.utils.json_to_sheet(summaryData.map(item => ({
            "Cost Center": item.cc,
            "Description": item.name,
            "Total Amount": item.total
        })));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Summary");
        XLSX.writeFile(wb, "Cost_Center_Summary.xlsx");
    });
});
