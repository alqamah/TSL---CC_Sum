document.addEventListener('DOMContentLoaded', () => {
    const fileUploadComp = document.getElementById('fileUpload');
    const statusText = document.getElementById('statusText');
    const generateBtn = document.getElementById('generateBtn');
    const outputSection = document.getElementById('outputSection');
    const summaryTableBody = document.querySelector('#summaryTable tbody');
    const totalShiftChargesDisplay = document.getElementById('totalShiftCharges');
    const totalOtChargesDisplay = document.getElementById('totalOtCharges');
    const totalBdChargesDisplay = document.getElementById('totalBdCharges');
    const grandTotalDisplay = document.getElementById('grandTotal');
    const errorContainer = document.getElementById('errorContainer');
    const errorMessage = document.getElementById('errorMessage');
    const exportBtn = document.getElementById('exportBtn');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const reportSection = document.getElementById('reportSection');
    const reportTableBody = document.querySelector('#reportTable tbody');
    const reportGrandTotalDisplay = document.getElementById('reportGrandTotal');
    const exportReportBtn = document.getElementById('exportReportBtn');

    let uploadedFile = null;
    let summaryData = [];
    let reportData = [];

    fileUploadComp.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        uploadedFile = file;
        statusText.textContent = `Selected: ${file.name}`;
        generateBtn.disabled = false;
        errorContainer.classList.add('hidden');
        outputSection.classList.add('hidden');
        reportSection.classList.add('hidden');
    });

    generateBtn.addEventListener('click', async () => {
        if (!uploadedFile) return;
        try {
            const rawRecords = await parseExcel(uploadedFile);
            renderSummary(rawRecords);
            outputSection.classList.remove('hidden');
        } catch (err) {
            showError("Error processing file: " + err.message);
        }
    });

    rateBtn.addEventListener('click', () => {
        renderRateCard();
        const modal = document.getElementById('rateCardModal');
        modal.classList.remove('hidden');
    });

    ccBtn.addEventListener('click', () => {
        renderCCDetails();
        const modal = document.getElementById('ccDetailsModal');
        modal.classList.remove('hidden');
    });

    // Modal close handlers for all modals
    const closeModalBtns = document.querySelectorAll('.close-modal');
    closeModalBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const modalId = btn.dataset.modal || btn.closest('.modal').id;
            document.getElementById(modalId).classList.add('hidden');
        });
    });

    window.addEventListener('click', (event) => {
        if (event.target.classList.contains('modal')) {
            event.target.classList.add('hidden');
        }
    });

    window.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            document.querySelectorAll('.modal:not(.hidden)').forEach(modal => {
                modal.classList.add('hidden');
            });
        }
    });

    function renderRateCard() {
        const tableBody = document.querySelector('#rateCardTable tbody');
        tableBody.innerHTML = '';

        RATE_CARD.forEach(rate => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${rate.crane_name}</td>
                <td>${formatCurrency(rate.hourly_charges)}</td>
                <td>${formatCurrency(rate.ot_amount_per_hr)}</td>
                <td>${formatCurrency(rate.penalty_per_hr)}</td>
            `;
            tableBody.appendChild(row);
        });
    }

    function renderCCDetails() {
        const tableBody = document.querySelector('#ccDetailsTable tbody');
        tableBody.innerHTML = '';

        // Sort by division, then by CC number
        const sortedEntries = Object.entries(CC_DATA).sort((a, b) => {
            if (a[1].division !== b[1].division) {
                return a[1].division.localeCompare(b[1].division);
            }
            return a[0].localeCompare(b[0]);
        });

        sortedEntries.forEach(([ccNumber, ccInfo]) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${ccNumber}</td>
                <td>${ccInfo.name}</td>
                <td>${ccInfo.division}</td>
            `;
            tableBody.appendChild(row);
        });
    }

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
                    const allRecords = [];

                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                        const sheetMeta = extractSheetMetadata(jsonData);
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

    //TODO: Improve Machine Name Extraction Logic to Generalise for alternate naming conventions
    function extractSheetMetadata(rows) {
        let machineName = "Unknown";
        if (rows.length > 2 && rows[2] && rows[2][0]) {
            const cellA3 = String(rows[2][0]).trim();
            const machineMatch = cellA3.match(/MACHINE\s*NAME[:\-\s]+([A-Za-z0-9\-\(\)]+)/i);
            if (machineMatch) {
                machineName = machineMatch[1].trim();
            }
        }
        return { machineName };
    }

    function extractDataRows(rows, sheetMeta) {
        const records = [];
        let headerRowIndex = -1;
        let colIndices = { shiftQty: -1, otQty: -1, bdHour: -1, costCenter: -1 };

        for (let i = 0; i < rows.length; i++) {
            const rowStrings = rows[i].map(cell => String(cell).toUpperCase().trim());
            const shiftIdx = rowStrings.findIndex(s => s.includes("SHIFT") && s.includes("QTY"));
            const otIdx = rowStrings.findIndex(s => s.includes("OT") && s.includes("QTY"));
            const bdIdx = rowStrings.findIndex(s => s.includes("BREAKDOWN") && s.includes("HOUR"));
            const ccIdx = rowStrings.findIndex(s => s === "COST CENTER" || s === "CC");

            if (shiftIdx !== -1 && otIdx !== -1 && ccIdx !== -1) {
                headerRowIndex = i;
                colIndices = { shiftQty: shiftIdx, otQty: otIdx, bdHour: bdIdx, costCenter: ccIdx };
                break;
            }
        }

        if (headerRowIndex === -1) return records;

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            const cc = String(row[colIndices.costCenter] || "").trim();
            if (!cc || cc.toUpperCase() === "COST CENTER" || cc.toUpperCase() === "CC") continue;

            records.push({
                machine: sheetMeta.machineName,
                cc: cc,
                shiftQty: parseFloat(row[colIndices.shiftQty]) || 0,
                otQty: parseFloat(row[colIndices.otQty]) || 0,
                bdHour: colIndices.bdHour !== -1 ? (parseFloat(row[colIndices.bdHour]) || 0) : 0
            });
        }

        return records;
    }

    function findRate(machineName) {
        const defaultRate = { hourly_charges: 0, ot_amount_per_hr: 0, penalty_per_hr: 0 };
        const normMachine = machineName.toLowerCase().replace(/\([^)]*\)/g, '').replace(/[\s]/g, '').trim();

        for (const rateEntry of RATE_CARD) {
            const entryCrane = rateEntry.crane_name.toLowerCase().replace(/[\s]/g, '').trim();
            if (normMachine.includes(entryCrane) || entryCrane.includes(normMachine)) {
                return rateEntry;
            }
        }
        return defaultRate;
    }

    function formatCurrency(value) {
        return `₹ ${value.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }

    function renderSummary(records) {
        summaryTableBody.innerHTML = '';
        let grandShift = 0, grandOt = 0, grandBd = 0;
        let grandShiftCharges = 0, grandOtCharges = 0, grandBdCharges = 0, grandTotal = 0;

        records.forEach(item => {
            const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
            const rate = findRate(item.machine);
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
                <td title="₹${rate.hourly_charges} x ${item.shiftQty.toFixed(2)}h">${formatCurrency(shiftCharges)}</td>
                <td title="₹${rate.ot_amount_per_hr} x ${item.otQty.toFixed(2)}h">${formatCurrency(otCharges)}</td>
                <td title="₹${rate.penalty_per_hr} x ${item.bdHour.toFixed(2)}h">${formatCurrency(bdCharges)}</td>
                <td title="shift_charges + ot_charges - bd_charges">${formatCurrency(netTotal)}</td>
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

        totalShiftChargesDisplay.textContent = formatCurrency(grandShiftCharges);
        totalOtChargesDisplay.textContent = formatCurrency(grandOtCharges);
        totalBdChargesDisplay.textContent = formatCurrency(grandBdCharges);
        grandTotalDisplay.textContent = formatCurrency(grandTotal);

        summaryData = records;
    }

    exportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) return;

        const exportData = summaryData.map(item => {
            const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
            const rate = findRate(item.machine);
            const shiftCharges = item.shiftQty * rate.hourly_charges;
            const otCharges = item.otQty * rate.ot_amount_per_hr;
            const bdCharges = item.bdHour * rate.penalty_per_hr;
            const netTotal = shiftCharges + otCharges - bdCharges;

            return {
                "Machine": item.machine,
                "CC": item.cc,
                "CC Name": ccInfo.name,
                "Division": ccInfo.division,
                "Shift Qty": item.shiftQty,
                "Hourly Rate": rate.hourly_charges,
                "Shift Charges": shiftCharges,
                "OT Qty": item.otQty,
                "OT Rate/hr": rate.ot_amount_per_hr,
                "OT Charges": otCharges,
                "BD Hour": item.bdHour,
                "Penalty/hr": rate.penalty_per_hr,
                "BD Charges": bdCharges,
                "Net Total": netTotal
            };
        });

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Crane Billing Summary");
        XLSX.writeFile(wb, "crane_billing_summary.xlsx");
    });

    generateReportBtn.addEventListener('click', () => {
        if (summaryData.length === 0) {
            showError("No data available. Please generate summary first.");
            return;
        }

        const ccAggregation = {};

        summaryData.forEach(item => {
            const rate = findRate(item.machine);
            const netTotal = (item.shiftQty * rate.hourly_charges) + (item.otQty * rate.ot_amount_per_hr) - (item.bdHour * rate.penalty_per_hr);

            if (!ccAggregation[item.cc]) {
                const ccInfo = CC_DATA[item.cc] || { name: "Unknown", division: "Unknown" };
                ccAggregation[item.cc] = { cc: item.cc, ccName: ccInfo.name, division: ccInfo.division, totalAmount: 0 };
            }
            ccAggregation[item.cc].totalAmount += netTotal;
        });

        reportData = Object.values(ccAggregation).sort((a, b) => {
            if (a.division !== b.division) return a.division.localeCompare(b.division);
            return a.cc.localeCompare(b.cc);
        });

        renderReport(reportData);
        reportSection.classList.remove('hidden');
    });

    function renderReport(data) {
        reportTableBody.innerHTML = '';
        let grandTotal = 0;

        data.forEach(item => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${item.division}</td>
                <td>${item.cc}</td>
                <td>${item.ccName}</td>
                <td>${formatCurrency(item.totalAmount)}</td>
            `;
            reportTableBody.appendChild(row);
            grandTotal += item.totalAmount;
        });

        reportGrandTotalDisplay.textContent = formatCurrency(grandTotal);
    }

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
