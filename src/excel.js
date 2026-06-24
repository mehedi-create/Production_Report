import ExcelJS from 'exceljs';

export async function generateExcelReport(dbRows) {
    const workbook = new ExcelJS.Workbook();
    
    // --- SHEET 1: MONTHLY SUMMARY ENGINE ---
    const sheet1 = workbook.addWorksheet('Production Summary', { views: [{ showGridLines: true }] });
    
    // Title Blocks Styling
    sheet1.mergeCells('A1:AK1');
    const title = sheet1.getCell('A1');
    title.value = 'Zakaria Printing Unit-1';
    title.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1E3A8A' } };
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet1.getRow(1).height = 30;

    sheet1.mergeCells('A2:AK2');
    const subtitle = sheet1.getCell('A2');
    subtitle.value = 'MONTHLY Printing PRODUCTION SUMMARY, June-2026';
    subtitle.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF475569' } };
    subtitle.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet1.getRow(2).height = 20;

    // Build Master Matrix Table Headings
    const headers = ['SL', 'Buyer', 'Style No'];
    for (let d = 1; d <= 31; d++) headers.push(d.toString());
    headers.push('Monthly Total Pcs', 'CM Dzn', 'Total CM');
    
    const headerRow = sheet1.getRow(4);
    headerRow.values = headers;
    headerRow.height = 26;
    
    headerRow.eachCell((cell) => {
        cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    });

    // Set Professional Column Widths
    sheet1.getColumn(1).width = 5;   // SL
    sheet1.getColumn(2).width = 16;  // Buyer
    sheet1.getColumn(3).width = 18;  // Style No
    for (let i = 4; i <= 34; i++) sheet1.getColumn(i).width = 5.5; // Days Matrix 1-31
    sheet1.getColumn(35).width = 16; // Monthly Total Pcs
    sheet1.getColumn(36).width = 10; // CM Dzn
    sheet1.getColumn(37).width = 14; // Total CM

    // Grouping Data Logic
    const sections = ['Stone Attached', 'Neck Print', 'Table Print'];
    let globalCursor = 5;
    
    const grandTotals = Array(31).fill(0);
    let absoluteGrandPcs = 0;
    let absoluteGrandRevenue = 0;

    sections.forEach((sectionName) => {
        // Section Block Header Row Injector
        sheet1.mergeCells(`A${globalCursor}:AK${globalCursor}`);
        const secCell = sheet1.getCell(`A${globalCursor}`);
        secCell.value = sectionName.toUpperCase();
        secCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF2C3E50' } };
        secCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAECEE' } };
        secCell.alignment = { horizontal: 'left', vertical: 'middle' };
        sheet1.getRow(globalCursor).height = 22;
        globalCursor++;

        // Filter out items matching target structural category
        const sectionRows = dbRows.filter(r => r.print_type === sectionName);
        
        // Aggregate matrix matching distinct rows mapping (Buyer + Style)
        const matrixMap = {};
        sectionRows.forEach(item => {
            const key = `${item.buyer}|||${item.style}`;
            if (!matrixMap[key]) {
                matrixMap[key] = {
                    buyer: item.buyer,
                    style: item.style,
                    cm_dzn: item.cm_dzn,
                    days: Array(32).fill(0),
                    rowTotalPcs: 0
                };
            }
            const dayNum = new Date(item.date).getDate();
            if (dayNum >= 1 && dayNum <= 31) {
                matrixMap[key].days[dayNum] += item.quantity;
                matrixMap[key].rowTotalPcs += item.quantity;
            }
        });

        let slCounter = 1;
        const sectionDailyTotals = Array(32).fill(0);
        let sectionTotalPcs = 0;
        let sectionTotalRevenue = 0;

        // Loop over the unique rows inside the Section Group
        Object.values(matrixMap).forEach((rowObj) => {
            const row = sheet1.getRow(globalCursor);
            row.height = 20;

            row.getCell(1).value = slCounter++;
            row.getCell(2).value = rowObj.buyer;
            row.getCell(3).value = rowObj.style;

            // Mapping quantities to specific day columns D-AH
            for (let d = 1; d <= 31; d++) {
                const qty = rowObj.days[d];
                if (qty > 0) {
                    row.getCell(3 + d).value = qty;
                    sectionDailyTotals[d] += qty;
                    grandTotals[d - 1] += qty;
                }
            }

            // Calculation Logic inside row
            const totalRevenue = (rowObj.cm_dzn / 12) * rowObj.rowTotalPcs;
            row.getCell(35).value = rowObj.rowTotalPcs;
            row.getCell(36).value = rowObj.cm_dzn;
            row.getCell(37).value = totalRevenue;

            // Formats and Borders
            row.getCell(36).numFmt = '$#,##0.00';
            row.getCell(37).numFmt = '$#,##0.00';
            row.eachCell({ includeEmpty: true }, (c, colNum) => {
                if (colNum <= 37) {
                    c.font = { name: 'Arial', size: 9 };
                    c.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                    c.alignment = { horizontal: colNum > 3 ? 'center' : 'left', vertical: 'middle' };
                }
            });

            sectionTotalPcs += rowObj.rowTotalPcs;
            sectionTotalRevenue += totalRevenue;
            globalCursor++;
        });

        // Sub Total Injection Block
        const subtotalRow = sheet1.getRow(globalCursor);
        subtotalRow.height = 22;
        sheet1.mergeCells(`A${globalCursor}:C${globalCursor}`);
        subtotalRow.getCell(1).value = `Sub Total ${sectionName}`;
        subtotalRow.getCell(1).font = { name: 'Arial', size: 9, bold: true };
        subtotalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' };

        for (let d = 1; d <= 31; d++) {
            if (sectionDailyTotals[d] > 0) {
                subtotalRow.getCell(3 + d).value = sectionDailyTotals[d];
            }
        }
        subtotalRow.getCell(35).value = sectionTotalPcs;
        subtotalRow.getCell(37).value = sectionTotalRevenue;
        subtotalRow.getCell(37).numFmt = '$#,##0.00';

        subtotalRow.eachCell({ includeEmpty: true }, (c, colNum) => {
            if (colNum <= 37) {
                c.font = { name: 'Arial', size: 9, bold: true };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F4F4' } };
                c.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                if (colNum > 3) c.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });

        absoluteGrandPcs += sectionTotalPcs;
        absoluteGrandRevenue += sectionTotalRevenue;
        globalCursor += 2; // Spacing logic matching template
    });

    // Grand Total Injector Row
    const grandRow = sheet1.getRow(globalCursor);
    grandRow.height = 25;
    sheet1.mergeCells(`A${globalCursor}:C${globalCursor}`);
    grandRow.getCell(1).value = 'Grand Total';
    grandRow.getCell(1).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    grandRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' };

    for (let d = 1; d <= 31; d++) {
        if (grandTotals[d - 1] > 0) {
            grandRow.getCell(3 + d).value = grandTotals[d - 1];
        }
    }
    grandRow.getCell(35).value = absoluteGrandPcs;
    grandRow.getCell(37).value = absoluteGrandRevenue;
    grandRow.getCell(37).numFmt = '$#,##0.00';

    grandRow.eachCell({ includeEmpty: true }, (c, colNum) => {
        if (colNum <= 37) {
            c.font = { name: 'Arial', size: 9, bold: true, color: colNum <= 3 ? { argb: 'FFFFFFFF' } : { argb: 'FF000000' } };
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: colNum <= 3 ? { argb: 'FF10B981' } : { argb: 'FFEAFAF1' } };
            c.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
            if (colNum > 3) c.alignment = { horizontal: 'center', vertical: 'middle' };
        }
    });

    // --- SHEET 2: DAILY RUN PRODUCTION MATRIX LOG ---
    const sheet2 = workbook.addWorksheet('Daily Log Matrix Summary');
    sheet2.mergeCells('A1:I1');
    sheet2.getCell('A1').value = 'Zakaria Printing Unit-1 & Unit-2, Ghatail, Tangail.';
    sheet2.getCell('A1').font = { bold: true, size: 12 };
    sheet2.getCell('A1').alignment = { horizontal: 'center' };

    sheet2.mergeCells('A2:I2');
    sheet2.getCell('A2').value = 'Production report with CM Summary For the month of June - 2026';
    sheet2.getCell('A2').font = { italic: true, size: 10 };
    sheet2.getCell('A2').alignment = { horizontal: 'center' };

    const s2Headers = ['SL NO', 'Date / Heat Set', 'Stone Attached', 'Table Print', 'Production PCS', 'Revenue $', 'Employee', 'H.Set Machine Run', 'Remarks'];
    const s2HeaderRow = sheet2.getRow(4);
    s2HeaderRow.values = s2Headers;
    s2HeaderRow.eachCell((c) => {
        c.font = { bold: true, size: 10 };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD5D8DC' } };
        c.alignment = { horizontal: 'center' };
        c.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    });

    // Generate rows day by day (1 to 31)
    let s2Cursor = 5;
    for (let day = 1; day <= 31; day++) {
        const row = sheet2.getRow(s2Cursor);
        const dayStr = `${day}-Jun-26`;
        
        // Filter transactions occurring precisely on this day across sections
        const dayEntries = dbRows.filter(r => new Date(r.date).getDate() === day);
        
        const stoneQty = dayEntries.filter(r => r.print_type === 'Stone Attached').reduce((a, b) => a + b.quantity, 0);
        const neckQty = dayEntries.filter(r => r.print_type === 'Neck Print').reduce((a, b) => a + b.quantity, 0);
        const tableQty = dayEntries.filter(r => r.print_type === 'Table Print').reduce((a, b) => a + b.quantity, 0);
        
        const totalPcs = stoneQty + neckQty + tableQty;
        const totalRev = dayEntries.reduce((sum, item) => sum + ((item.cm_dzn / 12) * item.quantity), 0);

        row.getCell(1).value = day;
        row.getCell(2).value = dayStr;
        row.getCell(3).value = stoneQty > 0 ? stoneQty : '';
        row.getCell(4).value = neckQty > 0 ? neckQty : '';
        row.getCell(5).value = tableQty > 0 ? tableQty : '';
        row.getCell(6).value = totalPcs > 0 ? totalPcs : '';
        row.getCell(7).value = totalRev > 0 ? totalRev : '';
        row.getCell(8).value = totalPcs > 0 ? 15 : ''; // Static matching employee index signature
        row.getCell(9).value = totalPcs > 0 ? '4--2' : '';

        row.getCell(7).numFmt = '$#,##0.00';
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            cell.alignment = { horizontal: 'center' };
        });
        s2Cursor++;
    }

    return await workbook.xlsx.writeBuffer();
}
