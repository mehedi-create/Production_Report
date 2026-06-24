import ExcelJS from 'exceljs';

export async function generateExcelReport(results) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Summary');

    // ডিজাইনের স্টাইল সেটআপ
    const centerAlign = { vertical: 'middle', horizontal: 'center' };
    const borderStyle = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
    };
    const boldFont = { name: 'Arial', size: 10, bold: true };
    const titleFont = { name: 'Arial', size: 14, bold: true };

    // Header Row 1
    sheet.mergeCells('A1:D1');
    sheet.getCell('A1').value = 'Zakaria Knitwear Ltd (Printing)';
    sheet.getCell('A1').font = titleFont;
    sheet.getCell('A1').alignment = centerAlign;
    sheet.getCell('E1').value = 'Factory: Porabari, Ghatail, Tangail.';
    sheet.getCell('E1').font = { name: 'Arial', size: 12, bold: true };
    sheet.getCell('E1').alignment = { vertical: 'middle', horizontal: 'left' };

    // Header Row 2
    let reportMonth = new Date().toLocaleString('en-US', { month: 'short', year: 'numeric' });
    if (results.length > 0) reportMonth = new Date(results[0].date).toLocaleString('en-US', { month: 'short', year: 'numeric' });
    
    sheet.mergeCells('A2:AP2');
    sheet.getCell('A2').value = `MONTHLY  Printing PRODUCTION  SUMMARY, ${reportMonth}`;
    sheet.getCell('A2').font = titleFont;
    sheet.getCell('A2').alignment = centerAlign;

    // Table Header Row 4 & 5
    sheet.mergeCells('A4:A5'); sheet.getCell('A4').value = 'SL.';
    sheet.mergeCells('B4:B5'); sheet.getCell('B4').value = 'Buyer';
    sheet.mergeCells('C4:C5'); sheet.getCell('C4').value = 'Style No';
    sheet.mergeCells('D4:D5'); sheet.getCell('D4').value = 'Print Type';
    
    // ১ থেকে ৩১ তারিখ পর্যন্ত ডেট কলাম তৈরি
    for (let i = 1; i <= 31; i++) {
        const col = i + 4; // E কলাম থেকে শুরু (Index 5)
        sheet.getCell(4, col).value = i;
    }

    sheet.mergeCells('AJ4:AJ5'); sheet.getCell('AJ4').value = 'Monthly Total';
    sheet.mergeCells('AK4:AK5'); sheet.getCell('AK4').value = 'Pre.Month Total';
    sheet.mergeCells('AL4:AL5'); sheet.getCell('AL4').value = 'G. Total';
    
    // CM এর হেডার (AN, AO, AP)
    sheet.mergeCells('AN4:AN5'); 
    sheet.getCell('AN4').value = 'CM Dzn\n$';
    sheet.getCell('AN4').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    
    sheet.mergeCells('AO4:AO5'); 
    sheet.getCell('AO4').value = 'CM Pcs \n$';
    sheet.getCell('AO4').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    
    sheet.mergeCells('AP4:AP5'); 
    sheet.getCell('AP4').value = 'Total CM $';

    // হেডারে বর্ডার এবং স্টাইল অ্যাপ্লাই করা
    for(let col = 1; col <= 42; col++) {
        if(col === 39) continue; // AM কলাম ফাঁকা
        [4, 5].forEach(row => {
            const cell = sheet.getCell(row, col);
            cell.font = boldFont;
            cell.alignment = centerAlign;
            cell.border = borderStyle;
        });
    }

    // ডেটাবেস থেকে ডেটা গ্রুপ করা (বায়ার, স্টাইল, প্রিন্ট টাইপ এবং CM অনুযায়ী)
    const groups = {};
    results.forEach(row => {
        const dateObj = new Date(row.date);
        const day = dateObj.getDate();
        
        const key = `${row.buyer}|${row.style}|${row.print_type}|${row.cm_dzn}`;
        if (!groups[key]) {
            groups[key] = {
                buyer: row.buyer, style: row.style, print_type: row.print_type, cm_dzn: row.cm_dzn, days: Array(31).fill(0)
            };
        }
        groups[key].days[day - 1] += row.quantity;
    });

    let currentRow = 6;
    let sl = 1;
    const startRow = currentRow;

    // গ্রুপ করা ডেটা এক্সেলে বসানো
    for (const key in groups) {
        const data = groups[key];
        
        sheet.getCell(currentRow, 1).value = sl++;
        sheet.getCell(currentRow, 2).value = data.buyer;
        sheet.getCell(currentRow, 3).value = data.style;
        sheet.getCell(currentRow, 4).value = data.print_type;
        
        // তারিখ অনুযায়ী কোয়ান্টিটি বসানো
        for (let i = 0; i < 31; i++) {
            const qty = data.days[i];
            sheet.getCell(currentRow, 5 + i).value = qty > 0 ? qty : null;
        }

        // এক্সেল ফর্মুলা বসানো (Monthly Total, CM Pcs, Total CM)
        sheet.getCell(currentRow, 36).value = { formula: `SUM(E${currentRow}:AI${currentRow})` }; // AJ
        sheet.getCell(currentRow, 37).value = null; // AK (Pre.Month)
        sheet.getCell(currentRow, 38).value = { formula: `AJ${currentRow}+AK${currentRow}` }; // AL (G. Total)
        
        sheet.getCell(currentRow, 40).value = data.cm_dzn; // AN
        sheet.getCell(currentRow, 41).value = { formula: `AN${currentRow}/12` }; // AO
        sheet.getCell(currentRow, 42).value = { formula: `AJ${currentRow}*AO${currentRow}` }; // AP

        for(let col = 1; col <= 42; col++) {
            if(col === 39) continue;
            const cell = sheet.getCell(currentRow, col);
            cell.border = borderStyle;
            cell.alignment = centerAlign;
        }
        currentRow++;
    }

    const endRow = currentRow - 1;

    // Sub Total Row তৈরি
    sheet.mergeCells(`A${currentRow}:D${currentRow}`);
    const subTotalCell = sheet.getCell(`A${currentRow}`);
    subTotalCell.value = 'Sub Total';
    subTotalCell.font = boldFont;
    subTotalCell.alignment = { vertical: 'middle', horizontal: 'right' };
    
    if (endRow >= startRow) {
        for (let i = 5; i <= 38; i++) {
            const colLetter = sheet.getColumn(i).letter;
            sheet.getCell(currentRow, i).value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
        }
        sheet.getCell(currentRow, 42).value = { formula: `SUM(AP${startRow}:AP${endRow})` };
    }

    // Daily CM Row তৈরি
    const dailyCmRow = currentRow + 1;
    sheet.mergeCells(`A${dailyCmRow}:D${dailyCmRow}`);
    const dailyCmCell = sheet.getCell(`A${dailyCmRow}`);
    dailyCmCell.value = 'Daily CM';
    dailyCmCell.font = boldFont;
    dailyCmCell.alignment = { vertical: 'middle', horizontal: 'right' };

    if (endRow >= startRow) {
        for (let i = 5; i <= 35; i++) {
            const colLetter = sheet.getColumn(i).letter;
            // প্রতিদিনের Daily CM ক্যালকুলেশনের স্মার্ট ফর্মুলা
            sheet.getCell(dailyCmRow, i).value = { formula: `SUMPRODUCT(${colLetter}${startRow}:${colLetter}${endRow}, $AO$${startRow}:$AO$${endRow})` };
        }
        sheet.getCell(dailyCmRow, 42).value = { formula: `SUM(E${dailyCmRow}:AI${dailyCmRow})` };
    }

    // টোটালের ঘরে বর্ডার এবং বোল্ড স্টাইল
    [currentRow, dailyCmRow].forEach(r => {
        for(let col = 1; col <= 42; col++) {
            if(col === 39) continue;
            const cell = sheet.getCell(r, col);
            cell.border = borderStyle;
            cell.font = boldFont;
            if(col > 4) cell.alignment = centerAlign;
        }
    });

    // এক্সেল কলামের সাইজ আপনার ফাইলের মতো ঠিক করা
    sheet.getColumn(1).width = 5;
    sheet.getColumn(2).width = 15;
    sheet.getColumn(3).width = 15;
    sheet.getColumn(4).width = 15;
    for (let i = 5; i <= 35; i++) sheet.getColumn(i).width = 6;
    sheet.getColumn(36).width = 12;
    sheet.getColumn(37).width = 12;
    sheet.getColumn(38).width = 12;
    sheet.getColumn(39).width = 2; // ফাঁকা কলাম
    sheet.getColumn(40).width = 10;
    sheet.getColumn(41).width = 10;
    sheet.getColumn(42).width = 15;

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}
