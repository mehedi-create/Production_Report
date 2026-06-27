import ExcelJS from 'exceljs';

// ============================================
// DAILY REPORT - 100% ইউজারের ফরম্যাট অনুযায়ী
// ============================================
export async function generateDailyExcelReport(results, date) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Production Report');

    // ডেট ফরম্যাট
    const dateObj = new Date(date);
    const day = dateObj.getDate();
    const month = dateObj.toLocaleDateString('en-US', { month: 'long' });
    const year = dateObj.getFullYear();
    const formattedDate = `${month} ${day}, ${year}`;

    // স্টাইল সেটআপ
    const centerAlign = { vertical: 'middle', horizontal: 'center' };
    const leftAlign = { vertical: 'middle', horizontal: 'left' };
    const rightAlign = { vertical: 'middle', horizontal: 'right' };
    const borderStyle = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
    };
    const boldFont = { name: 'Arial', size: 10, bold: true };
    const titleFont = { name: 'Arial', size: 14, bold: true };

    // Row 1: Title
    sheet.mergeCells('A1:G1');
    sheet.getCell('A1').value = `Production Report Of ${formattedDate}`;
    sheet.getCell('A1').font = titleFont;
    sheet.getCell('A1').alignment = leftAlign;

    // Row 3: Company
    sheet.mergeCells('A3:G3');
    sheet.getCell('A3').value = 'Zakaria Knitwear Ltd (Printing)';
    sheet.getCell('A3').font = boldFont;
    sheet.getCell('A3').alignment = leftAlign;

    // Row 4: Address
    sheet.mergeCells('A4:G4');
    sheet.getCell('A4').value = 'Porabari, Ghatail, Tangail';
    sheet.getCell('A4').font = boldFont;
    sheet.getCell('A4').alignment = leftAlign;

    // Row 7: Headers
    const headers = [
        'Serial No', 'Buyer', 'Style', 'Print Type',
        'Production Qty (Pcs)', 'CM Per Dozen ($)', 'Total CM ($)'
    ];

    headers.forEach((header, index) => {
        const cell = sheet.getCell(7, index + 1);
        cell.value = header;
        cell.font = boldFont;
        cell.alignment = centerAlign;
        cell.border = borderStyle;
    });

    // ডাটা রো যোগ করা
    let currentRow = 8;
    let serialNo = 1;

    results.forEach(row => {
        // Serial No
        sheet.getCell(currentRow, 1).value = serialNo++;
        sheet.getCell(currentRow, 1).alignment = centerAlign;
        sheet.getCell(currentRow, 1).border = borderStyle;

        // Buyer
        sheet.getCell(currentRow, 2).value = row.buyer;
        sheet.getCell(currentRow, 2).alignment = leftAlign;
        sheet.getCell(currentRow, 2).border = borderStyle;

        // Style
        sheet.getCell(currentRow, 3).value = row.style;
        sheet.getCell(currentRow, 3).alignment = leftAlign;
        sheet.getCell(currentRow, 3).border = borderStyle;

        // Print Type
        sheet.getCell(currentRow, 4).value = row.print_type;
        sheet.getCell(currentRow, 4).alignment = leftAlign;
        sheet.getCell(currentRow, 4).border = borderStyle;

        // Production Qty
        sheet.getCell(currentRow, 5).value = row.quantity;
        sheet.getCell(currentRow, 5).alignment = centerAlign;
        sheet.getCell(currentRow, 5).border = borderStyle;

        // CM Per Dozen
        sheet.getCell(currentRow, 6).value = row.cm_dzn;
        sheet.getCell(currentRow, 6).alignment = centerAlign;
        sheet.getCell(currentRow, 6).border = borderStyle;

        // Total CM - ফর্মুলা: (E8/12)*F8
        sheet.getCell(currentRow, 7).value = { formula: `(E${currentRow}/12)*F${currentRow}` };
        sheet.getCell(currentRow, 7).alignment = centerAlign;
        sheet.getCell(currentRow, 7).border = borderStyle;

        currentRow++;
    });

    // টোটাল রো
    const totalRow = currentRow;
    sheet.mergeCells(`A${totalRow}:D${totalRow}`);
    sheet.getCell(`A${totalRow}`).value = 'Total';
    sheet.getCell(`A${totalRow}`).font = boldFont;
    sheet.getCell(`A${totalRow}`).alignment = centerAlign;
    sheet.getCell(`A${totalRow}`).border = borderStyle;

    // Total Production Qty
    if (currentRow > 8) {
        sheet.getCell(totalRow, 5).value = { formula: `SUM(E8:E${currentRow - 1})` };
        sheet.getCell(totalRow, 5).font = boldFont;
        sheet.getCell(totalRow, 5).alignment = centerAlign;
        sheet.getCell(totalRow, 5).border = borderStyle;
    }

    // Total CM
    if (currentRow > 8) {
        sheet.getCell(totalRow, 7).value = { formula: `SUM(G8:G${currentRow - 1})` };
        sheet.getCell(totalRow, 7).font = boldFont;
        sheet.getCell(totalRow, 7).alignment = centerAlign;
        sheet.getCell(totalRow, 7).border = borderStyle;
    }

    // কলাম উইডথ সেট (ইউজারের ফাইলের মতো)
    sheet.getColumn(1).width = 12;  // Serial No
    sheet.getColumn(2).width = 20;  // Buyer
    sheet.getColumn(3).width = 20;  // Style
    sheet.getColumn(4).width = 20;  // Print Type
    sheet.getColumn(5).width = 22;  // Production Qty
    sheet.getColumn(6).width = 20;  // CM Per Dozen
    sheet.getColumn(7).width = 18;  // Total CM

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// ============================================
// MONTHLY REPORT - 100% ইউজারের ফরম্যাট অনুযায়ী
// ============================================
export async function generateMonthlyExcelReport(results, month) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Monthly Production Matrix');

    // মাস এবং সাল পার্স
    const monthObj = new Date(month);
    const year = monthObj.getFullYear();
    const monthName = monthObj.toLocaleDateString('en-US', { month: 'long' });
    const reportTitle = `${monthName} ${year}`;

    // স্টাইল সেটআপ
    const centerAlign = { vertical: 'middle', horizontal: 'center' };
    const leftAlign = { vertical: 'middle', horizontal: 'left' };
    const rightAlign = { vertical: 'middle', horizontal: 'right' };
    const borderStyle = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
    };
    const boldFont = { name: 'Arial', size: 10, bold: true };
    const titleFont = { name: 'Arial', size: 14, bold: true };
    const headerFont = { name: 'Arial', size: 9, bold: true };

    // Row 1: Title
    sheet.mergeCells('A1:AK1');
    sheet.getCell('A1').value = `Production Report Of ${reportTitle}`;
    sheet.getCell('A1').font = titleFont;
    sheet.getCell('A1').alignment = leftAlign;

    // Row 3: Company
    sheet.mergeCells('A3:AK3');
    sheet.getCell('A3').value = 'Zakaria Knitwear Ltd (Printing)';
    sheet.getCell('A3').font = boldFont;
    sheet.getCell('A3').alignment = leftAlign;

    // Row 4: Address
    sheet.mergeCells('A4:AK4');
    sheet.getCell('A4').value = 'Porabari, Ghatail, Tangail';
    sheet.getCell('A4').font = boldFont;
    sheet.getCell('A4').alignment = leftAlign;

    // Row 7: Headers
    const mainHeaders = ['SL', 'Buyer', 'Style', 'Print Type', 'CM/Doz ($)'];
    mainHeaders.forEach((header, index) => {
        const cell = sheet.getCell(7, index + 1);
        cell.value = header;
        cell.font = headerFont;
        cell.alignment = centerAlign;
        cell.border = borderStyle;
    });

    // তারিখ হেডার (1-31)
    for (let i = 1; i <= 31; i++) {
        const cell = sheet.getCell(7, 5 + i);
        cell.value = i;
        cell.font = headerFont;
        cell.alignment = centerAlign;
        cell.border = borderStyle;
    }

    // অতিরিক্ত হেডার
    sheet.getCell(7, 37).value = 'Prev. Qty\n(Pcs)';
    sheet.getCell(7, 37).font = headerFont;
    sheet.getCell(7, 37).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    sheet.getCell(7, 37).border = borderStyle;

    sheet.getCell(7, 38).value = 'Curr. M Qty\n(Pcs)';
    sheet.getCell(7, 38).font = headerFont;
    sheet.getCell(7, 38).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    sheet.getCell(7, 38).border = borderStyle;

    sheet.getCell(7, 39).value = 'Total CM\n($)';
    sheet.getCell(7, 39).font = headerFont;
    sheet.getCell(7, 39).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    sheet.getCell(7, 39).border = borderStyle;

    // ডাটা গ্রুপ করা
    const groups = {};
    results.forEach(row => {
        const dateObj = new Date(row.date);
        const day = dateObj.getDate();

        const key = `${row.buyer}|${row.style}|${row.print_type}|${row.cm_dzn}`;
        if (!groups[key]) {
            groups[key] = {
                buyer: row.buyer,
                style: row.style,
                print_type: row.print_type,
                cm_dzn: row.cm_dzn,
                days: Array(32).fill(0),
                totalQty: 0
            };
        }
        groups[key].days[day] += row.quantity;
        groups[key].totalQty += row.quantity;
    });

    // ডাটা রো যোগ করা
    let currentRow = 8;
    let serialNo = 1;
    const startRow = currentRow;

    for (const key in groups) {
        const data = groups[key];

        // SL
        sheet.getCell(currentRow, 1).value = serialNo++;
        sheet.getCell(currentRow, 1).alignment = centerAlign;
        sheet.getCell(currentRow, 1).border = borderStyle;

        // Buyer
        sheet.getCell(currentRow, 2).value = data.buyer;
        sheet.getCell(currentRow, 2).alignment = leftAlign;
        sheet.getCell(currentRow, 2).border = borderStyle;

        // Style
        sheet.getCell(currentRow, 3).value = data.style;
        sheet.getCell(currentRow, 3).alignment = leftAlign;
        sheet.getCell(currentRow, 3).border = borderStyle;

        // Print Type
        sheet.getCell(currentRow, 4).value = data.print_type;
        sheet.getCell(currentRow, 4).alignment = leftAlign;
        sheet.getCell(currentRow, 4).border = borderStyle;

        // CM/Doz
        sheet.getCell(currentRow, 5).value = data.cm_dzn;
        sheet.getCell(currentRow, 5).alignment = centerAlign;
        sheet.getCell(currentRow, 5).border = borderStyle;

        // তারিখ 1-31
        for (let i = 1; i <= 31; i++) {
            const qty = data.days[i];
            sheet.getCell(currentRow, 5 + i).value = qty > 0 ? qty : null;
            sheet.getCell(currentRow, 5 + i).alignment = centerAlign;
            sheet.getCell(currentRow, 5 + i).border = borderStyle;
        }

        // Prev. Qty (খালি)
        sheet.getCell(currentRow, 37).value = null;
        sheet.getCell(currentRow, 37).border = borderStyle;

        // Curr. M Qty - SUM ফর্মুলা
        sheet.getCell(currentRow, 38).value = { formula: `SUM(F${currentRow}:AF${currentRow})` };
        sheet.getCell(currentRow, 38).alignment = centerAlign;
        sheet.getCell(currentRow, 38).border = borderStyle;

        // Total CM - IF ফর্মুলা (ইউজারের ফাইলের মতো)
        sheet.getCell(currentRow, 39).value = { formula: `IF(AG${currentRow}>0, (AG${currentRow}/12)*E${currentRow}, \"\")` };
        sheet.getCell(currentRow, 39).alignment = centerAlign;
        sheet.getCell(currentRow, 39).border = borderStyle;

        currentRow++;
    }

    const endRow = currentRow - 1;

    // টোটাল প্রোডাকশন কোয়ান্টিটি রো
    if (endRow >= startRow) {
        sheet.mergeCells(`A${currentRow}:E${currentRow}`);
        sheet.getCell(currentRow, 1).value = 'Total Production Qty (Pcs)';
        sheet.getCell(currentRow, 1).font = boldFont;
        sheet.getCell(currentRow, 1).alignment = rightAlign;
        sheet.getCell(currentRow, 1).border = borderStyle;

        // প্রতিটি তারিখ কলামের জন্য SUM
        for (let i = 1; i <= 31; i++) {
            const colLetter = sheet.getColumn(5 + i).letter;
            sheet.getCell(currentRow, 5 + i).value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
            sheet.getCell(currentRow, 5 + i).font = boldFont;
            sheet.getCell(currentRow, 5 + i).alignment = centerAlign;
            sheet.getCell(currentRow, 5 + i).border = borderStyle;
        }

        currentRow++;
    }

    // টোটাল CM ভ্যালু রো
    if (endRow >= startRow) {
        sheet.mergeCells(`A${currentRow}:E${currentRow}`);
        sheet.getCell(currentRow, 1).value = 'Total CM Value ($)';
        sheet.getCell(currentRow, 1).font = boldFont;
        sheet.getCell(currentRow, 1).alignment = rightAlign;
        sheet.getCell(currentRow, 1).border = borderStyle;

        // প্রতিটি তারিখ কলামের জন্য SUMPRODUCT
        for (let i = 1; i <= 31; i++) {
            const colLetter = sheet.getColumn(5 + i).letter;
            sheet.getCell(currentRow, 5 + i).value = {
                formula: `SUMPRODUCT(${colLetter}${startRow}:${colLetter}${endRow}, $E${startRow}:$E${endRow})/12`
            };
            sheet.getCell(currentRow, 5 + i).font = boldFont;
            sheet.getCell(currentRow, 5 + i).alignment = centerAlign;
            sheet.getCell(currentRow, 5 + i).border = borderStyle;
        }

        // টোটাল CM
        sheet.getCell(currentRow, 39).value = { formula: `SUM(AH${startRow}:AH${endRow})` };
        sheet.getCell(currentRow, 39).font = boldFont;
        sheet.getCell(currentRow, 39).alignment = centerAlign;
        sheet.getCell(currentRow, 39).border = borderStyle;
    }

    // কলাম উইডথ সেট (ইউজারের ফাইলের মতো)
    sheet.getColumn(1).width = 5;   // SL
    sheet.getColumn(2).width = 15;  // Buyer
    sheet.getColumn(3).width = 18;  // Style
    sheet.getColumn(4).width = 15;  // Print Type
    sheet.getColumn(5).width = 12;  // CM/Doz

    for (let i = 6; i <= 36; i++) {
        sheet.getColumn(i).width = 6;   // Dates 1-31
    }

    sheet.getColumn(37).width = 12;  // Prev. Qty
    sheet.getColumn(38).width = 12;  // Curr. M Qty
    sheet.getColumn(39).width = 15;  // Total CM

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// ব্যাকওয়ার্ড কম্প্যাটিবিলিটির জন্য
export async function generateExcelReport(results) {
    const now = new Date();
    const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    return generateMonthlyExcelReport(results, month);
}
