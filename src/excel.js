import ExcelJS from 'exceljs';

// ====================== HELPER FUNCTIONS ======================
const thin = { style: 'thin' };

function allThinBorder() {
    return { top: thin, bottom: thin, left: thin, right: thin };
}

function font(bold, size, color, italic = false) {
    return { bold, italic, size, color: { argb: color }, name: 'Calibri' };
}

function fill(argb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function align(horizontal, vertical, wrapText = false) {
    return { horizontal, vertical, wrapText };
}

// ====================== DAILY REPORT ======================
export async function generateDailyExcelReport(results, date) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Production Report');

    // কলাম উইডথ (আপনার এক্সেল থেকে)
    ws.getColumn(1).width = 13.85546875; // A
    ws.getColumn(2).width = 14;          // B
    ws.getColumn(3).width = 14;          // C
    ws.getColumn(4).width = 14;          // D
    ws.getColumn(5).width = 16;          // E
    ws.getColumn(6).width = 16;          // F
    ws.getColumn(7).width = 14;          // G

    // রো হাইট
    ws.getRow(1).height = 21;
    ws.getRow(7).height = 33;
    ws.getRow(8).height = 22.5;
    ws.getRow(9).height = 22.5;
    ws.getRow(10).height = 22.5;
    ws.getRow(11).height = 22.5;
    ws.getRow(12).height = 22.5;
    ws.getRow(13).height = 22.5;

    // ✅ ডেট ফরম্যাট - সিলেক্ট করা ডেট দেখাবে
    const dateObj = new Date(date);
    const day = dateObj.getDate();
    const month = dateObj.toLocaleDateString('en-US', { month: 'long' });
    const year = dateObj.getFullYear();
    const formattedDate = `${month} ${day}, ${year}`;

    // Row 1: Title
    ws.getRow(1).getCell(1).value = `Production Report Of ${formattedDate}`;
    ws.getRow(1).getCell(1).font = font(true, 16, 'FF1F4E78');

    // Row 3: Company
    ws.getRow(3).getCell(1).value = 'Zakaria Knitwear Limited (Printing)';
    ws.getRow(3).getCell(1).font = font(true, 12, 'FF333333');

    // Row 4: Address (ইটালিক)
    ws.getRow(4).getCell(1).value = 'Porabari, Ghatail, Tangail';
    ws.getRow(4).getCell(1).font = font(false, 10, 'FF595959', true);

    // Row 7: Headers
    const headers = [
        'Serial No', 'Buyer', 'Style', 'Print Type',
        'Production Qty (Pcs)', 'CM Per Dozen ($)', 'Total CM ($)'
    ];
    const headerBg = 'FF2F5597';

    const r7 = ws.getRow(7);
    for (let i = 0; i < headers.length; i++) {
        const cell = r7.getCell(i + 1);
        cell.value = headers[i];
        cell.font = font(true, 11, 'FFFFFFFF');
        cell.fill = fill(headerBg);
        cell.alignment = align('center', 'center', true);
        cell.border = allThinBorder();
    }

    // ডাটা রো যোগ করা
    let currentRow = 8;
    let serialNo = 1;

    results.forEach((row, index) => {
        const isEven = (index % 2 === 1);
        const rowFill = isEven ? fill('FFF2F5F9') : null;
        const wsRow = ws.getRow(currentRow);

        // Serial No
        wsRow.getCell(1).value = serialNo++;
        wsRow.getCell(1).font = font(false, 11, 'FF000000');
        wsRow.getCell(1).alignment = align('center', 'center');
        wsRow.getCell(1).border = allThinBorder();
        if (rowFill) wsRow.getCell(1).fill = rowFill;

        // Buyer
        wsRow.getCell(2).value = row.buyer;
        wsRow.getCell(2).font = font(false, 11, 'FF000000');
        wsRow.getCell(2).alignment = align('left', 'center');
        wsRow.getCell(2).border = allThinBorder();
        if (rowFill) wsRow.getCell(2).fill = rowFill;

        // Style
        wsRow.getCell(3).value = row.style;
        wsRow.getCell(3).font = font(false, 11, 'FF000000');
        wsRow.getCell(3).alignment = align('left', 'center');
        wsRow.getCell(3).border = allThinBorder();
        if (rowFill) wsRow.getCell(3).fill = rowFill;

        // Print Type
        wsRow.getCell(4).value = row.print_type;
        wsRow.getCell(4).font = font(false, 11, 'FF000000');
        wsRow.getCell(4).alignment = align('center', 'center');
        wsRow.getCell(4).border = allThinBorder();
        if (rowFill) wsRow.getCell(4).fill = rowFill;

        // Production Qty
        wsRow.getCell(5).value = row.quantity;
        wsRow.getCell(5).font = font(false, 11, 'FF000000');
        wsRow.getCell(5).alignment = align('right', 'center');
        wsRow.getCell(5).border = allThinBorder();
        if (rowFill) wsRow.getCell(5).fill = rowFill;

        // CM Per Dozen
        wsRow.getCell(6).value = row.cm_dzn;
        wsRow.getCell(6).font = font(false, 11, 'FF000000');
        wsRow.getCell(6).alignment = align('right', 'center');
        wsRow.getCell(6).border = allThinBorder();
        if (rowFill) wsRow.getCell(6).fill = rowFill;

        // Total CM - ফর্মুলা
        wsRow.getCell(7).value = { formula: `(E${currentRow}/12)*F${currentRow}` };
        wsRow.getCell(7).font = font(false, 11, 'FF000000');
        wsRow.getCell(7).alignment = align('right', 'center');
        wsRow.getCell(7).border = allThinBorder();
        if (rowFill) wsRow.getCell(7).fill = rowFill;

        currentRow++;
    });

    // টোটাল রো
    if (currentRow > 8) {
        const totalRow = currentRow;
        ws.mergeCells(`A${totalRow}:D${totalRow}`);

        const a13 = ws.getCell(`A${totalRow}`);
        a13.value = 'Total';
        a13.font = font(true, 11, 'FF000000');
        a13.fill = fill('FFE9EEF4');
        a13.alignment = align('left', 'center');
        a13.border = allThinBorder();

        for (let col = 2; col <= 4; col++) {
            const cell = ws.getRow(totalRow).getCell(col);
            cell.fill = fill('FFE9EEF4');
            cell.border = allThinBorder();
        }

        const e13 = ws.getRow(totalRow).getCell(5);
        e13.value = { formula: `SUM(E8:E${totalRow - 1})` };
        e13.font = font(true, 11, 'FF000000');
        e13.fill = fill('FFE9EEF4');
        e13.alignment = align('right', 'center');
        e13.border = allThinBorder();

        const f13 = ws.getRow(totalRow).getCell(6);
        f13.fill = fill('FFE9EEF4');
        f13.border = allThinBorder();

        const g13 = ws.getRow(totalRow).getCell(7);
        g13.value = { formula: `SUM(G8:G${totalRow - 1})` };
        g13.font = font(true, 11, 'FF000000');
        g13.fill = fill('FFE9EEF4');
        g13.alignment = align('right', 'center');
        g13.border = allThinBorder();
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// ====================== MONTHLY REPORT ======================
export async function generateMonthlyExcelReport(results, month) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Monthly Production Matrix');

    // কলাম উইডথ
    const colWidths = {
        1: 3.765625,    // A - SL
        2: 12.43,  // B - Buyer
        3: 8.86, // C - Style
        4: 11.45, // D - Print Type
        5: 4.43,  // E - CM/Doz
        37: 7, // AK - Prev. Qty
        38: 7.57,      // AL - Curr. M Qty
        39: 8.47265625, // AM - Total CM
    };

    for (let c = 1; c <= 39; c++) {
        ws.getColumn(c).width = colWidths[c] || 8;
    }

    // রো হাইট
    const rowHeights = { 1: 21, 6: 35.25 };
    for (let r = 7; r <= 23; r++) rowHeights[r] = 26.25;
    for (const [r, h] of Object.entries(rowHeights)) {
        ws.getRow(Number(r)).height = h;
    }

    // মাস ফরম্যাট
    const monthObj = new Date(month);
    const year = monthObj.getFullYear();
    const monthName = monthObj.toLocaleDateString('en-US', { month: 'long' });
    const reportTitle = `${monthName} ${year}`;

    // Row 1: Title
    ws.getRow(1).getCell(1).value = `Production Report Of ${reportTitle}`;
    ws.getRow(1).getCell(1).font = font(true, 16, 'FF1F4E78');

    // Row 3: Company
    ws.getRow(3).getCell(1).value = 'Zakaria Knitwear Limited (Printing)';
    ws.getRow(3).getCell(1).font = font(true, 12, 'FF333333');

    // Row 4: Address (ইটালিক)
    ws.getRow(4).getCell(1).value = 'Porabari, Ghatail, Tangail';
    ws.getRow(4).getCell(1).font = font(false, 10, 'FF595959', true);

    // Row 5: "Production Dates" merged header
    ws.mergeCells('F5:AJ5');
    const f5 = ws.getCell('F5');
    f5.value = 'Production Dates (1 to 31)';
    f5.font = font(true, 10, 'FFFFFFFF');
    f5.fill = fill('FF2F75B5');
    f5.alignment = align('center', 'center', true);
    f5.border = { top: thin, bottom: thin, left: thin, right: thin };

    for (let c = 7; c <= 36; c++) {
        const cell = ws.getRow(5).getCell(c);
        cell.border = { top: thin, bottom: thin };
    }
    ws.getRow(5).getCell(36).border = { top: thin, bottom: thin, right: thin };

    // Row 6: Column headers
    const r6 = ws.getRow(6);
    const darkBlue = 'FF1F4E78';
    const medBlue = 'FF2F75B5';
    const darkNav = 'FF16365C';

    const hdrs6 = [
        [1, 'SL', darkBlue],
        [2, 'Buyer', darkBlue],
        [3, 'Style', darkBlue],
        [4, 'Print Type', darkBlue],
        [5, 'CM/Doz ($)', darkBlue],
    ];

    for (const [col, val, bg] of hdrs6) {
        const cell = r6.getCell(col);
        cell.value = val;
        cell.font = font(true, 10, 'FFFFFFFF');
        cell.fill = fill(bg);
        cell.alignment = align('center', 'center', true);
        cell.border = allThinBorder();
    }

    // Date columns 1-31
    for (let d = 1; d <= 31; d++) {
        const cell = r6.getCell(5 + d);
        cell.value = d;
        cell.font = font(true, 10, 'FFFFFFFF');
        cell.fill = fill(medBlue);
        cell.alignment = align('center', 'center', true);
        cell.border = allThinBorder();
    }

    // Summary headers
    const summaryHdrs = [
        [37, 'Pre. Qty\n(Pcs)', darkNav],
        [38, 'Total Qty\n(Pcs)', darkNav],
        [39, 'Total CM\n($)', darkNav],
    ];

    for (const [col, val, bg] of summaryHdrs) {
        const cell = r6.getCell(col);
        cell.value = val;
        cell.font = font(true, 10, 'FFFFFFFF');
        cell.fill = fill(bg);
        cell.alignment = align('center', 'center', true);
        cell.border = allThinBorder();
    }

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
    let currentRow = 7;
    let serialNo = 1;
    const startRow = currentRow;

    for (const key in groups) {
        const data = groups[key];
        const isEven = (serialNo % 2 === 0);
        const rowFill = isEven ? fill('FFF2F5F9') : null;
        const wsRow = ws.getRow(currentRow);

        // SL
        wsRow.getCell(1).value = serialNo++;
        wsRow.getCell(1).font = font(false, 10, 'FF000000');
        wsRow.getCell(1).alignment = align('center', 'center');
        wsRow.getCell(1).border = allThinBorder();
        if (rowFill) wsRow.getCell(1).fill = rowFill;

        // Buyer
        wsRow.getCell(2).value = data.buyer;
        wsRow.getCell(2).font = font(false, 10, 'FF000000');
        wsRow.getCell(2).alignment = align('left', 'center');
        wsRow.getCell(2).border = allThinBorder();
        if (rowFill) wsRow.getCell(2).fill = rowFill;

        // Style
        wsRow.getCell(3).value = data.style;
        wsRow.getCell(3).font = font(false, 10, 'FF000000');
        wsRow.getCell(3).alignment = align('left', 'center');
        wsRow.getCell(3).border = allThinBorder();
        if (rowFill) wsRow.getCell(3).fill = rowFill;

        // Print Type
        wsRow.getCell(4).value = data.print_type;
        wsRow.getCell(4).font = font(false, 10, 'FF000000');
        wsRow.getCell(4).alignment = align('center', 'center');
        wsRow.getCell(4).border = allThinBorder();
        if (rowFill) wsRow.getCell(4).fill = rowFill;

        // CM/Doz
        wsRow.getCell(5).value = data.cm_dzn;
        wsRow.getCell(5).font = font(false, 10, 'FF000000');
        wsRow.getCell(5).alignment = align('right', 'center');
        wsRow.getCell(5).border = allThinBorder();
        if (rowFill) wsRow.getCell(5).fill = rowFill;

        // তারিখ কলাম (1-31)
        for (let d = 1; d <= 31; d++) {
            const cell = wsRow.getCell(5 + d);
            const qty = data.days[d];
            cell.value = qty > 0 ? qty : null;
            cell.font = font(false, 10, 'FF000000');
            cell.alignment = align('center', 'center', true);
            cell.border = allThinBorder();
            if (rowFill) cell.fill = rowFill;
        }

        // Prev. Qty
        const cell37 = wsRow.getCell(37);
        cell37.value = null;
        cell37.font = font(false, 10, 'FF000000');
        cell37.alignment = align('right', 'center');
        cell37.border = allThinBorder();
        if (rowFill) cell37.fill = rowFill;

        // Curr. M Qty - SUM ফর্মুলা
        const cell38 = wsRow.getCell(38);
        cell38.value = { formula: `SUM(F${currentRow}:AJ${currentRow})` };
        cell38.font = font(false, 10, 'FF000000');
        cell38.alignment = align('right', 'center');
        cell38.border = allThinBorder();
        if (rowFill) cell38.fill = rowFill;

        // Total CM - IF ফর্মুলা
        const cell39 = wsRow.getCell(39);
        cell39.value = { formula: `IF(AL${currentRow}>0,(AL${currentRow}/12)*E${currentRow},"")` };
        cell39.font = font(false, 10, 'FF000000');
        cell39.alignment = align('right', 'center');
        cell39.border = allThinBorder();
        if (rowFill) cell39.fill = rowFill;

        currentRow++;
    }

    const endRow = currentRow - 1;

    // টোটাল প্রোডাকশন কোয়ান্টিটি রো
    if (endRow >= startRow) {
        ws.mergeCells(`A${currentRow}:E${currentRow}`);
        const r22a = ws.getCell(`A${currentRow}`);
        r22a.value = 'Total Production Qty (Pcs)';
        r22a.font = font(true, 10, 'FF000000');
        r22a.fill = fill('FFE9EEF4');
        r22a.alignment = align('left', 'center');
        r22a.border = allThinBorder();

        for (let col = 6; col <= 39; col++) {
            const cell = ws.getRow(currentRow).getCell(col);
            if (col <= 36) {
                const colLetter = ws.getColumn(col).letter;
                cell.value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
            } else if (col === 37) {
                cell.value = { formula: `SUM(AK${startRow}:AK${endRow})` };
            } else if (col === 38) {
                cell.value = { formula: `SUM(AL${startRow}:AL${endRow})` };
            }
            cell.font = font(true, 10, 'FF000000');
            cell.fill = fill('FFE9EEF4');
            cell.alignment = align('right', 'center');
            cell.border = allThinBorder();
        }
        currentRow++;
    }

    // টোটাল CM ভ্যালু রো
    if (endRow >= startRow) {
        ws.mergeCells(`A${currentRow}:E${currentRow}`);
        const r23a = ws.getCell(`A${currentRow}`);
        r23a.value = 'Total CM Value ($)';
        r23a.font = font(true, 10, 'FF000000');
        r23a.fill = fill('FFD9E1F2');
        r23a.alignment = align('left', 'center');
        r23a.border = allThinBorder();

        for (let col = 6; col <= 39; col++) {
            const cell = ws.getRow(currentRow).getCell(col);
            if (col >= 6 && col <= 36) {
                const colLetter = ws.getColumn(col).letter;
                cell.value = { formula: `SUMPRODUCT(${colLetter}$${startRow}:${colLetter}$${endRow},$E$${startRow}:$E$${endRow})/12` };
            } else if (col === 39) {
                cell.value = { formula: `SUM(AM${startRow}:AM${endRow})` };
            }
            cell.font = font(true, 10, 'FF000000');
            cell.fill = fill('FFD9E1F2');
            cell.alignment = align('right', 'center');
            cell.border = allThinBorder();
        }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// ব্যাকওয়ার্ড কম্প্যাটিবিলিটির জন্য
export async function generateExcelReport(results) {
    const now = new Date();
    const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    return generateMonthlyExcelReport(results, month);
}
