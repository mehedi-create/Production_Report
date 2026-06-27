import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';

// কালার কনভার্সন ফাংশন (ARGB থেকে RGB)
function argbToRgb(argb) {
    // Remove alpha channel if present
    const hex = argb.replace(/^FF/, '');
    const r = parseInt(hex.substring(0, 2), 16) / 255;
    const g = parseInt(hex.substring(2, 4), 16) / 255;
    const b = parseInt(hex.substring(4, 6), 16) / 255;
    return rgb(r, g, b);
}

// ============================================
// DAILY PDF REPORT - এক্সেল ফরম্যাটের সাথে মিল রেখে
// ============================================
export async function generateDailyPdfReport(dbRows, date) {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([842, 595]); // A4 Landscape
    const { width, height } = page.getSize();

    const fontHelvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontHelveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // কালার (আপনার এক্সেল থেকে)
    const titleColor = argbToRgb('FF1F4E78');
    const companyColor = argbToRgb('FF333333');
    const addressColor = argbToRgb('FF595959');
    const headerBg = argbToRgb('FF2F5597');
    const evenRowBg = argbToRgb('FFF2F5F9');
    const totalRowBg = argbToRgb('FFE9EEF4');

    // ডেট ফরম্যাট
    const dateObj = new Date(date);
    const day = dateObj.getDate();
    const month = dateObj.toLocaleDateString('en-US', { month: 'long' });
    const year = dateObj.getFullYear();
    const formattedDate = `[Specify Date]`;

    // টাইটেল
    page.drawText(`Production Report Of ${formattedDate}`, {
        x: 50, y: height - 40, size: 16,
        font: fontHelveticaBold, color: titleColor
    });

    // কোম্পানি নাম
    page.drawText('Zakaria Knitwear Limited (Printing)', {
        x: 50, y: height - 60, size: 12,
        font: fontHelveticaBold, color: companyColor
    });

    // ঠিকানা (ইটালিক)
    page.drawText('Porabari, Ghatail, Tangail', {
        x: 50, y: height - 78, size: 10,
        font: fontHelvetica, color: addressColor
    });

    // টেবিল সেটআপ
    const startX = 40;
    let startY = height - 120;
    const colWidths = {
        sl: 40, buyer: 100, style: 100, printType: 100,
        qty: 80, cmDzn: 70, totalCm: 70
    };
    const tableWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);

    // হেডার ব্যাকগ্রাউন্ড
    page.drawRectangle({
        x: startX, y: startY - 20,
        width: tableWidth, height: 20,
        color: headerBg
    });

    // হেডার টেক্সট (সাদা)
    page.drawText('Serial No', { x: startX + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Buyer', { x: startX + colWidths.sl + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Style', { x: startX + colWidths.sl + colWidths.buyer + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Print Type', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Production Qty (Pcs)', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('CM Per Dozen ($)', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.qty + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Total CM ($)', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.qty + colWidths.cmDzn + 5, y: startY - 14, size: 11, font: fontHelveticaBold, color: rgb(1, 1, 1) });

    startY -= 20;

    // ডাটা রো
    let sl = 1;
    let totalQty = 0;
    let totalCm = 0;

    dbRows.forEach(row => {
        const isEven = (sl % 2 === 0);
        if (isEven) {
            page.drawRectangle({
                x: startX, y: startY - 18,
                width: tableWidth, height: 18,
                color: evenRowBg
            });
        }

        page.drawText(sl.toString(), { x: startX + 5, y: startY - 14, size: 11, font: fontHelvetica });
        page.drawText(row.buyer.substring(0, 15), { x: startX + colWidths.sl + 5, y: startY - 14, size: 11, font: fontHelvetica });
        page.drawText(row.style.substring(0, 15), { x: startX + colWidths.sl + colWidths.buyer + 5, y: startY - 14, size: 11, font: fontHelvetica });
        page.drawText(row.print_type.substring(0, 15), { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + 5, y: startY - 14, size: 11, font: fontHelvetica });
        page.drawText(row.quantity.toString(), { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + 10, y: startY - 14, size: 11, font: fontHelvetica });
        page.drawText(`$${row.cm_dzn.toFixed(2)}`, { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.qty + 5, y: startY - 14, size: 11, font: fontHelvetica });

        const rowTotalCm = (row.quantity / 12) * row.cm_dzn;
        page.drawText(`$${rowTotalCm.toFixed(2)}`, { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.qty + colWidths.cmDzn + 5, y: startY - 14, size: 11, font: fontHelvetica });

        page.drawLine({
            start: { x: startX, y: startY - 18 },
            end: { x: width - 40, y: startY - 18 },
            thickness: 0.5, color: rgb(0.85, 0.85, 0.85)
        });

        totalQty += row.quantity;
        totalCm += rowTotalCm;
        startY -= 18;
        sl++;
    });

    // টোটাল রো
    page.drawRectangle({
        x: startX, y: startY - 18,
        width: tableWidth, height: 18,
        color: totalRowBg
    });

    page.drawText('Total', { x: startX + 5, y: startY - 14, size: 11, font: fontHelveticaBold });
    page.drawText(totalQty.toString(), { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + 10, y: startY - 14, size: 11, font: fontHelveticaBold });
    page.drawText(`$${totalCm.toFixed(2)}`, { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.qty + colWidths.cmDzn + 5, y: startY - 14, size: 11, font: fontHelveticaBold });

    page.drawLine({
        start: { x: startX, y: startY - 18 },
        end: { x: width - 40, y: startY - 18 },
        thickness: 0.5, color: rgb(0.85, 0.85, 0.85)
    });

    return await pdfDoc.save();
}

// ============================================
// MONTHLY PDF REPORT - এক্সেল ফরম্যাটের সাথে মিল রেখে
// ============================================
export async function generateMonthlyPdfReport(dbRows, month) {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([842, 595]); // A4 Landscape
    const { width, height } = page.getSize();

    const fontHelvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontHelveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // কালার (আপনার এক্সেল থেকে)
    const titleColor = argbToRgb('FF1F4E78');
    const companyColor = argbToRgb('FF333333');
    const addressColor = argbToRgb('FF595959');
    const headerBg = argbToRgb('FF2F75B5');
    const darkNav = argbToRgb('FF16365C');
    const evenRowBg = argbToRgb('FFF2F5F9');
    const totalRowBg1 = argbToRgb('FFE9EEF4');
    const totalRowBg2 = argbToRgb('FFD9E1F2');
    const sectionHeaderBg = argbToRgb('FF92969A'); // Approximate for light grey

    // মাস ফরম্যাট
    const monthObj = new Date(month);
    const year = monthObj.getFullYear();
    const monthName = monthObj.toLocaleDateString('en-US', { month: 'long' });
    const reportTitle = `${monthName} ${year}`;

    // টাইটেল
    page.drawText(`Production Report Of ${reportTitle}`, {
        x: 50, y: height - 40, size: 16,
        font: fontHelveticaBold, color: titleColor
    });

    // কোম্পানি নাম
    page.drawText('Zakaria Knitwear Limited (Printing)', {
        x: 50, y: height - 60, size: 12,
        font: fontHelveticaBold, color: companyColor
    });

    // ঠিকানা
    page.drawText('Porabari, Ghatail, Tangail', {
        x: 50, y: height - 78, size: 10,
        font: fontHelvetica, color: addressColor
    });

    // "Production Dates" হেডার
    page.drawText('Production Dates (1 to 31)', {
        x: 200, y: height - 100, size: 10,
        font: fontHelveticaBold, color: rgb(1, 1, 1)
    });

    // টেবিল সেটআপ
    const startX = 40;
    let startY = height - 130;
    const colWidths = {
        sl: 25, buyer: 75, style: 90, printType: 70, cmDzn: 45,
        day: 22, prevQty: 50, currQty: 50, totalCm: 60
    };

    const headerWidth = colWidths.sl + colWidths.buyer + colWidths.style +
                       colWidths.printType + colWidths.cmDzn +
                       (31 * colWidths.day) + colWidths.prevQty + colWidths.currQty + colWidths.totalCm;

    // হেডার ব্যাকগ্রাউন্ড
    page.drawRectangle({
        x: startX, y: startY - 20,
        width: headerWidth, height: 20,
        color: headerBg
    });

    // হেডার টেক্সট (সাদা)
    page.drawText('SL', { x: startX + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Buyer', { x: startX + colWidths.sl + 10, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Style', { x: startX + colWidths.sl + colWidths.buyer + 10, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Print Type', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('CM/Doz ($)', { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });

    // তারিখ হেডার (1-31)
    let currentDaysX = startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.cmDzn;
    for (let d = 1; d <= 31; d++) {
        page.drawText(d.toString(), { x: currentDaysX + 2, y: startY - 14, size: 8, font: fontHelveticaBold, color: rgb(1, 1, 1) });
        currentDaysX += colWidths.day;
    }

    // অতিরিক্ত হেডার
    page.drawText('Prev. Qty', { x: currentDaysX + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Curr. M Qty', { x: currentDaysX + colWidths.prevQty + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });
    page.drawText('Total CM', { x: currentDaysX + colWidths.prevQty + colWidths.currQty + 5, y: startY - 14, size: 10, font: fontHelveticaBold, color: rgb(1, 1, 1) });

    startY -= 20;

    // সেকশন অনুযায়ী ডাটা গ্রুপ করা
    const sections = ['Stone Attached', 'Neck Print', 'Table Print'];

    sections.forEach(sectionName => {
        const sectionRows = dbRows.filter(r => r.print_type === sectionName);
        if (sectionRows.length === 0) return;

        // সেকশন হেডার
        page.drawRectangle({
            x: startX, y: startY - 18,
            width: headerWidth, height: 18,
            color: sectionHeaderBg
        });
        page.drawText(sectionName.toUpperCase(), {
            x: startX + 5, y: startY - 13, size: 9,
            font: fontHelveticaBold, color: rgb(0.2, 0.2, 0.2)
        });
        startY -= 18;

        // ডাটা গ্রুপ করা
        const matrixMap = {};
        sectionRows.forEach(item => {
            const key = `${item.buyer}|||${item.style}`;
            if (!matrixMap[key]) {
                matrixMap[key] = {
                    buyer: item.buyer,
                    style: item.style,
                    cm_dzn: item.cm_dzn,
                    print_type: item.print_type,
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

        let sl = 1;
        Object.values(matrixMap).forEach(row => {
            const isEven = (sl % 2 === 0);
            if (isEven) {
                page.drawRectangle({
                    x: startX, y: startY - 16,
                    width: headerWidth, height: 16,
                    color: evenRowBg
                });
            }

            page.drawText((sl++).toString(), { x: startX + 5, y: startY - 12, size: 9, font: fontHelvetica });
            page.drawText(row.buyer.substring(0, 12), { x: startX + colWidths.sl + 5, y: startY - 12, size: 9, font: fontHelvetica });
            page.drawText(row.style.substring(0, 14), { x: startX + colWidths.sl + colWidths.buyer + 5, y: startY - 12, size: 9, font: fontHelvetica });
            page.drawText(row.print_type.substring(0, 12), { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + 5, y: startY - 12, size: 9, font: fontHelvetica });
            page.drawText(`$${row.cm_dzn.toFixed(2)}`, { x: startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + 5, y: startY - 12, size: 9, font: fontHelvetica });

            let dx = startX + colWidths.sl + colWidths.buyer + colWidths.style + colWidths.printType + colWidths.cmDzn;
            for (let d = 1; d <= 31; d++) {
                if (row.days[d] > 0) {
                    page.drawText(row.days[d].toString(), { x: dx + 1, y: startY - 12, size: 8, font: fontHelvetica });
                }
                dx += colWidths.day;
            }

            const totalCM = (row.cm_dzn / 12) * row.rowTotalPcs;
            page.drawText(row.rowTotalPcs.toString(), { x: dx + 5, y: startY - 12, size: 9, font: fontHelvetica });
            page.drawText(`$${totalCM.toFixed(2)}`, { x: dx + colWidths.prevQty + colWidths.currQty + 5, y: startY - 12, size: 9, font: fontHelvetica });

            page.drawLine({
                start: { x: startX, y: startY - 16 },
                end: { x: width - 40, y: startY - 16 },
                thickness: 0.5, color: rgb(0.85, 0.85, 0.85)
            });

            startY -= 16;
        });
        startY -= 10;
    });

    return await pdfDoc.save();
}

// ব্যাকওয়ার্ড কম্প্যাটিবিলিটির জন্য
export async function generatePdfReport(dbRows) {
    const now = new Date();
    const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    return generateMonthlyPdfReport(dbRows, month);
}
