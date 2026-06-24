import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';

export async function generatePdfReport(dbRows) {
    const pdfDoc = await PDFDocument.create();
    // A4 Horizontal Landscape View Orientation Specification
    const page = pdfDoc.addPage([842, 595]); 
    const { width, height } = page.getSize();
    
    const fontHelvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontHelveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // Document Titles Header Block
    page.drawText('Zakaria Printing Unit-1', { x: 50, y: height - 40, size: 18, font: fontHelveticaBold, color: rgb(0.12, 0.23, 0.54) });
    page.drawText('MONTHLY Printing PRODUCTION SUMMARY, June-2026', { x: 50, y: height - 60, size: 11, font: fontHelvetica });

    // Table Header Structural Settings
    const startX = 40;
    let startY = height - 90;
    const colWidths = { sl: 25, buyer: 75, style: 90, days: 14, totalPcs: 65, cmDzn: 45, totalCm: 65 };
    
    // Draw Primary Headers Block Text Matrix
    page.drawRectangle({ x: startX, y: startY - 20, width: width - 80, height: 20, color: rgb(0.2, 0.28, 0.37) });
    
    page.drawText('SL', { x: startX + 5, y: startY - 14, size: 9, font: fontHelveticaBold, color: rgb(1,1,1) });
    page.drawText('Buyer', { x: startX + colWidths.sl + 10, y: startY - 14, size: 9, font: fontHelveticaBold, color: rgb(1,1,1) });
    page.drawText('Style No', { x: startX + colWidths.sl + colWidths.buyer + 10, y: startY - 14, size: 9, font: fontHelveticaBold, color: rgb(1,1,1) });

    let currentDaysX = startX + colWidths.sl + colWidths.buyer + colWidths.style;
    for (let d = 1; d <= 31; d++) {
        page.drawText(d.toString(), { x: currentDaysX + 2, y: startY - 14, size: 7, font: fontHelveticaBold, color: rgb(1,1,1) });
        currentDaysX += colWidths.days;
    }
    
    page.drawText('Total Pcs', { x: currentDaysX + 5, y: startY - 14, size: 8, font: fontHelveticaBold, color: rgb(1,1,1) });
    page.drawText('CM Dzn', { x: currentDaysX + colWidths.totalPcs + 5, y: startY - 14, size: 8, font: fontHelveticaBold, color: rgb(1,1,1) });
    page.drawText('Total CM', { x: currentDaysX + colWidths.totalPcs + colWidths.cmDzn + 5, y: startY - 14, size: 8, font: fontHelveticaBold, color: rgb(1,1,1) });

    startY -= 20;

    const sections = ['Stone Attached', 'Neck Print', 'Table Print'];
    
    sections.forEach((sectionName) => {
        const sectionRows = dbRows.filter(r => r.print_type === sectionName);
        if (sectionRows.length === 0) return;

        // Draw Section Banner Row
        page.drawRectangle({ x: startX, y: startY - 18, width: width - 80, height: 18, color: rgb(0.92, 0.93, 0.95) });
        page.drawText(sectionName.toUpperCase(), { x: startX + 5, y: startY - 13, size: 9, font: fontHelveticaBold, color: rgb(0.1, 0.2, 0.3) });
        startY -= 18;

        // Group rows internally inside specific blocks
        const matrixMap = {};
        sectionRows.forEach(item => {
            const key = `${item.buyer}|||${item.style}`;
            if (!matrixMap[key]) {
                matrixMap[key] = { buyer: item.buyer, style: item.style, cm_dzn: item.cm_dzn, days: Array(32).fill(0), rowTotalPcs: 0 };
            }
            const dayNum = new Date(item.date).getDate();
            if (dayNum >= 1 && dayNum <= 31) {
                matrixMap[key].days[dayNum] += item.quantity;
                matrixMap[key].rowTotalPcs += item.quantity;
            }
        });

        let sl = 1;
        Object.values(matrixMap).forEach((row) => {
            // Alternating Row Background Color System
            if (sl % 2 === 0) {
                page.drawRectangle({ x: startX, y: startY - 16, width: width - 80, height: 16, color: rgb(0.98, 0.98, 0.99) });
            }

            // Draw Inner Boundaries / Text Content Matrix Cells
            page.drawText((sl++).toString(), { x: startX + 5, y: startY - 12, size: 8, font: fontHelvetica });
            page.drawText(row.buyer.substring(0, 12), { x: startX + colWidths.sl + 5, y: startY - 12, size: 8, font: fontHelvetica });
            page.drawText(row.style.substring(0, 14), { x: startX + colWidths.sl + colWidths.buyer + 5, y: startY - 12, size: 8, font: fontHelvetica });

            let dx = startX + colWidths.sl + colWidths.buyer + colWidths.style;
            for (let d = 1; d <= 31; d++) {
                if (row.days[d] > 0) {
                    page.drawText(row.days[d].toString(), { x: dx + 1, y: startY - 12, size: 7, font: fontHelvetica, color: rgb(0.2, 0.2, 0.2) });
                }
                dx += colWidths.days;
            }

            const totalCM = (row.cm_dzn / 12) * row.rowTotalPcs;
            page.drawText(row.rowTotalPcs.toString(), { x: dx + 5, y: startY - 12, size: 8, font: fontHelvetica });
            page.drawText(`$${row.cm_dzn.toFixed(2)}`, { x: dx + colWidths.totalPcs + 5, y: startY - 12, size: 8, font: fontHelvetica });
            page.drawText(`$${totalCM.toFixed(2)}`, { x: dx + colWidths.totalPcs + colWidths.cmDzn + 5, y: startY - 12, size: 8, font: fontHelvetica });

            // Bottom border line per transaction row
            page.drawLine({ start: { x: startX, y: startY - 16 }, end: { x: width - 40, y: startY - 16 }, thickness: 0.5, color: rgb(0.85, 0.85, 0.85) });

            startY -= 16;
        });
        startY -= 10; // Extra spacing separation line block
    });

    return await pdfDoc.save();
}
