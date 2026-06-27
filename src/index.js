import { generateDailyExcelReport, generateMonthlyExcelReport, generateExcelReport } from './excel.js';
import { generateDailyPdfReport, generateMonthlyPdfReport, generatePdfReport } from './pdf.js';

const HTML_CONTENT = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ZKL Printing Production Reporting System</title>
    <style>
        :root { --primary: #1e3a8a; --primary-hover: #172554; --bg: #f8fafc; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); margin: 0; padding: 20px; color: #1e293b; }
        .container { max-width: 700px; margin: 40px auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05); }
        .view { display: none; }
        .active { display: block; }
        h2 { text-align: center; color: var(--primary); margin-bottom: 25px; font-weight: 700; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }

        .menu-grid { display: flex; flex-direction: column; gap: 15px; margin-top: 30px; }
        .menu-btn { background: white; border: 2px solid var(--primary); color: var(--primary); padding: 20px; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: 600; transition: all 0.2s ease; text-align: center; }
        .menu-btn:hover { background: var(--primary); color: white; transform: translateY(-2px); }

        label { display: block; margin-top: 15px; font-weight: 600; font-size: 14px; color: #475569; }
        input { width: 100%; padding: 12px; margin-top: 6px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; font-size: 15px; }
        input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(30,58,138,0.1); }

        .autocomplete-container { position: relative; }
        .autocomplete-items { position: absolute; border: 1px solid #cbd5e1; border-top: none; z-index: 99; top: 100%; left: 0; right: 0; max-height: 200px; overflow-y: auto; background-color: #fff; border-radius: 0 0 6px 6px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        .autocomplete-items div { padding: 12px; cursor: pointer; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
        .autocomplete-items div:hover { background-color: #f8fafc; color: var(--primary); font-weight: 600; }

        .btn-submit { background: var(--primary); color: white; border: none; padding: 14px; width: 100%; border-radius: 6px; font-size: 16px; font-weight: 600; cursor: pointer; margin-top: 25px; }
        .btn-submit:hover { background: var(--primary-hover); }
        .btn-back { background: #64748b; color: white; border: none; padding: 12px; width: 100%; border-radius: 6px; font-size: 15px; cursor: pointer; margin-top: 10px; }
        .btn-back:hover { background: #475569; }

        .table-responsive { overflow-x: auto; margin-top: 15px; border-radius: 8px; border: 1px solid #cbd5e1; }
        table { width: 100%; border-collapse: collapse; min-width: 500px; font-size: 14px; }
        th, td { border-bottom: 1px solid #cbd5e1; padding: 12px; text-align: left; }
        th { background-color: #f8fafc; color: #475569; }
        .btn-delete { background-color: #ef4444; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; }
        .btn-delete:hover { background-color: #dc2626; }

        #notification {
            position: fixed; top: -100px; left: 50%; transform: translateX(-50%);
            background: white; padding: 16px 24px; border-radius: 8px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15); transition: top 0.4s cubic-bezier(0.68, -0.55, 0.265, 1.55);
            z-index: 99999; display: flex; align-items: center; gap: 12px; font-weight: 600;
            width: 85%; max-width: 450px; font-size: 15px; line-height: 1.4;
        }
        #notification.show { top: 30px; }
        #notification.error { border-left: 6px solid #ef4444; color: #b91c1c; }
        #notification.success { border-left: 6px solid #10b981; color: #047857; }
        .icon { font-size: 20px; }

        .report-option {
            display: flex;
            align-items: center;
            padding: 20px;
            margin: 15px 0;
            border: 2px solid #e2e8f0;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.2s ease;
            background: white;
        }
        .report-option:hover {
            border-color: var(--primary);
            box-shadow: 0 4px 12px rgba(30,58,138,0.1);
        }
        .report-icon {
            font-size: 28px;
            margin-right: 15px;
            color: var(--primary);
        }
        .report-content {
            flex: 1;
        }
        .report-title {
            font-size: 18px;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 5px;
        }
        .report-desc {
            font-size: 14px;
            color: #64748b;
        }
        .download-arrow {
            color: #94a3b8;
            font-size: 20px;
        }

        .format-selector {
            margin-top: 20px;
            padding: 15px;
            background: #f8fafc;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
        }
        .format-buttons {
            display: flex;
            gap: 10px;
            margin-top: 15px;
        }
        .format-btn {
            flex: 1;
            padding: 12px;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .format-btn.excel {
            background: #10b981;
            color: white;
        }
        .format-btn.excel:hover {
            background: #059669;
        }
        .format-btn.pdf {
            background: #ef4444;
            color: white;
        }
        .format-btn.pdf:hover {
            background: #dc2626;
        }

        .date-selector, .month-selector {
            margin-top: 20px;
            padding: 15px;
            background: #f8fafc;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
        }
        .date-selector label, .month-selector label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #475569;
        }
        .date-selector input, .month-selector input {
            width: 100%;
            padding: 10px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 14px;
        }
    </style>
</head>
<body>

<div id="notification">
    <span class="icon" id="noti-icon"></span>
    <span id="noti-text"></span>
</div>

<div class="container">

    <div id="view-home" class="view active">
        <h2>Zakaria Printing Unit-1</h2>
        <div class="menu-grid">
            <div class="menu-btn" onclick="navigate('view-update')">📝 Data Update</div>
            <div class="menu-btn" onclick="navigate('view-history')" style="color: #0ea5e9; border-color: #0ea5e9;">📊 Data Control & History</div>
            <div class="menu-btn" onclick="navigate('view-download')" style="color: #10b981; border-color: #10b981;">💾 Download Reports</div>
        </div>
    </div>

    <div id="view-update" class="view">
        <h2>Production Entry Form</h2>
        <form id="productionForm" autocomplete="off">
            <label for="date">Production Date</label>
            <input type="date" id="date" required>

            <label for="buyer">Buyer Name</label>
            <div class="autocomplete-container">
                <input type="text" id="buyer" placeholder="Click to see existing or type to search..." required>
            </div>

            <label for="style">Style No</label>
            <div class="autocomplete-container">
                <input type="text" id="style" placeholder="Click to see existing or type to search..." required>
            </div>

            <label for="print_type">Print Type / Section</label>
            <div class="autocomplete-container">
                <input type="text" id="print_type" placeholder="Click to see existing or type to search..." required>
            </div>

            <label for="cm_dzn">CM per Dozen ($) <span style="color:#94a3b8; font-weight:normal;">(Optional if already saved before)</span></label>
            <input type="number" id="cm_dzn" step="0.0001" placeholder="e.g., 1.40 (Leave empty to use last price)">

            <label for="quantity">Quantity (Pcs)</label>
            <input type="number" id="quantity" placeholder="e.g., 1500" required>

            <button type="submit" class="btn-submit">Save Entry to Cloud D1</button>
            <button type="button" onclick="navigate('view-home')" class="btn-back">Cancel & Go Back</button>
        </form>
    </div>

    <div id="view-history" class="view">
        <h2>Data History & Control</h2>
        <p style="font-size: 14px; color: #64748b; text-align:center;">Recent updates are shown at the top.</p>

        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Buyer</th>
                        <th>Style</th>
                        <th>Qty</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody id="history-tbody">
                    <tr><td colspan="5" style="text-align:center;">Loading...</td></tr>
                </tbody>
            </table>
        </div>
        <button type="button" onclick="navigate('view-home')" class="btn-back" style="margin-top: 20px;">Back to Home</button>
    </div>

    <div id="view-download" class="view">
        <h2>Report Download Options</h2>

        <div class="report-option" onclick="showDailyOptions()">
            <div class="report-icon">📥</div>
            <div class="report-content">
                <div class="report-title">ডাউনলোড ডেইলি রিপোর্ট</div>
                <div class="report-desc">Select a specific date to download daily production report</div>
            </div>
            <div class="download-arrow">→</div>
        </div>

        <div class="report-option" onclick="showMonthlyOptions()">
            <div class="report-icon">📥</div>
            <div class="report-content">
                <div class="report-title">ডাউনলোড মান্থলি রিপোর্ট</div>
                <div class="report-desc">Select a specific month to download monthly production report</div>
            </div>
            <div class="download-arrow">→</div>
        </div>

        <div id="daily-section" style="display: none;">
            <h2>Daily Report Download</h2>

            <div class="date-selector">
                <label for="daily-date">Select Date</label>
                <input type="date" id="daily-date" required>
            </div>

            <div class="format-selector">
                <div class="format-buttons">
                    <button class="format-btn excel" onclick="downloadDailyReport('excel')">📊 Download Excel</button>
                    <button class="format-btn pdf" onclick="downloadDailyReport('pdf')">📄 Download PDF</button>
                </div>
            </div>

            <button type="button" onclick="backToDownloadOptions()" class="btn-back">Back to Options</button>
        </div>

        <div id="monthly-section" style="display: none;">
            <h2>Monthly Report Download</h2>

            <div class="month-selector">
                <label for="month-year">Select Month and Year</label>
                <input type="month" id="month-year" required>
            </div>

            <div class="format-selector">
                <div class="format-buttons">
                    <button class="format-btn excel" onclick="downloadMonthlyReport('excel')">📊 Download Excel</button>
                    <button class="format-btn pdf" onclick="downloadMonthlyReport('pdf')">📄 Download PDF</button>
                </div>
            </div>

            <button type="button" onclick="backToDownloadOptions()" class="btn-back">Back to Options</button>
        </div>
    </div>
</div>

<script>
    let dbData = { buyers: [], styles: [], print_types: [] };

    function showNotification(message, type) {
        const noti = document.getElementById('notification');
        const notiText = document.getElementById('noti-text');
        const notiIcon = document.getElementById('noti-icon');
        noti.className = type === 'error' ? 'error' : 'success';
        notiText.innerText = message;
        notiIcon.innerText = type === 'error' ? '⚠️' : '✅';
        noti.classList.add('show');
        setTimeout(() => noti.classList.remove('show'), 3500);
    }

    function navigate(viewId) {
        document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
        document.getElementById(viewId).classList.add('active');

        document.getElementById('daily-section').style.display = 'none';
        document.getElementById('monthly-section').style.display = 'none';

        if(viewId === 'view-update') loadDatabaseSuggestions();
        if(viewId === 'view-history') loadHistory();
    }

    function showDailyOptions() {
        document.getElementById('daily-section').style.display = 'block';
        document.getElementById('monthly-section').style.display = 'none';

        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        document.getElementById('daily-date').value = year + '-' + month + '-' + day;
    }

    function showMonthlyOptions() {
        document.getElementById('monthly-section').style.display = 'block';
        document.getElementById('daily-section').style.display = 'none';

        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        document.getElementById('month-year').value = year + '-' + month;
    }

    function backToDownloadOptions() {
        document.getElementById('daily-section').style.display = 'none';
        document.getElementById('monthly-section').style.display = 'none';
    }

    function downloadDailyReport(format) {
        const date = document.getElementById('daily-date').value;
        if (!date) {
            alert('Please select a date');
            return;
        }

        if (format === 'excel') {
            window.location.href = '/api/daily-excel?date=' + date;
        } else if (format === 'pdf') {
            window.location.href = '/api/daily-pdf?date=' + date;
        }
    }

    function downloadMonthlyReport(format) {
        const monthYear = document.getElementById('month-year').value;
        if (!monthYear) {
            alert('Please select a month');
            return;
        }

        if (format === 'excel') {
            window.location.href = '/api/monthly-excel?month=' + monthYear;
        } else if (format === 'pdf') {
            window.location.href = '/api/monthly-pdf?month=' + monthYear;
        }
    }

    function setupAutocomplete(inputId, dataType) {
        const inp = document.getElementById(inputId);

        inp.addEventListener("input", function() {
            let a, b, val = this.value;
            closeAllLists();

            a = document.createElement("DIV");
            a.setAttribute("id", this.id + "autocomplete-list");
            a.setAttribute("class", "autocomplete-items");
            this.parentNode.appendChild(a);

            let arr = [];
            if(dataType === 'buyer') arr = dbData.buyers;
            if(dataType === 'style') arr = dbData.styles;
            if(dataType === 'print_type') arr = dbData.print_types;

            let matches = arr;
            if(val) matches = arr.filter(item => item.toUpperCase().includes(val.toUpperCase()));

            for (let i = 0; i < matches.length; i++) {
                b = document.createElement("DIV");
                b.innerHTML = matches[i];
                b.innerHTML += "<input type='hidden' value='" + matches[i].replace(/'/g, "&#39;") + "'>";
                b.addEventListener("click", function() {
                    inp.value = this.getElementsByTagName("input")[0].value;
                    closeAllLists();
                });
                a.appendChild(b);
            }
        });

        inp.addEventListener("click", function() {
            this.dispatchEvent(new Event('input'));
        });
    }

    function closeAllLists(elmnt) {
        var x = document.getElementsByClassName("autocomplete-items");
        for (var i = 0; i < x.length; i++) {
            if (elmnt != x[i]) x[i].parentNode.removeChild(x[i]);
        }
    }
    document.addEventListener("click", function (e) {
        if(e.target.tagName !== "INPUT") closeAllLists();
    });

    setupAutocomplete("buyer", "buyer");
    setupAutocomplete("style", "style");
    setupAutocomplete("print_type", "print_type");

    async function loadDatabaseSuggestions() {
        try {
            const res = await fetch('/api/suggestions');
            const result = await res.json();
            if (res.ok) dbData = result;
        } catch (err) { console.error("Could not load suggestions."); }
    }

    async function loadHistory() {
        try {
            const res = await fetch('/api/history');
            const data = await res.json();
            const tbody = document.getElementById('history-tbody');
            if(data.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No data found.</td></tr>';
                return;
            }
            tbody.innerHTML = data.map(row =>
                '<tr>' +
                    '<td>' + row.date + '</td>' +
                    '<td>' + row.buyer + '</td>' +
                    '<td>' + row.style + '</td>' +
                    '<td>' + row.quantity + '</td>' +
                    '<td><button class="btn-delete" onclick="deleteEntry(' + row.id + ')">Delete</button></td>' +
                '</tr>'
            ).join('');
        } catch(err) {
            showNotification('Failed to load history', 'error');
        }
    }

    async function deleteEntry(id) {
        if(!confirm("Are you sure you want to delete this entry? This cannot be undone.")) return;
        try {
            const res = await fetch('/api/delete', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id })
            });
            const result = await res.json();
            if(result.success) {
                showNotification('Entry deleted successfully', 'success');
                loadHistory();
            } else {
                showNotification('Error deleting: ' + result.error, 'error');
            }
        } catch(e) {
            showNotification('Network Error while deleting.', 'error');
        }
    }

    document.getElementById('productionForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const data = {
            date: document.getElementById('date').value,
            buyer: document.getElementById('buyer').value,
            style: document.getElementById('style').value,
            print_type: document.getElementById('print_type').value,
            cm_dzn: document.getElementById('cm_dzn').value,
            quantity: parseInt(document.getElementById('quantity').value)
        };

        try {
            const res = await fetch('/api/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
            const result = await res.json();

            if (res.ok && result.success) {
                showNotification('Success: Data securely saved!', 'success');
                document.getElementById('productionForm').reset();
                navigate('view-home');
            } else {
                showNotification('Error: ' + (result.error || 'Submission failed.'), 'error');
            }
        } catch (err) {
            showNotification('Critical Error: Could not connect to Server.', 'error');
        }
    });
</script>
</body>
</html>
`;

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);
        const headers = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "GET, POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" };

        if (request.method === "OPTIONS") return new Response(null, { headers });

        if (request.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
            return new Response(HTML_CONTENT, { headers: { ...headers, "Content-Type": "text/html; charset=utf-8" } });
        }

        // Auto-suggestions API
        if (request.method === "GET" && url.pathname === "/api/suggestions") {
            try {
                const buyersData = await env.DB.prepare("SELECT DISTINCT buyer FROM production WHERE buyer IS NOT NULL AND buyer != '' ORDER BY buyer ASC").all();
                const stylesData = await env.DB.prepare("SELECT DISTINCT style FROM production WHERE style IS NOT NULL AND style != '' ORDER BY style ASC").all();
                const ptData = await env.DB.prepare("SELECT DISTINCT print_type FROM production WHERE print_type IS NOT NULL AND print_type != '' ORDER BY print_type ASC").all();

                return new Response(JSON.stringify({
                    buyers: buyersData.results.map(r => r.buyer),
                    styles: stylesData.results.map(r => r.style),
                    print_types: Array.from(new Set([...['Stone Attached', 'Neck Print', 'Table Print'], ...ptData.results.map(r => r.print_type)]))
                }), { headers: { ...headers, "Content-Type": "application/json" } });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { ...headers, "Content-Type": "application/json" } }); }
        }

        // History API
        if (request.method === "GET" && url.pathname === "/api/history") {
            try {
                const { results } = await env.DB.prepare("SELECT id, date, buyer, style, quantity FROM production ORDER BY date DESC, id DESC LIMIT 200").all();
                return new Response(JSON.stringify(results), { headers: { ...headers, "Content-Type": "application/json" } });
            } catch (err) {
                return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { ...headers, "Content-Type": "application/json" } });
            }
        }

        // Delete Entry API
        if (request.method === "POST" && url.pathname === "/api/delete") {
            try {
                const { id } = await request.json();
                await env.DB.prepare("DELETE FROM production WHERE id = ?").bind(id).run();
                return new Response(JSON.stringify({ success: true }), { headers: { ...headers, "Content-Type": "application/json" } });
            } catch (err) {
                return new Response(JSON.stringify({ success: false, error: err.message }), { status: 500, headers: { ...headers, "Content-Type": "application/json" } });
            }
        }

        // Save Form API
        if (request.method === "POST" && url.pathname === "/api/save") {
            try {
                const body = await request.json();
                let finalCmDzn = body.cm_dzn;

                if (finalCmDzn === "" || finalCmDzn === null || finalCmDzn === undefined) {
                    const previousEntry = await env.DB.prepare("SELECT cm_dzn FROM production WHERE buyer = ? AND style = ? AND cm_dzn IS NOT NULL ORDER BY id DESC LIMIT 1")
                        .bind(body.buyer, body.style).first();

                    if (previousEntry) {
                        finalCmDzn = previousEntry.cm_dzn;
                    } else {
                        return new Response(JSON.stringify({ success: false, error: "This is a new Style. Please provide CM per Dozen ($) for the first time!" }), { status: 400, headers: { ...headers, "Content-Type": "application/json" } });
                    }
                } else {
                    finalCmDzn = parseFloat(finalCmDzn);
                }

                await env.DB.prepare("INSERT INTO production (date, buyer, style, print_type, cm_dzn, quantity) VALUES (?, ?, ?, ?, ?, ?)")
                    .bind(body.date, body.buyer, body.style, body.print_type, finalCmDzn, body.quantity).run();

                return new Response(JSON.stringify({ success: true }), { headers: { ...headers, "Content-Type": "application/json" } });
            } catch (err) { return new Response(JSON.stringify({ success: false, error: err.message }), { status: 500, headers: { ...headers, "Content-Type": "application/json" } }); }
        }

        // ==================== NEW API ENDPOINTS ====================

        // Daily Excel Report
        if (request.method === "GET" && url.pathname === "/api/daily-excel") {
            try {
                const date = url.searchParams.get('date');
                if (!date) {
                    return new Response(JSON.stringify({ error: 'Date parameter is required' }), { status: 400, headers: { ...headers, "Content-Type": "application/json" } });
                }

                const { results } = await env.DB.prepare("SELECT * FROM production WHERE date = ? ORDER BY print_type ASC, buyer ASC, style ASC").bind(date).all();
                const buffer = await generateDailyExcelReport(results, date);

                const dateObj = new Date(date);
                const formattedDate = dateObj.toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' }).replace(/\//g, '-');

                return new Response(buffer, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Content-Disposition": "attachment; filename=\"ZKL_Daily_Report_" + formattedDate + ".xlsx\""
                    }
                });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        // Daily PDF Report
        if (request.method === "GET" && url.pathname === "/api/daily-pdf") {
            try {
                const date = url.searchParams.get('date');
                if (!date) {
                    return new Response(JSON.stringify({ error: 'Date parameter is required' }), { status: 400, headers: { ...headers, "Content-Type": "application/json" } });
                }

                const { results } = await env.DB.prepare("SELECT * FROM production WHERE date = ? ORDER BY print_type ASC, buyer ASC, style ASC").bind(date).all();
                const pdfBytes = await generateDailyPdfReport(results, date);

                const dateObj = new Date(date);
                const formattedDate = dateObj.toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' }).replace(/\//g, '-');

                return new Response(pdfBytes, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/pdf",
                        "Content-Disposition": "attachment; filename=\"ZKL_Daily_Report_" + formattedDate + ".pdf\""
                    }
                });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        // Monthly Excel Report
        if (request.method === "GET" && url.pathname === "/api/monthly-excel") {
            try {
                const month = url.searchParams.get('month');
                if (!month) {
                    return new Response(JSON.stringify({ error: 'Month parameter is required (format: YYYY-MM)' }), { status: 400, headers: { ...headers, "Content-Type": "application/json" } });
                }

                const { results } = await env.DB.prepare("SELECT * FROM production WHERE strftime('%Y-%m', date) = ? ORDER BY print_type ASC, date ASC").bind(month).all();
                const buffer = await generateMonthlyExcelReport(results, month);

                const monthObj = new Date(month);
                const monthName = monthObj.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(/\//g, '-');

                return new Response(buffer, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Content-Disposition": "attachment; filename=\"ZKL_Monthly_Report_" + monthName + ".xlsx\""
                    }
                });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        // Monthly PDF Report
        if (request.method === "GET" && url.pathname === "/api/monthly-pdf") {
            try {
                const month = url.searchParams.get('month');
                if (!month) {
                    return new Response(JSON.stringify({ error: 'Month parameter is required (format: YYYY-MM)' }), { status: 400, headers: { ...headers, "Content-Type": "application/json" } });
                }

                const { results } = await env.DB.prepare("SELECT * FROM production WHERE strftime('%Y-%m', date) = ? ORDER BY print_type ASC, date ASC").bind(month).all();
                const pdfBytes = await generateMonthlyPdfReport(results, month);

                const monthObj = new Date(month);
                const monthName = monthObj.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(/\//g, '-');

                return new Response(pdfBytes, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/pdf",
                        "Content-Disposition": "attachment; filename=\"ZKL_Monthly_Report_" + monthName + ".pdf\""
                    }
                });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        // Original Excel API (for backward compatibility)
        if (request.method === "GET" && url.pathname === "/api/excel") {
            try {
                const { results } = await env.DB.prepare("SELECT * FROM production ORDER BY print_type ASC, date ASC").all();
                const buffer = await generateExcelReport(results);
                return new Response(buffer, { headers: { ...headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": "attachment; filename=\"ZKL_Printing_Report.xlsx\"" } });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        // Original PDF API (for backward compatibility)
        if (request.method === "GET" && url.pathname === "/api/pdf") {
            try {
                const { results } = await env.DB.prepare("SELECT * FROM production ORDER BY print_type ASC, date ASC").all();
                const pdfBytes = await generatePdfReport(results);
                return new Response(pdfBytes, { headers: { ...headers, "Content-Type": "application/pdf", "Content-Disposition": "attachment; filename=\"ZKL_Printing_Report.pdf\"" } });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

        return new Response("Not Found", { status: 404, headers });
    }
};
