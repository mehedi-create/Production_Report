import { generateExcelReport } from './excel.js';
import { generatePdfReport } from './pdf.js';

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
        
        /* Custom Smooth Autocomplete Dropdown */
        .autocomplete-container { position: relative; }
        .autocomplete-items { position: absolute; border: 1px solid #cbd5e1; border-top: none; z-index: 99; top: 100%; left: 0; right: 0; max-height: 200px; overflow-y: auto; background-color: #fff; border-radius: 0 0 6px 6px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        .autocomplete-items div { padding: 12px; cursor: pointer; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
        .autocomplete-items div:hover { background-color: #f8fafc; color: var(--primary); font-weight: 600; }
        
        /* Buttons */
        .btn-submit { background: var(--primary); color: white; border: none; padding: 14px; width: 100%; border-radius: 6px; font-size: 16px; font-weight: 600; cursor: pointer; margin-top: 25px; }
        .btn-submit:hover { background: var(--primary-hover); }
        .btn-back { background: #64748b; color: white; border: none; padding: 12px; width: 100%; border-radius: 6px; font-size: 15px; cursor: pointer; margin-top: 10px; }
        .btn-back:hover { background: #475569; }
        
        /* Downloads Box */
        .download-box { border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; background: #f8fafc; }
        .btn-dl { background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600; text-decoration: none; }
        .btn-dl:hover { background: #059669; }
        .btn-dl.pdf { background: #ef4444; }
        .btn-dl.pdf:hover { background: #dc2626; }
        
        /* History Table */
        .table-responsive { overflow-x: auto; margin-top: 15px; border-radius: 8px; border: 1px solid #cbd5e1; }
        table { width: 100%; border-collapse: collapse; min-width: 500px; font-size: 14px; }
        th, td { border-bottom: 1px solid #cbd5e1; padding: 12px; text-align: left; }
        th { background-color: #f8fafc; color: #475569; }
        .btn-delete { background-color: #ef4444; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; }
        .btn-delete:hover { background-color: #dc2626; }

        /* Notification UI */
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
        <h2>Report Export Engine</h2>
        <div class="download-box">
            <div>
                <div class="download-title">Excel Master Spreadsheet</div>
                <small style="color:#64748b;">Full month grid matrix, auto formulas</small>
            </div>
            <a href="/api/excel" target="_blank" class="btn-dl">Export XLSX</a>
        </div>
        <div class="download-box">
            <div>
                <div class="download-title">PDF Executive Report</div>
                <small style="color:#64748b;">Formatted matrix ready for printing</small>
            </div>
            <a href="/api/pdf" target="_blank" class="btn-dl pdf">Export PDF</a>
        </div>
        <button type="button" onclick="navigate('view-home')" class="btn-back">Back to Home</button>
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
        
        if(viewId === 'view-update') loadDatabaseSuggestions();
        if(viewId === 'view-history') loadHistory();
    }

    // Custom Smooth Autocomplete Logic (Replaces the ugly datalist)
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

    // Initialize Autocomplete
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

    // Load History Data
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

    // Delete Entry Logic
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
                loadHistory(); // Reload history table instantly
            } else {
                showNotification('Error deleting: ' + result.error, 'error');
            }
        } catch(e) {
            showNotification('Network Error while deleting.', 'error');
        }
    }

    // Save Form Logic (With Optional CM Dzn Support)
    document.getElementById('productionForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const data = {
            date: document.getElementById('date').value,
            buyer: document.getElementById('buyer').value,
            style: document.getElementById('style').value,
            print_type: document.getElementById('print_type').value,
            cm_dzn: document.getElementById('cm_dzn').value, // This can now be empty
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
                // তারিখ অনুযায়ী নতুন এন্ট্রিগুলো সবার ওপরে দেখাবে
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

        // Save Form API (Handles Optional CM Dzn automatically)
        if (request.method === "POST" && url.pathname === "/api/save") {
            try {
                const body = await request.json();
                let finalCmDzn = body.cm_dzn;

                // যদি ইউজার প্রাইস ফাঁকা রাখে, ডাটাবেস থেকে পূর্বের প্রাইস খুঁজে বের করবে
                if (finalCmDzn === "" || finalCmDzn === null || finalCmDzn === undefined) {
                    const previousEntry = await env.DB.prepare("SELECT cm_dzn FROM production WHERE buyer = ? AND style = ? AND cm_dzn IS NOT NULL ORDER BY id DESC LIMIT 1")
                        .bind(body.buyer, body.style).first();
                    
                    if (previousEntry) {
                        finalCmDzn = previousEntry.cm_dzn; // আগের প্রাইস বসিয়ে দিল
                    } else {
                        // যদি নতুন স্টাইল হয় এবং আগে প্রাইস না থাকে
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

        if (request.method === "GET" && url.pathname === "/api/excel") {
            try {
                const { results } = await env.DB.prepare("SELECT * FROM production ORDER BY print_type ASC, date ASC").all();
                const buffer = await generateExcelReport(results);
                return new Response(buffer, { headers: { ...headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": "attachment; filename=\"ZKL_Printing_Report.xlsx\"" } });
            } catch (err) { return new Response(JSON.stringify({ error: err.message }), { status: 500, headers }); }
        }

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
