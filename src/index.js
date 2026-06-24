import { generateExcelReport } from './excel.js';
import { generatePdfReport } from './pdf.js';

// ফ্রন্টএন্ড এইচটিএমএল কোড ডাইনামিক ডাটালিস্ট ও এরর ট্র্যাকিং সহ
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
        .menu-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 30px; }
        .menu-btn { background: white; border: 2px solid var(--primary); color: var(--primary); padding: 30px 20px; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: 600; transition: all 0.2s ease; text-align: center; }
        .menu-btn:hover { background: var(--primary); color: white; transform: translateY(-2px); }
        label { display: block; margin-top: 15px; font-weight: 600; font-size: 14px; color: #475569; }
        input { width: 100%; padding: 12px; margin-top: 6px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; font-size: 15px; }
        input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(30,58,138,0.1); }
        .btn-submit { background: var(--primary); color: white; border: none; padding: 14px; width: 100%; border-radius: 6px; font-size: 16px; font-weight: 600; cursor: pointer; margin-top: 25px; }
        .btn-submit:hover { background: var(--primary-hover); }
        .btn-back { background: #64748b; color: white; border: none; padding: 12px; width: 100%; border-radius: 6px; font-size: 15px; cursor: pointer; margin-top: 10px; }
        .btn-back:hover { background: #475569; }
        .download-box { border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; background: #f8fafc; }
        .download-title { font-weight: 600; font-size: 16px; }
        .btn-dl { background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600; text-decoration: none; }
        .btn-dl:hover { background: #059669; }
        .btn-dl.pdf { background: #ef4444; }
        .btn-dl.pdf:hover { background: #dc2626; }
    </style>
</head>
<body>
<div class="container">
    
    <div id="view-home" class="view active">
        <h2>Zakaria Printing Unit-1 Control Panel</h2>
        <div class="menu-grid">
            <div class="menu-btn" onclick="navigate('view-update')">📝<br><br>Data Update</div>
            <div class="menu-btn" onclick="navigate('view-download')">💾<br><br>Download Reports</div>
        </div>
    </div>

    <div id="view-update" class="view">
        <h2>Production Entry Form</h2>
        <form id="productionForm">
            <label>Production Date</label>
            <input type="date" id="date" required>
            
            <label>Buyer Name</label>
            <input type="text" id="buyer" list="buyer-suggestions" placeholder="Click to see existing or type to filter..." autocomplete="off" required>
            <datalist id="buyer-suggestions"></datalist>
            
            <label>Style No</label>
            <input type="text" id="style" list="style-suggestions" placeholder="Click to see existing or type to filter..." autocomplete="off" required>
            <datalist id="style-suggestions"></datalist>
            
            <label>Print Type / Section</label>
            <input type="text" id="print_type" list="print-type-suggestions" placeholder="Click to see sections or type to filter..." autocomplete="off" required>
            <datalist id="print-type-suggestions"></datalist>

            <label>CM per Dozen ($)</label>
            <input type="number" id="cm_dzn" step="0.0001" placeholder="e.g., 1.40" required>

            <label>Quantity (Pcs)</label>
            <input type="number" id="quantity" placeholder="e.g., 1500" required>

            <button type="submit" class="btn-submit">Save Entry to Cloud D1</button>
            <button type="button" onclick="navigate('view-home')" class="btn-back">Cancel & Go Back</button>
        </form>
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
    function navigate(viewId) {
        document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
        document.getElementById(viewId).classList.add('active');
        if(viewId === 'view-update') {
            loadDatabaseSuggestions();
        }
    }

    // ডাটাবেস থেকে বায়ার, স্টাইল এবং প্রিন্ট টাইপ ডাটা এনে লিস্টে সাজানো
    async function loadDatabaseSuggestions() {
        try {
            const res = await fetch('/api/suggestions');
            const result = await res.json();
            
            if (!res.ok) {
                alert('⚠️ System Alert: Failed to fetch suggestions from DB.\\nReason: ' + (result.error || 'Unknown Cloud Error'));
                return;
            }

            // Populate Buyers
            const buyerList = document.getElementById('buyer-suggestions');
            buyerList.innerHTML = '';
            result.buyers.forEach(name => {
                let opt = document.createElement('option');
                opt.value = name;
                buyerList.appendChild(opt);
            });

            // Populate Styles
            const styleList = document.getElementById('style-suggestions');
            styleList.innerHTML = '';
            result.styles.forEach(styleNo => {
                let opt = document.createElement('option');
                opt.value = styleNo;
                styleList.appendChild(opt);
            });

            // Populate Print Types
            const ptList = document.getElementById('print-type-suggestions');
            ptList.innerHTML = '';
            result.print_types.forEach(pt => {
                let opt = document.createElement('option');
                opt.value = pt;
                ptList.appendChild(opt);
            });

        } catch (err) {
            alert('❌ Network Error: Unable to sync auto-suggestions with D1 Cloud.');
        }
    }

    // ডাটা সাবমিশন এবং অ্যাডভান্সড এরর হ্যান্ডলিং অ্যালার্ট
    document.getElementById('productionForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const data = {
            date: document.getElementById('date').value,
            buyer: document.getElementById('buyer').value,
            style: document.getElementById('style').value,
            print_type: document.getElementById('print_type').value,
            cm_dzn: parseFloat(document.getElementById('cm_dzn').value),
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
                alert('🎉 Success: Data securely saved to D1 SQL Matrix!');
                document.getElementById('productionForm').reset();
                navigate('view-home');
            } else { 
                // ডাটাবেসের সুনির্দিষ্ট এরর মেসেজ ইউজারের সামনে পপআপ করবে
                alert('❌ Database Server Error (Code ' + res.status + '):\\n\\n→ ' + (result.error || 'Submission failed.')); 
            }
        } catch (err) { 
            alert('❌ Critical Network Error: Could not connect to Cloudflare Worker Engine.'); 
        }
    });
</script>
</body>
</html>
`;

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);

        const headers = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type"
        };

        if (request.method === "OPTIONS") {
            return new Response(null, { headers });
        }

        // ROUTE: GET /
        if (request.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
            return new Response(HTML_CONTENT, {
                headers: { ...headers, "Content-Type": "text/html; charset=utf-8" }
            });
        }

        // NEW ROUTE: GET /api/suggestions (অটো ফিল্টারিং ডাটা প্রোভাইডার)
        if (request.method === "GET" && url.pathname === "/api/suggestions") {
            try {
                const buyersData = await env.DB.prepare(`SELECT DISTINCT buyer FROM production WHERE buyer IS NOT NULL AND buyer != '' ORDER BY buyer ASC`).all();
                const stylesData = await env.DB.prepare(`SELECT DISTINCT style FROM production WHERE style IS NOT NULL AND style != '' ORDER BY style ASC`).all();
                const ptData = await env.DB.prepare(`SELECT DISTINCT print_type FROM production WHERE print_type IS NOT NULL AND print_type != '' ORDER BY print_type ASC`).all();

                const buyers = buyersData.results.map(r => r.buyer);
                const styles = stylesData.results.map(r => r.style);
                
                // ডিফল্ট ৩টি ক্যাটাগরি ফিক্সড রাখা হলো, ডাটাবেসে নতুন কিছু যোগ হলে তাও যুক্ত হবে
                const defaultPt = ['Stone Attached', 'Neck Print', 'Table Print'];
                const dbPt = ptData.results.map(r => r.print_type);
                const print_types = Array.from(new Set([...defaultPt, ...dbPt]));

                return new Response(JSON.stringify({ buyers, styles, print_types }), {
                    headers: { ...headers, "Content-Type": "application/json" }
                });
            } catch (err) {
                // যদি টেবিল তৈরি না করা থাকে তবে তার সুনির্দিষ্ট মেসেজ রিটার্ন করবে
                return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { ...headers, "Content-Type": "application/json" } });
            }
        }

        // ROUTE: POST /api/save
        if (request.method === "POST" && url.pathname === "/api/save") {
            try {
                const body = await request.json();
                const { date, buyer, style, print_type, cm_dzn, quantity } = body;

                await env.DB.prepare(
                    `INSERT INTO production (date, buyer, style, print_type, cm_dzn, quantity) VALUES (?, ?, ?, ?, ?, ?)`
                ).bind(date, buyer, style, print_type, cm_dzn, quantity).run();

                return new Response(JSON.stringify({ success: true }), { 
                    headers: { ...headers, "Content-Type": "application/json" } 
                });
            } catch (err) {
                // ক্যাচ ব্লকে JSON এরর রেসপন্স ফিক্স করা হয়েছে যেন ফ্রন্টএন্ডে পড়া যায়
                return new Response(JSON.stringify({ success: false, error: err.message }), { 
                    status: 500, 
                    headers: { ...headers, "Content-Type": "application/json" } 
                });
            }
        }

        // ROUTE: GET /api/excel
        if (request.method === "GET" && url.pathname === "/api/excel") {
            try {
                const { results } = await env.DB.prepare(`SELECT * FROM production ORDER BY print_type ASC, date ASC`).all();
                const buffer = await generateExcelReport(results);
                return new Response(buffer, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Content-Disposition": "attachment; filename=\"Zakaria_Printing_Report_2026.xlsx\""
                    }
                });
            } catch (err) {
                return new Response(JSON.stringify({ error: err.message }), { status: 500, headers });
            }
        }

        // ROUTE: GET /api/pdf
        if (request.method === "GET" && url.pathname === "/api/pdf") {
            try {
                const { results } = await env.DB.prepare(`SELECT * FROM production ORDER BY print_type ASC, date ASC`).all();
                const pdfBytes = await generatePdfReport(results);
                return new Response(pdfBytes, {
                    headers: {
                        ...headers,
                        "Content-Type": "application/pdf",
                        "Content-Disposition": "attachment; filename=\"Zakaria_Printing_Report_2026.pdf\""
                    }
                });
            } catch (err) {
                return new Response(JSON.stringify({ error: err.message }), { status: 500, headers });
            }
        }

        return new Response("Not Found", { status: 404, headers });
    }
};
