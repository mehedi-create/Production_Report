import { generateExcelReport } from './excel.js';
import { generatePdfReport } from './pdf.js';

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);

        // Security / CORS Headers
        const headers = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type"
        };

        if (request.method === "OPTIONS") {
            return new Response(null, { headers });
        }

        // ROUTE: POST /api/save
        if (request.method === "POST" && url.pathname === "/api/save") {
            try {
                const body = await request.json();
                const { date, buyer, style, print_type, cm_dzn, quantity } = body;

                if (!date || !buyer || !style || !print_type || cm_dzn === undefined || !quantity) {
                    return new Response(JSON.stringify({ error: "Missing properties" }), { status: 400, headers });
                }

                await env.DB.prepare(
                    `INSERT INTO production (date, buyer, style, print_type, cm_dzn, quantity) VALUES (?, ?, ?, ?, ?, ?)`
                ).bind(date, buyer, style, print_type, cm_dzn, quantity).run();

                return new Response(JSON.stringify({ success: true }), { headers: { ...headers, "Content-Type": "application/json" } });
            } catch (err) {
                return new Response(JSON.stringify({ error: err.message }), { status: 500, headers });
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
                return new Response("Excel Generator Error: " + err.message, { status: 500, headers });
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
                return new Response("PDF Generator Error: " + err.message, { status: 500, headers });
            }
        }

        // Fallback Asset Router for Cloudflare Pages handling
        return assetFallback(request, env);
    }
};

async function assetFallback(request, env) {
    return fetch(request);
}
