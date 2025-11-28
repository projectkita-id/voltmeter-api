import cors from "cors";
import express from 'express';
import ExcelJS from "exceljs";
import puppeteer from "puppeteer";
import { PrismaClient } from "@prisma/client";

const app = express();
const prisma = new PrismaClient();

app.use(cors())
app.use(express.json());


// GET

app.get("/data", async (req, res) => {
    const log = await prisma.log.findMany();
    res.json(log);
})

app.get("/data/:deviceId", async (req, res) => {
    const { deviceId } = req.params;
    
    try {
        const log = await prisma.log.findMany({
            where: { deviceId }
        });

        res.json({
            deviceId,
            count: log.length,
            data: log
        })
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        res.status(500).json({ error: "Internal server error: " + message })
    }
})

app.get("/data/:deviceId/export", async (req, res) => {
    const { deviceId } = req.params;

    try {
        const latest = await prisma.record.findFirst({
            where: { deviceId },
            orderBy: { timestamp: 'desc' }
        });

        if (!latest) {
            return res.status(404).json({ error: "No records found for this unit" });
        }

        const logs = await prisma.log.findMany({
            where: { recordId: latest.id },
            orderBy: { cell: 'asc' }
        });

        // Excel
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet(`Record ${latest.id}`);

        const blockSize = 31;
        const colWidths = [10, 15, 30];
        const headers = ['Cell', 'Volt', 'Timestamp'];

        const minBlocks = 3;
        const calculatedBlocks = Math.ceil(logs.length / blockSize);
        const numBlocks = Math.max(minBlocks, calculatedBlocks);
        const maxRows = 1 + blockSize;

        for (let block = 0; block < numBlocks; block++) {
            const startCol = block * colWidths.length + 1;

            colWidths.forEach((width, idx) => {
                sheet.getColumn(startCol + idx).width = width;
            });

            headers.forEach((h, idx) => {
                const cell = sheet.getCell(1, startCol + idx);
                cell.value = h;
                cell.font = { bold: true, size: 12 };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0C4DE' } };
                cell.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };
            });

            for (let r = 2; r <= maxRows; r++) {
                for (let idx = 0; idx < colWidths.length; idx++) {
                    const cell = sheet.getCell(r, startCol + idx);
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" },
                    };
                }
            }
        }

        logs.forEach((log: any, i: number) => {
            const block = Math.floor(i / blockSize);
            const posInBlock = i % blockSize;
            const startCol = block * colWidths.length + 1;
            const rowNum = posInBlock + 2;

            const values = [log.cell, log.volt, log.timestamp];
            values.forEach((v, idx) => {
                const cell = sheet.getCell(rowNum, startCol + idx);
                cell.value = v;
            });
        });

        res.setHeader(
            'Content-Disposition',
            `attachment; filename=record_${deviceId}_${latest.id}.xlsx`
        );

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);

        console.error(err);
        res.status(500).json({ error: "Failed to export data: " + message })
    }
});

app.get("/data/:deviceId/export/pdf", async (req, res) => {
    const { deviceId } = req.params;

    const latestRecord = await prisma.record.findFirst({
        where: { deviceId },
        orderBy: { id: "desc" },
        include: { logs: true }
    });

    if (!latestRecord) {
        return res.status(404).json({ error: "No data found" });
    }

    // buat HTML yang nanti dirender jadi PDF
    const html = `
        <html>
        <head>
            <style>
                body { font-family: Arial; padding: 20px; }
                h2 { text-align: center; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td {
                    border: 1px solid #000;
                    padding: 8px;
                    text-align: center;
                }
                th {
                    background: #eee;
                    font-weight: bold;
                }
            </style>
        </head>
        <body>
            <h2>Export Report — ${deviceId}</h2>
            <p>Record ID: ${latestRecord.id}</p>
            <p>Timestamp: ${latestRecord.timestamp}</p>

            <table>
                <tr>
                    <th>Cell</th>
                    <th>Volt</th>
                </tr>
                ${latestRecord.logs
                    .map((log: any) => `
                        <tr>
                            <td>${log.cell}</td>
                            <td>${log.volt}</td>
                        </tr>`)
                    .join("")}
            </table>
        </body>
        </html>
    `;

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    await page.setContent(html);
    const pdfBuffer = await page.pdf({ format: "A4" });
    await browser.close();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${deviceId}.pdf"`);
    res.send(pdfBuffer);
})

// POST

app.post("/data", async (req, res) => {
    const { deviceId, cell, volt } = req.body;

    if (!deviceId || !cell || !Array.isArray(volt)) {
        return res.status(400).json({ error: "Invalid request body" });
    }

    if (volt.length !== cell) {
        return res.status(400).json({ error: "Volt array length does not match cell count" });
    }

    try {
        const previousCount = await prisma.record.count({
            where: { deviceId }
        });
        
        const now = new Date();
        const desc = `Pengambilan ke-${previousCount + 1}`;

        const record = await prisma.record.create({
            data: {
                deviceId,
                cell,
                timestamp: now,
                detail: desc
            }
        });

        const logs = [];

        for (let i = 0; i < cell; i++) {
            const log = await prisma.log.create({
                data: {
                    deviceId,
                    cell: i + 1,
                    volt: volt[i],
                    timestamp: now,
                    recordId: record.id
                }
            });

            logs.push(log);
        }

        res.json({
            message: "Data recorded successfully",
            record,
            data: logs
        })

    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        res.status(500).json({ error: "Internal server error: " + message });
    }

})

app.listen(3000, () => {
    console.log("Server is running on http://localhost:3000");
})
