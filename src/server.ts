import cors from "cors";
import express from 'express';
import ExcelJS from "exceljs";
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

        sheet.columns = [
            { header: 'Cell', key: 'cell', width: 10 },
            { header: 'Volt', key: 'volt', width: 15 },
            { header: 'Timestamp', key: 'timestamp', width: 25 }
        ];

        const header = sheet.getRow(1);
        header.font = { bold: true };
        header.alignment = { horizontal: 'center', vertical: 'middle' };
        header.commit();

        logs.forEach((log) => {
            sheet.addRow({
                cell: log.cell,
                volt: log.volt,
                timestamp: log.timestamp
            });
            sheet.eachRow((row) => {
                row.eachCell((cell) => {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                });
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