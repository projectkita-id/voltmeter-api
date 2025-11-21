import cors from "cors";
import express from 'express';
import { PrismaClient } from "@prisma/client";

const app = express();
const prisma = new PrismaClient();

app.use(cors())
app.use(express.json());

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

app.post("/data", async (req, res) => {
    const { deviceId, cell, volt } = req.body;

    if (!deviceId || !cell || !Array.isArray(volt)) {
        return res.status(400).json({ error: "Invalid request body" });
    }

    if (volt.length !== cell) {
        return res.status(400).json({ error: "Volt array length does not match cell count" });
    }

    try {
        const result = [];

        for (let i = 0; i < cell; i++) {
            const log = await prisma.log.upsert({
                where: {
                    deviceId_cell: {
                        deviceId,
                        cell: String(i + 1)
                    }
                },
                update: {
                    volt: volt[i],
                    timestamp: new Date()
                },
                create: {
                    deviceId,
                    cell: String(i + 1),
                    volt: volt[i]
                }
            });

            result.push(log);
        }

        res.json({
            message: "Data logged successfully",
            count: result.length,
            data: result
        })
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        res.status(500).json({ error: "Internal server error: " + message });
    }

})

app.listen(3000, () => {
    console.log("Server is running on http://localhost:3000");
})