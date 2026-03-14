import cors from "cors";
import express from "express";
import ExcelJS from "exceljs";
import pkg from "@prisma/client";

const { PrismaClient } = pkg;

const app = express();
const prisma = new PrismaClient();

app.use(cors());
app.use(express.json());

/**
 * Lebar kolom FIX per blok untuk tabel utama (rapat/press isi)
 */
function setFixedBlockWidths(sheet, numBlocks, gap) {
  const WIDTH_CELL = 6;
  const WIDTH_VOLT = 9;
  const WIDTH_TIME = 12;
  const GAP_WIDTH = 2;

  for (let block = 0; block < numBlocks; block++) {
    const startCol = block * (3 + gap) + 1;

    sheet.getColumn(startCol + 0).width = WIDTH_CELL; // Cell
    sheet.getColumn(startCol + 1).width = WIDTH_VOLT; // Volt
    sheet.getColumn(startCol + 2).width = WIDTH_TIME; // Timestamp

    for (let g = 0; g < gap; g++) {
      sheet.getColumn(startCol + 3 + g).width = GAP_WIDTH;
    }
  }
}

app.get("/data", async (req, res) => {
  const log = await prisma.log.findMany();
  res.json(log);
});

app.get("/data/:deviceId", async (req, res) => {
  const { deviceId } = req.params;

  try {
    const record = await prisma.record.findMany({
      where: { deviceId },
      orderBy: { id: "asc" },
    });

    if (record.length === 0) {
      return res.json({ deviceId, records: [] });
    }

    const result = [];

    for (const rec of record) {
      const log = await prisma.log.findMany({
        where: { recordId: rec.id },
        orderBy: { cell: "asc" },
      });

      result.push({
        detail: rec.detail,
        total: rec.total,
        plus: rec.plus,
        minus: rec.minus,
        log,
      });
    }

    res.json({
      deviceId,
      data: result,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Internal server error: " + message });
  }
});

app.get("/data/:deviceId/export", async (req, res) => {
  const { deviceId } = req.params;

  try {
    const latest = await prisma.record.findFirst({
      where: { deviceId },
      orderBy: { timestamp: "desc" },
    });

    if (!latest) {
      return res.status(404).json({ error: "No records found for this unit" });
    }

    const logs = await prisma.log.findMany({
      where: { recordId: latest.id },
      orderBy: { cell: "asc" },
    });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(`Record ${latest.id}`);

    const blockSize = 31;
    const gap = 1;
    const colsPerBlock = 3;
    const headers = ["Cell", "Volt", "Timestamp"];
    const extra = ["Total Tegangan", "Positif - GND", "Negatif - GND"];

    const minBlocks = 3;
    const calculatedBlocks = Math.ceil(logs.length / blockSize);
    const numBlocks = Math.max(minBlocks, calculatedBlocks);
    const maxRows = 1 + blockSize;

    // Grid + header per blok
    for (let block = 0; block < numBlocks; block++) {
      const startCol = block * (colsPerBlock + gap) + 1;

      headers.forEach((h, idx) => {
        const cell = sheet.getCell(1, startCol + idx);
        cell.value = h;
        cell.font = { bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB0C4DE" } };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      for (let r = 2; r <= maxRows; r++) {
        for (let idx = 0; idx < colsPerBlock; idx++) {
          const cell = sheet.getCell(r, startCol + idx);
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }

        for (let g = 0; g < gap; g++) {
          const cell = sheet.getCell(r, startCol + colsPerBlock + g);
          cell.value = null;
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
    }

    // Isi logs
    logs.forEach((log, i) => {
      const block = Math.floor(i / blockSize);
      const posInBlock = i % blockSize;
      const startCol = block * (colsPerBlock + gap) + 1;
      const rowNum = posInBlock + 2;

      sheet.getCell(rowNum, startCol + 0).value = log.cell;

      const voltCell = sheet.getCell(rowNum, startCol + 1);
      voltCell.value = log.volt;

      const timeCell = sheet.getCell(rowNum, startCol + 2);
      timeCell.value = log.timestamp;
      timeCell.numFmt = "dd/mm/yyyy";
    });

    // ✅ Extra tetap di bawah kiri (A..C), tapi label dibuat lebih lebar dengan MERGE
    const extraStartCol = 1; // A
    const valueCol = 3; // C (nilai di kolom C)
    const labelMergeEndCol = 2; // B (label merge A-B). Kalau masih kepotong, ganti jadi 3 (A-C) dan nilai pindah ke D.

    // baris kosong pemisah
    for (let i = 0; i < colsPerBlock; i++) {
      sheet.getCell(maxRows + 1, extraStartCol + i).value = null;
    }

    extra.forEach((title, idx) => {
      const extraRow = maxRows + 2 + idx;

      // merge label A-B agar lebar
      sheet.mergeCells(extraRow, extraStartCol, extraRow, labelMergeEndCol);

      const titleCell = sheet.getCell(extraRow, extraStartCol);
      titleCell.value = title;
      titleCell.font = { bold: true, size: 12 };
      titleCell.alignment = {
        horizontal: "left",
        vertical: "middle",
        wrapText: false, // biar tidak turun baris
      };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB0C4DE" } };
      titleCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      const values = [latest.total, latest.plus, latest.minus];
      const valueCell = sheet.getCell(extraRow, valueCol);
      valueCell.value = values[idx];
      valueCell.font = { bold: true, size: 12 };
      valueCell.alignment = { horizontal: "center", vertical: "middle" };
      valueCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // ✅ Lebar tabel utama rapat (ini tetap)
    setFixedBlockWidths(sheet, numBlocks, gap);

    // Tambahan kecil: pastikan kolom value (C) cukup untuk angka 11,1 dll
    sheet.getColumn(3).width = Math.max(sheet.getColumn(3).width || 0, 10);

    res.setHeader("Content-Disposition", `attachment; filename=record_${deviceId}_${latest.id}.xlsx`);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(err);
    res.status(500).json({ error: "Failed to export data: " + message });
  }
});

app.post("/data", async (req, res) => {
  const { deviceId, standardCount, standard, all, plus, minus } = req.body;

  if (!deviceId || !standardCount || !Array.isArray(standard)) {
    return res.status(400).json({ error: "Invalid request body" });
  }

  if (standard.length !== standardCount) {
    return res.status(400).json({ error: "Volt array length does not match cell count" });
  }

  try {
    const previousCount = await prisma.record.count({ where: { deviceId } });

    const now = new Date();
    const desc = `Pengambilan ke-${previousCount + 1}`;

    const record = await prisma.record.create({
      data: {
        deviceId,
        cell: standardCount,
        detail: desc,
        timestamp: now,
        total: all,
        plus,
        minus,
      },
    });

    const logs = [];

    for (let i = 0; i < standardCount; i++) {
      const log = await prisma.log.create({
        data: {
          deviceId,
          cell: i + 1,
          volt: standard[i],
          timestamp: now,
          recordId: record.id,
        },
      });

      logs.push(log);
    }

    res.json({
      message: "Data recorded successfully",
      record,
      data: logs,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Internal server error: " + message });
  }
});

app.listen(3000, '0.0.0.0', () => {
  console.log("Server is running on http://localhost:3000");
});
