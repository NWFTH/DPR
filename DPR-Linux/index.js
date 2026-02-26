require("dotenv").config();
const sql = require("mssql");
const ExcelJS = require("exceljs");
const cron = require("node-cron");
const fs = require("fs");
const path = require("path");

const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_DATABASE,
  server: process.env.DB_SERVER,
  options: {
    instanceName: process.env.DB_INSTANCE,
    encrypt: false,
    trustServerCertificate: true,
  },
};

async function runRpa(daysToProcess = []) {
  let pool;
  try {
    const now = new Date();
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const currentMonth = monthNames[now.getMonth()];
    const currentYear = now.getFullYear();

    const fileName = `A Daily production report ${currentMonth} ${currentYear}.xlsx`;
    const filePath = path.join(__dirname, "ChicagoReport", fileName);

    if (!fs.existsSync(filePath)) throw new Error(`Template missing: ${fileName}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    pool = await sql.connect(config);

    if (daysToProcess.length === 0) {
      daysToProcess = [now.getDate()];
    }

    for (const day of daysToProcess) {
      const dateStr = `${currentYear}-${(now.getMonth() + 1).toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
      console.log(`[${new Date().toLocaleTimeString()}] Syncing: ${dateStr}`);

      const result = await pool.request().query(`
                SELECT LineType, CustKey, OrderNumber, OrderQty, TotalBags, ProductCode, 
                       StartTime, EndTime, PartialBagsKG, TotalStdUnitKG 
                FROM [dbo].[Chicago_Report_INF] 
                WHERE CAST(ReportDate AS DATE) = '${dateStr}'
                ORDER BY StartTime ASC
            `);

      const worksheet = workbook.worksheets.find(s => s.name.trim() === day.toString());
      if (!worksheet) continue;

      // Clear rows 5-500
      for (let i = 5; i <= 500; i++) {
        const row = worksheet.getRow(i);
        ["B", "C", "D", "E", "F", "H", "I", "J", "K", "L", "Y", "Z", "AA", "AB", "AC", "AE", "AG", "AF", "AH", "AI"].forEach(col => {
          row.getCell(col).value = null;
        });
      }

      let leftRow = 5;
      let rightRow = 5;

      result.recordset.forEach((row) => {
        let hour = 0;
        if (row.StartTime instanceof Date) hour = row.StartTime.getHours();
        
        const isDayShift = hour >= 7 && hour <= 18;
        const targetRow = isDayShift ? leftRow++ : rightRow++;
        const excelRow = worksheet.getRow(targetRow);

        const map = isDayShift 
          ? { B: "LineType", C: "CustKey", D: "OrderNumber", E: "OrderQty", F: "ProductCode", H: "StartTime", I: "EndTime", J: "TotalBags", K: "PartialBagsKG", L: "TotalStdUnitKG" }
          : { Y: "LineType", Z: "CustKey", AA: "OrderNumber", AB: "OrderQty", AC: "ProductCode", AE: "StartTime", AF: "EndTime", AG: "TotalBags", AH: "PartialBagsKG", AI: "TotalStdUnitKG" };

        Object.keys(map).forEach(col => {
          const val = row[map[col]];
          excelRow.getCell(col).value = val == null ? "" : val;
        });
      });
    }

    // Set Active Sheet
    const activeSheetIndex = workbook.worksheets.findIndex(s => s.name.trim() === now.getDate().toString());
    if (activeSheetIndex !== -1) {
      workbook.views = [{ activeTab: activeSheetIndex }];
    }

    // Save locally
    await workbook.xlsx.writeFile(filePath);

    // --- UPDATED EXPORT LOGIC ---
    const driveP = "P:\\Production\\DPR\\" + fileName;
    const uncPath = "\\\\th-bp-filesvr.nwfth.com\\Production\\DPR\\" + fileName;

    try {
      fs.copyFileSync(filePath, driveP);
      console.log("Auto-exported to P: Drive");
    } catch (e) {
      console.log("Drive P not found. Trying Network Path...");
      try {
        fs.copyFileSync(filePath, uncPath);
        console.log("Auto-exported via UNC Path");
      } catch (i) {
        console.error("Export failed: Check network permissions.");
      }
    }

    console.log(`Completed ${fileName}`);
  } catch (err) {
    console.error("RPA Error:", err.message);
  } finally {
    if (pool) await pool.close();
  }
}

async function start() {
  const now = new Date();
  const today = now.getDate();

  console.log(`Startup: Syncing month-to-date (Days 1 to ${today})...`);
  let syncDays = Array.from({ length: today }, (_, i) => i + 1);
  await runRpa(syncDays);

  console.log("Monitoring today's data every 5 minutes...");
  cron.schedule("*/5 * * * *", () => {
    runRpa([new Date().getDate()]);
  });
}

start();