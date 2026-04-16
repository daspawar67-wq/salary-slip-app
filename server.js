const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const puppeteer = require("puppeteer");
const archiver = require("archiver");
const path = require("path");

const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

/* =============================
   ✅ RENDER FILE SYSTEM FIX
   ============================= */

const UPLOAD_DIR = "/tmp/uploads";
const SLIPS_DIR = "/tmp/slips";

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);
if (!fs.existsSync(SLIPS_DIR)) fs.mkdirSync(SLIPS_DIR);

const upload = multer({ dest: UPLOAD_DIR });

/* =============================
   DATE FORMAT
   ============================= */

function formatDate(value) {
  if (!value) return "";

  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return date.toLocaleDateString("en-IN");
  }

  const date = new Date(value);
  if (!isNaN(date)) return date.toLocaleDateString("en-IN");

  return value;
}

/* =============================
   UPLOAD ROUTE
   ============================= */

app.post("/upload", upload.single("excel"), async (req, res) => {
  try {

    const selectedMonth = req.body.month;

    const formattedMonth = selectedMonth
      ? new Date(selectedMonth).toLocaleString("en-IN", {
          month: "long",
          year: "numeric",
        })
      : new Date().toLocaleString("en-IN", {
          month: "long",
          year: "numeric",
        });

    /* Excel Read */
    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    const template = fs.readFileSync("template.html", "utf8");

    /* Clean old slips */
    fs.rmSync(SLIPS_DIR, { recursive: true, force: true });
    fs.mkdirSync(SLIPS_DIR);

    /* =============================
       ✅ PUPPETEER FIX FOR RENDER
       ============================= */

const browser = await puppeteer.launch({
  headless: true,
  args: ["--no-sandbox", "--disable-setuid-sandbox"]
});

    for (let emp of data) {

      let earnings = emp["Total Earnings"] || 0;
      let deductions = emp["Total Deductions"] || 0;
      let netSalary = emp["Net Salary"] || (earnings - deductions);

      let html = template
        .replace(/{{name}}/g, emp["Employee Name"] || "")
        .replace(/{{uan}}/g, emp["Employee ID"] || "")
        .replace(/{{designation}}/g, emp["Designation"] || "")
        .replace(/{{department}}/g, emp["Department"] || "")
        .replace(/{{doj}}/g, formatDate(emp["Date of Joining"]))
        .replace(/{{shift}}/g, "General")
        .replace(/{{gross}}/g, emp["Gross Wages"] || 0)
        .replace(/{{leaves}}/g, emp["Leaves"] || 0)
        .replace(/{{totalDays}}/g, emp["Total Days in Month"] || 0)
        .replace(/{{weekoffs}}/g, emp["Week Offs"] || 0)
        .replace(/{{holidays}}/g, emp["Holidays"] || 0)
        .replace(/{{workingDays}}/g, emp["Working Days"] || 0)
        .replace(/{{present}}/g, emp["Present Days"] || 0)
        .replace(/{{absent}}/g, emp["Absent Days"] || 0)
        .replace(/{{basic}}/g, emp["Basic"] || 0)
        .replace(/{{hra}}/g, emp["HRA"] || 0)
        .replace(/{{allowance}}/g, emp["Special Allowance"] || 0)
        .replace(/{{incentive}}/g, emp["Incentive"] || 0)
        .replace(/{{totalEarnings}}/g, earnings)
        .replace(/{{epf}}/g, emp["EPF"] || 0)
        .replace(/{{esic}}/g, emp["ESIC"] || 0)
        .replace(/{{pt}}/g, emp["Professional Tax"] || 0)
        .replace(/{{attendanceDeduction}}/g, emp["Attendance Deductions"] || 0)
        .replace(/{{totalDeductions}}/g, deductions)
        .replace(/{{netSalary}}/g, netSalary)
        .replace(/{{month}}/g, formattedMonth);

      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: "networkidle0" });

      const fileName = (emp["Employee Name"] || "employee")
        .replace(/\s+/g, "_");

      await page.pdf({
        path: path.join(SLIPS_DIR, `${fileName}.pdf`),
        format: "A4",
        printBackground: true,
      });

      await page.close();
    }

    await browser.close();

    /* ZIP FILE */
    const zipPath = "/tmp/slips.zip";
    const archive = archiver("zip");
    const output = fs.createWriteStream(zipPath);

    archive.pipe(output);
    archive.directory(SLIPS_DIR, false);
    await archive.finalize();

    output.on("close", () => {
      res.download(zipPath);
    });

  } catch (err) {
    console.error(err);
    res.send("Error processing file");
  }
});

/* =============================
   ✅ RENDER PORT FIX
   ============================= */

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("Server running on port " + PORT);
});
