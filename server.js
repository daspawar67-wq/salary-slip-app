const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const puppeteer = require("puppeteer");
const archiver = require("archiver");

const app = express();

// 🔥 IMPORTANT
app.use(express.urlencoded({ extended: true }));

app.use(express.static("public"));

const upload = multer({ dest: "uploads/" });

// ✅ Date Fix Function
function formatDate(value) {
  if (!value) return "";

  // Excel serial number
  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return date.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric"
    });
  }

  const date = new Date(value);
  if (!isNaN(date)) {
    return date.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric"
    });
  }

  return value;
}

app.post("/upload", upload.single("excel"), async (req, res) => {
  try {

    // 📅 Month
    const selectedMonth = req.body.month;

    const formattedMonth = selectedMonth
      ? new Date(selectedMonth).toLocaleString("en-IN", {
          month: "long",
          year: "numeric"
        })
      : new Date().toLocaleString("en-IN", {
          month: "long",
          year: "numeric"
        });

    // 📊 Read Excel
    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    const template = fs.readFileSync("template.html", "utf8");

    // 🧹 Clean old slips
    if (fs.existsSync("slips")) {
      fs.rmSync("slips", { recursive: true, force: true });
    }
    fs.mkdirSync("slips");

    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox"]
    });

    for (let emp of data) {

      let earnings = emp["Total Earnings"] || 0;
      let deductions = emp["Total Deductions"] || 0;
      let netSalary = emp["Net Salary"] || (earnings - deductions);

      let html = template

        // 👤 Employee
        .replace(/{{name}}/g, emp["Employee Name"] || "")
        .replace(/{{uan}}/g, emp["Employee ID"] || "")
        .replace(/{{designation}}/g, emp["Designation"] || "")
        .replace(/{{department}}/g, emp["Department"] || "")
        .replace(/{{doj}}/g, formatDate(emp["Date of Joining"]))
        .replace(/{{shift}}/g, "General")

        // 📅 Attendance
        .replace(/{{gross}}/g, emp["Gross Wages"] || 0)
        .replace(/{{leaves}}/g, emp["Leaves"] || 0)
        .replace(/{{totalDays}}/g, emp["Total Days in Month"] || 0)
        .replace(/{{weekoffs}}/g, emp["Week Offs"] || 0)
        .replace(/{{holidays}}/g, emp["Holidays"] || 0)
        .replace(/{{workingDays}}/g, emp["Working Days"] || 0)
        .replace(/{{present}}/g, emp["Present Days"] || 0)
        .replace(/{{absent}}/g, emp["Absent Days"] || 0)

        // 💰 Earnings
        .replace(/{{basic}}/g, emp["Basic"] || 0)
        .replace(/{{hra}}/g, emp["HRA"] || 0)
        .replace(/{{allowance}}/g, emp["Special Allowance"] || 0)
        .replace(/{{incentive}}/g, emp["Incentive"] || 0)
        .replace(/{{totalEarnings}}/g, earnings)

        // ➖ Deductions
        .replace(/{{epf}}/g, emp["EPF"] || 0)
        .replace(/{{esic}}/g, emp["ESIC"] || 0)
        .replace(/{{pt}}/g, emp["Professional Tax"] || 0)
        .replace(/{{attendanceDeduction}}/g, emp["Attendance Deductions"] || 0)
        .replace(/{{totalDeductions}}/g, deductions)

        // 🧾 Final
        .replace(/{{netSalary}}/g, netSalary)
        .replace(/{{month}}/g, formattedMonth);

      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: "networkidle0" });

      const fileName = (emp["Employee Name"] || "employee").replace(/\s+/g, "_");

      await page.pdf({
        path: `slips/${fileName}.pdf`,
        format: "A4",
        printBackground: true
      });

      await page.close();
    }

    await browser.close();

    // 📦 ZIP
    const archive = archiver("zip");
    const output = fs.createWriteStream("slips.zip");

    archive.pipe(output);
    archive.directory("slips/", false);
    await archive.finalize();

    output.on("close", () => {
      res.download("slips.zip");
    });

  } catch (err) {
    console.error(err);
    res.send("Error processing file");
  }
});

app.listen(3000, () => {
  console.log("Server running at http://localhost:3000");
});