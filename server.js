const express = require("express");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const ExcelJS = require("exceljs");

const app = express();
const PORT = process.env.PORT || 3000; // Use environment PORT if available

const DATA_FILE = path.join(__dirname, "submissions.json");
const BROCHURE_FILE = path.join(__dirname, "public/broucher.pdf"); // Inside /public/

// Middleware
app.use(express.json());
app.use(cors());

// Serve static files from /public/
app.use("/public", express.static(path.join(__dirname, "public")));

// Ensure `submissions.json` exists
if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify({}, null, 2));
}

// ✅ Registration API (now under `/api/register`)
app.post("/api/register", (req, res) => {
    const { name, email, phone, group, school, location } = req.body;

    // Validation
    if (!name || !email || !phone || !group || !school || !location) {
        return res.status(400).json({ message: "All fields are required" });
    }

    // Read existing data
    let submissions = {};
    if (fs.existsSync(DATA_FILE)) {
        try {
            submissions = JSON.parse(fs.readFileSync(DATA_FILE, "utf-8"));
        } catch (error) {
            console.error("Error reading JSON:", error);
        }
    }

    // Prevent duplicate phone numbers
    if (submissions[phone]) {
        return res.status(409).json({ message: "Already registered" });
    }

    // Save new submission
    submissions[phone] = { name, email, phone, group, school, location };
    fs.writeFileSync(DATA_FILE, JSON.stringify(submissions, null, 2));

    res.status(201).json({ message: "Registration successful", downloadUrl: "/api/download" });
});

// ✅ Auto-Download Brochure API (now under `/api/download`)
app.get("/api/download", (req, res) => {
    if (!fs.existsSync(BROCHURE_FILE)) {
        return res.status(404).json({ message: "Brochure not found" });
    }

    res.download(BROCHURE_FILE, "broucher.pdf");
});

// ✅ Download Submissions as Excel File (now under `/api/download-excel`)
app.get("/api/download-excel", async (req, res) => {
    // Read submissions data
    let submissions = {};
    if (fs.existsSync(DATA_FILE)) {
        try {
            submissions = JSON.parse(fs.readFileSync(DATA_FILE, "utf-8"));
        } catch (error) {
            console.error("Error reading JSON:", error);
            return res.status(500).json({ message: "Error reading submissions data" });
        }
    }

    // If no submissions exist, return a message
    if (Object.keys(submissions).length === 0) {
        return res.status(404).json({ message: "No submissions found" });
    }

    // Create a new Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Submissions");

    // Define columns
    worksheet.columns = [
        { header: "Name", key: "name", width: 20 },
        { header: "Email", key: "email", width: 30 },
        { header: "Phone", key: "phone", width: 15 },
        { header: "Group", key: "group", width: 15 },
        { header: "School", key: "school", width: 25 },
        { header: "Location", key: "location", width: 25 },
    ];

    // Add rows from submissions data
    Object.values(submissions).forEach((submission) => {
        worksheet.addRow(submission);
    });

    // Write the workbook to a buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Set headers for Excel file download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="submissions.xlsx"');

    // Send the buffer as the response
    res.send(buffer);
});

// ✅ Start the server
app.listen(PORT, () => {
    console.log(`🚀 API running at https://forms.faceprepdev.shop/api/`);
});
