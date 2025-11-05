const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
const app = express();
app.use(cors({ origin: "*", methods: ["GET", "POST"], allowedHeaders: ["Content-Type"] }));
app.use(express.json());

const filePath = path.join(__dirname, "student_data.xlsx");

app.post("/submit-form", async (req, res) => {
  try {
    const {
      name,
      aadharNumber,
      voterCardNumber,
      bankAccountNumber,
      bankName,
      bankIfsc,
      organizationName,
      organizationAddress,
      email,
      latitude,
      longitude,
    } = req.body;

   
    if (
      !name ||
      !aadharNumber ||
      !voterCardNumber ||
      !bankAccountNumber ||
      !bankName ||
      !bankIfsc ||
      !organizationName ||
      !organizationAddress ||
      !email
    ) {
      return res.status(400).json({ message: "Missing required fields" });
    }

   
    const latNum = Number(latitude);
    const lonNum = Number(longitude);

    if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) {
      return res.status(400).json({ message: "Location is required (valid latitude and longitude)." });
    }

    let address = "";
    const geoRes = await fetch(
      `https://nominatim.openstreetmap.org/reverse?lat=${latNum}&lon=${lonNum}&format=json`,
      {
        headers: {
          "User-Agent": "SchoolApp/1.0 (dinesh@example.com)",
        },
      }
    );

    if (geoRes.ok) {
      const geoData = await geoRes.json();
      address = geoData.display_name || "Address not found";
    } else {
      address = "Failed to fetch address";
    }

    console.log("✅ Received data:", req.body);
    console.log("📍 Address:", address);

    
    const workbook = new ExcelJS.Workbook();
    let sheet;

    try {
     
      await workbook.xlsx.readFile(filePath);
      sheet = workbook.getWorksheet("Students");

      
      if (!sheet) {
        sheet = workbook.addWorksheet("Students");
        sheet.columns = [
          { header: "Name", key: "name", width: 25 },
          { header: "Adhaar No", key: "aadharNumber", width: 22 },
          { header: "Voter Card No", key: "voterCardNumber", width: 22 },
          { header: "Bank Account No", key: "bankAccountNumber", width: 24 },
          { header: "Bank Name", key: "bankName", width: 22 },
          { header: "Bank IFSC", key: "bankIfsc", width: 18 },
          { header: "Organization Name", key: "organizationName", width: 26 },
          { header: "Organization Address", key: "organizationAddress", width: 36 },
          { header: "Email", key: "email", width: 28 },
          { header: "Latitude", key: "latitude", width: 15 },
          { header: "Longitude", key: "longitude", width: 15 },
          { header: "Address", key: "address", width: 50 },
          { header: "Date/Time", key: "timestamp", width: 25 },
        ];
      }
    } catch (err) {
      
      sheet = workbook.addWorksheet("Students");
      sheet.columns = [
        { header: "Name", key: "name", width: 25 },
        { header: "Adhaar No", key: "aadharNumber", width: 22 },
        { header: "Voter Card No", key: "voterCardNumber", width: 22 },
        { header: "Bank Account No", key: "bankAccountNumber", width: 24 },
        { header: "Bank Name", key: "bankName", width: 22 },
        { header: "Bank IFSC", key: "bankIfsc", width: 18 },
        { header: "Organization Name", key: "organizationName", width: 26 },
        { header: "Organization Address", key: "organizationAddress", width: 36 },
        { header: "Email", key: "email", width: 28 },
        { header: "Latitude", key: "latitude", width: 15 },
        { header: "Longitude", key: "longitude", width: 15 },
        { header: "Address", key: "address", width: 50 },
        { header: "Date/Time", key: "timestamp", width: 25 },
      ];
    }

   
    sheet.addRow({
      name,
      aadharNumber,
      voterCardNumber,
      bankAccountNumber,
      bankName,
      bankIfsc,
      organizationName,
      organizationAddress,
      email,
      latitude: latNum,
      longitude: lonNum,
      address,
      timestamp: new Date().toLocaleString(),
    });

    
    await workbook.xlsx.writeFile(filePath);

    
    res.status(200).json({
      message: "Data saved successfully!",
      address,
    });
  } catch (error) {
    console.error(" Server error:", error);
    res.status(500).json({ message: "Server error", error: error.message });
  }
});

app.get("/download-excel", async (req, res) => {
  try {
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=student_data.xlsx");
    res.sendFile(filePath);
  } catch (error) {
    console.error(" Error sending Excel file:", error);
    res.status(500).json({ message: "Unable to download the Excel file." });
  }
});


app.listen(5000, () => console.log(" Server running on http://localhost:5000"));
