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
    const { name, email, course, latitude, longitude } = req.body;

   
    if (!name || !email || !course) {
      return res.status(400).json({ message: "Missing required fields" });
    }

   
    const geoRes = await fetch(
      `https://nominatim.openstreetmap.org/reverse?lat=${latitude}&lon=${longitude}&format=json`,
      {
        headers: {
          "User-Agent": "SchoolApp/1.0 (dinesh@example.com)", 
        },
      }
    );

    
    if (!geoRes.ok) {
      throw new Error("Failed to fetch address from OpenStreetMap");
    }

    const geoData = await geoRes.json();
    const address = geoData.display_name || "Address not found";

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
          { header: "Email", key: "email", width: 25 },
          { header: "Course", key: "course", width: 20 },
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
        { header: "Email", key: "email", width: 25 },
        { header: "Course", key: "course", width: 20 },
        { header: "Latitude", key: "latitude", width: 15 },
        { header: "Longitude", key: "longitude", width: 15 },
        { header: "Address", key: "address", width: 50 },
        { header: "Date/Time", key: "timestamp", width: 25 },
      ];
    }

   
    sheet.addRow({
      name,
      email,
      course,
      latitude,
      longitude,
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


app.listen(5000, () => console.log(" Server running on http://localhost:5000"));
