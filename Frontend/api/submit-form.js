const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ message: "Method Not Allowed" });
    return;
  }

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
    } = req.body || {};

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
      res.status(400).json({ message: "Missing required fields" });
      return;
    }

    const latNum = Number(latitude);
    const lonNum = Number(longitude);
    if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) {
      res.status(400).json({ message: "Location is required (valid latitude and longitude)." });
      return;
    }

    let address = "";
    try {
      const geoRes = await fetch(
        `https://nominatim.openstreetmap.org/reverse?lat=${latNum}&lon=${lonNum}&format=json`,
        { headers: { "User-Agent": "SchoolApp/1.0 (contact@example.com)" } }
      );
      if (geoRes.ok) {
        const geoData = await geoRes.json();
        address = geoData.display_name || "Address not found";
      } else {
        address = "Failed to fetch address";
      }
    } catch (e) {
      address = "Failed to fetch address";
    }

    const tmpDir = "/tmp";
    const filePath = path.join(tmpDir, "student_data.xlsx");

    const workbook = new ExcelJS.Workbook();
    let sheet;

    if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);
      sheet = workbook.getWorksheet("Students");
    }
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

    res.status(200).json({ message: "Data saved successfully!", address });
  } catch (error) {
    console.error("Serverless error:", error);
    res.status(500).json({ message: "Server error", error: error.message });
  }
};


