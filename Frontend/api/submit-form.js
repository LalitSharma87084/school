const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const nodemailer = require("nodemailer");

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

    try {
      const adminEmail = process.env.ADMIN_EMAIL;
      const gmailUser = process.env.GMAIL_USER;
      const gmailAppPassword = process.env.GMAIL_APP_PASSWORD;
      if (adminEmail && gmailUser && gmailAppPassword) {
        const transporter = nodemailer.createTransport({
          host: "smtp.gmail.com",
          port: 465,
          secure: true,
          auth: { user: gmailUser, pass: gmailAppPassword },
        });

        const buffer = await workbook.xlsx.writeBuffer();

        const html = `
          <div>
            <h3>New Student Submission</h3>
            <ul>
              <li><b>Name</b>: ${name}</li>
              <li><b>Adhaar No</b>: ${aadharNumber}</li>
              <li><b>Voter Card No</b>: ${voterCardNumber}</li>
              <li><b>Bank Account No</b>: ${bankAccountNumber}</li>
              <li><b>Bank Name</b>: ${bankName}</li>
              <li><b>Bank IFSC</b>: ${bankIfsc}</li>
              <li><b>Organization Name</b>: ${organizationName}</li>
              <li><b>Organization Address</b>: ${organizationAddress}</li>
              <li><b>Email</b>: ${email}</li>
              <li><b>Latitude</b>: ${latNum}</li>
              <li><b>Longitude</b>: ${lonNum}</li>
              <li><b>Address</b>: ${address}</li>
              <li><b>Timestamp</b>: ${new Date().toLocaleString()}</li>
            </ul>
          </div>
        `;

        await transporter.sendMail({
          from: gmailUser,
          to: adminEmail,
          subject: "New Form Submission",
          html,
          attachments: [
            { filename: "student_data.xlsx", content: Buffer.from(buffer) },
          ],
        });
      }
    } catch (e) {
      console.error("Email send error:", e);
    }

    // Persist a copy to Vercel Blob (private) so admin can view/download later
    try {
      const token = process.env.BLOB_READ_WRITE_TOKEN;
      if (token) {
        const { put } = await import("@vercel/blob");
        const key = `submissions/${Date.now()}-${Math.random()
          .toString(36)
          .slice(2)}.json`;
        const payload = {
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
          timestamp: new Date().toISOString(),
        };
        await put(key, JSON.stringify(payload), {
          access: "private",
          token,
          addRandomSuffix: false,
        });
      }
    } catch (e) {
      console.error("Blob store error:", e);
    }

    res.status(200).json({ message: "Data saved successfully!", address });
  } catch (error) {
    console.error("Serverless error:", error);
    res.status(500).json({ message: "Server error", error: error.message });
  }
};


