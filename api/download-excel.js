const path = require("path");
const fs = require("fs");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ message: "Method Not Allowed" });
    return;
  }

  try {
    const filePath = path.join("/tmp", "student_data.xlsx");
    if (!fs.existsSync(filePath)) {
      res.status(404).json({ message: "Excel file not found yet" });
      return;
    }
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=student_data.xlsx");
    res.sendFile(filePath);
  } catch (error) {
    console.error("Error sending Excel file:", error);
    res.status(500).json({ message: "Unable to download the Excel file." });
  }
};


