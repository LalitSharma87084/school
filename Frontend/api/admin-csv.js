module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ message: "Method Not Allowed" });
    return;
  }
  try {
    const { list, get } = await import("@vercel/blob");
    const { blobs } = await list({ prefix: "submissions/" });

    const rows = [];
    for (const b of blobs) {
      try {
        const response = await get(b.url);
        const j = await response.json();
        rows.push(j);
      } catch (e) {}
    }
    rows.sort((a, b) => (b.timestamp || "").localeCompare(a.timestamp || ""));

    const headers = [
      "name","aadharNumber","voterCardNumber","bankAccountNumber","bankName","bankIfsc","organizationName","organizationAddress","email","latitude","longitude","address","timestamp"
    ];
    const escape = (v) => {
      if (v === undefined || v === null) return "";
      const s = String(v).replace(/"/g, '""');
      return `"${s}"`;
    };
    const csv = [headers.join(",")]
      .concat(rows.map(r => headers.map(h => escape(r[h])).join(",")))
      .join("\n");

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", "attachment; filename=submissions.csv");
    res.status(200).send(csv);
  } catch (error) {
    console.error("Admin csv error:", error);
    res.status(500).json({ message: "Server error" });
  }
};


