module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ message: "Method Not Allowed" });
    return;
  }
  try {
    const { list, get } = await import("@vercel/blob");
    const { blobs } = await list({ prefix: "submissions/" });

    const items = [];
    for (const b of blobs) {
      try {
        const response = await get(b.url);
        const json = await response.json();
        items.push(json);
      } catch (e) {
        // skip broken entries
      }
    }
    // sort newest first
    items.sort((a, b) => (b.timestamp || "").localeCompare(a.timestamp || ""));
    res.status(200).json({ submissions: items });
  } catch (error) {
    console.error("Admin list error:", error);
    res.status(500).json({ message: "Server error" });
  }
};


