Office.onReady(() => {
  // Show a message if not running in Word
  if (!Office.context.requirements.isSetSupported("WordApi", "1.1")) {
    document.getElementById("hostNotice").style.display = "block";
    document.getElementById("hostNotice").innerText = "This add-in only works in Word.";
    return;
  }

  // Scan for clauses (mocked with content controls)
  document.getElementById("scan").onclick = async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      const clauses = [
        { title: "Limitation of Liability", tag: "clause_lol" },
        { title: "Indemnity", tag: "clause_indemnity" },
        { title: "Governing Law", tag: "clause_law" }
      ];

      clauses.forEach((clause) => {
        const range = body.insertParagraph(`🔍 ${clause.title}: [Insert clause text here]`, Word.InsertLocation.end);
        const cc = range.insertContentControl();
        cc.tag = clause.tag;
        cc.title = clause.title;
        cc.appearance = "BoundingBox";
      });

      await context.sync();
      logEvent("scan_complete", "Scanned and marked clauses");
    });
  };

  // Show "Why?" explanation
  document.getElementById("why").onclick = () => {
    const explanation = `This clause is flagged due to high risk exposure. Based on Policy-123 and past redlines.`;
    document.getElementById("whyOut").innerText = explanation;
    logEvent("why_shown", explanation);
  };

  // Apply suggestion to selected clause
  document.getElementById("apply").onclick = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText("✅ Suggested revision applied.", Word.InsertLocation.replace);
      await context.sync();
      logEvent("suggestion_applied", "User applied suggestion to selected clause");
    });
  };

  // Export audit log as CSV
  document.getElementById("export").onclick = () => {
    const rows = auditLog.map((e) => `${e.timestamp},${e.event},${e.details}`);
    const csv = ["timestamp,event,details", ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "audit_log.csv";
    link.click();
  };
});

// Simple audit log
const auditLog = [];

function logEvent(event, details) {
  const timestamp = new Date().toISOString();
  auditLog.push({ timestamp, event, details });

  const logDiv = document.getElementById("events");
  const entry = document.createElement("div");
  entry.className = "log-entry";
  entry.innerText = `${timestamp} — ${event}: ${details}`;
  logDiv.prepend(entry);
}
