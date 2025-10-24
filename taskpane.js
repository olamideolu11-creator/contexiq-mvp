Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setupWord();
  } else if (info.host === Office.HostType.PowerPoint) {
    setupPowerPoint();
  }
});

function setupWord() {
  document.getElementById("scan").onclick = async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      const clauses = [
        { title: "Limitation of Liability", tag: "clause_lol" },
        { title: "Indemnity", tag: "clause_indemnity" },
        { title: "Governing Law", tag: "clause_law" }
      ];

      clauses.forEach((clause) => {
        const para = body.insertParagraph(`🔍 ${clause.title}: [Insert clause text here]`, Word.InsertLocation.end);
        const cc = para.insertContentControl();
        cc.tag = clause.tag;
        cc.title = clause.title;
        cc.appearance = "BoundingBox";
      });

      await context.sync();
      logEvent("scan_complete", "Scanned and marked clauses");
    });
  };

  document.getElementById("why").onclick = () => {
    const explanation = `This clause is flagged due to high risk exposure. Based on Policy-123 and past redlines.`;
    document.getElementById("whyOut").innerText = explanation;
    logEvent("why_shown", explanation);
  };

  document.getElementById("apply").onclick = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText("✅ Suggested revision applied.", Word.InsertLocation.replace);
      await context.sync();
      logEvent("suggestion_applied", "User applied suggestion to selected clause");
    });
  };
}

  function setupPowerPoint() {
  document.getElementById("insertSlide").onclick = () => {
    Office.context.document.setSelectedDataAsync(
      "🔍 Context IQ Suggestion:\nConsider revising clause 4.2 for clarity and compliance.",
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          logEvent("slide_inserted", "Inserted suggestion text into PowerPoint slide");
        } else {
          logEvent("error", "Failed to insert suggestion: " + asyncResult.error.message);
        }
      }
    );
  };
}

// Audit log
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
