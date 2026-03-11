/* global Office */

// 🔥 BACKEND CONFIGURATION
// CHANGE THIS URL AFTER DEPLOYING TO RENDER
// const BACKEND_URL = "https://YOUR-APP-NAME.onrender.com"; 
// Example: "https://phishing-detector-abc123.onrender.com"
// Get this URL after Render deployment (Step 5 in deployment guide)
window.autoReportedMessageIds = new Set();

const BACKEND_URL =
  window.location.hostname === "localhost"
    ? "http://localhost:8000"
    : "https://phishbuster-backend-z1a7.onrender.com";





Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "block";

  // INITIAL UI STATE
  document.getElementById("resultsSection").style.display = "none";
  document.getElementById("scanSummary").style.display = "none";
  document.getElementById("appDescription").style.display = "block";
  document.getElementById("reportButton").style.display = "none";

  // 🔥 NEW: Check Auto-Scan Status on Load
  checkAutoScanStatus();

  // Button Event Listeners
  document.getElementById("run").onclick = analyzeCurrentEmail;
  // document.getElementById("scanUnread").onclick = scanUnreadEmails;
  document.getElementById("manageAutoScan").onclick = openAutoScanSettings;
});

// ==========================================================
// 🔥 NEW: CHECK AUTO-SCAN STATUS
// ==========================================================

function checkAutoScanStatus() {
  // For "Better UX" approach, we assume Power Automate flow is enabled
  // In production, you could call your backend to check if flow exists
  
  const statusBadge = document.getElementById("autoScanBadge");
  const statusText = document.getElementById("autoScanStatus");
  
  // Simulate checking status (in real implementation, call backend)
  setTimeout(() => {
    // Assume flow is enabled (user manages this in Power Automate)
    statusBadge.classList.add("status-active");
    statusText.innerText = "Active";
  }, 500);
}

// ==========================================================
// 🔥 NEW: OPEN POWER AUTOMATE SETTINGS
// ==========================================================

function openAutoScanSettings() {
  alert("Auto-scan settings are managed by the IT administrator.");
}


// ==========================================================
// 🔥 NEW: SCAN ALL UNREAD EMAILS
// ==========================================================

async function scanUnreadEmails() {
  try {
    document.getElementById("resultsSection").style.display = "none";
    document.getElementById("scanSummary").style.display = "block";
    document.getElementById("appDescription").style.display = "none";
    
    // Reset counters
    document.getElementById("safeCount").innerText = "...";
    document.getElementById("suspiciousCount").innerText = "...";
    document.getElementById("phishingCount").innerText = "...";
    document.getElementById("phishingList").innerHTML = "<p>🔍 Scanning unread emails...</p>";

    // Get mailbox
    const mailbox = Office.context.mailbox;
    
    // Get unread emails from inbox
    mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
      if (result.status === "succeeded") {
        const token = result.value;
        
        // Get REST API URL
        const restUrl = mailbox.restUrl;
        const filterQuery = "$filter=IsRead eq false&$top=20&$select=Subject,From,BodyPreview";
        const url = `${restUrl}/v2.0/me/messages?${filterQuery}`;
        
        // Fetch unread emails
        fetch(url, {
          method: "GET",
          headers: {
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json"
          }
        })
        .then(response => response.json())
        .then(data => {
          if (data.value && data.value.length > 0) {
            scanEmailsBatch(data.value);
          } else {
            document.getElementById("phishingList").innerHTML = 
              "<p>✅ No unread emails found.</p>";
          }
        })
        .catch(err => {
          console.error("Error fetching emails:", err);
          document.getElementById("phishingList").innerHTML = 
            "<p>❌ Could not fetch unread emails. Please try again.</p>";
        });
      }
    });
    
  } catch (err) {
    console.error("Error in scanUnreadEmails:", err);
    console.error("Auto scan failed.");

  }
}

// ==========================================================
// SCAN MULTIPLE EMAILS
// ==========================================================

async function scanEmailsBatch(emails) {
  let safeCount = 0;
  let suspiciousCount = 0;
  let phishingCount = 0;
  let phishingEmails = [];

  document.getElementById("phishingList").innerHTML = "<p>⏳ Analyzing emails...</p>";

  // Scan each email
  for (const email of emails) {
    try {
      const response = await fetch(`${BACKEND_URL}/analyze`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sender: email.From.EmailAddress.Address,
          subject: email.Subject,
          body: email.BodyPreview
        })
      });

      const data = await response.json();
      
      if (data.result === true) {
        phishingCount++;
        phishingEmails.push({
          subject: email.Subject,
          sender: email.From.EmailAddress.Address
        });
      } else if (data.result === "Suspicious") {
        suspiciousCount++;
      } else {
        safeCount++;
      }
      
    } catch (err) {
      console.error("Error analyzing email:", err);
    }
  }

  // Update UI
  document.getElementById("safeCount").innerText = safeCount;
  document.getElementById("suspiciousCount").innerText = suspiciousCount;
  document.getElementById("phishingCount").innerText = phishingCount;

  // Show phishing emails list
  if (phishingEmails.length > 0) {
    let listHtml = "<div class='phishing-list'><h4>🚨 Phishing Emails Detected:</h4>";
    phishingEmails.forEach(email => {
      listHtml += `
        <div class='phishing-item'>
          <strong>Subject:</strong> ${email.subject}<br>
          <strong>From:</strong> ${email.sender}
        </div>
      `;
    });
    listHtml += "</div>";
    document.getElementById("phishingList").innerHTML = listHtml;
  } else {
    document.getElementById("phishingList").innerHTML = 
      "<p style='color: #28a745; margin-top: 15px;'>✅ No phishing emails detected!</p>";
  }
}

// ==========================================================
// ANALYZE CURRENT EMAIL (ORIGINAL FUNCTION - KEPT)
// ==========================================================

function analyzeCurrentEmail() {
  try {
    const item = Office.context.mailbox.item;

    // Hide scan summary, show results section
    document.getElementById("scanSummary").style.display = "none";
    document.getElementById("resultsSection").style.display = "block";
    document.getElementById("appDescription").style.display = "none";

    // RESET UI
    document.getElementById("statusText").innerText = "Analyzing...";
    document.getElementById("confidenceText").innerText = "";
    document.getElementById("analysisMessage").innerText = "";
    document.getElementById("links").innerText = "";
    document.getElementById("suspiciousWords").innerText = "";
    document.getElementById("riskLevel").innerText = "";

    const statusBox = document.getElementById("statusBox");
    statusBox.style.backgroundColor = "#f3f2f1";
    statusBox.style.border = "1px solid #ccc";

    let subject = item.subject || "";

    item.body.getAsync(Office.CoercionType.Text, async function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        document.getElementById("statusText").innerText =
          "Failed to read email body.";
        return;
      }

      const body = result.value;
      const sender =
        item.from?.emailAddress?.address || "";

      try {
        const response = await fetch(`${BACKEND_URL}/analyze-outlook`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            senderName: item.from?.displayName || "",
            senderEmail: sender,
            userEmail: Office.context.mailbox.userProfile.emailAddress || "",
            subject: subject,
            body: body,
            messageId: item.itemId || ""
          })
        });

        if (!response.ok) {
          document.getElementById("statusText").innerText =
            "Backend error: " + response.status;
          return;
        }

        const data = await response.json();
        const riskScore = data.riskScore || 0;
        const riskLevel = data.riskLevel || "LOW";

        const meterBar = document.getElementById("riskMeterBar");
        const meterLabel = document.getElementById("riskMeterLabel");

        meterBar.style.width = riskScore + "%";

        if (riskLevel === "HIGH") {
          meterBar.style.background = "#E53935";
        } else if (riskLevel === "MEDIUM") {
          meterBar.style.background = "#FFC107";
        } else {
          meterBar.style.background = "#1DB954";
        }

        meterLabel.innerText = riskScore + "% — " + riskLevel + " RISK";
        // ================= INTERNAL / EXTERNAL BADGE =================
//         const originBadge = document.getElementById("originBadge");
//         const originLabel = document.getElementById("originLabel");

//         if (typeof data.isInternal !== "undefined") {
//           originBadge.style.display = "block";

//           if (data.isInternal) {
//             originLabel.innerText = "INTERNAL";
//             originLabel.className = "origin-badge origin-internal";
//           } else {
//             originLabel.innerText = "EXTERNAL";
//             originLabel.className = "origin-badge origin-external";
//           }
//         } else {
//           originBadge.style.display = "none";
// }

                // Set base analysis message first
        document.getElementById("analysisMessage").innerText =
          data.analysisMessage || data.aiExplanation || "No detailed analysis available.";

        // STORE LAST ANALYSIS FOR REPORTING
        // STORE LAST ANALYSIS FOR REPORTING
        window.lastAnalysis = {
          category: data.category,
          sender: sender,
          subject: subject,
          reason: data.reason,
          messageId: item.itemId || "",
          confidence: data.confidence || null
        };
        
        // Store current email info for manual reporting
        window.currentMessageId = item.itemId || "";
        window.currentSender = sender || "";

        window.currentUserEmail = Office.context.mailbox.userProfile.emailAddress || "";
        window.currentSubject = subject || "";

        // ================= FIXED LOGIC =================

        // ================= CATEGORY-BASED LOGIC =================

        const category = data.category; // SAFE | SUSPICIOUS | PHISHING
        const reportBtn = document.getElementById("reportButton");

        // Show button only for PHISHING or SUSPICIOUS
        if (category === "PHISHING" || category === "SUSPICIOUS") {
          reportBtn.style.display = "block";
        } else {
          reportBtn.style.display = "none";
        }

        const confidence = data.confidence ?? null;
        

        // ================= UI UPDATE =================

        const statusText = document.getElementById("statusText");
        const confidenceText = document.getElementById("confidenceText");

        if (category === "PHISHING") {
          statusBox.style.backgroundColor = "#fdecea";
          statusBox.style.border = "2px solid #d93025";
          statusText.innerText = "🚨 PHISHING DETECTED";
          statusText.style.color = "#d93025";
        
          confidenceText.innerText = "Confidence: " + confidence + "%";
          const msgId = item.itemId || "";
          if (!window.autoReportedMessageIds.has(msgId)) {
            window.autoReportedMessageIds.add(msgId);
            
            document.getElementById("analysisMessage").innerText += 
            "\n\n🚨 This phishing email was automatically reported to Admin.";
          }
        
        } else if (category === "SUSPICIOUS") {
          statusBox.style.backgroundColor = "#fff3cd";
          statusBox.style.border = "2px solid #ffc107";
          statusText.innerText = "⚠️ SUSPICIOUS EMAIL";
          statusText.style.color = "#856404";
        
          confidenceText.innerText = "";
        
        } else {
          statusBox.style.backgroundColor = "#e6f4ea";
          statusBox.style.border = "2px solid #188038";
          statusText.innerText = "✅ SAFE EMAIL";
          statusText.style.color = "#188038";
        
          confidenceText.innerText = "";
        }
        

        // REAL DETAILS
        const links = data.details?.links || [];
        const words = data.details?.suspiciousWords || [];

        document.getElementById("links").innerText =
          links.length > 0 ? links.join(", ") : "None";

        document.getElementById("suspiciousWords").innerText =
          words.length > 0 ? words.join(", ") : "None";

        document.getElementById("riskLevel").innerText = category;


      

      } catch (apiErr) {
        console.error("Frontend error after API call:", apiErr);
        document.getElementById("statusText").innerText =
          "Error while processing response. Check console.";
      }
    });

  } catch (err) {
    console.error(err);
    document.getElementById("statusText").innerText =
      "Unexpected error.";
  }
}
document.getElementById("reportButton").onclick = async function () {
  console.log("🚀 Report button clicked, sending to backend...");
  if (!window.lastAnalysis) {
    showReportSuccess("No analysis available to report.");
    return;
  }

  // 🚨 Only allow reporting for PHISHING / SUSPICIOUS
  if (!["PHISHING", "SUSPICIOUS"].includes(window.lastAnalysis.category)) {
    showReportSuccess("Only suspicious or phishing emails can be reported.");
    return;
  }

  const reportPayload = {
    messageId: window.currentMessageId || "",
    category: window.lastAnalysis.category,
    confidence: window.lastAnalysis.confidence || null,
    ruleHits: window.lastAnalysis.reason
      ? window.lastAnalysis.reason.split(";")
      : [],
    sender: window.currentSender || "",
    reportedBy: window.currentUserEmail || ""
  };
 


  try {
    const response = await fetch(`${BACKEND_URL}/report-to-admin`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        messageId: reportPayload.messageId,
        category: reportPayload.category,
        confidence: reportPayload.confidence,
        ruleHits: reportPayload.ruleHits,
        sender: reportPayload.sender,
        reportedBy: reportPayload.reportedBy
      })
    });

    if (response.ok) {
      showReportSuccess("Reported to IT Admin for review.");
    } else {
      console.error("Failed to report to admin.");
    }
    

  } catch (err) {
    console.error(err);
    showReportSuccess("Failed to report email.");
  }
};

function showReportSuccess(message) {
  const msgBox = document.getElementById("analysisMessage");
  msgBox.innerText = "✅ " + (message || "Reported successfully");
  msgBox.style.color = "green";
}



