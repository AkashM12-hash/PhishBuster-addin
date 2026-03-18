/* global Office */

const BACKEND_URL =
  window.location.hostname === "localhost"
    ? "http://localhost:8000"
    : "https://phishbuster-backend-z1a7.onrender.com";

Office.onReady(() => {
    analyzeEmail();
});

async function analyzeEmail() {

    const item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Text, async function (res) {

        const response = await fetch(`${BACKEND_URL}/analyze-outlook`, {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({
                subject: item.subject,
                body: res.value,
                senderEmail: item.from?.emailAddress?.address || ""
            })
        });

        const data = await response.json();

        updateUI(data);
    });
}

function updateUI(data) {

    const category = data.category;
    const riskScore = data.riskScore || 0;

    const badge = document.getElementById("statusBadge");
    const riskFill = document.getElementById("riskFill");
    const riskLabel = document.getElementById("riskLabel");
    const flowText = document.getElementById("flowText");
    const btn = document.getElementById("reportBtn");

    // ===== BADGE =====
    badge.innerText = category;

    if (category === "PHISHING") {
        badge.style.background = "#fdecea";
        badge.style.color = "#d93025";
    } else if (category === "SUSPICIOUS") {
        badge.style.background = "#fff3cd";
        badge.style.color = "#856404";
    } else {
        badge.style.background = "#e6f4ea";
        badge.style.color = "#188038";
    }

    // ===== RISK =====
    riskFill.style.width = riskScore + "%";

    if (riskScore > 70) {
        riskFill.style.background = "#d93025";
        riskLabel.innerText = "High Risk";
    } else if (riskScore > 40) {
        riskFill.style.background = "#ffc107";
        riskLabel.innerText = "Medium Risk";
    } else {
        riskFill.style.background = "#34a853";
        riskLabel.innerText = "Low Risk";
    }

    // ===== CONFIDENCE =====
    document.getElementById("confidenceText").innerText =
        "Confidence: " + (data.confidence || "N/A") + "%";

    // ===== INSIGHT =====
    document.getElementById("analysisMessage").innerText =
        data.analysisMessage || "";

    // ===== DETAILS =====
    document.getElementById("links").innerText =
        data.details?.links?.join(", ") || "None";

    document.getElementById("words").innerText =
        data.details?.suspiciousWords?.join(", ") || "None";

    // ===== FLOW =====
    if (category === "PHISHING") {
        flowText.innerText =
            "🚨 Automatically moved to Phishing folder and reported to Admin.";
    } else if (category === "SUSPICIOUS") {
        flowText.innerText =
            "⚠️ Moved to Suspicious folder for review.";
    } else {
        flowText.innerText =
            "✅ Email is safe and remains in inbox.";
    }

    // ===== BUTTON =====
    if (category === "PHISHING") {
        btn.style.display = "none";
    } else {
        btn.style.display = "block";
    }
}