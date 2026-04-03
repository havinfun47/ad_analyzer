/* ============================================================
   Scale Science — PDF Export
   Uses html2pdf.js (loaded from CDN)
   ============================================================ */

function exportToPDF() {
  const client    = currentClient;
  const dates     = currentDates;
  const clientName = client?.name || "Client";
  const since     = dates?.since || "";
  const until     = dates?.until || "";

  const filename  = `${clientName.replace(/[^a-z0-9]/gi, "_")}_MetaAdsReport_${since}_${until}.pdf`;

  const element   = document.getElementById("report");
  if (!element) return;

  // Temporarily show print header
  const ph = document.getElementById("print-header");
  if (ph) ph.style.display = "flex";

  // Hide controls during capture
  const hideEls = document.querySelectorAll(".no-print");
  hideEls.forEach(el => el.style.visibility = "hidden");

  const opt = {
    margin:       [10, 10, 10, 10],
    filename:     filename,
    image:        { type: "jpeg", quality: 0.95 },
    html2canvas:  {
      scale:       2,
      useCORS:     true,
      logging:     false,
      backgroundColor: "#0A0A0A"
    },
    jsPDF:        {
      unit:        "mm",
      format:      "a4",
      orientation: "landscape"
    },
    pagebreak:    { mode: ["avoid-all", "css", "legacy"] }
  };

  if (typeof html2pdf === "undefined") {
    alert("PDF export library not loaded. Please check your internet connection.");
    if (ph) ph.style.display = "none";
    hideEls.forEach(el => el.style.visibility = "");
    return;
  }

  html2pdf()
    .set(opt)
    .from(element)
    .save()
    .finally(() => {
      if (ph) ph.style.display = "none";
      hideEls.forEach(el => el.style.visibility = "");
    });
}

function initPDFExport() {
  document.getElementById("btn-export")?.addEventListener("click", exportToPDF);
}

// Also support native print
function initPrint() {
  window.addEventListener("beforeprint", () => {
    const ph = document.getElementById("print-header");
    if (ph) ph.style.display = "flex";
  });
  window.addEventListener("afterprint", () => {
    const ph = document.getElementById("print-header");
    if (ph) ph.style.display = "none";
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initPDFExport();
  initPrint();
});
