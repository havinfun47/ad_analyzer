/* ============================================================
   Scale Science — PDF Export
   Uses window.print() with dark @media print CSS.
   The browser's native print engine produces crisp, high-fidelity
   output — far better than html2canvas rasterisation.
   ============================================================ */

function exportToPDF() {
  const client     = currentClient;
  const dates      = currentDates;
  const clientName = (client?.name || "Client").replace(/[^a-z0-9]/gi, "_");
  const since      = dates?.since || "";
  const until      = dates?.until || "";

  // Set title so the browser uses it as the default PDF filename
  const origTitle  = document.title;
  document.title   = `${clientName}_MetaAdsReport_${since}_${until}`;

  window.print();

  // Restore after the print dialog closes (async — setTimeout is fine)
  setTimeout(() => { document.title = origTitle; }, 2000);
}

function initPDFExport() {
  document.getElementById("btn-export")?.addEventListener("click", exportToPDF);
}

// Show/hide print header around native browser print
window.addEventListener("beforeprint", () => {
  const ph = document.getElementById("print-header");
  if (ph) ph.style.display = "flex";
});

window.addEventListener("afterprint", () => {
  const ph = document.getElementById("print-header");
  if (ph) ph.style.display = "none";
});

document.addEventListener("DOMContentLoaded", initPDFExport);
