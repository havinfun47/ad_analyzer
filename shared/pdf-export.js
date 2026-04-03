/* ============================================================
   Scale Science — PDF Export
   Uses html2pdf.js (loaded from CDN)
   Captures the live dark theme via onclone — no print-inversion.
   ============================================================ */

function exportToPDF() {
  if (typeof html2pdf === "undefined") {
    alert("PDF library not loaded. Check your internet connection.");
    return;
  }

  const client     = currentClient;
  const dates      = currentDates;
  const clientName = client?.name || "Client";
  const since      = dates?.since  || "";
  const until      = dates?.until  || "";
  const filename   = `${clientName.replace(/[^a-z0-9]/gi, "_")}_MetaAdsReport_${since}_${until}.pdf`;

  const element = document.getElementById("report");
  if (!element) return;

  const opt = {
    margin:      [6, 6, 6, 6],
    filename,
    image:       { type: "jpeg", quality: 0.97 },
    html2canvas: {
      scale:           2,
      useCORS:         true,
      logging:         false,
      backgroundColor: "#0A0A0A",
      windowWidth:     1300,
      onclone(clonedDoc) {
        const root = clonedDoc.documentElement;

        // ── Force dark CSS variables ──────────────────────────
        root.style.setProperty("--bg",           "#0A0A0A");
        root.style.setProperty("--bg-surface",   "#111111");
        root.style.setProperty("--bg-elevated",  "#181818");
        root.style.setProperty("--bg-card",      "#141414");
        root.style.setProperty("--border",       "#222222");
        root.style.setProperty("--border-subtle","#1A1A1A");
        root.style.setProperty("--text",         "#F0EDE8");
        root.style.setProperty("--text-muted",   "#666666");
        root.style.setProperty("--text-dim",     "#3A3A3A");
        root.style.setProperty("--accent",       "#00FF6A");
        root.style.setProperty("--accent-dim",   "rgba(0,255,106,0.12)");
        root.style.setProperty("--red",          "#FF3B3B");

        clonedDoc.body.style.background = "#0A0A0A";
        clonedDoc.body.style.color      = "#F0EDE8";

        // ── Show print header ─────────────────────────────────
        const ph = clonedDoc.getElementById("print-header");
        if (ph) {
          ph.style.display    = "flex";
          ph.style.background = "#0A0A0A";
          ph.style.color      = "#F0EDE8";
          ph.style.borderBottom = "1px solid #333";
          ph.style.paddingBottom = "20px";
          ph.style.marginBottom  = "24px";
        }

        // ── Hide UI chrome (display:none so no empty gaps) ───
        ["nav", "custom-date-row", "token-screen"].forEach(id => {
          const el = clonedDoc.getElementById(id);
          if (el) el.style.display = "none";
        });
        clonedDoc.querySelectorAll(".no-print").forEach(el => {
          el.style.display = "none";
        });

        // ── Ensure card / table backgrounds are explicit ─────
        clonedDoc.querySelectorAll(".kpi-card").forEach(el => {
          el.style.background   = "#141414";
          el.style.borderColor  = "#1A1A1A";
        });
        clonedDoc.querySelectorAll(".table-wrapper").forEach(el => {
          el.style.background  = "#111111";
          el.style.borderColor = "#222222";
        });
        clonedDoc.querySelectorAll(".editorial-block").forEach(el => {
          el.style.background  = "#111111";
          el.style.borderColor = "#222222";
        });
        clonedDoc.querySelectorAll(".data-table thead tr").forEach(el => {
          el.style.background = "#181818";
        });
        clonedDoc.querySelectorAll(".kpi-badge.positive").forEach(el => {
          el.style.background = "rgba(0,255,106,0.12)";
          el.style.color      = "#00FF6A";
        });
        clonedDoc.querySelectorAll(".kpi-badge.negative").forEach(el => {
          el.style.background = "rgba(255,59,59,0.12)";
          el.style.color      = "#FF3B3B";
        });
      }
    },
    jsPDF: {
      unit:        "mm",
      format:      "a4",
      orientation: "landscape"
    },
    pagebreak: { mode: ["css", "legacy"] }
  };

  html2pdf().set(opt).from(element).save();
}

function initPDFExport() {
  document.getElementById("btn-export")?.addEventListener("click", exportToPDF);
}

// Native browser print — keep @media print styles for Ctrl+P
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
