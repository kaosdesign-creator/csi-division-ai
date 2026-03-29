import { useState, useCallback, useRef } from "react";

const CSI_DIVISIONS = {
  "00": "Procurement & Contracting Requirements",
  "01": "General Requirements",
  "02": "Existing Conditions",
  "03": "Concrete",
  "04": "Masonry",
  "05": "Metals",
  "06": "Wood, Plastics & Composites",
  "07": "Thermal & Moisture Protection",
  "08": "Openings",
  "09": "Finishes",
  "10": "Specialties",
  "11": "Equipment",
  "12": "Furnishings",
  "13": "Special Construction",
  "14": "Conveying Equipment",
  "21": "Fire Suppression",
  "22": "Plumbing",
  "23": "HVAC",
  "25": "Integrated Automation",
  "26": "Electrical",
  "27": "Communications",
  "28": "Electronic Safety & Security",
  "31": "Earthwork",
  "32": "Exterior Improvements",
  "33": "Utilities",
  "34": "Transportation",
  "35": "Waterway & Marine Construction",
  "40": "Process Integration",
  "41": "Material Processing & Handling",
  "48": "Electrical Power Generation",
};

const DIV_COLORS = {
  "00":"#64748B","01":"#64748B","02":"#92400E","03":"#6B7280",
  "04":"#B45309","05":"#0F766E","06":"#15803D","07":"#0369A1",
  "08":"#7C3AED","09":"#BE185D","10":"#DC2626","11":"#D97706",
  "12":"#0891B2","13":"#7C3AED","14":"#4F46E5","21":"#DC2626",
  "22":"#2563EB","23":"#EA580C","25":"#6D28D9","26":"#CA8A04",
  "27":"#0E7490","28":"#B91C1C","31":"#854D0E","32":"#166534",
  "33":"#1E40AF","34":"#374151","35":"#075985","40":"#581C87",
  "41":"#713F12","48":"#92400E",
};

const SYSTEM_PROMPT = `You are a senior construction estimator with deep expertise in CSI MasterFormat 2020.

Analyze this construction plan/drawing and identify EVERY element, note, callout, material, system, and specification visible. Leave nothing out.

Assign each item to the correct CSI MasterFormat 2020 Division (00-49).

Return ONLY valid JSON, no markdown, no preamble:
{
  "planDescription": "Brief description of plan type",
  "pageInfo": "Sheet number, title, scale, date if visible",
  "divisions": {
    "03": {
      "name": "Concrete",
      "items": [
        {"item": "4-inch concrete slab on grade", "detail": "6 mil vapor barrier, #4 rebar 18-inch OC EW", "location": "Floor area A"}
      ]
    }
  },
  "uncategorized": [],
  "totalItemCount": 0
}

Rules: Include EVERY visible element. General notes = Div 01. Existing/demo = Div 02. Only include divisions with actual items. Be exhaustive.`;

// Load PDF.js
function loadPdfJs() {
  return new Promise((resolve, reject) => {
    if (window.pdfjsLib) { resolve(window.pdfjsLib); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    s.onerror = () => reject(new Error("Failed to load PDF.js"));
    document.head.appendChild(s);
  });
}

async function renderPage(pdfData, pageNum, scale = 2.0) {
  const lib = await loadPdfJs();
  const pdf = await lib.getDocument({ data: pdfData }).promise;
  const page = await pdf.getPage(pageNum);
  const vp = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = vp.width; canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
  return { base64: canvas.toDataURL("image/jpeg", 0.9).split(",")[1], totalPages: pdf.numPages };
}

async function getPdfPageCount(pdfData) {
  const lib = await loadPdfJs();
  const pdf = await lib.getDocument({ data: pdfData }).promise;
  return pdf.numPages;
}

// ── Export helpers ──────────────────────────────────────────────
function exportToCSV(allResults, filename) {
  const rows = [["Division", "Division Name", "Item", "Specification / Detail", "Location", "Sheet"]];
  allResults.forEach(r => {
    Object.entries(r.divisions || {}).sort(([a],[b])=>parseInt(a)-parseInt(b)).forEach(([num, div]) => {
      (div.items || []).forEach(item => {
        rows.push([num, div.name || CSI_DIVISIONS[num] || "", item.item || "", item.detail || "", item.location || "", r.pageInfo || ""]);
      });
    });
  });
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  download(new Blob([csv], { type: "text/csv" }), filename + ".csv");
}

function exportToHTML(allResults, projectName) {
  // Merge all results into one division map
  const merged = {};
  allResults.forEach(r => {
    Object.entries(r.divisions || {}).forEach(([num, div]) => {
      if (!merged[num]) merged[num] = { name: div.name || CSI_DIVISIONS[num] || "", items: [] };
      (div.items || []).forEach(item => {
        merged[num].items.push({ ...item, sheet: r.pageInfo || "" });
      });
    });
  });

  const totalItems = Object.values(merged).reduce((s,d)=>s+(d.items?.length||0),0);
  const today = new Date().toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });

  const divRows = Object.entries(merged).sort(([a],[b])=>parseInt(a)-parseInt(b)).map(([num, div]) => {
    const color = DIV_COLORS[num] || "#6B7280";
    const itemRows = (div.items||[]).map(item => `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;font-size:12px;">${item.item||""}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#555;">${item.detail||""}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#777;">${item.location||""}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;color:#999;">${item.sheet||""}</td>
      </tr>`).join("");
    return `
    <div style="margin-bottom:20px;border:1px solid #e0e0e0;border-left:4px solid ${color};border-radius:4px;overflow:hidden;">
      <div style="background:#f8f8f8;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;">
        <div style="display:flex;align-items:center;gap:12px;">
          <span style="background:${color};color:#fff;font-weight:700;font-size:11px;padding:3px 8px;border-radius:3px;">DIV ${num}</span>
          <span style="font-weight:700;font-size:14px;">${div.name}</span>
        </div>
        <span style="font-size:12px;color:#666;">${div.items.length} items</span>
      </div>
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr style="background:#fafafa;">
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0;">ELEMENT</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0;">SPECIFICATION</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0;">LOCATION</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0;">SHEET</th>
        </tr></thead>
        <tbody>${itemRows}</tbody>
      </table>
    </div>`;
  }).join("");

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>${projectName} — CSI Analysis</title>
<style>body{font-family:Arial,sans-serif;margin:0;padding:0;color:#333;}
.cover{background:#1a1a2e;color:#fff;padding:60px 40px;}
.cover h1{margin:0 0 8px;font-size:28px;letter-spacing:1px;}
.cover p{margin:4px 0;opacity:0.7;font-size:13px;}
.stats{display:flex;gap:40px;margin-top:24px;}
.stat{text-align:center;}
.stat .num{font-size:32px;font-weight:900;color:#F59E0B;}
.stat .lbl{font-size:11px;opacity:0.6;letter-spacing:2px;}
.body{padding:32px 40px;}
@media print{.cover{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>
<div class="cover">
  <div style="font-size:11px;letter-spacing:3px;color:#F59E0B;margin-bottom:12px;">CSI MASTERFORMAT 2020 ANALYSIS</div>
  <h1>${projectName}</h1>
  <p>Generated ${today}</p>
  <div class="stats">
    <div class="stat"><div class="num">${Object.keys(merged).length}</div><div class="lbl">DIVISIONS</div></div>
    <div class="stat"><div class="num">${totalItems}</div><div class="lbl">TOTAL ITEMS</div></div>
    <div class="stat"><div class="num">${allResults.length}</div><div class="lbl">PAGES ANALYZED</div></div>
  </div>
</div>
<div class="body">${divRows}</div>
</body></html>`;

  download(new Blob([html], { type: "text/html" }), projectName.replace(/\s+/g,"_") + "_CSI_Report.html");
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); document.body.removeChild(a); }, 100);
}

// ── Main App ────────────────────────────────────────────────────
export default function App() {
  const [stage, setStage] = useState("upload"); // upload | select | analyzing | results
  const [pdfData, setPdfData] = useState(null);
  const [fileName, setFileName] = useState("");
  const [pageCount, setPageCount] = useState(0);
  const [pageThumbs, setPageThumbs] = useState({}); // pageNum -> base64
  const [selectedPages, setSelectedPages] = useState(new Set());
  const [loadingThumbs, setLoadingThumbs] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [progressCurrent, setProgressCurrent] = useState(0);
  const [progressTotal, setProgressTotal] = useState(0);
  const [progressMsg, setProgressMsg] = useState("");
  const [allResults, setAllResults] = useState([]);
  const [mergedDivisions, setMergedDivisions] = useState({});
  const [expandedDivs, setExpandedDivs] = useState({});
  const [projectName, setProjectName] = useState("");
  const [error, setError] = useState(null);
  const fileInputRef = useRef(null);
  const MAX_PAGES = 20;

  const processFile = async (f) => {
    if (!f.name.toLowerCase().endsWith(".pdf") && f.type !== "application/pdf") {
      setError("Please upload a PDF file."); return;
    }
    setError(null);
    setLoadingThumbs(true);
    setFileName(f.name);
    setProjectName(f.name.replace(/\.pdf$/i, "").replace(/_/g, " "));
    try {
      const buf = await f.arrayBuffer();
      const uint8 = new Uint8Array(buf);
      setPdfData(uint8);
      const count = await getPdfPageCount(uint8);
      setPageCount(count);
      // Pre-render first 6 thumbnails
      const thumbs = {};
      for (let i = 1; i <= Math.min(6, count); i++) {
        const { base64 } = await renderPage(uint8, i, 0.4);
        thumbs[i] = base64;
      }
      setPageThumbs(thumbs);
      setSelectedPages(new Set());
      setStage("select");
    } catch (e) { setError("Could not read PDF: " + e.message); }
    setLoadingThumbs(false);
  };

  const loadThumb = async (pageNum) => {
    if (pageThumbs[pageNum] || !pdfData) return;
    try {
      const { base64 } = await renderPage(pdfData, pageNum, 0.4);
      setPageThumbs(p => ({ ...p, [pageNum]: base64 }));
    } catch {}
  };

  const togglePage = (n) => {
    setSelectedPages(prev => {
      const next = new Set(prev);
      if (next.has(n)) next.delete(n);
      else if (next.size < MAX_PAGES) next.add(n);
      return next;
    });
  };

  const selectAll = () => {
    const pages = Array.from({ length: Math.min(pageCount, MAX_PAGES) }, (_, i) => i + 1);
    setSelectedPages(new Set(pages));
    // Load missing thumbs
    pages.forEach(p => loadThumb(p));
  };

  const clearAll = () => setSelectedPages(new Set());

  const runAnalysis = async () => {
    if (selectedPages.size === 0) return;
    setAnalyzing(true);
    setStage("analyzing");
    setError(null);
    const pages = Array.from(selectedPages).sort((a, b) => a - b);
    setProgressTotal(pages.length);
    setProgressCurrent(0);

    const results = [];
    for (let i = 0; i < pages.length; i++) {
      const pageNum = pages[i];
      setProgressMsg(`Analyzing page ${pageNum}...`);
      setProgressCurrent(i);
      try {
        const { base64 } = await renderPage(pdfData, pageNum, 2.0);
        const response = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 4000,
            system: SYSTEM_PROMPT,
            messages: [{
              role: "user",
              content: [
                { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
                { type: "text", text: `Analyze page ${pageNum} of this construction plan. Return only valid JSON.` },
              ],
            }],
          }),
        });
        if (!response.ok) {
          const e = await response.json();
          throw new Error(e.error?.message || "API error");
        }
        const data = await response.json();
        const text = data.content.map(b => b.text || "").join("");
        const match = text.match(/\{[\s\S]*\}/);
        if (match) results.push({ ...JSON.parse(match[0]), _page: pageNum });
      } catch (e) {
        results.push({ _page: pageNum, _error: e.message, divisions: {} });
      }
      setProgressCurrent(i + 1);
    }

    // Merge all into one division map
    const merged = {};
    results.forEach(r => {
      Object.entries(r.divisions || {}).forEach(([num, div]) => {
        if (!merged[num]) merged[num] = { name: div.name || CSI_DIVISIONS[num] || "", items: [] };
        (div.items || []).forEach(item => {
          merged[num].items.push({ ...item, _sheet: r.pageInfo || `Page ${r._page}` });
        });
      });
    });

    setAllResults(results);
    setMergedDivisions(merged);
    const exp = {};
    Object.keys(merged).forEach(k => (exp[k] = true));
    setExpandedDivs(exp);
    setProgressMsg("Complete!");
    setAnalyzing(false);
    setStage("results");
  };

  const totalItems = Object.values(mergedDivisions).reduce((s, d) => s + (d.items?.length || 0), 0);
  const errorPages = allResults.filter(r => r._error);

  // ── RENDER ──
  const font = "'Segoe UI','Helvetica Neue',Arial,sans-serif";

  // Upload stage
  if (stage === "upload") return (
    <div style={{ minHeight: "100vh", background: "#F8FAFC", fontFamily: font, display: "flex", flexDirection: "column" }}>
      {/* Header */}
      <div style={{ background: "#0F172A", padding: "16px 32px", display: "flex", alignItems: "center", gap: "12px" }}>
        <div style={{ background: "#F59E0B", color: "#000", fontWeight: "800", fontSize: "12px", padding: "4px 10px", borderRadius: "4px", letterSpacing: "1px" }}>CSI</div>
        <div style={{ color: "#fff", fontWeight: "700", fontSize: "16px" }}>Construction Plan Analyzer</div>
        <div style={{ color: "#64748B", fontSize: "12px", marginLeft: "4px" }}>MasterFormat 2020</div>
      </div>

      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "40px 20px" }}>
        <div style={{ maxWidth: "560px", width: "100%", textAlign: "center" }}>
          <div style={{ fontSize: "32px", marginBottom: "12px" }}>🏗️</div>
          <h1 style={{ margin: "0 0 8px", fontSize: "26px", fontWeight: "800", color: "#0F172A" }}>Analyze Your Construction Plans</h1>
          <p style={{ margin: "0 0 32px", color: "#64748B", fontSize: "14px" }}>Upload a PDF plan set. Select the pages you want analyzed. Get every element classified into CSI MasterFormat 2020 divisions — exportable to Excel, PDF, or Word.</p>

          <div
            onClick={() => fileInputRef.current.click()}
            onDrop={(e) => { e.preventDefault(); e.dataTransfer.files[0] && processFile(e.dataTransfer.files[0]); }}
            onDragOver={(e) => e.preventDefault()}
            style={{
              border: "2px dashed #CBD5E1", borderRadius: "12px", padding: "48px 32px",
              cursor: "pointer", background: "#fff", transition: "all 0.2s",
              marginBottom: "24px",
            }}
            onMouseEnter={e => { e.currentTarget.style.borderColor = "#F59E0B"; e.currentTarget.style.background = "#FFFBEB"; }}
            onMouseLeave={e => { e.currentTarget.style.borderColor = "#CBD5E1"; e.currentTarget.style.background = "#fff"; }}
          >
            <div style={{ fontSize: "40px", marginBottom: "12px" }}>📄</div>
            <div style={{ fontWeight: "700", color: "#0F172A", marginBottom: "6px" }}>Drop your PDF here</div>
            <div style={{ fontSize: "13px", color: "#94A3B8" }}>or click to browse · Up to {MAX_PAGES} pages per analysis</div>
            <input ref={fileInputRef} type="file" accept=".pdf,application/pdf" style={{ display: "none" }}
              onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
          </div>

          {loadingThumbs && <div style={{ color: "#64748B", fontSize: "13px" }}>⚙️ Loading PDF...</div>}
          {error && <div style={{ background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: "8px", padding: "12px 16px", color: "#DC2626", fontSize: "13px" }}>⚠ {error}</div>}

          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "12px", marginTop: "8px" }}>
            {[
              { icon: "📋", label: "Floor Plans", desc: "Layouts, schedules, callouts" },
              { icon: "📐", label: "Details & Sections", desc: "Assemblies, connections" },
              { icon: "⚡", label: "MEP Plans", desc: "Electrical, plumbing, HVAC" },
            ].map(c => (
              <div key={c.label} style={{ background: "#fff", border: "1px solid #E2E8F0", borderRadius: "8px", padding: "16px 12px", textAlign: "center" }}>
                <div style={{ fontSize: "22px", marginBottom: "6px" }}>{c.icon}</div>
                <div style={{ fontWeight: "700", fontSize: "12px", color: "#0F172A", marginBottom: "4px" }}>{c.label}</div>
                <div style={{ fontSize: "11px", color: "#94A3B8" }}>{c.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );

  // Page selection stage
  if (stage === "select") {
    const pages = Array.from({ length: pageCount }, (_, i) => i + 1);
    return (
      <div style={{ minHeight: "100vh", background: "#F8FAFC", fontFamily: font }}>
        <div style={{ background: "#0F172A", padding: "16px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ background: "#F59E0B", color: "#000", fontWeight: "800", fontSize: "12px", padding: "4px 10px", borderRadius: "4px", letterSpacing: "1px" }}>CSI</div>
            <div style={{ color: "#fff", fontWeight: "700", fontSize: "16px" }}>Select Pages to Analyze</div>
          </div>
          <button onClick={() => { setStage("upload"); setPdfData(null); setPageThumbs({}); }}
            style={{ background: "transparent", border: "1px solid #334155", color: "#94A3B8", padding: "6px 14px", borderRadius: "6px", fontSize: "12px", cursor: "pointer", fontFamily: font }}>
            ← New File
          </button>
        </div>

        <div style={{ maxWidth: "1100px", margin: "0 auto", padding: "24px 20px" }}>
          {/* Project name + controls */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "20px", flexWrap: "wrap", gap: "12px" }}>
            <div>
              <input value={projectName} onChange={e => setProjectName(e.target.value)}
                placeholder="Project Name"
                style={{ fontSize: "18px", fontWeight: "700", color: "#0F172A", border: "none", background: "transparent", outline: "none", fontFamily: font, padding: "0", minWidth: "300px" }} />
              <div style={{ fontSize: "12px", color: "#94A3B8", marginTop: "2px" }}>{fileName} · {pageCount} pages</div>
            </div>
            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
              <button onClick={selectAll} style={{ background: "#fff", border: "1px solid #CBD5E1", color: "#0F172A", padding: "8px 16px", borderRadius: "6px", fontSize: "12px", cursor: "pointer", fontFamily: font, fontWeight: "600" }}>
                Select All {pageCount > MAX_PAGES ? `(max ${MAX_PAGES})` : ""}
              </button>
              <button onClick={clearAll} style={{ background: "#fff", border: "1px solid #CBD5E1", color: "#64748B", padding: "8px 16px", borderRadius: "6px", fontSize: "12px", cursor: "pointer", fontFamily: font }}>
                Clear
              </button>
              <button
                onClick={runAnalysis}
                disabled={selectedPages.size === 0}
                style={{
                  background: selectedPages.size > 0 ? "#F59E0B" : "#E2E8F0",
                  color: selectedPages.size > 0 ? "#000" : "#94A3B8",
                  border: "none", padding: "8px 20px", borderRadius: "6px",
                  fontSize: "13px", fontWeight: "700", cursor: selectedPages.size > 0 ? "pointer" : "not-allowed",
                  fontFamily: font,
                }}>
                ⚡ Analyze {selectedPages.size > 0 ? `${selectedPages.size} Page${selectedPages.size > 1 ? "s" : ""}` : ""}
              </button>
            </div>
          </div>

          <div style={{ background: "#EFF6FF", border: "1px solid #BFDBFE", borderRadius: "8px", padding: "10px 16px", marginBottom: "20px", fontSize: "12px", color: "#1D4ED8" }}>
            💡 Click pages to select them for analysis. You can select up to {MAX_PAGES} pages at once. Select just the pages you need — plumbing sheets, electrical, structural, etc.
          </div>

          {/* Page grid */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: "12px" }}>
            {pages.map(n => {
              const selected = selectedPages.has(n);
              const thumb = pageThumbs[n];
              return (
                <div
                  key={n}
                  onClick={() => { togglePage(n); loadThumb(n); }}
                  style={{
                    border: selected ? "2px solid #F59E0B" : "2px solid #E2E8F0",
                    borderRadius: "8px", overflow: "hidden", cursor: "pointer",
                    background: selected ? "#FFFBEB" : "#fff",
                    transition: "all 0.15s", position: "relative",
                    boxShadow: selected ? "0 0 0 3px rgba(245,158,11,0.2)" : "none",
                  }}
                  onMouseEnter={e => { if(!selected) e.currentTarget.style.borderColor="#94A3B8"; }}
                  onMouseLeave={e => { if(!selected) e.currentTarget.style.borderColor="#E2E8F0"; }}
                >
                  {selected && (
                    <div style={{ position: "absolute", top: "6px", right: "6px", background: "#F59E0B", color: "#000", borderRadius: "50%", width: "20px", height: "20px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", fontWeight: "900", zIndex: 2 }}>✓</div>
                  )}
                  <div style={{ background: "#F1F5F9", height: "110px", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
                    {thumb
                      ? <img src={`data:image/jpeg;base64,${thumb}`} alt={`Page ${n}`} style={{ width: "100%", height: "100%", objectFit: "contain" }} />
                      : <div style={{ color: "#CBD5E1", fontSize: "11px" }} onClick={() => loadThumb(n)}>Page {n}</div>
                    }
                  </div>
                  <div style={{ padding: "6px 8px", fontSize: "11px", fontWeight: "600", color: selected ? "#92400E" : "#475569", textAlign: "center" }}>
                    Page {n}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  // Analyzing stage
  if (stage === "analyzing") {
    const pct = progressTotal > 0 ? Math.round((progressCurrent / progressTotal) * 100) : 0;
    return (
      <div style={{ minHeight: "100vh", background: "#F8FAFC", fontFamily: font, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "40px 20px" }}>
        <div style={{ background: "#0F172A", padding: "16px 32px", display: "flex", alignItems: "center", gap: "12px", position: "fixed", top: 0, left: 0, right: 0 }}>
          <div style={{ background: "#F59E0B", color: "#000", fontWeight: "800", fontSize: "12px", padding: "4px 10px", borderRadius: "4px" }}>CSI</div>
          <div style={{ color: "#fff", fontWeight: "700", fontSize: "16px" }}>Analyzing Plan...</div>
        </div>
        <div style={{ maxWidth: "480px", width: "100%", textAlign: "center", marginTop: "60px" }}>
          <div style={{ fontSize: "48px", marginBottom: "20px" }}>🔍</div>
          <h2 style={{ margin: "0 0 8px", color: "#0F172A", fontSize: "22px" }}>Classifying Elements</h2>
          <p style={{ color: "#64748B", fontSize: "14px", marginBottom: "32px" }}>{progressMsg}</p>

          <div style={{ background: "#E2E8F0", borderRadius: "99px", height: "10px", overflow: "hidden", marginBottom: "12px" }}>
            <div style={{ background: "#F59E0B", height: "100%", width: `${pct}%`, borderRadius: "99px", transition: "width 0.4s ease" }} />
          </div>
          <div style={{ fontSize: "13px", color: "#64748B" }}>
            <strong style={{ color: "#0F172A" }}>{progressCurrent}</strong> of <strong style={{ color: "#0F172A" }}>{progressTotal}</strong> pages complete · {pct}%
          </div>

          <div style={{ marginTop: "32px", background: "#fff", border: "1px solid #E2E8F0", borderRadius: "10px", padding: "20px 24px", textAlign: "left" }}>
            <div style={{ fontSize: "11px", color: "#94A3B8", letterSpacing: "1px", marginBottom: "10px" }}>WHAT'S HAPPENING</div>
            {["Rendering each page at high resolution", "Sending to Claude AI vision model", "Identifying every element on the plan", "Mapping to CSI MasterFormat 2020", "Merging results across all pages"].map((s, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: "10px", padding: "5px 0", fontSize: "13px", color: i < progressCurrent * 5 / progressTotal ? "#0F172A" : "#CBD5E1" }}>
                <span>{i < progressCurrent * 5 / progressTotal ? "✅" : "○"}</span> {s}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // Results stage
  if (stage === "results") {
    const divEntries = Object.entries(mergedDivisions).sort(([a],[b]) => parseInt(a)-parseInt(b));
    return (
      <div style={{ minHeight: "100vh", background: "#F8FAFC", fontFamily: font }}>
        {/* Header */}
        <div style={{ background: "#0F172A", padding: "14px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ background: "#F59E0B", color: "#000", fontWeight: "800", fontSize: "12px", padding: "4px 10px", borderRadius: "4px" }}>CSI</div>
            <div>
              <div style={{ color: "#fff", fontWeight: "700", fontSize: "15px" }}>{projectName || "Construction Plan Analysis"}</div>
              <div style={{ color: "#64748B", fontSize: "11px" }}>{allResults.length} pages analyzed · MasterFormat 2020</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: "8px" }}>
            <button onClick={() => exportToCSV(allResults, projectName || "CSI_Analysis")}
              style={{ background: "#166534", border: "none", color: "#fff", padding: "8px 14px", borderRadius: "6px", fontSize: "12px", fontWeight: "600", cursor: "pointer", fontFamily: font }}>
              📊 Excel / CSV
            </button>
            <button onClick={() => exportToHTML(allResults, projectName || "CSI Analysis")}
              style={{ background: "#1E40AF", border: "none", color: "#fff", padding: "8px 14px", borderRadius: "6px", fontSize: "12px", fontWeight: "600", cursor: "pointer", fontFamily: font }}>
              📄 PDF / Word Report
            </button>
            <button onClick={() => setStage("select")}
              style={{ background: "transparent", border: "1px solid #334155", color: "#94A3B8", padding: "8px 14px", borderRadius: "6px", fontSize: "12px", cursor: "pointer", fontFamily: font }}>
              ← Back
            </button>
            <button onClick={() => { setStage("upload"); setPdfData(null); setPageThumbs({}); setAllResults([]); setMergedDivisions({}); }}
              style={{ background: "transparent", border: "1px solid #334155", color: "#94A3B8", padding: "8px 14px", borderRadius: "6px", fontSize: "12px", cursor: "pointer", fontFamily: font }}>
              New File
            </button>
          </div>
        </div>

        <div style={{ maxWidth: "1200px", margin: "0 auto", padding: "24px 20px" }}>
          {/* Summary cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "14px", marginBottom: "24px" }}>
            {[
              { label: "Pages Analyzed", value: allResults.length, icon: "📄" },
              { label: "CSI Divisions", value: divEntries.length, icon: "📂" },
              { label: "Total Items", value: totalItems, icon: "📋" },
              { label: "Errors", value: errorPages.length, icon: errorPages.length > 0 ? "⚠️" : "✅" },
            ].map(c => (
              <div key={c.label} style={{ background: "#fff", border: "1px solid #E2E8F0", borderRadius: "10px", padding: "18px", textAlign: "center" }}>
                <div style={{ fontSize: "22px", marginBottom: "6px" }}>{c.icon}</div>
                <div style={{ fontSize: "26px", fontWeight: "800", color: "#0F172A" }}>{c.value}</div>
                <div style={{ fontSize: "11px", color: "#94A3B8", marginTop: "2px" }}>{c.label}</div>
              </div>
            ))}
          </div>

          {/* Export note */}
          <div style={{ background: "#F0FDF4", border: "1px solid #BBF7D0", borderRadius: "8px", padding: "10px 16px", marginBottom: "20px", fontSize: "12px", color: "#166534" }}>
            💡 <strong>Export tip:</strong> Click <strong>Excel / CSV</strong> to open in Excel and add cost columns. Click <strong>PDF / Word Report</strong> to download a formatted HTML report you can print to PDF or open in Word.
          </div>

          {/* Division cards */}
          {divEntries.map(([num, div]) => {
            const color = DIV_COLORS[num] || "#64748B";
            const isOpen = expandedDivs[num];
            return (
              <div key={num} style={{ background: "#fff", border: "1px solid #E2E8F0", borderLeft: `4px solid ${color}`, borderRadius: "8px", marginBottom: "8px", overflow: "hidden" }}>
                <div
                  onClick={() => setExpandedDivs(p => ({ ...p, [num]: !p[num] }))}
                  style={{ padding: "14px 18px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", userSelect: "none" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                    <div style={{ background: color, color: "#fff", fontWeight: "800", fontSize: "11px", padding: "3px 8px", borderRadius: "4px", minWidth: "36px", textAlign: "center" }}>
                      {num}
                    </div>
                    <div style={{ fontWeight: "700", fontSize: "14px", color: "#0F172A" }}>
                      {div.name}
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <div style={{ background: "#F1F5F9", color: "#475569", fontSize: "12px", fontWeight: "600", padding: "3px 10px", borderRadius: "99px" }}>
                      {div.items.length} items
                    </div>
                    <span style={{ color: "#CBD5E1", fontSize: "12px" }}>{isOpen ? "▲" : "▼"}</span>
                  </div>
                </div>
                {isOpen && (
                  <div style={{ borderTop: "1px solid #F1F5F9" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 180px 140px", padding: "6px 18px 6px 58px", background: "#F8FAFC", fontSize: "10px", color: "#94A3B8", fontWeight: "700", letterSpacing: "0.8px" }}>
                      <span>ELEMENT</span><span>SPECIFICATION / DETAIL</span><span>LOCATION</span><span>SHEET</span>
                    </div>
                    {div.items.map((item, idx) => (
                      <div key={idx} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 180px 140px", padding: "10px 18px 10px 58px", borderBottom: idx < div.items.length - 1 ? "1px solid #F8FAFC" : "none", alignItems: "start" }}>
                        <div style={{ fontSize: "13px", fontWeight: "600", color: "#0F172A", paddingRight: "12px" }}>{item.item}</div>
                        <div style={{ fontSize: "12px", color: "#475569", paddingRight: "12px" }}>{item.detail}</div>
                        <div style={{ fontSize: "12px", color: "#94A3B8", fontStyle: "italic" }}>{item.location}</div>
                        <div style={{ fontSize: "11px", color: "#CBD5E1" }}>{item._sheet}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}

          {/* Uncategorized */}
          {allResults.some(r => r.uncategorized?.length > 0) && (
            <div style={{ background: "#fff", border: "1px solid #FDE68A", borderLeft: "4px solid #F59E0B", borderRadius: "8px", padding: "16px 18px", marginTop: "8px" }}>
              <div style={{ fontWeight: "700", fontSize: "13px", color: "#92400E", marginBottom: "10px" }}>⚠ Items Needing Manual Review</div>
              {allResults.flatMap(r => (r.uncategorized || []).map(item => ({ item, sheet: r.pageInfo || `Page ${r._page}` }))).map((x, i) => (
                <div key={i} style={{ fontSize: "12px", color: "#78350F", padding: "3px 0" }}>• {x.item} <span style={{ color: "#D97706" }}>({x.sheet})</span></div>
              ))}
            </div>
          )}

          {/* Error pages */}
          {errorPages.length > 0 && (
            <div style={{ background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: "8px", padding: "14px 18px", marginTop: "8px" }}>
              <div style={{ fontWeight: "700", fontSize: "13px", color: "#DC2626", marginBottom: "8px" }}>Pages That Could Not Be Analyzed</div>
              {errorPages.map((r, i) => (
                <div key={i} style={{ fontSize: "12px", color: "#B91C1C", padding: "2px 0" }}>• Page {r._page}: {r._error}</div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  return null;
}
