function safeFilePart(text) {
  return String(text || "")
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "") || "문서";
}

function exportFileName(docTitle, personName, ext) {
  const d = new Date();
  const ymd = String(d.getFullYear()) +
    String(d.getMonth() + 1).padStart(2, "0") +
    String(d.getDate()).padStart(2, "0");
  return safeFilePart(docTitle) + "_" + safeFilePart(personName) + "_" + ymd + "." + (ext || "pdf");
}

function loadScript(src) {
  return new Promise(function (resolve, reject) {
    const s = document.createElement("script");
    s.src = src;
    s.onload = function () { resolve(); };
    s.onerror = function () { reject(new Error("라이브러리 로드 실패")); };
    document.head.appendChild(s);
  });
}

async function ensurePdfLib() {
  if (typeof html2pdf === "function") return;
  try {
    await loadScript("html2pdf.bundle.min.js");
  } catch (err) {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js");
  }
  if (typeof html2pdf !== "function") {
    throw new Error("PDF 변환 라이브러리를 불러오지 못했습니다.");
  }
}

function looksLikeHtml(bytes, mime) {
  if (looksLikePdf(bytes) || looksLikePng(bytes)) return false;
  if (String(mime || "").toLowerCase().indexOf("html") >= 0) return true;
  const head = Array.from(bytes.slice(0, 48)).map(function (b) {
    return String.fromCharCode(b);
  }).join("").toLowerCase().replace(/^\uFEFF/, "").trim();
  return head.indexOf("<!doctype html") === 0 || head.indexOf("<html") === 0;
}

function looksLikePdf(bytes) {
  return bytes.length >= 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46;
}

function looksLikePng(bytes) {
  return bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
}

async function blobToDataUri(blob) {
  return new Promise(function (resolve, reject) {
    const reader = new FileReader();
    reader.onload = function () { resolve(String(reader.result || "")); };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function downloadBlob(blob, filename) {
  const lower = String(filename || "").toLowerCase();
  if (lower.endsWith(".html") || lower.endsWith(".htm")) {
    throw new Error("HTML 파일은 저장하지 않습니다.");
  }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  if (looksLikeHtml(bytes, blob.type)) {
    throw new Error("HTML 파일은 저장하지 않습니다.");
  }
  let mime = "application/pdf";
  let name = filename;
  if (looksLikePng(bytes)) {
    mime = "image/png";
    name = filename.replace(/\.pdf$/i, ".png");
  } else if (!looksLikePdf(bytes)) {
    throw new Error("PDF/이미지 변환에 실패했습니다.");
  }
  const typed = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(typed);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.rel = "noopener";
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  return { blob: typed, filename: name, mime: mime };
}

function hideChrome() {
  const nodes = document.querySelectorAll(".topbar, .foot, .modal, .msg");
  const prev = [];
  nodes.forEach(function (el) {
    prev.push(el.style.display);
    el.style.display = "none";
  });
  return function restore() {
    nodes.forEach(function (el, i) { el.style.display = prev[i]; });
  };
}

async function waitForImages(root) {
  const imgs = Array.from((root && root.querySelectorAll) ? root.querySelectorAll("img") : []);
  await Promise.all(imgs.map(function (img) {
    if (img.complete) return Promise.resolve();
    return new Promise(function (resolve) {
      img.onload = img.onerror = resolve;
    });
  }));
}

function ignoreHidden(el) {
  return !!(el && (el.hidden || (el.getAttribute && el.getAttribute("hidden") !== null)));
}

async function pdfBlobFromSheet(sheet) {
  await waitForImages(sheet);
  const blob = await html2pdf()
    .set({
      margin: [8, 8, 8, 8],
      image: { type: "jpeg", quality: 0.95 },
      html2canvas: {
        scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false,
        ignoreElements: ignoreHidden
      },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] }
    })
    .from(sheet)
    .outputPdf("blob");
  return blob;
}

async function pngBlobFromSheet(sheet) {
  if (typeof html2canvas === "function") {
    const canvas = await html2canvas(sheet, { scale: 2, useCORS: true, backgroundColor: "#ffffff" });
    return await new Promise(function (resolve, reject) {
      canvas.toBlob(function (b) {
        if (b) resolve(b);
        else reject(new Error("이미지 변환에 실패했습니다."));
      }, "image/png");
    });
  }
  const dataUri = await html2pdf()
    .set({
      image: { type: "png", quality: 1 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false }
    })
    .from(sheet)
    .outputImg("datauristring");
  const res = await fetch(dataUri);
  return res.blob();
}

async function captureSheetPdf(sheet, filename) {
  if (!sheet) throw new Error("저장할 문서를 찾지 못했습니다.");
  await ensurePdfLib();
  const restore = hideChrome();
  try {
    let blob = null;
    let name = filename;
    let kind = "pdf";
    try {
      blob = await pdfBlobFromSheet(sheet);
      const head = new Uint8Array(await blob.slice(0, 48).arrayBuffer());
      if (looksLikeHtml(head, blob.type) || !looksLikePdf(head)) blob = null;
    } catch (err) {
      blob = null;
    }
    if (!blob) {
      blob = await pngBlobFromSheet(sheet);
      const pngHead = new Uint8Array(await blob.slice(0, 48).arrayBuffer());
      if (looksLikeHtml(pngHead, blob.type) || !looksLikePng(pngHead)) {
        throw new Error("PDF/이미지 변환에 실패했습니다.");
      }
      name = filename.replace(/\.pdf$/i, ".png");
      kind = "png";
    }
    const dataUri = await blobToDataUri(blob);
    return { blob: blob, dataUri: dataUri, filename: name, kind: kind };
  } finally {
    restore();
  }
}

async function downloadCaptured(captured) {
  const saved = await downloadBlob(captured.blob, captured.filename);
  captured.filename = saved.filename;
  captured.kind = saved.mime.indexOf("png") >= 0 ? "png" : "pdf";
  return captured;
}

function downloadDataUri(dataUri, filename) {
  return fetch(dataUri)
    .then(function (res) { return res.blob(); })
    .then(function (blob) { return downloadBlob(blob, filename); });
}

function exportPayloadFields(captured) {
  if (captured.kind === "png") return { png_base64: captured.dataUri };
  return { pdf_base64: captured.dataUri };
}

function showDownloadLink(container, captured, message) {
  if (!container) return;
  const url = URL.createObjectURL(captured.blob);
  container.className = "msg ok";
  container.textContent = "";
  container.appendChild(document.createTextNode(message + " "));
  const a = document.createElement("a");
  a.href = url;
  a.download = captured.filename;
  a.rel = "noopener";
  a.textContent = captured.filename + " 받기";
  a.style.fontWeight = "700";
  a.style.textDecoration = "underline";
  container.appendChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 120000);
}
