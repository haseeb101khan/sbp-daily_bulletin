const form = document.querySelector("#generate-form");
const uploadZone = document.querySelector("#upload-zone");
const fileInput = document.querySelector("#excel-file");
const fileName = document.querySelector("#file-name");
const statusLine = document.querySelector("#status");
const button = document.querySelector("#generate-button");
const result = document.querySelector("#result");
const resultSummary = document.querySelector("#result-summary");
const previewLink = document.querySelector("#preview-link");
const downloadLink = document.querySelector("#download-link");
const warnings = document.querySelector("#warnings");
const warningList = document.querySelector("#warning-list");
let previewObjectUrl = null;
let downloadObjectUrl = null;

function revokeGeneratedLinks() {
  if (previewObjectUrl) {
    URL.revokeObjectURL(previewObjectUrl);
    previewObjectUrl = null;
  }
  if (downloadObjectUrl) {
    URL.revokeObjectURL(downloadObjectUrl);
    downloadObjectUrl = null;
  }
}

function blobFromBase64(base64, contentType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: contentType });
}

fileInput.addEventListener("change", () => {
  const hasFile = fileInput.files.length > 0;
  uploadZone.classList.toggle("has-file", hasFile);
  fileName.textContent = hasFile
    ? `File uploaded: ${fileInput.files[0].name}`
    : "Choose the filled template to generate a Word bulletin.";
  statusLine.textContent = hasFile ? "Excel file uploaded and ready." : "";
  statusLine.classList.remove("error");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  revokeGeneratedLinks();
  result.hidden = true;
  warnings.hidden = true;
  warningList.innerHTML = "";
  statusLine.classList.remove("error");

  if (!fileInput.files.length) {
    statusLine.textContent = "Choose a completed Excel file first.";
    statusLine.classList.add("error");
    return;
  }

  const data = new FormData(form);
  button.disabled = true;
  button.textContent = "Generating...";
  statusLine.textContent = "Building the bulletin preview and Word report.";

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: data,
    });
    const payload = await response.json();
    if (!payload.ok) {
      throw new Error(payload.error || "The report could not be generated.");
    }

    previewObjectUrl = URL.createObjectURL(
      new Blob([payload.previewHtml], { type: "text/html; charset=utf-8" }),
    );
    downloadObjectUrl = URL.createObjectURL(
      blobFromBase64(
        payload.docxBase64,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ),
    );

    previewLink.href = previewObjectUrl;
    downloadLink.href = downloadObjectUrl;
    downloadLink.download = payload.filename;
    resultSummary.textContent = `${payload.articleCount} articles grouped into ${payload.storyCount} stories across ${payload.sectionCount} sections.`;
    result.hidden = false;
    statusLine.textContent = "Report generated.";

    if (payload.warnings && payload.warnings.length) {
      payload.warnings.forEach((item) => {
        const li = document.createElement("li");
        li.textContent = item;
        warningList.appendChild(li);
      });
      warnings.hidden = false;
    }
  } catch (error) {
    statusLine.textContent = error.message;
    statusLine.classList.add("error");
  } finally {
    button.disabled = false;
    button.textContent = "Generate Report";
  }
});
