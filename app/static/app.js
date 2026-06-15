const form = document.querySelector("#generate-form");
const sourceInput = document.querySelector("#source_docx");
const templateInput = document.querySelector("#template_docx");
const sourceName = document.querySelector("#source-name");
const sourceSize = document.querySelector("#source-size");
const outputName = document.querySelector("#output-name");
const footerName = document.querySelector("#footer-name");
const submit = document.querySelector("#submit");
const statusTitle = document.querySelector("#status-title");
const statusText = document.querySelector("#status-text");
const runAnalysisButton = document.querySelector("#run-analysis");
const analysisBox = document.querySelector("#analysis");
const correctionsArea = document.querySelector("#corrections-area");
const correctionCount = document.querySelector("#correction-count");
const correctionList = document.querySelector("#correction-list");
const applyAllButton = document.querySelector("#apply-all");
const linksArea = document.querySelector("#links-area");
const linkCount = document.querySelector("#link-count");
const linkList = document.querySelector("#link-list");
const rerunButton = document.querySelector("#rerun");

let latestAnalysis = null;
const appliedCorrectionIndexes = new Set();

function safeSlug(value) {
  return value.trim().replaceAll(".", "-").replace(/[^A-Za-z0-9-]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
}

function currentFilename() {
  const data = new FormData(form);
  const parts = [
    data.get("unit_code"),
    data.get("year"),
    data.get("term"),
    data.get("week_number") ? `week-${data.get("week_number")}` : "",
    data.get("version") ? `v${data.get("version")}` : "",
  ].filter(Boolean).map(safeSlug);
  if (parts.length) return `${parts.join("_")}.docx`;
  return sourceInput.files[0]?.name || "source-document.docx";
}

function refreshFilename() {
  const name = currentFilename();
  outputName.textContent = name;
  footerName.textContent = name;
}

sourceInput.addEventListener("change", () => {
  const file = sourceInput.files[0];
  if (file) {
    sourceName.textContent = file.name;
    sourceSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;
  }
  resetAnalysis();
  runAnalysisButton.disabled = !file;
  refreshFilename();
});

form.addEventListener("input", refreshFilename);

runAnalysisButton.addEventListener("click", runAnalysis);
rerunButton.addEventListener("click", runAnalysis);
applyAllButton.addEventListener("click", () => {
  (latestAnalysis?.corrections || []).forEach((_, index) => appliedCorrectionIndexes.add(index));
  correctionList.querySelectorAll(".correction-card").forEach((card) => {
    card.classList.add("applied");
    const button = card.querySelector("button");
    if (button) {
      button.disabled = true;
      button.textContent = "Applied";
    }
  });
  applyAllButton.disabled = true;
  updateAppliedStatus();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!sourceInput.files[0]) return;

  submit.disabled = true;
  submit.textContent = "Generating...";
  statusTitle.textContent = "Working";
  statusText.textContent = "Uploading document.";

  const data = new FormData(form);
  if (!templateInput.files[0]) data.delete("template_docx");
  data.set("accepted_corrections", JSON.stringify(appliedCorrections()));

  try {
    const start = await fetchWithRetry("/generate", { method: "POST", body: data });
    if (!start.ok) throw new Error(await start.text());
    const job = await start.json();
    await poll(job.job_id);
  } catch (error) {
    statusTitle.textContent = "Generation failed";
    statusText.textContent = error.message;
    submit.disabled = false;
    submit.textContent = "⇩ Generate & download";
  }
});

async function poll(jobId) {
  const timer = setInterval(async () => {
    const response = await fetch(`/jobs/${jobId}`);
    const job = await response.json();
    statusTitle.textContent = job.status.replaceAll("_", " ");
    statusText.textContent = job.error || job.message || "Working.";
    if (job.status === "completed") {
      clearInterval(timer);
      window.location.href = `/download/${jobId}`;
      submit.disabled = false;
      submit.textContent = "⇩ Generate & download";
      statusTitle.textContent = "Download started";
      statusText.textContent = `Saved as ${job.output_filename}`;
    }
    if (job.status === "failed") {
      clearInterval(timer);
      submit.disabled = false;
      submit.textContent = "⇩ Generate & download";
    }
  }, 900);
}

async function fetchWithRetry(url, options = {}, attempts = 2) {
  let lastError;
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    try {
      return await fetch(url, options);
    } catch (error) {
      lastError = error;
      await delay(350);
    }
  }
  throw lastError;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function downloadJob(jobId) {
  const link = document.createElement("a");
  link.href = `/download/${jobId}`;
  link.download = "";
  document.body.append(link);
  link.click();
  link.remove();
}

async function poll(jobId) {
  return new Promise((resolve, reject) => {
    const timer = setInterval(async () => {
      let job;
      try {
        const response = await fetchWithRetry(`/jobs/${jobId}`);
        if (!response.ok) throw new Error(await response.text());
        job = await response.json();
      } catch (error) {
        clearInterval(timer);
        submit.disabled = false;
        submit.textContent = "Generate & download";
        reject(error);
        return;
      }

      statusTitle.textContent = job.status.replaceAll("_", " ");
      statusText.textContent = job.error || job.message || "Working.";
      if (job.status === "completed") {
        clearInterval(timer);
        downloadJob(jobId);
        submit.disabled = false;
        submit.textContent = "Generate & download";
        statusTitle.textContent = "Download started";
        statusText.textContent = `Saved as ${job.output_filename}`;
        resolve(job);
      }
      if (job.status === "failed") {
        clearInterval(timer);
        submit.disabled = false;
        submit.textContent = "Generate & download";
        reject(new Error(job.error || "Generation failed."));
      }
    }, 900);
  });
}

refreshFilename();

function resetAnalysis() {
  latestAnalysis = null;
  appliedCorrectionIndexes.clear();
  analysisBox.classList.add("hidden");
  correctionsArea.classList.add("hidden");
  linksArea.classList.add("hidden");
  rerunButton.classList.add("hidden");
  correctionList.replaceChildren();
  linkList.replaceChildren();
  correctionCount.textContent = "Suggested corrections (0)";
  linkCount.textContent = "Hyperlinks (0)";
  applyAllButton.disabled = true;
  runAnalysisButton.classList.remove("hidden");
}

async function runAnalysis() {
  if (!sourceInput.files[0]) return;
  runAnalysisButton.disabled = true;
  rerunButton.disabled = true;
  runAnalysisButton.textContent = "Analyzing...";
  analysisBox.classList.remove("hidden");
  analysisBox.textContent = "Checking grammar, consistency, and hyperlinks.";
  statusTitle.textContent = "Analyzing";
  statusText.textContent = "Running AI proofread and link check.";

  const data = new FormData(form);
  data.delete("template_docx");

  try {
    const response = await fetch("/analyze", { method: "POST", body: data });
    if (!response.ok) throw new Error(await response.text());
    latestAnalysis = await response.json();
    renderAnalysis(latestAnalysis);
    statusTitle.textContent = "Analysis complete";
    statusText.textContent = analysisStatusText(latestAnalysis);
  } catch (error) {
    analysisBox.textContent = `Analysis failed: ${friendlyError(error.message)}`;
    statusTitle.textContent = "Analysis failed";
    statusText.textContent = friendlyError(error.message);
  } finally {
    runAnalysisButton.textContent = "Run AI proofread & link check";
    runAnalysisButton.disabled = !sourceInput.files[0];
    rerunButton.disabled = false;
  }
}

function renderAnalysis(result) {
  runAnalysisButton.classList.add("hidden");
  rerunButton.classList.remove("hidden");
  analysisBox.classList.remove("hidden");
  const warnings = result.warnings || [];
  analysisBox.textContent = warnings.length
    ? `${result.summary || "Analysis complete."} ${warnings[0]}`
    : result.summary || "Analysis complete.";

  correctionsArea.classList.remove("hidden");
  correctionList.replaceChildren();
  appliedCorrectionIndexes.clear();
  const corrections = result.corrections || [];
  correctionCount.textContent = `Suggested corrections (${corrections.length})`;
  applyAllButton.disabled = corrections.length === 0;
  if (corrections.length === 0) {
    correctionList.append(emptyLine("No suggested corrections."));
  } else {
    corrections.forEach((correction, index) => correctionList.append(correctionCard(correction, index)));
  }

  linksArea.classList.remove("hidden");
  linkList.replaceChildren();
  const links = result.hyperlinks || [];
  linkCount.textContent = `Hyperlinks (${links.length})`;
  if (links.length === 0) {
    linkList.append(emptyLine("No hyperlinks in document."));
  } else {
    links.forEach((link) => linkList.append(linkCard(link)));
  }
}

function correctionCard(correction, index) {
  const card = document.createElement("div");
  card.className = "correction-card";

  const text = document.createElement("div");
  const original = document.createElement("div");
  original.className = "original";
  original.textContent = correction.original;
  const replacement = document.createElement("div");
  replacement.className = "replacement";
  replacement.textContent = correction.replacement;
  const reason = document.createElement("small");
  reason.textContent = correction.reason || "Suggested proofreading correction.";
  text.append(original, replacement, reason);

  const apply = document.createElement("button");
  apply.type = "button";
  apply.className = "ghost";
  apply.textContent = "Apply";
  apply.addEventListener("click", () => {
    appliedCorrectionIndexes.add(index);
    card.classList.add("applied");
    apply.disabled = true;
    apply.textContent = "Applied";
    updateAppliedStatus();
  });
  card.append(text, apply);
  return card;
}

function appliedCorrections() {
  const corrections = latestAnalysis?.corrections || [];
  return [...appliedCorrectionIndexes]
    .sort((a, b) => a - b)
    .map((index) => corrections[index])
    .filter(Boolean)
    .map((correction) => ({
      block_index: correction.block_index,
      original: correction.original,
      replacement: correction.replacement,
    }));
}

function updateAppliedStatus() {
  const count = appliedCorrectionIndexes.size;
  const total = latestAnalysis?.corrections?.length || 0;
  statusTitle.textContent = count ? "Corrections applied" : "Analysis complete";
  statusText.textContent = count
    ? `${count} of ${total} corrections will be included in the generated DOCX.`
    : `${total} corrections available.`;
  applyAllButton.disabled = total === 0 || count === total;
}

function linkCard(link) {
  const card = document.createElement("div");
  card.className = "link-card";
  const label = document.createElement("strong");
  label.textContent = link.text || link.url;
  const url = document.createElement("small");
  url.textContent = link.url;
  const detail = document.createElement("small");
  detail.textContent = link.detail || "";
  const status = document.createElement("span");
  const statusName = link.status || "needs_review";
  status.className = `pill link-status ${statusName}`;
  status.textContent = statusLabel(statusName);
  const text = document.createElement("div");
  text.append(label, url);
  if (detail.textContent) text.append(detail);
  card.append(text, status);
  return card;
}

function analysisStatusText(result) {
  const corrections = result.corrections?.length || 0;
  const links = result.hyperlinks || [];
  const broken = links.filter((link) => link.status === "broken").length;
  const review = links.filter((link) => link.status === "needs_review").length;
  if (broken || review) {
    return `${corrections} corrections, ${links.length} hyperlinks, ${broken} broken, ${review} need review.`;
  }
  return `${corrections} corrections, ${links.length} hyperlinks checked.`;
}

function statusLabel(status) {
  if (status === "ok") return "OK";
  if (status === "broken") return "Broken";
  if (status === "needs_review") return "Review";
  return status;
}

function emptyLine(text) {
  const item = document.createElement("p");
  item.className = "empty-line";
  item.textContent = text;
  return item;
}

function friendlyError(message) {
  try {
    const parsed = JSON.parse(message);
    return parsed.detail || message;
  } catch {
    return message;
  }
}
