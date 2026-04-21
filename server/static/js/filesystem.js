// File System Access API wrapper + pipeline page controller.
// On Chrome/Edge: user picks a folder once, we read inputs and write
// outputs back into that same folder. On Firefox/Safari: normal
// <input type=file> upload and downloads go to the Downloads folder.

const FSA_SUPPORTED = typeof window.showDirectoryPicker === "function";

let folderHandle = null;
let selectedFiles = []; // [{ name, file, source: 'input' | 'fsa' }]

async function pickFolder(labelEl, autoloadBtn) {
  if (!FSA_SUPPORTED) {
    alert(
      "Your browser doesn't support folder picking.\n" +
      "Use Chrome or Edge for auto-save to a folder.\n" +
      "Files will still download normally in this browser."
    );
    return;
  }
  try {
    folderHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    labelEl.textContent = "📂 " + folderHandle.name + " (outputs will save here)";
    labelEl.classList.remove("muted");
    if (autoloadBtn) autoloadBtn.classList.remove("hidden");
  } catch (e) {
    if (e.name !== "AbortError") console.error(e);
  }
}

async function autoLoadFromFolder(patterns, fileListEl, runBtn, keepAll) {
  if (!folderHandle) return;
  selectedFiles = [];
  const picked = {};
  const all = [];
  for await (const entry of folderHandle.values()) {
    if (entry.kind !== "file") continue;
    for (const pat of patterns) {
      if (pat.test(entry.name)) {
        const file = await entry.getFile();
        if (keepAll) {
          all.push({ name: entry.name, file, source: "fsa" });
        } else {
          const key = pat.source;
          if (!picked[key] || file.lastModified > picked[key].file.lastModified) {
            picked[key] = { name: entry.name, file, source: "fsa" };
          }
        }
        break;
      }
    }
  }
  selectedFiles = keepAll ? all : Object.values(picked);
  renderFileList(fileListEl);
  runBtn.disabled = selectedFiles.length === 0;
}

function renderFileList(el) {
  el.innerHTML = "";
  for (const f of selectedFiles) {
    const li = document.createElement("li");
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = f.source === "fsa" ? "from folder" : "uploaded";
    li.appendChild(tag);
    const name = document.createElement("span");
    name.textContent = f.name;
    li.appendChild(name);
    el.appendChild(li);
  }
}

async function saveBlobToFolder(blob, filename) {
  const fh = await folderHandle.getFileHandle(filename, { create: true });
  const w = await fh.createWritable();
  await w.write(blob);
  await w.close();
}

function triggerBrowserDownload(url, filename) {
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function initPipelinePage(opts) {
  const pickBtn = document.getElementById("pick-folder");
  const folderLabel = document.getElementById("folder-label");
  const autoloadBtn = document.getElementById("autoload-btn");
  const input = document.getElementById("dxf-input");
  const fileList = document.getElementById("file-list");
  const runBtn = document.getElementById("run-btn");
  const outputCard = document.getElementById("output-card");
  const outputList = document.getElementById("output-list");
  const statusEl = document.getElementById("status");
  const logEl = document.getElementById("log");

  pickBtn.addEventListener("click", () => pickFolder(folderLabel, autoloadBtn));

  if (autoloadBtn && opts.autoloadPatterns) {
    autoloadBtn.addEventListener("click", () =>
      autoLoadFromFolder(opts.autoloadPatterns, fileList, runBtn, !!opts.autoloadAll)
    );
  }

  input.addEventListener("change", () => {
    selectedFiles = Array.from(input.files).map(file => ({
      name: file.name, file, source: "input"
    }));
    renderFileList(fileList);
    runBtn.disabled = selectedFiles.length === 0;
  });

  runBtn.addEventListener("click", async () => {
    if (selectedFiles.length === 0) return;
    runBtn.disabled = true;
    outputCard.classList.remove("hidden");
    outputList.innerHTML = "";
    logEl.classList.add("hidden");
    statusEl.innerHTML = '<span class="status-running">⏳ Uploading and running pipeline…</span>';

    const fd = new FormData();
    for (const f of selectedFiles) {
      fd.append("files", f.file, f.name);
    }

    try {
      const res = await fetch(opts.endpoint, { method: "POST", body: fd });
      const data = await res.json();

      if (!data.ok) {
        statusEl.innerHTML = `<span class="status-fail">✗ ${escapeHtml(data.error || "Pipeline failed")}</span>`;
        if (data.stderr) {
          logEl.textContent = data.stderr;
          logEl.classList.remove("hidden");
        }
        runBtn.disabled = false;
        return;
      }

      statusEl.innerHTML = `<span class="status-ok">✓ Done (${data.outputs.length} files)</span>`;

      // Save or offer each output
      let savedToFolder = 0;
      for (const name of data.outputs) {
        const url = `${opts.downloadBase}/${encodeURIComponent(data.job_id)}/${encodeURIComponent(name)}`;
        const li = document.createElement("li");

        if (folderHandle) {
          try {
            const resp = await fetch(url);
            const blob = await resp.blob();
            await saveBlobToFolder(blob, name);
            savedToFolder++;
            li.innerHTML = `<span class="tag" style="background:#dcfce7;color:#166534">saved</span><span>${escapeHtml(name)}</span>`;
          } catch (err) {
            console.error("FSA save failed, falling back to download", err);
            triggerBrowserDownload(url, name);
            li.innerHTML = `<span class="tag">downloaded</span><span>${escapeHtml(name)}</span>`;
          }
        } else {
          triggerBrowserDownload(url, name);
          li.innerHTML = `<span class="tag">downloaded</span><span>${escapeHtml(name)}</span>`;
        }
        outputList.appendChild(li);
      }

      if (folderHandle && savedToFolder === data.outputs.length) {
        statusEl.innerHTML += ` <span class="muted">→ saved to <code>${escapeHtml(folderHandle.name)}</code></span>`;
      }

      if (data.stdout) {
        logEl.textContent = data.stdout;
      }
    } catch (err) {
      statusEl.innerHTML = `<span class="status-fail">✗ Network error: ${escapeHtml(err.message)}</span>`;
    } finally {
      runBtn.disabled = false;
    }
  });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

window.initPipelinePage = initPipelinePage;
