/* WordPreview — upload a .docx, fill {{placeholders}}, preview live, download. */
(() => {
    "use strict";

    const el = (id) => document.getElementById(id);
    const dropzone = el("dropzone");
    const fileInput = el("fileInput");
    const browseBtn = el("browseBtn");
    const resetBtn = el("resetBtn");
    const fileMeta = el("fileMeta");
    const fileNameEl = el("fileName");
    const fieldsWrap = el("fieldsWrap");
    const fieldsForm = el("fieldsForm");
    const fieldCount = el("fieldCount");
    const emptyFields = el("emptyFields");
    const actions = el("actions");
    const downloadBtn = el("downloadBtn");
    const statusEl = el("status");
    const previewEl = el("preview");
    const previewHint = el("previewHint");
    const previewOverlay = el("previewOverlay");

    // Current session state.
    let currentFile = null;          // File object of the uploaded .docx
    let placeholders = [];           // string[] distinct keys
    let renderSeq = 0;               // guards out-of-order async renders

    // --- upload wiring --------------------------------------------------------

    browseBtn.addEventListener("click", (e) => { e.stopPropagation(); fileInput.click(); });
    dropzone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });

    ["dragenter", "dragover"].forEach((evt) =>
        dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add("dragover"); }));
    ["dragleave", "drop"].forEach((evt) =>
        dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove("dragover"); }));
    dropzone.addEventListener("drop", (e) => {
        const f = e.dataTransfer?.files?.[0];
        if (f) handleFile(f);
    });

    resetBtn.addEventListener("click", () => resetAll());

    // --- core flow ------------------------------------------------------------

    async function handleFile(file) {
        if (!file.name.toLowerCase().endsWith(".docx")) {
            setStatus("Please choose a .docx file.", "err");
            return;
        }
        currentFile = file;
        fileNameEl.textContent = file.name;
        fileMeta.hidden = false;
        dropzone.hidden = true;
        setStatus("Analyzing…");

        try {
            const fd = new FormData();
            fd.append("file", file);
            const res = await fetch("/api/analyze", { method: "POST", body: fd });
            const data = await res.json();
            if (!res.ok) throw new Error(data.error || "Analyze failed.");

            placeholders = data.placeholders || [];
            buildForm(placeholders);
            setStatus("");
            previewHint.textContent = placeholders.length
                ? `${placeholders.length} field${placeholders.length === 1 ? "" : "s"}`
                : "No placeholders";
            await renderPreview(); // initial render (values empty)
        } catch (err) {
            setStatus(err.message, "err");
            previewHint.textContent = "Error";
        }
    }

    function buildForm(keys) {
        fieldsForm.innerHTML = "";
        fieldCount.textContent = String(keys.length);

        if (keys.length === 0) {
            fieldsWrap.hidden = true;
            emptyFields.hidden = false;
            actions.hidden = false; // still allow downloading a copy
            return;
        }
        emptyFields.hidden = true;
        fieldsWrap.hidden = false;
        actions.hidden = false;

        for (const key of keys) {
            const wrap = document.createElement("div");
            wrap.className = "field";

            const label = document.createElement("label");
            label.htmlFor = fieldId(key);
            label.innerHTML = `<span class="key"></span>`;
            label.querySelector(".key").textContent = key;

            // Multi-line-ish keys (address, notes, description) get a textarea.
            const multiline = /address|notes?|description|comment|details|body/i.test(key);
            const input = multiline ? document.createElement("textarea") : document.createElement("input");
            if (!multiline) input.type = "text";
            input.id = fieldId(key);
            input.dataset.key = key;
            input.placeholder = key;
            input.addEventListener("input", onFieldInput);

            wrap.append(label, input);
            fieldsForm.append(wrap);
        }
    }

    const onFieldInput = debounce(() => renderPreview(), 350);

    function collectValues() {
        const values = {};
        for (const input of fieldsForm.querySelectorAll("[data-key]")) {
            values[input.dataset.key] = input.value;
        }
        return values;
    }

    // --- preview --------------------------------------------------------------

    async function renderPreview() {
        if (!currentFile) return;
        const seq = ++renderSeq;
        previewOverlay.hidden = false;

        try {
            const blob = await fillDocx(collectValues());
            if (seq !== renderSeq) return; // a newer render superseded this one

            previewEl.innerHTML = "";
            await window.docx.renderAsync(blob, previewEl, undefined, {
                className: "docx",
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                breakPages: true,
                experimental: true,
            });
        } catch (err) {
            if (seq === renderSeq) {
                previewEl.innerHTML =
                    `<div class="placeholder-empty">Preview failed: ${escapeHtml(err.message)}</div>`;
            }
        } finally {
            if (seq === renderSeq) previewOverlay.hidden = true;
        }
    }

    // Calls the backend to substitute values and returns the resulting .docx blob.
    async function fillDocx(values) {
        const fd = new FormData();
        fd.append("file", currentFile);
        fd.append("values", JSON.stringify(values));
        const res = await fetch("/api/fill", { method: "POST", body: fd });
        if (!res.ok) {
            let msg = "Fill failed.";
            try { msg = (await res.json()).error || msg; } catch { /* ignore */ }
            throw new Error(msg);
        }
        return await res.blob();
    }

    // --- download -------------------------------------------------------------

    downloadBtn.addEventListener("click", async () => {
        if (!currentFile) return;
        downloadBtn.disabled = true;
        setStatus("Generating…");
        try {
            const blob = await fillDocx(collectValues());
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = currentFile.name.replace(/\.docx$/i, "") + "-filled.docx";
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 4000);
            setStatus("Downloaded.", "ok");
        } catch (err) {
            setStatus(err.message, "err");
        } finally {
            downloadBtn.disabled = false;
        }
    });

    // --- helpers --------------------------------------------------------------

    function resetAll() {
        currentFile = null;
        placeholders = [];
        fileInput.value = "";
        fieldsForm.innerHTML = "";
        previewEl.innerHTML = "";
        fileMeta.hidden = true;
        fieldsWrap.hidden = true;
        emptyFields.hidden = true;
        actions.hidden = true;
        dropzone.hidden = false;
        previewHint.textContent = "Upload a document to begin";
        setStatus("");
    }

    function setStatus(msg, kind) {
        statusEl.textContent = msg || "";
        statusEl.className = "status" + (kind ? " " + kind : "");
    }

    function fieldId(key) { return "f_" + key.replace(/[^a-z0-9]+/gi, "_"); }

    function debounce(fn, ms) {
        let t;
        return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, (c) =>
            ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    }
})();
