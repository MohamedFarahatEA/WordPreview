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
    let fields = [];                 // field descriptors: {key,type,format,options}
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

            fields = data.fields || [];
            buildForm(fields);
            setStatus("");
            previewHint.textContent = fields.length
                ? `${fields.length} field${fields.length === 1 ? "" : "s"}`
                : "No placeholders";
            await renderPreview(); // initial render (values empty)
        } catch (err) {
            setStatus(err.message, "err");
            previewHint.textContent = "Error";
        }
    }

    function buildForm(list) {
        fieldsForm.innerHTML = "";
        fieldCount.textContent = String(list.length);

        if (list.length === 0) {
            fieldsWrap.hidden = true;
            emptyFields.hidden = false;
            actions.hidden = false; // still allow downloading a copy
            return;
        }
        emptyFields.hidden = true;
        fieldsWrap.hidden = false;
        actions.hidden = false;

        for (const field of list) {
            const wrap = document.createElement("div");
            wrap.className = "field";

            const label = document.createElement("label");
            label.htmlFor = fieldId(field.key);
            const keySpan = document.createElement("span");
            keySpan.className = "key";
            keySpan.textContent = field.key;
            keySpan.setAttribute("dir", "auto");
            label.append(keySpan);

            const control = createControl(field);
            wrap.append(label, control);
            fieldsForm.append(wrap);
        }
    }

    // Builds the right input control for a field descriptor {key,type,format,options}.
    function createControl(field) {
        let control;
        switch (field.type) {
            case "date":
                control = document.createElement("input");
                control.type = "date";
                break;
            case "select":
                control = document.createElement("select");
                control.append(new Option("— Select —", ""));
                for (const opt of field.options || []) control.append(new Option(opt, opt));
                break;
            case "multiline":
                control = document.createElement("textarea");
                control.placeholder = field.key;
                break;
            default:
                control = document.createElement("input");
                control.type = "text";
                control.placeholder = field.key;
        }
        control.id = fieldId(field.key);
        control.dataset.key = field.key;
        // Auto-detect direction: Arabic values align RTL, Latin values LTR.
        control.setAttribute("dir", "auto");
        // `input` covers text/textarea/date; `change` covers <select>.
        control.addEventListener("input", onFieldInput);
        control.addEventListener("change", onFieldInput);
        return control;
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
            if (seq === renderSeq) applyAutoDirection(previewEl);
        } catch (err) {
            if (seq === renderSeq) {
                previewEl.innerHTML =
                    `<div class="placeholder-empty">Preview failed: ${escapeHtml(err.message)}</div>`;
            }
        } finally {
            if (seq === renderSeq) previewOverlay.hidden = true;
        }
    }

    // docx-preview only right-aligns paragraphs that carry an explicit RTL property.
    // Word, however, auto-detects Arabic/Hebrew and shows it RTL. Match that in the
    // preview: for any block whose first strong character is RTL, set dir="rtl" and,
    // if it's left/start-aligned, flip it to right — while leaving center/justify
    // paragraphs alone so intentional layouts survive. Preview-only; the downloaded
    // file is never touched.
    function applyAutoDirection(root) {
        const blocks = root.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li, td, th, div");
        for (const el of blocks) {
            if (!isRtlText(el.textContent || "")) continue;
            el.setAttribute("dir", "rtl");
            const align = getComputedStyle(el).textAlign;
            if (align === "left" || align === "start") el.style.textAlign = "right";
        }
    }

    // First-strong-character heuristic (same rule the browser uses for dir="auto"):
    // true if the first letter that has a direction is Arabic/Hebrew.
    function isRtlText(s) {
        for (const ch of s) {
            const c = ch.codePointAt(0);
            if ((c >= 0x0590 && c <= 0x08FF) || (c >= 0xFB50 && c <= 0xFDFF) || (c >= 0xFE70 && c <= 0xFEFF))
                return true;
            if ((c >= 0x41 && c <= 0x5A) || (c >= 0x61 && c <= 0x7A) || (c >= 0xC0 && c <= 0x24F))
                return false;
        }
        return false;
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
        fields = [];
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
