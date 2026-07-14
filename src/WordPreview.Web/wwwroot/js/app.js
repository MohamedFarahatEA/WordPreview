/* WordPreview — fill {{placeholders}} in a .docx, preview live, download.
   Two editing modes: a side panel, or inline inputs overlaid on the preview. */
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
    const previewScroll = el("previewScroll");
    const previewHint = el("previewHint");
    const previewOverlay = el("previewOverlay");
    const overlay = el("overlay");
    const modeToggle = el("modeToggle");

    // Locator sentinels the server wraps around each value in "marks" mode.
    const MARK_RX = /(\d+)([\s\S]*?)/g;

    // Session state.
    let currentFile = null;              // uploaded .docx File
    let fields = [];                     // [{key,type,format,options}]
    let mode = "panel";                  // "panel" | "inline"
    let renderSeq = 0;                   // guards out-of-order async renders
    const values = Object.create(null);  // key -> string (single source of truth)
    const overlayControls = new Map();   // anchorId -> control element

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

    // --- mode toggle ----------------------------------------------------------

    modeToggle.addEventListener("click", (e) => {
        const btn = e.target.closest(".mode-btn");
        if (!btn) return;
        setMode(btn.dataset.mode);
    });

    function setMode(next) {
        if (next === mode) return;
        mode = next;
        for (const b of modeToggle.querySelectorAll(".mode-btn"))
            b.classList.toggle("is-active", b.dataset.mode === mode);
        document.body.classList.toggle("mode-inline", mode === "inline");
        if (mode !== "inline") clearOverlays();
        renderPreview();
    }

    // reposition overlays when the preview scrolls or the window resizes
    previewScroll.addEventListener("scroll", () => { if (mode === "inline") requestAnimationFrame(layoutOverlays); });
    window.addEventListener("resize", () => { if (mode === "inline") renderPreview(); });

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
            for (const k in values) delete values[k];
            for (const f of fields) values[f.key] = "";

            buildForm(fields);
            modeToggle.hidden = fields.length === 0;
            setStatus("");
            previewHint.textContent = fields.length
                ? `${fields.length} field${fields.length === 1 ? "" : "s"}`
                : "No placeholders";
            await renderPreview();
        } catch (err) {
            setStatus(err.message, "err");
            previewHint.textContent = "Error";
        }
    }

    // --- side-panel form ------------------------------------------------------

    function buildForm(list) {
        fieldsForm.innerHTML = "";
        fieldCount.textContent = String(list.length);

        if (list.length === 0) {
            fieldsWrap.hidden = true;
            emptyFields.hidden = false;
            actions.hidden = false;
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

            const control = buildControl(field, "panel");
            control.id = fieldId(field.key);
            wrap.append(label, control);
            fieldsForm.append(wrap);
        }
    }

    // Builds a control (input/textarea/date/select) bound to values[field.key].
    // `where` is "panel" or "inline" (only affects class + placeholder).
    function buildControl(field, where) {
        let control;
        switch (field.type) {
            case "date":
                control = document.createElement("input");
                control.type = "date";
                break;
            case "select":
                control = document.createElement("select");
                control.append(new Option(where === "inline" ? "—" : "— Select —", ""));
                for (const opt of field.options || []) control.append(new Option(opt, opt));
                break;
            case "multiline":
                control = document.createElement(where === "inline" ? "input" : "textarea");
                if (where === "inline") control.type = "text";
                control.placeholder = field.key;
                break;
            default:
                control = document.createElement("input");
                control.type = "text";
                control.placeholder = field.key;
        }
        control.dataset.key = field.key;
        control.value = values[field.key] ?? "";
        control.setAttribute("dir", "auto");
        const handler = () => setValue(field.key, control.value, control);
        control.addEventListener("input", handler);
        control.addEventListener("change", handler);
        return control;
    }

    // Update the model, mirror to every other control for that key, re-render.
    function setValue(key, v, sourceEl) {
        values[key] = v;
        const selector = `[data-key="${cssEscape(key)}"]`;
        for (const other of document.querySelectorAll(selector)) {
            if (other !== sourceEl && other.value !== v) other.value = v;
        }
        scheduleRender();
    }

    const scheduleRender = debounce(() => renderPreview(), 350);

    function collectValues() {
        return Object.assign(Object.create(null), values);
    }

    // --- preview + overlays ---------------------------------------------------

    async function renderPreview() {
        if (!currentFile) return;
        const seq = ++renderSeq;
        const inline = mode === "inline";
        previewOverlay.hidden = false;

        try {
            const blob = await fillDocx(collectValues(), inline);
            if (seq !== renderSeq) return;

            previewEl.innerHTML = "";
            await window.docx.renderAsync(blob, previewEl, undefined, {
                className: "docx",
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                breakPages: true,
                experimental: true,
            });
            if (seq !== renderSeq) return;

            applyAutoDirection(previewEl);
            if (inline) {
                extractAnchors(previewEl);   // strip sentinels, wrap values in .ph-anchor
                layoutOverlays();
            }
        } catch (err) {
            if (seq === renderSeq)
                previewEl.innerHTML =
                    `<div class="placeholder-empty">Preview failed: ${escapeHtml(err.message)}</div>`;
        } finally {
            if (seq === renderSeq) previewOverlay.hidden = true;
        }
    }

    // Calls the backend to substitute values; `marks` requests locator sentinels.
    async function fillDocx(vals, marks) {
        const fd = new FormData();
        fd.append("file", currentFile);
        fd.append("values", JSON.stringify(vals));
        if (marks) fd.append("marks", "true");
        const res = await fetch("/api/fill", { method: "POST", body: fd });
        if (!res.ok) {
            let msg = "Fill failed.";
            try { msg = (await res.json()).error || msg; } catch { /* ignore */ }
            throw new Error(msg);
        }
        return await res.blob();
    }

    // Replace each sentinel-wrapped value with a <span class="ph-anchor"> we can
    // measure, and delete the sentinel characters so nothing shows through.
    function extractAnchors(root) {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        const targets = [];
        let node;
        while ((node = walker.nextNode())) {
            if (node.nodeValue && node.nodeValue.indexOf("") !== -1) targets.push(node);
        }
        for (const n of targets) {
            const text = n.nodeValue;
            const color = getComputedStyle(n.parentElement).color;
            const frag = document.createDocumentFragment();
            let last = 0, m;
            MARK_RX.lastIndex = 0;
            while ((m = MARK_RX.exec(text))) {
                frag.appendChild(document.createTextNode(text.slice(last, m.index)));
                const span = document.createElement("span");
                span.className = "ph-anchor";
                span.dataset.fieldIndex = m[1];
                span.dataset.color = color;
                if (m[2].length) {
                    span.textContent = m[2];
                } else {
                    // Empty blank: reserve space in the flow so the box doesn't
                    // overlap the following text (see .ph-anchor.is-empty).
                    span.textContent = "​"; // ZWSP keeps a line box
                    span.classList.add("is-empty");
                }
                frag.appendChild(span);
                last = m.index + m[0].length;
            }
            frag.appendChild(document.createTextNode(text.slice(last)));
            n.parentNode.replaceChild(frag, n);
        }
    }

    // Position (create/reuse/remove) an overlay control over every anchor.
    function layoutOverlays() {
        if (mode !== "inline") return;
        const anchors = previewEl.querySelectorAll(".ph-anchor");
        const scRect = previewScroll.getBoundingClientRect();
        const occ = Object.create(null);
        const seen = new Set();

        for (const a of anchors) {
            const fi = a.dataset.fieldIndex;
            const field = fields[fi];
            if (!field) continue;
            const j = (occ[fi] = (occ[fi] === undefined ? 0 : occ[fi] + 1));
            const id = fi + "#" + j;
            seen.add(id);

            let ctrl = overlayControls.get(id);
            if (!ctrl) {
                ctrl = buildControl(field, "inline");
                ctrl.classList.add("ph-input");
                overlay.appendChild(ctrl);
                overlayControls.set(id, ctrl);
            }

            const r = a.getBoundingClientRect();
            const cs = getComputedStyle(a);
            const lineH = parseFloat(cs.fontSize) * 1.35;
            const h = Math.max(r.height || 0, lineH);
            const top = r.top - scRect.top + previewScroll.scrollTop - (h - (r.height || lineH)) / 2;
            const left = r.left - scRect.left + previewScroll.scrollLeft;
            const isSelect = ctrl.tagName === "SELECT";

            ctrl.style.top = Math.round(top) + "px";
            ctrl.style.left = Math.round(left) + "px";
            ctrl.style.height = Math.round(h) + "px";
            ctrl.style.width = Math.max(Math.ceil(r.width) + 8, isSelect ? 56 : 44) + "px";
            ctrl.style.fontSize = cs.fontSize;
            ctrl.style.fontFamily = cs.fontFamily;
            ctrl.style.fontWeight = cs.fontWeight;
            ctrl.style.fontStyle = cs.fontStyle;
            ctrl.style.color = a.dataset.color || "#111";
            ctrl.style.direction = cs.direction;
            ctrl.style.textAlign = cs.direction === "rtl" ? "right" : "left";
            if (document.activeElement !== ctrl && ctrl.value !== (values[field.key] ?? ""))
                ctrl.value = values[field.key] ?? "";
        }

        for (const [id, ctrl] of overlayControls) {
            if (!seen.has(id)) { ctrl.remove(); overlayControls.delete(id); }
        }
    }

    function clearOverlays() {
        for (const [, ctrl] of overlayControls) ctrl.remove();
        overlayControls.clear();
    }

    // --- download -------------------------------------------------------------

    downloadBtn.addEventListener("click", async () => {
        if (!currentFile) return;
        downloadBtn.disabled = true;
        setStatus("Generating…");
        try {
            const blob = await fillDocx(collectValues(), false);
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

    // --- preview direction (RTL) ----------------------------------------------

    // Match Word's auto-detect: Arabic paragraphs render RTL / right-aligned.
    function applyAutoDirection(root) {
        const blocks = root.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li, td, th, div");
        for (const node of blocks) {
            if (!isRtlText(node.textContent || "")) continue;
            node.setAttribute("dir", "rtl");
            const align = getComputedStyle(node).textAlign;
            if (align === "left" || align === "start") node.style.textAlign = "right";
        }
    }

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

    // --- helpers --------------------------------------------------------------

    function resetAll() {
        currentFile = null;
        fields = [];
        for (const k in values) delete values[k];
        clearOverlays();
        fileInput.value = "";
        fieldsForm.innerHTML = "";
        previewEl.innerHTML = "";
        fileMeta.hidden = true;
        fieldsWrap.hidden = true;
        emptyFields.hidden = true;
        actions.hidden = true;
        modeToggle.hidden = true;
        dropzone.hidden = false;
        previewHint.textContent = "Upload a document to begin";
        setStatus("");
    }

    function setStatus(msg, kind) {
        statusEl.textContent = msg || "";
        statusEl.className = "status" + (kind ? " " + kind : "");
    }

    function fieldId(key) { return "f_" + key.replace(/[^a-z0-9]+/gi, "_"); }

    function cssEscape(s) {
        return window.CSS && CSS.escape ? CSS.escape(s) : s.replace(/["\\]/g, "\\$&");
    }

    function debounce(fn, ms) {
        let t;
        return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, (c) =>
            ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    }
})();
