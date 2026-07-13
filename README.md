# WordPreview

Turn a Word (`.docx`) file into a fillable web form, preview it live, and download
the finished document — **with the original formatting and page size fully preserved**.

Blanks are marked in the Word file with double curly braces, e.g. `{{client_name}}`,
`{{Date}}`, `{{CaseNo}}`. WordPreview finds them, builds a form, and substitutes your
values straight into the original OpenXML — it never converts the document to HTML and
back, so styles, fonts, tables, headers/footers, margins, page size, and right-to-left
(Arabic/Hebrew) layout all stay exactly as authored.

## How it works

1. **Upload** a `.docx`. The server scans it for `{{placeholder}}` tokens using the
   Open XML SDK, correctly handling tokens that Word has split across multiple runs
   (e.g. `{{pro` + `ject_` + `name}}`).
2. **Fill** the auto-generated form. Fields whose names look long-form
   (`address`, `notes`, `description`, …) render as multi-line text areas.
3. **Preview** updates live — rendered from the *actual generated `.docx`* with
   [docx-preview], so what you see is exactly what you'll download.
4. **Download** the filled `.docx`.

Newlines in a field become real Word line breaks. Unknown tokens are left untouched.
The app is **stateless** — the original file stays in the browser and is sent with each
request; nothing is persisted on the server.

## Tech

- **.NET 10** / ASP.NET Core minimal API
- **DocumentFormat.OpenXml** for reading and rewriting the document
- Vanilla JS front end + [docx-preview] (vendored under `wwwroot/lib`)

## Run

```bash
cd src/WordPreview.Web
dotnet run
```

Then open the URL printed in the console (e.g. `http://localhost:5080`).

## Project layout

```
WordPreview.slnx
src/WordPreview.Web/
  Program.cs                       # endpoints: /api/analyze, /api/fill
  Services/DocxPlaceholderService.cs  # scan + fill engine (format-preserving)
  Models/ApiModels.cs
  wwwroot/                         # index.html, css, js, vendored libs
```

## API

| Method | Route          | Body                                   | Returns                        |
|--------|----------------|----------------------------------------|--------------------------------|
| POST   | `/api/analyze` | multipart `file`                       | `{ fileName, placeholders[] }` |
| POST   | `/api/fill`    | multipart `file` + `values` (JSON map) | the filled `.docx`             |

Max upload: 25 MB. Only `.docx` is accepted.

[docx-preview]: https://github.com/VolodymyrBaydalka/docxjs
