using WordPreview.Web.Models;
using WordPreview.Web.Services;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddSingleton<DocxPlaceholderService>();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

const long MaxUploadBytes = 25 * 1024 * 1024; // 25 MB
const string DocxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

static bool IsDocx(IFormFile file) =>
    file.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);

// Analyze: return the distinct {{placeholders}} found in the uploaded document.
app.MapPost("/api/analyze", async (HttpRequest request, DocxPlaceholderService svc) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected a multipart form upload." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null || file.Length == 0)
        return Results.BadRequest(new { error = "No file uploaded." });
    if (!IsDocx(file))
        return Results.BadRequest(new { error = "Only .docx files are supported." });
    if (file.Length > MaxUploadBytes)
        return Results.BadRequest(new { error = "File exceeds the 25 MB limit." });

    try
    {
        await using var stream = file.OpenReadStream();
        var fields = svc.ExtractFields(stream);
        return Results.Ok(new AnalyzeResponse(file.FileName, fields));
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"Could not read the document: {ex.Message}" });
    }
});

// Fill: return a new .docx with placeholder values substituted in.
app.MapPost("/api/fill", async (HttpRequest request, DocxPlaceholderService svc) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected a multipart form upload." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null || file.Length == 0)
        return Results.BadRequest(new { error = "No file uploaded." });
    if (!IsDocx(file))
        return Results.BadRequest(new { error = "Only .docx files are supported." });
    if (file.Length > MaxUploadBytes)
        return Results.BadRequest(new { error = "File exceeds the 25 MB limit." });

    // Values arrive as a JSON object string in the "values" field.
    var values = new Dictionary<string, string>(StringComparer.Ordinal);
    var valuesJson = form["values"].ToString();
    if (!string.IsNullOrWhiteSpace(valuesJson))
    {
        try
        {
            var parsed = System.Text.Json.JsonSerializer
                .Deserialize<Dictionary<string, string>>(valuesJson);
            if (parsed is not null)
                foreach (var kv in parsed) values[kv.Key] = kv.Value ?? string.Empty;
        }
        catch (System.Text.Json.JsonException)
        {
            return Results.BadRequest(new { error = "Invalid values payload." });
        }
    }

    try
    {
        await using var stream = file.OpenReadStream();
        var bytes = svc.Fill(stream, values);
        var outName = BuildOutputName(file.FileName);
        return Results.File(bytes, DocxContentType, outName);
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"Could not generate the document: {ex.Message}" });
    }
});

app.Run();

static string BuildOutputName(string original)
{
    var name = Path.GetFileNameWithoutExtension(original);
    return $"{name}-filled.docx";
}
