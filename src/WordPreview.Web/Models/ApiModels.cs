namespace WordPreview.Web.Models;

/// <summary>
/// A field the user fills in. <see cref="Type"/> is one of
/// "text", "multiline", "date", or "select".
/// </summary>
public sealed record PlaceholderField(
    string Key,
    string Type,
    string? Format,
    IReadOnlyList<string>? Options);

/// <summary>Result of analyzing an uploaded .docx.</summary>
public sealed record AnalyzeResponse(string FileName, IReadOnlyList<PlaceholderField> Fields);
