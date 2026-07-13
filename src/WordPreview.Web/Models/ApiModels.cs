namespace WordPreview.Web.Models;

/// <summary>Result of analyzing an uploaded .docx.</summary>
public sealed record AnalyzeResponse(string FileName, IReadOnlyList<string> Placeholders);
