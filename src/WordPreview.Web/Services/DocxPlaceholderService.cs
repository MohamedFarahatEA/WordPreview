using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordPreview.Web.Services;

/// <summary>
/// Scans and fills <c>{{placeholder}}</c> tokens inside a .docx while leaving every
/// other byte of formatting (styles, page size, margins, tables, headers/footers)
/// untouched. Works directly on the OpenXML — it never converts to HTML and back.
/// </summary>
public sealed partial class DocxPlaceholderService
{
    // {{ name }} — inner text is any chars except braces; spaces around are trimmed.
    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}", RegexOptions.Compiled)]
    private static partial Regex PlaceholderRegex();

    /// <summary>Returns the distinct placeholder keys in first-seen order.</summary>
    public IReadOnlyList<string> ExtractPlaceholders(Stream docxStream)
    {
        var seen = new List<string>();
        var set = new HashSet<string>(StringComparer.Ordinal);

        // Open read-only on a copy so the caller's stream stays reusable.
        using var ms = ToSeekableCopy(docxStream);
        using var doc = WordprocessingDocument.Open(ms, false);

        foreach (var para in EnumerateParagraphs(doc))
        {
            var full = GetParagraphText(para);
            if (full.Length == 0) continue;
            foreach (Match m in PlaceholderRegex().Matches(full))
            {
                var key = m.Groups[1].Value.Trim();
                if (key.Length > 0 && set.Add(key)) seen.Add(key);
            }
        }
        return seen;
    }

    /// <summary>
    /// Produces a new .docx with placeholders replaced by <paramref name="values"/>.
    /// Unknown tokens are left as-is. Newlines in a value become Word line breaks.
    /// </summary>
    public byte[] Fill(Stream docxStream, IReadOnlyDictionary<string, string> values)
    {
        var ms = ToSeekableCopy(docxStream); // becomes the output buffer
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            foreach (var para in EnumerateParagraphs(doc))
                ReplaceInParagraph(para, values);
        }
        return ms.ToArray();
    }

    // --- traversal ------------------------------------------------------------

    private static IEnumerable<Paragraph> EnumerateParagraphs(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) yield break;

        if (main.Document?.Body is { } body)
            foreach (var p in body.Descendants<Paragraph>()) yield return p;

        foreach (var header in main.HeaderParts)
            if (header.Header is { } h)
                foreach (var p in h.Descendants<Paragraph>()) yield return p;

        foreach (var footer in main.FooterParts)
            if (footer.Footer is { } f)
                foreach (var p in f.Descendants<Paragraph>()) yield return p;

        if (main.FootnotesPart?.Footnotes is { } fn)
            foreach (var p in fn.Descendants<Paragraph>()) yield return p;

        if (main.EndnotesPart?.Endnotes is { } en)
            foreach (var p in en.Descendants<Paragraph>()) yield return p;
    }

    private static string GetParagraphText(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var t in DirectTextNodes(para)) sb.Append(t.Text);
        return sb.ToString();
    }

    // Text nodes belonging to this paragraph, in document order. (A paragraph does
    // not nest other paragraphs in normal documents, so Descendants is safe.)
    private static List<Text> DirectTextNodes(Paragraph para) => para.Descendants<Text>().ToList();

    // --- replacement ----------------------------------------------------------

    private static void ReplaceInParagraph(Paragraph para, IReadOnlyDictionary<string, string> values)
    {
        var nodes = DirectTextNodes(para);
        if (nodes.Count == 0) return;

        var starts = new int[nodes.Count];
        var lens = new int[nodes.Count];
        var sb = new StringBuilder();
        for (int i = 0, pos = 0; i < nodes.Count; i++)
        {
            var text = nodes[i].Text ?? string.Empty;
            starts[i] = pos;
            lens[i] = text.Length;
            pos += text.Length;
            sb.Append(text);
        }

        var full = sb.ToString();
        var matches = PlaceholderRegex().Matches(full);
        if (matches.Count == 0) return;

        // Apply right-to-left so earlier matches keep valid global offsets.
        for (int mi = matches.Count - 1; mi >= 0; mi--)
        {
            var m = matches[mi];
            var key = m.Groups[1].Value.Trim();
            if (!values.TryGetValue(key, out var value)) continue; // leave unknown tokens
            ReplaceRange(nodes, starts, lens, m.Index, m.Index + m.Length, value ?? string.Empty);
        }
    }

    private static void ReplaceRange(
        List<Text> nodes, int[] starts, int[] lens,
        int rangeStart, int rangeEnd, string replacement)
    {
        for (int i = 0; i < nodes.Count; i++)
        {
            int nStart = starts[i];
            int nEnd = nStart + lens[i];
            if (nEnd <= rangeStart || nStart >= rangeEnd) continue; // no overlap

            var text = nodes[i].Text ?? string.Empty;
            int from = Math.Clamp(Math.Max(rangeStart, nStart) - nStart, 0, text.Length);
            int to = Math.Clamp(Math.Min(rangeEnd, nEnd) - nStart, 0, text.Length);
            string prefix = text[..from];
            string suffix = text[to..];

            bool ownsStart = nStart <= rangeStart && rangeStart < nEnd;
            if (!ownsStart)
            {
                // A middle/end node: drop the covered portion, keep the rest.
                nodes[i].Text = prefix + suffix;
                nodes[i].Space = SpaceProcessingModeValues.Preserve;
                continue;
            }

            // The node that owns the start of the placeholder receives the value.
            if (replacement.IndexOf('\n') < 0 && replacement.IndexOf('\r') < 0)
            {
                nodes[i].Text = prefix + replacement + suffix;
                nodes[i].Space = SpaceProcessingModeValues.Preserve;
            }
            else
            {
                WriteMultiline(nodes[i], prefix, replacement, suffix);
            }
        }
    }

    // Writes a value containing newlines as Text/Break siblings within the same run,
    // so line breaks render as real Word breaks while inheriting the run's formatting.
    private static void WriteMultiline(Text startNode, string prefix, string replacement, string suffix)
    {
        if (startNode.Parent is not Run run)
        {
            startNode.Text = prefix + replacement.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ') + suffix;
            return;
        }

        var lines = replacement.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        startNode.Text = prefix + lines[0];
        startNode.Space = SpaceProcessingModeValues.Preserve;

        OpenXmlElement anchor = startNode;
        for (int li = 1; li < lines.Length; li++)
        {
            var br = new Break();
            run.InsertAfter(br, anchor);
            anchor = br;

            var t = new Text(lines[li]) { Space = SpaceProcessingModeValues.Preserve };
            run.InsertAfter(t, anchor);
            anchor = t;
        }

        if (suffix.Length > 0)
        {
            var st = new Text(suffix) { Space = SpaceProcessingModeValues.Preserve };
            run.InsertAfter(st, anchor);
        }
    }

    // --- helpers --------------------------------------------------------------

    private static MemoryStream ToSeekableCopy(Stream source)
    {
        var ms = new MemoryStream();
        if (source.CanSeek) source.Position = 0;
        source.CopyTo(ms);
        ms.Position = 0;
        return ms;
    }
}
