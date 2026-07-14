using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordPreview.Web.Models;

namespace WordPreview.Web.Services;

/// <summary>
/// Scans and fills placeholder tokens inside a .docx while leaving every other byte
/// of formatting (styles, page size, margins, tables, headers/footers) untouched.
/// Works directly on the OpenXML — it never converts to HTML and back.
///
/// Token syntax:
///   {{ key }}                        text (or inferred date/multiline from the name)
///   {{ key | text }}                 force plain text
///   {{ key | date }}                 date picker, output yyyy-MM-dd
///   {{ key | date : dd/MM/yyyy }}    date picker, custom output format
///   {{ key | select : A, B, C }}     dropdown with the given options
///   {{ key | multiline }}            multi-line text area
///
/// Inserted values are also tagged with a run direction (w:rtl) inferred from their
/// script, so Arabic values render right-to-left and Latin values left-to-right no
/// matter which direction the surrounding paragraph uses.
/// </summary>
public sealed partial class DocxPlaceholderService
{
    // {{ ...anything but braces... }} — inner text parsed by ParseSpec.
    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}", RegexOptions.Compiled)]
    private static partial Regex PlaceholderRegex();

    private const string DefaultDateFormat = "yyyy-MM-dd";

    private enum FieldType { Text, Multiline, Date, Select }

    private sealed record Spec(string Key, FieldType Type, string? Format, List<string>? Options);

    /// <summary>Returns the distinct fields in first-seen order.</summary>
    public IReadOnlyList<PlaceholderField> ExtractFields(Stream docxStream)
    {
        var order = new List<string>();
        var byKey = new Dictionary<string, Spec>(StringComparer.Ordinal);

        using var ms = ToSeekableCopy(docxStream);
        using var doc = WordprocessingDocument.Open(ms, false);

        foreach (var para in EnumerateParagraphs(doc))
        {
            var full = GetParagraphText(para);
            if (full.Length == 0) continue;
            foreach (Match m in PlaceholderRegex().Matches(full))
            {
                var spec = ParseSpec(m.Groups[1].Value);
                if (spec.Key.Length == 0 || byKey.ContainsKey(spec.Key)) continue;
                byKey[spec.Key] = spec;
                order.Add(spec.Key);
            }
        }

        return order.Select(k =>
        {
            var s = byKey[k];
            return new PlaceholderField(s.Key, s.Type.ToString().ToLowerInvariant(), s.Format, s.Options);
        }).ToList();
    }

    /// <summary>
    /// Produces a new .docx with placeholders replaced by <paramref name="values"/>.
    /// Values are formatted per each token's spec (e.g. dates). Unknown tokens are
    /// left as-is; newlines in a value become Word line breaks.
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

    // --- spec parsing ---------------------------------------------------------

    private static Spec ParseSpec(string inner)
    {
        int pipe = inner.IndexOf('|');
        if (pipe < 0)
        {
            var key = inner.Trim();
            // Format left null = "author specified none"; FormatValue picks the
            // default (LTR yyyy-MM-dd, or day-first dd-MM-yyyy in Arabic context).
            return new Spec(key, InferType(key), null, null);
        }

        var name = inner[..pipe].Trim();
        var rest = inner[(pipe + 1)..].Trim();

        string typeToken = rest;
        string arg = string.Empty;
        int colon = rest.IndexOf(':');
        if (colon >= 0)
        {
            typeToken = rest[..colon].Trim();
            arg = rest[(colon + 1)..].Trim();
        }

        switch (typeToken.ToLowerInvariant())
        {
            case "date":
                return new Spec(name, FieldType.Date,
                    string.IsNullOrWhiteSpace(arg) ? null : arg, null);
            case "select":
            case "dropdown":
            case "choice":
                var opts = arg.Split(',')
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();
                return new Spec(name, FieldType.Select, null, opts);
            case "multiline":
            case "textarea":
                return new Spec(name, FieldType.Multiline, null, null);
            case "text":
                return new Spec(name, FieldType.Text, null, null);
            default:
                return new Spec(name, InferType(name), null, null);
        }
    }

    private static FieldType InferType(string key)
    {
        if (Regex.IsMatch(key, "date", RegexOptions.IgnoreCase)) return FieldType.Date;
        if (Regex.IsMatch(key, "address|notes?|description|comment|details|body", RegexOptions.IgnoreCase))
            return FieldType.Multiline;
        return FieldType.Text;
    }

    // Arabic day-first default (dd-MM-yyyy) so RTL dates read day → month → year
    // left-to-right (day on the left, year on the right).
    private const string DefaultRtlDateFormat = "dd-MM-yyyy";

    private static string FormatValue(Spec spec, string raw, bool rtlContext)
    {
        if (spec.Type == FieldType.Date && !string.IsNullOrWhiteSpace(raw))
        {
            // Client sends ISO (yyyy-MM-dd) from the date picker; format per the spec.
            if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                var fmt = spec.Format ?? DefaultDateFormat;
                // In an Arabic/RTL paragraph, a plain date (no explicit format) uses
                // the day-first order and Arabic-Indic digits. An explicit template
                // format is always honored as-is.
                bool useRtlDefault = rtlContext && spec.Format is null;
                if (useRtlDefault) fmt = DefaultRtlDateFormat;

                var formatted = dt.ToString(fmt, CultureInfo.InvariantCulture);
                return rtlContext ? ToArabicIndicDigits(formatted) : formatted;
            }
        }
        return raw;
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

        // Is this an Arabic/RTL paragraph? Dates in RTL context use Arabic-Indic digits.
        bool rtlContext = ParagraphIsRtl(para, full);

        // Apply right-to-left so earlier matches keep valid global offsets.
        for (int mi = matches.Count - 1; mi >= 0; mi--)
        {
            var m = matches[mi];
            var spec = ParseSpec(m.Groups[1].Value);
            if (!values.TryGetValue(spec.Key, out var raw)) continue; // leave unknown tokens
            var value = FormatValue(spec, raw ?? string.Empty, rtlContext);
            // Dates keep the paragraph's direction (a number renders LTR internally);
            // other fields get a direction tag from their own script.
            ReplaceRange(nodes, starts, lens, m.Index, m.Index + m.Length, value,
                tagDirection: spec.Type != FieldType.Date);
        }
    }

    private static bool ParagraphIsRtl(Paragraph para, string text)
    {
        var bidi = para.ParagraphProperties?.GetFirstChild<BiDi>();
        if (bidi is not null)
            return bidi.Val is null || bidi.Val.Value;
        return DetectDir(text) == TextDir.Rtl;
    }

    // 0-9 -> Arabic-Indic ٠-٩ (U+0660..U+0669). Other characters (/, -, .) unchanged.
    private static string ToArabicIndicDigits(string s)
    {
        char[]? buf = null;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] >= '0' && s[i] <= '9')
            {
                buf ??= s.ToCharArray();
                buf[i] = (char)(0x0660 + (s[i] - '0'));
            }
        }
        return buf is null ? s : new string(buf);
    }

    private static void ReplaceRange(
        List<Text> nodes, int[] starts, int[] lens,
        int rangeStart, int rangeEnd, string replacement, bool tagDirection)
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
                nodes[i].Text = prefix + suffix;
                nodes[i].Space = SpaceProcessingModeValues.Preserve;
                continue;
            }

            if (replacement.IndexOf('\n') < 0 && replacement.IndexOf('\r') < 0)
            {
                nodes[i].Text = prefix + replacement + suffix;
                nodes[i].Space = SpaceProcessingModeValues.Preserve;
            }
            else
            {
                WriteMultiline(nodes[i], prefix, replacement, suffix);
            }

            // When the placeholder occupies this run by itself, tag the run's
            // direction from the value's script so Arabic values render RTL and
            // Latin values render LTR — regardless of the paragraph's direction.
            if (tagDirection && prefix.Length == 0 && suffix.Length == 0 && nodes[i].Parent is Run valueRun)
                ApplyDirection(valueRun, replacement);
        }
    }

    private enum TextDir { Neutral, Ltr, Rtl }

    // Dominant script of a string: RTL (Arabic/Hebrew), LTR (Latin), or Neutral
    // (only digits/punctuation/whitespace — keep the inherited direction).
    private static TextDir DetectDir(string s)
    {
        int rtl = 0, ltr = 0;
        foreach (var ch in s)
        {
            int c = ch;
            // Hebrew + Arabic blocks (U+0590–U+08FF), Arabic presentation forms
            // A (U+FB50–U+FDFF) and B (U+FE70–U+FEFF).
            if ((c >= 0x0590 && c <= 0x08FF) ||
                (c >= 0xFB50 && c <= 0xFDFF) ||
                (c >= 0xFE70 && c <= 0xFEFF))
                rtl++;
            // Latin letters (Basic + Latin-1 Supplement + Extended-A/B, U+00C0–U+024F).
            else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                     (c >= 0x00C0 && c <= 0x024F))
                ltr++;
        }
        if (rtl == 0 && ltr == 0) return TextDir.Neutral;
        return rtl > ltr ? TextDir.Rtl : TextDir.Ltr;
    }

    private static void ApplyDirection(Run run, string value)
    {
        var dir = DetectDir(value);
        if (dir == TextDir.Neutral) return; // dates/numbers: inherit paragraph direction

        var rpr = run.GetFirstChild<RunProperties>();
        if (rpr is null) { rpr = new RunProperties(); run.PrependChild(rpr); }

        var rtl = rpr.GetFirstChild<RightToLeftText>();
        if (rtl is null) { rtl = new RightToLeftText(); rpr.AppendChild(rtl); }
        rtl.Val = OnOffValue.FromBoolean(dir == TextDir.Rtl);
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
