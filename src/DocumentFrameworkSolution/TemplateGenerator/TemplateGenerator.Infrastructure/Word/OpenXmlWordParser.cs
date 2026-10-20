using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Models;

namespace TemplateGenerator.Infrastructure.Word;

public sealed class OpenXmlWordParser : IWordDocumentParser
{
    public async Task<WordConversionResult> ParseAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        if (!stream.CanRead)
            throw new InvalidOperationException("Input stream is not readable.");

        try
        {
            if (stream.CanSeek)
                stream.Position = 0;

            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream, cancellationToken).ConfigureAwait(false);
            memoryStream.Position = 0;

            using var document = WordprocessingDocument.Open(memoryStream, false);
            var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Word document body not found.");

            var html = new StringBuilder();
            var css = new StringBuilder();

            html.AppendLine("<div class=\"doc-root\">");

            foreach (var child in body.ChildElements)
            {
                html.AppendLine(ConvertElement(child));
            }

            html.AppendLine("</div>");

            css.AppendLine(GetDefaultCss());

            return new WordConversionResult
            {
                Html = html.ToString(),
                Css = css.ToString(),
                Messages = Array.Empty<string>()
            };
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to parse the Word document using OpenXML.", ex);
        }
    }

    private static string ConvertElement(OpenXmlElement element)
    {
        return element switch
        {
            Paragraph paragraph => ConvertParagraph(paragraph),
            Table table => ConvertTable(table),
            Break breakElement when breakElement.Type == BreakValues.Page => "<div class=\"page-break\"></div>",
            Break breakElement when breakElement.Type == BreakValues.TextWrapping => "<br />",
            _ => string.Empty
        };
    }

    private static string ConvertParagraph(Paragraph paragraph)
    {
        if (WordStyleResolver.IsListParagraph(paragraph))
            return ConvertListParagraph(paragraph);

        var tagName = WordStyleResolver.ResolveHeadingTag(paragraph);

        var paragraphStyle = WordStyleResolver.ResolveParagraphStyle(paragraph.ParagraphProperties);
        var cssInline = paragraphStyle.ToCssRule();

        var innerHtml = ConvertRuns(paragraph.Elements<Run>(), paragraph.ParagraphProperties);

        if (tagName == "p")
        {
            return string.IsNullOrWhiteSpace(cssInline)
                ? $"<p class=\"doc-paragraph\">{innerHtml}</p>"
                : $"<p class=\"doc-paragraph\" style=\"{cssInline}\">{innerHtml}</p>";
        }

        var headingCss = cssInline;
        return string.IsNullOrWhiteSpace(headingCss)
            ? $"<{tagName} class=\"doc-heading\">{innerHtml}</{tagName}>"
            : $"<{tagName} class=\"doc-heading\" style=\"{headingCss}\">{innerHtml}</{tagName}>";
    }

    private static string ConvertListParagraph(Paragraph paragraph)
    {
        var level = WordStyleResolver.GetListLevel(paragraph);
        var isOrdered = true; // can be made more precise with numbering definition lookup

        var listTag = isOrdered ? "ol" : "ul";
        var content = ConvertRuns(paragraph.Elements<Run>(), paragraph.ParagraphProperties);

        return $@"<{listTag} class=""doc-list doc-list-level-{level}"">
                    <li>{content}</li>
                 </{listTag}>";
    }

    private static string ConvertTable(Table table)
    {
        var sb = new StringBuilder();
        sb.Append("<table class=\"doc-table\" style=\"");
        sb.Append("width: 100%; border-collapse: collapse; border: 1px solid #d1d5db;\">");

        foreach (var row in table.Elements<TableRow>())
        {
            sb.Append("<tr>");

            foreach (var cell in row.Elements<TableCell>())
            {
                var cellContent = cell.InnerText.Trim();
                var cellHtml = ConvertCellContent(cell);
                sb.Append($"<td class=\"doc-cell\" style=\"border:1px solid #d1d5db; padding:6px 8px; vertical-align:top;\">{cellHtml}</td>");
            }

            sb.Append("</tr>");
        }

        sb.Append("</table>");
        return sb.ToString();
    }

    private static string ConvertCellContent(TableCell cell)
    {
        var content = new StringBuilder();

        foreach (var child in cell.ChildElements)
        {
            if (child is Paragraph paragraph)
                content.Append(ConvertParagraph(paragraph));
            else if (child is Table nestedTable)
                content.Append(ConvertTable(nestedTable));
            else if (child is Hyperlink hyperlink)
                content.Append(ConvertHyperlink(hyperlink));
            else
            {
                content.Append(child.InnerText);
            }
        }

        return content.ToString();
    }

    private static string ConvertRuns(IEnumerable<Run> runs, ParagraphProperties? paragraphProperties)
    {
        var sb = new StringBuilder();

        foreach (var run in runs)
        {
            var text = run.InnerText;

            if (string.IsNullOrWhiteSpace(text))
                continue;

            sb.Append(ConvertRun(run));
        }

        return sb.ToString();
    }

    private static string ConvertRun(Run run)
    {
        var runStyle = WordStyleResolver.ResolveRunStyle(run.RunProperties);

        var text = System.Net.WebUtility.HtmlEncode(run.InnerText);

        if (run.Elements<Hyperlink>().Any())
        {
            var hyperlink = run.Parent as Hyperlink;
            if (hyperlink is not null)
            {
                var href = hyperlink.Anchor ?? "#";
                return $"<a class=\"doc-link\" href=\"{href}\">{ApplyInlineFormatting(text, run.RunProperties)}</a>";
            }
        }

        return ApplyInlineFormatting(text, run.RunProperties);
    }

    private static string ApplyInlineFormatting(string text, RunProperties? runProperties)
    {
        if (runProperties is null)
            return text;

        var result = text;

        if (runProperties.Bold is not null)
            result = $"<strong>{result}</strong>";

        if (runProperties.Italic is not null)
            result = $"<em>{result}</em>";

        if (runProperties.Underline is not null && runProperties.Underline.Val != UnderlineValues.None)
            result = $"<u>{result}</u>";

        if (runProperties.Color is not null && !string.IsNullOrWhiteSpace(runProperties.Color.Val?.Value))
            result = $"<span style=\"color: #{runProperties.Color.Val.Value};\">{result}</span>";

        return result;
    }

    private static string ConvertHyperlink(Hyperlink hyperlink)
    {
        var text = System.Net.WebUtility.HtmlEncode(hyperlink.InnerText.Trim());
        var anchor = hyperlink.Anchor ?? "#";

        return $"<a class=\"doc-link\" href=\"{anchor}\">{text}</a>";
    }

    private static string GetDefaultCss()
    {
        return """
.doc-root {
  font-family: "Segoe UI", Arial, sans-serif;
  color: #111;
  line-height: 1.5;
  background: #fff;
}

.doc-paragraph {
  margin: 0 0 0.75rem 0;
  white-space: normal;
}

.doc-heading {
  margin: 0.75rem 0 0.5rem 0;
  font-weight: 700;
  color: #111;
}

h1.doc-heading { font-size: 2rem; }
h2.doc-heading { font-size: 1.6rem; }
h3.doc-heading { font-size: 1.35rem; }
h4.doc-heading { font-size: 1.2rem; }
h5.doc-heading { font-size: 1.05rem; }
h6.doc-heading { font-size: 0.95rem; }

.doc-table {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #d1d5db;
  margin-top: 0.75rem;
  margin-bottom: 0.75rem;
}

.doc-cell {
  border: 1px solid #d1d5db;
  padding: 6px 8px;
  vertical-align: top;
}

.doc-list {
  margin: 0.25rem 0 0.75rem 1.5rem;
  padding-left: 0.75rem;
}

.doc-link {
  color: #1d4ed8;
  text-decoration: underline;
}

.page-break {
  break-after: page;
  page-break-after: always;
  height: 0;
}
""";
    }
}