using System.Text;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Models;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Rendering;

public sealed class HtmlDocumentRenderer : IHtmlDocumentRenderer
{
    public string Render(WordConversionResult result)
    {
        if (result is null)
            throw new ArgumentNullException(nameof(result));

        return Render(result.Html, result.Css);
    }

    public string Render(DocumentModel document)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));

        var htmlBuilder = new StringBuilder();

        foreach (var section in document.Sections)
        {
            htmlBuilder.AppendLine("<section class=\"doc-section\">");
            foreach (var block in section.Blocks)
            {
                if (!string.IsNullOrWhiteSpace(block.Html))
                {
                    htmlBuilder.AppendLine(block.Html);
                    continue;
                }

                htmlBuilder.AppendLine(RenderBlock(block));
            }
            htmlBuilder.AppendLine("</section>");
        }

        var html = htmlBuilder.ToString();
        var css = GetPrintCss();

        return BuildDocumentHtml(html, css);
    }

    public string Render(string html, string css)
    {
        if (string.IsNullOrWhiteSpace(html))
            return BuildDocumentHtml(string.Empty, css);

        var effectiveCss = string.IsNullOrWhiteSpace(css)
            ? GetPrintCss()
            : css;

        return BuildDocumentHtml(html, effectiveCss);
    }

    private static string RenderBlock(Block block)
    {
        return block.Type switch
        {
            BlockType.Heading => $"<h2 class=\"doc-heading\">{HtmlEncode(block.Text)}</h2>",
            BlockType.Paragraph => $"<p class=\"doc-paragraph\">{HtmlEncode(block.Text)}</p>",
            BlockType.PageBreak => "<div class=\"page-break\"></div>",
            _ => $"<p class=\"doc-paragraph\">{HtmlEncode(block.Text)}</p>"
        };
    }

    private static string BuildDocumentHtml(string bodyHtml, string css)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset=\"utf-8\" />");
        sb.AppendLine("  <style>");
        sb.AppendLine(css);
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("  <div class=\"document-shell\">");
        sb.AppendLine(bodyHtml);
        sb.AppendLine("  </div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }

    private static string HtmlEncode(string? value)
        => System.Net.WebUtility.HtmlEncode(value ?? string.Empty);

    private static string GetPrintCss()
    {
        return """
@page {
  size: A4 portrait;
  margin: 12mm 12mm 12mm 12mm;
}

html, body {
  margin: 0;
  padding: 0;
  background: white;
  color: #111;
  font-family: "Segoe UI", Arial, sans-serif;
  line-height: 1.5;
}

body {
  display: flex;
  justify-content: center;
}

.document-shell {
  width: 100%;
  max-width: 210mm;
  padding: 0;
}

.doc-section {
  page-break-inside: avoid;
}

.doc-paragraph,
.doc-heading,
.doc-list,
.doc-table,
.doc-image,
.doc-link {
  break-inside: avoid;
}

.doc-heading {
  margin: 0.75rem 0 0.5rem 0;
  font-weight: 700;
  color: #111;
}

h1.doc-heading { font-size: 2rem; }
h2.doc-heading { font-size: 1.6rem; }
h3.doc-heading { font-size: 1.35rem; }

.doc-table {
  width: 100%;
  border-collapse: collapse;
  margin: 0.75rem 0;
}

.doc-table th,
.doc-table td {
  border: 1px solid #d1d5db;
  padding: 6px 8px;
  vertical-align: top;
  text-align: left;
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
