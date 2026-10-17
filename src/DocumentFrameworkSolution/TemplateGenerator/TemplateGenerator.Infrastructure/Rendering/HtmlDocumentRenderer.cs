using System.Text;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Rendering;

public sealed class HtmlDocumentRenderer : IHtmlDocumentRenderer
{
    public string Render(DocumentModel document)
    {
        var sb = new StringBuilder();

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset=\"utf-8\" />");
        sb.AppendLine("  <style>");
        sb.AppendLine(PrintCss());
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("  <div class=\"document\">");

        foreach (var section in document.Sections)
        {
            sb.AppendLine("    <section class=\"section\">");
            foreach (var block in section.Blocks)
            {
                sb.AppendLine(RenderBlock(block));
            }
            sb.AppendLine("    </section>");
        }

        sb.AppendLine("  </div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }

    private static string RenderBlock(Block block)
    {
        if (!string.IsNullOrWhiteSpace(block.Html))
            return block.Html;

        return block.Type switch
        {
            BlockType.Heading => $"<h2>{System.Net.WebUtility.HtmlEncode(block.Text ?? string.Empty)}</h2>",
            BlockType.Paragraph => $"<p>{System.Net.WebUtility.HtmlEncode(block.Text ?? string.Empty)}</p>",
            BlockType.PageBreak => "<div class=\"page-break\"></div>",
            _ => $"<p>{System.Net.WebUtility.HtmlEncode(block.Text ?? string.Empty)}</p>"
        };
    }

    private static string PrintCss()
        => """
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
}

body {
  display: flex;
  justify-content: center;
}

.document {
  width: 100%;
  max-width: 210mm;
  padding: 0;
}

.section {
  page-break-inside: avoid;
}

p, h1, h2, h3, h4, h5, h6, table, ul, ol {
  break-inside: avoid;
}

.page-break {
  page-break-after: always;
  break-after: page;
  height: 0;
}

table {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #bbb;
}

th, td {
  border: 1px solid #bbb;
  padding: 6px 8px;
}
""";
}
