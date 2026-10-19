using System.Text;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Rendering;

public sealed class HtmlGenerator
{
    public string Generate(DocumentModel document)
    {
        var sb = new StringBuilder();

        foreach (var section in document.Sections)
        {
            sb.AppendLine("<section class=\"section\">");
            foreach (var block in section.Blocks)
            {
                sb.AppendLine(GenerateBlock(block));
            }

            sb.AppendLine("</section>");
        }

        return sb.ToString();
    }

    private static string GenerateBlock(Block block)
    {
        if (!string.IsNullOrWhiteSpace(block.Html))
            return block.Html;

        return block.Type switch
        {
            BlockType.Heading => $"<h2 class=\"heading\">{HtmlEncode(block.Text)}</h2>",
            BlockType.PageBreak => "<div class=\"page-break\"></div>",
            BlockType.Paragraph => $"<p class=\"paragraph\">{HtmlEncode(block.Text)}</p>",
            _ => $"<p class=\"paragraph\">{HtmlEncode(block.Text)}</p>"
        };
    }

    private static string HtmlEncode(string? value) => System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
}
