using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Rendering;

public sealed class StyleNormalizer
{
    public DocumentModel Normalize(DocumentModel document)
    {
        foreach (var section in document.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block.Style is null)
                    block.Style = new StyleToken();

                if (block.Type == BlockType.Heading)
                {
                    block.Style.FontSizePt = Math.Max(block.Style.FontSizePt, 18d);
                    block.Style.Bold = true;
                    block.Style.MarginBottomPt ??= 8d;
                }

                if (block.Type == BlockType.Paragraph)
                {
                    block.Style.FontSizePt = Math.Max(block.Style.FontSizePt, 11d);
                    block.Style.MarginBottomPt ??= 6d;
                }

                if (string.IsNullOrWhiteSpace(block.Style.Alignment))
                    block.Style.Alignment = "left";
            }
        }

        return document;
    }
}
