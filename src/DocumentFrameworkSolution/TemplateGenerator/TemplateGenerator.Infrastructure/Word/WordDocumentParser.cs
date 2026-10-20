using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Models;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Word;

public sealed class WordDocumentParser : IWordDocumentParser
{
    public async Task<DocumentModel> ParseAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, cancellationToken);
        memoryStream.Position = 0;

        using var document = WordprocessingDocument.Open(memoryStream, false);
        var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Document body not found.");

        var model = new DocumentModel();
        var section = new Section();

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case Paragraph paragraph:
                    section.Blocks.Add(ParseParagraph(paragraph));
                    break;
                case Table table:
                    section.Blocks.Add(ParseTable(table));
                    break;
                case Break breakElement when breakElement.Type == BreakValues.Page:
                    section.Blocks.Add(new Block { Type = BlockType.PageBreak, Html = "<div class=\"page-break\"></div>" });
                    break;
            }
        }

        model.Sections.Add(section);
        return model;
    }

    private static Block ParseParagraph(Paragraph paragraph)
    {
        var text = paragraph.InnerText.Trim();
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";

        var block = new Block
        {
            Type = styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) ? BlockType.Heading : BlockType.Paragraph,
            Text = text,
            Style = new StyleToken
            {
                Name = styleId,
                Alignment = paragraph.ParagraphProperties?.Justification?.Val?.Value.ToString(),
                MarginTopPt = ToPointValue(paragraph.ParagraphProperties?.SpacingBetweenLines?.Before?.Value),
                MarginBottomPt = ToPointValue(paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value)
            }
        };

        foreach (var run in paragraph.Elements<Run>())
        {
            var runText = run.InnerText;
            if (string.IsNullOrWhiteSpace(runText))
                continue;

            block.Runs.Add(new InlineRun
            {
                Text = runText,
                Bold = run.RunProperties?.Bold is not null,
                Italic = run.RunProperties?.Italic is not null,
                Underline = run.RunProperties?.Underline is not null,
                Color = run.RunProperties?.Color?.Val?.Value
            });
        }

        block.Html = BuildParagraphHtml(block);
        return block;
    }

    private static Block ParseTable(Table table)
    {
        return new Block
        {
            Type = BlockType.Table,
            Html = "<table class=\"document-table\"><tbody></tbody></table>"
        };
    }

    private static double? ToPointValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        return double.TryParse(value, out var result) ? result : null;
    }

    private static string BuildParagraphHtml(Block block)
    {
        if (block.Runs.Count == 0)
            return $"<p>{System.Net.WebUtility.HtmlEncode(block.Text ?? string.Empty)}</p>";

        var html = new System.Text.StringBuilder();
        html.Append("<p>");

        foreach (var run in block.Runs)
        {
            var text = System.Net.WebUtility.HtmlEncode(run.Text);

            if (run.Bold)
                text = $"<strong>{text}</strong>";

            if (run.Italic)
                text = $"<em>{text}</em>";

            if (run.Underline)
                text = $"<u>{text}</u>";

            html.Append($"<span>{text}</span>");
        }

        html.Append("</p>");
        return html.ToString();
    }

    Task<WordConversionResult> IWordDocumentParser.ParseAsync(Stream stream, CancellationToken cancellationToken)
    {
        throw new NotImplementedException();
    }
}
