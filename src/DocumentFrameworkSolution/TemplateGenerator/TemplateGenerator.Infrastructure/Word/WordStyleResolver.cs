using DocumentFormat.OpenXml.Wordprocessing;

namespace TemplateGenerator.Infrastructure.Word;

public static class WordStyleResolver
{
    public static CssStyleToken ResolveParagraphStyle(ParagraphProperties? paragraphProperties)
    {
        var style = new CssStyleToken
        {
            ClassName = "doc-paragraph"
        };

        if (paragraphProperties is null)
            return style;

        var alignment = paragraphProperties.Justification?.Val?.Value;
        if (alignment is not null)
        {
            string textAlign;

            if (alignment.Value == JustificationValues.Center)
            {
                textAlign = "center";
            }
            else if (alignment.Value == JustificationValues.Right)
            {
                textAlign = "right";
            }    
            else if (alignment.Value == JustificationValues.Both)
            {
                textAlign = "justify";
            }
            else
            {
                textAlign = "left";
            }

            style.Declarations["text-align"] = textAlign;
        }

        var spacing = paragraphProperties.SpacingBetweenLines;
        if (spacing is not null)
        {
            if (spacing.Before is not null && int.TryParse(spacing.Before.Value, out var beforePts))
                style.Declarations["margin-top"] = $"{beforePts / 20.0:F2}pt";

            if (spacing.After is not null && int.TryParse(spacing.After.Value, out var afterPts))
                style.Declarations["margin-bottom"] = $"{afterPts / 20.0:F2}pt";

            if (spacing.Line is not null && spacing.Line.Value is not null)
            {
                var lineValue = spacing.Line.Value;
                style.Declarations["line-height"] = lineValue switch
                {
                    "240" => "1.2",
                    "360" => "1.5",
                    "480" => "2",
                    _ => "normal"
                };
            }
        }

        var indent = paragraphProperties.Indentation;
        if (indent is not null)
        {
            if (indent.Left is not null)
                style.Declarations["margin-left"] = $"{(double.Parse(indent.Left.Value??"0") / 20.0):F2}pt";

            if (indent.Right is not null)
                style.Declarations["margin-right"] = $"{(double.Parse(indent.Right.Value??"0") / 20.0):F2}pt";
        }

        return style;
    }

    public static CssStyleToken ResolveRunStyle(RunProperties? runProperties)
    {
        var style = new CssStyleToken
        {
            ClassName = "doc-run"
        };

        if (runProperties is null)
            return style;

        if (runProperties.Bold is not null)
            style.Declarations["font-weight"] = "700";

        if (runProperties.Italic is not null)
            style.Declarations["font-style"] = "italic";

        if (runProperties.Underline is not null && runProperties.Underline.Val != UnderlineValues.None)
            style.Declarations["text-decoration"] = "underline";

        if (runProperties.Color is not null && !string.IsNullOrWhiteSpace(runProperties.Color.Val?.Value))
            style.Declarations["color"] = $"#{runProperties.Color.Val.Value}";

        if (runProperties.FontSize is not null && !string.IsNullOrWhiteSpace(runProperties.FontSize.Val?.Value))
        {
            if (double.TryParse(runProperties.FontSize.Val.Value, out var fontSizeValue))
                style.Declarations["font-size"] = $"{fontSizeValue / 2.0:F2}pt";
        }

        //if (runProperties.FontHint is not null)
        //    style.Declarations["font-family"] = "\"Segoe UI\", sans-serif";

        return style;
    }

    public static string ResolveHeadingTag(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;

        return styleId.ToLowerInvariant() switch
        {
            "heading1" => "h1",
            "heading2" => "h2",
            "heading3" => "h3",
            "heading4" => "h4",
            "heading5" => "h5",
            "heading6" => "h6",
            _ => "p"
        };
    }

    public static bool IsHeadingLike(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
    }

    public static string ResolveTableStyle(Table table)
    {
        var style = new CssStyleToken
        {
            ClassName = "doc-table"
        };

        style.Declarations["width"] = "100%";
        style.Declarations["border-collapse"] = "collapse";
        style.Declarations["border"] = "1px solid #d1d5db";
        style.Declarations["margin-top"] = "0.75rem";
        style.Declarations["margin-bottom"] = "0.75rem";

        return style.ToCssRule();
    }

    public static int GetListLevel(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr is null)
            return 0;

        var ilvl = numPr.NumberingLevelReference?.Val?.Value;
        if (ilvl is not null)
            return ilvl.Value;

        var levelVal = numPr.NumberingId is not null ? 0 : 0;

        return 0;
    }

    public static bool IsListParagraph(Paragraph paragraph)
    {
        return paragraph.ParagraphProperties?.NumberingProperties is not null;
    }
}
