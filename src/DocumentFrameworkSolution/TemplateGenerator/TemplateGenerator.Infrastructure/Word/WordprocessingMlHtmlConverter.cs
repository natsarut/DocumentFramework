using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using TemplateGenerator.Application.Models;

namespace TemplateGenerator.Infrastructure.Word;

internal sealed class WordprocessingMlHtmlConverter
{
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const long EmusPerPixel = 9525;
    private readonly MainDocumentPart _mainPart;
    private OpenXmlPart _activePart;
    private string? _activeParagraphStyleId;
    private readonly List<string> _messages = [];

    private WordprocessingMlHtmlConverter(MainDocumentPart mainPart)
    {
        _mainPart = mainPart;
        _activePart = mainPart;
    }

    public static WordConversionResult Convert(MainDocumentPart mainPart, CancellationToken cancellationToken)
    {
        var converter = new WordprocessingMlHtmlConverter(mainPart);
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Word document body not found.");
        var html = new StringBuilder();
        var css = new StringBuilder(DefaultDocumentCss);

        html.Append("<div class=\"doc-root\"");
        if (converter.ResolveDocumentDefaultFontFamily() is { } defaultFontFamily)
            html.Append(" style=\"font-family:").Append(Attr(defaultFontFamily)).Append('"');
        html.AppendLine(">");
        converter.AppendHeadersAndFooters(html, body, "headerReference", "header", cancellationToken);
        converter.AppendDocumentSections(html, css, body, cancellationToken);
        converter.AppendNotes(html, cancellationToken);
        converter.AppendHeadersAndFooters(html, body, "footerReference", "footer", cancellationToken);
        html.AppendLine("</div>");

        return new WordConversionResult
        {
            Html = html.ToString(),
            Css = css.ToString(),
            Messages = converter._messages.AsReadOnly()
        };
    }

    private void AppendBlocks(StringBuilder output, IEnumerable<OpenXmlElement> elements, CancellationToken cancellationToken)
    {
        var blockElements = elements.Where(IsBlockElement).ToArray();
        for (var index = 0; index < blockElements.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var element = blockElements[index];
            if (NameIs(element, "p") && GetNumberingId(element) is not null)
            {
                var listId = GetNumberingId(element);
                var level = GetListLevel(element);
                var isOrdered = IsOrderedList(listId!, level);
                var tag = isOrdered ? "ol" : "ul";
                output.Append('<').Append(tag).Append(" class=\"doc-list doc-list-level-").Append(level).Append('"');
                var listStyle = new Dictionary<string, string>(StringComparer.Ordinal);

                if (isOrdered)
                {
                    listStyle["list-style-type"] = GetOrderedListStyleType(listId!, level);
                }
                else
                {
                    listStyle["list-style-type"] = GetUnorderedListStyleType(listId!, level);
                }
                    
                if (ResolveNumberingMarkerFontFamily(element, listId!, level) is { } markerFont)
                    listStyle["--doc-marker-font"] = QuoteCssFontFamily(markerFont);
                output.Append(StyleAttribute(listStyle));
                output.Append('>');
                while (index < blockElements.Length && NameIs(blockElements[index], "p") &&
                       GetNumberingId(blockElements[index]) == listId && GetListLevel(blockElements[index]) == level)
                {
                    Dictionary<string, string> styleList = GetParagraphStyle(blockElements[index]);

                    // Remove unnecessary style.
                    styleList.Remove("margin-left");

                    output.Append("<li").Append(StyleAttribute(styleList)).Append('>');
                    output.Append(ConvertParagraphContent(blockElements[index], cancellationToken));
                    output.Append("</li>");
                    index++;
                    cancellationToken.ThrowIfCancellationRequested();
                }
                index--;
                output.Append("</").Append(tag).AppendLine(">");
                continue;
            }

            output.AppendLine(ConvertBlock(element, cancellationToken));
        }
    }

    private void AppendDocumentSections(StringBuilder output, StringBuilder css, OpenXmlElement body, CancellationToken cancellationToken)
    {
        var elements = new List<OpenXmlElement>();
        var sectionNumber = 0;
        var startNextSectionOnNewPage = false;

        void AppendSection(OpenXmlElement? sectionProperties, bool hasNextSection)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sectionNumber++;
            var pageName = "doc-section-" + sectionNumber.ToString(CultureInfo.InvariantCulture);
            if (startNextSectionOnNewPage)
                output.AppendLine("<div class=\"page-break\"></div>");
            startNextSectionOnNewPage = false;
            output.Append("<section class=\"doc-section ").Append(pageName).Append("\" style=\"page:").Append(pageName).AppendLine(";\">");
            AppendBlocks(output, elements, cancellationToken);
            output.AppendLine("</section>");
            AppendPageRules(css, sectionProperties, pageName);

            var breakType = GetAttribute(FindChild(sectionProperties, "type"), "val");
            startNextSectionOnNewPage = hasNextSection && breakType is null or "nextPage" or "oddPage" or "evenPage";
        }

        foreach (var child in body.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (NameIs(child, "sectPr"))
            {
                AppendSection(child, false);
                elements.Clear();
                continue;
            }

            elements.Add(child);
            var paragraphSectionProperties = NameIs(child, "p")
                ? FindChild(FindChild(child, "pPr"), "sectPr")
                : null;
            if (paragraphSectionProperties is not null)
            {
                AppendSection(paragraphSectionProperties, true);
                elements.Clear();
            }
        }

        if (elements.Count > 0 || sectionNumber == 0)
            AppendSection(FindChild(body, "sectPr"), false);
    }

    private string ConvertBlock(OpenXmlElement element, CancellationToken cancellationToken)
    {
        if (NameIs(element, "p"))
            return ConvertParagraph(element, cancellationToken);
        if (NameIs(element, "tbl"))
            return ConvertTable(element, cancellationToken);
        if (NameIs(element, "lastRenderedPageBreak"))
            return string.Empty;
        if (NameIs(element, "sdt"))
        {
            var content = FindChild(element, "sdtContent");
            return content is null ? string.Empty : ConvertBlocksToString(content.ChildElements, cancellationToken);
        }
        if (NameIs(element, "customXml") || NameIs(element, "altChunk"))
        {
            if (NameIs(element, "altChunk"))
            {
                _messages.Add("An alternative-format chunk was omitted because its external content cannot be safely converted.");
                return string.Empty;
            }
            return ConvertBlocksToString(element.ChildElements, cancellationToken);
        }
        return string.Empty;
    }

    private string ConvertBlocksToString(IEnumerable<OpenXmlElement> elements, CancellationToken cancellationToken)
    {
        var output = new StringBuilder();
        AppendBlocks(output, elements, cancellationToken);
        return output.ToString();
    }

    private string ConvertParagraph(OpenXmlElement paragraph, CancellationToken cancellationToken)
    {
        var styleId = GetAttribute(FindChild(paragraph, "pPr") is { } properties ? FindChild(properties, "pStyle") : null, "val") ?? string.Empty;
        var tag = styleId.ToLowerInvariant() switch
        {
            "heading1" => "h1", "heading2" => "h2", "heading3" => "h3",
            "heading4" => "h4", "heading5" => "h5", "heading6" => "h6",
            "title" => "h1", "subtitle" => "h2", "quote" or "intensequote" => "blockquote",
            _ => "p"
        };
        var paragraphStyle = GetParagraphStyle(paragraph);
        if (styleId.ToLowerInvariant() is "heading1" or "heading2" or "heading3" or "heading4" or "heading5" or "heading6")
            paragraphStyle["font-weight"] = ResolveHeadingFontWeight(styleId);
        return $"<{tag} class=\"{(tag == "p" ? "doc-paragraph" : "doc-heading") }\"{StyleAttribute(paragraphStyle)}>{ConvertParagraphContent(paragraph, cancellationToken)}</{tag}>";
    }

    private string ResolveHeadingFontWeight(string styleId)
    {
        var bold = GetStyleProperties(styleId, "rPr")
            .Select(properties => FindChild(properties, "b"))
            .FirstOrDefault(property => property is not null)
            ?? FindChild(GetDocumentDefaultRunProperties(), "b");
        return IsOn(bold) ? "700" : "400";
    }

    private string ConvertParagraphContent(OpenXmlElement paragraph, CancellationToken cancellationToken)
    {
        var output = new StringBuilder();
        var field = new FieldState();
        var previousParagraphStyleId = _activeParagraphStyleId;
        _activeParagraphStyleId = GetParagraphStyleId(paragraph);
        try
        {
            foreach (var child in paragraph.ChildElements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (NameIs(child, "pPr"))
                    continue;
                output.Append(ConvertInline(child, field, cancellationToken));
            }
            if (field.Active)
                output.Append(RenderField(field.Instruction.ToString(), field.Result.ToString()));
            return output.ToString();
        }
        finally
        {
            _activeParagraphStyleId = previousParagraphStyleId;
        }
    }

    private string ConvertInline(OpenXmlElement element, FieldState field, CancellationToken cancellationToken)
    {
        var localName = element.LocalName;
        if (localName == "r")
            return ConvertRun(element, field, cancellationToken);
        if (localName is "hyperlink" or "smartTag" or "sdt" or "customXml")
        {
            if (localName == "hyperlink")
            {
                var content = ConvertChildrenInline(element, field, cancellationToken);
                var href = GetAttribute(element, "anchor") is { Length: > 0 } anchor
                    ? "#" + anchor
                    : ResolveHyperlink(GetAttribute(element, "id"));
                return href is null ? content : $"<a class=\"doc-link\" href=\"{Attr(href)}\">{content}</a>";
            }
            if (localName == "sdt")
            {
                var content = FindChild(element, "sdtContent");
                return content is null ? string.Empty : ConvertChildrenInline(content, field, cancellationToken);
            }
            return ConvertChildrenInline(element, field, cancellationToken);
        }
        if (localName is "fldSimple")
        {
            var instruction = GetAttribute(element, "instr") ?? string.Empty;
            var result = ConvertChildrenInline(element, field, cancellationToken);
            return RenderField(instruction, result);
        }
        if (localName == "t" || localName == "delText" || localName == "instrText")
        {
            var rawText = element.InnerText;
            var text = System.Net.WebUtility.HtmlEncode(rawText);
            if (!field.Active && (GetAttribute(element, "space", "http://www.w3.org/XML/1998/namespace") == "preserve" || rawText.Contains("  ", StringComparison.Ordinal)))
                text = "<span class=\"doc-preserve\">" + text + "</span>";
            if (field.Active)
            {
                if (field.InInstruction && localName == "instrText")
                    field.Instruction.Append(element.InnerText);
                else if (!field.InInstruction)
                    field.Result.Append(text);
                return string.Empty;
            }
            return localName == "instrText" ? string.Empty : text;
        }
        if (localName == "fldChar")
        {
            var type = GetAttribute(element, "fldCharType") ?? GetAttribute(element, "type");
            if (type == "begin")
            {
                field.Reset();
                field.Active = true;
                field.InInstruction = true;
            }
            else if (type == "separate")
                field.InInstruction = false;
            else if (type == "end" && field.Active)
            {
                var result = RenderField(field.Instruction.ToString(), field.Result.ToString());
                field.Reset();
                return result;
            }
            return string.Empty;
        }
        if (localName == "br")
        {
            var type = GetAttribute(element, "type");
            return type == "page" ? "<span class=\"page-break\"></span>" : "<br />";
        }
        if (localName == "lastRenderedPageBreak")
            return string.Empty;
        if (localName is "cr")
            return "<br />";
        if (localName == "tab")
            return "<span class=\"doc-tab\">&#9;</span>";
        if (localName == "drawing" || localName == "pict")
            return ConvertDrawing(element, cancellationToken);
        if (localName is "footnoteReference" or "endnoteReference")
        {
            var noteKind = localName == "footnoteReference" ? "footnote" : "endnote";
            var id = GetAttribute(element, "id") ?? string.Empty;
            return $"<sup class=\"doc-note-ref\"><a href=\"#{noteKind}-{Attr(id)}\">{Attr(id)}</a></sup>";
        }
        if (localName == "sym")
        {
            var hex = GetAttribute(element, "char");
            return int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var codePoint) &&
                   Rune.IsValid(codePoint) ? System.Net.WebUtility.HtmlEncode(new Rune(codePoint).ToString()) : string.Empty;
        }
        if (localName is "oMath" or "oMathPara")
            return $"<span class=\"doc-equation\">{System.Net.WebUtility.HtmlEncode(element.InnerText)}</span>";
        if (localName == "noBreakHyphen") return "&#8209;";
        if (localName == "softHyphen") return "&shy;";
        return string.Empty;
    }

    private string ConvertChildrenInline(OpenXmlElement parent, FieldState field, CancellationToken cancellationToken)
    {
        var output = new StringBuilder();
        foreach (var child in parent.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (NameIs(child, "rPr") || NameIs(child, "pPr"))
                continue;
            if (NameIs(child, "r"))
                output.Append(ConvertRun(child, field, cancellationToken));
            else
                output.Append(ConvertInline(child, field, cancellationToken));
        }
        return output.ToString();
    }

    private string ConvertRun(OpenXmlElement run, FieldState field, CancellationToken cancellationToken)
    {
        var output = new StringBuilder();
        var properties = FindChild(run, "rPr");
        var sources = GetRunFormattingSources(properties);
        var standardFontSize = GetFontSizeCss(FindRunProperty(sources, "sz"));
        var complexScriptFontSize = GetFontSizeCss(FindRunProperty(sources, "szCs"));
        var hasDifferentComplexScriptSize = complexScriptFontSize is not null && complexScriptFontSize != standardFontSize;
        foreach (var child in run.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (NameIs(child, "rPr"))
                continue;
            if (NameIs(child, "t") || NameIs(child, "delText"))
                output.Append(ConvertRunText(child, field, hasDifferentComplexScriptSize ? complexScriptFontSize : null, sources));
            else
                output.Append(ConvertInline(child, field, cancellationToken));
        }
        if (field.Active)
        {
            if (!field.InInstruction)
                field.Result.Append(output);
            return string.Empty;
        }
        var content = output.ToString();
        if (content.Length == 0)
            return content;
        return ApplyRunFormatting(content, properties, run.InnerText);
    }

    /// <summary>
    /// Applies formatting to a run of text based on the provided properties and returns the formatted HTML string.
    /// </summary>
    /// <param name="content">The text content of the run.</param>
    /// <param name="properties">The properties element containing formatting information.</param>
    /// <param name="text">The original text of the run.</param>
    /// <returns>The formatted HTML string.</returns>
    private string ApplyRunFormatting(string content, OpenXmlElement? properties, string text)
    {
        var sources = GetRunFormattingSources(properties);
        var style = new Dictionary<string, string>(StringComparer.Ordinal);

        if (ResolveRunFontFamily(sources, text) is { } fontFamily)
        { 
            style["font-family"] = fontFamily; 
        }
            
        if (GetFontSizeCss(FindRunProperty(sources, "sz")) is { } fontSize)
        {
            style["font-size"] = fontSize;
        }
            
        var color = GetAttribute(FindRunProperty(sources, "color"), "val");

        if (IsHexColor(color))
        { 
            style["color"] = $"#{color}"; 
        }

        var highlight = GetAttribute(FindRunProperty(sources, "highlight"), "val");

        if (TryGetHighlightColor(highlight, out var highlightColor)) 
        {
            style["background-color"] = highlightColor;
        }

        var spacing = GetAttribute(FindRunProperty(sources, "spacing"), "val");

        if (int.TryParse(spacing, NumberStyles.Integer, CultureInfo.InvariantCulture, out var spacingValue))
        {
            style["letter-spacing"] = (spacingValue / 20d).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        }

        string? underlineValue = GetAttribute(FindRunProperty(sources, "u"), "val");

        if (TryGetUnderlineStyle(underlineValue, out var underlineStyle))
        {
            style["text-decoration"] = underlineStyle;
        }

        string? underlineColor = GetAttribute(FindRunProperty(sources, "u"), "color");

        if (IsHexColor(underlineColor))
        {
            style["text-decoration-color"] = $"#{underlineColor}";
        }

        if (IsOn(FindRunProperty(sources, "b")))
        { 
            content = $"<strong>{content}</strong>"; 
        }

        if (IsOn(FindRunProperty(sources, "i"))) 
        { 
            content = $"<em>{content}</em>"; 
        }

        if (IsOn(FindRunProperty(sources, "strike")) || IsOn(FindRunProperty(sources, "dstrike")))
        {
            content = $"<del>{content}</del>";
        }

        if (IsOn(FindRunProperty(sources, "caps"))) 
        { 
            style["text-transform"] = "uppercase"; 
        }

        if (IsOn(FindRunProperty(sources, "smallCaps"))) 
        { 
            style["font-variant"] = "small-caps"; 
        }

        if (IsOn(FindRunProperty(sources, "vanish"))) 
        { 
            style["display"] = "none"; 
        }

        var vertical = GetAttribute(FindRunProperty(sources, "vertAlign"), "val");

        if (vertical == "superscript") 
        { 
            content = $"<sup>{content}</sup>"; 
        }

        if (vertical == "subscript") 
        { 
            content = $"<sub>{content}</sub>"; 
        }

        return style.Count == 0 ? content : $"<span{StyleAttribute(style)}>{content}</span>";
    }

    private string ConvertRunText(OpenXmlElement textElement, FieldState field, string? complexScriptFontSize, List<OpenXmlElement?> sources)
    {
        var rawText = textElement.InnerText;
        var runes = rawText.EnumerateRunes().ToArray();
        var parenthesisFontSlots = GetParenthesisFontSlots(runes);
        var output = new StringBuilder();
        var segment = new StringBuilder();
        Dictionary<string, string>? segmentStyle = null;

        void AppendSegment()
        {
            if (segment.Length == 0 || segmentStyle is null)
                return;

            var encoded = System.Net.WebUtility.HtmlEncode(segment.ToString());
            if (segmentStyle.Count == 0)
                output.Append(encoded);
            else
                output.Append("<span").Append(StyleAttribute(segmentStyle)).Append('>').Append(encoded).Append("</span>");
            segment.Clear();
        }

        for (var index = 0; index < runes.Length; index++)
        {
            var style = new Dictionary<string, string>(StringComparer.Ordinal);
            if (complexScriptFontSize is not null && IsComplexScriptTextRune(runes, index))
                style["font-size"] = complexScriptFontSize;
            if (parenthesisFontSlots.TryGetValue(index, out var slot) && ResolveRunFontFamily(sources, slot, rawText) is { } fontFamily)
                style["font-family"] = QuoteCssFontFamily(fontFamily);

            if (segmentStyle is not null &&
                (style.Count != segmentStyle.Count || style.Any(item => !segmentStyle.TryGetValue(item.Key, out var value) || value != item.Value)))
                AppendSegment();
            segmentStyle = style;
            segment.Append(runes[index]);
        }
        AppendSegment();

        var text = output.ToString();
        if (!field.Active && (GetAttribute(textElement, "space", "http://www.w3.org/XML/1998/namespace") == "preserve" || rawText.Contains("  ", StringComparison.Ordinal)))
            text = "<span class=\"doc-preserve\">" + text + "</span>";
        if (field.Active)
        {
            if (!field.InInstruction)
                field.Result.Append(text);
            return string.Empty;
        }
        return text;
    }

    private static string? GetFontSizeCss(OpenXmlElement? sizeProperty)
    {
        var size = GetAttribute(sizeProperty, "val");
        return double.TryParse(size, NumberStyles.Number, CultureInfo.InvariantCulture, out var halfPoints)
            ? (halfPoints / 2).ToString("0.##", CultureInfo.InvariantCulture) + "pt"
            : null;
    }

    private static bool IsComplexScriptTextRune(Rune[] runes, int index)
    {
        if (IsComplexScriptCharacter(runes[index].Value))
            return true;
        if (!Rune.IsWhiteSpace(runes[index]))
            return false;

        for (var previous = index - 1; previous >= 0; previous--)
        {
            if (Rune.IsWhiteSpace(runes[previous]))
                continue;
            return IsComplexScriptCharacter(runes[previous].Value);
        }
        for (var next = index + 1; next < runes.Length; next++)
        {
            if (Rune.IsWhiteSpace(runes[next]))
                continue;
            return IsComplexScriptCharacter(runes[next].Value);
        }
        return false;
    }

    private List<OpenXmlElement?> GetRunFormattingSources(OpenXmlElement? runProperties)
    {
        var sources = new List<OpenXmlElement?> { runProperties };
        var characterStyleId = GetAttribute(FindChild(runProperties, "rStyle"), "val");
        sources.AddRange(GetStyleProperties(characterStyleId, "rPr"));
        sources.AddRange(GetStyleProperties(_activeParagraphStyleId, "rPr"));
        sources.Add(GetDocumentDefaultRunProperties());
        return sources;
    }

    private static OpenXmlElement? FindRunProperty(IEnumerable<OpenXmlElement?> sources, string propertyName)
        => sources.Select(source => FindChild(source, propertyName)).FirstOrDefault(property => property is not null);

    private string? ResolveRunFontFamily(List<OpenXmlElement?> sources, string text)
    {
        var families = new List<string>();
        foreach (var slot in GetFontSlots(text))
        {
            var family = ResolveRunFontFamily(sources, slot, text);
            if (!string.IsNullOrWhiteSpace(family) && !families.Contains(family, StringComparer.OrdinalIgnoreCase))
                families.Add(family);
        }

        return families.Count == 0 ? null : string.Join(", ", families.Select(QuoteCssFontFamily));
    }

    private string? ResolveRunFontFamily(IEnumerable<OpenXmlElement?> sources, string slot, string text)
    {
        foreach (var source in sources)
        {
            var fonts = FindChild(source, "rFonts");
            if (fonts is null)
                continue;

            var family = ResolveDeclaredFont(fonts, slot, text);
            if (!string.IsNullOrWhiteSpace(family))
                return family;
        }

        return null;
    }

    private string? ResolveDocumentDefaultFontFamily()
    {
        var defaults = GetDocumentDefaultRunProperties();
        if (defaults is null)
            return null;
        var fonts = FindChild(defaults, "rFonts");
        var family = fonts is null ? null : ResolveDeclaredFont(fonts, "ascii", string.Empty);
        return string.IsNullOrWhiteSpace(family) ? null : QuoteCssFontFamily(family);
    }

    private string? GetParagraphStyleId(OpenXmlElement paragraph)
    {
        var styleId = GetAttribute(FindChild(FindChild(paragraph, "pPr"), "pStyle"), "val");
        if (!string.IsNullOrWhiteSpace(styleId))
            return styleId;

        var defaultStyle = _mainPart.StyleDefinitionsPart?.Styles?.ChildElements.FirstOrDefault(element =>
            NameIs(element, "style") && GetAttribute(element, "type") == "paragraph" &&
            GetAttribute(element, "default") is "1" or "true" or "on");
        return GetAttribute(defaultStyle, "styleId");
    }

    private IEnumerable<OpenXmlElement> GetStyleProperties(string? styleId, string propertyName)
    {
        var styles = _mainPart.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId))
        {
            var style = styles?.ChildElements.FirstOrDefault(element =>
                NameIs(element, "style") && GetAttribute(element, "styleId") == styleId);
            if (style is null)
                yield break;

            if (FindChild(style, propertyName) is { } properties)
                yield return properties;
            styleId = GetAttribute(FindChild(style, "basedOn"), "val");
        }
    }

    private OpenXmlElement? GetDocumentDefaultRunProperties()
    {
        var defaults = FindChild(_mainPart.StyleDefinitionsPart?.Styles, "docDefaults");
        return FindChild(FindChild(defaults, "rPrDefault"), "rPr");
    }

    private string? ResolveDeclaredFont(OpenXmlElement fonts, string slot, string text)
    {
        var (fontAttributes, themeAttributes) = slot switch
        {
            "eastAsia" => (["eastAsia"], ["eastAsiaTheme"]),
            "cs" => (["cs"], ["cstheme"]),
            _ => (new[] { "ascii", "hAnsi" }, new[] { "asciiTheme", "hAnsiTheme" })
        };

        foreach (var attribute in fontAttributes)
        {
            var font = GetAttribute(fonts, attribute);
            if (!string.IsNullOrWhiteSpace(font))
                return font;
        }
        foreach (var attribute in themeAttributes)
        {
            var themeName = GetAttribute(fonts, attribute);
            if (!string.IsNullOrWhiteSpace(themeName) && ResolveThemeFont(themeName, slot, text) is { } themeFont)
                return themeFont;
        }
        return null;
    }

    private string? ResolveThemeFont(string themeName, string slot, string text)
    {
        var themeFont = themeName.StartsWith("major", StringComparison.OrdinalIgnoreCase)
            ? "majorFont"
            : themeName.StartsWith("minor", StringComparison.OrdinalIgnoreCase)
                ? "minorFont"
                : null;
        if (themeFont is null)
            return null;

        var themeElements = FindChild(_mainPart.ThemePart?.RootElement, "themeElements");
        var fontScheme = FindChild(themeElements, "fontScheme");
        var fontSet = FindChild(fontScheme, themeFont);
        var fontElement = FindChild(fontSet, slot == "eastAsia" ? "ea" : slot == "cs" ? "cs" : "latin");
        var typeface = GetAttribute(fontElement, "typeface");
        if (!string.IsNullOrWhiteSpace(typeface))
            return typeface;

        var script = GetThemeScript(text, slot);
        var supplementalFont = script is null ? null : fontSet?.ChildElements.FirstOrDefault(element =>
            NameIs(element, "font") && GetAttribute(element, "script") == script);
        typeface = GetAttribute(supplementalFont, "typeface");
        return string.IsNullOrWhiteSpace(typeface) ? null : typeface;
    }

    private static string? GetThemeScript(string text, string slot)
    {
        foreach (var rune in text.EnumerateRunes())
        {
            var value = rune.Value;
            if (slot == "eastAsia")
            {
                if (value is >= 0x3040 and <= 0x30FF or >= 0x31F0 and <= 0x31FF)
                    return "Jpan";
                if (value is >= 0xAC00 and <= 0xD7AF)
                    return "Hang";
                if (value is >= 0x2E80 and <= 0xA4CF or >= 0xF900 and <= 0xFAFF or >= 0x20000 and <= 0x323AF)
                    return "Hans";
            }
            else if (slot == "cs")
            {
                if (value is >= 0x0590 and <= 0x05FF)
                    return "Hebr";
                if (value is >= 0x0600 and <= 0x08FF)
                    return "Arab";
                if (value is >= 0x0E00 and <= 0x0E7F)
                    return "Thai";
            }
        }
        return null;
    }

    private static List<string> GetFontSlots(string text)
    {
        var slots = new List<string>();
        foreach (var rune in text.EnumerateRunes())
        {
            var value = rune.Value;
            var slot = IsEastAsianCharacter(value) ? "eastAsia" : IsComplexScriptCharacter(value) ? "cs" : "ascii";
            if (!slots.Contains(slot, StringComparer.Ordinal))
                slots.Add(slot);
        }
        if (slots.Count == 0)
            slots.Add("ascii");
        return slots;
    }

    private static bool IsEastAsianCharacter(int value)
        => value is >= 0x2E80 and <= 0xA4CF or >= 0xAC00 and <= 0xD7AF or >= 0xF900 and <= 0xFAFF
            or >= 0xFE30 and <= 0xFE4F or >= 0xFF00 and <= 0xFFEF or >= 0x20000 and <= 0x323AF;

    private static Dictionary<int, string> GetParenthesisFontSlots(Rune[] runes)
    {
        var pairs = new List<(int OpeningIndex, int ClosingIndex)>();
        var openings = new List<(int Index, int ClosingValue)>();
        for (var index = 0; index < runes.Length; index++)
        {
            var value = runes[index].Value;
            if (value is 0x28 or 0xFF08)
            {
                openings.Add((index, value == 0x28 ? 0x29 : 0xFF09));
                continue;
            }

            for (var openingIndex = openings.Count - 1; openingIndex >= 0; openingIndex--)
            {
                if (openings[openingIndex].ClosingValue != value)
                    continue;

                pairs.Add((openings[openingIndex].Index, index));
                openings.RemoveAt(openingIndex);
                break;
            }
        }

        var fontSlots = new Dictionary<int, string>();
        foreach (var (openingIndex, closingIndex) in pairs)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            string? firstSlot = null;
            for (var index = openingIndex + 1; index < closingIndex; index++)
            {
                var rune = runes[index];
                var category = Rune.GetUnicodeCategory(rune);
                if (!Rune.IsLetter(rune) && category is not (UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark))
                    continue;

                var slot = IsEastAsianCharacter(rune.Value) ? "eastAsia" : IsComplexScriptCharacter(rune.Value) ? "cs" : "ascii";
                firstSlot ??= slot;
                counts[slot] = counts.GetValueOrDefault(slot) + 1;
            }

            if (firstSlot is null)
                continue;

            var dominantSlot = firstSlot;
            var dominantCount = 0;
            foreach (var (slot, count) in counts)
            {
                if (count > dominantCount)
                {
                    dominantSlot = slot;
                    dominantCount = count;
                }
            }
            fontSlots[openingIndex] = dominantSlot;
            fontSlots[closingIndex] = dominantSlot;
        }

        return fontSlots;
    }

    private static bool IsComplexScriptCharacter(int value)
        => value is >= 0x0590 and <= 0x08FF or >= 0x0900 and <= 0x0DFF or >= 0x0E00 and <= 0x0E7F
            or >= 0x1000 and <= 0x109F or >= 0x1780 and <= 0x17FF;

    private static string QuoteCssFontFamily(string family)
    {
        var escaped = family.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\d ", StringComparison.Ordinal)
            .Replace("\n", "\\a ", StringComparison.Ordinal);
        return "\"" + escaped + "\"";
    }

    private string ConvertTable(OpenXmlElement table, CancellationToken cancellationToken)
    {
        var tableBorders = FindChild(FindChild(table, "tblPr"), "tblBorders");
        var tableClass = tableBorders is null ? "doc-table doc-table-default-borders" : "doc-table";
        var output = new StringBuilder("<table class=\"").Append(tableClass).Append("\">");
        var grid = FindChild(table, "tblGrid");
        if (grid is not null)
        {
            output.Append("<colgroup>");
            foreach (var column in grid.ChildElements.Where(child => NameIs(child, "gridCol")))
            {
                var width = GetAttribute(column, "w");
                output.Append(TryTwips(width, out var points)
                    ? $"<col style=\"width:{points.ToString("0.##", CultureInfo.InvariantCulture)}pt\" />"
                    : "<col />");
            }
            output.Append("</colgroup>");
        }
        output.Append("<tbody>");
        var rows = table.ChildElements.Where(child => NameIs(child, "tr")).ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = rows[rowIndex];
            var rowProperties = FindChild(row, "trPr");
            var isHeader = FindChild(rowProperties, "tblHeader") is not null;
            output.Append("<tr>");
            var cells = row.ChildElements.Where(child => NameIs(child, "tc")).ToArray();
            for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
            {
                var cell = cells[cellIndex];
                var properties = FindChild(cell, "tcPr");
                var gridSpan = ParsePositiveInt(GetAttribute(FindChild(properties, "gridSpan"), "val"), 1);
                var merge = GetAttribute(FindChild(properties, "vMerge"), "val");
                if (FindChild(properties, "vMerge") is not null && merge != "restart")
                    continue;
                var tag = isHeader ? "th" : "td";
                var attrs = new StringBuilder();
                if (gridSpan > 1) attrs.Append(" colspan=\"").Append(gridSpan).Append('"');
                var rowSpan = merge == "restart" ? GetRowSpan(rows, rowIndex, cellIndex) : 1;
                if (rowSpan > 1) attrs.Append(" rowspan=\"").Append(rowSpan).Append('"');
                var cellStyle = GetCellStyle(properties, tableBorders, rowIndex, cellIndex, cells.Length, rows.Length);
                if (cellStyle.Count > 0) attrs.Append(StyleAttribute(cellStyle));
                output.Append('<').Append(tag).Append(attrs).Append('>');
                AppendBlocks(output, cell.ChildElements, cancellationToken);
                output.Append("</").Append(tag).Append('>');
            }
            output.AppendLine("</tr>");
        }
        output.Append("</tbody></table>");
        return output.ToString();
    }

    private static int GetRowSpan(OpenXmlElement[] rows, int rowIndex, int cellIndex)
    {
        var span = 1;
        for (var next = rowIndex + 1; next < rows.Length; next++)
        {
            var cells = rows[next].ChildElements.Where(child => NameIs(child, "tc")).ToArray();
            if (cellIndex >= cells.Length || FindChild(FindChild(cells[cellIndex], "tcPr"), "vMerge") is null)
                break;
            span++;
        }
        return span;
    }

    private static Dictionary<string, string> GetCellStyle(
        OpenXmlElement? properties,
        OpenXmlElement? tableBorders,
        int rowIndex,
        int cellIndex,
        int cellCount,
        int rowCount)
    {
        var style = new Dictionary<string, string>(StringComparer.Ordinal);
        var shading = GetAttribute(FindChild(properties, "shd"), "fill");
        if (IsHexColor(shading)) style["background-color"] = "#" + shading;
        var vertical = GetAttribute(FindChild(properties, "vAlign"), "val");
        if (vertical is "top" or "center" or "bottom") style["vertical-align"] = vertical == "center" ? "middle" : vertical;

        var cellBorders = FindChild(properties, "tcBorders");
        foreach (var side in new[] { "top", "right", "bottom", "left" })
        {
            var tableSide = side switch
            {
                "top" => rowIndex == 0 ? "top" : "insideH",
                "bottom" => rowIndex == rowCount - 1 ? "bottom" : "insideH",
                "left" => cellIndex == 0 ? "left" : "insideV",
                _ => cellIndex == cellCount - 1 ? "right" : "insideV"
            };
            var border = FindChild(cellBorders, side) ?? FindChild(tableBorders, tableSide);
            if (GetBorderCss(border) is { } borderCss)
                style["border-" + side] = borderCss;
        }

        foreach (var side in new[] { "insideH", "insideV" })
        {
            var border = FindChild(cellBorders, side);
            var cssSide = side == "insideH" ? "bottom" : "right";
            if (GetBorderCss(border) is { } borderCss)
                style["border-" + cssSide] = borderCss;
        }
        return style;
    }

    private static string? GetBorderCss(OpenXmlElement? border)
    {
        if (border is null)
            return null;
        var value = GetAttribute(border, "val");
        if (value is "nil" or "none")
            return "none";

        var size = ParsePositiveInt(GetAttribute(border, "sz"), 4) / 8d;
        var color = GetAttribute(border, "color");
        var cssColor = IsHexColor(color) ? "#" + color : "#000000";
        return $"{size.ToString("0.###", CultureInfo.InvariantCulture)}pt solid {cssColor}";
    }

    private Dictionary<string, string> GetParagraphStyle(OpenXmlElement paragraph)
    {
        var style = new Dictionary<string, string>(StringComparer.Ordinal);
        AddParagraphProperties(style, GetDocumentDefaultParagraphProperties());
        var paragraphStyleId = GetParagraphStyleId(paragraph);
        foreach (var properties in GetStyleProperties(paragraphStyleId, "pPr").Reverse())
            AddParagraphProperties(style, properties);
        AddParagraphProperties(style, FindChild(paragraph, "pPr"));
        return style;
    }

    private OpenXmlElement? GetDocumentDefaultParagraphProperties()
    {
        var defaults = FindChild(_mainPart.StyleDefinitionsPart?.Styles, "docDefaults");
        return FindChild(FindChild(defaults, "pPrDefault"), "pPr");
    }

    private static void AddParagraphProperties(Dictionary<string, string> style, OpenXmlElement? properties)
    {
        var alignment = GetAttribute(FindChild(properties, "jc"), "val");
        if (alignment is not null)
            style["text-align"] = alignment switch { "center" => "center", "right" or "end" => "right", "both" or "distribute" => "justify", _ => "left" };
        var spacing = FindChild(properties, "spacing");
        AddTwips(style, "margin-top", GetAttribute(spacing, "before"));
        AddTwips(style, "margin-bottom", GetAttribute(spacing, "after"));
        var line = GetAttribute(spacing, "line");
        if (int.TryParse(line, NumberStyles.Integer, CultureInfo.InvariantCulture, out var lineValue))
        {
            var lineRule = GetAttribute(spacing, "lineRule");
            var lineHeight = lineRule switch
            {
                "exact" => (lineValue / 20d).ToString("0.##", CultureInfo.InvariantCulture) + "pt",
                "atLeast" => "max(1em, " + (lineValue / 20d).ToString("0.##", CultureInfo.InvariantCulture) + "pt)",
                _ => (lineValue / 240d).ToString("0.###", CultureInfo.InvariantCulture)
            };
            style["line-height"] = lineHeight;
        }
        ApplyToggleProperty(style, properties, "keepNext", "break-after", "avoid-page");
        ApplyToggleProperty(style, properties, "keepLines", "break-inside", "avoid-page");
        ApplyToggleProperty(style, properties, "pageBreakBefore", "break-before", "page");
        if (FindChild(properties, "widowControl") is { } widowControl)
        {
            var minimumLines = IsOn(widowControl) ? "2" : "1";
            style["orphans"] = minimumLines;
            style["widows"] = minimumLines;
        }
        var indent = FindChild(properties, "ind");
        AddTwips(style, "margin-left", GetAttribute(indent, "left") ?? GetAttribute(indent, "start"));
        AddTwips(style, "margin-right", GetAttribute(indent, "right") ?? GetAttribute(indent, "end"));
        AddTwips(style, "text-indent", GetAttribute(indent, "firstLine"));
        if (GetAttribute(indent, "hanging") is { } hanging && TryTwips(hanging, out var hangingPt))
        {
            style["text-indent"] = (-hangingPt).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
            style["padding-left"] = hangingPt.ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        }
        var shading = GetAttribute(FindChild(properties, "shd"), "fill");
        if (IsHexColor(shading)) style["background-color"] = "#" + shading;
        var borders = FindChild(properties, "pBdr");
        foreach (var side in new[] { "top", "right", "bottom", "left" })
        {
            var border = FindChild(borders, side);
            var color = GetAttribute(border, "color");
            if (border is not null && GetAttribute(border, "val") is not ("nil" or "none"))
                style["border-" + side] = $"{(ParsePositiveInt(GetAttribute(border, "sz"), 4) / 8d).ToString("0.##", CultureInfo.InvariantCulture)}pt solid {(IsHexColor(color) ? "#" + color : "#d1d5db")}";
        }
    }

    private static void ApplyToggleProperty(Dictionary<string, string> style, OpenXmlElement? properties, string elementName, string cssProperty, string cssValue)
    {
        if (FindChild(properties, elementName) is not { } element)
            return;

        if (IsOn(element))
            style[cssProperty] = cssValue;
        else
            style.Remove(cssProperty);
    }

    private bool IsOrderedList(string numId, int level)
    {
        var format = GetAttribute(GetNumberingLevelProperty(numId, level, "numFmt"), "val");
        return format is not ("bullet" or "none");
    }

    private string? ResolveNumberingMarkerFontFamily(OpenXmlElement paragraph, string numId, int level)
    {
        var levelElement = GetNumberingLevelElement(numId, level);
        var markerText = GetAttribute(FindChild(levelElement, "lvlText"), "val") ?? string.Empty;
        var paragraphRunProperties = FindChild(FindChild(paragraph, "pPr"), "rPr");
        var rFonts = FindChild(paragraphRunProperties, "rFonts");

        if (rFonts is not null)
        {
            var markerSlot = GetFontSlots(markerText).FirstOrDefault() ?? "ascii";
            var markerFamily = ResolveDeclaredFont(rFonts, markerSlot, markerText);
            if (!string.IsNullOrWhiteSpace(markerFamily))
                return markerFamily;
        }

        var defaultFonts = FindChild(GetDocumentDefaultRunProperties(), "rFonts");
        return defaultFonts is null
            ? null
            : ResolveDeclaredFont(defaultFonts, GetFontSlots(markerText).FirstOrDefault() ?? "ascii", markerText);
    }

    /// <summary>
    /// Gets the CSS ordered-list style type for a numbering level.
    /// </summary>
    /// <param name="numId">The numbering identifier.</param>
    /// <param name="level">The numbering level.</param>
    /// <returns>The CSS ordered-list style type, or <c>decimal</c> when the format is unsupported.</returns>
    private string GetOrderedListStyleType(string numId, int level)
    {
        var format = GetAttribute(GetNumberingLevelProperty(numId, level, "numFmt"), "val");
        return format switch
        {
            "decimal" => "decimal",
            "lowerLetter" => "lower-alpha",
            "upperLetter" => "upper-alpha",
            "lowerRoman" => "lower-roman",
            "upperRoman" => "upper-roman",
            _ => "decimal"
        };
    }

    /// <summary>
    /// Gets the CSS unordered-list style type for a numbering level.
    /// </summary>
    /// <param name="numId">The numbering identifier.</param>
    /// <param name="level">The numbering level.</param>
    /// <returns>A style type of <c>disc</c>, <c>circle</c>, <c>square</c>, or <c>none</c>.</returns>
    private string GetUnorderedListStyleType(string numId, int level)
    {
        var marker = GetAttribute(GetNumberingLevelProperty(numId, level, "lvlText"), "val");
        if (GetAttribute(GetNumberingLevelProperty(numId, level, "numFmt"), "val") == "none" || string.IsNullOrEmpty(marker))
            return "none";

        return marker switch
        {
            "\uF0B7" or "\u2022" => "disc",
            "o" or "\u25CB" or "\u25E6" => "circle",
            "\uF0A7" or "\u25AA" or "\u25A0" => "square",
            _ => "disc"
        };
    }

    private OpenXmlElement? GetNumberingLevelProperty(string numId, int level, string propertyName)
    {
        var levelElement = GetNumberingLevelElement(numId, level);
        return FindChild(levelElement, propertyName);
    }

    private OpenXmlElement? GetNumberingLevelElement(string numId, int level)
    {
        var numbering = _mainPart.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
            return null;

        var instance = numbering.ChildElements.FirstOrDefault(element => NameIs(element, "num") && GetAttribute(element, "numId") == numId);
        var levelOverride = instance?.ChildElements.FirstOrDefault(element =>
            NameIs(element, "lvlOverride") && ParsePositiveInt(GetAttribute(element, "ilvl"), 0) == level);
        if (FindChild(levelOverride, "lvl") is { } overriddenLevel)
            return overriddenLevel;

        var abstractId = GetAttribute(instance is null ? null : FindChild(instance, "abstractNumId"), "val");
        var abstractNum = numbering.ChildElements.FirstOrDefault(element => NameIs(element, "abstractNum") && GetAttribute(element, "abstractNumId") == abstractId);
        return abstractNum?.ChildElements.FirstOrDefault(element =>
            NameIs(element, "lvl") && ParsePositiveInt(GetAttribute(element, "ilvl"), 0) == level);
    }

    private void AppendHeadersAndFooters(StringBuilder output, OpenXmlElement body, string referenceName, string elementName, CancellationToken cancellationToken)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var section in body.Descendants().Where(element => NameIs(element, "sectPr")).Prepend(FindChild(body, "sectPr")).Where(element => element is not null))
        {
            foreach (var reference in section!.ChildElements.Where(element => NameIs(element, referenceName)))
            {
                var relationshipId = GetAttribute(reference, "id", RelationshipNamespace);
                if (relationshipId is null || !seen.Add(relationshipId)) continue;
                var part = _mainPart.GetPartById(relationshipId);
                if (part.RootElement is not { } root) continue;
                var referenceType = GetAttribute(reference, "type");
                var type = referenceType is "first" or "even" or "default" ? referenceType : "default";
                cancellationToken.ThrowIfCancellationRequested();
                var previousPart = _activePart;
                _activePart = part;
                try
                {
                    var content = ConvertBlocksToString(root.ChildElements, cancellationToken);
                    output.Append('<').Append(elementName)
                        .Append(" class=\"doc-").Append(elementName).Append(" doc-").Append(type).Append("-page\">")
                        .Append(content).Append("</").Append(elementName).AppendLine(">");
                }
                finally
                {
                    _activePart = previousPart;
                }
            }
        }
    }

    private void AppendNotes(StringBuilder output, CancellationToken cancellationToken)
    {
        AppendNotePart(output, _mainPart.FootnotesPart, "footnote", cancellationToken);
        AppendNotePart(output, _mainPart.EndnotesPart, "endnote", cancellationToken);
    }

    private void AppendNotePart(StringBuilder output, OpenXmlPart? notePart, string kind, CancellationToken cancellationToken)
    {
        var notes = notePart?.RootElement;
        if (notes is null) return;
        var selected = notes.ChildElements.Where(note => NameIs(note, kind) && GetAttribute(note, "id") is not ("-1" or "0")).ToArray();
        if (selected.Length == 0) return;
        output.Append("<aside class=\"doc-").Append(kind).Append("s\"><ol>");
        var previousPart = _activePart;
        _activePart = notePart!;
        try
        {
            foreach (var note in selected)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var id = GetAttribute(note, "id") ?? string.Empty;
                output.Append("<li id=\"").Append(kind).Append('-').Append(Attr(id)).Append("\">");
                AppendBlocks(output, note.ChildElements, cancellationToken);
                output.Append("</li>");
            }
        }
        finally
        {
            _activePart = previousPart;
        }
        output.Append("</ol></aside>");
    }

    private string ConvertDrawing(OpenXmlElement drawing, CancellationToken cancellationToken)
    {
        var blip = drawing.Descendants().FirstOrDefault(element => NameIs(element, "blip"));
        var relationshipId = GetAttribute(blip, "embed", RelationshipNamespace);
        if (relationshipId is not null && _activePart.GetPartById(relationshipId) is ImagePart imagePart)
        {
            using var stream = imagePart.GetStream();
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            var contentType = imagePart.ContentType.ToLowerInvariant();
            if (contentType is "image/png" or "image/jpeg" or "image/gif" or "image/webp" or "image/bmp")
            {
                var extents = drawing.Descendants().FirstOrDefault(element => NameIs(element, "extent"));
                var width = ParsePositiveLong(GetAttribute(extents, "cx"));
                var height = ParsePositiveLong(GetAttribute(extents, "cy"));
                var dimensions = width > 0 && height > 0
                    ? $" width=\"{(width / (double)EmusPerPixel).ToString("0.##", CultureInfo.InvariantCulture)}\" height=\"{(height / (double)EmusPerPixel).ToString("0.##", CultureInfo.InvariantCulture)}\""
                    : string.Empty;
                var alt = GetAttribute(drawing.Descendants().FirstOrDefault(element => NameIs(element, "docPr")), "descr") ?? "";
                return $"<img class=\"doc-image\" src=\"data:{contentType};base64,{System.Convert.ToBase64String(memory.ToArray())}\" alt=\"{Attr(alt)}\"{dimensions} />";
            }
            _messages.Add($"An image with unsupported content type '{contentType}' was omitted.");
            return string.Empty;
        }

        var shape = drawing.Descendants().FirstOrDefault(element => NameIs(element, "sp") || NameIs(element, "shape"));
        if (shape is not null)
            return NameIs(shape, "shape") ? ConvertVmlShape(shape, cancellationToken) : ConvertShape(shape, cancellationToken);

        var textBox = drawing.Descendants().FirstOrDefault(element => NameIs(element, "txbxContent"));
        if (textBox is not null)
            return $"<div class=\"doc-textbox\">{ConvertBlocksToString(textBox.ChildElements, cancellationToken)}</div>";

        if (drawing.Descendants().Any(element => NameIs(element, "graphicData")))
            _messages.Add("A chart or SmartArt drawing was omitted because its embedded visual model is not supported.");
        return string.Empty;
    }

    private string ConvertVmlShape(OpenXmlElement shape, CancellationToken cancellationToken)
    {
        var textBox = shape.Descendants().FirstOrDefault(element => NameIs(element, "txbxContent"));
        var content = textBox is null
            ? System.Net.WebUtility.HtmlEncode(shape.InnerText)
            : ConvertBlocksToString(textBox.ChildElements, cancellationToken);
        var width = GetVmlDimension(GetAttribute(shape, "style"), "width");
        var height = GetVmlDimension(GetAttribute(shape, "style"), "height");
        var style = new Dictionary<string, string>(StringComparer.Ordinal);
        if (width is not null) style["width"] = width;
        if (height is not null) style["height"] = height;
        return $"<div class=\"doc-shape doc-vml-shape\"{StyleAttribute(style)}>{content}</div>";
    }

    private static string? GetVmlDimension(string? style, string property)
    {
        if (style is null) return null;
        var match = Regex.Match(style, $@"(?:^|;)\s*{property}\s*:\s*([0-9]+(?:\.[0-9]+)?(?:px|pt|in|cm|mm))", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return match.Success ? match.Groups[1].Value : null;
    }

    private string ConvertShape(OpenXmlElement shape, CancellationToken cancellationToken)
    {
        var geometry = GetAttribute(shape.Descendants().FirstOrDefault(element => NameIs(element, "prstGeom")), "prst") ?? "rect";
        if (geometry is not ("rect" or "roundRect" or "ellipse" or "line" or "triangle" or "diamond"))
            _messages.Add($"Shape preset '{geometry}' was rendered as a rectangle.");
        var preset = geometry switch
        {
            "roundRect" => "roundRect", "ellipse" => "ellipse", "line" => "line", "triangle" => "triangle", "diamond" => "diamond", _ => "rect"
        };
        var properties = FindChild(shape, "spPr");
        var fill = GetAttribute(properties?.Descendants().FirstOrDefault(element => NameIs(element, "srgbClr")), "val");
        var fillCss = IsHexColor(fill) ? "#" + fill : "#ffffff";
        var textBody = FindChild(shape, "txBody");
        var wordTextBody = shape.Descendants().FirstOrDefault(element => NameIs(element, "txbxContent"));
        var text = textBody is not null
            ? ConvertBlocksToString(textBody.ChildElements, cancellationToken)
            : wordTextBody is not null
                ? ConvertBlocksToString(wordTextBody.ChildElements, cancellationToken)
                : string.Empty;
        var extents = FindChild(properties, "xfrm") is { } transform ? FindChild(transform, "ext") : null;
        var width = ParsePositiveLong(GetAttribute(extents, "cx"));
        var height = ParsePositiveLong(GetAttribute(extents, "cy"));
        var w = width > 0 ? width / (double)EmusPerPixel : 160;
        var h = height > 0 ? height / (double)EmusPerPixel : 80;
        var svgShape = preset switch
        {
            "ellipse" => $"<ellipse cx=\"{SvgNumber(w / 2)}\" cy=\"{SvgNumber(h / 2)}\" rx=\"{SvgNumber(w / 2)}\" ry=\"{SvgNumber(h / 2)}\" fill=\"{fillCss}\" stroke=\"#333\" />",
            "line" => $"<line x1=\"0\" y1=\"0\" x2=\"{SvgNumber(w)}\" y2=\"{SvgNumber(h)}\" stroke=\"#333\" />",
            "triangle" => $"<polygon points=\"{SvgNumber(w / 2)},0 {SvgNumber(w)},{SvgNumber(h)} 0,{SvgNumber(h)}\" fill=\"{fillCss}\" stroke=\"#333\" />",
            "diamond" => $"<polygon points=\"{SvgNumber(w / 2)},0 {SvgNumber(w)},{SvgNumber(h / 2)} {SvgNumber(w / 2)},{SvgNumber(h)} 0,{SvgNumber(h / 2)}\" fill=\"{fillCss}\" stroke=\"#333\" />",
            _ => $"<rect x=\"1\" y=\"1\" width=\"{SvgNumber(Math.Max(0, w - 2))}\" height=\"{SvgNumber(Math.Max(0, h - 2))}\" rx=\"{(preset == "roundRect" ? "8" : "0")}\" fill=\"{fillCss}\" stroke=\"#333\" />"
        };
        return $"<div class=\"doc-shape\"><svg role=\"img\" viewBox=\"0 0 {SvgNumber(w)} {SvgNumber(h)}\" width=\"{SvgNumber(w)}\" height=\"{SvgNumber(h)}\">{svgShape}</svg>{text}</div>";
    }

    private string? ResolveHyperlink(string? relationshipId)
    {
        if (relationshipId is null) return null;
        var relationship = _activePart.HyperlinkRelationships.FirstOrDefault(item => item.Id == relationshipId);
        if (relationship is null) return null;
        var uri = relationship.Uri;
        return uri.IsAbsoluteUri
            ? uri.Scheme is "https" or "http" or "mailto" ? uri.ToString() : null
            : uri.ToString().StartsWith('#') ? uri.ToString() : null;
    }

    private static string RenderField(string instruction, string result)
    {
        var match = Regex.Match(instruction.Trim(), @"^([A-Za-z]+)(?:\s+(.*))?$", RegexOptions.CultureInvariant);
        if (!match.Success) return result;
        var fieldName = match.Groups[1].Value.ToUpperInvariant();
        if (fieldName == "MERGEFIELD")
        {
            var mergeMatch = Regex.Match(match.Groups[2].Value, "\\\"([^\\\"]+)\\\"|([^\\s\\\\]+)", RegexOptions.CultureInvariant);
            var mergeName = mergeMatch.Groups[1].Success ? mergeMatch.Groups[1].Value : mergeMatch.Groups[2].Value;
            return $"<span data-field=\"MERGEFIELD\" data-mergefield=\"{Attr(mergeName)}\">{{{{{Attr(mergeName)}}}}}</span>";
        }
        if (fieldName == "TOC")
            return $"<nav class=\"doc-toc\" data-field=\"TOC\">{result}</nav>";
        if (fieldName == "HYPERLINK")
        {
            var address = Regex.Match(match.Groups[2].Value, "\\\"([^\\\"]+)\\\"", RegexOptions.CultureInvariant).Groups[1].Value;
            var href = SafeExternalUrl(address);
            return href is null ? result : $"<a class=\"doc-link\" href=\"{Attr(href)}\">{result}</a>";
        }
        var known = fieldName is "PAGE" or "NUMPAGES" or "SECTIONPAGES" or "DATE" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" or "FILENAME" or "AUTHOR" or "PAGEREF" or "REF";
        if (fieldName == "DOCPROPERTY")
        {
            var property = Regex.Match(match.Groups[2].Value, "\\\"([^\\\"]+)\\\"|([^\\s\\\\]+)", RegexOptions.CultureInvariant);
            var propertyName = property.Groups[1].Success ? property.Groups[1].Value : property.Groups[2].Value;
            return $"<span data-prop=\"{Attr(propertyName)}\">{result}</span>";
        }
        if (!known) return result;
        var filenameField = fieldName == "FILENAME" && match.Groups[2].Value.Contains("\\p", StringComparison.OrdinalIgnoreCase)
            ? "FILENAME_PATH"
            : fieldName;
        var cssClass = fieldName switch { "PAGE" => " doc-field-page", "NUMPAGES" => " doc-field-pages", _ => string.Empty };
        var attribute = fieldName == "PAGEREF" ? $" data-pageref=\"{Attr(match.Groups[2].Value.Trim().Trim('"'))}\"" : string.Empty;
        var fieldResult = fieldName is "PAGE" or "NUMPAGES" ? string.Empty : result;
        return $"<span class=\"doc-field{cssClass}\" data-field=\"{filenameField}\"{attribute}>{fieldResult}</span>";
    }

    private static void AppendPageRules(StringBuilder css, OpenXmlElement? section, string pageName)
    {
        if (section is null) return;
        var size = FindChild(section, "pgSz");
        var width = ParsePositiveLong(GetAttribute(size, "w"));
        var height = ParsePositiveLong(GetAttribute(size, "h"));
        if (GetAttribute(size, "orient") == "landscape" && width > 0 && height > 0 && width < height)
            (width, height) = (height, width);
        var margins = FindChild(section, "pgMar");
        var sizeCss = width > 0 && height > 0
            ? $"{(width / 1440d).ToString("0.###", CultureInfo.InvariantCulture)}in {(height / 1440d).ToString("0.###", CultureInfo.InvariantCulture)}in"
            : "A4 portrait";
        var marginProperties = new List<string>();
        foreach (var side in new[] { "top", "right", "bottom", "left" })
        {
            var value = GetAttribute(margins, side);
            if (TryTwips(value, out var points))
                marginProperties.Add($"margin-{side}:{(points / 72d).ToString("0.###", CultureInfo.InvariantCulture)}in");
        }
        css.Append("\n@page ").Append(pageName).Append(" { size: ").Append(sizeCss).Append(';');
        if (marginProperties.Count > 0)
            css.Append(' ').Append(string.Join("; ", marginProperties)).Append(';');
        css.AppendLine(" }");

        if (width > 0 && height > 0)
        {
            var screenMargins = string.Join("; ", sourceArray.Select(side =>
            {
                var value = GetAttribute(margins, side);
                var padding = TryTwips(value, out var points)
                    ? (points / 72d).ToString("0.###", CultureInfo.InvariantCulture) + "in"
                    : "0";
                return $"padding-{side}:{padding}";
            }));
            css.Append("\n@media screen { .doc-section.").Append(pageName)
                .Append(" { box-sizing:border-box; width:").Append((width / 1440d).ToString("0.###", CultureInfo.InvariantCulture))
                .Append("in; min-height:").Append((height / 1440d).ToString("0.###", CultureInfo.InvariantCulture))
                .Append("in; ").Append(screenMargins)
                .AppendLine("; margin:24px auto; background:#fff; box-shadow:0 1px 5px #0003; } .doc-root { background:#f3f4f6; padding:1px 0; } }");
        }
    }

    private static void AddTwips(Dictionary<string, string> style, string name, string? value)
    {
        if (TryTwips(value, out var points))
            style[name] = (points+1.5).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
    }

    private static bool TryTwips(string? value, out double points)
    {
        if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var twips))
        {
            points = twips / 20;
            return true;
        }
        points = 0;
        return false;
    }

    private static string StyleAttribute(Dictionary<string, string> style)
        => style.Count == 0 ? string.Empty : " style=\"" + Attr(string.Join(";", style.Select(item => item.Key + ":" + item.Value))) + "\"";

    private static string? SafeExternalUrl(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return Uri.TryCreate(value, UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https" or "mailto"
            ? uri.ToString()
            : null;
    }

    private static bool IsOn(OpenXmlElement? element)
    {
        if (element is null) return false;
        var value = GetAttribute(element, "val");
        return value is null || value is "1" or "true" or "on";
    }

    private static int GetListLevel(OpenXmlElement paragraph)
    {
        var pPr = FindChild(paragraph, "pPr");
        var ilvl = GetAttribute(FindChild(FindChild(pPr, "numPr"), "ilvl"), "val");
        return int.TryParse(ilvl, NumberStyles.Integer, CultureInfo.InvariantCulture, out var level) ? Math.Clamp(level, 0, 8) : 0;
    }

    private static string? GetNumberingId(OpenXmlElement paragraph)
        => GetAttribute(FindChild(FindChild(FindChild(paragraph, "pPr"), "numPr"), "numId"), "val");

    private static bool IsBlockElement(OpenXmlElement element)
        => element.LocalName is "p" or "tbl" or "lastRenderedPageBreak" or "sdt" or "customXml" or "altChunk";

    private static OpenXmlElement? FindChild(OpenXmlElement? parent, string localName)
        => parent?.ChildElements.FirstOrDefault(element => NameIs(element, localName));

    private static bool NameIs(OpenXmlElement? element, string localName)
        => element is not null && element.LocalName == localName;

    private static string? GetAttribute(OpenXmlElement? element, string localName, string? namespaceUri = null)
    {
        if (element is null) return null;
        var attribute = element.GetAttributes().FirstOrDefault(item => item.LocalName == localName && (namespaceUri is null || item.NamespaceUri == namespaceUri));
        return attribute.LocalName is null ? null : attribute.Value;
    }

    private static int ParsePositiveInt(string? value, int fallback)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0 ? parsed : fallback;

    private static long ParsePositiveLong(string? value)
        => long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0 ? parsed : 0;

    private static bool IsHexColor(string? value)
        => value is { Length: 6 } && value.All(Uri.IsHexDigit);

    private static bool TryGetHighlightColor(string? value, out string color)
    {
        color = value?.ToLowerInvariant() switch
        {
            "yellow" => "#ffff00", "green" => "#00ff00", "cyan" => "#00ffff", "magenta" => "#ff00ff",
            "blue" => "#0000ff", "red" => "#ff0000", "darkblue" => "#00008b", "darkcyan" => "#008b8b",
            "darkgreen" => "#006400", "darkmagenta" => "#8b008b", "darkred" => "#8b0000", "darkyellow" => "#808000",
            "darkgray" => "#a9a9a9", "lightgray" => "#d3d3d3", "black" => "#000000", "white" => "#ffffff",
            _ => string.Empty
        };
        return color.Length > 0;
    }

    private static bool TryGetUnderlineStyle(string? value, out string style)
    {
        style = value?.ToLowerInvariant() switch
        {
            "single" => "underline",
            "double" => "underline double",
            "dotted" => "underline dotted",
            "dash" => "underline dashed 1px",
            "wave" => "underline wavy 1px",
            "thick" => "underline 2px",
            _ => string.Empty
        };
        return style.Length > 0;
    }

    private static string SvgNumber(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    private static string Attr(string value) => System.Net.WebUtility.HtmlEncode(value);

    private sealed class FieldState
    {
        public bool Active { get; set; }
        public bool InInstruction { get; set; }
        public StringBuilder Instruction { get; } = new();
        public StringBuilder Result { get; } = new();
        public void Reset()
        {
            Active = false;
            InInstruction = false;
            Instruction.Clear();
            Result.Clear();
        }
    }

    private const string DefaultDocumentCss = """
h1.doc-heading { font-size: 2rem; } h2.doc-heading { font-size: 1.6rem; } h3.doc-heading { font-size: 1.35rem; }
h4.doc-heading { font-size: 1.2rem; } h5.doc-heading { font-size: 1.05rem; } h6.doc-heading { font-size: 0.95rem; }
.doc-root { font-family: "Segoe UI", Arial, sans-serif; color: #111; line-height: 1; background: #fff; }
.doc-paragraph { margin: 0; }
.doc-heading { margin: 0; font-weight: normal; }
.doc-table { width: 100%; border-collapse: collapse; margin: 0.75rem 0; }
.doc-table th, .doc-table td { padding: 6px 8px; vertical-align: top; }
.doc-table-default-borders th, .doc-table-default-borders td { border: 0.5pt solid #000; }
.doc-list { margin: 0.25rem 0 0.75rem 1.5rem; padding-left: 0.75rem; }
.doc-list li::marker { font-family: var(--doc-marker-font, inherit); }
.doc-link { color: #1d4ed8; text-decoration: underline; }
.doc-image { max-width: 100%; height: auto; object-fit: contain; }
.doc-textbox, .doc-shape { display: inline-block; vertical-align: middle; }
.doc-shape svg { max-width: 100%; }
.doc-header, .doc-footer { display: block; }
 .page-break { display: block; break-after: page; page-break-after: always; height: 0; }
.doc-preserve { white-space: pre-wrap; }
.doc-field-page::after { content: counter(page); }
.doc-field-pages::after { content: counter(pages); }
.doc-tab { white-space: pre; }
.doc-equation { font-family: "Cambria Math", serif; }
""";
    internal static readonly string[] sourceArray = ["top", "right", "bottom", "left"];
}
