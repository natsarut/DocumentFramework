namespace TemplateGenerator.Domain;

public enum BlockType
{
    Paragraph,
    Heading,
    List,
    Table,
    Image,
    PageBreak
}

public enum PaperSize
{
    A4,
    Letter,
    Legal
}

public enum PageOrientation
{
    Portrait,
    Landscape
}

public sealed class DocumentModel
{
    public List<Section> Sections { get; set; } = new();
    public List<StyleToken> Styles { get; set; } = new();
}

public sealed class Section
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public List<Block> Blocks { get; set; } = new();
    public bool StartsOnNewPage { get; set; }
}

public sealed class Block
{
    public BlockType Type { get; set; }
    public string? Text { get; set; }
    public string? Html { get; set; }
    public List<InlineRun> Runs { get; set; } = new();
    public StyleToken Style { get; set; } = new();
}

public sealed class InlineRun
{
    public string Text { get; set; } = string.Empty;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? Color { get; set; }
}

public sealed class StyleToken
{
    public string Name { get; set; } = string.Empty;
    public string? FontFamily { get; set; }
    public double FontSizePt { get; set; } = 11;
    public string? Color { get; set; }
    public string? Alignment { get; set; }
    public double? LineHeight { get; set; }
    public double? MarginTopPt { get; set; }
    public double? MarginBottomPt { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
}

public sealed class PageSettings
{
    public PaperSize Size { get; set; } = PaperSize.A4;
    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;
    public double MarginTopMm { get; set; } = 12;
    public double MarginRightMm { get; set; } = 12;
    public double MarginBottomMm { get; set; } = 12;
    public double MarginLeftMm { get; set; } = 12;
}
