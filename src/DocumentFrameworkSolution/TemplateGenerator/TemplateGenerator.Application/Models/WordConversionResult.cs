namespace TemplateGenerator.Application.Models;

public sealed class WordConversionResult
{
    public string Html { get; set; } = string.Empty;
    public IReadOnlyList<string> Messages { get; set; } = Array.Empty<string>();
}