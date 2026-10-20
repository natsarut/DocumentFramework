namespace TemplateGenerator.Infrastructure.Word;

public sealed class CssStyleToken
{
    public string ClassName { get; set; } = string.Empty;
    public Dictionary<string, string> Declarations { get; set; } = new();

    public string ToCssRule()
    {
        if (Declarations.Count == 0)
            return string.Empty;

        return string.Join(" ",
            Declarations.Select(kvp => $"{kvp.Key}: {kvp.Value};"));
    }
}