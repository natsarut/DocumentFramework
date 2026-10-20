using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Models;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Services;

public interface IDocumentConversionService
{
    Task<WordConversionResult> ParseWordAsync(Stream stream, CancellationToken cancellationToken = default);
    string RenderHtml(WordConversionResult result);
    Task<byte[]> ExportPdfAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default);
}

public sealed class DocumentConversionService : IDocumentConversionService
{
    private readonly IWordDocumentParser _wordDocumentParser;
    private readonly IHtmlDocumentRenderer _htmlDocumentRenderer;
    private readonly IPdfExportService _pdfExportService;

    public DocumentConversionService(
        IWordDocumentParser wordDocumentParser,
        IHtmlDocumentRenderer htmlDocumentRenderer,
        IPdfExportService pdfExportService)
    {
        _wordDocumentParser = wordDocumentParser;
        _htmlDocumentRenderer = htmlDocumentRenderer;
        _pdfExportService = pdfExportService;
    }

    public async Task<WordConversionResult> ParseWordAsync(Stream stream, CancellationToken cancellationToken = default)
        => await _wordDocumentParser.ParseAsync(stream, cancellationToken);

    public string RenderHtml(WordConversionResult result)
        => _htmlDocumentRenderer.Render(result);

    public Task<byte[]> ExportPdfAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default)
        => _pdfExportService.ExportAsync(html, pageSettings, cancellationToken);
}
