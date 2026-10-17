using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Services;

public interface IDocumentConversionService
{
    Task<DocumentModel> ParseWordAsync(Stream stream, CancellationToken cancellationToken = default);
    string RenderHtml(DocumentModel document);
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

    public Task<DocumentModel> ParseWordAsync(Stream stream, CancellationToken cancellationToken = default)
        => _wordDocumentParser.ParseAsync(stream, cancellationToken);

    public string RenderHtml(DocumentModel document)
        => _htmlDocumentRenderer.Render(document);

    public Task<byte[]> ExportPdfAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default)
        => _pdfExportService.ExportAsync(html, pageSettings, cancellationToken);
}
