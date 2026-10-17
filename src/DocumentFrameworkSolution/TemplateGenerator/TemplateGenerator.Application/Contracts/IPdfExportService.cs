using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Contracts;

public interface IPdfExportService
{
    Task<byte[]> ExportAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default);
}
