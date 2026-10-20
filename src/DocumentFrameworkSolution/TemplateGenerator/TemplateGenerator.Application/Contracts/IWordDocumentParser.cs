using TemplateGenerator.Application.Models;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Contracts;

public interface IWordDocumentParser
{
    //Task<DocumentModel> ParseAsync(Stream stream, CancellationToken cancellationToken = default);
    Task<WordConversionResult> ParseAsync(Stream stream, CancellationToken cancellationToken = default);
}
