using DocumentFormat.OpenXml.Packaging;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Models;

namespace TemplateGenerator.Infrastructure.Word;

public sealed class OpenXmlWordParser : IWordDocumentParser
{
    public async Task<WordConversionResult> ParseAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead)
            throw new InvalidOperationException("Input stream is not readable.");

        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek)
            stream.Position = 0;

        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        memoryStream.Position = 0;

        using var document = WordprocessingDocument.Open(memoryStream, false);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Word document main part not found.");
        return WordprocessingMlHtmlConverter.Convert(mainPart, cancellationToken);
    }
}
