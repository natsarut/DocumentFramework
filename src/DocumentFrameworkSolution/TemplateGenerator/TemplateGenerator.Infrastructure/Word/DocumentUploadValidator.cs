using TemplateGenerator.Application.Contracts;

namespace TemplateGenerator.Infrastructure.Word;

public sealed class DocumentUploadValidator : IDocumentUploadValidator
{
    public void Validate(Stream stream, string fileName, long maxSizeBytes = 20L * 1024 * 1024)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("File name is required.", nameof(fileName));

        var extension = Path.GetExtension(fileName);
        if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Only .docx files are supported.");

        if (stream.CanSeek && stream.Length > maxSizeBytes)
            throw new InvalidOperationException($"File exceeds the maximum supported size of {maxSizeBytes} bytes.");
    }
}
