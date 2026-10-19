namespace TemplateGenerator.Application.Contracts;

public interface IDocumentUploadValidator
{
    void Validate(Stream stream, string fileName, long maxSizeBytes = 20L * 1024 * 1024);
}
