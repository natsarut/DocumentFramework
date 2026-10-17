using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Contracts;

public interface IHtmlDocumentRenderer
{
    string Render(DocumentModel document);
}
