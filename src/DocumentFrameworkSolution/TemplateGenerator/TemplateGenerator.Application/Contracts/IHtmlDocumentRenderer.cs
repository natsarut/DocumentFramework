using TemplateGenerator.Application.Models;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Application.Contracts;

public interface IHtmlDocumentRenderer
{
    string Render(WordConversionResult result);
    string Render(DocumentModel document);
    string Render(string html, string css);
}
