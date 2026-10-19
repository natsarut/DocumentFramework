using Microsoft.FluentUI.AspNetCore.Components;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Application.Services;
using TemplateGenerator.Client.Pages;
using TemplateGenerator.Components;
using TemplateGenerator.Infrastructure.Pdf;
using TemplateGenerator.Infrastructure.Rendering;
using TemplateGenerator.Infrastructure.Word;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();
builder.Services.AddFluentUIComponents();

builder.Services.AddScoped<IDocumentUploadValidator, DocumentUploadValidator>();
builder.Services.AddScoped<IWordDocumentParser, WordDocumentParser>();
builder.Services.AddScoped<IHtmlDocumentRenderer, HtmlDocumentRenderer>();
builder.Services.AddScoped<IPdfExportService, PdfExportService>();
builder.Services.AddScoped<IDocumentConversionService, DocumentConversionService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseWebAssemblyDebugging();
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(TemplateGenerator.Client._Imports).Assembly);

app.Run();
