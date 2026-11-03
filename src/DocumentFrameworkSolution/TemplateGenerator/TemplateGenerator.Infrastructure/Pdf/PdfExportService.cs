using PuppeteerSharp;
using PuppeteerSharp.Media;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Pdf;

public sealed class PdfExportService : IPdfExportService
{
    public async Task<byte[]> ExportAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(pageSettings);
        cancellationToken.ThrowIfCancellationRequested();

        var browserFetcher = new BrowserFetcher();
        var installedBrowser = await browserFetcher.DownloadAsync();
        cancellationToken.ThrowIfCancellationRequested();

        await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions
        {
            Headless = true,
            ExecutablePath = installedBrowser.GetExecutablePath()
        });
        await using var page = await browser.NewPageAsync();
        await page.SetViewportAsync(new ViewPortOptions { Width = 1440, Height = 1200 });

        var printCss = BuildPageCss(pageSettings);

        await page.SetContentAsync(
            $"""
            <html>
              <head>
                <style>{printCss}</style>
              </head>
              <body>{html}</body>
            </html>
            """);
        await page.EvaluateExpressionAsync("document.fonts.ready");
        cancellationToken.ThrowIfCancellationRequested();

        var pdfBytes = await page.PdfDataAsync(new PdfOptions
        {
            PrintBackground = true,
            Landscape = pageSettings.Orientation == PageOrientation.Landscape,
            Format = pageSettings.Size switch
            {
                PaperSize.A4 => PaperFormat.A4,
                PaperSize.Legal => PaperFormat.Legal,
                _ => PaperFormat.Letter
            },
            MarginOptions = new MarginOptions
            {
                Top = $"{pageSettings.MarginTopMm}mm",
                Right = $"{pageSettings.MarginRightMm}mm",
                Bottom = $"{pageSettings.MarginBottomMm}mm",
                Left = $"{pageSettings.MarginLeftMm}mm"
            }
        });

        cancellationToken.ThrowIfCancellationRequested();
        return pdfBytes;
    }

    private static string BuildPageCss(PageSettings pageSettings)
    {
        var size = pageSettings.Size switch
        {
            PaperSize.A4 => "A4",
            PaperSize.Legal => "Legal",
            _ => "Letter"
        };

        var orientation = pageSettings.Orientation == PageOrientation.Landscape ? "landscape" : "portrait";

        return $@"
        @page {{
          size: {size} {orientation};
          margin: {pageSettings.MarginTopMm}mm {pageSettings.MarginRightMm}mm {pageSettings.MarginBottomMm}mm {pageSettings.MarginLeftMm}mm;
        }}

        body {{
          margin: 0;
          font-family: 'Segoe UI', sans-serif;
          color: #111;
          background: white;
        }}

        .document {{
          width: 100%;
        }}

        .page-break {{
          page-break-after: always;
          break-after: page;
        }}

        table, p, h1, h2, h3, h4, h5, h6, ul, ol {{
          break-inside: avoid;
        }}";
    }
}
