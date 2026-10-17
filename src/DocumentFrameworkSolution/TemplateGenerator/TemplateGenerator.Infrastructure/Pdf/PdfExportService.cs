using Microsoft.Playwright;
using TemplateGenerator.Application.Contracts;
using TemplateGenerator.Domain;

namespace TemplateGenerator.Infrastructure.Pdf;

public sealed class PdfExportService : IPdfExportService
{
    public async Task<byte[]> ExportAsync(string html, PageSettings pageSettings, CancellationToken cancellationToken = default)
    {
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = true
        });

        var page = await browser.NewPageAsync(new BrowserNewPageOptions
        {
            ViewportSize = new() { Width = 1440, Height = 1200 }
        });

        var printCss = BuildPageCss(pageSettings);

        await page.SetContentAsync(
            $"""
            <html>
              <head>
                <style>{printCss}</style>
              </head>
              <body>{html}</body>
            </html>
            """,
            new PageSetContentOptions { WaitUntil = WaitUntilState.NetworkIdle });

        var pdfBytes = await page.PdfAsync(new PagePdfOptions
        {
            PrintBackground = true,
            Landscape = pageSettings.Orientation == PageOrientation.Landscape,
            Format = pageSettings.Size switch
            {
                PaperSize.A4 => PaperFormat.A4,
                PaperSize.Legal => PaperFormat.Legal,
                _ => PaperFormat.Letter
            },
            Margin = new Margin
            {
                Top = $"{pageSettings.MarginTopMm}mm",
                Right = $"{pageSettings.MarginRightMm}mm",
                Bottom = $"{pageSettings.MarginBottomMm}mm",
                Left = $"{pageSettings.MarginLeftMm}mm"
            }
        });

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
