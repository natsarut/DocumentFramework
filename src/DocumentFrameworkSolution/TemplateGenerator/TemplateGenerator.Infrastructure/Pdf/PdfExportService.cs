using PuppeteerSharp;
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

        await page.SetContentAsync(html);
        var hasWordSectionPageRules = html.Contains("@page doc-section-", StringComparison.OrdinalIgnoreCase);
        await page.AddStyleTagAsync(new AddTagOptions { Content = BuildPageCss(pageSettings, !hasWordSectionPageRules) });
        await page.EvaluateExpressionAsync("document.fonts.ready");
        cancellationToken.ThrowIfCancellationRequested();

        var pdfBytes = await page.PdfDataAsync(new PdfOptions
        {
            PrintBackground = true,
            PreferCSSPageSize = true
        });

        cancellationToken.ThrowIfCancellationRequested();
        return pdfBytes;
    }

    private static string BuildPageCss(PageSettings pageSettings, bool includeDefaultPageRule)
    {
        var pageRule = string.Empty;
        if (includeDefaultPageRule)
        {
            var size = pageSettings.Size switch
            {
                PaperSize.A4 => "A4",
                PaperSize.Legal => "Legal",
                _ => "Letter"
            };
            var orientation = pageSettings.Orientation == PageOrientation.Landscape ? "landscape" : "portrait";
            pageRule = $"@page {{ size: {size} {orientation}; margin: {pageSettings.MarginTopMm}mm {pageSettings.MarginRightMm}mm {pageSettings.MarginBottomMm}mm {pageSettings.MarginLeftMm}mm; }}";
        }

        return $@"
        {pageRule}

        body {{
          margin: 0;
          color: #111;
          background: white;
        }}

        .document {{
          width: 100%;
        }}

        .page-break {{
          page-break-after: always;
          break-after: page;
        }}";
    }
}
