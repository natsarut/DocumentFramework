# TemplateGenerator: Word-to-HTML and PDF Design

## 1. Objective

Build a .NET 10 web application in the `TemplateGenerator` project that can:

- import a Microsoft Word document (`.docx`)
- convert the document structure into semantic HTML + CSS
- render the result with Razor-based templates
- produce printable PDF output that preserves page layout and page breaks similar to Microsoft Word

This solution should be suitable for enterprise reporting and print-ready documents, not just casual webpage conversion.

## 2. Scope

### In scope
- Word document upload and parsing
- conversion of Word styles and structure to HTML/CSS
- Razor templating for dynamic document data
- HTML preview and PDF export
- page-break handling aligned with Word layout
- print-ready CSS for A4/Letter/Custom formats

### Out of scope for v1
- full WYSIWYG editing of Word content in browser
- live collaboration or document revision history
- OCR or image-based PDF reconstruction
- support for all Word features in the first release (advanced tracked changes, comments, macros, complex embedded objects)

## 3. Core Design Principles

1. Use Word as the source-of-truth for content structure.
2. Convert to semantic HTML instead of raw DOM dumps.
3. Keep page layout logic in CSS rather than JavaScript.
4. Use Razor as the dynamic templating layer on top of the converted document model.
5. Use browser-based print rendering for PDF output so pagination mimics Word more closely than server-only PDF conversion.
6. Separate parsing, normalization, templating, and rendering into explicit services.

## 4. Architecture

### 4.1 Solution layout

The existing `TemplateGenerator` app is a Blazor Web App using Fluent UI Blazor v5. The design should extend it with clear service boundaries:

- `TemplateGenerator` (Blazor app)
  - UI pages and Fluent UI components
  - upload, preview, template configuration screens
- `TemplateGenerator.Application`
  - document conversion use cases
  - DTOs and contracts
  - template rendering orchestration
- `TemplateGenerator.Infrastructure`
  - Word OpenXML parsing
  - HTML generation and CSS normalization
  - PDF rendering adapters
  - file storage and template persistence
- `TemplateGenerator.Domain`
  - document model, page model, style tokens, print settings

### 4.2 Main components

#### A. Document Upload & Validation
- accepts `.docx` files
- validates MIME type and file integrity
- stores a document package in temp or persistent storage
- creates a `DocumentPackage` record containing file metadata, template metadata, and output targets

#### B. Word Parser (OpenXML-based)
- uses `DocumentFormat.OpenXml` to inspect the Word document structure
- reads paragraphs, runs, tables, headings, lists, images, page break markers, and style IDs
- converts the document into a normalized intermediate model:
  - `DocumentModel`
  - `Section`
  - `Block`
  - `InlineRun`
  - `Table`
  - `ImageAsset`
  - `StyleRef`

#### C. Style Normalizer
- maps Word style definitions to CSS tokens
- resolves paragraph spacing, font size, line height, alignment, border, background color
- resolves heading hierarchy (`Heading1` ... `Heading9`)
- converts Word list numbering into HTML list semantics

#### D. HTML Generator
- creates semantic HTML from the normalized model
- outputs HTML fragments and page-structure wrappers
- attaches class names derived from Word styles and document settings

#### E. Razor Template Engine
- receives the generated document model and business data
- executes Razor templates to inject dynamic values into the HTML structure
- supports `.cshtml` layouts and partials for repeating sections

#### F. PDF Renderer
- renders the final HTML using a headless browser (Playwright preferred, Chromium-based)
- uses `@media print` CSS to match page dimensions and pagination
- prints to PDF with a fixed page size and margins

## 5. Processing Flow

1. User uploads `.docx` file via Fluent UI page.
2. Application validates and stores the file.
3. Word parser reads OOXML content and extracts paragraphs, tables, images, styles, and break markers.
4. Style normalizer resolves styles into document tokens.
5. HTML generator converts blocks into structured HTML.
6. Razor template engine merges document content with runtime data.
7. Final HTML is rendered in a preview pane.
8. Browser print pipeline exports PDF using print CSS.
9. PDF is stored or downloaded.

## 6. Word-to-HTML Conversion Strategy

### 6.1 Structural mapping

| Word element | HTML target | Notes |
|---|---|---|
| Paragraph | `<p>` | preserve spacing/alignment |
| Heading | `<h1>`...`<h6>` | map style IDs to heading levels |
| Table | `<table>` | preserve width, border, cell padding |
| List | `<ul>` / `<ol>` | support numbering levels |
| Image | `<img>` / `<figure>` | embed base64 or local asset |
| Break | `<div class="page-break">` or CSS `page-break` | use explicit pagination markers |
| Hyperlink | `<a>` | preserve target and styling |
| Run formatting | `<span>` | bold, italic, underline, colored text |

### 6.2 Important Word features to handle early

- paragraph spacing
- line spacing
- page breaks (`w:pPr/w:pageBreakBefore`, section breaks)
- tabs and alignment
- list numbering and indentation
- tables with row/col spans
- headers and footers (optional v1 support)
- bullet/numbered lists
- inline formatting (bold/italic/underline/strike)

### 6.3 Notable limitations for v1

- tracked changes and comments are not rendered as first-class output
- complex Word features like text boxes, smart art, and embedded Excel objects are deferred
- advanced layout logic that depends on Word's native engine is approximated with CSS rules

## 7. Page-Break Strategy for PDF Output

This is the most important requirement: the PDF should appear like a Word document, including page breaks and print layout.

### 7.1 Layout model

Use a fixed page model:
- document body split into printable pages
- each page has consistent margins and page size
- some blocks are explicitly marked as page starters or page terminators
- page break rules are encoded in CSS and structural tags

### 7.2 Recommended print CSS

```css
@page {
  size: A4 portrait;
  margin: 12mm 12mm 12mm 12mm;
  marks: crop cross;
}

html, body {
  margin: 0;
  padding: 0;
  font-family: "Segoe UI", sans-serif;
  background: white;
}

.document {
  width: 100%;
  max-width: 210mm;
  margin: 0 auto;
}

.page {
  page-break-after: always;
  break-after: page;
}

.section {
  break-inside: avoid;
}

.table,
.img,
.p,
.h1,
.h2,
.h3,
.list {
  break-inside: avoid;
}
```

### 7.3 Rules for Word-like pagination

- use `break-after: page` for explicit page breaks
- use `break-inside: avoid` for tables, lists, and large blocks
- avoid floating layout containers that cause unpredictable pagination
- use stable margins and padding that match Word print settings
- include a `PageSettings` model to support A4, Letter, legal, landscape, etc.

## 8. Data Model Design

```csharp
public class DocumentPackage
{
	public Guid Id { get; set; }
	public string FileName { get; set; }
	public string TemplateName { get; set; }
	public PageSettings PageSettings { get; set; }
	public DocumentModel Document { get; set; }
}

public class DocumentModel
{
	public List<Section> Sections { get; set; } = new();
	public List<StyleToken> Styles { get; set; } = new();
}

public class Section
{
	public string Id { get; set; }
	public List<Block> Blocks { get; set; } = new();
	public bool StartsOnNewPage { get; set; }
}

public class Block
{
	public BlockType Type { get; set; }
	public string? Html { get; set; }
	public string? Text { get; set; }
	public List<InlineRun> Runs { get; set; } = new();
	public StyleToken Style { get; set; } = new();
}

public class PageSettings
{
	public PaperSize Size { get; set; }
	public PageOrientation Orientation { get; set; }
	public double MarginTopMm { get; set; }
	public double MarginRightMm { get; set; }
	public double MarginBottomMm { get; set; }
	public double MarginLeftMm { get; set; }
}
```

## 9. Razor Templating Design

The templating layer should remain separate from the Word conversion layer.

### Example pattern

```cshtml
@page "/preview/{id:guid}"
@using TemplateGenerator.Application.Models

<div class="document">
	@foreach (var section in Model.Sections)
	{
		<section class="section">
			@foreach (var block in section.Blocks)
			{
				@Html.Raw(block.Html)
			}
		</section>
	}
</div>
```

### Template capabilities
- dynamic data binding
- loops and conditionals
- layout and partial components
- document-level variables like client name, number, date, title
- optional custom themes and print styles

## 10. PDF Export Strategy

### Preferred option: headless browser print

Use Playwright or Puppeteer to render the final HTML into PDF.

Benefits:
- better pagination consistency
- stronger CSS print support
- native handling of page breaks and margins
- closer match to browser rendering than pure .NET PDF libraries

### Alternative option

Use a .NET PDF library only if browser-based rendering is not acceptable for enterprise constraints. This is a fallback, not the preferred design.

## 11. Recommended Technology Choices

### .NET 10 + Blazor
- App remains in the current stack and aligns with the repo
- Fluent UI Blazor v5 provides a modern admin UI and upload flow

### OpenXML library
- robust for reading `.docx` XML structure
- suitable for conversion of paragraphs, tables, lists, styles and page metadata

### Razor templating
- use existing .NET and Blazor capabilities
- keep document generation dynamic and maintainable

### PDF rendering
- Playwright for browser-based print output
- CSS media print as the source of pagination rules

## 12. Proposed Project Structure

```text
src/DocumentFrameworkSolution/
  TemplateGenerator/
	TemplateGenerator/
	  Components/
	  Pages/
	  Services/
	  wwwroot/
	TemplateGenerator.Application/
	  Contracts/
	  Models/
	  UseCases/
	TemplateGenerator.Infrastructure/
	  Word/
	  Rendering/
	  Templates/
	TemplateGenerator.Domain/
	  Entities/
	  ValueObjects/
	  Enums/
```

## 13. Delivery Plan

### Phase 1: Proof of concept
- accept `.docx` file upload
- parse paragraphs and basic styles
- convert to HTML preview
- use a fixed print layout

### Phase 2: Core conversion
- support headings, tables, lists, images, hyperlinks
- convert Word styles into CSS tokens
- implement page-break detection and print CSS

### Phase 3: Templating integration
- add Razor rendering
- allow dynamic binding to business data
- add versioned templates and parameter models

### Phase 4: PDF export
- render HTML in browser
- export to PDF using print settings
- validate against sample Word documents

### Phase 5: Hardening
- security validation for uploaded files
- logging and telemetry
- large-document performance tuning
- production deployment and caching strategy

## 14. Key Risks and Mitigations

### Risk: Word styling is too complex to reproduce exactly
Mitigation: focus on the most common print styles first and treat advanced Word features as a compatibility layer.

### Risk: pagination differs from Word even with CSS
Mitigation: keep a style token map and allow per-template page settings to be tuned.

### Risk: PDF output varies by browser engine
Mitigation: pin Chromium version for the print service and keep a single rendering path.

### Risk: large documents cause memory or performance issues
Mitigation: stream document processing and render page-by-page where possible.

## 15. Recommendation

For this project, the best approach is:

- Blazor + Fluent UI v5 for the user experience
- .NET 10 for the application layer
- OpenXML for Word parsing
- Razor for dynamic template rendering
- Playwright Chromium for print/PDF generation
- CSS `@page`, `break-after`, and `break-inside` as the core page-layout engine

This combination gives the closest match to Microsoft Word formatting while preserving the architectural clarity expected for the Document Framework solution.
