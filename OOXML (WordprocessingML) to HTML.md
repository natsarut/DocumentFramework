# OOXML (WordprocessingML) to HTML/CSS Mapping Reference

เอกสารนี้รวบรวม Mapping สำหรับการแปลง DOCX/OOXML ไปเป็น HTML/CSS/SVG โดยครอบคลุม

- Document Structure
- Paragraph
- Run Properties
- Table
- Field Codes
- Page Number
- Header/Footer
- TextBox
- Shape / AutoShape

---

# 1. Document Structure

| OOXML | ความหมาย | HTML | CSS |
|---------|----------|--------|--------|
| w:document | Root Document | html | - |
| w:body | เนื้อหา | body | - |
| w:sectPr | Section | section | page properties |
| w:p | Paragraph | p | margin,text-align |
| w:r | Run | span | font styles |
| w:t | Text | TextNode | - |
| w:hyperlink | Hyperlink | a | color,text-decoration |
| w:br | Line Break | br | - |
| w:cr | Carriage Return | br | - |
| w:tab | Tab | &nbsp; | white-space |
| w:lastRenderedPageBreak | Page Break | div.page-break | page-break-after |

---

# 2. Paragraph Properties (w:pPr)

## Alignment

| OOXML | CSS |
|---------|---------|
| left | text-align:left |
| center | text-align:center |
| right | text-align:right |
| both | text-align:justify |

### Example

```xml
<w:jc w:val="center"/>
```

```css
text-align:center;
```

---

## Indentation

| OOXML | CSS |
|---------|---------|
| w:left | margin-left |
| w:right | margin-right |
| w:firstLine | text-indent |
| w:hanging | hanging indent |

---

## Spacing

| OOXML | CSS |
|---------|---------|
| w:before | margin-top |
| w:after | margin-bottom |
| w:line | line-height |

---

## Border

| OOXML | CSS |
|---------|---------|
| w:pBdr/w:top | border-top |
| w:pBdr/w:bottom | border-bottom |
| w:pBdr/w:left | border-left |
| w:pBdr/w:right | border-right |

---

## Shading

| OOXML | CSS |
|---------|---------|
| w:shd | background-color |

---

# 3. Run Properties (w:rPr)

## Font Mapping

| OOXML | CSS |
|---------|---------|
| w:rFonts | font-family |
| w:sz | font-size |
| w:color | color |
| w:highlight | background-color |
| w:spacing | letter-spacing |

---

## Style Mapping

| OOXML | HTML/CSS |
|---------|---------|
| w:b | font-weight:bold |
| w:i | font-style:italic |
| w:u | text-decoration:underline |
| w:strike | line-through |
| w:dstrike | line-through |
| w:caps | text-transform:uppercase |
| w:smallCaps | font-variant:small-caps |
| w:vanish | display:none |

---

## Superscript / Subscript

| OOXML | HTML |
|---------|---------|
| superscript | sup |
| subscript | sub |

---

# 4. Styles

## Paragraph Styles

| Word Style | HTML |
|------------|------|
| Heading1 | h1 |
| Heading2 | h2 |
| Heading3 | h3 |
| Heading4 | h4 |
| Heading5 | h5 |
| Heading6 | h6 |
| Title | h1 |
| Subtitle | h2 |
| Normal | p |
| Quote | blockquote |

---

# 5. Lists

| OOXML | HTML |
|---------|---------|
| w:numPr | ol / ul |
| Bullet | ul |
| Numbering | ol |
| List Item | li |

---

# 6. Tables

| OOXML | HTML |
|---------|---------|
| w:tbl | table |
| w:tr | tr |
| w:tc | td |
| w:tblGrid | colgroup |
| w:gridCol | col |

## Cell Merge

| OOXML | HTML |
|---------|---------|
| w:gridSpan | colspan |
| w:vMerge | rowspan |

## Cell Style

| OOXML | CSS |
|---------|---------|
| w:tcBorders | border |
| w:vAlign | vertical-align |
| w:shd | background |

---

# 7. Images

| OOXML | HTML |
|---------|---------|
| w:drawing | img |
| wp:inline | img |
| wp:anchor | img |
| a:blip | src |

---

# 8. Page Layout

## Page Size

```xml
<w:pgSz/>
```

```css
@page {
  size:A4;
}
```

## Margin

```xml
<w:pgMar/>
```

```css
@page {
  margin:1in;
}
```

---

# 9. Fields

## Simple Field

```xml
<w:fldSimple w:instr="PAGE"/>
```

## Complex Field

```xml
<w:fldChar/>
<w:instrText/>
```

---

# 10. Page Number Fields

| Field | Mapping |
|---------|---------|
| PAGE | counter(page) |
| NUMPAGES | counter(pages) |
| SECTIONPAGES | section counter |

## Example

```html
Page
<span class="page"></span>
of
<span class="pages"></span>
```

```css
.page::after{
  content:counter(page);
}

.pages::after{
  content:counter(pages);
}
```

---

# 11. Date Fields

| Word Field | HTML |
|------------|------|
| DATE | data-field="DATE" |
| CREATEDATE | data-field="CREATEDATE" |
| SAVEDATE | data-field="SAVEDATE" |
| PRINTDATE | data-field="PRINTDATE" |

---

# 12. Document Properties Fields

| Word Field | HTML |
|------------|------|
| Title | data-prop="Title" |
| Subject | data-prop="Subject" |
| Keywords | data-prop="Keywords" |
| Company | data-prop="Company" |
| Comments | data-prop="Comments" |

---

# 13. File Information Fields

| Word Field | HTML |
|------------|------|
| FILENAME | data-field="FILENAME" |
| FILENAME \p | data-field="FILENAME_PATH" |

---

# 14. References

| Field | HTML |
|---------|---------|
| REF | hyperlink |
| PAGEREF | data-pageref |
| HYPERLINK | a |

---

# 15. TOC

| Word | HTML |
|---------|---------|
| TOC | nav.toc |

Example

```html
<nav class="toc">
    <ul>
        <li><a href="#h1">Chapter 1</a></li>
    </ul>
</nav>
```

---

# 16. Mail Merge

| Word Field | HTML |
|------------|------|
| MERGEFIELD CustomerName | {{CustomerName}} |

---

# 17. Footnote / Endnote

| OOXML | HTML |
|---------|---------|
| footnoteReference | sup > a |
| endnoteReference | sup > a |

---

# 18. Header/Footer

## Header Types

| OOXML | HTML |
|---------|---------|
| default | header |
| first | header.first-page |
| even | header.even-page |
| odd | header.odd-page |

---

## Header/Footer Elements

| OOXML | HTML |
|---------|---------|
| w:hdr | header |
| w:ftr | footer |
| headerReference | header |
| footerReference | footer |

---

## Common Header/Footer Fields

| Field | Mapping |
|---------|---------|
| PAGE | counter(page) |
| NUMPAGES | counter(pages) |
| DATE | data-field="DATE" |
| AUTHOR | data-field="AUTHOR" |
| FILENAME | data-field="FILENAME" |

---

# 19. TextBox

## Main Elements

| OOXML | HTML |
|---------|---------|
| v:textbox | div.textbox |
| wps:txbx | div.textbox |
| w:txbxContent | div.textbox-content |

---

## Position

| OOXML | CSS |
|---------|---------|
| positionH | left |
| positionV | top |
| wp:anchor | position:absolute |
| wp:inline | display:inline-block |

---

## Appearance

| OOXML | CSS |
|---------|---------|
| a:solidFill | background-color |
| a:ln | border |
| rot | transform:rotate |
| a:outerShdw | box-shadow |
| lIns/rIns/tIns/bIns | padding |

---

## Text

| OOXML | CSS |
|---------|---------|
| algn="ctr" | text-align:center |
| anchor="ctr" | align-items:center |
| vert="vert" | writing-mode:vertical-rl |

---

# 20. Shape & AutoShape

## Preset Geometries

| OOXML | Shape | HTML/SVG |
|------------|------------|--------------|
| rect | Rectangle | div |
| roundRect | Rounded Rectangle | div |
| ellipse | Circle/Ellipse | svg ellipse |
| line | Line | svg line |
| triangle | Triangle | svg polygon |
| diamond | Diamond | svg polygon |
| pentagon | Pentagon | svg polygon |
| hexagon | Hexagon | svg polygon |
| octagon | Octagon | svg polygon |
| parallelogram | Parallelogram | svg polygon |
| trapezoid | Trapezoid | svg polygon |
| chevron | Chevron | svg polygon |

---

## Arrow Shapes

| OOXML | Shape |
|---------|---------|
| rightArrow | Right Arrow |
| leftArrow | Left Arrow |
| upArrow | Up Arrow |
| downArrow | Down Arrow |
| leftRightArrow | Two-direction Arrow |

แนะนำ Render เป็น SVG Path

---

## FlowChart Shapes

| OOXML | Shape |
|---------|---------|
| flowChartProcess | Process |
| flowChartDecision | Decision |
| flowChartTerminator | Terminator |
| flowChartDocument | Document |
| flowChartData | Data |

---

## Shape Effects

| OOXML | CSS |
|---------|---------|
| a:solidFill | background-color |
| a:gradFill | linear-gradient |
| a:pattFill | pattern background |
| a:ln | border |
| a:outerShdw | box-shadow |
| a:glow | filter:drop-shadow |
| a:reflection | reflection effect |

---

## Rotation

Word uses:

```text
1 degree = 60000
```

Example

```xml
rot="1800000"
```

Result

```css
transform:rotate(30deg);
```

---

## Flip

| OOXML | CSS |
|---------|---------|
| flipH="1" | scaleX(-1) |
| flipV="1" | scaleY(-1) |

---

# Recommended Production Rendering Strategy

| Object Type | Recommended Renderer |
|-------------|----------------------|
| Text | HTML |
| Paragraph | HTML |
| Lists | HTML |
| Table | HTML |
| Image | HTML |
| Header/Footer | HTML |
| Field | HTML + Runtime |
| TextBox | HTML + CSS |
| Shape | SVG |
| AutoShape | SVG |
| FlowChart | SVG |
| Equation | MathML |
| Chart | SVG / Canvas |
| SmartArt | SVG |

---

# Coverage

เอกสารฉบับนี้ครอบคลุม OOXML ที่พบในเอกสาร Word ทั่วไปประมาณ 95%+

- Text
- Paragraph
- Styles
- Lists
- Table
- Images
- Fields
- Page Number
- TOC
- References
- Mail Merge
- Header/Footer
- TextBox
- Shape
- AutoShape
- FlowChart Shapes

เหมาะสำหรับใช้เป็น Reference ในการพัฒนา DOCX → HTML Renderer, DOCX Viewer, PDF Generator และ Document Conversion Platform

---

# Reference
- [Word processing](https://learn.microsoft.com/en-us/office/open-xml/word/overview)