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
| w:lastRenderedPageBreak | Pagination hint from Word's previous layout | Ignored | Reflowed by the HTML renderer |

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
| w:keepNext | break-after: avoid-page |
| w:keepLines | break-inside: avoid-page |
| w:pageBreakBefore | break-before: page |

`docDefaults` paragraph spacing/line-height is applied before paragraph styles and direct paragraph properties so later definitions override earlier ones like in Word.

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
| w:sz | font-size for standard scripts |
| w:szCs | font-size for complex scripts, including Thai |
| w:color | color |
| w:highlight | background-color |
| w:spacing | letter-spacing |

ฟอนต์และรูปแบบ run จะเลือกตามกลุ่มอักษรของข้อความ (`ascii`/`hAnsi`, `eastAsia`, `cs`) และรองรับฟอนต์ธีม (`*Theme`) โดยไล่ค่าจาก run, character/paragraph style ที่สืบทอดผ่าน `basedOn` และ `docDefaults` ขนาดของ complex script รวมถึงภาษาไทยใช้ `w:szCs` แยกจาก `w:sz`; หาก run มีหลายกลุ่มอักษร จะกำหนดขนาด complex script เฉพาะช่วงอักษรที่เกี่ยวข้อง ระยะห่างและการจัดย่อหน้าก็สืบทอดจาก paragraph styles เช่นกัน ชื่อฟอนต์ Unicode และช่องว่างซ้ำในข้อความจะถูกเก็บไว้ใน HTML/CSS

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

# Page Size in OOXML (WordprocessingML)

## Overview

ขนาดหน้ากระดาษ (Page Size) ใน Microsoft Word ถูกกำหนดโดย Element `w:pgSz` ซึ่งอยู่ภายใน Element `w:sectPr` (Section Properties)

เนื่องจาก Word รองรับหลาย Section ภายในเอกสารเดียว แต่ละ Section สามารถกำหนดขนาดหน้ากระดาษ (Page Size) และการวางแนว (Orientation) ได้อย่างอิสระ

---

# XML Structure

```xml
<w:sectPr>

    <w:pgSz
        w:w="11906"
        w:h="16838"/>

</w:sectPr>
```

---

# Element Definition

| Element / Attribute | Description |
|---------------------|-------------|
| `w:sectPr` | Section Properties |
| `w:pgSz` | Page Size Definition |
| `w:w` | Page Width |
| `w:h` | Page Height |
| `w:orient` | Page Orientation |

---

# Measurement Unit

ค่าของ `w:w` และ `w:h` ใช้หน่วย **Twips** (Twentieths of a Point)

```text
1 Point = 20 Twips
1 Inch = 72 Points
1 Inch = 1440 Twips
```

---

# Common Page Sizes

## A4 Portrait

```xml
<w:pgSz
    w:w="11906"
    w:h="16838"/>
```

### Dimensions

```text
Width  = 11906 / 1440 = 8.27 in
Height = 16838 / 1440 = 11.69 in
```

Equivalent size:

```text
210 mm × 297 mm
```

---

## A4 Landscape

```xml
<w:pgSz
    w:w="16838"
    w:h="11906"
    w:orient="landscape"/>
```

Equivalent size:

```text
297 mm × 210 mm
```

---

## Letter

```xml
<w:pgSz
    w:w="12240"
    w:h="15840"/>
```

Equivalent size:

```text
8.5" × 11"
```

---

## Legal

```xml
<w:pgSz
    w:w="12240"
    w:h="20160"/>
```

Equivalent size:

```text
8.5" × 14"
```

---

# Page Orientation

Orientation ถูกกำหนดด้วย Attribute `w:orient`

## Portrait (Default)

```xml
<w:pgSz
    w:w="11906"
    w:h="16838"/>
```

หรือ

```xml
<w:pgSz
    w:w="11906"
    w:h="16838"
    w:orient="portrait"/>
```

---

## Landscape

```xml
<w:pgSz
    w:w="16838"
    w:h="11906"
    w:orient="landscape"/>
```

---

# HTML/CSS Mapping

## Standard A4

```css
@page {
    size: A4;
}
```

---

## Standard Letter

```css
@page {
    size: Letter;
}
```

---

## Custom Page Size

### OOXML

```xml
<w:pgSz
    w:w="11906"
    w:h="16838"/>
```

### HTML/CSS

```css
@page {
    size: 210mm 297mm;
}
```

---

# Relationship with Section Properties

`w:pgSz` เป็นส่วนหนึ่งของ `w:sectPr`

```xml
<w:sectPr>

    <w:pgSz/>

    <w:pgMar/>

    <w:cols/>

    <w:headerReference/>

    <w:footerReference/>

    <w:pgNumType/>

</w:sectPr>
```

---

# Related Section Settings

| Element | Purpose |
|----------|---------|
| `w:pgSz` | Page Size |
| `w:pgMar` | Page Margins |
| `w:cols` | Multi-column Layout |
| `w:headerReference` | Header Assignment |
| `w:footerReference` | Footer Assignment |
| `w:pgNumType` | Page Number Settings |
| `w:titlePg` | Different First Page |
| `w:evenAndOddHeaders` | Different Odd/Even Pages |

---

# Example: Complete Section Definition

```xml
<w:sectPr>

    <w:pgSz
        w:w="11906"
        w:h="16838"/>

    <w:pgMar
        w:top="1440"
        w:right="1440"
        w:bottom="1440"
        w:left="1440"/>

    <w:headerReference
        w:type="default"
        r:id="rId5"/>

    <w:footerReference
        w:type="default"
        r:id="rId6"/>

</w:sectPr>
```

---

# Renderer Mapping Summary

| OOXML Element | Renderer Mapping |
|--------------|------------------|
| `w:sectPr` | Section Container |
| `w:pgSz` | Page Size |
| `w:w` | Width |
| `w:h` | Height |
| `w:orient` | Portrait / Landscape |
| `w:pgMar` | CSS Margin |
| `w:headerReference` | Header |
| `w:footerReference` | Footer |

แต่ละ section จะสร้าง CSS named page (`@page doc-section-N`) จากความกว้าง/ความสูงใน `w:pgSz` และขอบบน/ขวา/ล่าง/ซ้ายที่ระบุใน `w:pgMar` โดยแปลงหน่วย twips เป็นนิ้ว หาก Word ไม่ได้ระบุขอบบางด้าน จะไม่กำหนดค่าแทนเอง เพื่อให้ค่าหน้ากระดาษเริ่มต้นของการ export มีผลเฉพาะด้านนั้น การส่งออก PDF ต้องให้ CSS `@page` มีลำดับความสำคัญเหนือค่าขนาดและขอบกระดาษส่วนกลาง

---

# Recommendation

สำหรับ DOCX → HTML/PDF Renderer ควรใช้ `w:pgSz` เป็น Source of Truth สำหรับ

- Page Width
- Page Height
- Page Orientation

และใช้ร่วมกับ `w:pgMar` เพื่อคำนวณพื้นที่แสดงผล (Printable Area) ของแต่ละ Section อย่างถูกต้อง

---

# Rendering Word Page Margins in HTML

## Overview

ใน Word ระยะขอบกระดาษ (Page Margin) ถูกกำหนดด้วย Element `w:pgMar` ภายใน `w:sectPr`

```xml
<w:sectPr>

    <w:pgSz
        w:w="11906"
        w:h="16838"/>

    <w:pgMar
        w:top="1440"
        w:right="1440"
        w:bottom="1440"
        w:left="1440"/>

</w:sectPr>
```

> สำหรับ HTML ที่แสดงบนหน้าจอ ให้แปลงขนาดหน้าและค่า Margin เป็นความกว้าง/ความสูงและ **CSS Padding** ของ Page Container เพื่อจำลองแผ่นกระดาษของ Word ส่วนการพิมพ์/ส่งออก PDF ให้ใช้ `@page` margins และไม่ใช้ padding จำลองซ้ำ

---

## Page Margin Attributes

| Attribute | Description |
|------------|------------|
| `w:top` | Top Margin |
| `w:right` | Right Margin |
| `w:bottom` | Bottom Margin |
| `w:left` | Left Margin |
| `w:header` | Header Position |
| `w:footer` | Footer Position |
| `w:gutter` | Binding Margin |

---

## Unit Conversion

ค่าทั้งหมดใน `w:pgMar` ใช้หน่วย **Twips**

```text
1 inch = 1440 twips
1 inch = 25.4 mm
```

ตัวอย่าง

```xml
<w:pgMar
    w:top="1440"
    w:right="1440"
    w:bottom="1440"
    w:left="1440"/>
```

เท่ากับ

```text
Top    = 25.4 mm
Right  = 25.4 mm
Bottom = 25.4 mm
Left   = 25.4 mm
```

---

## HTML Mapping

### HTML/CSS

ตัวแปลงสร้าง `<section class="doc-section doc-section-N">` สำหรับแต่ละ section ของ Word โดยใช้ขนาดหน้าจาก `w:pgSz` และระยะขอบจาก `w:pgMar`:

```css
@media screen {
  .doc-section.doc-section-1 {
    box-sizing: border-box;
    width: 8.5in;
    min-height: 11in;
    padding-top: 1in;
    padding-right: 1in;
    padding-bottom: 1in;
    padding-left: 1in;
  }
}

@page doc-section-1 {
  size: 8.5in 11in;
  margin-top: 1in;
  margin-right: 1in;
  margin-bottom: 1in;
  margin-left: 1in;
}
```

บนหน้าจอใช้ CSS padding เพื่อให้เห็นพื้นที่ขอบกระดาษ ส่วนการพิมพ์และส่งออก PDF ใช้ `@page` margins โดยไม่ใช้ padding จำลองซ้ำ

---

## Printable Area

การคำนวณพื้นที่เนื้อหา

```text
Content Width
=
Page Width
− Left Margin
− Right Margin
```

ตัวอย่าง A4

```text
210
− 25.4
− 25.4
=
159.2 mm
```

---

## Reading Margins with C#

```csharp
var margin = sectionProperties.PageMargin;

double top =
    margin.Top!.Value * 25.4 / 1440;

double right =
    margin.Right!.Value * 25.4 / 1440;

double bottom =
    margin.Bottom!.Value * 25.4 / 1440;

double left =
    margin.Left!.Value * 25.4 / 1440;
```

สร้าง CSS

```csharp
var padding = $@"
padding:
    {top:F2}mm
    {right:F2}mm
    {bottom:F2}mm
    {left:F2}mm;";
```

---

## Header/Footer Position

Word แยกตำแหน่ง Header และ Footer จาก Margin

```xml
<w:pgMar
    w:top="1440"
    w:bottom="1440"
    w:header="720"
    w:footer="720"/>
```

```text
Top Margin      = 25.4 mm
Header Position = 12.7 mm

Bottom Margin   = 25.4 mm
Footer Position = 12.7 mm
```

---

## Browser PDF Rendering

```css
@page {
    size: A4;

    margin:
        25.4mm
        25.4mm
        25.4mm
        25.4mm;
}
```

---

## Best Practice

1. อ่านขนาดกระดาษจาก `w:pgSz`
2. อ่านระยะขอบจาก `w:pgMar`
3. สร้าง Page Container แทนกระดาษจริง
4. ใช้ `padding` แทน `margin`
5. กำหนด `box-sizing: border-box`
6. คำนวณ Printable Area จาก Page Size และ Margins

---

## Mapping Summary

| OOXML | HTML/CSS |
|---------|-----------|
| `w:pgSz` | Width / Height |
| `w:pgMar` | Padding |
| `w:top` | padding-top |
| `w:right` | padding-right |
| `w:bottom` | padding-bottom |
| `w:left` | padding-left |
| `w:header` | Header Position |
| `w:footer` | Footer Position |

> แนวทาง **Page Container + Padding + Box Sizing** ให้ผลลัพธ์ใกล้เคียง Microsoft Word มากที่สุดทั้งการ Preview บนเว็บและการ Export PDF

---

# Mapping ระหว่าง OOXML `<w:u>` และ CSS `text-decoration-style`

> หมายเหตุ: CSS ไม่สามารถจำลองรูปแบบ Underline ของ Microsoft Word ได้ครบทุกแบบ 100% บางค่าอาจต้องใช้ Border, SVG หรือ Custom Rendering เพื่อให้ได้ผลลัพธ์ใกล้เคียงที่สุด

| Word OOXML (`w:u`) | Word Appearance | CSS Mapping | Notes |
|-------------------|------------------|-------------|--------|
| `single` | เส้นใต้เดี่ยว | `text-decoration: underline;` | รองรับตรงที่สุด |
| `words` | ขีดเส้นใต้เฉพาะคำ | `text-decoration: underline;` | CSS ไม่มี equivalent โดยตรง |
| `double` | เส้นใต้คู่ | `text-decoration-line: underline; text-decoration-style: double;` | Browser สมัยใหม่รองรับ |
| `thick` | เส้นใต้หนา | `text-decoration-thickness: 2px;` | ใกล้เคียง |
| `dotted` | เส้นใต้จุด | `text-decoration-style: dotted;` | ใกล้เคียง |
| `dottedHeavy` | เส้นใต้จุดหนา | `text-decoration-style: dotted; text-decoration-thickness: 2px;` | ประมาณค่า |
| `dash` | เส้นใต้ขีด | `text-decoration-style: dashed;` | ใกล้เคียง |
| `dashedHeavy` | เส้นใต้ขีดหนา | `text-decoration-style: dashed; text-decoration-thickness: 2px;` | ประมาณค่า |
| `dashLong` | เส้นใต้ขีดยาว | `border-bottom: 1px dashed;` | CSS ไม่มีเส้นยาวแบบ Word |
| `dashLongHeavy` | เส้นใต้ขีดยาวหนา | `border-bottom: 2px dashed;` | ประมาณค่า |
| `dotDash` | จุด-ขีด | ไม่รองรับโดยตรง | ใช้ SVG หรือ Border Image |
| `dotDashHeavy` | จุด-ขีด (หนา) | ไม่รองรับโดยตรง | Custom Render |
| `dotDotDash` | จุด-จุด-ขีด | ไม่รองรับโดยตรง | Custom Render |
| `dotDotDashHeavy` | จุด-จุด-ขีด (หนา) | ไม่รองรับโดยตรง | Custom Render |
| `wave` | เส้นหยัก | `text-decoration-style: wavy;` | รองรับโดยตรง |
| `wavyHeavy` | เส้นหยักหนา | `text-decoration-style: wavy; text-decoration-thickness: 2px;` | ใกล้เคียง |
| `wavyDouble` | เส้นหยักคู่ | ไม่รองรับโดยตรง | SVG แนะนำ |
| `none` | ไม่มีเส้นใต้ | `text-decoration: none;` | ตรงกัน |
| `nil` | ยกเลิกเส้นใต้ทั้งหมด | `text-decoration: none;` | ตรงกัน |

---

## CSS Mapping ที่รองรับได้โดยตรง

```css
.single {
    text-decoration-line: underline;
}

.double {
    text-decoration-line: underline;
    text-decoration-style: double;
}

.dotted {
    text-decoration-line: underline;
    text-decoration-style: dotted;
}

.dashed {
    text-decoration-line: underline;
    text-decoration-style: dashed;
}

.wavy {
    text-decoration-line: underline;
    text-decoration-style: wavy;
}
```

---

## กลุ่มที่ Browser รองรับได้ดี

| Word Value | CSS Equivalent |
|------------|----------------|
| `single` | ✅ |
| `double` | ✅ |
| `dotted` | ✅ |
| `dash` | ✅ |
| `wave` | ✅ |
| `thick` | ✅ |
| `none` | ✅ |

---

## กลุ่มที่ต้องใช้ Custom Rendering

| Word Value |
|------------|
| `dotDash` |
| `dotDashHeavy` |
| `dotDotDash` |
| `dotDotDashHeavy` |
| `dashLong` |
| `dashLongHeavy` |
| `wavyDouble` |

แนวทางที่แนะนำ:

- SVG Underline
- Border Image
- Pseudo Element (`::after`)
- Canvas/SVG Renderer

---

## Recommendation สำหรับ DOCX → HTML Renderer

```text
single         → text-decoration
double         → text-decoration-style:double
dotted         → text-decoration-style:dotted
dash           → text-decoration-style:dashed
wave           → text-decoration-style:wavy
thick          → text-decoration-thickness
อื่น ๆ         → SVG Renderer
```

สำหรับ Renderer ที่ต้องการความเหมือน Word มากกว่า 95% ควร Render กลุ่ม `dotDash`, `dotDotDash`, `dashLong` และ `wavyDouble` ด้วย SVG แทน CSS เนื่องจาก CSS ยังไม่มี Mapping ที่ตรงกับรูปแบบของ Word อย่างสมบูรณ์

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