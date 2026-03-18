# OOXML Quick Reference for PowerPoint

Use this reference when editing slide XML directly via the unpack/pack workflow.

## Units

| Unit | Abbreviation | Conversion |
|------|-------------|------------|
| English Metric Unit | EMU | 914400 EMU = 1 inch |
| Point | pt | 72 pt = 1 inch |
| Half-point | half-pt | `sz="2400"` = 24pt (font size) |
| Hundredths of a point | centi-pt | `sz="1400"` = 14pt |

Common conversions:
- 1 inch = 914400 EMU = 72 pt
- 14pt font = `sz="1400"`
- 0.5 inch margin = 457200 EMU

## Slide Dimensions (16:9)

```
Width:  cx="12192000" EMU  (13.333 inches)
Height: cy="6858000"  EMU  (7.5 inches)
```

## Basic Slide Structure

```xml
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <!-- Shapes go here -->
    </p:spTree>
  </p:cSld>
</p:sld>
```

## Text Box

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="TextBox 1"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="685800" y="274320"/>      <!-- position -->
      <a:ext cx="10800000" cy="914400"/>   <!-- size -->
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="2400" b="1" dirty="0"/>
        <a:t>Title text here</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

## Text Formatting

```xml
<!-- Bold -->
<a:rPr b="1"/>

<!-- Italic -->
<a:rPr i="1"/>

<!-- Underline -->
<a:rPr u="sng"/>

<!-- Font size (14pt) and typeface -->
<a:rPr sz="1400" dirty="0">
  <a:latin typeface="Calibri"/>
</a:rPr>

<!-- Colour -->
<a:rPr>
  <a:solidFill>
    <a:srgbClr val="333333"/>
  </a:solidFill>
</a:rPr>

<!-- Theme colour -->
<a:rPr>
  <a:solidFill>
    <a:schemeClr val="dk1"/>
  </a:solidFill>
</a:rPr>
```

## Paragraphs

```xml
<!-- Left-aligned paragraph with spacing -->
<a:p>
  <a:pPr algn="l">
    <a:spcBef><a:spcPts val="600"/></a:spcBef>
    <a:spcAft><a:spcPts val="300"/></a:spcAft>
  </a:pPr>
  <a:r>
    <a:rPr lang="en-US" sz="1400" dirty="0"/>
    <a:t>Paragraph text</a:t>
  </a:r>
</a:p>

<!-- Bullet list -->
<a:p>
  <a:pPr lvl="0">
    <a:buChar char="•"/>
  </a:pPr>
  <a:r><a:t>First bullet</a:t></a:r>
</a:p>

<!-- Numbered list -->
<a:p>
  <a:pPr lvl="0">
    <a:buAutoNum type="arabicPeriod"/>
  </a:pPr>
  <a:r><a:t>First item</a:t></a:r>
</a:p>
```

## Shapes

```xml
<!-- Rectangle (MXA standard — no rounded corners) -->
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Rectangle 1"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="685800" y="1645920"/>
      <a:ext cx="3600000" cy="360000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill>
      <a:srgbClr val="1E2761"/>
    </a:solidFill>
    <a:ln w="12700">
      <a:solidFill>
        <a:srgbClr val="000000"/>
      </a:solidFill>
    </a:ln>
  </p:spPr>
</p:sp>
```

## Lines and Connectors

```xml
<!-- Horizontal line -->
<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="4" name="Connector 1"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="685800" y="1200000"/>
      <a:ext cx="10800000" cy="0"/>
    </a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="12700">
      <a:solidFill>
        <a:srgbClr val="CCCCCC"/>
      </a:solidFill>
    </a:ln>
  </p:spPr>
</p:cxnSp>
```

## Images

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="5" name="Picture 1"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="685800" y="1645920"/>
      <a:ext cx="4572000" cy="3429000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

Image must be added to `ppt/media/` and referenced in slide rels:
```xml
<!-- ppt/slides/_rels/slide1.xml.rels -->
<Relationship Id="rId2"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/image1.png"/>
```

## Tables

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="6" name="Table 1"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="685800" y="1645920"/>
    <a:ext cx="10800000" cy="3600000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="5400000"/>
          <a:gridCol w="5400000"/>
        </a:tblGrid>
        <a:tr h="500000">
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:rPr b="1"/><a:t>Header 1</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:rPr b="1"/><a:t>Header 2</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

## Presentation Structure Files

| File | Purpose |
|------|---------|
| `ppt/presentation.xml` | Slide list (`<p:sldIdLst>`), slide size |
| `ppt/slides/slide{N}.xml` | Individual slide content |
| `ppt/slides/_rels/slide{N}.xml.rels` | Slide relationships (layout, images) |
| `ppt/slideLayouts/` | Layout templates |
| `ppt/slideMasters/` | Master templates |
| `ppt/theme/theme1.xml` | Colours and fonts |
| `ppt/media/` | Images and media |
| `ppt/_rels/presentation.xml.rels` | Presentation relationships |
| `[Content_Types].xml` | MIME types for all files |

## Slide Operations

**Reorder slides:** Rearrange `<p:sldId>` elements in `ppt/presentation.xml`:
```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId2"/>
  <p:sldId id="257" r:id="rId3"/>
</p:sldIdLst>
```

**Delete slide:** Remove `<p:sldId>` from presentation.xml, remove relationship from `ppt/_rels/presentation.xml.rels`, remove override from `[Content_Types].xml`, delete slide file and its rels file.

**Add slide:** Create slide XML, add to `[Content_Types].xml`, add relationship in `ppt/_rels/presentation.xml.rels`, add `<p:sldId>` to `<p:sldIdLst>`, create slide rels file pointing to a layout.

## Common Pitfalls

- **Smart quotes:** Use XML entities (`&#x201C;` `&#x201D;` `&#x2018;` `&#x2019;`) not raw unicode
- **Whitespace:** Add `xml:space="preserve"` on `<a:t>` with leading/trailing spaces
- **defusedxml:** Use `defusedxml.minidom` not `xml.etree.ElementTree` (corrupts namespaces)
- **IDs must be unique:** `<p:cNvPr id="N">` — no two shapes on a slide share the same id
- **dirty attribute:** Add `dirty="0"` to `<a:rPr>` to prevent spell-check underlining
- **Geometry:** Always use `prst="rect"` for MXA. Never `roundRect`.
