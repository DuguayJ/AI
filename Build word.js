#!/usr/bin/env node
/**
 * Universal Document Builder — Word (.docx) generator
 * Usage: node build_word.js '<json>'
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, HeadingLevel, LevelFormat, TableOfContents,
  PageBreak, TabStopType, TabStopPosition,
} = require("docx");

// ─── Brand Palette ────────────────────────────────────────────────────────────
const C = {
  darkBlue:   "1F3864",
  midBlue:    "2E5FA3",
  lightBlue:  "D6E4F7",
  accentBlue: "4A90D9",
  darkGray:   "2C2C2C",
  medGray:    "595959",
  lightGray:  "F4F6F9",
  borderGray: "D0D7E2",
  white:      "FFFFFF",
  warnYellow: "FEF9EC",
  warnBorder: "F0C040",
  errorPink:  "FDF0F0",
  errorBorder:"E07070",
  successGreen:"F0FAF4",
  successBdr: "4CAF82",
  teal:       "1A7F8E",
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const px = (pt) => pt * 2; // half-points
const rgb = (hex) => ({ r: parseInt(hex.slice(0,2),16), g: parseInt(hex.slice(2,4),16), b: parseInt(hex.slice(4,6),16) });

function border(color = C.borderGray, size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}
function noBorder() {
  return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
}
function allBorders(color, size) {
  const b = border(color, size);
  return { top: b, bottom: b, left: b, right: b };
}
function noAllBorders() {
  const b = noBorder();
  return { top: b, bottom: b, left: b, right: b };
}

// ─── Reusable paragraph factories ─────────────────────────────────────────────
function spacer(before = 80, after = 80) {
  return new Paragraph({ spacing: { before, after }, children: [new TextRun("")] });
}

function divider(color = C.midBlue, size = 8) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    children: [new TextRun("")],
  });
}

function heading1(text, color = C.darkBlue) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, color, bold: true, size: px(15), font: "Arial" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.midBlue, space: 2 } },
  });
}

function heading2(text, color = C.midBlue) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, color, bold: true, size: px(12), font: "Arial" })],
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: px(10.5), font: "Arial", color: C.darkGray, ...opts })],
  });
}

function bulletItem(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: px(10.5), font: "Arial", color: C.darkGray })],
  });
}

// ─── Cover page ───────────────────────────────────────────────────────────────
function buildCoverPage(data) {
  const items = [];

  // Dark header bar (simulated via shaded paragraph)
  items.push(new Paragraph({
    spacing: { before: 0, after: 0 },
    shading: { fill: C.darkBlue, type: ShadingType.CLEAR },
    children: [new TextRun({ text: " ", size: px(2) })],
  }));

  items.push(spacer(600));

  // Document type label
  if (data.documentType) {
    items.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({
        text: data.documentType.toUpperCase(),
        font: "Arial", size: px(10), bold: true,
        color: C.accentBlue, characterSpacing: 80,
      })],
    }));
  }

  // Main title
  items.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({
      text: data.reportTitle,
      font: "Arial", size: px(28), bold: true, color: C.darkBlue,
    })],
  }));

  // Subtitle
  if (data.subtitle) {
    items.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 360 },
      children: [new TextRun({
        text: data.subtitle, font: "Arial", size: px(14), color: C.midBlue,
      })],
    }));
  }

  // Accent rule
  items.push(new Paragraph({
    spacing: { before: 0, after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: C.accentBlue, space: 1 } },
    children: [new TextRun("")],
  }));

  items.push(spacer(200));

  // Metadata grid (2-col table)
  const metaRows = [
    ["Organization", data.organization || "—"],
    ["Environment / Site", data.environment || "—"],
    ["Assessment Date", data.date || "—"],
    ["Classification", data.classification || "—"],
    ["Prepared By", data.preparedBy || "—"],
    ["Version", data.version || "1.0"],
  ];

  const metaTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 6760],
    borders: noAllBorders(),
    rows: metaRows.map(([label, val], i) =>
      new TableRow({
        children: [
          new TableCell({
            width: { size: 2600, type: WidthType.DXA },
            borders: noAllBorders(),
            shading: { fill: i % 2 === 0 ? C.lightGray : C.white, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 80 },
            children: [new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: px(9.5), bold: true, color: C.medGray })] })],
          }),
          new TableCell({
            width: { size: 6760, type: WidthType.DXA },
            borders: noAllBorders(),
            shading: { fill: i % 2 === 0 ? C.lightGray : C.white, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 80 },
            children: [new Paragraph({ children: [new TextRun({ text: val, font: "Arial", size: px(9.5), color: C.darkGray })] })],
          }),
        ],
      })
    ),
  });

  items.push(metaTable);
  items.push(spacer(300));

  // Status badge (color-coded)
  const statusColors = {
    "Completed":     { fill: C.successGreen, text: C.teal },
    "In Progress":   { fill: C.warnYellow,   text: "8A6000" },
    "Pending Review":{ fill: C.lightBlue,    text: C.midBlue },
    "Draft":         { fill: C.lightGray,    text: C.medGray },
  };
  const sc = statusColors[data.status] || { fill: C.lightGray, text: C.medGray };

  if (data.status) {
    items.push(new Table({
      width: { size: 2200, type: WidthType.DXA },
      columnWidths: [2200],
      borders: noAllBorders(),
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: 2200, type: WidthType.DXA },
          borders: { top: border(sc.text, 4), bottom: border(sc.text, 4), left: border(sc.text, 4), right: border(sc.text, 4) },
          shading: { fill: sc.fill, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `STATUS: ${data.status.toUpperCase()}`, font: "Arial", size: px(9), bold: true, color: sc.text })],
          })],
        })],
      })],
    }));
  }

  // Page break to content
  items.push(new Paragraph({ children: [new PageBreak()] }));
  return items;
}

// ─── Section builder ──────────────────────────────────────────────────────────
function buildSection(sec, listRefs) {
  const items = [];
  items.push(heading1(sec.title));

  for (const block of (sec.blocks || [])) {
    switch (block.type) {

      case "text":
        items.push(bodyText(block.content));
        break;

      case "bullets":
        for (const item of block.items || [])
          items.push(bulletItem(item, "bullets"));
        break;

      case "numbered":
        for (const item of block.items || [])
          items.push(bulletItem(item, "numbers"));
        break;

      case "callout": {
        const calloutStyle = {
          info:    { fill: C.lightBlue,   bdrColor: C.midBlue,     label: "ℹ INFO" },
          warning: { fill: C.warnYellow,  bdrColor: C.warnBorder,  label: "⚠ WARNING" },
          error:   { fill: C.errorPink,   bdrColor: C.errorBorder, label: "✕ CRITICAL" },
          success: { fill: C.successGreen,bdrColor: C.successBdr,  label: "✓ NOTE" },
        };
        const cs = calloutStyle[block.style] || calloutStyle.info;
        items.push(spacer(60, 40));
        items.push(new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [280, 9080],
          borders: noAllBorders(),
          rows: [new TableRow({
            children: [
              new TableCell({
                width: { size: 280, type: WidthType.DXA },
                borders: noAllBorders(),
                shading: { fill: cs.bdrColor, type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [new TextRun("")] })],
              }),
              new TableCell({
                width: { size: 9080, type: WidthType.DXA },
                borders: noAllBorders(),
                shading: { fill: cs.fill, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 160 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: cs.label, font: "Arial", size: px(8.5), bold: true, color: cs.bdrColor })] }),
                  new Paragraph({ children: [new TextRun({ text: block.content, font: "Arial", size: px(10), color: C.darkGray })] }),
                ],
              }),
            ],
          })],
        }));
        items.push(spacer(40, 60));
        break;
      }

      case "table": {
        const cols = block.columns || [];
        const rows = block.rows || [];
        const totalW = 9360;
        const colW = Math.floor(totalW / Math.max(cols.length, 1));
        const colWidths = cols.map((_, i) => i === cols.length - 1 ? totalW - colW * (cols.length - 1) : colW);

        const headerRow = new TableRow({
          tableHeader: true,
          children: cols.map((col, i) =>
            new TableCell({
              width: { size: colWidths[i], type: WidthType.DXA },
              borders: allBorders(C.midBlue, 4),
              shading: { fill: C.darkBlue, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [new TextRun({ text: col, font: "Arial", size: px(9.5), bold: true, color: C.white })],
              })],
            })
          ),
        });

        const dataRows = rows.map((row, ri) =>
          new TableRow({
            children: row.map((cell, ci) =>
              new TableCell({
                width: { size: colWidths[ci], type: WidthType.DXA },
                borders: allBorders(C.borderGray, 4),
                shading: { fill: ri % 2 === 0 ? C.white : C.lightGray, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: String(cell), font: "Arial", size: px(10), color: C.darkGray })] })],
              })
            ),
          })
        );

        items.push(spacer(60, 40));
        items.push(new Table({
          width: { size: totalW, type: WidthType.DXA },
          columnWidths: colWidths,
          rows: [headerRow, ...dataRows],
        }));
        items.push(spacer(40, 80));
        break;
      }

      case "key_value": {
        const kvRows = (block.items || []).map(([k, v], i) =>
          new TableRow({
            children: [
              new TableCell({
                width: { size: 2800, type: WidthType.DXA },
                borders: allBorders(C.borderGray, 4),
                shading: { fill: C.lightGray, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 140, right: 100 },
                children: [new Paragraph({ children: [new TextRun({ text: k, font: "Arial", size: px(10), bold: true, color: C.darkBlue })] })],
              }),
              new TableCell({
                width: { size: 6560, type: WidthType.DXA },
                borders: allBorders(C.borderGray, 4),
                shading: { fill: C.white, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: v, font: "Arial", size: px(10), color: C.darkGray })] })],
              }),
            ],
          })
        );
        items.push(new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 6560],
          rows: kvRows,
        }));
        items.push(spacer(80));
        break;
      }

      case "heading2":
        items.push(heading2(block.content));
        break;
    }
  }

  return items;
}

// ─── Header / Footer ──────────────────────────────────────────────────────────
function buildHeader(data) {
  return new Header({
    children: [
      new Paragraph({
        spacing: { before: 0, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 2 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: data.reportTitle, font: "Arial", size: px(8.5), bold: true, color: C.darkBlue }),
          new TextRun({ text: `\t${data.classification || "Confidential — Internal Use Only"}`, font: "Arial", size: px(8), color: C.medGray, italics: true }),
        ],
      }),
    ],
  });
}

function buildFooter(data) {
  return new Footer({
    children: [
      new Paragraph({
        spacing: { before: 80, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.borderGray, space: 2 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: `${data.organization || "Organization"} · ${data.date || ""}`, font: "Arial", size: px(8), color: C.medGray }),
          new TextRun({ text: "\tPage ", font: "Arial", size: px(8), color: C.medGray }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: px(8), color: C.midBlue, bold: true }),
          new TextRun({ text: " of ", font: "Arial", size: px(8), color: C.medGray }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: px(8), color: C.medGray }),
        ],
      }),
    ],
  });
}

// ─── Main builder ─────────────────────────────────────────────────────────────
function buildDocument(data) {
  const listRefs = {};

  const coverItems = buildCoverPage(data);

  // TOC (optional)
  const tocItems = data.includeToc ? [
    heading1("Table of Contents"),
    new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
    new Paragraph({ children: [new PageBreak()] }),
  ] : [];

  // Sections
  const sectionItems = [];
  for (const sec of (data.sections || [])) {
    sectionItems.push(...buildSection(sec, listRefs));
    sectionItems.push(spacer(120, 60));
  }

  const doc = new Document({
    numbering: {
      config: [
        { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 560, hanging: 280 } } } }] },
        { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 560, hanging: 280 } } } }] },
      ],
    },
    styles: {
      default: { document: { run: { font: "Arial", size: px(10.5), color: C.darkGray } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: px(15), bold: true, font: "Arial", color: C.darkBlue },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: px(12), bold: true, font: "Arial", color: C.midBlue },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1200, right: 1296, bottom: 1200, left: 1296 },
        },
      },
      headers: { default: buildHeader(data) },
      footers: { default: buildFooter(data) },
      children: [
        ...coverItems,
        ...tocItems,
        ...sectionItems,
      ],
    }],
  });

  return doc;
}

// ─── Entry point ──────────────────────────────────────────────────────────────
const raw = process.argv[2];
if (!raw) { console.error("No JSON passed"); process.exit(1); }

const data = JSON.parse(raw);
const outPath = process.argv[3] || "/mnt/user-data/outputs/document.docx";

const doc = buildDocument(data);
Packer.toBuffer(doc).then(buf => {
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buf);
  console.log("Written:", outPath);
}).catch(e => { console.error(e); process.exit(1); });
