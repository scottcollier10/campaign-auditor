#!/usr/bin/env node
// ============================================================
// campaign-audit-report.js — Standalone DOCX Generator
// ============================================================
// Usage: echo '{"analysis":{...},"metadata":{...}}' | node campaign-audit-report.js
// Output: Writes DOCX to /tmp/campaign-audit-YYYY-MM-DD.docx
//         Prints JSON to stdout: {"filepath":"/tmp/...","filename":"...","size":1234}
//
// Deploy: Place this file at /home/node/campaign-audit-report.js on your Render instance
//         (or wherever your n8n service has filesystem access)
// ============================================================

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require('docx');

// ---- READ STDIN ----
let inputData = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => { inputData += chunk; });
process.stdin.on('end', async () => {
  try {
    const input = JSON.parse(inputData);
    const analysis = input.analysis || {};
    const metadata = input.metadata || {};
    const buffer = await generateReport(analysis, metadata);

    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `campaign-audit-${dateStr}.docx`;
    const filepath = `/tmp/${filename}`;

    fs.writeFileSync(filepath, buffer);

    // Output result as JSON to stdout for n8n to capture
    process.stdout.write(JSON.stringify({
      filepath,
      filename,
      size: buffer.length,
      generated: true
    }));
  } catch (err) {
    process.stderr.write(`DOCX generation error: ${err.message}\n${err.stack}`);
    process.exit(1);
  }
});

// ---- REPORT GENERATOR ----

async function generateReport(analysis, metadata) {

  const C = {
    primary: "1A3A5C", secondary: "2C7BE5", accent: "E8F4FD",
    success: "00A854", warning: "F5A623", danger: "E74C3C",
    lightGray: "F8F9FA", mediumGray: "E9ECEF", darkGray: "495057",
    headerBg: "1A3A5C", headerText: "FFFFFF", tableBorder: "DEE2E6"
  };

  const border = { style: BorderStyle.SINGLE, size: 1, color: C.tableBorder };
  const borders = { top: border, bottom: border, left: border, right: border };
  const noBorder = { style: BorderStyle.NONE, size: 0 };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
  const pad = { top: 60, bottom: 60, left: 120, right: 120 };

  // ---- Helpers ----

  function sc(score) {
    if (!score) return C.darkGray;
    if (score.startsWith('A')) return C.success;
    if (score.startsWith('B')) return C.secondary;
    if (score.startsWith('C')) return C.warning;
    return C.danger;
  }

  function hCell(text, w) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA },
      shading: { fill: C.headerBg, type: ShadingType.CLEAR }, margins: pad, verticalAlign: "center",
      children: [new Paragraph({ alignment: AlignmentType.LEFT,
        children: [new TextRun({ text, bold: true, color: C.headerText, font: "Arial", size: 18 })] })]
    });
  }

  function dCell(text, w, opts = {}) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA },
      shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
      margins: pad, verticalAlign: "center",
      children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text), bold: opts.bold || false,
          color: opts.color || C.darkGray, font: "Arial", size: opts.size || 18 })] })]
    });
  }

  function heading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 160 },
      children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: C.primary })]
    });
  }

  function headingWithScore(text, score) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 160 },
      children: [
        new TextRun({ text: text + "  ", font: "Arial", size: 28, bold: true, color: C.primary }),
        new TextRun({ text: `Rating: ${score}`, font: "Arial", size: 22, color: sc(score) })
      ]
    });
  }

  function bullet(text, color) {
    return new Paragraph({
      spacing: { after: 60 }, indent: { left: 360 },
      children: [
        new TextRun({ text: "\u2022  ", font: "Arial", size: 18, color: color || C.secondary }),
        new TextRun({ text, font: "Arial", size: 18, color: C.darkGray })
      ]
    });
  }

  function spacer(after) {
    return new Paragraph({ spacing: { after: after || 200 }, children: [] });
  }

  // ---- Build content ----

  const children = [];
  const portalId = metadata.portalId || 'N/A';
  const dateDisplay = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

  // === TITLE ===
  children.push(
    spacer(100),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 60 },
      children: [new TextRun({ text: "Campaign Audit Report", font: "Arial", size: 40, bold: true, color: C.primary })] }),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200 },
      children: [
        new TextRun({ text: `HubSpot Portal ${portalId}`, font: "Arial", size: 22, color: C.darkGray }),
        new TextRun({ text: "  |  ", font: "Arial", size: 22, color: C.mediumGray }),
        new TextRun({ text: dateDisplay, font: "Arial", size: 22, color: C.darkGray })
      ] })
  );

  // === OVERALL SCORE CARD ===
  const score = analysis.overallScore || 'N/A';
  const summary = analysis.executiveSummary || 'No summary available.';

  children.push(
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 7020],
      rows: [new TableRow({ children: [
        new TableCell({ borders: noBorders, width: { size: 2340, type: WidthType.DXA },
          shading: { fill: sc(score), type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 200, right: 200 }, verticalAlign: "center",
          children: [new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: score, font: "Arial", size: 56, bold: true, color: "FFFFFF" })] })] }),
        new TableCell({ borders: noBorders, width: { size: 7020, type: WidthType.DXA },
          shading: { fill: C.accent, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 240, right: 240 }, verticalAlign: "center",
          children: [
            new Paragraph({ spacing: { after: 80 },
              children: [new TextRun({ text: "Overall Score", font: "Arial", size: 22, bold: true, color: C.primary })] }),
            new Paragraph({
              children: [new TextRun({ text: summary, font: "Arial", size: 18, color: C.darkGray })] })
          ] })
      ] })]
    }),
    spacer(300)
  );

  // === BENCHMARK COMPARISON ===
  if (analysis.benchmarkComparison) {
    children.push(heading("Benchmark Comparison"));

    const bc = analysis.benchmarkComparison;
    const bw = [2800, 1640, 1640, 1640, 1640];
    const rows = [new TableRow({ children: [
      hCell("Metric", 2800), hCell("Yours", 1640), hCell("Industry Avg", 1640),
      hCell("Difference", 1640), hCell("Verdict", 1640)
    ] })];

    Object.entries(bc).forEach(([key, val], i) => {
      if (!val || typeof val !== 'object') return;
      const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());
      const diff = (val.yours - val.industry).toFixed(1);
      const sign = parseFloat(diff) >= 0 ? "+" : "";
      const isReverseMetric = key === 'bounceRate' || key === 'unsubRate';
      const isGood = isReverseMetric ? val.yours <= val.industry : val.yours >= val.industry;
      const fill = i % 2 === 0 ? C.lightGray : undefined;

      rows.push(new TableRow({ children: [
        dCell(label, 2800, { bold: true, fill }),
        dCell(`${val.yours}%`, 1640, { align: AlignmentType.CENTER, fill }),
        dCell(`${val.industry}%`, 1640, { align: AlignmentType.CENTER, fill }),
        dCell(`${sign}${diff}%`, 1640, { align: AlignmentType.CENTER, color: isGood ? C.success : C.danger, bold: true, fill }),
        dCell(isGood ? "Above" : "Below", 1640, { align: AlignmentType.CENTER, color: isGood ? C.success : C.warning, bold: true, fill })
      ] }));
    });

    children.push(
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: bw, rows }),
      spacer(300)
    );
  }

  // === EMAIL PERFORMANCE ===
  if (analysis.emailAnalysis) {
    const ea = analysis.emailAnalysis;
    children.push(headingWithScore("Email Performance", ea.overallRating || 'N/A'));

    if (ea.emails && ea.emails.length > 0) {
      const ew = [3200, 800, 1280, 1280, 2800];
      const eRows = [new TableRow({ children: [
        hCell("Email", 3200), hCell("Score", 800), hCell("Open %", 1280),
        hCell("Click %", 1280), hCell("Insight", 2800)
      ] })];

      ea.emails.forEach((email, i) => {
        const fill = i % 2 === 0 ? C.lightGray : undefined;
        eRows.push(new TableRow({ children: [
          dCell(email.name || 'Unknown', 3200, { bold: true, fill }),
          dCell(email.score || '-', 800, { align: AlignmentType.CENTER, bold: true, color: sc(email.score), fill }),
          dCell(email.openRate != null ? `${email.openRate}%` : '-', 1280, { align: AlignmentType.CENTER, fill }),
          dCell(email.clickRate != null ? `${email.clickRate}%` : '-', 1280, { align: AlignmentType.CENTER, fill }),
          dCell(email.insight || '', 2800, { fill, size: 16 })
        ] }));
      });

      children.push(
        new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: ew, rows: eRows }),
        spacer(120)
      );
    }

    if (ea.patterns && ea.patterns.length > 0) {
      children.push(new Paragraph({ spacing: { before: 120, after: 80 },
        children: [new TextRun({ text: "Key Patterns", font: "Arial", size: 22, bold: true, color: C.primary })] }));
      ea.patterns.forEach(p => children.push(bullet(p, C.secondary)));
    }

    children.push(spacer(200));
  }

  // === AUDIENCE HEALTH ===
  if (analysis.audienceHealth) {
    const ah = analysis.audienceHealth;
    children.push(headingWithScore("Audience Health", ah.score || 'N/A'));

    const stats = [
      { val: String(ah.totalContacts || 0), label: "Total Contacts", color: C.primary },
      { val: `${ah.activeRate || 0}%`, label: "Active Rate", color: (ah.activeRate || 0) > 30 ? C.success : C.danger },
      { val: `${ah.disengagedRate || 0}%`, label: "Disengaged", color: C.darkGray },
      { val: `${ah.optOutRate || 0}%`, label: "Opted Out", color: C.darkGray }
    ];

    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 2340, 2340, 2340],
      rows: [new TableRow({ children: stats.map((s, i) =>
        new TableCell({
          borders: noBorders, width: { size: 2340, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? C.accent : C.lightGray, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 160, right: 160 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
              children: [new TextRun({ text: s.val, font: "Arial", size: 32, bold: true, color: s.color })] }),
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: s.label, font: "Arial", size: 16, color: C.darkGray })] })
          ]
        })
      ) })]
    }), spacer(120));

    if (ah.findings && ah.findings.length > 0) {
      children.push(new Paragraph({ spacing: { before: 120, after: 80 },
        children: [new TextRun({ text: "Findings", font: "Arial", size: 22, bold: true, color: C.primary })] }));
      ah.findings.forEach(f => children.push(bullet(f, C.danger)));
    }

    if (ah.risks && ah.risks.length > 0) {
      children.push(new Paragraph({ spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: "Risks", font: "Arial", size: 22, bold: true, color: C.danger })] }));
      ah.risks.forEach(r => children.push(new Paragraph({
        spacing: { after: 60 }, indent: { left: 360 },
        children: [
          new TextRun({ text: "\u26A0  ", font: "Arial", size: 18 }),
          new TextRun({ text: r, font: "Arial", size: 18, color: C.darkGray })
        ]
      })));
    }

    children.push(spacer(200));
  }

  // === LIFECYCLE FUNNEL ===
  if (analysis.funnelAnalysis) {
    const fa = analysis.funnelAnalysis;
    children.push(headingWithScore("Lifecycle Funnel", fa.score || 'N/A'));

    if (fa.distribution) {
      const total = Object.values(fa.distribution).reduce((a, b) => a + b, 0);
      const fw = [3120, 1560, 4680];
      const fRows = [new TableRow({ children: [hCell("Stage", 3120), hCell("Count", 1560), hCell("% of Total", 4680)] })];

      Object.entries(fa.distribution).forEach(([stage, count], i) => {
        const pct = total > 0 ? ((count / total) * 100).toFixed(1) : '0.0';
        const fill = i % 2 === 0 ? C.lightGray : undefined;
        fRows.push(new TableRow({ children: [
          dCell(stage.charAt(0).toUpperCase() + stage.slice(1), 3120, { bold: true, fill }),
          dCell(String(count), 1560, { align: AlignmentType.CENTER, fill }),
          dCell(`${pct}%`, 4680, { fill })
        ] }));
      });

      children.push(
        new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: fw, rows: fRows }),
        spacer(120)
      );
    }

    if (fa.gaps && fa.gaps.length > 0) {
      children.push(new Paragraph({ spacing: { before: 120, after: 80 },
        children: [new TextRun({ text: "Funnel Gaps", font: "Arial", size: 22, bold: true, color: C.primary })] }));
      fa.gaps.forEach(g => children.push(bullet(g, C.warning)));
    }
  }

  // === RECOMMENDATIONS ===
  if (analysis.recommendations && analysis.recommendations.length > 0) {
    children.push(new Paragraph({ children: [new TextRun({ children: [new PageBreak()] })] }));
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 100, after: 200 },
      children: [new TextRun({ text: "Strategic Recommendations", font: "Arial", size: 28, bold: true, color: C.primary })]
    }));

    analysis.recommendations.forEach((rec, i) => {
      const pColor = rec.priority === 'high' ? C.danger : rec.priority === 'medium' ? C.warning : C.success;
      const effortLabel = rec.effort === 'quick_win' ? 'Quick Win' : rec.effort === 'project' ? 'Project' : rec.effort || 'Initiative';

      children.push(new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [600, 8760],
        rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, width: { size: 600, type: WidthType.DXA },
            shading: { fill: pColor, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: "center",
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: String(i + 1), font: "Arial", size: 28, bold: true, color: "FFFFFF" })] })] }),
          new TableCell({ borders: noBorders, width: { size: 8760, type: WidthType.DXA },
            shading: { fill: C.lightGray, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 200, right: 200 }, verticalAlign: "center",
            children: [new Paragraph({
              children: [
                new TextRun({ text: rec.title || 'Recommendation', font: "Arial", size: 22, bold: true, color: C.primary }),
                new TextRun({ text: `   ${(rec.priority || '').toUpperCase()} PRIORITY`, font: "Arial", size: 16, bold: true, color: pColor }),
                new TextRun({ text: `  |  ${effortLabel}`, font: "Arial", size: 16, color: C.darkGray })
              ] })] })
        ] })]
      }), spacer(60));

      if (rec.description) {
        children.push(new Paragraph({ spacing: { after: 60 }, indent: { left: 600 },
          children: [new TextRun({ text: rec.description, font: "Arial", size: 18, color: C.darkGray })] }));
      }

      if (rec.expectedImpact) {
        children.push(new Paragraph({ spacing: { after: 200 }, indent: { left: 600 },
          children: [
            new TextRun({ text: "Expected Impact: ", font: "Arial", size: 18, bold: true, color: C.secondary }),
            new TextRun({ text: rec.expectedImpact, font: "Arial", size: 18, italics: true, color: C.darkGray })
          ] }));
      }
    });
  }

  // Footer
  const tier = metadata.tier || 'HubSpot';
  const dataSource = metadata.dataSource || 'API + CSV';
  children.push(
    spacer(300),
    new Paragraph({ spacing: { before: 200 },
      children: [new TextRun({
        text: `Report generated by Campaign Auditor v1  |  Data source: ${dataSource}  |  ${tier}`,
        font: "Arial", size: 16, color: C.mediumGray, italics: true
      })] })
  );

  // === ASSEMBLE ===

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 20 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: C.primary },
          paragraph: { spacing: { before: 200, after: 160 }, outlineLevel: 0 } }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1200, right: 1440, bottom: 1200, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({ alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: "Campaign Audit Report  |  ", font: "Arial", size: 14, color: C.mediumGray }),
              new TextRun({ text: `Portal ${portalId}`, font: "Arial", size: 14, color: C.mediumGray, bold: true })
            ] })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({ alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Page ", font: "Arial", size: 14, color: C.mediumGray }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 14, color: C.mediumGray })
            ] })]
        })
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}
