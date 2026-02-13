#!/usr/bin/env node
// ============================================================
// campaign-audit-report.js — Standalone DOCX Generator
// ============================================================
// ZERO template literals — uses string concatenation only
// to survive any copy/paste/download path without corruption.
//
// Usage: echo '{"analysis":{...},"metadata":{...}}' | node campaign-audit-report.js
// Output: DOCX to /tmp, JSON result to stdout
// ============================================================

var fs = require('fs');
var docx = require('docx');

var Document = docx.Document;
var Packer = docx.Packer;
var Paragraph = docx.Paragraph;
var TextRun = docx.TextRun;
var Table = docx.Table;
var TableRow = docx.TableRow;
var TableCell = docx.TableCell;
var Header = docx.Header;
var Footer = docx.Footer;
var AlignmentType = docx.AlignmentType;
var HeadingLevel = docx.HeadingLevel;
var BorderStyle = docx.BorderStyle;
var WidthType = docx.WidthType;
var ShadingType = docx.ShadingType;
var PageNumber = docx.PageNumber;
var PageBreak = docx.PageBreak;

// ---- READ STDIN ----
var inputData = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', function(chunk) { inputData += chunk; });
process.stdin.on('end', function() {
  try {
    var input = JSON.parse(inputData);
    var analysis = input.analysis || {};
    var metadata = input.metadata || {};
    generateReport(analysis, metadata).then(function(buffer) {
      var dateStr = new Date().toISOString().slice(0, 10);
      var filename = 'campaign-audit-' + dateStr + '.docx';
      var filepath = '/tmp/' + filename;
      fs.writeFileSync(filepath, buffer);
      process.stdout.write(JSON.stringify({
        filepath: filepath,
        filename: filename,
        size: buffer.length,
        generated: true,
        base64: buffer.toString('base64')
      }));
    }).catch(function(err) {
      process.stderr.write('DOCX generation error: ' + err.message + '\n' + err.stack);
      process.exit(1);
    });
  } catch (err) {
    process.stderr.write('JSON parse error: ' + err.message);
    process.exit(1);
  }
});

// ---- REPORT GENERATOR ----

function generateReport(analysis, metadata) {

  var C = {
    primary: "1A3A5C", secondary: "2C7BE5", accent: "E8F4FD",
    success: "00A854", warning: "F5A623", danger: "E74C3C",
    lightGray: "F8F9FA", mediumGray: "E9ECEF", darkGray: "495057",
    headerBg: "1A3A5C", headerText: "FFFFFF", tableBorder: "DEE2E6"
  };

  var border = { style: BorderStyle.SINGLE, size: 1, color: C.tableBorder };
  var borders = { top: border, bottom: border, left: border, right: border };
  var noBorder = { style: BorderStyle.NONE, size: 0 };
  var noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
  var pad = { top: 60, bottom: 60, left: 120, right: 120 };

  function sc(score) {
    if (!score) return C.darkGray;
    if (score.startsWith('A')) return C.success;
    if (score.startsWith('B')) return C.secondary;
    if (score.startsWith('C')) return C.warning;
    return C.danger;
  }

  function hCell(text, w) {
    return new TableCell({
      borders: borders, width: { size: w, type: WidthType.DXA },
      shading: { fill: C.headerBg, type: ShadingType.CLEAR }, margins: pad, verticalAlign: "center",
      children: [new Paragraph({ alignment: AlignmentType.LEFT,
        children: [new TextRun({ text: text, bold: true, color: C.headerText, font: "Arial", size: 18 })] })]
    });
  }

  function dCell(text, w, opts) {
    opts = opts || {};
    return new TableCell({
      borders: borders, width: { size: w, type: WidthType.DXA },
      shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
      margins: pad, verticalAlign: "center",
      children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text), bold: opts.bold || false,
          color: opts.color || C.darkGray, font: "Arial", size: opts.size || 18 })] })]
    });
  }

  function sectionHeading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 160 },
      children: [new TextRun({ text: text, font: "Arial", size: 28, bold: true, color: C.primary })]
    });
  }

  function sectionHeadingWithScore(text, score) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 160 },
      children: [
        new TextRun({ text: text + "  ", font: "Arial", size: 28, bold: true, color: C.primary }),
        new TextRun({ text: "Rating: " + score, font: "Arial", size: 22, color: sc(score) })
      ]
    });
  }

  function bullet(text, color) {
    return new Paragraph({
      spacing: { after: 60 }, indent: { left: 360 },
      children: [
        new TextRun({ text: "\u2022  ", font: "Arial", size: 18, color: color || C.secondary }),
        new TextRun({ text: text, font: "Arial", size: 18, color: C.darkGray })
      ]
    });
  }

  function warningBullet(text) {
    return new Paragraph({
      spacing: { after: 60 }, indent: { left: 360 },
      children: [
        new TextRun({ text: "\u26A0  ", font: "Arial", size: 18 }),
        new TextRun({ text: text, font: "Arial", size: 18, color: C.darkGray })
      ]
    });
  }

  function spacer(after) {
    return new Paragraph({ spacing: { after: after || 200 }, children: [] });
  }

  function subHeading(text, color) {
    return new Paragraph({ spacing: { before: 120, after: 80 },
      children: [new TextRun({ text: text, font: "Arial", size: 22, bold: true, color: color || C.primary })] });
  }

  // ---- Build content ----

  var children = [];
  var portalId = (metadata && metadata.portalId) ? metadata.portalId : 'N/A';
  var dateDisplay = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

  // === TITLE ===
  children.push(spacer(100));
  children.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 60 },
    children: [new TextRun({ text: "Campaign Audit Report", font: "Arial", size: 40, bold: true, color: C.primary })] }));
  children.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200 },
    children: [
      new TextRun({ text: "HubSpot Portal " + portalId, font: "Arial", size: 22, color: C.darkGray }),
      new TextRun({ text: "  |  ", font: "Arial", size: 22, color: C.mediumGray }),
      new TextRun({ text: dateDisplay, font: "Arial", size: 22, color: C.darkGray })
    ] }));

  // === OVERALL SCORE CARD ===
  var score = analysis.overallScore || 'N/A';
  var summary = analysis.executiveSummary || 'No summary available.';

  children.push(new Table({
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
  }));
  children.push(spacer(300));

  // === BENCHMARK COMPARISON ===
  if (analysis.benchmarkComparison) {
    children.push(sectionHeading("Benchmark Comparison"));

    var bc = analysis.benchmarkComparison;
    var bw = [2800, 1640, 1640, 1640, 1640];
    var bRows = [new TableRow({ children: [
      hCell("Metric", 2800), hCell("Yours", 1640), hCell("Industry Avg", 1640),
      hCell("Difference", 1640), hCell("Verdict", 1640)
    ] })];

    var bcKeys = Object.keys(bc);
    for (var bi = 0; bi < bcKeys.length; bi++) {
      var key = bcKeys[bi];
      var val = bc[key];
      if (!val || typeof val !== 'object') continue;
      var label = key.replace(/([A-Z])/g, ' $1').replace(/^./, function(s) { return s.toUpperCase(); });
      var diff = (val.yours - val.industry).toFixed(1);
      var sign = parseFloat(diff) >= 0 ? "+" : "";
      var isReverseMetric = key === 'bounceRate' || key === 'unsubRate';
      var isGood = isReverseMetric ? val.yours <= val.industry : val.yours >= val.industry;
      var fill = bi % 2 === 0 ? C.lightGray : undefined;

      bRows.push(new TableRow({ children: [
        dCell(label, 2800, { bold: true, fill: fill }),
        dCell(val.yours + "%", 1640, { align: AlignmentType.CENTER, fill: fill }),
        dCell(val.industry + "%", 1640, { align: AlignmentType.CENTER, fill: fill }),
        dCell(sign + diff + "%", 1640, { align: AlignmentType.CENTER, color: isGood ? C.success : C.danger, bold: true, fill: fill }),
        dCell(isGood ? "Above" : "Below", 1640, { align: AlignmentType.CENTER, color: isGood ? C.success : C.warning, bold: true, fill: fill })
      ] }));
    }

    children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: bw, rows: bRows }));
    children.push(spacer(300));
  }

  // === EMAIL PERFORMANCE ===
  if (analysis.emailAnalysis) {
    var ea = analysis.emailAnalysis;
    children.push(sectionHeadingWithScore("Email Performance", ea.overallRating || 'N/A'));

    if (ea.emails && ea.emails.length > 0) {
      var ew = [3200, 800, 1280, 1280, 2800];
      var eRows = [new TableRow({ children: [
        hCell("Email", 3200), hCell("Score", 800), hCell("Open %", 1280),
        hCell("Click %", 1280), hCell("Insight", 2800)
      ] })];

      for (var ei = 0; ei < ea.emails.length; ei++) {
        var email = ea.emails[ei];
        var eFill = ei % 2 === 0 ? C.lightGray : undefined;
        var openStr = email.openRate != null ? email.openRate + "%" : "-";
        var clickStr = email.clickRate != null ? email.clickRate + "%" : "-";
        eRows.push(new TableRow({ children: [
          dCell(email.name || 'Unknown', 3200, { bold: true, fill: eFill }),
          dCell(email.score || '-', 800, { align: AlignmentType.CENTER, bold: true, color: sc(email.score), fill: eFill }),
          dCell(openStr, 1280, { align: AlignmentType.CENTER, fill: eFill }),
          dCell(clickStr, 1280, { align: AlignmentType.CENTER, fill: eFill }),
          dCell(email.insight || '', 2800, { fill: eFill, size: 16 })
        ] }));
      }

      children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: ew, rows: eRows }));
      children.push(spacer(120));
    }

    if (ea.patterns && ea.patterns.length > 0) {
      children.push(subHeading("Key Patterns"));
      for (var pi = 0; pi < ea.patterns.length; pi++) {
        children.push(bullet(ea.patterns[pi], C.secondary));
      }
    }

    children.push(spacer(200));
  }

  // === AUDIENCE HEALTH ===
  if (analysis.audienceHealth) {
    var ah = analysis.audienceHealth;
    children.push(sectionHeadingWithScore("Audience Health", ah.score || 'N/A'));

    var activeRate = ah.activeRate || 0;
    var disengagedRate = ah.disengagedRate || 0;
    var optOutRate = ah.optOutRate || 0;

    var stats = [
      { val: String(ah.totalContacts || 0), label: "Total Contacts", color: C.primary },
      { val: activeRate + "%", label: "Active Rate", color: activeRate > 30 ? C.success : C.danger },
      { val: disengagedRate + "%", label: "Disengaged", color: C.darkGray },
      { val: optOutRate + "%", label: "Opted Out", color: C.darkGray }
    ];

    var statCells = [];
    for (var si = 0; si < stats.length; si++) {
      var s = stats[si];
      statCells.push(new TableCell({
        borders: noBorders, width: { size: 2340, type: WidthType.DXA },
        shading: { fill: si % 2 === 0 ? C.accent : C.lightGray, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
            children: [new TextRun({ text: s.val, font: "Arial", size: 32, bold: true, color: s.color })] }),
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: s.label, font: "Arial", size: 16, color: C.darkGray })] })
        ]
      }));
    }

    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 2340, 2340, 2340],
      rows: [new TableRow({ children: statCells })]
    }));
    children.push(spacer(120));

    if (ah.findings && ah.findings.length > 0) {
      children.push(subHeading("Findings"));
      for (var fi = 0; fi < ah.findings.length; fi++) {
        children.push(bullet(ah.findings[fi], C.danger));
      }
    }

    if (ah.risks && ah.risks.length > 0) {
      children.push(subHeading("Risks", C.danger));
      for (var ri = 0; ri < ah.risks.length; ri++) {
        children.push(warningBullet(ah.risks[ri]));
      }
    }

    children.push(spacer(200));
  }

  // === LIFECYCLE FUNNEL ===
  if (analysis.funnelAnalysis) {
    var fa = analysis.funnelAnalysis;
    children.push(sectionHeadingWithScore("Lifecycle Funnel", fa.score || 'N/A'));

    if (fa.distribution) {
      var distKeys = Object.keys(fa.distribution);
      var total = 0;
      for (var ti = 0; ti < distKeys.length; ti++) {
        total += fa.distribution[distKeys[ti]];
      }

      var fw = [3120, 1560, 4680];
      var fRows = [new TableRow({ children: [hCell("Stage", 3120), hCell("Count", 1560), hCell("% of Total", 4680)] })];

      for (var di = 0; di < distKeys.length; di++) {
        var stage = distKeys[di];
        var count = fa.distribution[stage];
        var pct = total > 0 ? ((count / total) * 100).toFixed(1) : '0.0';
        var fFill = di % 2 === 0 ? C.lightGray : undefined;
        fRows.push(new TableRow({ children: [
          dCell(stage.charAt(0).toUpperCase() + stage.slice(1), 3120, { bold: true, fill: fFill }),
          dCell(String(count), 1560, { align: AlignmentType.CENTER, fill: fFill }),
          dCell(pct + "%", 4680, { fill: fFill })
        ] }));
      }

      children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: fw, rows: fRows }));
      children.push(spacer(120));
    }

    if (fa.gaps && fa.gaps.length > 0) {
      children.push(subHeading("Funnel Gaps"));
      for (var gi = 0; gi < fa.gaps.length; gi++) {
        children.push(bullet(fa.gaps[gi], C.warning));
      }
    }
  }

  // === RECOMMENDATIONS ===
  if (analysis.recommendations && analysis.recommendations.length > 0) {
    children.push(new Paragraph({ children: [new TextRun({ children: [new PageBreak()] })] }));
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 100, after: 200 },
      children: [new TextRun({ text: "Strategic Recommendations", font: "Arial", size: 28, bold: true, color: C.primary })]
    }));

    for (var rci = 0; rci < analysis.recommendations.length; rci++) {
      var rec = analysis.recommendations[rci];
      var pColor = rec.priority === 'high' ? C.danger : rec.priority === 'medium' ? C.warning : C.success;
      var effortLabel = rec.effort === 'quick_win' ? 'Quick Win' : rec.effort === 'project' ? 'Project' : (rec.effort || 'Initiative');
      var priorityText = (rec.priority || '').toUpperCase();

      children.push(new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [600, 8760],
        rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, width: { size: 600, type: WidthType.DXA },
            shading: { fill: pColor, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: "center",
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: String(rci + 1), font: "Arial", size: 28, bold: true, color: "FFFFFF" })] })] }),
          new TableCell({ borders: noBorders, width: { size: 8760, type: WidthType.DXA },
            shading: { fill: C.lightGray, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 200, right: 200 }, verticalAlign: "center",
            children: [new Paragraph({
              children: [
                new TextRun({ text: rec.title || 'Recommendation', font: "Arial", size: 22, bold: true, color: C.primary }),
                new TextRun({ text: "   " + priorityText + " PRIORITY", font: "Arial", size: 16, bold: true, color: pColor }),
                new TextRun({ text: "  |  " + effortLabel, font: "Arial", size: 16, color: C.darkGray })
              ] })] })
        ] })]
      }));
      children.push(spacer(60));

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
    }
  }

  // Footer
  var tier = (metadata && metadata.tier) ? metadata.tier : 'HubSpot';
  var dataSource = (metadata && metadata.dataSource) ? metadata.dataSource : 'API + CSV';
  children.push(spacer(300));
  children.push(new Paragraph({ spacing: { before: 200 },
    children: [new TextRun({
      text: "Report generated by Campaign Auditor v1  |  Data source: " + dataSource + "  |  " + tier,
      font: "Arial", size: 16, color: C.mediumGray, italics: true
    })] }));

  // === ASSEMBLE ===

  var doc = new Document({
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
              new TextRun({ text: "Portal " + portalId, font: "Arial", size: 14, color: C.mediumGray, bold: true })
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
      children: children
    }]
  });

  return Packer.toBuffer(doc);
}
