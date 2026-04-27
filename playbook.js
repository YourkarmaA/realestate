const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Color Palette ──────────────────────────────────────────────────────────
const BRAND_BLUE   = "1A3C6B";
const ACCENT_TEAL  = "0E7C7B";
const LIGHT_BLUE   = "D6E4F7";
const LIGHT_TEAL   = "D0EDEC";
const LIGHT_GRAY   = "F5F5F5";
const MID_GRAY     = "E8E8E8";
const WHITE        = "FFFFFF";
const TEXT_DARK    = "1A1A1A";
const TEXT_GRAY    = "555555";
const ACCENT_GOLD  = "C8890A";
const LIGHT_GOLD   = "FFF4DC";

// ── Borders ────────────────────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder   = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ────────────────────────────────────────────────────────────────
function hr(color = "CCCCCC", thickness = 6) {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: thickness, color, space: 1 } },
    spacing: { before: 80, after: 80 }
  });
}

function spacer(pts = 120) {
  return new Paragraph({ children: [], spacing: { before: pts, after: 0 } });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, font: "Arial", color: WHITE, size: 36 })],
    shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
    spacing: { before: 240, after: 160 },
    indent: { left: 200, right: 200 }
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, font: "Arial", color: BRAND_BLUE, size: 28 })],
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT_TEAL, space: 1 } }
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, font: "Arial", color: ACCENT_TEAL, size: 24 })],
    spacing: { before: 200, after: 80 }
  });
}

function body(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 22, color: TEXT_DARK, ...options })],
    spacing: { before: 60, after: 60 },
    alignment: AlignmentType.JUSTIFIED
  });
}

function boldBody(text) {
  return body(text, { bold: true });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [new TextRun({ text, font: "Arial", size: 22, color: TEXT_DARK })],
    spacing: { before: 40, after: 40 }
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    children: [new TextRun({ text, font: "Arial", size: 22, color: TEXT_DARK })],
    spacing: { before: 40, after: 40 }
  });
}

function calloutBox(label, text, bgColor = LIGHT_BLUE, labelColor = BRAND_BLUE) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: allBorders,
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: label, bold: true, font: "Arial", size: 22, color: labelColor })],
                spacing: { before: 0, after: 60 }
              }),
              new Paragraph({
                children: [new TextRun({ text, font: "Arial", size: 21, color: TEXT_DARK })],
                spacing: { before: 0, after: 0 }
              })
            ]
          })
        ]
      })
    ]
  });
}

function twoColTable(headers, rows, col1Width = 2800, col2Width = 6560) {
  const totalW = col1Width + col2Width;
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: allBorders,
        width: { size: col1Width, type: WidthType.DXA },
        shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: headers[0], bold: true, font: "Arial", size: 22, color: WHITE })] })]
      }),
      new TableCell({
        borders: allBorders,
        width: { size: col2Width, type: WidthType.DXA },
        shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: headers[1], bold: true, font: "Arial", size: 22, color: WHITE })] })]
      })
    ]
  });

  const dataRows = rows.map((row, i) => new TableRow({
    children: [
      new TableCell({
        borders: allBorders,
        width: { size: col1Width, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? LIGHT_BLUE : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: row[0], bold: true, font: "Arial", size: 21, color: TEXT_DARK })] })]
      }),
      new TableCell({
        borders: allBorders,
        width: { size: col2Width, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? LIGHT_BLUE : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: row[1], font: "Arial", size: 21, color: TEXT_DARK })] })]
      })
    ]
  }));

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: [col1Width, col2Width],
    rows: [headerRow, ...dataRows]
  });
}

function threeColTable(headers, rows) {
  const w = [2600, 3380, 3380];
  const total = 9360;
  const mkCell = (text, bold, bg, color = TEXT_DARK) => new TableCell({
    borders: allBorders,
    width: { size: w[0], type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold, font: "Arial", size: 21, color })] })]
  });

  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: allBorders,
      width: { size: w[i], type: WidthType.DXA },
      shading: { fill: ACCENT_TEAL, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22, color: WHITE })] })]
    }))
  });

  const dataRows = rows.map((row, i) => new TableRow({
    children: row.map((cell, j) => new TableCell({
      borders: allBorders,
      width: { size: w[j], type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? LIGHT_TEAL : WHITE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 21, color: TEXT_DARK })] })]
    }))
  }));

  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: w,
    rows: [headerRow, ...dataRows]
  });
}

function metricCard(label, value, sub = "") {
  return new TableCell({
    borders: allBorders,
    width: { size: 2200, type: WidthType.DXA },
    shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
    margins: { top: 120, bottom: 120, left: 120, right: 120 },
    children: [
      new Paragraph({ children: [new TextRun({ text: value, bold: true, font: "Arial", size: 36, color: BRAND_BLUE })], alignment: AlignmentType.CENTER }),
      new Paragraph({ children: [new TextRun({ text: label, bold: true, font: "Arial", size: 19, color: TEXT_DARK })], alignment: AlignmentType.CENTER }),
      ...(sub ? [new Paragraph({ children: [new TextRun({ text: sub, font: "Arial", size: 18, color: TEXT_GRAY })], alignment: AlignmentType.CENTER })] : [])
    ]
  });
}

function metricsRow(cards) {
  const colW = Math.floor(9360 / cards.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: cards.map(() => colW),
    rows: [new TableRow({ children: cards })]
  });
}

function chapterCover(num, title, subtitle) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
            margins: { top: 400, bottom: 400, left: 400, right: 400 },
            children: [
              new Paragraph({ children: [new TextRun({ text: `CHAPTER ${num}`, font: "Arial", size: 20, color: ACCENT_GOLD, bold: true })], alignment: AlignmentType.CENTER }),
              new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 44, color: WHITE, bold: true })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } }),
              new Paragraph({ children: [new TextRun({ text: subtitle, font: "Arial", size: 22, color: "AACCEE" })], alignment: AlignmentType.CENTER })
            ]
          })
        ]
      })
    ]
  });
}

// ── Main Document ──────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25CB", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
          { level: 2, format: LevelFormat.BULLET, text: "\u25AA", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }
        ]
      },
      {
        reference: "numbers",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }
        ]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: TEXT_DARK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 36, bold: true, font: "Arial", color: WHITE }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: BRAND_BLUE }, paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial", color: ACCENT_TEAL }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "Real Estate Performance Marketing Playbook", font: "Arial", size: 18, color: TEXT_GRAY }),
              new TextRun({ text: "\t", font: "Arial", size: 18 }),
              new TextRun({ text: "Confidential | Internal Use", font: "Arial", size: 18, color: TEXT_GRAY, italics: true })
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT_TEAL, space: 1 } }
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "South City Realtors  |  Performance Marketing Division  \t", font: "Arial", size: 18, color: TEXT_GRAY }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: TEXT_GRAY })
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT_TEAL, space: 1 } }
          })
        ]
      })
    },
    children: [

      // ═══════════════════════════════════════════════════════════════════
      // COVER PAGE
      // ═══════════════════════════════════════════════════════════════════
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders,
                width: { size: 9360, type: WidthType.DXA },
                shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
                margins: { top: 800, bottom: 200, left: 600, right: 600 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "REAL ESTATE", font: "Arial", size: 28, color: ACCENT_GOLD, bold: true })], alignment: AlignmentType.CENTER }),
                  new Paragraph({ children: [new TextRun({ text: "PERFORMANCE MARKETING", font: "Arial", size: 56, color: WHITE, bold: true })], alignment: AlignmentType.CENTER, spacing: { before: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: "PLAYBOOK", font: "Arial", size: 72, color: WHITE, bold: true })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 } }),
                  new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT_GOLD } }, spacing: { before: 0, after: 200 } }),
                  new Paragraph({ children: [new TextRun({ text: "The Complete A-to-Z Guide for Paid Media in Real Estate", font: "Arial", size: 26, color: "AACCEE", italics: true })], alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: "Google Ads \u2022 Meta Ads \u2022 Tracking \u2022 Reporting \u2022 Optimization", font: "Arial", size: 22, color: "AACCEE" })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 600 } }),
                ]
              })
            ]
          })
        ]
      }),

      spacer(400),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4500, 4860],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: allBorders,
                width: { size: 4500, type: WidthType.DXA },
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "Prepared by", font: "Arial", size: 18, color: TEXT_GRAY })], spacing: { after: 40 } }),
                  new Paragraph({ children: [new TextRun({ text: "Performance Marketing Team", font: "Arial", size: 22, bold: true, color: BRAND_BLUE }) ]}),
                  new Paragraph({ children: [new TextRun({ text: "South City Realtors, Bengaluru", font: "Arial", size: 20, color: TEXT_DARK }) ]}),
                ]
              }),
              new TableCell({
                borders: allBorders,
                width: { size: 4860, type: WidthType.DXA },
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "Coverage Period", font: "Arial", size: 18, color: TEXT_GRAY })], spacing: { after: 40 } }),
                  new Paragraph({ children: [new TextRun({ text: "June 2025 \u2013 February 2026", font: "Arial", size: 22, bold: true, color: BRAND_BLUE }) ]}),
                  new Paragraph({ children: [new TextRun({ text: "Platforms: Google Ads | Meta Ads (Facebook & Instagram)", font: "Arial", size: 20, color: TEXT_DARK }) ]}),
                ]
              })
            ]
          })
        ]
      }),

      spacer(200),

      // KPI Strip
      metricsRow([
        metricCard("Monthly Leads", "300\u2013350", "vs 280 target"),
        metricCard("CPL Reduction", "\u221234%", "\u20B9820 \u2192 \u20B9540"),
        metricCard("Site-Visit Rate", "~18%", "Rental vertical"),
        metricCard("Quality Score", "5 \u2192 7", "Google Ads avg."),
      ]),

      spacer(300),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // INTRODUCTION
      // ═══════════════════════════════════════════════════════════════════
      h1("Introduction & How to Use This Playbook"),
      body("This playbook is the operational bible for performance marketing in a real estate agency context. It captures every tactic, decision, framework and lesson learned from running paid media campaigns across Google Ads and Meta Ads for property buyer leads, rental leads, and owner listing acquisitions in a competitive urban market."),
      spacer(80),
      body("Whether you are a new marketing executive joining the team, a freelancer taking over campaigns, or agency leadership seeking to understand what drives ROI, this document gives you everything you need to run, optimize, and scale paid media independently."),
      spacer(80),

      h2("Who This Playbook Is For"),
      bullet("Performance marketing executives managing Google and Meta campaigns"),
      bullet("Marketing managers who review dashboards and monthly reports"),
      bullet("Agency leadership evaluating budget allocation and ROI"),
      bullet("Freelancers or contractors temporarily running campaigns"),
      bullet("New hires onboarding to the paid media function"),

      spacer(80),
      h2("How This Playbook Is Organized"),
      body("The playbook is divided into 10 chapters, each dedicated to a distinct pillar of performance marketing. You can read it end-to-end for a full picture, or jump directly to the chapter most relevant to your immediate need."),
      spacer(80),

      twoColTable(
        ["Chapter", "What It Covers"],
        [
          ["Ch. 1 \u2013 Strategy & Architecture", "Campaign structure, objectives, budget allocation, and account organization"],
          ["Ch. 2 \u2013 Google Ads", "Search, Performance Max, Display, and Branded campaigns"],
          ["Ch. 3 \u2013 Meta Ads", "Lead Gen, Conversions, creative formats, audience targeting"],
          ["Ch. 4 \u2013 Owner Listing Campaigns", "Separate strategy for homeowner acquisition on Meta"],
          ["Ch. 5 \u2013 Retargeting", "Warm audience strategy across Google and Meta"],
          ["Ch. 6 \u2013 Tracking & Attribution", "GTM, Meta Pixel, UTM parameters, GA4"],
          ["Ch. 7 \u2013 Creative Strategy & Testing", "A/B testing, fatigue management, creative briefs"],
          ["Ch. 8 \u2013 Reporting & Dashboards", "Looker Studio, weekly reporting, monthly business reviews"],
          ["Ch. 9 \u2013 Cross-Channel Strategy", "How Google and Meta work together, CRM integration, Lookalikes"],
          ["Ch. 10 \u2013 Processes & Team Workflow", "Daily, weekly, monthly routines, sales team coordination"]
        ]
      ),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 1
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("1", "Strategy & Architecture", "Building the foundation before you spend a single rupee"),
      spacer(160),

      h2("1.1  The Four Campaign Verticals"),
      body("Every rupee of paid media budget is allocated across four distinct business objectives. Keeping these verticals separate is non-negotiable \u2014 it ensures clean attribution, objective-specific optimization, and faster decisions when something underperforms."),
      spacer(80),

      threeColTable(
        ["Vertical", "Objective", "Primary Platform"],
        [
          ["Buyer Leads", "Generate inquiries from people looking to purchase a property", "Google Search + Meta"],
          ["Rental Leads", "Generate tenant inquiries for listed rental properties", "Google Search + Meta"],
          ["Owner Listings \u2013 Sale", "Acquire homeowners willing to list their property for sale", "Meta"],
          ["Owner Listings \u2013 Rent", "Acquire homeowners willing to list their property for rent", "Meta"]
        ]
      ),
      spacer(80),

      calloutBox("Why Keep Verticals Separate?",
        "Mixing buyer and rental objectives inside the same campaign prevents the algorithm from optimizing correctly. Budget bleeds from high-priority verticals. Attribution becomes murky. You cannot read performance cleanly enough to make weekly decisions. One campaign = one goal.",
        LIGHT_GOLD, ACCENT_GOLD),

      spacer(160),
      h2("1.2  Account Architecture Principles"),
      body("A disciplined account structure is what separates a campaign that scales from one that becomes a mess after three months. Follow these rules from day one."),
      spacer(80),
      bullet("One campaign per vertical and per platform. Never mix buyer and rental inside the same campaign."),
      bullet("Each campaign has its own budget, bidding strategy, and primary conversion goal."),
      bullet("Ad groups within Google Search campaigns are structured by locality and property type \u2014 e.g., \"2BHK Whitefield Sale\", \"3BHK Koramangala Rent\"."),
      bullet("Meta ad sets are structured by audience type \u2014 cold interest, cold location, warm retargeting, Lookalike."),
      bullet("Retargeting campaigns always live in their own separate campaign, never inside prospecting campaigns."),
      bullet("Naming conventions are enforced consistently: [Platform]_[Vertical]_[Audience/AdGroup]_[Date]"),

      spacer(80),
      h2("1.3  Budget Allocation Framework"),
      body("Monthly paid media budget ranges from \u20B92.5L to \u20B93L. Allocation across platforms and verticals is reviewed monthly and adjusted based on performance data and seasonality."),
      spacer(80),

      threeColTable(
        ["Vertical", "Typical Budget Split", "Rationale"],
        [
          ["Buyer Leads \u2013 Google", "35\u201340%", "High-intent search traffic; highest lead quality despite higher CPL"],
          ["Rental Leads \u2013 Google", "15\u201320%", "Moderate volume; strong with location-radius targeting"],
          ["Owner Listing \u2013 Meta", "20\u201325%", "Lower CPL (\u20B9380\u2013420); critical for inventory pipeline"],
          ["Buyer/Rental \u2013 Meta", "15\u201320%", "Supports volume; video creative drives CTR"],
          ["Branded Search", "3\u20135%", "Captures warm Meta-exposed users; lowest CPL in account"]
        ]
      ),
      spacer(80),

      bullet("Shift budget toward verticals that beat CPL targets by >10% consistently for 3 weeks."),
      bullet("Reduce budget for verticals that exceed CPL target by >20% for more than 2 consecutive weeks."),
      bullet("Never cut a vertical to zero without sales team consultation \u2014 some verticals feed pipeline even at higher CPL."),
      bullet("Keep a 5\u201310% flex reserve for burst opportunities (new project launches, festive campaigns)."),

      spacer(160),
      h2("1.4  Bidding Strategy Selection"),
      body("Choose your bidding strategy based on where a campaign is in its lifecycle. Automated strategies require data; they perform poorly when conversion history is thin."),
      spacer(80),

      twoColTable(
        ["Campaign Stage", "Recommended Bidding Strategy"],
        [
          ["New campaign (0\u201350 conversions/month)", "Manual CPC \u2014 gives you control while building conversion data"],
          ["Growing campaign (50\u2013100 conversions/month)", "Maximize Conversions without a target \u2014 lets algorithm explore"],
          ["Mature campaign (100+ conversions/month)", "Target CPA \u2014 set based on your CPL target, allows algorithmic efficiency"],
          ["Brand awareness / Display", "Target CPM or Maximize Reach"],
          ["Meta \u2013 Lead Gen campaigns", "Lowest Cost (default) until you have 50+ leads, then Cost Cap"],
          ["Meta \u2013 Retargeting", "Lowest Cost always \u2014 warm audience does not need aggressive bidding"]
        ]
      ),
      spacer(80),

      calloutBox("Real Result: Target CPA vs Manual CPC",
        "A/B tested Manual CPC against Target CPA on the buyer leads campaign using Google's built-in Campaign Experiment feature. After the learning phase (approximately 2 weeks), Target CPA reduced CPL by 14%. The campaign had accumulated over 120 conversions before switching, which was the prerequisite for the strategy to work correctly.",
        LIGHT_TEAL, ACCENT_TEAL),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 2
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("2", "Google Ads", "Search, Performance Max, Display & Branded campaigns"),
      spacer(160),

      h2("2.1  Google Search Campaigns"),
      h3("Campaign Setup"),
      bullet("Campaign type: Search"),
      bullet("Goal: Leads (form submissions, click-to-call, WhatsApp clicks)"),
      bullet("Network: Search Network only. Disable Search Partners and Display Network expansion on new campaigns \u2014 re-evaluate after 30 days."),
      bullet("Location: Target Bengaluru at the campaign level. Refine to specific localities at the ad group level using location bid adjustments."),
      bullet("Language: English + Kannada"),
      bullet("Ad schedule: Run 24/7 initially; after 30 days, analyze hourly performance and add bid adjustments for peak hours (typically 6\u201310 PM on weekdays)."),

      spacer(80),
      h3("Ad Group Structure"),
      body("Structure ad groups by locality + property type. This maintains tight keyword-to-ad relevance which directly impacts Quality Score and CPL."),
      spacer(60),

      twoColTable(
        ["Ad Group Name", "Example Keywords"],
        [
          ["2BHK_Whitefield_Sale", "2bhk for sale in whitefield, 2 bedroom apartment whitefield buy, flat for sale whitefield bangalore"],
          ["3BHK_Koramangala_Sale", "3bhk flat koramangala sale, 3 bedroom apartment koramangala, buy apartment koramangala"],
          ["2BHK_HSR_Rent", "2bhk for rent in hsr layout, 2 bedroom flat rent hsr, flat to rent hsr bangalore"],
          ["1BHK_Marathahalli_Rent", "1bhk rent marathahalli, 1 bedroom flat marathahalli, apartments for rent marathahalli"]
        ]
      ),
      spacer(80),

      h3("Keyword Match Types"),
      bullet("Start with Phrase Match for all keywords. It gives enough control while allowing natural language variations."),
      bullet("Add Exact Match for your 5\u201310 highest-converting queries once you have 60+ days of Search Terms data."),
      bullet("Avoid Broad Match in the early months. If you use it later, pair it with Smart Bidding and a strong negative list."),
      bullet("Review the Search Terms Report every week. Convert high-performing search terms into new exact match keywords."),

      spacer(80),
      h3("Negative Keywords"),
      body("Negative keywords are one of the highest-ROI activities in Google Ads management. Run a weekly review without exception."),
      spacer(60),

      calloutBox("Week 1 Negatives: Add These Before You Launch",
        "jobs, career, fresher, internship, free, low cost, cheap, reviews, complaint, news, youtube, wikipedia, images, pictures, gallery, sqft rate, property tax, bank loan, home loan rate, vastu, 2bhk meaning, real estate meaning, how to invest, real estate news",
        LIGHT_GOLD, ACCENT_GOLD),

      spacer(80),
      bullet("Add 60+ negatives in Month 1. Download Search Terms weekly and scan for irrelevant queries."),
      bullet("Maintain a Master Negative Keyword List in a shared Google Sheet. Add to it weekly; apply to all campaigns."),
      bullet("Create a Negative Keyword List at the MCC or account level and link it to all campaigns to save time."),
      bullet("Common irrelevant categories in real estate: job-related, DIY/construction, financial products, news/media, competitor brand names (unless you have a competitor campaign strategy)."),

      spacer(80),
      h3("Ad Copy Guidelines"),
      bullet("Each ad group should have a minimum of 3 Responsive Search Ads (RSAs)."),
      bullet("Include the locality name in at least 2 headlines: \"2BHK Flats in Whitefield\", \"Whitefield Property \u2013 Act Now\""),
      bullet("Lead with the strongest value proposition in Headline 1: Price, availability, USP."),
      bullet("Use all 15 headline slots and all 4 description slots. Give Google's algorithm enough assets to optimize."),
      bullet("Pin Headline 1 if it contains your primary keyword for Quality Score relevance."),
      bullet("Use ad extensions: Sitelinks (4 minimum), Callouts (6 minimum), Structured Snippets (property types), Call extension, Location extension."),

      spacer(80),
      h3("Quality Score Optimization"),
      body("Quality Score (QS) directly affects your CPL. A higher QS means you pay less per click for the same position."),
      spacer(60),

      threeColTable(
        ["QS Component", "What Drives It", "Action to Take"],
        [
          ["Expected CTR", "Ad relevance to keyword; historical CTR", "Add keyword in Headline 1; test multiple ad variations"],
          ["Ad Relevance", "How closely ad matches keyword intent", "Restructure ad groups so one ad group = one theme"],
          ["Landing Page Experience", "Page relevance, load speed, mobile UX", "Match landing page headline to Search query; compress images; use AMP or fast-loading pages"]
        ]
      ),
      spacer(80),

      calloutBox("Real Result: QS Improvement",
        "Average Quality Score improved from 5 to 7 across key ad groups over 6 months by restructuring ad copy to match keyword intent and aligning landing page headlines with Search queries. This reduced CPL from \u20B9820 to \u20B9540 \u2014 a 34% reduction.",
        LIGHT_TEAL, ACCENT_TEAL),

      spacer(160),
      h2("2.2  Performance Max Campaigns"),
      body("Performance Max (PMax) is Google's AI-driven campaign type that serves across all Google channels \u2014 Search, Display, YouTube, Gmail, Maps, and Discover \u2014 from a single campaign."),
      spacer(80),

      h3("When to Use PMax in Real Estate"),
      bullet("Use PMax as a supplement to, not a replacement for, Search campaigns."),
      bullet("PMax works well for broad awareness and capturing mid-funnel intent signals."),
      bullet("Run PMax with branded exclusions to prevent it from cannibalizing your Branded Search campaign."),

      h3("Asset Group Setup"),
      bullet("Create separate Asset Groups per property type and locality, mirroring your Search ad group structure."),
      bullet("Upload a minimum of: 5 headlines, 5 long headlines, 5 descriptions, 5 images (landscape + portrait), 1 video (or Google will auto-generate one from your images)."),
      bullet("Use real property photography \u2014 not stock images. CTR and conversion rate are significantly higher."),
      bullet("Add Audience Signals to guide the algorithm: upload your CRM lead list, add in-market audiences for Real Estate, add remarketing lists."),

      h3("PMax Reporting Limitations"),
      bullet("PMax provides limited transparency on where clicks are serving. Use the Insights tab to monitor."),
      bullet("Check the Search Terms report inside PMax weekly \u2014 it now shows some search query data."),
      bullet("If Search campaigns are underperforming after PMax launch, check whether PMax is bidding on your core keywords. Use campaign-level brand exclusions."),

      spacer(160),
      h2("2.3  Display Campaigns"),
      body("Display campaigns serve visual banner ads across Google's Display Network. In real estate, they are primarily used for retargeting and brand awareness rather than direct lead generation."),
      spacer(80),

      bullet("Use Responsive Display Ads \u2014 provide 5 images, 5 headlines, 5 descriptions, and 1 logo."),
      bullet("Target audiences: Website visitors (retargeting), Custom Intent audiences (searched for property-related terms), In-Market: Real Estate."),
      bullet("Avoid broad demographic targeting on Display for direct lead generation \u2014 CPL will be high and lead quality low."),
      bullet("Use Display for retargeting users who visited your landing page but did not convert. This warms them up before they see your Search or Meta ads again."),
      bullet("Set frequency caps: no more than 3 impressions per user per day to avoid ad fatigue."),

      spacer(160),
      h2("2.4  Branded Search Campaign"),
      body("A branded campaign targets people who search for your agency name \u2014 e.g., \"South City Realtors\", \"South City properties Bengaluru\". This is often the most undervalued campaign in the account."),
      spacer(80),

      calloutBox("Why Branded Search Exists",
        "Many potential clients see your Meta ads but do not click immediately. Later they search your agency name on Google. Without a Branded campaign, a competitor could appear in that moment and steal the lead. Branded Search captures warm intent at the bottom of the funnel at the lowest CPL in the account (\u20B9220\u2013280).",
        LIGHT_BLUE, BRAND_BLUE),

      spacer(80),
      bullet("Use Exact and Phrase match for your agency name and common variations."),
      bullet("Set bids aggressively \u2014 you need position 1 for your own brand terms."),
      bullet("Keep separate from non-brand campaigns for clean budget and performance tracking."),
      bullet("Link branded CPL performance to your Meta spend to demonstrate cross-channel attribution to leadership."),
      bullet("Exclude branded terms from all non-branded campaigns using a Brand Exclusion List."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 3
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("3", "Meta Ads", "Facebook & Instagram: Lead Generation, Conversions & Creative"),
      spacer(160),

      h2("3.1  Campaign Types on Meta"),
      threeColTable(
        ["Objective", "When to Use", "Best Format"],
        [
          ["Lead Generation", "Primary objective for all verticals. Form fills directly inside Facebook/Instagram.", "Carousel, Single Image, Video"],
          ["Conversions", "When you have a high-traffic landing page and strong Pixel data (50+ conversions/week).", "Video, Single Image"],
          ["Traffic", "Top-of-funnel for building Pixel audiences before launching conversion campaigns.", "Single Image"],
          ["Reach / Brand Awareness", "Festive season or new project announcements. Not for direct lead gen.", "Video, Carousel"]
        ]
      ),
      spacer(80),

      h2("3.2  Ad Set Structure & Audience Targeting"),
      body("Each ad set contains one audience segment and one budget. Keep audiences separate so performance is readable."),
      spacer(80),

      h3("Location Targeting"),
      bullet("Target by specific pin codes rather than broad city-level targeting. Bengaluru has micro-markets with very different buyer profiles."),
      bullet("For Whitefield buyer campaigns, target pin codes: 560066, 560048, 560037 and adjacent codes."),
      bullet("Use a 2\u20135 km radius around target localities for rental campaigns to reach people already in or near that area."),
      bullet("Exclude locations with consistently poor lead quality (review monthly with sales team)."),

      spacer(80),
      h3("Demographic Targeting"),
      bullet("Age: 25\u201355 for buyer campaigns. 22\u201345 for rental campaigns. 30\u201360 for owner listing campaigns."),
      bullet("Income: Use the Top 10\u201325% income brackets for premium property campaigns. Widen for affordable housing."),
      bullet("Relationship status: Recently married is a strong life-event signal for both buyer and rental campaigns."),
      bullet("Life Events: Newly married, Recently moved, Away from hometown \u2014 highly relevant for rental."),

      spacer(80),
      h3("Interest-Based Targeting"),
      bullet("Real estate interests: Property Investment, Home Ownership, Real Estate Investing, MagicBricks, 99acres, Housing.com"),
      bullet("Lifestyle signals: Interior design, Home improvement, Architecture, Luxury goods (for premium properties)"),
      bullet("Financial signals: Personal finance, Mutual fund investing, Stock market (buyers are often investors)"),
      bullet("Do not stack too many interests \u2014 it over-narrows your audience. Keep broad-to-moderate targeting and let Meta's algorithm optimize."),

      spacer(80),
      h3("Custom Audiences & Exclusions"),
      bullet("Upload your CRM lead list monthly as a Custom Audience and exclude from all cold prospecting campaigns."),
      bullet("Create a Website Visitors Custom Audience (180 days) for retargeting."),
      bullet("Create a Lead Form Openers Custom Audience (people who opened but did not submit) for soft retargeting."),
      bullet("Exclude existing leads and clients from new lead gen campaigns to avoid wasting budget on people who have already converted."),

      spacer(80),
      h3("Lookalike Audiences"),
      bullet("Build 1% Lookalike from your CRM converted leads list \u2014 this consistently outperforms interest-based targeting by 15\u201320% in CPL."),
      bullet("Minimum source audience: 100 people. Optimal: 500\u20131,000 people for meaningful signal."),
      bullet("Test 1% vs 2\u20133% Lookalike \u2014 broader Lookalikes lower CPL but may reduce lead quality."),
      bullet("Refresh your source audience every 30 days by re-uploading your updated CRM list."),

      spacer(160),
      h2("3.3  Meta Lead Forms: Best Practices"),
      body("The Meta Instant Form is the primary conversion point for lead generation campaigns. The quality of your form directly determines the quality of your leads."),
      spacer(80),

      h3("Form Type Selection"),
      bullet("Use More Volume form for high-CPL campaigns where quantity is the priority."),
      bullet("Use Higher Intent form for owner listing campaigns \u2014 it adds a review step before submission, reducing accidental leads."),

      spacer(80),
      h3("Qualifying Questions (The Most Important Setting)"),
      body("Adding 2\u20133 qualifying questions is what separates a form that generates 50 poor leads from one that generates 30 qualified ones."),
      spacer(60),

      twoColTable(
        ["Vertical", "Qualifying Questions to Add"],
        [
          ["Buyer Leads", "What is your budget range? (dropdown: \u20B940L\u201360L / \u20B960L\u201381 Cr / \u20B91 Cr+) | What size are you looking for? (1BHK / 2BHK / 3BHK) | When are you planning to buy? (Within 3 months / 3\u20136 months / Just exploring)"],
          ["Rental Leads", "Which area do you prefer? (Whitefield / Koramangala / HSR / Other) | What is your monthly budget? (dropdown) | When do you need to move in?"],
          ["Owner Listing \u2013 Sale", "Do you own a property you want to sell? (Yes/No) | What is the property type? (Apartment / Villa / Plot) | Which locality is it in?"],
          ["Owner Listing \u2013 Rent", "Do you own a property you want to rent out? (Yes/No) | Is the property currently vacant? (Yes / No, tenant leaving soon) | Which area?"]
        ]
      ),
      spacer(80),

      calloutBox("Real Result: Qualifying Questions on Rental Forms",
        "Adding area preference and budget range as qualifying questions to Meta lead forms, combined with tightening location radius targeting, achieved a consistent lead-to-site-visit conversion rate of approximately 18% in the rental vertical.",
        LIGHT_TEAL, ACCENT_TEAL),

      spacer(160),
      h2("3.4  Creative Formats & Strategy"),
      h3("Format Performance by Vertical"),
      threeColTable(
        ["Format", "Best For", "Key Finding"],
        [
          ["Video Walkthrough", "Rental and buyer campaigns, cold audiences", "Outperformed static images by ~30% in CTR for rental vertical; 60% of Meta budget shifted to video"],
          ["Carousel", "Showcasing multiple properties or floor plans", "Strong for buyer campaigns with diverse inventory"],
          ["Single Image", "Retargeting and owner listing campaigns", "Simple and effective when copy is strong; lower production cost"],
          ["Story / Reel", "Younger audience; rental campaigns for 22\u201330 age group", "Short-form video; use vertical format (9:16)"]
        ]
      ),
      spacer(80),

      h3("Creative Brief Template"),
      body("Every creative request to the freelance designer should include the following elements. A clear brief prevents revision cycles and maintains quality."),
      spacer(60),
      bullet("Objective: What should the viewer feel or do?"),
      bullet("Format: Carousel / Single Image / Video. Dimensions specified."),
      bullet("Headline: Primary text to feature prominently in the creative."),
      bullet("Value Proposition: The one thing that makes this property or offer compelling."),
      bullet("Visual Reference: Property photos, mood board, or reference ad."),
      bullet("Call to Action: \"Enquire Now\", \"Get Free Valuation\", \"Schedule a Visit\""),
      bullet("Deadline: Always set 5\u20137 days from brief submission for production turnaround."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 4
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("4", "Owner Listing Campaigns", "Acquiring homeowners who want to sell or rent out their property"),
      spacer(160),

      h2("4.1  Why Owner Listings Are a Separate Vertical"),
      body("Owner listing campaigns target homeowners, not property seekers. The psychology, messaging, creative, and qualification criteria are completely different from buyer or renter campaigns. Running them inside the same campaign as buyer leads will destroy performance on both sides."),
      spacer(80),

      twoColTable(
        ["Dimension", "Owner Listing Campaign vs. Buyer/Rental Campaign"],
        [
          ["Target Person", "Homeowner who wants to sell or rent their existing property"],
          ["Core Motivation", "Get a good price, find a verified buyer/tenant quickly, avoid hassle"],
          ["Messaging Angle", "Free valuation, faster discovery, verified inquiries, zero upfront cost"],
          ["Platform", "Meta only (Facebook + Instagram); homeowners are not typically searching Google for this"],
          ["CPL Target", "\u20B9380\u2013420 \u2014 approximately 15% below initial team target"],
          ["Lead Volume", "60\u201380 leads per month across sale + rent listing verticals"],
          ["Form Type", "Higher Intent (adds a review step to reduce accidental submissions)"]
        ]
      ),

      spacer(160),
      h2("4.2  Campaign Structure"),
      bullet("Campaign 1: Owner Listing \u2013 Sale (targeting homeowners wanting to sell)"),
      bullet("Campaign 2: Owner Listing \u2013 Rent (targeting homeowners wanting to rent out)"),
      bullet("Keep budgets separate so each objective is independently measurable."),
      bullet("Ad Set structure: one ad set per audience type (interest-based, Lookalike, retargeting)."),

      spacer(80),
      h2("4.3  Audience Targeting for Owner Listings"),
      bullet("Age 35\u201365: Most property owners fall in this range."),
      bullet("Homeowner status: Facebook allows targeting \"likely homeowners\" under Behaviors > Residential Profiles."),
      bullet("Life Events: Recently renovated home, homeowner-adjacent interests (interior design, property investment, real estate news)."),
      bullet("Build a 1% Lookalike from previously converted owner listing leads \u2014 this becomes your most efficient audience after 3\u20136 months."),

      spacer(80),
      h2("4.4  Messaging Framework"),
      body("Owner listing messaging must address the homeowner's fears (will my property sit unsold? will I get low-quality tenants?) and desires (quick sale, good price, verified buyers)."),
      spacer(60),

      twoColTable(
        ["Message Angle", "Example Ad Copy Direction"],
        [
          ["Speed", "\"Find a verified buyer for your property in 30 days. No brokerage. No hassle.\""],
          ["Valuation", "\"Get a FREE market valuation for your property. Know what your home is worth today.\""],
          ["Quality Buyers/Tenants", "\"We screen every buyer. You only meet serious, pre-qualified prospects.\""],
          ["Trust", "\"500+ successful transactions in Bengaluru. Your property is in safe hands.\""],
          ["Simplicity", "\"List your property in 5 minutes. We handle everything else.\""]
        ]
      ),

      spacer(160),
      h2("4.5  Lead Pre-Qualification"),
      body("Owner listing leads are pre-qualified via form questions before being passed to the sales team. This protects sales team time and increases conversion rate from lead to listing."),
      spacer(60),
      bullet("Question 1: Do you own a property you want to list? (Yes / No)"),
      bullet("Question 2: What type of property? (Apartment / Villa / Independent House / Plot)"),
      bullet("Question 3: Which locality is your property in? (Dropdown of key localities)"),
      bullet("Question 4: What is your listing intent? (Want to Sell / Want to Rent Out / Still Deciding)"),
      bullet("Only leads answering \"Yes\" to ownership and with a clear locality are passed to sales. Ambiguous responses go into a nurture queue."),

      spacer(160),
      h2("4.6  Form Abandonment Retargeting"),
      body("A significant portion of potential leads open the form but do not submit. This retargeting layer recovers a portion of that lost interest."),
      spacer(60),

      bullet("Create a Custom Audience: Lead Form \u2013 Opened but Not Submitted (last 14 days)"),
      bullet("Run a separate Meta campaign targeting this audience with a softer follow-up message (e.g., \"Still thinking about listing your property? Here's what we offer.\")"),
      bullet("Use single-image or text-heavy creative \u2014 keep it simple and low pressure."),
      bullet("Set frequency cap at 2\u20133 per day for this retargeting audience to avoid being intrusive."),
      bullet("This layer reduced form abandonment impact by an estimated 20% in Q4 2024."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 5
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("5", "Retargeting Strategy", "Converting warm audiences at 40\u201350% lower CPL"),
      spacer(160),

      h2("5.1  Why Retargeting Is Non-Negotiable"),
      body("Retargeting audiences \u2014 people who have already visited your website or interacted with your ads \u2014 consistently deliver CPL 40\u201350% lower than cold-audience campaigns. They are the most efficient spend in your account. Never sacrifice retargeting budget for prospecting budget."),
      spacer(80),

      h2("5.2  Retargeting Audience Segments"),
      threeColTable(
        ["Audience", "Platform", "Recommended Window"],
        [
          ["Website visitors (all pages)", "Google Display + Meta", "30 days"],
          ["Landing page visitors (did not convert)", "Google Display + Meta", "30 days"],
          ["Meta lead form openers (not submitted)", "Meta only", "14 days"],
          ["Video viewers (watched 50%+)", "Meta only", "30 days"],
          ["Instagram profile visitors", "Meta only", "60 days"],
          ["CRM leads (not yet visited site)", "Meta via Custom Audience upload", "Evergreen"]
        ]
      ),
      spacer(80),

      h2("5.3  Google Retargeting (Display)"),
      bullet("Link Google Ads to Google Analytics 4 to import Audiences."),
      bullet("Create audience segments in GA4: All Users > 30 days, Landed on property page > No conversion."),
      bullet("Run Responsive Display Ads with property images and a clear CTA (\"Schedule a Visit\", \"Call Now\")."),
      bullet("Bid strategy: Target CPA or Target ROAS based on volume. Max CPM if volume is low."),
      bullet("Frequency cap: 3 impressions/user/day on Display retargeting."),

      spacer(80),
      h2("5.4  Meta Retargeting"),
      bullet("Set budget: 15\u201320% of total Meta budget allocated to retargeting campaigns."),
      bullet("Creative: Change the angle from the cold ad. If cold ad was aspirational (\"Dream home in Whitefield\"), retargeting should be practical (\"Viewed properties in Whitefield? Schedule a visit today.\")"),
      bullet("Exclude all retargeting audiences from cold prospecting campaigns to prevent overlap and budget waste."),
      bullet("Use dynamic property ads (DPA) if your inventory is on a product catalog. DPA automatically shows the exact property a user viewed on your website."),

      spacer(80),
      calloutBox("Retargeting CPL Benchmark",
        "The Meta retargeting campaign consistently delivers CPL 40\u201350% lower than cold-audience campaigns. This is because warm audiences already recognize the brand and have demonstrated intent by visiting the website. Protect this budget line even when overall budgets are under pressure.",
        LIGHT_GOLD, ACCENT_GOLD),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 6
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("6", "Tracking & Attribution", "Google Tag Manager, Meta Pixel, UTM Parameters & GA4"),
      spacer(160),

      h2("6.1  Conversion Tracking Architecture"),
      body("Accurate conversion tracking is the single most important technical setup in your account. Without it, your bidding algorithms optimize on noise, your reports are wrong, and your decisions are based on guesswork."),
      spacer(80),

      twoColTable(
        ["Conversion Event", "Platform", "Priority"],
        [
          ["Form Submission", "Google Ads + Meta Pixel + GA4", "Primary \u2014 the main conversion goal"],
          ["Click-to-Call", "Google Ads + GA4", "Secondary"],
          ["WhatsApp Button Click", "Google Ads + Meta Pixel + GA4", "Secondary"],
          ["Page View (Landing Page)", "Meta Pixel (ViewContent)", "Informational"],
          ["Lead (Meta Pixel)", "Meta Pixel", "Primary for Meta campaigns"],
          ["Video View (50%, 75%)", "Meta Pixel", "Retargeting audience builder"]
        ]
      ),

      spacer(160),
      h2("6.2  Google Tag Manager Setup"),
      h3("Container Setup"),
      bullet("Create one GTM container per website. Install the container snippet in <head> and <body> of every page."),
      bullet("Publish a debug-mode test with Google Tag Assistant before going live."),

      spacer(60),
      h3("Tags to Configure"),
      numbered("Google Ads Conversion Tracking tag \u2014 fires on form submission thank-you page or on successful form submit event."),
      numbered("GA4 Configuration tag \u2014 fires on all pages. Use your GA4 Measurement ID."),
      numbered("GA4 Event tag for Form Submission \u2014 fire on the same trigger as your Google Ads conversion."),
      numbered("GA4 Event tag for Click-to-Call \u2014 fire on click of phone number link."),
      numbered("GA4 Event tag for WhatsApp button click \u2014 fire on click of WhatsApp link."),
      numbered("Meta Pixel Base Code tag \u2014 fires on all pages."),
      numbered("Meta Pixel Lead event tag \u2014 fires on form submission."),
      numbered("Meta Pixel ViewContent event tag \u2014 fires on property listing page view."),

      spacer(80),
      h3("Trigger Configuration"),
      bullet("Thank-you page trigger: Page URL contains \"/thank-you\" or \"/success\". Most reliable method."),
      bullet("Form submit trigger: Use GTM's built-in Form Submission trigger if no redirect to thank-you page. Enable \"Wait for Tags\" and \"Check Validation\"."),
      bullet("Click trigger for call/WhatsApp: Use Click URL contains \"tel:\" and \"wa.me\" respectively."),

      spacer(160),
      h2("6.3  Meta Pixel Configuration"),
      bullet("Install via GTM (recommended) rather than hardcoding in HTML."),
      bullet("Verify installation with Meta Pixel Helper Chrome extension and Meta Events Manager Test Events tool."),
      bullet("Events to configure: PageView (all pages), ViewContent (property listing pages), Lead (form submission), Contact (WhatsApp/call click)."),
      bullet("Enable Advanced Matching in Meta Events Manager: pass hashed email and phone from form submissions to improve attribution accuracy."),
      bullet("Set up Conversions API (CAPI) as a server-side complement to Pixel to recover iOS 14+ data loss. This is critical \u2014 browser-based Pixel alone underreports by 20\u201330%."),

      spacer(80),
      calloutBox("Conversions API (CAPI) \u2014 Why It Matters",
        "Apple's iOS 14 App Tracking Transparency update blocks Meta Pixel from firing in approximately 30% of mobile Safari sessions. Conversions API sends the same events server-side, bypassing the browser block. Setting up CAPI typically recovers 15\u201325% of lost conversions and significantly improves campaign optimization signal.",
        LIGHT_BLUE, BRAND_BLUE),

      spacer(160),
      h2("6.4  UTM Parameter Framework"),
      body("Apply UTM parameters to every ad URL on every platform without exception. This gives you platform-independent data in GA4 to cross-reference against platform-reported numbers."),
      spacer(80),

      twoColTable(
        ["Parameter", "Google Ads Value (example)"],
        [
          ["utm_source", "google"],
          ["utm_medium", "cpc"],
          ["utm_campaign", "buyer_leads_whitefield_search"],
          ["utm_content", "2bhk_sale_ad_v2"],
          ["utm_term", "{keyword} (dynamic insertion)"]
        ]
      ),

      spacer(60),
      twoColTable(
        ["Parameter", "Meta Ads Value (example)"],
        [
          ["utm_source", "facebook"],
          ["utm_medium", "paid_social"],
          ["utm_campaign", "rental_leads_koramangala_lg"],
          ["utm_content", "video_walkthrough_2bhk_v1"],
          ["utm_term", "life_event_married (audience label)"]
        ]
      ),
      spacer(80),
      bullet("Maintain a UTM builder spreadsheet. Every new ad creative or landing page gets its UTM built and recorded before launch."),
      bullet("Use GA4 as an independent attribution source. Compare platform-reported leads with GA4-attributed sessions monthly to identify discrepancy."),
      bullet("A discrepancy of 20\u201330% between Google Ads reported conversions and GA4 is normal (attribution window differences). A discrepancy >40% indicates a tracking issue that needs investigation."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 7
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("7", "Creative Strategy & Testing", "A/B testing, fatigue management & the creative production cycle"),
      spacer(160),

      h2("7.1  The A/B Testing Framework"),
      body("Every creative decision must be backed by a test. Assumptions are expensive in paid media. The discipline is: one variable at a time, documented before and after."),
      spacer(80),

      h3("Test Design Rules"),
      bullet("Change only ONE variable per test. Testing two things simultaneously makes it impossible to identify what caused the result."),
      bullet("Run 3\u20134 Meta creative tests per month. Any more dilutes budget and statistical significance."),
      bullet("Minimum test duration: 7 days. Minimum spend: \u20B95,000 per variant. Without this threshold, results are noise."),
      bullet("Define success criteria before launching the test. \"The winner is the ad with lower CPL after 7 days and \u20B95,000 spend per variant.\""),

      spacer(80),
      h3("Test Documentation Template"),
      twoColTable(
        ["Field", "What to Fill"],
        [
          ["Test ID", "Sequential number, e.g., META-TEST-012"],
          ["Date Launched", "DD/MM/YYYY"],
          ["Campaign / Ad Set", "Exact name of where the test ran"],
          ["Variable Tested", "e.g., Image vs Video, Short copy vs Long copy"],
          ["Hypothesis", "\"We believe video will outperform image because property seekers want to experience the space visually.\""],
          ["Control (A)", "Description of the existing/baseline creative"],
          ["Variant (B)", "Description of the new creative"],
          ["Result", "CPL, CTR, and lead volume for A and B after test period"],
          ["Decision", "Which variant won and what action was taken (scale / pause / iterate)"],
          ["Learning", "One-line takeaway for future reference"]
        ]
      ),
      spacer(80),

      calloutBox("Real Test Result: Video vs Static Image",
        "A/B test on Meta for the rental vertical: Video walkthroughs vs. static property images. Video outperformed static by approximately 30% in CTR. Decision: Shifted 60% of Meta creative budget to video formats. This test is now a foundational assumption for the rental vertical.",
        LIGHT_TEAL, ACCENT_TEAL),

      spacer(80),
      h3("What to Test (Priority Order)"),
      numbered("Format: Image vs. Video vs. Carousel \u2014 highest impact variable"),
      numbered("Hook / Headline: Different angles (price-led, FOMO, aspirational, social proof)"),
      numbered("Copy length: Short (2 lines) vs. long (5+ lines)"),
      numbered("CTA button: \"Get Free Quote\" vs. \"Schedule a Visit\" vs. \"Learn More\""),
      numbered("Audience: Interest-based vs. Lookalike vs. Life-event targeting"),
      numbered("Offer: Free valuation vs. Limited units vs. Zero brokerage"),

      spacer(160),
      h2("7.2  Creative Fatigue Management"),
      body("Ad fatigue is the single biggest silent killer of Meta campaign performance. When users see the same ad too many times, CTR drops, CPM rises, and CPL climbs \u2014 all without a single change to your targeting or bids."),
      spacer(80),

      h3("Fatigue Warning Signals"),
      bullet("CTR drops >25% week-over-week on the same creative"),
      bullet("Frequency exceeds 3.0 for cold audiences within a week"),
      bullet("CPL increases >20% week-over-week without targeting change"),
      bullet("Negative comments increase on ad (\"I keep seeing this\")"),

      spacer(80),
      h3("Creative Refresh Schedule"),
      twoColTable(
        ["Audience Type", "Creative Refresh Frequency"],
        [
          ["Cold Audiences (prospecting)", "Every 3\u20134 weeks"],
          ["Retargeting Audiences", "Every 10\u201314 days (sees more frequency)"],
          ["Owner Listing Audiences", "Every 3\u20134 weeks (smaller audience = faster fatigue)"],
          ["Branded / Warm Audiences", "Monthly refresh is sufficient"]
        ]
      ),
      spacer(80),

      h3("The Production Pipeline"),
      body("The creative production cycle runs continuously. New creatives must always be ready before the current ones fatigue."),
      spacer(60),
      numbered("Week 1: Brief the designer on next set of creatives (3\u20134 new ads)"),
      numbered("Week 1\u20132: Designer produces assets (5\u20137 day turnaround)"),
      numbered("Week 2: Review and approve creatives; upload to Meta as inactive ads"),
      numbered("Week 3\u20134: Launch new creatives as current ones approach fatigue threshold"),
      numbered("Week 4: Brief the next round. The cycle never stops."),

      spacer(160),
      h2("7.3  Google Ads A/B Testing"),
      h3("Campaign Experiments"),
      bullet("Use Google Ads' built-in Campaign Experiment feature for structured A/B tests on bidding strategy, ad copy, and landing pages."),
      bullet("The experiment splits traffic between Control and Experiment, ensuring a fair comparison."),
      bullet("Tested: Manual CPC vs Target CPA on buyer leads campaign. Result: Target CPA reduced CPL by 14% after learning phase."),

      spacer(80),
      h3("Ad Copy Testing in RSAs"),
      bullet("Google's RSA machine learning tests headline and description combinations automatically. Check the \"Asset Performance\" report monthly to see which combinations perform best."),
      bullet("Pin strong performers: If Headline 1 with the locality name consistently outperforms, pin it to ensure it always shows."),
      bullet("Remove Low-performing assets regularly. A \"Low\" rating asset should be replaced after 30 days of data."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 8
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("8", "Reporting & Dashboards", "Weekly dashboards, monthly reviews & communicating with leadership"),
      spacer(160),

      h2("8.1  Metrics That Matter"),
      body("Not all metrics are worth tracking. Focus on the metrics that directly reflect business outcomes."),
      spacer(80),

      threeColTable(
        ["Metric", "Definition", "Target / Benchmark"],
        [
          ["Cost Per Lead (CPL)", "Total spend \u00F7 Total leads", "Buyer: <\u20B9600 | Rental: <\u20B9500 | Owner Listing: <\u20B9430"],
          ["Lead Volume", "Total leads submitted per vertical per week", "300\u2013350/month total; 280 team minimum"],
          ["Lead-to-Site-Visit Rate", "Leads that result in a physical site visit", "Target \u226518% (rental); \u226512% (buyer)"],
          ["CTR (Click-Through Rate)", "Clicks \u00F7 Impressions", "Google Search: >5% | Meta: >1.5%"],
          ["Quality Score (Google)", "Google's 1\u201310 rating for ad group relevance", "Target: 7+ on all key ad groups"],
          ["Frequency (Meta)", "Average times each person has seen your ad", "Cold: <2.5/week | Retargeting: <5/week"],
          ["Budget Utilization", "Actual spend vs. planned budget", "95\u2013100% (under-delivery = missed opportunity)"],
          ["Cost Per Deal Estimate", "Total spend \u00F7 Deals closed (from CRM data)", "Reviewed monthly; benchmarked by vertical"]
        ]
      ),
      spacer(80),

      h2("8.2  Looker Studio Dashboard"),
      body("The weekly performance dashboard is built and maintained in Google Looker Studio. It is the single source of truth for campaign performance, shared with agency leadership every Monday morning."),
      spacer(80),

      h3("Dashboard Structure"),
      bullet("Page 1 \u2013 Executive Summary: Total spend, total leads, blended CPL, lead volume vs target, week-over-week trend."),
      bullet("Page 2 \u2013 Google Ads: Impressions, CTR, CPL, QS trend, Search Impression Share, budget utilization per campaign."),
      bullet("Page 3 \u2013 Meta Ads: Reach, frequency, CTR, CPL, lead volume by ad set, creative performance table."),
      bullet("Page 4 \u2013 Vertical Breakdown: Leads, CPL, and budget by vertical (buyer / rental / owner listing sale / owner listing rent)."),
      bullet("Page 5 \u2013 Lead Quality Bridge: Lead volume from ads vs. qualified leads confirmed by sales (requires weekly CRM data input)."),

      spacer(80),
      h3("Data Sources to Connect"),
      bullet("Google Ads \u2014 native Looker Studio connector"),
      bullet("Meta Ads \u2014 use a third-party connector (Supermetrics, Porter Metrics, or Data Slayer)"),
      bullet("GA4 \u2014 native Looker Studio connector"),
      bullet("Google Sheets \u2014 for manual CRM data entry (lead quality feedback from sales team)"),

      spacer(160),
      h2("8.3  Weekly Reporting Routine"),
      body("The weekly report is not a vanity exercise \u2014 it is the primary decision-making document. Every number should trigger a question: Is this better or worse than last week? Why? What will we do about it?"),
      spacer(80),

      numbered("Pull all data from Google Ads, Meta Ads, and GA4 into Looker Studio. Dashboard auto-refreshes."),
      numbered("Review Google Ads Search Terms Report. Add negatives. Note new converting search terms."),
      numbered("Check Meta creative fatigue signals: CTR trend, frequency, CPL week-over-week."),
      numbered("Compare platform-reported leads with GA4 attributed sessions. Flag discrepancies >30%."),
      numbered("Update the shared tracking Google Sheet with weekly actuals (spend, leads, CPL per vertical)."),
      numbered("Prepare a 5-bullet weekly summary for leadership: highlight one win, one concern, one action taken."),
      numbered("Send dashboard link + summary by Monday 10 AM."),

      spacer(160),
      h2("8.4  Monthly Business Review"),
      body("The monthly business review is a structured presentation to agency leadership. It covers the full month's performance, insights, and the plan for the next month."),
      spacer(80),

      h3("MBR Slide Structure"),
      numbered("Month Summary: Total spend, total leads, blended CPL, leads vs target (with percentage achievement)."),
      numbered("Vertical Performance: Table showing leads, CPL, and budget for each of the 4 verticals."),
      numbered("Lead Quality Analysis: Qualified leads %, leads-to-site-visit rate, cost-per-deal estimate (if CRM data available)."),
      numbered("Key Wins: 2\u20133 specific achievements with numbers. E.g., \"Video CPL \u20B965 lower than static this month.\""),
      numbered("Key Learnings: What did not work and why. Intellectual honesty builds trust with leadership."),
      numbered("Next Month Plan: Budget allocation, campaigns to launch or pause, creative themes, tests planned."),

      spacer(80),
      calloutBox("Presenting to Leadership: Plain Language Rule",
        "Present all findings in plain language. Replace \"ROAS\" with \"return per rupee spent\". Replace \"CTR\" with \"percentage of people who clicked\". Replace \"frequency\" with \"how many times each person saw our ad\". Leadership makes better decisions when they understand the numbers without needing a paid media dictionary.",
        LIGHT_GOLD, ACCENT_GOLD),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 9
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("9", "Cross-Channel Strategy", "How Google and Meta work together to reduce blended CPL"),
      spacer(160),

      h2("9.1  The Full-Funnel Picture"),
      body("Google and Meta are not independent channels \u2014 they are two parts of a continuous consumer journey. A prospect may see your property video on Instagram, search your brand name on Google three days later, and fill a lead form the following morning. Only by running both platforms \u2014 and connecting their data \u2014 can you see and optimize the full funnel."),
      spacer(80),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "TOP OF FUNNEL", bold: true, font: "Arial", size: 22, color: WHITE })], alignment: AlignmentType.CENTER }), new Paragraph({ children: [new TextRun({ text: "Awareness", font: "Arial", size: 20, color: "AACCEE" })], alignment: AlignmentType.CENTER })] }),
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: ACCENT_TEAL, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "MID FUNNEL", bold: true, font: "Arial", size: 22, color: WHITE })], alignment: AlignmentType.CENTER }), new Paragraph({ children: [new TextRun({ text: "Consideration", font: "Arial", size: 20, color: "AADDDD" })], alignment: AlignmentType.CENTER })] }),
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: ACCENT_GOLD, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "BOTTOM OF FUNNEL", bold: true, font: "Arial", size: 22, color: WHITE })], alignment: AlignmentType.CENTER }), new Paragraph({ children: [new TextRun({ text: "Conversion", font: "Arial", size: 20, color: "FFF0AA" })], alignment: AlignmentType.CENTER })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Meta video/carousel ads reach people in Bengaluru who match buyer/renter profiles but are not yet searching actively.", font: "Arial", size: 20, color: TEXT_DARK })] })] }),
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: LIGHT_TEAL, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Meta retargeting + Google Display follow up. Prospects research properties, visit your landing page.", font: "Arial", size: 20, color: TEXT_DARK })] })] }),
              new TableCell({ borders: allBorders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Google Search + Branded Search capture the high-intent search. Meta Lead form closes the loop.", font: "Arial", size: 20, color: TEXT_DARK })] })] })
            ]
          })
        ]
      }),

      spacer(160),
      h2("9.2  The Meta-to-Branded-Search Pipeline"),
      body("The Branded Search campaign (\u20B9220\u2013280 CPL) exists because of upstream Meta investment. Users who saw a Meta awareness ad and developed interest later searched the agency name on Google. This is the most concrete cross-channel attribution proof in the account."),
      spacer(80),
      bullet("Track Branded Search impression volume monthly. A spike in branded impressions after a high-spend Meta month confirms the relationship."),
      bullet("Do not cut Meta budget to \"save money\" without understanding that Branded Search performance will decline with it."),
      bullet("When presenting to leadership, show the combined blended CPL: (Meta spend + Google Branded spend) \u00F7 (Meta leads + Branded leads). This is always significantly lower than Meta CPL alone."),

      spacer(160),
      h2("9.3  CRM Integration & Audience Management"),
      h3("CRM to Meta Custom Audiences"),
      bullet("Export the full CRM lead list monthly. Upload to Meta as a Customer List Custom Audience."),
      bullet("Match fields: Email, Phone Number, First Name, Last Name, City. More fields = higher match rate."),
      bullet("Use this audience for TWO purposes: (1) Exclusion from cold campaigns, (2) Source for Lookalike generation."),

      spacer(80),
      h3("Lookalike Audience Strategy"),
      bullet("Create 1% Lookalike from converted leads (people who became actual clients, not just form fills if CRM tracks this)."),
      bullet("1% Lookalike CPL is consistently 15\u201320% lower than interest-based targeting."),
      bullet("Test 1% vs 2\u20133% Lookalike quarterly. The right percentage depends on audience size and budget."),
      bullet("If your converted leads list is <100 people: use form-fill leads as the source. If >500: use converted clients."),

      spacer(80),
      h3("Sales Team Feedback Loop"),
      body("The paid media function does not operate in a vacuum. The sales team is your most valuable source of lead quality data."),
      spacer(60),
      bullet("Weekly sync with the sales team: Review leads from the past week by source, locality, and property type."),
      bullet("Ask specifically: Which sources are sending ready-to-visit leads? Which localities are generating the most callbacks? Which audience segments are converting to site visits?"),
      bullet("Feed this data back: Pause underperforming localities in campaign targeting. Increase bids on localities with high site-visit conversion. Adjust qualifying questions if a particular segment consistently fails to convert."),
      bullet("Create a shared Google Sheet: Sales team logs \"Lead Quality\" (Hot / Warm / Cold / Junk) against each lead source. Review this weekly to calibrate campaigns."),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // CHAPTER 10
      // ═══════════════════════════════════════════════════════════════════
      chapterCover("10", "Processes & Team Workflow", "Daily, weekly & monthly routines that keep campaigns performing"),
      spacer(160),

      h2("10.1  Daily Routine (20\u201330 minutes)"),
      numbered("Check Google Ads and Meta Ads dashboards for anomalies: major CPL spikes, budget exhaustion before midday, zero-lead days."),
      numbered("Check for disapproved ads or policy violations (Google and Meta both send notifications, but check directly)."),
      numbered("Review overnight conversion data. If a campaign has zero conversions for 48 hours, investigate immediately."),
      numbered("Check campaign budget pacing: Are campaigns on track to spend their budget evenly? Over-delivery and under-delivery both need correction."),

      spacer(80),
      h2("10.2  Weekly Routine (3\u20134 hours)"),
      numbered("Google Ads \u2013 Download Search Terms Report. Review irrelevant queries. Add negatives to Master Negative Keyword List. Apply list to all campaigns."),
      numbered("Google Ads \u2013 Review Quality Scores by ad group. Flag any ad group scoring <5 for immediate copy and landing page review."),
      numbered("Google Ads \u2013 Check impression share. If Impression Share lost to budget >15%, consider budget increase request for that vertical."),
      numbered("Meta Ads \u2013 Review creative performance: CTR, CPL, Frequency by ad creative. Flag creatives approaching fatigue."),
      numbered("Meta Ads \u2013 Review audience performance by ad set. Pause underperforming audiences after 7 days and \u20B95,000+ spend with no leads."),
      numbered("Update Looker Studio dashboard. Verify data sources are pulling correctly."),
      numbered("Sales team sync: collect lead quality feedback. Update campaign targeting if needed."),
      numbered("Log any tests launched or completed in the Test Documentation Sheet."),
      numbered("Send weekly performance summary to leadership (5 bullets maximum)."),

      spacer(80),
      h2("10.3  Monthly Routine (6\u20138 hours)"),
      numbered("Full budget reconciliation: Actual spend vs. planned spend per vertical per platform. Note variances and reasons."),
      numbered("Lookalike audience refresh: Re-upload CRM list to Meta. Rebuild Lookalike audiences from the updated source."),
      numbered("Audience exclusion refresh: Re-upload current CRM leads list as exclusion Custom Audience."),
      numbered("Landing page performance review in GA4: Bounce rate, time on page, form completion rate. Flag pages below 30% form completion rate for redesign."),
      numbered("Keyword performance review in Google Ads: Pause keywords with zero conversions after 30 days and >200 clicks."),
      numbered("Meta creative audit: Archive all paused, low-performing creatives. Ensure the active creative library is clean and current."),
      numbered("Monthly Business Review preparation: compile the MBR deck and present to agency leadership."),
      numbered("Next-month planning: set budgets, plan tests, brief designer on new creatives, confirm team targets."),

      spacer(160),
      h2("10.4  Team & Vendor Coordination"),
      h3("Freelance Designer"),
      bullet("Brief: 3\u20134 new creatives every month per vertical. Provide written brief with all specifications (format, dimensions, copy, visual reference)."),
      bullet("Turnaround: 5\u20137 days from brief to delivery. Build this into the creative calendar."),
      bullet("Review: One round of revisions is standard. Be specific in revision notes \u2014 vague feedback wastes time."),
      bullet("File formats: Request source files (PSD/AI) plus final exports (JPG/PNG/MP4) for every creative. Store in a shared Google Drive folder organized by month and vertical."),

      spacer(80),
      h3("Sales Team"),
      bullet("Weekly sync meeting: 30 minutes. Agenda: lead quality review by source, locality performance, pipeline update."),
      bullet("Shared tracking sheet: Sales logs lead quality; marketing logs campaign data. One shared document, two teams."),
      bullet("Escalation protocol: If sales team flags a specific source as consistently delivering junk leads, pause that source within 24 hours. Do not wait for the weekly review."),

      spacer(160),
      h2("10.5  Seasonal Campaign Planning"),
      body("Real estate in Bengaluru has predictable demand patterns. Plan campaigns around these seasonal peaks to maximize budget efficiency."),
      spacer(80),

      threeColTable(
        ["Period", "Demand Signal", "Campaign Adjustment"],
        [
          ["Jan \u2013 Mar", "Q4 corporate relocations; new project launches post-New Year", "Increase buyer lead budget 10\u201315%; launch new project-specific campaigns"],
          ["Apr \u2013 Jun", "Summer slowdown; school year end relocation", "Shift budget toward rental campaigns; test owner listing campaigns harder"],
          ["Jul \u2013 Sep", "Monsoon \u2013 moderate activity; festive preview", "Maintain base budgets; run awareness campaigns building up to festive season"],
          ["Oct \u2013 Dec", "Diwali and festive season \u2013 highest demand period", "Increase total budget 20\u201330%; launch time-limited offers; increase creative refresh rate"]
        ]
      ),

      spacer(200),
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════
      // APPENDIX
      // ═══════════════════════════════════════════════════════════════════
      h1("Appendix: Quick-Reference Benchmarks & Checklists"),
      spacer(80),

      h2("A.1  Campaign Launch Checklist"),
      body("Before any new campaign goes live, verify every item on this checklist."),
      spacer(60),

      twoColTable(
        ["Item", "Platform"],
        [
          ["\u2610  Campaign objective matches business goal", "Google + Meta"],
          ["\u2610  Separate campaign per vertical per platform", "Google + Meta"],
          ["\u2610  Conversion tracking verified in Tag Assistant / Meta Test Events", "Google + Meta"],
          ["\u2610  UTM parameters applied to all ad URLs", "Google + Meta"],
          ["\u2610  Negative keyword list applied (Google) | CRM exclusion audience applied (Meta)", "Google + Meta"],
          ["\u2610  Ad copy reviewed: no trademarks, no superlatives (\"Best\", \"#1\") without proof", "Google + Meta"],
          ["\u2610  Landing page mobile load time <3 seconds", "All"],
          ["\u2610  Lead form includes qualifying questions", "Meta"],
          ["\u2610  Retargeting campaign set up alongside prospecting campaign", "Google + Meta"],
          ["\u2610  Looker Studio dashboard updated with new campaign", "All"],
          ["\u2610  Sales team briefed on new lead source and what to expect", "All"]
        ]
      ),

      spacer(160),
      h2("A.2  CPL Benchmarks by Vertical"),
      threeColTable(
        ["Vertical", "Target CPL", "Alert Threshold (Pause/Review)"],
        [
          ["Buyer Leads \u2013 Google Search", "\u20B9540\u2013600", "> \u20B9750 for 2 consecutive weeks"],
          ["Rental Leads \u2013 Google Search", "\u20B9450\u2013520", "> \u20B9650 for 2 consecutive weeks"],
          ["Buyer Leads \u2013 Meta", "\u20B9480\u2013580", "> \u20B9700 for 2 consecutive weeks"],
          ["Rental Leads \u2013 Meta", "\u20B9400\u2013500", "> \u20B9620 for 2 consecutive weeks"],
          ["Owner Listing \u2013 Sale (Meta)", "\u20B9380\u2013420", "> \u20B9550 for 2 consecutive weeks"],
          ["Owner Listing \u2013 Rent (Meta)", "\u20B9380\u2013420", "> \u20B9550 for 2 consecutive weeks"],
          ["Branded Search \u2013 Google", "\u20B9220\u2013280", "> \u20B9400 (indicates competitor bidding on brand)"],
          ["Retargeting (Meta)", "40\u201350% below respective cold CPL", "> Cold campaign CPL (retargeting should always be cheaper)"]
        ]
      ),

      spacer(160),
      h2("A.3  Glossary of Terms"),
      twoColTable(
        ["Term", "Definition"],
        [
          ["CPL", "Cost Per Lead \u2014 total spend divided by total leads generated"],
          ["CTR", "Click-Through Rate \u2014 percentage of people who clicked your ad after seeing it"],
          ["Quality Score", "Google's 1\u201310 rating for each keyword based on expected CTR, ad relevance, and landing page experience"],
          ["ROAS", "Return on Ad Spend \u2014 revenue generated per rupee spent on ads"],
          ["CPA", "Cost Per Acquisition \u2014 cost to acquire one customer/conversion"],
          ["Impression Share", "Percentage of total available impressions your ads actually received"],
          ["Frequency", "Average number of times one person has seen your ad within a given period"],
          ["Lookalike Audience", "Meta audience built to mirror the characteristics of your existing customers or leads"],
          ["Custom Audience", "Audience built from your own data \u2014 CRM list, website visitors, video viewers"],
          ["CAPI", "Conversions API \u2014 server-side event tracking that supplements the Meta Pixel"],
          ["GTM", "Google Tag Manager \u2014 tag management system for deploying tracking codes without editing website code"],
          ["GA4", "Google Analytics 4 \u2014 Google's current analytics platform"],
          ["PMax", "Performance Max \u2014 Google's AI-driven campaign type that serves across all Google channels"],
          ["RSA", "Responsive Search Ad \u2014 Google's ad format that tests multiple headline and description combinations"],
          ["UTM", "Urchin Tracking Module \u2014 URL parameters that identify the source, medium, and campaign of a click"],
          ["DPA", "Dynamic Product Ads \u2014 Meta ad format that automatically shows relevant inventory to each user based on website behavior"]
        ]
      ),

      spacer(200),

      // Final note
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders,
                width: { size: 9360, type: WidthType.DXA },
                shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
                margins: { top: 300, bottom: 300, left: 400, right: 400 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "This playbook is a living document.", bold: true, font: "Arial", size: 26, color: WHITE })], alignment: AlignmentType.CENTER }),
                  new Paragraph({ children: [new TextRun({ text: "Update it every quarter with new benchmark data, test results, and lessons learned.", font: "Arial", size: 22, color: "AACCEE" })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } }),
                  new Paragraph({ children: [new TextRun({ text: "The best playbook is one that reflects actual experience, not just theory.", font: "Arial", size: 22, color: "AACCEE", italics: true })], alignment: AlignmentType.CENTER })
                ]
              })
            ]
          })
        ]
      })

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/RE_Performance_Marketing_Playbook.docx", buffer);
  console.log("Done!");
});
