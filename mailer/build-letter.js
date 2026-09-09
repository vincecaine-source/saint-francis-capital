const fs = require('fs');
const path = require('path');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType,
  VerticalAlign, PageNumber,
} = docx;

const DIR = __dirname;
const img = (name) => fs.readFileSync(path.join(DIR, name));
const sq = (s) => s.replace(/'/g, '\u2019');

const NAVY = '0B1F3A';
const GOLD = 'A97C2F';
const MUTED = '5B6472';
const HAIRLINE = 'D8CFB8';

// ---- helpers -------------------------------------------------------------

function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 268 },
    alignment: AlignmentType.LEFT,
    children: [new TextRun({ text: sq(text), font: 'Calibri', size: 23, color: '1A1A1A', ...opts })],
  });
}

function bodyParaRuns(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: 200, line: 300 },
    alignment: AlignmentType.LEFT,
    children: runs,
    ...opts,
  });
}

// ---- letterhead header -----------------------------------------------

const logoDims = { w: 1465, h: 743 };
const logoWidth = 205;
const logoHeight = Math.round(logoWidth * (logoDims.h / logoDims.w));

const header = new Header({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [
        new ImageRun({
          type: 'png',
          data: img('letter-logo.png'),
          transformation: { width: logoWidth, height: logoHeight },
          altText: { title: 'Saint Francis Capital', description: 'Saint Francis Capital logo', name: 'logo' },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: HAIRLINE, space: 8 } },
      children: [
        new TextRun({
          text: '359 ROCKWELL CHURCH ROAD NE, WINDER, GA 30680   |   800.832.9682   |   STFRANCAP.COM',
          font: 'Calibri',
          size: 14,
          color: MUTED,
          characterSpacing: 10,
        }),
      ],
    }),
  ],
});

const footer = new Footer({
  children: [
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE, space: 8 } },
      spacing: { before: 80 },
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: 'Saint Francis Capital is an independent Registered Investment Advisor and is not affiliated with, endorsed by, or sponsored by United Airlines, ALPA, or Charles Schwab. This letter is for informational purposes only and does not constitute investment, legal, or tax advice. Personal experiences described are individual to the author and are not indicative of future investment results.',
          font: 'Calibri',
          size: 12,
          color: MUTED,
          italics: true,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60 },
      children: [
        new TextRun({ text: '359 Rockwell Church Road NE, Winder, GA 30680  ·  800.832.9682  ·  stfrancap.com  ·  Page ', font: 'Calibri', size: 12, color: MUTED }),
        new TextRun({ children: [PageNumber.CURRENT], font: 'Calibri', size: 12, color: MUTED }),
        new TextRun({ text: ' of ', font: 'Calibri', size: 12, color: MUTED }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Calibri', size: 12, color: MUTED }),
      ],
    }),
  ],
});

// ---- body content ----------------------------------------------------

const dateline = new Paragraph({
  spacing: { after: 260 },
  children: [new TextRun({ text: 'September 2026', font: 'Calibri', size: 22, color: '1A1A1A' })],
});

const salutation = new Paragraph({
  spacing: { after: 200 },
  children: [new TextRun({ text: 'Dear Captain [Last Name],', font: 'Calibri', size: 23, color: '1A1A1A' })],
});

const openHook = new Paragraph({
  spacing: { after: 220 },
  children: [
    new TextRun({
      text: 'Every pilot has one flight he never forgets.',
      font: 'Georgia', size: 25, color: NAVY, italics: true, bold: true,
    }),
  ],
});

const storyParas = [
  `Mine was in the winter of 2005, flying the left seat of a Boeing 737-800 into Oakland. It was my leg — I had the controls, and my First Officer had the radio and the checklists.`,
  `It was an El Niño year, and we picked up moderate turbulence for the last hour of the flight at 36,000 feet. That held until Oakland Center cleared us to begin our descent. That's when it got worse — a lot worse.`,
  `For the next thirty minutes, we flew through nearly continuous severe turbulence. The forecast had called for moderate chop on the way down. In roughly 22,000 flight hours by that point, I had never experienced continuous severe turbulence. This was it.`,
  `Ten minutes before descent, I'd told the flight attendants to secure the cabin and strap in. They told me afterward that everything left sitting on empty seats — laptops, phones, a baby bottle — found new homes rows away from where it started.`,
  `We were cleared into Oakland, runway 30. The ceiling was around 3,000 feet with good visibility. The winds were out of 250 degrees at 25 knots, gusting to 35 — about 50 degrees off our nose, and close to the maximum crosswind limit for the aircraft. Every airline aircraft has one, for a simple reason: you don't want to drag a wingtip on landing.`,
  `We broke out of the clouds about 3,000 feet up, five miles out, right on the centerline — I remember being surprised by that, looking out the First Officer's side window to find the runway. We held that crab angle all the way through the flare, which is where it all comes together, or it doesn't: the landing gear has to align with the runway, so you add right rudder with your right foot while adding left aileron with your hands to counteract it. Done right, the airplane looks banked to the left without actually turning — aligned with the runway, left wheel touching down first, closer to the ground than usual but exactly where it needs to be.`,
  `Most days in that seat, you think, they actually pay me for this. This was not most days.`,
  `The landing, when it came, was the smoothest, most centered landing of my entire career.`,
];

const moralPara = new Paragraph({
  spacing: { before: 60, after: 200 },
  children: [
    new TextRun({ text: 'Sometimes you prepare for the worst, and you turn out the best.'.replace(/'/g, '’'), font: 'Georgia', size: 24, color: NAVY, bold: true, italics: true }),
  ],
});

const bridgeParas = [
  `I've thought about that landing many times in the years since — not just as a pilot, but as someone who now spends his days helping fellow aviators plan for a different kind of final approach: retirement.`,
  `The parallel isn't subtle. We didn't get through that descent by hoping for calm air. We got through it because the checklists, the crosswind charts, and the procedures were already second nature long before we needed them. The landing wasn't luck — it was preparation meeting discipline, at exactly the moment it mattered.`,
  `Retirement works the same way. Markets get turbulent. Pension elections, Social Security timing, tax brackets, and TIP decisions all shift underneath you, often with little warning. The pilots who land it well aren't the ones who avoided the turbulence — they're the ones who had a plan built for it long before the descent began.`,
];

const founderPara = new Paragraph({
  spacing: { after: 180 },
  children: [
    new TextRun({ text: sq(`That's why I started Saint Francis Capital in 2013 — to build that kind of plan for pilots, by a pilot. Over more than three decades in the cockpit, I flew Boeing 727, 737, 757, 767, and DC-10 aircraft based out of Denver, Guam, Manila, Los Angeles, Houston, Cleveland, Honolulu, and Newark, and served as an elected union officer for my fellow pilots. I built this firm to bring the same discipline to retirement planning for United pilots that we brought to the flight deck — and every month, I share more of that thinking directly with clients in our Monthly Client Letter.`), font: 'Calibri', size: 23, color: '1A1A1A' }),
  ],
});

const pullQuoteBlock = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [9000],
  borders: {
    top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    left: { style: BorderStyle.SINGLE, size: 24, color: GOLD },
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: { size: 9000, type: WidthType.DXA },
          margins: { top: 100, bottom: 100, left: 260, right: 160 },
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            left: { style: BorderStyle.SINGLE, size: 24, color: GOLD },
          },
          children: [
            new Paragraph({
              spacing: { after: 80 },
              children: [
                new TextRun({
                  text: '\u201CYou\u2019ve spent a career preparing for every flight. Your final approach to retirement deserves the same care.\u201D',
                  font: 'Georgia', size: 24, italics: true, color: NAVY,
                }),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: '— Captain Kaye Riggs, Founder', font: 'Calibri', size: 20, color: MUTED })],
            }),
          ],
        }),
      ],
    }),
  ],
});

const ctaPara = new Paragraph({
  spacing: { before: 180, after: 160 },
  children: [
    new TextRun({
      text: `I recorded a short video — about ninety seconds — sharing more of this story and what it means for the decisions ahead of you. Scan the code on this page to watch it, or reach out directly to my colleague Kelly Holliday to schedule a complimentary conversation about where you stand.`,
      font: 'Calibri', size: 23, color: '1A1A1A',
    }),
  ],
});

const closing1 = new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: 'Warm regards,', font: 'Calibri', size: 23, color: '1A1A1A' })] });

// ---- signature + QR two-column block ----------------------------------

const photoDims = { w: 300, h: 300 };
const photoWidth = 62;
const photoHeight = Math.round(photoWidth * (photoDims.h / photoDims.w));

const qrDims = { w: 1080, h: 1248 };
const qrWidth = 155;
const qrHeight = Math.round(qrWidth * (qrDims.h / qrDims.w));

const signatureCell = new TableCell({
  width: { size: 5600, type: WidthType.DXA },
  verticalAlign: VerticalAlign.TOP,
  borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  margins: { top: 60, right: 260, bottom: 0, left: 0 },
  children: [
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new ImageRun({
          type: 'png',
          data: img('kaye-signature-photo.png'),
          transformation: { width: photoWidth, height: photoHeight },
          altText: { title: 'Captain Kaye Riggs', description: 'Photo of Captain Kaye Riggs', name: 'founder-photo' },
        }),
      ],
    }),
    new Paragraph({ spacing: { after: 10 }, children: [new TextRun({ text: 'Captain Kaye Riggs', font: 'Georgia', size: 24, bold: true, color: NAVY })] }),
    new Paragraph({ spacing: { after: 10 }, children: [new TextRun({ text: 'Founder, Saint Francis Capital', font: 'Calibri', size: 20, color: '1A1A1A' })] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: '(737/757/767/CASA-212)', font: 'Calibri', size: 18, color: MUTED })] }),
    new Paragraph({
      spacing: { after: 6 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE, space: 6 } },
      children: [new TextRun({ text: 'Kelly Holliday  ·  Client Service', font: 'Calibri', size: 18, bold: true, color: NAVY })],
    }),
    new Paragraph({ spacing: { after: 2 }, children: [new TextRun({ text: '770.851.0934', font: 'Calibri', size: 18, color: '1A1A1A' })] }),
    new Paragraph({ children: [new TextRun({ text: 'kelly.holliday@stfrancap.com', font: 'Calibri', size: 18, color: '1A1A1A' })] }),
  ],
});

const qrCell = new TableCell({
  width: { size: 3800, type: WidthType.DXA },
  verticalAlign: VerticalAlign.TOP,
  borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  margins: { top: 60, left: 260, bottom: 0, right: 0 },
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          type: 'png',
          data: img('../endcard/qr-pilot-message-print.png'),
          transformation: { width: qrWidth, height: qrHeight },
          altText: { title: 'QR code', description: 'Scan to watch a 90-second message from Saint Francis Capital', name: 'qr-code' },
        }),
      ],
    }),
  ],
});

const signatureTable = new Table({
  width: { size: 9400, type: WidthType.DXA },
  columnWidths: [5600, 3800],
  rows: [new TableRow({ children: [signatureCell, qrCell] })],
});

// ---- assemble document --------------------------------------------------

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Calibri', size: 23, color: '1A1A1A' } } },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1500, right: 1350, bottom: 1300, left: 1350 },
        },
      },
      headers: { default: header },
      footers: { default: footer },
      children: [
        dateline,
        salutation,
        openHook,
        ...storyParas.map((t) => bodyPara(t)),
        moralPara,
        ...bridgeParas.map((t) => bodyPara(t)),
        founderPara,
        pullQuoteBlock,
        ctaPara,
        closing1,
        signatureTable,
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(DIR, 'Pilot-Story-Letter.docx'), buf);
  console.log('Saved Pilot-Story-Letter.docx');
});
