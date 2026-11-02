// Builds the Website Requirements Questionnaire for D S Seeds (OPC) Private Limited.
// Usage: node scripts/build_questionnaire.js  (run from repo root)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
  WidthType, BorderStyle, ShadingType, Header, Footer, PageNumber, PageBreak,
  TabStopType, HeadingLevel, TableOfContents,
} = require("docx");

const VERSION = "1.0";
const DATE = "2 November 2026";
const CLIENT = "D S Seeds (OPC) Private Limited";
const BRAND = "Embetronix Technologies";
const OUT = path.join(__dirname, "..", "docs", "01-requirements",
  `DS_Seeds_Website_Requirements_Questionnaire_v${VERSION}.docx`);

const PRIMARY = "1F4E5F";   // deep teal
const ACCENT = "E8F1F2";    // light tint
const LINE = "9AA5AB";
const FONT = "Calibri";
const CONTENT_W = 9638;     // A4 width 11906 - 2 x 1134 margins

// ---------- helpers ----------
let qCounter = {};
let currentSection = "";
const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 100 }, ...opts.para,
  children: [new TextRun({ text, font: FONT, size: 21, ...opts.run })],
});
const note = (text) => new Paragraph({
  spacing: { before: 40, after: 100 },
  children: [new TextRun({ text, font: FONT, size: 18, italics: true, color: "5A6A70" })],
});
const section = (letter, title, intro) => {
  currentSection = letter; qCounter[letter] = 0;
  const out = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1, pageBreakBefore: true,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: `Section ${letter}: ${title}`, font: FONT })],
    }),
  ];
  if (intro) out.push(note(intro));
  return out;
};
const sub = (title) => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 },
  children: [new TextRun({ text: title, font: FONT })],
});
const qnum = () => `${currentSection}${++qCounter[currentSection]}.`;
const question = (text, hint) => {
  const out = [new Paragraph({
    keepNext: true, spacing: { before: 200, after: 60 },
    children: [
      new TextRun({ text: qnum() + " ", font: FONT, size: 21, bold: true, color: PRIMARY }),
      new TextRun({ text, font: FONT, size: 21, bold: true }),
    ],
  })];
  if (hint) out.push(new Paragraph({
    keepNext: true, spacing: { after: 40 },
    children: [new TextRun({ text: hint, font: FONT, size: 18, italics: true, color: "5A6A70" })],
  }));
  return out;
};
const lines = (n = 1) => Array.from({ length: n }, () => new Paragraph({
  spacing: { before: 200, after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 1 } },
  children: [new TextRun({ text: "", font: FONT, size: 21 })],
}));
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};
// Checkbox grid; "Other" automatically gets a write-in line.
const checks = (options, cols = 3, other = true) => {
  const opts = [...options];
  if (other) opts.push("Other (please specify):");
  const w = Math.floor(CONTENT_W / cols);
  const rows = [];
  for (let i = 0; i < opts.length; i += cols) {
    const cells = [];
    for (let c = 0; c < cols; c++) {
      const t = opts[i + c];
      cells.push(new TableCell({
        borders: noBorders, width: { size: w, type: WidthType.DXA },
        margins: { top: 30, bottom: 30, left: 60, right: 60 },
        children: [new Paragraph({ children: t ? [
          new TextRun({ text: "☐  ", font: "Segoe UI Symbol", size: 22 }),
          new TextRun({ text: t, font: FONT, size: 20 }),
        ] : [] })],
      }));
    }
    rows.push(new TableRow({ cantSplit: true, children: cells }));
  }
  const out = [new Table({ width: { size: w * cols, type: WidthType.DXA },
    columnWidths: Array(cols).fill(w), borders: noBorders, rows })];
  if (other) out.push(...lines(1));
  return out;
};
const yesNo = (extra = []) => checks(["Yes", "No", ...extra], 4, false);
const cell = (text, width, { head = false, shade = null } = {}) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  shading: head ? { fill: PRIMARY, type: ShadingType.CLEAR, color: "auto" }
    : shade ? { fill: shade, type: ShadingType.CLEAR, color: "auto" } : undefined,
  margins: { top: 70, bottom: 70, left: 100, right: 100 },
  children: [new Paragraph({ children: [new TextRun({
    text, font: FONT, size: head ? 19 : 19, bold: head, color: head ? "FFFFFF" : "222222" })] })],
});
const grid = (headers, widths, rowsData, minRows = 0) => {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFC8CC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const mk = (cells, head) => new TableRow({ cantSplit: true, tableHeader: head,
    children: cells.map((t, i) => { const c = cell(t, widths[i], { head, shade: !head && i === 0 && t ? ACCENT : null }); c.options = c.options; return c; }) });
  const rows = [mk(headers, true), ...rowsData.map(r => mk(r, false))];
  for (let i = rowsData.length; i < minRows; i++) rows.push(mk(headers.map(() => ""), false));
  // apply borders
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths, borders: { ...borders, insideHorizontal: border, insideVertical: border }, rows });
};
const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });

// ---------- content ----------
const children = [];

// Cover page
children.push(
  new Paragraph({ spacing: { before: 1800, after: 120 }, children: [
    new TextRun({ text: BRAND.toUpperCase(), font: FONT, size: 24, bold: true, color: PRIMARY, characterSpacing: 40 })] }),
  new Paragraph({ spacing: { after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: PRIMARY, space: 6 } },
    children: [new TextRun({ text: "Website Requirements Questionnaire", font: FONT, size: 52, bold: true, color: "1A1A1A" })] }),
  new Paragraph({ spacing: { after: 600 }, children: [
    new TextRun({ text: "Company profile, product portfolio, content and asset requirements", font: FONT, size: 26, color: "4A5A60" })] }),
);
const coverRows = [
  ["Prepared for", CLIENT],
  ["Prepared by", BRAND],
  ["Document version", VERSION],
  ["Date issued", DATE],
  ["Date returned (to be filled by client)", ""],
  ["Filled in by (name and designation)", ""],
];
children.push(grid(["Detail", "Information"], [3600, 6038], coverRows));
children.push(new Paragraph({ children: [new PageBreak()] }));

// How to use
children.push(
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 160 },
    children: [new TextRun({ text: "About this questionnaire", font: FONT })] }),
  p("Thank you for choosing us to build your company website. This questionnaire helps us understand your business, your products and the information you want visitors to see. Your answers become the foundation of the website's content, structure and design."),
  p("You do not need any technical knowledge to fill it in. Please answer in your own words; short notes are perfectly fine, and we will refine the wording with you later."),
  sub("How to fill it in"),
  p("Where you see a box (☐), mark the options that apply. You can type an X next to it, replace it with ☒, or tick it by hand on a printed copy. More than one option may be marked unless the question says otherwise."),
  p("Write your answers on the blank lines below each question. If you need more space, add lines or attach a separate sheet and note the question number (for example, \"see attached sheet for D4\")."),
  p("If a question does not apply to your business, write \"Not applicable\". If you are unsure, write \"Not sure\" and we will discuss it during our review call."),
  p("Wherever we ask for documents, photographs or files, you can send them by email, WhatsApp or a shared Google Drive folder. Please name each file clearly (for example, \"Product_Paddy_DS101_field.jpg\")."),
  sub("What we will do with your answers"),
  p("Information you mark for public display will appear on the website. Everything else is for our internal reference only and will be kept confidential. Before the website goes live, you will review and approve every page."),
  sub("Sections in this questionnaire"),
);
const sectionList = [
  ["A", "Business identity and registration"],
  ["B", "Contact details and locations"],
  ["C", "About the company"],
  ["D", "Product portfolio and pricing"],
  ["E", "Website goals and target audience"],
  ["F", "Website pages and features"],
  ["G", "Newsletters, news and updates"],
  ["H", "Brand identity and design preferences"],
  ["I", "Photographs, videos and documents"],
  ["J", "Existing online presence"],
  ["K", "Legal and policy pages"],
  ["L", "Points of contact and approvals"],
  ["M", "Consolidated checklist of information and files"],
  ["N", "Confirmation"],
];
children.push(grid(["Section", "Topic"], [1400, 8238], sectionList));

// ---------- Section A ----------
children.push(...section("A", "Business identity and registration",
  "These details help us present your company accurately and build trust with buyers. For each item, please also tell us whether it may be shown publicly on the website."));
children.push(...question("Registered legal name of the company (exactly as on the Certificate of Incorporation)"), ...lines(1));
children.push(...question("Brand or trade name used in the market, if different from the legal name"), ...lines(1));
children.push(...question("Tagline or slogan, if any"), ...lines(1));
children.push(...question("Statutory registration details", "Please fill in the numbers you have. In the last column, write Yes if it may be shown on the website, or No if it is for our records only."));
children.push(grid(["Registration", "Number / details", "Show on website? (Yes / No)"], [3600, 3838, 2200], [
  ["Corporate Identification Number (CIN)", "", ""],
  ["Date of incorporation", "", ""],
  ["PAN of the company", "", ""],
  ["GSTIN (one per state, if more than one)", "", ""],
  ["Udyam (MSME) Registration Number", "", ""],
  ["Seed dealer / seed licence under the Seeds Act, 1966 and Seeds (Control) Order, 1983 (number, issuing state, validity)", "", ""],
  ["Seed production / processing licence, if applicable", "", ""],
  ["Importer Exporter Code (IEC), if applicable", "", ""],
  ["Trademark registration (brand name / logo) with application or registration number", "", ""],
  ["Plant variety registration under the PPV&FR Act, 2001, if any", "", ""],
  ["Other licences or registrations", "", ""],
]));
children.push(...question("Certifications and quality accreditations held", "Mark all that apply."));
children.push(...checks(["ISO 9001 (Quality Management)", "Seed certification by State Seed Certification Agency",
  "Truthfully Labelled (TL) seed compliance", "NABL-accredited seed testing laboratory", "DSIR-recognised in-house R&D unit",
  "OECD Seed Scheme", "Organic certification (NPOP)", "Startup India recognition", "None at present"], 2));
children.push(...question("Do you have scanned copies of the above certificates that can be displayed on the website?"), ...yesNo(["Some of them"]));
children.push(...question("Associations, memberships or affiliations", "For example, National Seed Association of India (NSAI), state seed associations, chambers of commerce, ICAR or agricultural university tie-ups."), ...lines(2));
children.push(...question("Awards, recognitions or government empanelments received (with year)"), ...lines(3));

// ---------- Section B ----------
children.push(...section("B", "Contact details and locations",
  "Making it easy for buyers to reach you is the main purpose of this website, so please be as complete as possible."));
children.push(...question("Registered office address (full address with PIN code)"), ...lines(2));
children.push(...question("Other locations you would like to show on the website", "Add rows as needed. Type examples: corporate office, processing plant, warehouse, R&D or trial farm, branch or regional office."));
children.push(grid(["Location type", "Full address with PIN code", "Phone", "Show on website?"], [2000, 4238, 1800, 1600], [], 5));
children.push(...question("Primary phone number(s) for enquiries"), ...lines(1));
children.push(...question("WhatsApp number for buyer enquiries (if different from the above)"), ...lines(1));
children.push(...question("Toll-free number, if any"), ...lines(1));
children.push(...question("Email address(es) to display on the website", "For example, a general enquiries address and a separate dealer or sales address."), ...lines(2));
children.push(...question("Where should enquiries submitted through the website be delivered?", "Mark all that apply and write the address or number."));
children.push(...checks(["Email (address):", "WhatsApp (number):", "SMS (number):"], 1, false), ...lines(1));
children.push(...question("Who will respond to website enquiries, and what is your target response time?"), ...lines(2));
children.push(...question("Business days and working hours"), ...lines(1));
children.push(...question("Do you want a Google Map of your office or plant on the Contact page?"), ...yesNo());
children.push(...question("Regional or state-wise sales contacts to be listed (name, area covered, phone)", "Leave blank if all enquiries go to one central number."));
children.push(grid(["Name", "Designation", "States / districts covered", "Phone / WhatsApp", "Email"], [1900, 1700, 2238, 1900, 1900], [], 5));

// ---------- Section C ----------
children.push(...section("C", "About the company",
  "This section shapes the \"About Us\" page, which builds credibility with dealers, distributors and institutional buyers."));
children.push(...question("In two or three sentences, how would you describe D S Seeds to a new buyer?"), ...lines(3));
children.push(...question("Company history: when and why was the business started, and what are the key milestones so far?", "Year-wise milestones work well, for example: 2018 started trading, 2021 opened processing unit."), ...lines(5));
children.push(...question("Vision statement (where you want the company to be in the future)"), ...lines(2));
children.push(...question("Mission statement (what the company does every day to get there)"), ...lines(2));
children.push(...question("Core values or guiding principles"), ...lines(2));
children.push(...question("What makes your seeds and your company different from competitors? (top three to five reasons)"), ...lines(4));
children.push(...question("Founder / Director profile", "Name, designation, qualifications, years of experience, and a short message from the Director, if you wish to include one. Please also send a professional photograph."), ...lines(5));
children.push(...question("Do you want to show key team members on the website?"), ...yesNo());
children.push(grid(["Name", "Designation", "Qualification / expertise", "Photo available?"], [2400, 2400, 3038, 1800], [], 5));
children.push(...question("Infrastructure and facilities", "Mark all that you have and would like to describe on the website."));
children.push(...checks(["Seed processing plant", "Grading and cleaning machinery", "Seed treatment facility", "Packaging unit",
  "Warehouse / storage", "Cold storage / controlled-humidity storage", "Seed testing laboratory", "Research and development farm",
  "Demonstration / trial plots", "Own seed production area", "Contract seed production with farmers"], 2));
children.push(...question("Please describe the capacity or scale where relevant", "For example, processing capacity in tonnes per day, storage capacity, area under seed production in acres, number of contract farmers."), ...lines(3));
children.push(...question("Describe your quality assurance process", "For example, field inspection, germination testing, genetic purity testing, moisture checks, seed treatment."), ...lines(4));
children.push(...question("Research and development activities, if any", "Breeding programmes, variety development, collaborations with universities or institutes."), ...lines(3));
children.push(...question("Geographic reach: which states or regions do you currently supply to?"), ...lines(2));
children.push(...question("Approximate size of your dealer / distributor network (number of dealers, states)"), ...lines(1));
children.push(...question("Do you export? If yes, to which countries?"), ...yesNo(), ...lines(1));
children.push(...question("Notable customers or partners you would like to mention", "For example, government departments, FPOs, NGOs, large growers. Please confirm you have permission to name them."), ...lines(3));
children.push(...question("Corporate social responsibility or farmer-welfare activities you would like to highlight", "For example, farmer training camps, field days, demonstrations, advisory services."), ...lines(3));

// ---------- Section D ----------
children.push(...section("D", "Product portfolio and pricing",
  "The product catalogue is the heart of the website. We will create one page for each product so buyers can understand its features and contact you about it."));
children.push(sub("Product range overview"));
children.push(...question("Which crop categories do you deal in?", "Mark all that apply."));
children.push(...checks(["Paddy / rice", "Wheat", "Maize", "Millets (bajra, jowar, ragi, etc.)", "Pulses (arhar, moong, urad, gram, etc.)",
  "Oilseeds (mustard, groundnut, soybean, sunflower, sesame)", "Cotton", "Jute", "Vegetables", "Fruits / melons",
  "Flowers", "Fodder / forage crops", "Green manure crops", "Spices", "Microgreens / exotic"], 3));
children.push(...question("Types of seed you offer", "Mark all that apply."));
children.push(...checks(["Hybrid seeds", "Open-pollinated / improved varieties", "Research (proprietary) varieties",
  "Public-sector (ICAR / SAU) varieties", "Foundation seed", "Certified seed", "Truthfully Labelled seed", "Organic seed", "Treated seed"], 3));
children.push(...question("Approximately how many products (varieties / hybrids) will be listed on the website at launch?"), ...lines(1));
children.push(...question("How often do you add new products or discontinue old ones?"));
children.push(...checks(["Every season", "Once a year", "Occasionally", "Rarely"], 4));
children.push(...question("Do you also sell or plan to list non-seed products?", "For example, fertilisers, bio-stimulants, crop protection, farm tools. These may be shown in a separate section."), ...yesNo(), ...lines(1));
children.push(...question("Are there brand names or sub-brands under which products are sold? Please list them."), ...lines(2));
children.push(...question("Do you have an existing product catalogue or brochure (printed or PDF)?", "If yes, please share it. It often answers many of the questions below."), ...yesNo());

children.push(sub("Pricing"));
children.push(...question("How would you like prices to appear on the website?", "Mark one."));
children.push(...checks(["MRP shown for each pack size", "\"Starting from\" price only", "Price shown only for some products",
  "\"Price on request\" with an enquiry button", "Separate dealer / bulk price list available on request"], 2));
children.push(...question("Should prices be shown inclusive of GST?"), ...yesNo(["Seeds are GST-exempt", "Not sure"]));
children.push(...question("How often do prices change, and who in your team will send us updated prices?"), ...lines(2));
children.push(...question("Should the website mention minimum order quantities, bulk discounts or dealer margins?"), ...yesNo(), ...lines(1));
children.push(...question("Should prices vary by state or region?"), ...yesNo(), ...lines(1));

children.push(sub("Product information sheet"));
children.push(note("Please fill in one sheet for each product. You can copy this table as many times as needed, or ask us for an Excel version of this sheet if that is easier for a long product list."));
children.push(...question("Preferred format for providing product details"));
children.push(...checks(["This Word document", "Excel sheet (we will send a template)", "Existing catalogue / brochure"], 3, false));
const productRows = [
  ["Product / brand name", ""], ["Variety or hybrid code", ""], ["Crop", ""],
  ["Seed type (hybrid / OP / research / certified / TL)", ""], ["Season (Kharif / Rabi / Zaid / all seasons)", ""],
  ["Recommended states / agro-climatic zones", ""], ["Suitable soil type and irrigation (irrigated / rainfed)", ""],
  ["Sowing window (months)", ""], ["Duration to maturity / first harvest (days)", ""],
  ["Yield potential (per acre or hectare)", ""], ["Plant characteristics (height, habit, tillers, etc.)", ""],
  ["Grain / fruit / pod characteristics (colour, shape, size, weight)", ""],
  ["Tolerance or resistance to pests and diseases", ""], ["Tolerance to stress (drought, flood, heat, salinity, lodging)", ""],
  ["Seed rate (per acre or hectare)", ""], ["Recommended spacing", ""],
  ["Key selling points (three to five short points)", ""], ["Market preference (e.g. good cooking quality, long shelf life)", ""],
  ["Germination % and physical / genetic purity %", ""], ["Seed treatment applied, if any", ""],
  ["Pack sizes available", ""], ["Price / MRP for each pack size", ""],
  ["Package of practices or crop guide available? (Yes / No)", ""],
  ["Farmer testimonials available for this product? (Yes / No)", ""],
  ["Photographs available (packet, crop in field, harvested produce)", ""],
  ["Anything that must NOT be shown publicly", ""],
];
children.push(grid(["Field", "Details"], [4000, 5638], productRows));

// ---------- Section E ----------
children.push(...section("E", "Website goals and target audience",
  "Knowing who will visit the website and what you want them to do helps us decide what to highlight on each page."));
children.push(...question("What are the most important goals for the website?", "Mark all that apply, then number your top three in order of importance on the line below."));
children.push(...checks(["Get enquiries from new buyers", "Recruit new dealers / distributors", "Show the full product catalogue",
  "Build trust and credibility", "Educate farmers on crop practices", "Share company news and newsletters",
  "Support tenders and institutional sales", "Attract export buyers", "Recruit staff"], 3));
children.push(...lines(1));
children.push(...question("Who are your main buyers?", "Mark all that apply."));
children.push(...checks(["Individual farmers", "Progressive / large farmers", "Seed dealers and retailers", "Distributors",
  "Farmer Producer Organisations (FPOs)", "Co-operative societies", "Government departments and agencies",
  "NGOs and development programmes", "Agri-input companies", "Export buyers"], 3));
children.push(...question("Which buyer group is the most important to attract through the website?"), ...lines(1));
children.push(...question("When a buyer visits the website, what is the one thing you most want them to do?", "Mark one."));
children.push(...checks(["Call us", "Message us on WhatsApp", "Fill in an enquiry form", "Apply to become a dealer",
  "Download the product catalogue", "Visit our office or plant"], 3));
children.push(...question("Which languages should the website be available in?"));
children.push(...checks(["English", "Hindi", "Bengali", "Odia", "Assamese", "Telugu", "Marathi", "Punjabi"], 4));
children.push(...question("If more than one language: will you provide translated content, or should we arrange translation?"));
children.push(...checks(["We will provide translations", "Please arrange translation", "Not sure yet"], 3, false));
children.push(...question("Who are your main competitors? (company names and website addresses, if known)"), ...lines(3));
children.push(...question("How do buyers find you today?", "Mark all that apply."));
children.push(...checks(["Word of mouth", "Dealer network", "Field staff", "Agricultural fairs / Krishi melas", "Social media",
  "Google search", "IndiaMART / TradeIndia", "Newspaper / radio advertising", "Government tenders"], 3));

// ---------- Section F ----------
children.push(...section("F", "Website pages and features",
  "Below is a list of pages commonly found on seed company websites. Mark the ones you want. We will recommend a final structure after reviewing your answers."));
children.push(...question("Which pages would you like on the website?"));
children.push(...checks(["Home", "About Us", "Director's message", "Our team", "Infrastructure / facilities", "Quality assurance",
  "Research and development", "Products (catalogue)", "Individual product pages", "Crop guides / package of practices",
  "Newsletters", "News and events", "Photo gallery", "Video gallery", "Farmer testimonials", "Certificates and licences",
  "Become a dealer", "Dealer / distributor locator", "Downloads (brochures, catalogues)", "Frequently asked questions",
  "Careers", "Contact Us"], 3));
children.push(...question("Should buyers be able to search or filter products?", "Mark the filters that would be useful."));
children.push(...checks(["Search by name", "Filter by crop", "Filter by season", "Filter by state / region", "Filter by seed type",
  "Filter by duration", "No filters needed"], 3, false));
children.push(...question("What information should the enquiry form ask for?", "Mark all that apply. Fewer fields usually mean more enquiries."));
children.push(...checks(["Name", "Mobile number", "Email", "State", "District", "Buyer type (farmer / dealer / other)",
  "Product of interest", "Quantity required", "Message"], 3));
children.push(...question("If you want a \"Become a dealer\" page, what information do you need from applicants?", "For example, firm name, location, existing licence, GSTIN, years in business, current brands handled."), ...lines(3));
children.push(...question("Contact options to show on every page"));
children.push(...checks(["Click-to-call button", "WhatsApp chat button", "Enquiry form button", "Email link"], 4, false));
children.push(...question("Would you like visitors to download any documents?", "Mark all that apply and please supply the files."));
children.push(...checks(["Company profile", "Product catalogue", "Price list", "Crop guides", "Dealer application form", "Certificates"], 3));
children.push(...question("Should any part of the website be visible only to registered dealers (for example, a dealer price list)?"), ...yesNo(["Maybe later"]));
children.push(...question("Do you need a Careers page with current job openings? If yes, who will provide openings?"), ...yesNo(), ...lines(1));
children.push(...question("Frequently asked questions you receive from buyers", "List the questions you are asked most often; we can turn these into an FAQ page."), ...lines(5));
children.push(...question("Anything else you have seen on another website that you would like on yours?"), ...lines(3));

// ---------- Section G ----------
children.push(...section("G", "Newsletters, news and updates",
  "You mentioned that you want to publish newsletters. This section helps us understand what that means for you so we can plan the right setup."));
children.push(...question("What do you mean by \"newsletter\"?", "Mark all that apply."));
children.push(...checks(["A PDF newsletter that visitors can read or download", "Articles / news posts on the website",
  "Email newsletters sent to subscribers", "WhatsApp broadcast updates"], 2));
children.push(...question("Do you want visitors to be able to subscribe to receive newsletters?"));
children.push(...checks(["Yes, by email", "Yes, by WhatsApp", "No"], 3, false));
children.push(...question("How often will you publish a newsletter or update?"));
children.push(...checks(["Weekly", "Monthly", "Quarterly", "Every season (Kharif / Rabi)", "Occasionally"], 3, false));
children.push(...question("What topics will the newsletters cover?", "Mark all that apply."));
children.push(...checks(["New product launches", "Seasonal sowing advisories", "Crop care tips", "Field day and event reports",
  "Farmer success stories", "Company announcements", "Government schemes and subsidies", "Market and weather updates"], 2));
children.push(...question("Who will write the newsletter content, and who will approve it before publishing?"), ...lines(2));
children.push(...question("Do you have past newsletters, press releases or news articles to add to the website at launch?"), ...yesNo(), ...lines(1));
children.push(...question("Would you like to update news and newsletters yourselves, or send them to us to publish?"));
children.push(...checks(["We will update it ourselves", "Please publish it for us", "Not sure yet"], 3, false));
children.push(...question("If you will update it yourselves, who in your team will do this, and how comfortable are they with computers?"), ...lines(2));

// ---------- Section H ----------
children.push(...section("H", "Brand identity and design preferences",
  "These answers guide how the website looks and feels."));
children.push(...question("Do you have a company logo?"), ...yesNo(["Needs redesign"]));
children.push(...question("In which formats is your logo available?", "Original design files give the sharpest result. Ask your logo designer or printer if you are not sure."));
children.push(...checks(["AI / EPS / SVG / PDF (original design file)", "High-resolution PNG with transparent background",
  "JPG only", "Only on printed material", "Not sure"], 2, false));
children.push(...question("Brand colours", "Colour codes if you know them, or describe them (for example, \"dark green and yellow as on our seed packets\")."), ...lines(2));
children.push(...question("Do you have brand guidelines, or fonts used on your packaging and brochures?"), ...yesNo(), ...lines(1));
children.push(...question("How should the website feel to a visitor?", "Mark up to three."));
children.push(...checks(["Modern and clean", "Earthy and natural", "Traditional and trustworthy", "Scientific and research-led",
  "Bold and energetic", "Premium", "Simple and farmer-friendly"], 3));
children.push(...question("Websites you like (any industry) and what you like about them", "Please give at least two or three examples."));
children.push(grid(["Website address", "What you like about it"], [3800, 5838], [], 3));
children.push(...question("Websites you do not like, and why"));
children.push(grid(["Website address", "What you dislike about it"], [3800, 5838], [], 2));
children.push(...question("Should the website design match your seed packet and brochure design?"), ...yesNo(["Partly"]));
children.push(...question("Any colours, images or styles that must be avoided?"), ...lines(2));

// ---------- Section I ----------
children.push(...section("I", "Photographs, videos and documents",
  "Good, original photographs make the biggest difference to how professional a website looks. Please send the highest-quality versions you have; photos forwarded through WhatsApp lose quality, so original files are preferred."));
children.push(...question("Which of the following photographs can you provide?", "Mark all that you have."));
children.push(...checks(["Seed packets / packaging (each product)", "Crops in the field (each product)", "Harvested produce / grain / fruit",
  "Close-up of seeds", "Office building", "Processing plant and machinery", "Laboratory", "Warehouse / storage",
  "R&D / trial farm", "Director / founder portrait", "Team photographs", "Farmer meetings and field days",
  "Trade fairs and events", "Awards and certificates", "Happy farmers with your crops"], 3));
children.push(...question("Would you like us to arrange a professional photo or video shoot?"), ...yesNo(["Discuss later"]));
children.push(...question("If not, may we use licensed stock photographs where your own are not available (for example, general farm scenes)?"), ...yesNo());
children.push(...question("Do you have videos to show on the website?", "For example, a company film, product demonstrations, farmer testimonials, field day recordings, or a YouTube channel."), ...yesNo(), ...lines(2));
children.push(...question("Farmer or dealer testimonials", "Name, village / district, state, crop and product used, their comment, and a photo if possible. Please confirm the person has agreed to be featured."));
children.push(grid(["Name", "Location", "Product used", "Comment / result", "Photo? Consent?"], [1700, 1800, 1700, 2738, 1700], [], 4));
children.push(...question("Who will write the text (content) for the website?"));
children.push(...checks(["We will write it", "Please write it based on our inputs", "Combination of both"], 3, false));
children.push(...question("Do you own the rights to all photographs and material you will send us?", "If some photos were taken by an outside photographer or agency, please confirm you have permission to use them online."), ...yesNo(["Not sure"]));

// ---------- Section J ----------
children.push(...section("J", "Existing online presence",
  "Simple questions about what you already have. If you are unsure about any of these, just write \"Not sure\" and we will check for you."));
children.push(...question("Do you already have a website? If yes, what is its address?"), ...yesNo(), ...lines(1));
children.push(...question("Do you already own a website name (domain), such as dsseeds.com? If yes, which one?"), ...yesNo(["Not sure"]), ...lines(1));
children.push(...question("If not, please suggest two or three website names you would like, in order of preference"), ...lines(2));
children.push(...question("Do you currently use company email addresses (for example, info@yourcompany.com)?"), ...yesNo(["Not sure"]), ...lines(1));
children.push(...question("Social media pages to link from the website"));
children.push(grid(["Platform", "Page address / handle", "Who manages it?"], [2400, 4638, 2600], [
  ["Facebook", "", ""], ["Instagram", "", ""], ["YouTube", "", ""], ["LinkedIn", "", ""], ["X (Twitter)", "", ""], ["Other", "", ""],
]));
children.push(...question("Do you have a Google Business Profile (your listing on Google Maps)?"), ...yesNo(["Not sure"]));
children.push(...question("Are you listed on online business directories or marketplaces?"));
children.push(...checks(["IndiaMART", "TradeIndia", "JustDial", "Agri-input apps (e.g. BigHaat, AgroStar)", "Government e-Marketplace (GeM)", "None"], 3));

// ---------- Section K ----------
children.push(...section("K", "Legal and policy pages",
  "Every business website in India should carry a few standard policy pages. Please mark the ones you want; we can prepare draft text for your review, which you should have checked by your legal adviser before publishing."));
children.push(...question("Which policy pages would you like on the website?"));
children.push(...checks(["Privacy policy (how enquiry details are used)", "Terms of use", "Disclaimer (results may vary with conditions)",
  "Seed replacement / complaint policy", "Refund policy", "Cookie policy", "Grievance redressal contact",
  "Copyright notice"], 2));
children.push(...question("Do you already have any of these policies written (for packaging, dealers or other purposes)? Please share them."), ...yesNo(), ...lines(1));
children.push(...question("What is your current policy if a buyer reports a germination or quality problem?", "This helps us draft a clear complaint or replacement policy."), ...lines(3));
children.push(...question("Is there any specific statement that must appear with product claims?", "For example, \"Yield depends on soil, climate and crop management practices.\""), ...lines(2));
children.push(...question("Name and contact of the person to be listed for complaints or grievances, if any"), ...lines(1));

// ---------- Section L ----------
children.push(...section("L", "Points of contact and approvals"));
children.push(...question("Main point of contact for this project"));
children.push(grid(["Detail", "Information"], [3200, 6438], [
  ["Name", ""], ["Designation", ""], ["Mobile / WhatsApp", ""], ["Email", ""], ["Best time to call", ""],
]));
children.push(...question("Who gives final approval of content and design? (if different from above)"), ...lines(1));
children.push(...question("Preferred way of communicating with us"));
children.push(...checks(["Phone call", "WhatsApp", "Email", "In-person meeting", "Video call"], 3, false));
children.push(...question("Preferred language for discussions"));
children.push(...checks(["English", "Hindi", "Bengali"], 3));
children.push(...question("Is there a target date or event by which you need the website live?", "For example, a product launch, Kharif / Rabi season start, or a trade fair."), ...lines(1));
children.push(...question("Anything else you would like us to know?"), ...lines(4));

// ---------- Section M ----------
children.push(...section("M", "Consolidated checklist of information and files",
  "Use this list to collect everything in one go. Mark the last column as you send each item."));
const assets = [
  ["Company logo", "AI / EPS / SVG / PDF, or high-resolution PNG", "H2"],
  ["Certificate of Incorporation and licence copies (if to be displayed)", "PDF or clear scan", "A4, A6"],
  ["Certification and award certificates", "PDF or clear scan", "A5, A8"],
  ["Existing catalogue, brochure, leaflets", "PDF or print copy", "D7"],
  ["Product details for every product", "Product sheet or Excel", "D13 sheet"],
  ["Current price list", "PDF or Excel", "D8 to D12"],
  ["Product photographs (packet, field, produce)", "Original JPG / PNG", "I1"],
  ["Facility and infrastructure photographs", "Original JPG / PNG", "I1"],
  ["Director and team photographs", "Original JPG / PNG", "C7, C8"],
  ["Event and farmer meeting photographs", "Original JPG / PNG", "I1"],
  ["Company videos or YouTube links", "MP4 or link", "I4"],
  ["Farmer / dealer testimonials with consent", "Text and photo", "I5"],
  ["Past newsletters, press releases, news", "PDF / Word / links", "G6"],
  ["Crop guides / package of practices", "PDF / Word", "F6"],
  ["Dealer application form, if any", "PDF / Word", "F4"],
  ["Existing policy documents", "PDF / Word", "K2"],
  ["Social media page links", "Links", "J5"],
];
children.push(grid(["Item", "Preferred format", "Related question", "Sent? (☐)"],
  [3900, 2700, 1638, 1400], assets.map(a => [...a, "☐"])));
children.push(note("Question references in this checklist are indicative. Please send whatever you have; missing items can be added later."));

// ---------- Section N ----------
children.push(...section("N", "Confirmation"));
children.push(p("I confirm that the information provided in this questionnaire is accurate to the best of my knowledge, and that D S Seeds (OPC) Private Limited holds the rights to the logos, photographs and materials supplied for use on the website."));
children.push(spacer());
children.push(grid(["Detail", "Information"], [3200, 6438], [
  ["Name", ""], ["Designation", ""], ["Signature", ""], ["Company seal (if applicable)", ""], ["Date", ""], ["Place", ""],
]));
children.push(spacer());
children.push(note("Please return the completed questionnaire to Embetronix Technologies. After we receive it, we will schedule a short review call to go through your answers and agree the website structure."));

// ---------- document ----------
const doc = new Document({
  creator: BRAND,
  title: `Website Requirements Questionnaire - ${CLIENT}`,
  description: "Client questionnaire for website requirements",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: FONT, color: PRIMARY }, paragraph: { spacing: { before: 0, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: FONT, color: "2E6E80" }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1300, right: 1134, bottom: 1134, left: 1134, header: 600, footer: 500 } }, titlePage: true },
    headers: {
      default: new Header({ children: [new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: PRIMARY, space: 4 } },
        children: [
          new TextRun({ text: BRAND, font: FONT, size: 18, bold: true, color: PRIMARY }),
          new TextRun({ text: "\tWebsite Requirements Questionnaire", font: FONT, size: 18, color: "5A6A70" }),
        ] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: `Confidential – prepared for ${CLIENT}  |  v${VERSION}`, font: FONT, size: 16, color: "7A868B" }),
          new TextRun({ children: ["\tPage ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: "7A868B" }),
        ] })] }),
      first: new Footer({ children: [new Paragraph({ children: [
        new TextRun({ text: "Confidential – for client use only", font: FONT, size: 16, color: "7A868B" })] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log("Wrote", OUT); });
