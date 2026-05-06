const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak, TabStopType,
  TabStopPosition, Header, Footer, ImageRun
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── COLORS ────────────────────────────────────────────────────────────────
const C = {
  blue:    "1A56B0",
  ltblue:  "D6E4F7",
  navy:    "0D3B6E",
  green:   "1B7A4A",
  ltgreen: "D4EDDA",
  orange:  "C05C00",
  ltorange:"FFF3CD",
  red:     "A61C00",
  ltred:   "FDECEA",
  gray:    "5A5A5A",
  ltgray:  "F2F4F7",
  midgray: "CCCCCC",
  white:   "FFFFFF",
  darktext:"1A1A1A",
};

// ─── BORDER HELPERS ─────────────────────────────────────────────────────────
const bdr = (color = C.midgray, size = 4) => ({
  top:    { style: BorderStyle.SINGLE, size, color },
  bottom: { style: BorderStyle.SINGLE, size, color },
  left:   { style: BorderStyle.SINGLE, size, color },
  right:  { style: BorderStyle.SINGLE, size, color },
});
const noBorder = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

// ─── PARAGRAPH HELPERS ──────────────────────────────────────────────────────
const sp = (before = 0, after = 0, line) =>
  line ? { before, after, line, lineRule: "exact" } : { before, after };

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, bold: true, size: 36, color: C.navy, font: "Arial" })],
  spacing: sp(320, 160),
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.blue, space: 4 } },
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, bold: true, size: 28, color: C.blue, font: "Arial" })],
  spacing: sp(280, 120),
});
const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, bold: true, size: 24, color: C.navy, font: "Arial" })],
  spacing: sp(200, 80),
});
const body = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, size: 22, font: "Arial", color: C.darktext, ...opts })],
  spacing: sp(80, 80),
});
const bold = (text) => new TextRun({ text, bold: true, size: 22, font: "Arial", color: C.darktext });
const code = (text) => new TextRun({ text, font: "Courier New", size: 20, color: C.navy });
const gap = (sz = 160) => new Paragraph({ children: [new TextRun("")], spacing: sp(0, sz) });

// ─── BULLET HELPER ──────────────────────────────────────────────────────────
const bullet = (text, indent = 0, opts = {}) => new Paragraph({
  numbering: { reference: "bullets", level: indent },
  children: [new TextRun({ text, size: 22, font: "Arial", color: C.darktext, ...opts })],
  spacing: sp(40, 40),
});
const subbullet = (text) => bullet(text, 1);

// ─── CELL HELPER ────────────────────────────────────────────────────────────
const cell = (text, w, { fill, bold: b, color, align, isHeader } = {}) => new TableCell({
  borders: bdr(C.midgray, 4),
  width: { size: w, type: WidthType.DXA },
  shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  children: [new Paragraph({
    alignment: align || AlignmentType.LEFT,
    children: [new TextRun({
      text: String(text),
      size: isHeader ? 20 : 20,
      bold: b || isHeader,
      font: "Arial",
      color: color || (isHeader ? C.white : C.darktext),
    })],
    spacing: sp(40, 40),
  })],
});
const hdrRow = (cells, widths, fills) => new TableRow({
  tableHeader: true,
  children: cells.map((c, i) => cell(c, widths[i], { fill: fills ? fills[i] : C.navy, isHeader: true })),
});
const dataRow = (cells, widths, fills) => new TableRow({
  children: cells.map((c, i) => cell(c, widths[i], fills ? { fill: fills[i] } : {})),
});

// ─── CALLOUT BOX (simulated via single-cell table) ───────────────────────────
const callout = (label, lines, fillColor, labelColor) => {
  const children = [
    new Paragraph({
      children: [new TextRun({ text: label, bold: true, size: 22, font: "Arial", color: labelColor || C.navy })],
      spacing: sp(0, 60),
    }),
    ...lines.map(l => new Paragraph({
      children: [new TextRun({ text: l, size: 20, font: "Arial", color: C.darktext })],
      spacing: sp(20, 20),
    })),
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: bdr(labelColor || C.blue, 6),
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children,
      })],
    })],
  });
};

// ─── COVER PAGE ─────────────────────────────────────────────────────────────
const coverSection = {
  properties: {
    page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
  },
  children: [
    gap(1400),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "PRICEWISE", bold: true, size: 72, font: "Arial", color: C.navy })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Product Price Comparison Platform", size: 36, font: "Arial", color: C.blue })],
      spacing: sp(80, 80),
    }),
    gap(80),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.blue } },
      children: [new TextRun("")],
    }),
    gap(120),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "COMPLETE IMPLEMENTATION PLAN", bold: true, size: 30, font: "Arial", color: C.gray })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Architecture · Data Sources · Scripts · Technology Stack", size: 22, font: "Arial", color: C.gray })],
      spacing: sp(60, 60),
    }),
    gap(1000),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Version 1.0  ·  May 2026  ·  Confidential — Internal Use Only", size: 20, font: "Arial", color: C.gray })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ]
};

// ─── MAIN CONTENT ────────────────────────────────────────────────────────────
const mainSection = {
  properties: {
    page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } }
  },
  headers: {
    default: new Header({
      children: [new Paragraph({
        children: [
          new TextRun({ text: "PriceWise — Complete Implementation Plan", size: 18, font: "Arial", color: C.gray }),
          new TextRun({ text: "    |    Confidential", size: 18, font: "Arial", color: C.midgray }),
        ],
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.midgray } },
        spacing: sp(0, 80),
      })],
    }),
  },
  footers: {
    default: new Footer({
      children: [new Paragraph({
        children: [
          new TextRun({ text: "PriceWise  ·  Page ", size: 18, font: "Arial", color: C.gray }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: C.gray }),
          new TextRun({ text: " of ", size: 18, font: "Arial", color: C.gray }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: C.gray }),
        ],
        alignment: AlignmentType.RIGHT,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.midgray } },
        spacing: sp(80, 0),
      })],
    }),
  },
  children: [

    // ══════════════════════════════════════════════════════
    // SECTION 1 — EXECUTIVE SUMMARY
    // ══════════════════════════════════════════════════════
    h1("1. Executive Summary"),
    body("PriceWise is a real-time product price comparison platform targeting Indian e-commerce consumers. It aggregates live pricing, discount, and quality data from 8+ major platforms including Amazon India, Flipkart, Myntra, Meesho, Snapdeal, Nykaa, Croma, and Reliance Digital, and surfaces an intelligent composite Value Score to help shoppers make the best purchasing decision in seconds."),
    gap(60),
    body("This document is the complete implementation blueprint — covering every layer of the system: where to get data, how to collect it, what scripts to write, which technologies to use, and how all pieces fit together. It is structured for a team of 2–4 engineers to execute across 9 months."),
    gap(120),

    callout("🎯  What PriceWise Does", [
      "→  User searches for any product (e.g. 'Samsung Galaxy S24')",
      "→  PriceWise simultaneously queries Amazon, Flipkart, Myntra, Meesho, Snapdeal, Nykaa, Croma, Reliance Digital",
      "→  System normalizes prices, parses all active offers & discounts, aggregates ratings",
      "→  Proprietary Value Score (0–100) ranks platforms from best to worst deal",
      "→  User sees one unified table with Buy Now deep-links to each platform",
    ], C.ltblue, C.navy),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 2 — HOW IT WORKS (END TO END)
    // ══════════════════════════════════════════════════════
    h1("2. How It Works — End-to-End Flow"),
    body("The following is the complete journey of a single user search request through every layer of the PriceWise system:"),
    gap(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [800, 2200, 6360],
      rows: [
        hdrRow(["Step", "Layer", "What Happens"], [800, 2200, 6360]),
        dataRow(["1", "Frontend (Next.js)", "User types product name in search bar → React sends GET /api/v1/search?q=... to API Gateway"], [800, 2200, 6360], [C.ltgray, C.ltgray, C.white]),
        dataRow(["2", "API Gateway", "Node.js/Express receives request → validates JWT → checks Redis cache (30-min TTL)"], [800, 2200, 6360], [C.white, C.white, C.ltgray]),
        dataRow(["3", "Cache Hit", "Redis returns pre-computed comparison data → API sends directly to frontend (<50ms)"], [800, 2200, 6360], [C.ltgray, C.ltgray, C.white]),
        dataRow(["4", "Cache Miss", "Elasticsearch searched for product entity → on-demand scraping triggered across all 8 platforms in parallel"], [800, 2200, 6360], [C.white, C.white, C.ltgray]),
        dataRow(["5", "Scraper Workers", "Scrapy/Playwright spiders fetch live product pages → extract price, MRP, offers, rating, review count, delivery time"], [800, 2200, 6360], [C.ltgray, C.ltgray, C.white]),
        dataRow(["6", "Processing Pipeline", "Price Normalizer → Offer Parser → Product Matcher → Quality Aggregator → Value Score Calculator"], [800, 2200, 6360], [C.white, C.white, C.ltgray]),
        dataRow(["7", "Storage", "Processed data written to PostgreSQL + cached in Redis + index updated in Elasticsearch"], [800, 2200, 6360], [C.ltgray, C.ltgray, C.white]),
        dataRow(["8", "API Response", "Ranked JSON payload returned: platform listings sorted by Value Score + price history + recommendation badge"], [800, 2200, 6360], [C.white, C.white, C.ltgray]),
        dataRow(["9", "Frontend Render", "React renders comparison table, Recharts price history chart, Value Score badges, Buy Now deep-links"], [800, 2200, 6360], [C.ltgray, C.ltgray, C.white]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 3 — DATA SOURCES
    // ══════════════════════════════════════════════════════
    h1("3. Data Sources — Where to Get Every Data Point"),
    body("PriceWise relies on two types of data sources: official APIs (where available) and web scraping (for platforms without APIs). Below is the complete breakdown per platform."),
    gap(120),

    h2("3.1  Amazon India — Official API + Scraping"),
    callout("Amazon Product Advertising API (PA API 5.0)", [
      "URL: https://webservices.amazon.in/paapi5/",
      "Registration: affiliate-program.amazon.in → apply for Associates account → get Access Key + Secret Key",
      "What it gives you: Product title, ASIN, price, MRP, discount %, Prime badge, rating, review count, images, category, availability, Buy Now affiliate deep-link",
      "Rate Limit: 1 request/second on new accounts; scales with affiliate sales volume",
      "Cost: Free (earns affiliate commission on referred purchases)",
      "Endpoint: SearchItems (search by keyword), GetItems (fetch by ASIN), GetVariations",
    ], C.ltgreen, C.green),
    gap(80),
    body("When to scrape Amazon: Offers, bank EMI deals, and coupons are NOT in PA API — scrape the product page HTML for these using Playwright (they are rendered client-side by React on Amazon's pages)."),
    gap(120),

    h2("3.2  Flipkart — Partial API + Scraping"),
    callout("Flipkart Affiliate API", [
      "URL: https://affiliate.flipkart.com/static/docs/",
      "Registration: affiliate.flipkart.com → register → get Tracking ID + App ID + Token",
      "What it gives you: Product title, price, MRP, discount %, category feed in JSON/XML, product URL",
      "Limitation: Gives batch catalog feeds, NOT real-time prices — must supplement with scraping",
      "Feed refresh: Every 2–4 hours",
      "Scraping needed for: Live price, active coupons, bank offers, rating, review count, delivery SLA",
    ], C.ltgreen, C.green),
    gap(80),
    body("Flipkart's product pages are server-side rendered — use Scrapy (requests + BeautifulSoup) first; fall back to Playwright if content is behind JS lazy load."),
    gap(120),

    h2("3.3  Myntra — Full Scraping"),
    callout("Myntra (No Official API)", [
      "Method: Web scraping via Playwright (site uses heavy client-side React rendering)",
      "Search URL pattern: https://www.myntra.com/{product-slug}?rawQuery={keyword}",
      "Product data embedded in: window.__myx object in page source (parse JSON directly — faster than DOM scraping)",
      "Key fields to extract: brandName, name, mrp, price, discountDisplayLabel, ratings, ratingCount, sizes",
      "Anti-bot: Requires cookie rotation + user-agent spoofing; sessions expire; use Playwright stealth plugin",
      "Rate limit: ~10–15 req/min per IP; rotate via residential proxies",
    ], C.ltorange, C.orange),
    gap(120),

    h2("3.4  Meesho — Full Scraping"),
    callout("Meesho (No Official API)", [
      "Method: Playwright scraping (heavy React SPA, no SSR)",
      "Search URL: https://www.meesho.com/search?q={keyword}",
      "Product data: Embedded in __NEXT_DATA__ JSON in page source — parse with Python json.loads()",
      "Key fields: name, price, mrp, discountPercent, ratings, ratingCount, supplierName",
      "Note: Meesho prices vary by pin code (logistics zone) — always set a reference pincode in request headers",
      "Headers needed: X-Latitude, X-Longitude or pincode cookie to get consistent pricing",
    ], C.ltorange, C.orange),
    gap(120),

    h2("3.5  Snapdeal, Nykaa, Croma, Reliance Digital — Full Scraping"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1800, 2500, 2300, 2760],
      rows: [
        hdrRow(["Platform", "Search URL Pattern", "Rendering", "Data Location"], [1800, 2500, 2300, 2760]),
        dataRow(["Snapdeal", "snapdeal.com/search#q={keyword}", "Server-side", "HTML DOM — .product-price, .product-desc-rating"], [1800, 2500, 2300, 2760], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Nykaa", "nykaa.com/search/result/?q={keyword}", "Client-side (React)", "__NEXT_DATA__ JSON in page source"], [1800, 2500, 2300, 2760], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Croma", "croma.com/searchB?q={keyword}", "Client-side (React)", "window.__INITIAL_STATE__ JSON in page source"], [1800, 2500, 2300, 2760], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Reliance Digital", "reliancedigital.in/search?q={keyword}", "Server-side", "HTML DOM — .pdp-price, .customer-reviews"], [1800, 2500, 2300, 2760], [C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 4 — SCRIPTS TO WRITE
    // ══════════════════════════════════════════════════════
    h1("4. Scripts to Write — Complete List with Purpose"),
    body("Every script in the PriceWise system is listed below, organized by layer. File paths follow the project repository structure."),
    gap(120),

    h2("4.1  Scraper Scripts (Python)"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 2400, 3760],
      rows: [
        hdrRow(["Script File", "Platform / Purpose", "What It Does"], [3200, 2400, 3760]),
        dataRow(["scrapers/amazon_api.py", "Amazon PA API", "Calls SearchItems & GetItems endpoints; normalizes JSON response to PriceWise schema; handles pagination"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/amazon_offers.py", "Amazon — Offers Only", "Playwright script: navigates to product page, extracts bank offers, coupon codes, EMI options from offer modal"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
        dataRow(["scrapers/flipkart_feed.py", "Flipkart Affiliate Feed", "Downloads & parses Flipkart XML/JSON catalog feed; caches to disk; extracts base price + URL"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/flipkart_live.py", "Flipkart — Live Price", "Scrapy spider: fetches product page, extracts live price, coupons, ratings via CSS selectors"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
        dataRow(["scrapers/myntra.py", "Myntra", "Playwright: loads page, reads window.__myx JSON, extracts all product fields; handles pagination"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/meesho.py", "Meesho", "Playwright: intercepts __NEXT_DATA__ from page source; sets pincode header; extracts price, rating"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
        dataRow(["scrapers/snapdeal.py", "Snapdeal", "Scrapy spider using BeautifulSoup; CSS selectors for price, rating, stock; handles JS-free pages"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/nykaa.py", "Nykaa", "Playwright: reads __NEXT_DATA__ JSON; extracts beauty product price, brand, variants, ratings"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
        dataRow(["scrapers/croma.py", "Croma", "Playwright: reads window.__INITIAL_STATE__; extracts electronics price, specs, delivery SLA"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/reliance_digital.py", "Reliance Digital", "Scrapy spider: BeautifulSoup on product listing pages; extracts price, rating, pincode-based availability"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
        dataRow(["scrapers/proxy_manager.py", "Proxy Rotation", "Manages residential proxy pool (BrightData or Oxylabs API); rotates IP per request; tracks ban status"], [3200, 2400, 3760], [C.ltgray, C.white, C.ltgray]),
        dataRow(["scrapers/stealth_browser.py", "Anti-Detection", "Configures Playwright with stealth plugin, random user agents, viewport randomization, cookie injection"], [3200, 2400, 3760], [C.white, C.ltgray, C.white]),
      ],
    }),
    gap(120),

    h2("4.2  Data Processing Scripts (Python)"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 6160],
      rows: [
        hdrRow(["Script File", "What It Does"], [3200, 6160]),
        dataRow(["processing/normalizer.py", "Converts all prices to INR (handles USD, GBP via forex API); strips ₹ symbols, commas, formatting; validates price sanity (not 0, not 999999)"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["processing/offer_parser.py", "Regex + ML classifier to extract structured offers from unstructured promo text. E.g. '10% off on HDFC cards above ₹5000' → {type: bank_offer, bank: HDFC, discount: 10%, min_order: 5000}"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["processing/product_matcher.py", "NLP entity resolution: TF-IDF vectorizer on product titles → cosine similarity matrix → clusters listings that refer to same product. Uses spaCy NER to extract brand, model, specs"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["processing/quality_aggregator.py", "Weighted average of ratings across platforms (weight = log(review_count+1)); statistical outlier detection for fake reviews (Z-score on rating distribution)"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["processing/value_score.py", "Computes 0–100 Value Score per platform listing. Inputs: effective_price, offer_value, weighted_rating, log_review_count, delivery_days. Weights configurable per mode (Price-First / Quality-First / Balanced)"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["processing/effective_price.py", "Computes final payable price: MRP → apply % discount → apply coupon → apply cashback → apply bank offer (cascading). Returns both effective_price and breakdown dict"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["processing/price_history.py", "Reads historical pricing from PostgreSQL; generates time-series arrays for 30/90/180-day windows; computes min/max/avg/current; detects price drop events"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["processing/alert_engine.py", "Cron job: compares current prices against user watchlist thresholds; triggers email (SendGrid) and push notifications (Firebase FCM) when price drops below target"], [3200, 6160], [C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("4.3  Backend API Scripts (Node.js / TypeScript)"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 6160],
      rows: [
        hdrRow(["File", "What It Does"], [3200, 6160]),
        dataRow(["api/routes/search.ts", "GET /api/v1/search?q={query} — validates query, checks Redis cache, triggers Elasticsearch lookup, calls scraper workers if needed, returns ranked results"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["api/routes/compare.ts", "GET /api/v1/compare/{productId} — fetches full platform comparison data from PostgreSQL, computes/fetches Value Scores, returns complete comparison payload"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["api/routes/priceHistory.ts", "GET /api/v1/price-history/{productId}?days=90 — queries PostgreSQL price_history table, returns time-series JSON for Recharts chart"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["api/routes/watchlist.ts", "POST/DELETE /api/v1/watchlist — authenticated endpoints to add/remove products from user watchlist; stores target price threshold; validates JWT"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["api/routes/auth.ts", "POST /api/v1/auth/register & /login — bcrypt password hashing, JWT signing (RS256), Google OAuth 2.0 via Passport.js, refresh token rotation"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["api/middleware/cache.ts", "Redis middleware: checks cache before DB hit; sets TTL (30 min for search results, 1 hour for price history); cache invalidation on price update"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["api/middleware/rateLimiter.ts", "Express-rate-limit: 100 req/min per IP on search, 20 req/min on scrape-trigger endpoints; Cloudflare WAF as outer layer"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["api/graphql/schema.ts", "GraphQL schema (Apollo Server): Product, Platform, PriceHistory, ValueScore, User, Watchlist types; resolvers delegating to same service layer as REST"], [3200, 6160], [C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("4.4  Database Scripts (SQL + Migrations)"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 6160],
      rows: [
        hdrRow(["File", "What It Does"], [3200, 6160]),
        dataRow(["db/migrations/001_schema.sql", "Creates core tables: products, platform_listings, price_history, offers, users, watchlists, alerts. Defines indexes on product_id, platform, scraped_at"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["db/migrations/002_indexes.sql", "Partial indexes for active listings only; BRIN index on scraped_at (time-series); GIN index on product title (full-text search backup)"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["db/seeds/platforms.sql", "Inserts platform metadata: name, base URL, scraper config (user-agent, rate limit, selector version), is_active flag"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["db/queries/comparison.sql", "Complex JOIN query fetching latest listing per platform per product with computed effective_price and value_score — used by compare.ts route"], [3200, 6160], [C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("4.5  Scheduling Scripts (Apache Airflow DAGs)"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 6160],
      rows: [
        hdrRow(["DAG File", "Schedule & Purpose"], [3200, 6160]),
        dataRow(["dags/scrape_amazon.py", "Every 1 hour — Amazon PA API refresh for top 10,000 tracked products; parallel ASIN batches of 10"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["dags/scrape_flipkart.py", "Every 2 hours — Flipkart feed download + live scraping for cached product URLs"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["dags/scrape_others.py", "Every 4 hours — Myntra, Meesho, Snapdeal, Nykaa, Croma, Reliance Digital scrapers (rotating, staggered)"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["dags/price_alerts.py", "Every 30 minutes — Alert engine checks watchlists against current prices; sends notifications"], [3200, 6160], [C.white, C.ltgray]),
        dataRow(["dags/reindex_search.py", "Every 6 hours — Syncs PostgreSQL product catalog to Elasticsearch index; handles incremental updates"], [3200, 6160], [C.ltgray, C.white]),
        dataRow(["dags/retrain_matcher.py", "Weekly — Retrains spaCy NER model and TF-IDF product matcher on newly accumulated product title data"], [3200, 6160], [C.white, C.ltgray]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 5 — TECHNOLOGY STACK
    // ══════════════════════════════════════════════════════
    h1("5. Complete Technology Stack"),
    body("Every technology choice is listed below with the exact version, why it was chosen over alternatives, and where to get it."),
    gap(120),

    h2("5.1  Frontend Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["Next.js", "14.x", "React framework with SSR for SEO-friendly product pages and fast initial load. Chosen over CRA for SSR & image optimization.", "nextjs.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["React", "18.x", "UI component library. Concurrent mode for responsive comparisons. Industry standard.", "react.dev", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["TypeScript", "5.x", "Static typing reduces runtime bugs in complex data transformations. Mandatory for team-scale projects.", "typescriptlang.org", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Tailwind CSS", "3.x", "Utility-first CSS. Mobile-first responsive design. No CSS file bloat. Chosen over MUI for full customizability.", "tailwindcss.com", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Redux Toolkit", "2.x", "Global state for user preferences, scoring mode, watchlist. RTK Query for server state. Replaces Context + useReducer.", "redux-toolkit.js.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["React Query (TanStack)", "5.x", "Data fetching, caching, background refresh for comparison results. Handles loading/error states elegantly.", "tanstack.com/query", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Recharts", "2.x", "SVG-based price history charts, Value Score radars. Pure React, no D3 dependency overhead.", "recharts.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Framer Motion", "11.x", "Page transitions, comparison card animations. Makes UI feel polished and fast.", "framer.com/motion", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
      ],
    }),
    gap(120),

    h2("5.2  Backend Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["Node.js", "20 LTS", "API Gateway runtime. High concurrency for parallel scraper result aggregation. V8 engine, mature ecosystem.", "nodejs.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Express.js", "4.x", "REST API framework. Minimal, flexible, huge middleware ecosystem. Chosen over Fastify for familiarity.", "expressjs.com", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Apollo Server", "4.x", "GraphQL server. Enables flexible client-driven queries for the comparison table. Replaces over-fetching REST.", "apollographql.com", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Passport.js", "0.7.x", "OAuth 2.0 for Google login. Modular strategy system. Industry standard for Node.js auth.", "passportjs.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["python-jose / jsonwebtoken", "latest", "JWT signing & verification (RS256). Asymmetric keys for API gateway / Python service interop.", "npm / PyPI", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("5.3  Scraping & Data Collection Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["Scrapy", "2.11.x", "High-performance async web crawling for SSR platforms (Snapdeal, Reliance Digital). Built-in rate limiting, retry, proxy middleware.", "scrapy.org", "BSD"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Playwright (Python)", "1.44.x", "Headless browser for JS-heavy SPA platforms (Myntra, Meesho, Nykaa, Croma). Chosen over Selenium for speed and async support.", "playwright.dev", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["playwright-stealth", "1.x", "Anti-detection plugin for Playwright: removes automation fingerprints, randomizes WebGL, canvas, AudioContext signatures.", "PyPI", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["BeautifulSoup4", "4.12.x", "HTML parser for SSR page scraping within Scrapy spiders. lxml backend for speed.", "PyPI", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["httpx", "0.27.x", "Async HTTP client for API connectors (Amazon PA API, Flipkart Affiliate). HTTP/2 support.", "python-httpx.org", "BSD"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Bright Data / Oxylabs", "SaaS", "Residential proxy pool for IP rotation. Needed to prevent bans on Flipkart/Myntra. ~$15/GB.", "brightdata.com", "Paid SaaS"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
      ],
    }),
    gap(120),

    h2("5.4  NLP & Machine Learning Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["spaCy", "3.7.x", "NER pipeline for extracting brand names, model numbers, specs from product titles. Pre-trained en_core_web_md model as base. Fine-tune on e-commerce product corpus.", "spacy.io", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["scikit-learn", "1.5.x", "TF-IDF vectorizer + cosine similarity for product matching. LogisticRegression for offer type classifier.", "scikit-learn.org", "BSD"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["sentence-transformers", "3.x", "SBERT embeddings for better product matching beyond TF-IDF. Use paraphrase-multilingual-MiniLM-L12-v2 for Hindi product names.", "sbert.net", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["pandas + numpy", "2.x / 1.x", "Data manipulation for price normalization, statistical outlier detection (Z-score for fake review flagging), price history aggregation.", "pypi.org", "BSD"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Apache Airflow", "2.9.x", "DAG-based job scheduler for orchestrating all scraping and processing tasks with retry, SLA monitoring, alerting.", "airflow.apache.org", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("5.5  Database & Storage Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["PostgreSQL", "16.x", "Primary relational database. Products, listings, price_history, users, watchlists. BRIN indexes for time-series pricing data. Chosen over MySQL for JSONB support & window functions.", "postgresql.org", "PostgreSQL"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Redis", "7.x", "In-memory cache for search results (TTL 30 min), session tokens, rate limit counters. 50–100x faster than DB for hot data.", "redis.io", "BSD"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Elasticsearch", "8.x", "Full-text product search index. Fuzzy matching for typos, synonym handling (e.g. 'mobile' = 'phone'), faceted filtering by category/brand.", "elastic.co", "SSPL"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["AWS S3 / Cloudflare R2", "SaaS", "Object storage for product images fetched from platforms. R2 is chosen for zero egress fees. CloudFront CDN for image delivery.", "aws.amazon.com / cloudflare.com", "Paid SaaS"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Prisma ORM", "5.x", "Type-safe DB client for Node.js API layer. Auto-generates TypeScript types from PostgreSQL schema. Migrations built-in.", "prisma.io", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["SQLAlchemy", "2.x", "ORM for Python scraping & processing services. Async support with asyncpg driver for high-throughput writes.", "sqlalchemy.org", "MIT"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
      ],
    }),
    gap(120),

    h2("5.6  Infrastructure & DevOps Technologies"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 1200, 3160, 2000, 1000],
      rows: [
        hdrRow(["Technology", "Version", "Purpose & Why Chosen", "Get It From", "License"], [2000, 1200, 3160, 2000, 1000]),
        dataRow(["Docker", "26.x", "Containerizes each service (scrapers, API, Airflow, etc.) for consistent environments. docker-compose for local dev.", "docker.com", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Kubernetes (K8s)", "1.30.x", "Orchestrates containers in production. Auto-scales scraper workers on demand. Use AWS EKS or GKE managed K8s.", "kubernetes.io", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["GitHub Actions", "N/A", "CI/CD: lint → test → build Docker image → push to ECR → deploy to K8s. Separate workflows for frontend, API, scrapers.", "github.com", "Free/Paid"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Prometheus + Grafana", "latest", "System health monitoring: scraper success rates, API latency, cache hit ratio, DB query times. Alerts via PagerDuty.", "prometheus.io / grafana.com", "Apache 2"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Cloudflare", "SaaS", "DNS, DDoS protection, WAF, rate limiting at edge. Zero-cost CDN for static assets. Workers for edge caching.", "cloudflare.com", "Free/Paid"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Sentry", "SaaS", "Error tracking for frontend (React) and backend (Node, Python). Source maps for minified JS. Slack alerts on new errors.", "sentry.io", "Free/Paid"], [2000, 1200, 3160, 2000, 1000], [C.white, C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["AWS Secrets Manager", "SaaS", "Stores API keys, DB credentials, proxy credentials. Rotated automatically. No secrets in env files or git.", "aws.amazon.com", "Paid SaaS"], [2000, 1200, 3160, 2000, 1000], [C.ltgray, C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(120),

    h2("5.7  Third-Party APIs & Services"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 3960, 1800, 1200],
      rows: [
        hdrRow(["Service", "Purpose", "Where to Get", "Cost"], [2400, 3960, 1800, 1200]),
        dataRow(["Amazon PA API 5.0", "Official Amazon product data: price, ASIN, rating, images, affiliate deep-links", "affiliate-program.amazon.in", "Free (affiliate)"], [2400, 3960, 1800, 1200], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Flipkart Affiliate API", "Catalog feed with prices and product URLs", "affiliate.flipkart.com", "Free (affiliate)"], [2400, 3960, 1800, 1200], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["ExchangeRate-API", "INR currency conversion for any non-INR prices", "exchangerate-api.com", "Free tier (1500 req/mo)"], [2400, 3960, 1800, 1200], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["SendGrid", "Transactional email for price drop alerts, registration, OTP", "sendgrid.com", "Free (100 emails/day)"], [2400, 3960, 1800, 1200], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Firebase FCM", "Browser & Android push notifications for price alerts", "firebase.google.com", "Free"], [2400, 3960, 1800, 1200], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Google OAuth 2.0", "Social login (Sign in with Google)", "console.cloud.google.com", "Free"], [2400, 3960, 1800, 1200], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Snyk", "Dependency vulnerability scanning in CI pipeline", "snyk.io", "Free tier available"], [2400, 3960, 1800, 1200], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Bright Data", "Residential proxy pool for anti-ban scraping", "brightdata.com", "~$15/GB"], [2400, 3960, 1800, 1200], [C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 6 — DATABASE SCHEMA
    // ══════════════════════════════════════════════════════
    h1("6. Database Schema (PostgreSQL)"),
    body("The following are the core table definitions for PriceWise's PostgreSQL database. All tables use UUIDs as primary keys for horizontal scaling compatibility."),
    gap(120),

    h2("6.1  Core Tables"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1600, 1600, 1200, 4960],
      rows: [
        hdrRow(["Table", "Column", "Type", "Description"], [1600, 1600, 1200, 4960]),
        dataRow(["products", "id", "UUID PK", "Unique product entity (after matching across platforms)"], [1600, 1600, 1200, 4960], [C.ltblue, C.white, C.ltgray, C.white]),
        dataRow(["", "canonical_name", "TEXT", "Normalized product name (e.g. 'Samsung Galaxy S24 256GB Black')"], [1600, 1600, 1200, 4960], [C.ltblue, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "brand", "TEXT", "Extracted brand name (spaCy NER)"], [1600, 1600, 1200, 4960], [C.ltblue, C.white, C.ltgray, C.white]),
        dataRow(["", "category", "TEXT", "Top-level category (Electronics, Fashion, etc.)"], [1600, 1600, 1200, 4960], [C.ltblue, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "embedding", "VECTOR(384)", "SBERT embedding for similarity search (pgvector extension)"], [1600, 1600, 1200, 4960], [C.ltblue, C.white, C.ltgray, C.white]),
        dataRow(["platform_listings", "id", "UUID PK", "One row per product per platform per scrape run"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "product_id", "UUID FK", "References products.id"], [1600, 1600, 1200, 4960], [C.ltorange, C.white, C.ltgray, C.white]),
        dataRow(["", "platform", "TEXT", "amazon | flipkart | myntra | meesho | snapdeal | nykaa | croma | reliancedigital"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "mrp", "NUMERIC(12,2)", "Maximum Retail Price (before discount)"], [1600, 1600, 1200, 4960], [C.ltorange, C.white, C.ltgray, C.white]),
        dataRow(["", "listed_price", "NUMERIC(12,2)", "Platform's listed/selling price"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "effective_price", "NUMERIC(12,2)", "Final price after all offers (computed by effective_price.py)"], [1600, 1600, 1200, 4960], [C.ltorange, C.white, C.ltgray, C.white]),
        dataRow(["", "rating", "NUMERIC(3,2)", "Weighted average star rating"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "review_count", "INTEGER", "Number of reviews (log-scaled in Value Score)"], [1600, 1600, 1200, 4960], [C.ltorange, C.white, C.ltgray, C.white]),
        dataRow(["", "delivery_days", "INTEGER", "Estimated delivery SLA in days"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "product_url", "TEXT", "Deep-link to platform product page (Buy Now)"], [1600, 1600, 1200, 4960], [C.ltorange, C.white, C.ltgray, C.white]),
        dataRow(["", "scraped_at", "TIMESTAMPTZ", "When this row was scraped (BRIN indexed)"], [1600, 1600, 1200, 4960], [C.ltorange, C.ltgray, C.white, C.ltgray]),
        dataRow(["price_history", "id", "BIGSERIAL PK", "Time-series table (one row per scrape per platform per product)"], [1600, 1600, 1200, 4960], [C.ltgreen, C.white, C.ltgray, C.white]),
        dataRow(["", "listing_id", "UUID FK", "References platform_listings.id"], [1600, 1600, 1200, 4960], [C.ltgreen, C.ltgray, C.white, C.ltgray]),
        dataRow(["", "effective_price", "NUMERIC(12,2)", "Effective price at this point in time"], [1600, 1600, 1200, 4960], [C.ltgreen, C.white, C.ltgray, C.white]),
        dataRow(["", "recorded_at", "TIMESTAMPTZ", "Timestamp of price recording"], [1600, 1600, 1200, 4960], [C.ltgreen, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 7 — VALUE SCORE ALGORITHM DETAIL
    // ══════════════════════════════════════════════════════
    h1("7. Value Score Algorithm — Detailed Implementation"),
    body("The Value Score is PriceWise's core intellectual property. It converts multi-dimensional product data into a single 0–100 score per platform listing, allowing objective comparison."),
    gap(120),

    h2("7.1  Formula"),
    callout("Value Score Formula", [
      "ValueScore = Σ (component_score_i × weight_i)  where each component is normalized to 0–100",
      "",
      "Price Score      = 100 × (max_price - effective_price) / (max_price - min_price)",
      "Offer Score      = min(100, total_offer_value_percent × 2)   [capped at 100]",
      "Quality Score    = (weighted_avg_rating - 1) / 4 × 100       [maps 1–5 stars to 0–100]",
      "Review Score     = min(100, log10(review_count + 1) / log10(10001) × 100)  [log-scaled, 10K+ reviews = 100]",
      "Delivery Score   = 100 × (max_days - delivery_days) / (max_days - min_days)",
    ], C.ltblue, C.navy),
    gap(120),

    h2("7.2  Weight Distribution by Mode"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2160, 2400, 2400, 2400],
      rows: [
        hdrRow(["Component", "Balanced Mode", "Price-First Mode", "Quality-First Mode"], [2160, 2400, 2400, 2400]),
        dataRow(["Price Score", "40%", "65%", "20%"], [2160, 2400, 2400, 2400], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Offer Score", "20%", "10%", "10%"], [2160, 2400, 2400, 2400], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Quality Score", "25%", "15%", "60%"], [2160, 2400, 2400, 2400], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Review Score", "10%", "5%", "7%"], [2160, 2400, 2400, 2400], [C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["Delivery Score", "5%", "5%", "3%"], [2160, 2400, 2400, 2400], [C.ltgray, C.white, C.ltgray, C.white]),
        dataRow(["Total", "100%", "100%", "100%"], [2160, 2400, 2400, 2400], [C.ltblue, C.ltblue, C.ltblue, C.ltblue]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 8 — RISKS & MITIGATIONS (DETAILED)
    // ══════════════════════════════════════════════════════
    h1("8. Risks & Detailed Mitigations"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 1000, 5960],
      rows: [
        hdrRow(["Risk", "Severity", "Detailed Mitigation Plan"], [2400, 1000, 5960]),
        dataRow(["Platform bans / IP blocks", "HIGH", "Layer 1: Bright Data residential proxies (rotate per request). Layer 2: playwright-stealth (remove automation fingerprints). Layer 3: Human-like delays (random 2–8 sec between requests). Layer 4: Session cookies + realistic browser headers. Layer 5: CAPTCHA solver integration (2captcha / Anti-captcha API) as last resort"], [2400, 1000, 5960], [C.ltred, C.ltred, C.white]),
        dataRow(["Amazon PA API rate limits", "MEDIUM", "Implement exponential backoff with jitter (start 1s, max 60s). Batch ASIN lookups (max 10 per request). Aggressive Redis caching (1-hour TTL for PA API data). Request queue with Celery for smoothing burst traffic. Multiple affiliate accounts for higher rate limits"], [2400, 1000, 5960], [C.ltorange, C.ltorange, C.white]),
        dataRow(["Product matching errors", "MEDIUM", "Three-stage matching: TF-IDF cosine (fast, 90% accuracy) → SBERT embedding (slower, 96% accuracy) → human review queue for low-confidence matches (<0.75 score). Build labeled dataset of matched pairs over time for continuous ML retraining (weekly Airflow DAG)"], [2400, 1000, 5960], [C.ltorange, C.ltorange, C.white]),
        dataRow(["Scraped HTML structure changes", "MEDIUM", "Version all CSS selectors in platform config JSON. Monitor scraper success rate via Prometheus — alert at <95%. Auto-detect structural changes by comparing field null rates. Telegram bot notifies dev team within 5 min of selector failure"], [2400, 1000, 5960], [C.ltorange, C.ltorange, C.white]),
        dataRow(["Legal / ToS scraping issues", "MEDIUM", "Prefer official APIs wherever available (Amazon PA API, Flipkart Affiliate). Respect robots.txt for background crawls. Do not scrape personal user data. Consult IP lawyer before launch. Implement opt-out if platforms send legal notice. Store scraped data only for comparison, not redistribution"], [2400, 1000, 5960], [C.ltorange, C.ltorange, C.white]),
        dataRow(["Price data staleness", "LOW-MED", "Short TTL (30 min) on Redis cache. On-demand scrape refresh triggered on each search (cache miss). Mark data with scraped_at timestamp; show 'Last updated X min ago' in UI. Airflow background refresh every 1–6 hours depending on platform volatility"], [2400, 1000, 5960], [C.ltgreen, C.ltgreen, C.white]),
        dataRow(["Database performance under load", "LOW", "PostgreSQL read replicas for search queries. Connection pooling via PgBouncer (transaction mode, 200 connections). BRIN index on price_history.recorded_at (time-series). Partition price_history by month. Elasticsearch absorbs full-text search load. Redis absorbs 80%+ of read traffic"], [2400, 1000, 5960], [C.ltgreen, C.ltgreen, C.white]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 9 — DEVELOPMENT ROADMAP
    // ══════════════════════════════════════════════════════
    h1("9. Development Roadmap — Month by Month"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [900, 1200, 4360, 1500, 1400],
      rows: [
        hdrRow(["Phase", "Timeline", "Deliverables", "Team Focus", "Status"], [900, 1200, 4360, 1500, 1400]),
        dataRow(["1", "Month 1–2", "Core architecture setup, PostgreSQL schema, Redis + Elasticsearch config, Amazon PA API connector, Flipkart scraper, basic search API, Next.js homepage + search bar", "Backend + DevOps", "In Progress"], [900, 1200, 4360, 1500, 1400], [C.ltblue, C.white, C.ltgray, C.white, C.ltorange]),
        dataRow(["2", "Month 3–4", "5 more platform scrapers (Myntra, Meesho, Snapdeal, Nykaa, Croma), product matching NLP engine, Value Score algorithm, offer parser, full comparison table UI with sorting", "Full-stack + ML", "Planned"], [900, 1200, 4360, 1500, 1400], [C.ltblue, C.ltgray, C.white, C.ltgray, C.ltgray]),
        dataRow(["3", "Month 5–6", "User auth (JWT + Google OAuth), watchlists, price drop alerts (SendGrid + FCM), price history charts (Recharts), Reliance Digital scraper, 30/90/180-day history UI", "Backend + Frontend", "Planned"], [900, 1200, 4360, 1500, 1400], [C.ltblue, C.white, C.ltgray, C.white, C.ltgray]),
        dataRow(["4", "Month 7–8", "Mobile PWA (offline support, install prompt), dark mode, personalization engine, A/B testing (Value Score modes), performance optimization (Lighthouse ≥ 90)", "Frontend + QA", "Planned"], [900, 1200, 4360, 1500, 1400], [C.ltblue, C.ltgray, C.white, C.ltgray, C.ltgray]),
        dataRow(["5", "Month 9+", "ML-based deal prediction (price drop forecasting), Chrome extension (compare while browsing), affiliate revenue integration, international platform expansion", "ML + Growth", "Future"], [900, 1200, 4360, 1500, 1400], [C.ltblue, C.white, C.ltgray, C.white, C.ltgray]),
      ],
    }),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 10 — PROJECT STRUCTURE
    // ══════════════════════════════════════════════════════
    h1("10. Repository Structure"),
    callout("Recommended Monorepo Layout (pricewise/)", [
      "├── apps/",
      "│   ├── web/                    ← Next.js frontend (TypeScript + Tailwind)",
      "│   └── api/                    ← Node.js Express + GraphQL backend",
      "├── services/",
      "│   ├── scrapers/               ← Python Scrapy + Playwright spiders (one file per platform)",
      "│   ├── processing/             ← Python data pipeline (normalizer, matcher, value_score, etc.)",
      "│   └── scheduler/              ← Apache Airflow DAGs",
      "├── db/",
      "│   ├── migrations/             ← SQL migration files (001_schema.sql, 002_indexes.sql, ...)",
      "│   ├── seeds/                  ← Seed data (platforms, categories)",
      "│   └── queries/                ← Reusable SQL queries",
      "├── infra/",
      "│   ├── docker/                 ← Dockerfiles per service",
      "│   ├── k8s/                    ← Kubernetes manifests (deployments, services, ingress)",
      "│   └── terraform/              ← AWS infrastructure as code",
      "├── .github/workflows/          ← GitHub Actions CI/CD pipelines",
      "├── docker-compose.yml          ← Local development environment",
      "└── README.md",
    ], C.ltgray, C.navy),
    gap(200),

    // ══════════════════════════════════════════════════════
    // SECTION 11 — GETTING STARTED
    // ══════════════════════════════════════════════════════
    h1("11. Getting Started — Setup Sequence"),
    body("Follow this exact sequence to set up a working local development environment from scratch:"),
    gap(80),
    h2("Step 1 — Prerequisites"),
    bullet("Install Node.js 20 LTS from nodejs.org"),
    bullet("Install Python 3.11 from python.org"),
    bullet("Install Docker Desktop from docker.com"),
    bullet("Install git from git-scm.com"),
    gap(80),
    h2("Step 2 — External Account Registrations (Do First)"),
    bullet("Amazon Associates: affiliate-program.amazon.in → Get Access Key + Secret Key + Partner Tag"),
    bullet("Flipkart Affiliate: affiliate.flipkart.com → Get App ID + Token"),
    bullet("Google Cloud Console: console.cloud.google.com → Enable OAuth 2.0 → Get Client ID + Secret"),
    bullet("SendGrid: sendgrid.com → Get API Key (free tier, 100 emails/day)"),
    bullet("Firebase: firebase.google.com → Create project → Enable FCM → Get Server Key"),
    bullet("Bright Data: brightdata.com → Create Residential proxy zone → Get endpoint + credentials"),
    gap(80),
    h2("Step 3 — Local Environment"),
    bullet("Clone repository: git clone https://github.com/your-org/pricewise"),
    bullet("Copy .env.example to .env and fill in all API keys from Step 2"),
    bullet("Run: docker-compose up -d (starts PostgreSQL, Redis, Elasticsearch, Airflow)"),
    bullet("Run DB migrations: npm run db:migrate (from apps/api/)"),
    bullet("Install Python deps: pip install -r services/scrapers/requirements.txt"),
    bullet("Install Playwright browsers: playwright install chromium"),
    bullet("Run API: npm run dev (from apps/api/) → http://localhost:3001"),
    bullet("Run frontend: npm run dev (from apps/web/) → http://localhost:3000"),
    bullet("Run Airflow UI: http://localhost:8080 (admin/admin) → Enable scraping DAGs"),
    gap(200),

    // ══════════════════════════════════════════════════════
    // APPENDIX
    // ══════════════════════════════════════════════════════
    h1("Appendix — Key URLs & Resources"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3000, 6360],
      rows: [
        hdrRow(["Resource", "URL"], [3000, 6360]),
        dataRow(["Amazon PA API 5.0 Docs", "https://webservices.amazon.in/paapi5/documentation/"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["Flipkart Affiliate Docs", "https://affiliate.flipkart.com/static/docs/"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["Playwright Python Docs", "https://playwright.dev/python/docs/intro"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["Scrapy Documentation", "https://docs.scrapy.org/en/latest/"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["spaCy Documentation", "https://spacy.io/usage"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["Apache Airflow Docs", "https://airflow.apache.org/docs/"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["Elasticsearch 8.x Docs", "https://www.elastic.co/guide/en/elasticsearch/reference/8.0/"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["Next.js Documentation", "https://nextjs.org/docs"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["Prisma ORM Docs", "https://www.prisma.io/docs"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["Recharts API", "https://recharts.org/en-US/api"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["Bright Data Proxy Docs", "https://docs.brightdata.com/"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["SendGrid Docs", "https://docs.sendgrid.com/"], [3000, 6360], [C.white, C.ltgray]),
        dataRow(["Firebase FCM Docs", "https://firebase.google.com/docs/cloud-messaging"], [3000, 6360], [C.ltgray, C.white]),
        dataRow(["sentence-transformers", "https://www.sbert.net/docs/"], [3000, 6360], [C.white, C.ltgray]),
      ],
    }),
    gap(160),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "— PriceWise  ·  Smarter Shopping Starts Here —", size: 20, font: "Arial", color: C.gray, italics: true })],
      spacing: sp(200, 80),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "This document is confidential and intended for internal project use only.", size: 18, font: "Arial", color: C.midgray, italics: true })],
    }),
  ]
};

// ─── BUILD & SAVE ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 600, hanging: 300 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1000, hanging: 300 } } } },
        ],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: C.navy },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: C.blue },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: C.navy },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [coverSection, mainSection],
});

Packer.toBuffer(doc).then(buf => {
  const outputPath = path.join(__dirname, '..', 'PriceWise_Implementation_Plan.docx');
  fs.writeFileSync(outputPath, buf);
  console.log("Done! File saved to: " + outputPath);
});
