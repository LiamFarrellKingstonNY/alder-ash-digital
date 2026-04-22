"use strict";

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// ─── Icons ────────────────────────────────────────────────────────────────────
const {
  FaMobileAlt, FaBolt, FaSearchDollar, FaRobot, FaRocket, FaChartLine,
  FaCamera, FaMapMarkerAlt, FaCheckCircle, FaStar, FaEnvelope, FaPhone,
  FaUsers, FaLeaf, FaArrowRight, FaLightbulb, FaClock, FaShieldAlt,
  FaHandshake, FaTools, FaStore, FaTrophy
} = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#FFFFFF", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconBase64(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  green:  "254D3B",   // Forest green – primary
  cream:  "F7F2E8",   // Warm cream
  charcoal: "22221F", // Charcoal
  gold:   "C4892A",   // Amber/gold accent
  greenDark: "1A3529", // Deeper green for contrast
  greenMid:  "2E6B52", // Mid green
  creamDark: "EAE3D0", // Slightly darker cream for cards
  white:  "FFFFFF",
};

// ─── Shadow factory (must be fresh each call) ─────────────────────────────────
const mkShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.18 });

// ─── Helpers ─────────────────────────────────────────────────────────────────
function greenCard(slide, x, y, w, h) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.greenMid }, line: { color: C.gold, width: 1 } });
}
function creamCard(slide, x, y, w, h) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.cream }, shadow: mkShadow() });
}
function goldAccentBar(slide, x, y, h) {
  slide.addShape("rect", { x, y, w: 0.07, h, fill: { color: C.gold } });
}

async function buildDeck() {
  const pres = new pptxgen();
  pres.layout  = "LAYOUT_16x9";
  pres.author  = "Alder & Ash Digital";
  pres.title   = "Alder & Ash Digital – Pitch Deck";

  const W = 10, H = 5.625; // slide dimensions

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 1 – Title
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.green };

    // Left dark panel
    s.addShape("rect", { x: 0, y: 0, w: 5.8, h: H, fill: { color: C.greenDark } });

    // Gold vertical accent
    s.addShape("rect", { x: 5.8, y: 0, w: 0.07, h: H, fill: { color: C.gold } });

    // Right decorative dots grid (cream circles)
    for (let row = 0; row < 5; row++) {
      for (let col = 0; col < 4; col++) {
        s.addShape("ellipse", {
          x: 6.4 + col * 0.65, y: 0.6 + row * 1.0, w: 0.12, h: 0.12,
          fill: { color: C.cream, transparency: 60 }
        });
      }
    }

    // Leaf icon top-left
    const leafIcon = await iconBase64(FaLeaf, "#C4892A", 256);
    s.addImage({ data: leafIcon, x: 0.55, y: 0.45, w: 0.52, h: 0.52 });

    // Company name
    s.addText("ALDER & ASH DIGITAL", {
      x: 0.5, y: 1.1, w: 5.2, h: 0.55,
      fontSize: 28, bold: true, color: C.gold,
      fontFace: "Georgia", align: "left",
      charSpacing: 3, margin: 0
    });

    // Tagline
    s.addText("Websites That Work\nAs Hard As You Do", {
      x: 0.5, y: 1.8, w: 5.0, h: 1.4,
      fontSize: 34, bold: true, color: C.white,
      fontFace: "Georgia", align: "left", margin: 0
    });

    // Sub-tagline
    s.addText("Hudson Valley Web Design & AI Automation", {
      x: 0.5, y: 3.3, w: 5.0, h: 0.45,
      fontSize: 14, color: C.cream, fontFace: "Calibri", align: "left",
      italic: true, margin: 0
    });

    // Bottom gold bar
    s.addShape("rect", { x: 0.5, y: 3.95, w: 2.4, h: 0.05, fill: { color: C.gold } });

    // Right side tagline decoration
    s.addText("Local Roots.\nDigital Results.", {
      x: 6.0, y: 1.8, w: 3.7, h: 1.2,
      fontSize: 20, color: C.cream, fontFace: "Georgia",
      align: "center", italic: true, margin: 0
    });

    // Hudson Valley text bottom right
    s.addText("Hudson Valley, NY", {
      x: 6.0, y: 4.8, w: 3.7, h: 0.5,
      fontSize: 11, color: C.gold, fontFace: "Calibri",
      align: "center", charSpacing: 2, margin: 0
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 2 – The Problem
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.charcoal };

    // Top gold band
    s.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: C.gold } });

    // Section label
    s.addText("THE PROBLEM", {
      x: 0.5, y: 0.25, w: 9, h: 0.4,
      fontSize: 11, color: C.gold, fontFace: "Calibri",
      bold: true, charSpacing: 4, align: "left", margin: 0
    });

    // Title
    s.addText("Local Businesses Are Getting Left Behind", {
      x: 0.5, y: 0.72, w: 9, h: 0.75,
      fontSize: 30, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });

    // Pain point cards (2 rows × 2 cols)
    const pains = [
      { icon: FaMobileAlt, label: "No Mobile Optimization", desc: "Over half your visitors leave immediately if your site isn't mobile-friendly." },
      { icon: FaSearchDollar, label: "No Lead Capture", desc: "Traffic arrives but disappears — no forms, no follow-up, no pipeline." },
      { icon: FaRobot, label: "Manual Follow-Ups", desc: "Hours lost chasing leads by hand when automation could do it instantly." },
      { icon: FaShieldAlt, label: "Outdated & Untrustworthy", desc: "Slow, stale sites signal neglect. Prospects pick a competitor instead." },
    ];

    const cols = [0.45, 5.15];
    const rows = [1.68, 3.38];

    for (let i = 0; i < pains.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = cols[col], y = rows[row];
      const cw = 4.5, ch = 1.5;

      s.addShape("rect", { x, y, w: cw, h: ch, fill: { color: "2E2E2B" }, shadow: mkShadow() });
      // Gold left accent
      s.addShape("rect", { x, y, w: 0.07, h: ch, fill: { color: C.gold } });

      const ic = await iconBase64(pains[i].icon, "#C4892A", 256);
      s.addImage({ data: ic, x: x + 0.22, y: y + 0.35, w: 0.4, h: 0.4 });

      s.addText(pains[i].label, {
        x: x + 0.75, y: y + 0.18, w: cw - 0.85, h: 0.38,
        fontSize: 14, bold: true, color: C.cream,
        fontFace: "Georgia", align: "left", margin: 0
      });
      s.addText(pains[i].desc, {
        x: x + 0.75, y: y + 0.58, w: cw - 0.88, h: 0.76,
        fontSize: 11, color: "B8B0A0", fontFace: "Calibri",
        align: "left", margin: 0
      });
    }

    // Stat callouts bottom
    s.addShape("rect", { x: 0.45, y: 4.95, w: 4.5, h: 0.52, fill: { color: C.green } });
    s.addText("75% of consumers judge credibility by website design", {
      x: 0.55, y: 4.97, w: 4.3, h: 0.48,
      fontSize: 11, bold: true, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0, italic: true
    });
    s.addShape("rect", { x: 5.15, y: 4.95, w: 4.5, h: 0.52, fill: { color: C.green } });
    s.addText("60% of web traffic is mobile — is your site ready?", {
      x: 5.25, y: 4.97, w: 4.3, h: 0.48,
      fontSize: 11, bold: true, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0, italic: true
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 3 – The Solution
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    // Left dark panel
    s.addShape("rect", { x: 0, y: 0, w: 4.0, h: H, fill: { color: C.green } });

    // Left panel content
    s.addText("THE\nSOLUTION", {
      x: 0.35, y: 0.7, w: 3.2, h: 1.6,
      fontSize: 38, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });
    s.addShape("rect", { x: 0.35, y: 2.45, w: 1.6, h: 0.05, fill: { color: C.gold } });
    s.addText("AI-powered websites built for\nHudson Valley businesses —\nin days, not months.", {
      x: 0.35, y: 2.6, w: 3.25, h: 1.4,
      fontSize: 14, color: C.cream, fontFace: "Calibri",
      align: "left", margin: 0, italic: true
    });

    const rocket = await iconBase64(FaRocket, "#C4892A", 256);
    s.addImage({ data: rocket, x: 0.4, y: 4.4, w: 0.55, h: 0.55 });
    s.addText("Built to convert.", {
      x: 1.1, y: 4.45, w: 2.6, h: 0.45,
      fontSize: 13, bold: true, color: C.gold,
      fontFace: "Georgia", margin: 0
    });

    // Right panel – solution features
    const features = [
      { icon: FaBolt,        label: "Lightning Fast Websites",   desc: "Speed-optimized, mobile-first builds that score 95+ on Google." },
      { icon: FaRobot,       label: "AI Automation Built-In",    desc: "Lead follow-ups, appointment reminders, and more — on autopilot." },
      { icon: FaChartLine,   label: "Conversion-Optimized",      desc: "Every element designed to turn visitors into paying customers." },
      { icon: FaCamera,      label: "Visual Storytelling",       desc: "Photography-rooted design that makes your brand unforgettable." },
    ];

    const fStartX = 4.35;
    const fStartY = 0.55;
    const fGap    = 1.17;

    for (let i = 0; i < features.length; i++) {
      const y = fStartY + i * fGap;
      // Circle background
      s.addShape("ellipse", {
        x: fStartX, y: y + 0.08, w: 0.62, h: 0.62,
        fill: { color: C.green }
      });
      const ic = await iconBase64(features[i].icon, "#F7F2E8", 256);
      s.addImage({ data: ic, x: fStartX + 0.1, y: y + 0.18, w: 0.42, h: 0.42 });

      s.addText(features[i].label, {
        x: fStartX + 0.82, y: y + 0.08, w: 4.7, h: 0.35,
        fontSize: 15, bold: true, color: C.charcoal,
        fontFace: "Georgia", align: "left", margin: 0
      });
      s.addText(features[i].desc, {
        x: fStartX + 0.82, y: y + 0.44, w: 4.7, h: 0.52,
        fontSize: 12, color: "5A5750", fontFace: "Calibri",
        align: "left", margin: 0
      });

      if (i < features.length - 1) {
        s.addShape("rect", { x: fStartX + 0.82, y: y + fGap - 0.06, w: 4.8, h: 0.01, fill: { color: "D8D0C0" } });
      }
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 4 – Services Overview
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    // Top bar
    s.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: C.green } });

    s.addText("SERVICES", {
      x: 0.5, y: 0.22, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.green,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("Three Ways to Work Together", {
      x: 0.5, y: 0.62, w: 9, h: 0.65,
      fontSize: 30, bold: true, color: C.charcoal,
      fontFace: "Georgia", align: "left", margin: 0
    });

    // Columns
    const cols = [
      {
        label: "Launch Package", price: "$3,500",
        tag: "One-Time", color: C.greenMid,
        items: ["5-page website", "Mobile-first design", "Speed optimized", "Contact form + lead capture", "Google Analytics setup", "2 revision rounds", "Launch in 5 days"],
        icon: FaRocket
      },
      {
        label: "Growth System", price: "$5,500+",
        tag: "One-Time", color: C.green,
        items: ["Everything in Launch", "AI chatbot / lead bot", "Automated email follow-up", "CRM integration", "Booking / scheduling", "SEO foundation", "Analytics dashboard"],
        icon: FaChartLine
      },
      {
        label: "Growth Retainer", price: "$250/mo",
        tag: "Monthly", color: C.charcoal,
        items: ["Monthly content updates", "Performance monitoring", "Ongoing SEO", "Priority support", "Monthly strategy call", "New page adds included"],
        icon: FaHandshake
      },
    ];

    const colX = [0.38, 3.68, 6.98];
    const cardW = 2.98, cardH = 4.38;

    for (let i = 0; i < cols.length; i++) {
      const c = cols[i];
      const x = colX[i];

      s.addShape("rect", { x, y: 1.45, w: cardW, h: cardH, fill: { color: C.white }, shadow: mkShadow() });
      // Header bar
      s.addShape("rect", { x, y: 1.45, w: cardW, h: 1.0, fill: { color: c.color } });

      const ic = await iconBase64(c.icon, "#F7F2E8", 256);
      s.addImage({ data: ic, x: x + 0.2, y: 1.58, w: 0.42, h: 0.42 });

      s.addText(c.label, {
        x: x + 0.72, y: 1.52, w: cardW - 0.82, h: 0.38,
        fontSize: 13, bold: true, color: C.cream,
        fontFace: "Georgia", align: "left", margin: 0
      });
      s.addText(c.tag, {
        x: x + 0.72, y: 1.9, w: cardW - 0.82, h: 0.28,
        fontSize: 10, color: C.gold,
        fontFace: "Calibri", align: "left", margin: 0, italic: true
      });

      // Price
      s.addShape("rect", { x, y: 2.45, w: cardW, h: 0.55, fill: { color: C.gold } });
      s.addText(c.price, {
        x, y: 2.45, w: cardW, h: 0.55,
        fontSize: 22, bold: true, color: C.white,
        fontFace: "Georgia", align: "center", valign: "middle", margin: 0
      });

      // Items
      const itemTexts = c.items.map((item, j) => ({
        text: item,
        options: {
          bullet: false, breakLine: j < c.items.length - 1,
          fontSize: 11, fontFace: "Calibri", color: C.charcoal
        }
      }));

      // Add each item with a custom checkmark prefix
      let itemY = 3.1;
      for (const item of c.items) {
        s.addText([
          { text: "✓  ", options: { color: C.greenMid, bold: true, fontSize: 11, fontFace: "Calibri" } },
          { text: item,  options: { color: C.charcoal,  fontSize: 11, fontFace: "Calibri" } },
        ], { x: x + 0.2, y: itemY, w: cardW - 0.3, h: 0.32, margin: 0 });
        itemY += 0.3;
      }
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 5 – How It Works
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.green };

    s.addText("HOW IT WORKS", {
      x: 0.5, y: 0.28, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.gold,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("From First Call to Launch — Fast", {
      x: 0.5, y: 0.7, w: 9, h: 0.65,
      fontSize: 30, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });

    const steps = [
      { num: "01", label: "Discovery Call",   icon: FaPhone,    desc: "We learn your business, goals, and competition. No jargon." },
      { num: "02", label: "Design & Build",    icon: FaTools,    desc: "Your site is designed and built in 2–5 days. You review, we refine." },
      { num: "03", label: "Launch",            icon: FaRocket,   desc: "We handle the tech — hosting, domain, speed tuning, go-live." },
      { num: "04", label: "Grow",              icon: FaTrophy,   desc: "Ongoing automation, content, and strategy to keep growing." },
    ];

    const stepY = 1.65;
    const stepW = 2.1;

    for (let i = 0; i < steps.length; i++) {
      const x = 0.4 + i * 2.42;
      const st = steps[i];

      // Card
      s.addShape("rect", { x, y: stepY, w: stepW, h: 3.0, fill: { color: "1A3529" }, shadow: mkShadow() });

      // Number circle
      s.addShape("ellipse", { x: x + 0.65, y: stepY + 0.2, w: 0.7, h: 0.7, fill: { color: C.gold } });
      s.addText(st.num, {
        x: x + 0.65, y: stepY + 0.2, w: 0.7, h: 0.7,
        fontSize: 16, bold: true, color: C.white,
        fontFace: "Georgia", align: "center", valign: "middle", margin: 0
      });

      // Icon
      const ic = await iconBase64(st.icon, "#F7F2E8", 256);
      s.addImage({ data: ic, x: x + 0.73, y: stepY + 1.08, w: 0.55, h: 0.55 });

      // Label
      s.addText(st.label, {
        x: x + 0.12, y: stepY + 1.75, w: stepW - 0.2, h: 0.45,
        fontSize: 14, bold: true, color: C.gold,
        fontFace: "Georgia", align: "center", margin: 0
      });

      // Description
      s.addText(st.desc, {
        x: x + 0.12, y: stepY + 2.22, w: stepW - 0.2, h: 0.7,
        fontSize: 11, color: C.cream,
        fontFace: "Calibri", align: "center", margin: 0
      });

      // Arrow between steps
      if (i < steps.length - 1) {
        const arrowIc = await iconBase64(FaArrowRight, "#C4892A", 256);
        s.addImage({ data: arrowIc, x: x + stepW + 0.12, y: stepY + 1.22, w: 0.32, h: 0.32 });
      }
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 6 – Why Us
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    // Left accent bar
    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: H, fill: { color: C.green } });

    s.addText("WHY US", {
      x: 0.4, y: 0.25, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.green,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("What Sets Alder & Ash Apart", {
      x: 0.4, y: 0.65, w: 6.5, h: 0.7,
      fontSize: 30, bold: true, color: C.charcoal,
      fontFace: "Georgia", align: "left", margin: 0
    });

    const reasons = [
      { icon: FaRobot,        label: "AI-First Approach",          desc: "Every site is built with automation hooks — chatbots, email workflows, lead scoring." },
      { icon: FaBolt,         label: "Speed: Days, Not Months",    desc: "Launch in 2–5 days. Real results immediately, not after a 3-month agency slog." },
      { icon: FaChartLine,    label: "Conversion-Focused Design",  desc: "We don't build pretty sites. We build sites that turn visitors into customers." },
      { icon: FaCamera,       label: "Photography Background",     desc: "Rooted in visual storytelling — we know what makes people stop and pay attention." },
      { icon: FaMapMarkerAlt, label: "Local Trust",                desc: "Hudson Valley based. We understand local markets, community, and the businesses here." },
    ];

    const startY = 1.52;
    const rowH   = 0.75;

    for (let i = 0; i < reasons.length; i++) {
      const y = startY + i * rowH;
      const r = reasons[i];

      // Subtle row stripe on even
      if (i % 2 === 0) {
        s.addShape("rect", { x: 0.12, y: y - 0.04, w: W - 0.12, h: rowH - 0.05, fill: { color: "EEE8DA" } });
      }

      // Circle icon
      s.addShape("ellipse", { x: 0.35, y: y + 0.08, w: 0.5, h: 0.5, fill: { color: C.green } });
      const ic = await iconBase64(r.icon, "#F7F2E8", 256);
      s.addImage({ data: ic, x: 0.42, y: y + 0.15, w: 0.35, h: 0.35 });

      // Label
      s.addText(r.label, {
        x: 1.05, y: y + 0.08, w: 2.8, h: 0.33,
        fontSize: 13, bold: true, color: C.charcoal,
        fontFace: "Georgia", align: "left", margin: 0
      });
      // Description
      s.addText(r.desc, {
        x: 1.05, y: y + 0.38, w: 8.55, h: 0.28,
        fontSize: 11, color: "5A5750",
        fontFace: "Calibri", align: "left", margin: 0
      });
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 7 – Results
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.charcoal };

    s.addText("RESULTS", {
      x: 0.5, y: 0.28, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.gold,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("Numbers That Matter", {
      x: 0.5, y: 0.7, w: 9, h: 0.65,
      fontSize: 30, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });

    const stats = [
      { num: "3×",    label: "More Leads",         desc: "Average lead increase for clients in 90 days" },
      { num: "5",     label: "Days to Launch",      desc: "From Discovery Call to a live, optimized website" },
      { num: "98",    label: "Mobile Score",        desc: "Google PageSpeed mobile score, out of 100" },
      { num: "$0",    label: "Ad Spend Needed",     desc: "Organic-first strategy — no paid ads required" },
    ];

    const cardX = [0.38, 2.88, 5.38, 7.88];
    const cardW = 2.28;

    for (let i = 0; i < stats.length; i++) {
      const x = cardX[i];
      s.addShape("rect", { x, y: 1.6, w: cardW, h: 3.4, fill: { color: "2C2C29" }, shadow: mkShadow() });
      // Gold top accent
      s.addShape("rect", { x, y: 1.6, w: cardW, h: 0.06, fill: { color: C.gold } });

      s.addText(stats[i].num, {
        x, y: 1.85, w: cardW, h: 1.25,
        fontSize: 64, bold: true, color: C.gold,
        fontFace: "Georgia", align: "center", valign: "middle", margin: 0
      });
      s.addText(stats[i].label, {
        x, y: 3.18, w: cardW, h: 0.45,
        fontSize: 14, bold: true, color: C.cream,
        fontFace: "Georgia", align: "center", margin: 0
      });
      s.addShape("rect", { x: x + 0.5, y: 3.66, w: cardW - 1.0, h: 0.04, fill: { color: C.gold } });
      s.addText(stats[i].desc, {
        x, y: 3.78, w: cardW, h: 0.72,
        fontSize: 11, color: "9A9488",
        fontFace: "Calibri", align: "center", margin: 0
      });
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 8 – Case Study
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addText("CASE STUDY", {
      x: 0.5, y: 0.22, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.green,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("Catskill Peak Contracting — Website Transformation", {
      x: 0.5, y: 0.6, w: 9, h: 0.65,
      fontSize: 24, bold: true, color: C.charcoal,
      fontFace: "Georgia", align: "left", margin: 0
    });
    s.addText("A local general contractor in the Hudson Valley", {
      x: 0.5, y: 1.22, w: 9, h: 0.35,
      fontSize: 13, color: "7A7268",
      fontFace: "Calibri", align: "left", italic: true, margin: 0
    });

    // BEFORE column
    s.addShape("rect", { x: 0.4, y: 1.68, w: 4.3, h: 3.55, fill: { color: "E8E2D4" } });
    s.addShape("rect", { x: 0.4, y: 1.68, w: 4.3, h: 0.48, fill: { color: "8A8070" } });
    s.addText("BEFORE", {
      x: 0.4, y: 1.68, w: 4.3, h: 0.48,
      fontSize: 14, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle",
      charSpacing: 3, margin: 0
    });

    const befores = [
      ["4.8s", "Page Load Time"],
      ["12", "Leads / Month"],
      ["34%", "Mobile Score"],
      ["0", "Automated Follow-Ups"],
      ["2018", "Last Site Update"],
    ];
    befores.forEach(([val, lbl], i) => {
      s.addText(val, {
        x: 0.55, y: 2.3 + i * 0.56, w: 1.5, h: 0.44,
        fontSize: 22, bold: true, color: "8A4A4A",
        fontFace: "Georgia", align: "right", valign: "middle", margin: 0
      });
      s.addText(lbl, {
        x: 2.2, y: 2.3 + i * 0.56, w: 2.2, h: 0.44,
        fontSize: 13, color: C.charcoal,
        fontFace: "Calibri", valign: "middle", margin: 0
      });
    });

    // Arrow
    const arrowIc = await iconBase64(FaArrowRight, "#C4892A", 256);
    s.addImage({ data: arrowIc, x: 4.82, y: 3.18, w: 0.42, h: 0.42 });

    // AFTER column
    s.addShape("rect", { x: 5.35, y: 1.68, w: 4.3, h: 3.55, fill: { color: C.green } });
    s.addShape("rect", { x: 5.35, y: 1.68, w: 4.3, h: 0.48, fill: { color: C.greenDark } });
    s.addText("AFTER", {
      x: 5.35, y: 1.68, w: 4.3, h: 0.48,
      fontSize: 14, bold: true, color: C.gold,
      fontFace: "Calibri", align: "center", valign: "middle",
      charSpacing: 3, margin: 0
    });

    const afters = [
      ["0.9s", "Page Load Time"],
      ["41",   "Leads / Month"],
      ["98%",  "Mobile Score"],
      ["100%", "Automated Follow-Ups"],
      ["2025", "Launched in 5 Days"],
    ];
    afters.forEach(([val, lbl], i) => {
      s.addText(val, {
        x: 5.5, y: 2.3 + i * 0.56, w: 1.5, h: 0.44,
        fontSize: 22, bold: true, color: C.gold,
        fontFace: "Georgia", align: "right", valign: "middle", margin: 0
      });
      s.addText(lbl, {
        x: 7.15, y: 2.3 + i * 0.56, w: 2.3, h: 0.44,
        fontSize: 13, color: C.cream,
        fontFace: "Calibri", valign: "middle", margin: 0
      });
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 9 – About
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    // Top green panel
    s.addShape("rect", { x: 0, y: 0, w: W, h: 1.55, fill: { color: C.green } });

    s.addText("ABOUT", {
      x: 0.5, y: 0.22, w: 9, h: 0.35,
      fontSize: 11, bold: true, color: C.gold,
      fontFace: "Calibri", charSpacing: 4, align: "left", margin: 0
    });
    s.addText("The Person Behind the Work", {
      x: 0.5, y: 0.62, w: 9, h: 0.65,
      fontSize: 28, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });

    // Avatar placeholder (circle)
    s.addShape("ellipse", { x: 0.5, y: 1.8, w: 2.1, h: 2.1, fill: { color: C.greenMid } });
    const cameraIc = await iconBase64(FaCamera, "#F7F2E8", 256);
    s.addImage({ data: cameraIc, x: 0.92, y: 2.22, w: 1.25, h: 1.25 });

    // Gold ring around avatar
    s.addShape("ellipse", { x: 0.5, y: 1.8, w: 2.1, h: 2.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: C.gold, width: 2.5 } });

    // Placeholder name
    s.addText("Your Name Here", {
      x: 0.5, y: 3.98, w: 2.1, h: 0.4,
      fontSize: 12, bold: true, color: C.charcoal,
      fontFace: "Georgia", align: "center", margin: 0
    });
    s.addText("Founder", {
      x: 0.5, y: 4.36, w: 2.1, h: 0.3,
      fontSize: 10, color: C.greenMid, fontFace: "Calibri",
      align: "center", italic: true, margin: 0
    });

    // Bio right side
    goldAccentBar(s, 3.0, 1.72, 3.45);

    s.addText("Rooted in Hudson Valley.", {
      x: 3.22, y: 1.72, w: 6.4, h: 0.48,
      fontSize: 18, bold: true, color: C.charcoal,
      fontFace: "Georgia", align: "left", margin: 0
    });

    const bioPoints = [
      { icon: FaCamera,      text: "Background in professional photography — visual storytelling is in the DNA of every project." },
      { icon: FaMapMarkerAlt,text: "Based in the Hudson Valley, working exclusively with local and regional businesses." },
      { icon: FaRobot,       text: "Early adopter of AI tools and automation — obsessed with making small businesses more efficient." },
      { icon: FaLeaf,        text: "Passionate about helping community businesses compete with the big guys, on any budget." },
    ];

    let bioY = 2.38;
    for (const bp of bioPoints) {
      s.addShape("ellipse", { x: 3.2, y: bioY + 0.04, w: 0.38, h: 0.38, fill: { color: C.green } });
      const ic = await iconBase64(bp.icon, "#F7F2E8", 256);
      s.addImage({ data: ic, x: 3.27, y: bioY + 0.10, w: 0.25, h: 0.25 });

      s.addText(bp.text, {
        x: 3.72, y: bioY + 0.03, w: 5.85, h: 0.42,
        fontSize: 12, color: "5A5750",
        fontFace: "Calibri", align: "left", margin: 0
      });
      bioY += 0.65;
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 10 – Let's Talk (CTA)
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.greenDark };

    // Large decorative dot grid
    for (let row = 0; row < 8; row++) {
      for (let col = 0; col < 6; col++) {
        s.addShape("ellipse", {
          x: 0.3 + col * 0.85, y: 0.25 + row * 0.65, w: 0.1, h: 0.1,
          fill: { color: C.cream, transparency: 75 }
        });
      }
    }

    // Right content area
    s.addShape("rect", { x: 4.8, y: 0, w: 5.2, h: H, fill: { color: C.green } });
    s.addShape("rect", { x: 4.8, y: 0, w: 0.07, h: H, fill: { color: C.gold } });

    // Left headline
    s.addText("Let's\nBuild\nSomething.", {
      x: 0.45, y: 0.7, w: 4.1, h: 3.2,
      fontSize: 46, bold: true, color: C.cream,
      fontFace: "Georgia", align: "left", margin: 0
    });
    s.addText("Your free audit is waiting.", {
      x: 0.45, y: 4.05, w: 4.1, h: 0.5,
      fontSize: 15, color: C.gold,
      fontFace: "Calibri", align: "left", italic: true, margin: 0
    });

    // Right CTA content
    s.addText("GET YOUR FREE\nWEBSITE AUDIT", {
      x: 5.1, y: 0.55, w: 4.7, h: 1.0,
      fontSize: 22, bold: true, color: C.gold,
      fontFace: "Georgia", align: "left",
      charSpacing: 1, margin: 0
    });

    s.addShape("rect", { x: 5.1, y: 1.62, w: 4.0, h: 0.04, fill: { color: C.gold } });

    s.addText("We'll review your current site and show you exactly\nwhat's costing you leads — for free.", {
      x: 5.1, y: 1.78, w: 4.65, h: 0.72,
      fontSize: 13, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0
    });

    // CTA button
    s.addShape("rect", { x: 5.1, y: 2.68, w: 3.75, h: 0.6, fill: { color: C.gold } });
    s.addText("Schedule Your Free Audit →", {
      x: 5.1, y: 2.68, w: 3.75, h: 0.6,
      fontSize: 14, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0
    });

    // Contact items
    const emailIc = await iconBase64(FaEnvelope, "#C4892A", 256);
    const phoneIc = await iconBase64(FaPhone,    "#C4892A", 256);
    const mapIc   = await iconBase64(FaMapMarkerAlt, "#C4892A", 256);

    s.addImage({ data: emailIc, x: 5.1, y: 3.55, w: 0.3, h: 0.3 });
    s.addText("hello@alderandash.com", {
      x: 5.52, y: 3.55, w: 4.15, h: 0.32,
      fontSize: 13, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0
    });

    s.addImage({ data: phoneIc, x: 5.1, y: 4.02, w: 0.3, h: 0.3 });
    s.addText("(845) 555-0000", {
      x: 5.52, y: 4.02, w: 4.15, h: 0.32,
      fontSize: 13, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0
    });

    s.addImage({ data: mapIc, x: 5.1, y: 4.48, w: 0.3, h: 0.3 });
    s.addText("Hudson Valley, New York", {
      x: 5.52, y: 4.48, w: 4.15, h: 0.32,
      fontSize: 13, color: C.cream,
      fontFace: "Calibri", align: "left", margin: 0
    });

    // Website
    s.addText("alderandash.com", {
      x: 5.1, y: 5.1, w: 4.5, h: 0.35,
      fontSize: 11, color: C.gold,
      fontFace: "Calibri", align: "left",
      italic: true, charSpacing: 1, margin: 0
    });
  }

  // ──────────────────────────────────────────────────────────────────────────
  const outPath = "/Users/ashai/Desktop/alder-ash-digital/Alder-Ash-Pitch-Deck.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("✅  Written:", outPath);
}

buildDeck().catch(err => { console.error(err); process.exit(1); });
