/**
 * RohWinBghit — Master's Thesis Defense Presentation Redesign
 * 33 slides of polished, visual content using Forest & Moss design system.
 * Usage: node "generate_presentation_v2.js"
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

// ─── DESIGN SYSTEM CONSTANTS ────────────────────────────────────────────────
const COLOR_PRIMARY_DARK = "1B5E3B";
const COLOR_MEDIUM_GREEN = "2E7D52";
const COLOR_LIGHT_GREEN  = "4CAF72";
const COLOR_GOLD         = "C9A84C";
const COLOR_BG_LIGHT     = "F7F5F0";
const COLOR_BG_DARK      = "0F3D2B";
const COLOR_WHITE        = "FFFFFF";
const COLOR_DARK_TEXT    = "1A1A1A";
const COLOR_MUTED_TEXT   = "5A5A5A";
const COLOR_LIGHT_GRAY   = "E0E0E0";
const COLOR_RED          = "C0392B";

const FONT_TITLE = "Cambria";
const FONT_BODY  = "Calibri";

// Shadow Factory Function (Avoids mutation bugs)
const makeShadow = () => ({
    type: "outer",
    color: "000000",
    blur: 6,
    offset: 2,
    angle: 45,
    opacity: 0.12
});

// Layout Dimensions
const SW = 10.0;
const SH = 5.625;

// Image Paths (corrected folder case to Screen_Mobile and matched actual assets)
const IMG_LOGO          = path.join(__dirname, "Screen_Mobile", "logo_bus.png");
const IMG_UNIV_LOGO     = path.join(__dirname, "figures", "logo_universite.png");
const IMG_SPLASH        = path.join(__dirname, "figures", "v2_welcome_splash.jpg");
const IMG_INSCRIP       = path.join(__dirname, "Screen_Mobile", "02_inscription_role.png");
const IMG_RESERVE       = path.join(__dirname, "Screen_Mobile", "06_recherche_details_reservation.png");
const IMG_QR_DEPART     = path.join(__dirname, "Screen_Mobile", "07_ticket_confirme_qr.png");

// New Technical diagrams and screenshots
const IMG_USE_CASE      = path.join(__dirname, "figures", "fig_use_case.png");
const IMG_CLASS_DIAGRAM = path.join(__dirname, "figures", "fig_class_diagram.png");
const IMG_DEPLOYMENT    = path.join(__dirname, "figures", "fig_deployment2.png");
const IMG_JEST_RESULTS  = path.join(__dirname, "Screen_Mobile", "jest_results.png");
const IMG_SEQ_KYC       = path.join(__dirname, "figures", "fig_sequence_kyc.png");
const IMG_KYC_INTRO     = path.join(__dirname, "Screen_Mobile", "17_kyc_intro.png");
const IMG_KYC_FACE      = path.join(__dirname, "Screen_Mobile", "20_kyc_capture_visage.png");
const IMG_SEQ_CHECKIN   = path.join(__dirname, "figures", "fig_sequence_checkin.png");
const IMG_SEQ_BOOKING   = path.join(__dirname, "figures", "fig_sequence_booking.png");
const IMG_ARCHITECTURE_DIAG = path.join(__dirname, "figures", "Architecture.png");
const IMG_DEPLOYMENT_CLOUD   = path.join(__dirname, "figures", "déploiement_cloud.png");

// Helper to encode image files to Base64 data URLs for reliable pptxgenjs embedding in Node.js
const toBase64 = (filePath) => {
    if (!fs.existsSync(filePath)) return null;
    const ext = path.extname(filePath).toLowerCase().replace(".", "");
    const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : `image/${ext}`;
    const base64 = fs.readFileSync(filePath, { encoding: "base64" });
    return `${mime};base64,${base64}`;
};

// Helper to check if file exists, if not logs warning
const verifyPath = (p, label) => {
    if (!fs.existsSync(p)) {
        console.warn(`⚠️ Warning: Image file not found for ${label} at: ${p}`);
    }
};

verifyPath(IMG_LOGO, "Logo Bus");
verifyPath(IMG_UNIV_LOGO, "University Logo");
verifyPath(IMG_SPLASH, "Splash Screen");
verifyPath(IMG_INSCRIP, "Inscription Screen");
verifyPath(IMG_RESERVE, "Reserve Screen");
verifyPath(IMG_QR_DEPART, "QR Scan Screen");
verifyPath(IMG_USE_CASE, "Use Case Diagram");
verifyPath(IMG_CLASS_DIAGRAM, "Class Diagram");
verifyPath(IMG_DEPLOYMENT, "Deployment Diagram");
verifyPath(IMG_JEST_RESULTS, "Jest Results Screenshot");
verifyPath(IMG_SEQ_KYC, "KYC Sequence Diagram");
verifyPath(IMG_KYC_INTRO, "KYC Screen 1 (Intro)");
verifyPath(IMG_KYC_FACE, "KYC Screen 2 (Face)");
verifyPath(IMG_SEQ_CHECKIN, "Check-in Sequence Diagram");
verifyPath(IMG_SEQ_BOOKING, "Booking Sequence Diagram");
verifyPath(IMG_ARCHITECTURE_DIAG, "Architecture Diagram");
verifyPath(IMG_DEPLOYMENT_CLOUD, "Deployment Cloud Screen");

// ─── INITIALIZE PRESENTATION ────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane";
pres.title = "RohWinBghit — Soutenance de Master en Génie Logiciel";

// ─── TEMPLATE BUILDERS ──────────────────────────────────────────────────────

/** Adds the standard content slide headers (section title, label, page number) */
function applyContentSlideTemplate(slide, title, sectionLabel, slideNumText) {
    slide.background = { color: COLOR_BG_LIGHT };

    // Section Label (Top Left)
    slide.addText(sectionLabel.toUpperCase(), {
        x: 0.5, y: 0.3, w: 7.0, h: 0.25,
        fontSize: 9, bold: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "left"
    });

    // Slide Number (Top Right)
    slide.addText(slideNumText, {
        x: 8.5, y: 0.3, w: 1.0, h: 0.25,
        fontSize: 10, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "right"
    });

    // Slide Title (Top Left, Cambria 32pt bold)
    slide.addText(title, {
        x: 0.5, y: 0.5, w: 8.5, h: 0.6,
        fontSize: 32, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "left"
    });
}

/** Builds a dark background section divider slide */
function buildSectionDivider(slideNum, title, subtitle, badgeText) {
    const s = pres.addSlide();
    s.background = { color: COLOR_BG_DARK };

    // Large Watermark Number (opacity simulated using light-colored text with high transparency)
    s.addText(String(slideNum).padStart(2, "0"), {
        x: 2.0, y: 1.0, w: 6.0, h: 3.5,
        fontSize: 180, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center", valign: "middle",
        transparency: 90
    });

    // Section Title
    s.addText(title, {
        x: 0.5, y: 2.0, w: 9.0, h: 0.8,
        fontSize: 44, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center", valign: "middle"
    });

    // Subtitle Question
    s.addText(subtitle, {
        x: 0.5, y: 2.9, w: 9.0, h: 0.5,
        fontSize: 16, italic: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });

    // Section Badge (Top Right)
    s.addShape("ellipse", {
        x: 8.8, y: 0.35, w: 0.6, h: 0.6,
        fill: { color: COLOR_GOLD }, line: { color: COLOR_GOLD, width: 0 }
    });
    s.addText(badgeText, {
        x: 8.8, y: 0.35, w: 0.6, h: 0.6,
        fontSize: 12, bold: true, color: COLOR_BG_DARK,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });

    return s;
}

/** Renders a card shape helper */
function drawCard(slide, x, y, w, h, options = {}) {
    const fill = options.fill || COLOR_WHITE;
    const line = options.line || { color: "E7ECE8", width: 0.5 };
    slide.addShape("roundRect", {
        x, y, w, h,
        fill: { color: fill },
        line: line,
        rectRadius: 0.12,
        shadow: makeShadow()
    });
}

// =============================================================================
// SLIDE 01 — COVER SLIDE (Dark Background)
// =============================================================================
{
    const s = pres.addSlide();
    s.background = { color: COLOR_BG_DARK };

    // University Header
    s.addText("Université Abou Bekr Belkaïd – Tlemcen\nFaculté des Sciences – Département d'Informatique", {
        x: 0.5, y: 0.35, w: 6.0, h: 0.6,
        fontSize: 13, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "left"
    });

    // Arabic Header (Right Aligned)
    s.addText("جامعة أبي بكر بلقايد – تلمسان\nكلية العلوم – قسم الإعلام الآلي", {
        x: 6.5, y: 0.35, w: 3.0, h: 0.6,
        fontSize: 13, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "right"
    });

    // University Logo (Centered Top)
    const logoData = toBase64(IMG_UNIV_LOGO);
    if (logoData) {
        s.addImage({ data: logoData, x: 4.65, y: 0.95, w: 0.7, h: 0.7 });
    }

    // Title
    s.addText("RohWinBghit", {
        x: 0.5, y: 1.7, w: 9.0, h: 0.5,
        fontSize: 38, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center"
    });

    s.addText("روح وين بغيت", {
        x: 0.5, y: 2.2, w: 9.0, h: 0.4,
        fontSize: 26, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center"
    });

    // Project Description (Italic Calibri 13pt, light green)
    s.addText("Plateforme mobile multiplateforme intelligente de covoiturage inter-wilayas sécurisée\nadaptée au contexte algérien", {
        x: 0.5, y: 2.65, w: 9.0, h: 0.5,
        fontSize: 13, italic: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "center"
    });

    // Subtitle / Framework
    s.addText("Mémoire de fin d'études de Master — Option : Génie Logiciel (GL)\nArrêté Ministériel 1275 — « Un diplôme, une Startup »", {
        x: 0.5, y: 3.2, w: 9.0, h: 0.5,
        fontSize: 11.5, bold: true, color: COLOR_GOLD,
        fontFace: FONT_BODY, align: "center"
    });

    // Presented / Supervised columns
    s.addText("Présenté par :\nAHMED BACHA Djamel Eddine\nBELHORMA Sidi Mohammed Reduane\n\nSoutenu le : 20 juin 2026", {
        x: 0.5, y: 3.8, w: 4.25, h: 1.1,
        fontSize: 10.5, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center"
    });

    s.addText("Jury de Soutenance :\nPrésidente : Mme BENMAHDI Meriem\nEncadrante : Mme BENLEDGHEM Rafika\nExaminatrice : Mme BENMANSOUR Fazilet\nExperte I2E : Mme SELADJI Yassamine", {
        x: 5.25, y: 3.8, w: 4.25, h: 1.1,
        fontSize: 10.5, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center"
    });

    // Tagline / Date Footer
    s.addShape("rect", { x: 0, y: 5.0, w: SW, h: 0.625, fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 } });
    s.addText("Année Universitaire : 2025/2026   •   SIMPLE  •  SÛR  •  ABORDABLE", {
        x: 0, y: 5.0, w: SW, h: 0.625,
        fontSize: 13, bold: true, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 02 — TABLE OF CONTENTS
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Plan de la Présentation", "Plan", "02/33");

    const sections = [
        { num: "01", title: "Introduction & Problématique", sub: "Mobilité inter-wilayas • Enjeux", active: true },
        { num: "02", title: "Contexte & Cadre Légal", sub: "Loi 25-11 • Concurrence", active: false },
        { num: "03", title: "Méthodologie & Architecture", sub: "Agile • Stack • Pipeline KYC", active: false },
        { num: "04", title: "Conception & Modélisation", sub: "UML • Cas d'utilisation", active: false },
        { num: "05", title: "Résultats & Validation", sub: "Tests • SUS • Hypothèses", active: false },
        { num: "06", title: "Modèle Économique", sub: "BMC • Prévisions financières", active: false },
        { num: "07", title: "Conclusion & Perspectives", sub: "Apports • Limitations • Futur", active: false }
    ];

    // Grid layout: 4 on top row, 3 on bottom row (centered)
    const cardW = 2.05;
    const cardH = 1.6;
    const yTop = 1.4;
    const yBot = 3.25;

    // Top row
    sections.slice(0, 4).forEach((sec, i) => {
        const x = 0.5 + i * (cardW + 0.27);
        drawCard(s, x, yTop, cardW, cardH, {
            fill: sec.active ? "E8F5EE" : COLOR_WHITE,
            line: sec.active ? { color: COLOR_PRIMARY_DARK, width: 1.5 } : { color: "E7ECE8", width: 0.5 }
        });

        // Circle number badge
        s.addShape("ellipse", {
            x: x + 0.15, y: yTop + 0.15, w: 0.35, h: 0.35,
            fill: { color: sec.active ? COLOR_PRIMARY_DARK : COLOR_MEDIUM_GREEN },
            line: { width: 0 }
        });
        s.addText(sec.num, {
            x: x + 0.15, y: yTop + 0.15, w: 0.35, h: 0.35,
            fontSize: 10, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Section Title
        s.addText(sec.title, {
            x: x + 0.15, y: yTop + 0.6, w: cardW - 0.3, h: 0.5,
            fontSize: 13, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left", valign: "top"
        });

        // Subtitle
        s.addText(sec.sub, {
            x: x + 0.15, y: yTop + 1.15, w: cardW - 0.3, h: 0.4,
            fontSize: 10, italic: true, color: COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "left"
        });
    });

    // Bottom row (3 cards, centered)
    const startXBot = 1.66;
    sections.slice(4).forEach((sec, i) => {
        const x = startXBot + i * (cardW + 0.27);
        drawCard(s, x, yBot, cardW, cardH, {
            fill: COLOR_WHITE,
            line: { color: "E7ECE8", width: 0.5 }
        });

        // Circle number badge
        s.addShape("ellipse", {
            x: x + 0.15, y: yBot + 0.15, w: 0.35, h: 0.35,
            fill: { color: COLOR_MEDIUM_GREEN },
            line: { width: 0 }
        });
        s.addText(sec.num, {
            x: x + 0.15, y: yBot + 0.15, w: 0.35, h: 0.35,
            fontSize: 10, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Section Title
        s.addText(sec.title, {
            x: x + 0.15, y: yBot + 0.6, w: cardW - 0.3, h: 0.5,
            fontSize: 13, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left", valign: "top"
        });

        // Subtitle
        s.addText(sec.sub, {
            x: x + 0.15, y: yBot + 1.15, w: cardW - 0.3, h: 0.4,
            fontSize: 10, italic: true, color: COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "left"
        });
    });
}

// =============================================================================
// SLIDE 03 — SECTION DIVIDER 01 (Introduction & Problématique)
// =============================================================================
buildSectionDivider("01", "Introduction & Problématique", "Où ? Quand ? Dans quel contexte ?", "01");

// =============================================================================
// SLIDE 04 — LA MOBILITÉ INTER-WILAYAS EN ALGÉRIE
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "La Mobilité Inter-Wilayas en Algérie", "Introduction & Problématique", "04/33");

    const kpis = [
        { num: "6.3M", label: "étudiants", desc: "d'étudiants universitaires effectuent des trajets inter-wilayas chaque année" },
        { num: "80%", label: "sièges vidés", desc: "des véhicules particuliers roulent avec des sièges vides sur les axes routiers" },
        { num: "0", label: "Plateformes KYC", desc: "plateformes nationales de covoiturage longue distance sécurisées par KYC biométrique" }
    ];

    const cardW = 2.8;
    const cardH = 2.7;

    kpis.forEach((kpi, i) => {
        const x = 0.5 + i * 3.1;
        const y = 1.4;

        drawCard(s, x, y, cardW, cardH);

        // KPI Icon Stand-In
        s.addShape("ellipse", {
            x: x + 1.1, y: y + 0.2, w: 0.6, h: 0.6,
            fill: { color: "E8F5EE" },
            line: { color: COLOR_PRIMARY_DARK, width: 1 }
        });
        s.addText(String(i + 1), {
            x: x + 1.1, y: y + 0.2, w: 0.6, h: 0.6,
            fontSize: 16, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Stat Number
        s.addText(kpi.num, {
            x: x + 0.1, y: y + 0.9, w: cardW - 0.2, h: 0.6,
            fontSize: 48, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Label
        s.addText(kpi.label.toUpperCase(), {
            x: x + 0.1, y: y + 1.55, w: cardW - 0.2, h: 0.25,
            fontSize: 11, bold: true, color: COLOR_GOLD,
            fontFace: FONT_BODY, align: "center"
        });

        // Description
        s.addText(kpi.desc, {
            x: x + 0.2, y: y + 1.85, w: cardW - 0.4, h: 0.7,
            fontSize: 12, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "center"
        });
    });

    // Warning Footer Bar
    const fy = 4.4;
    s.addShape("roundRect", {
        x: 0.5, y: fy, w: 9.0, h: 0.5,
        fill: { color: COLOR_GOLD },
        line: { width: 0 },
        rectRadius: 0.08
    });
    s.addText("⚠️  Risques : Aucune vérification d'identité  •  Paiement en espèces uniquement  •  Groupes Facebook non sécurisés", {
        x: 0.5, y: fy, w: 9.0, h: 0.5,
        fontSize: 11, bold: true, color: COLOR_DARK_TEXT,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 05 — LE PROBLÈME DE CONFIANCE & SÉCURITÉ
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Le Problème de Confiance & Sécurité", "Introduction & Problématique", "05/33");

    // Research Question Box
    const qy = 1.25;
    s.addShape("roundRect", {
        x: 0.5, y: qy, w: 9.0, h: 0.95,
        fill: { color: "E8F5EE" },
        line: { color: COLOR_PRIMARY_DARK, width: 1.5 },
        rectRadius: 0.12
    });
    s.addText("QUESTION CENTRALE DE RECHERCHE", {
        x: 0.5, y: qy + 0.1, w: 9.0, h: 0.2,
        fontSize: 9, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY, align: "center"
    });
    s.addText("Comment concevoir une plateforme de covoiturage inter-wilayas intégrant un pipeline biométrique KYC pour garantir la confiance et la sécurité des utilisateurs ?", {
        x: 0.7, y: qy + 0.3, w: 8.6, h: 0.55,
        fontSize: 14, italic: true, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "center", valign: "middle"
    });

    // 4 problem cards in a row
    const probs = [
        { title: "Identité", letter: "I", desc: "Vérification de l'identité réelle des conducteurs et passagers." },
        { title: "Paiement", letter: "P", desc: "Intégration du paiement local par cartes monétiques (CIB, Edahabia)." },
        { title: "Résilience", letter: "R", desc: "Architecture API tolérante aux coupures de connexion mobile interurbaine." },
        { title: "Couverture", letter: "C", desc: "Desservir efficacement et de manière optimisée les 69 wilayas du pays." }
    ];

    const cy = 2.45;
    const cardW = 2.025;
    const cardH = 2.3;

    probs.forEach((p, i) => {
        const x = 0.5 + i * (cardW + 0.3);
        drawCard(s, x, cy, cardW, cardH);

        // Abbreviation Circle Badge (Stand-in for icon)
        s.addShape("ellipse", {
            x: x + (cardW - 0.5) / 2, y: cy + 0.2, w: 0.5, h: 0.5,
            fill: { color: COLOR_PRIMARY_DARK },
            line: { width: 0 }
        });
        s.addText(p.letter, {
            x: x + (cardW - 0.5) / 2, y: cy + 0.2, w: 0.5, h: 0.5,
            fontSize: 14, bold: true, color: COLOR_WHITE,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Title
        s.addText(p.title, {
            x: x + 0.1, y: cy + 0.8, w: cardW - 0.2, h: 0.3,
            fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "center"
        });

        // Description
        s.addText(p.desc, {
            x: x + 0.15, y: cy + 1.15, w: cardW - 0.3, h: 1.0,
            fontSize: 11, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "center", valign: "top"
        });
    });
}

// =============================================================================
// SLIDE 06 — TROIS HYPOTHÈSES DE TRAVAIL
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Trois Hypothèses de Travail", "Introduction & Problématique", "06/33");

    const hyps = [
        {
            num: "H1",
            title: "KYC Biométrique",
            badge: "INNOVATION",
            color: COLOR_MEDIUM_GREEN,
            desc: "Un pipeline d'IA combinant OCR, ArcFace et détection de vivacité peut vérifier l'identité des utilisateurs avec un seuil de confiance ≥ 0.65 (conducteurs) et ≥ 0.60 (passagers)."
        },
        {
            num: "H2",
            title: "Usabilité Mobile",
            badge: "PROTOTYPAGE",
            color: "028090",
            desc: "L'application mobile obtient un score SUS ≥ 68/100, validant une expérience utilisateur de qualité « Good » selon les standards ISO."
        },
        {
            num: "H3",
            title: "Résilience API",
            badge: "SOLVABILITÉ",
            color: COLOR_GOLD,
            desc: "L'architecture microservices avec circuit breaker, retry patterns et health checks maintient un taux de disponibilité ≥ 99%."
        }
    ];

    const cardW = 9.0;
    const cardH = 1.1;

    hyps.forEach((h, i) => {
        const y = 1.3 + i * 1.25;

        drawCard(s, 0.5, y, cardW, cardH);

        // H-number badge (Gold rounded box)
        s.addShape("roundRect", {
            x: 0.7, y: y + 0.25, w: 0.6, h: 0.6,
            fill: { color: COLOR_PRIMARY_DARK },
            line: { width: 0 },
            rectRadius: 0.1
        });
        s.addText(h.num, {
            x: 0.7, y: y + 0.25, w: 0.6, h: 0.6,
            fontSize: 20, bold: true, color: COLOR_GOLD,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Title
        s.addText(h.title, {
            x: 1.45, y: y + 0.2, w: 3.5, h: 0.35,
            fontSize: 16, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left"
        });

        // Badge pill right
        s.addShape("roundRect", {
            x: 7.7, y: y + 0.22, w: 1.5, h: 0.3,
            fill: { color: h.color },
            line: { width: 0 },
            rectRadius: 0.2
        });
        s.addText(h.badge, {
            x: 7.7, y: y + 0.22, w: 1.5, h: 0.3,
            fontSize: 9, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Description text
        s.addText(h.desc, {
            x: 1.45, y: y + 0.55, w: 7.7, h: 0.45,
            fontSize: 12, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "left", valign: "top"
        });
    });
}

// =============================================================================
// SLIDE 07 — SECTION DIVIDER 02 (Contexte & Cadre Légal)
// =============================================================================
buildSectionDivider("02", "Contexte & Cadre Légal", "Quel marché ? Quelle réglementation ?", "02");

// =============================================================================
// SLIDE 08 — ANALYSE CONCURRENTIELLE
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Analyse Concurrentielle", "Contexte & Cadre Légal", "08/33");

    // Table Data
    const tableRows = [
        [
            { text: "Critère", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "BlaBlaCar", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Yassir", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Nroho", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "inDrive", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Covoit'Dz", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "RohWinBghit", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } }
        ],
        [
            { text: "Couverture DZ" },
            { text: "Absent", options: { color: COLOR_RED } },
            { text: "Urbaine" },
            { text: "Nationale" },
            { text: "Nationale" },
            { text: "Nationale" },
            { text: "69 Wilayas", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ],
        [
            { text: "KYC Biométrique" },
            { text: "Standard" },
            { text: "Standard" },
            { text: "CIN + Selfie" },
            { text: "Documentaire" },
            { text: "Néant", options: { color: COLOR_RED } },
            { text: "Liveness + Anti-spoof", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ],
        [
            { text: "Seuils par Rôle" },
            { text: "Non" },
            { text: "Non" },
            { text: "Non" },
            { text: "Non" },
            { text: "Non" },
            { text: "Oui (≥0.65 / ≥0.60)", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ],
        [
            { text: "Paiement Local" },
            { text: "Non", options: { color: COLOR_RED } },
            { text: "Partiel", options: { color: COLOR_GOLD, bold: true } },
            { text: "Non (cash)" },
            { text: "Non (cash)" },
            { text: "Non", options: { color: COLOR_RED } },
            { text: "Chargily Pay V2", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ],
        [
            { text: "Tarif Dynamique" },
            { text: "Non" },
            { text: "Oui", options: { color: COLOR_PRIMARY_DARK, bold: true } },
            { text: "Non" },
            { text: "Non" },
            { text: "Fixe" },
            { text: "Oui (m ∈ [1.0; 2.5])", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ],
        [
            { text: "Sécurité PII" },
            { text: "Non doc." },
            { text: "Non doc." },
            { text: "Non doc." },
            { text: "Non doc." },
            { text: "Non doc." },
            { text: "AES-256-GCM", options: { color: COLOR_PRIMARY_DARK, bold: true } }
        ]
    ];

    s.addTable(tableRows, {
        x: 0.5, y: 1.2, w: 9.0, h: 2.9,
        border: { type: "solid", color: "E7ECE8", size: 1 },
        align: "center",
        valign: "middle",
        fontFace: FONT_BODY,
        fontSize: 10,
        colWidths: [1.8, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2]
    });

    // Takeaway Footer Card
    const fy = 4.25;
    s.addShape("roundRect", {
        x: 0.5, y: fy, w: 9.0, h: 0.6,
        fill: { color: COLOR_BG_DARK },
        line: { width: 0 },
        rectRadius: 0.08
    });
    s.addText("RohWinBghit résout les 5 lacunes du marché en combinant KYC biométrique, paiement local et couverture nationale", {
        x: 0.5, y: fy, w: 9.0, h: 0.6,
        fontSize: 11.5, bold: true, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 09 — CADRE RÉGLEMENTAIRE : LOI 25-11
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Cadre Réglementaire : Loi 25-11", "Contexte & Cadre Légal", "09/33");

    s.addText("Protection des données personnelles dans le contexte algérien", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.3,
        fontSize: 13, italic: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY
    });

    const regulations = [
        { title: "Consentement Explicite", art: "Art. 7", desc: "Collecte de données sensibles (biométrie) uniquement après accord explicite de l'utilisateur." },
        { title: "Chiffrement AES-256", art: "Art. 42", desc: "Données biométriques et d'identification stockées sous forme chiffrée côté serveur." },
        { title: "Traitement Local", art: "Art. 44", desc: "Traitement des données personnelles sur le territoire algérien sans transfert transfrontalier." },
        { title: "Droit à l'Effacement", art: "Art. 34", desc: "Garantir à l'utilisateur la suppression définitive de ses données sur demande." }
    ];

    const cardW = 4.35;
    const cardH = 1.4;

    regulations.forEach((r, i) => {
        const row = Math.floor(i / 2);
        const col = i % 2;
        const x = 0.5 + col * (cardW + 0.3);
        const y = 1.5 + row * (cardH + 0.35);

        drawCard(s, x, y, cardW, cardH);

        // Title
        s.addText(r.title, {
            x: x + 0.2, y: y + 0.15, w: cardW - 1.2, h: 0.35,
            fontSize: 16, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left"
        });

        // Art Badge (Gold badge circle)
        s.addShape("ellipse", {
            x: x + cardW - 0.85, y: y + 0.15, w: 0.7, h: 0.3,
            fill: { color: COLOR_GOLD }, line: { width: 0 }
        });
        s.addText(r.art, {
            x: x + cardW - 0.85, y: y + 0.15, w: 0.7, h: 0.3,
            fontSize: 9, bold: true, color: COLOR_BG_DARK,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Description
        s.addText(r.desc, {
            x: x + 0.2, y: y + 0.55, w: cardW - 0.4, h: 0.7,
            fontSize: 12, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "left", valign: "top"
        });
    });
}

// =============================================================================
// SLIDE 10 — SECTION DIVIDER 03 (Méthodologie & Architecture)
// =============================================================================
buildSectionDivider("03", "Méthodologie & Architecture", "Comment ? Avec quelles technologies ?", "03");

// =============================================================================
// SLIDE 11 — GESTION AGILE : 5 SPRINTS EN 10 SEMAINES
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Gestion Agile : 5 Sprints en 10 Semaines", "Méthodologie & Architecture", "11/33");

    const sprints = [
        { name: "Sprint 1 : Auth & Profils", time: "Semaines 1-2", pts: "21 SP", pct: 100 },
        { name: "Sprint 2 : Trajets & Réservation", time: "Semaines 3-4", pts: "34 SP", pct: 100 },
        { name: "Sprint 3 : Pipeline KYC Biométrique", time: "Semaines 5-6", pts: "28 SP", pct: 100 },
        { name: "Sprint 4 : Paiement & QR Code", time: "Semaines 7-8", pts: "25 SP", pct: 100 },
        { name: "Sprint 5 : Admin Console & Tests", time: "Semaines 9-10", pts: "22 SP", pct: 100 }
    ];

    const rowH = 0.55;
    const gap = 0.15;
    const startY = 1.3;

    sprints.forEach((sp, i) => {
        const y = startY + i * (rowH + gap);

        // Row background card
        drawCard(s, 0.5, y, 9.0, rowH);

        // Sprint name
        s.addText(sp.name, {
            x: 0.7, y: y + 0.08, w: 2.3, h: 0.22,
            fontSize: 12, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left"
        });

        // Time / Weeks
        s.addText(sp.time, {
            x: 0.7, y: y + 0.32, w: 2.3, h: 0.18,
            fontSize: 9.5, italic: true, color: COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "left"
        });

        // Green SP pill
        s.addShape("roundRect", {
            x: 3.2, y: y + 0.12, w: 0.9, h: 0.3,
            fill: { color: COLOR_MEDIUM_GREEN },
            line: { width: 0 },
            rectRadius: 0.2
        });
        s.addText(sp.pts, {
            x: 3.2, y: y + 0.12, w: 0.9, h: 0.3,
            fontSize: 10, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Custom Progress Bar
        const barX = 4.4;
        const barY = y + 0.22;
        const barW = 4.4;
        const barH = 0.15;

        // Gray base
        s.addShape("roundRect", {
            x: barX, y: barY, w: barW, h: barH,
            fill: { color: COLOR_LIGHT_GRAY },
            line: { width: 0 },
            rectRadius: 0.5
        });

        // Filled Portion (100% since project is completed)
        s.addShape("roundRect", {
            x: barX, y: barY, w: barW, h: barH,
            fill: { color: COLOR_PRIMARY_DARK },
            line: { width: 0 },
            rectRadius: 0.5
        });
    });

    // Agile Velocity Footer Card
    const fy = 4.95;
    s.addShape("roundRect", {
        x: 0.5, y: fy, w: 9.0, h: 0.45,
        fill: { color: "E8F5EE" },
        line: { color: COLOR_MEDIUM_GREEN, width: 0.5 },
        rectRadius: 0.08
    });
    s.addText("Cycle Projet Global : 9 Mois   •   Sprints de Développement : 10 Semaines (130 SP)", {
        x: 0.5, y: fy, w: 9.0, h: 0.45,
        fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 12 — STACK TECHNIQUE MODERNE
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Stack Technique Moderne", "Méthodologie & Architecture", "12/33");

    const layers = [
        { title: "FRONTEND MOBILE", content: "React Native (Expo SDK 54), TypeScript, Mapbox SDK, Socket.io client, Zustand." },
        { title: "FRONTEND WEB", content: "React.js, Vite, TailwindCSS (Console d'administration web)." },
        { title: "BACKEND API", content: "Node.js, Express.js (API modulaire), FastAPI (Python IA Pipeline), WebSockets." },
        { title: "DATA & INFRA", content: "PostgreSQL 16, Redis 7 (Cache / BullMQ queue), Docker, NGINX Reverse Proxy." }
    ];

    const cardW = 4.35;
    const cardH = 1.3;

    layers.forEach((l, i) => {
        const row = Math.floor(i / 2);
        const col = i % 2;
        const x = 0.5 + col * (cardW + 0.3);
        const y = 1.3 + row * (cardH + 0.2);

        drawCard(s, x, y, cardW, cardH);

        // Header label badge
        s.addShape("roundRect", {
            x: x + 0.2, y: y + 0.15, w: 1.8, h: 0.25,
            fill: { color: COLOR_BG_DARK },
            line: { width: 0 },
            rectRadius: 0.1
        });
        s.addText(l.title, {
            x: x + 0.2, y: y + 0.15, w: 1.8, h: 0.25,
            fontSize: 8.5, bold: true, color: COLOR_LIGHT_GREEN,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Content
        s.addText(l.content, {
            x: x + 0.2, y: y + 0.48, w: cardW - 0.4, h: 0.7,
            fontSize: 11.5, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "left", valign: "top"
        });
    });

    // Metric Badges (Bottom Row)
    const metrics = [
        { title: "Sécurité", detail: "JWT + Bcrypt + AES-256" },
        { title: "Performance", detail: "Redis caching < 200ms" },
        { title: "Qualité", detail: "147/151 tests Jest 97.4%" }
    ];

    metrics.forEach((m, i) => {
        const x = 0.5 + i * 3.1;
        const y = 4.35;
        const w = 2.8;
        const h = 0.8;

        drawCard(s, x, y, w, h);

        s.addText(m.title, {
            x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.25,
            fontSize: 13, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "center"
        });

        s.addText(m.detail, {
            x: x + 0.15, y: y + 0.38, w: w - 0.3, h: 0.35,
            fontSize: 11, color: COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "center"
        });
    });
}

// =============================================================================
// SLIDE 13 — ARCHITECTURE MULTI-SERVICES DISTRIBUÉE
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Architecture Multi-Services Distribuée", "Méthodologie & Architecture", "13/33");

    const archImgData = toBase64(IMG_ARCHITECTURE_DIAG);
    if (archImgData) {
        s.addImage({
            data: archImgData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}
// =============================================================================
// SLIDE 14 — WORKFLOW SIMPLIFIÉ KYC (Pipeline Horizontal Modernisé)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Workflow Simplifié KYC", "Méthodologie & Architecture", "14/33");

    // Introduction text
    s.addText("Pour garantir la sécurité et la conformité, chaque utilisateur suit un parcours de validation d'identité fluide et automatisé en 5 étapes clés.", {
        x: 0.5, y: 1.15, w: 9.0, h: 0.35,
        fontSize: 11, italic: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "left"
    });

    const steps = [
        { num: "1", title: "Candidat", desc: "Lancement du processus de certification depuis le profil utilisateur.", icon: "👤" },
        { num: "2", title: "Documents", desc: "Upload en direct des photos de la CNI (recto/verso) et d'un selfie.", icon: "📂" },
        { num: "3", title: "Analyse IA", desc: "OCR FastAPI, comparaison faciale ArcFace et test de vivacité (Liveness).", icon: "🤖" },
        { num: "4", title: "Décision", desc: "Validation automatique instantanée ou renvoi en modération manuelle.", icon: "⚖" },
        { num: "5", title: "Certifié", desc: "Badge de confiance vert octroyé pour un accès complet.", icon: "✓" }
    ];

    const cardW = 1.4;
    const cardH = 2.4;
    const gap = 0.45;
    const startX = 0.5;
    const ly = 1.7;

    steps.forEach((st, i) => {
        const cx = startX + i * (cardW + gap);
        const isLast = i === steps.length - 1;

        // Draw Card
        if (isLast) {
            // Highlighted final step card
            drawCard(s, cx, ly, cardW, cardH, {
                fill: "E8F5EE",
                line: { color: COLOR_MEDIUM_GREEN, width: 1.5 }
            });
        } else {
            // Standard step card
            drawCard(s, cx, ly, cardW, cardH);
        }

        // Left vertical indicator line inside card
        s.addShape("rect", {
            x: cx, y: ly + 0.1, w: 0.06, h: cardH - 0.2,
            fill: { color: isLast ? COLOR_MEDIUM_GREEN : COLOR_PRIMARY_DARK },
            line: { width: 0 }
        });

        // Step Number Badge (Circle)
        s.addShape("ellipse", {
            x: cx + 0.2, y: ly + 0.2, w: 0.32, h: 0.32,
            fill: { color: isLast ? COLOR_MEDIUM_GREEN : COLOR_PRIMARY_DARK },
            line: { width: 0 }
        });
        s.addText(st.num, {
            x: cx + 0.2, y: ly + 0.2, w: 0.32, h: 0.32,
            fontSize: 10, bold: true, color: isLast ? COLOR_GOLD : COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Icon Text (Unicode Symbol)
        s.addText(st.icon, {
            x: cx + 0.7, y: ly + 0.15, w: cardW - 0.8, h: 0.4,
            fontSize: 20, color: isLast ? COLOR_MEDIUM_GREEN : COLOR_PRIMARY_DARK,
            fontFace: FONT_BODY, align: "right", valign: "middle"
        });

        // Step Title
        s.addText(st.title, {
            x: cx + 0.2, y: ly + 0.65, w: cardW - 0.3, h: 0.25,
            fontSize: 11.5, bold: true, color: isLast ? COLOR_MEDIUM_GREEN : COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left"
        });

        // Step Description
        s.addText(st.desc, {
            x: cx + 0.2, y: ly + 0.95, w: cardW - 0.3, h: 1.3,
            fontSize: 8.5, color: isLast ? COLOR_PRIMARY_DARK : COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "left", valign: "top"
        });

        // Draw connecting arrow if not the last step
        if (!isLast) {
            const arrowX = cx + cardW + 0.05;
            s.addShape("line", {
                x: arrowX, y: ly + (cardH / 2), w: gap - 0.1, h: 0,
                line: { color: COLOR_GOLD, width: 2, endArrowType: "triangle" }
            });
        }
    });

    // Performance statistics badge at the bottom
    s.addShape("roundRect", {
        x: 0.5, y: 4.35, w: 9.0, h: 0.45,
        fill: { color: "E8F5EE" },
        line: { color: COLOR_MEDIUM_GREEN, width: 0.5 },
        rectRadius: 0.04
    });
    s.addText("Pipeline Optimisé : Temps moyen de traitement < 3.2 secondes  •  Taux de fausse acceptation (FAR) < 0.01%", {
        x: 0.5, y: 4.35, w: 9.0, h: 0.45,
        fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}
// =============================================================================
// SLIDE 15 — SECTION DIVIDER 04 (Conception & Modélisation)
// =============================================================================
buildSectionDivider("04", "Conception & Modélisation", "Quels sont les modèles UML ?", "04");

// =============================================================================
// SLIDE 16 — CAS D'UTILISATION GLOBAL (Affichage Diagramme Unique)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Cas d'Utilisation Global", "Conception & Modélisation", "16/33");

    const useCaseData = toBase64(IMG_USE_CASE);
    if (useCaseData) {
        s.addImage({
            data: useCaseData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}

// =============================================================================
// SLIDE 17 — SCHÉMA RELATIONNEL DE DONNÉES (Affichage Diagramme Unique)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Schéma Relationnel de Données", "Conception & Modélisation", "17/33");

    const classDiagData = toBase64(IMG_CLASS_DIAGRAM);
    if (classDiagData) {
        s.addImage({
            data: classDiagData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}

// =============================================================================
// SLIDE 18 — DIAGRAMME DE SÉQUENCE : RÉSERVATION (Booking) (Affichage Diagramme Unique)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Diagramme de Séquence : Réservation", "Conception & Modélisation", "18/33");

    const bookingSeqData = toBase64(IMG_SEQ_BOOKING);
    if (bookingSeqData) {
        s.addImage({
            data: bookingSeqData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}

// =============================================================================
// SLIDE 19 — DIAGRAMME DE SÉQUENCE : VALIDATION KYC (UML) (Affichage Diagramme Unique)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Diagramme de Séquence : Validation KYC", "Conception & Modélisation", "19/33");

    const kycSeqData = toBase64(IMG_SEQ_KYC);
    if (kycSeqData) {
        s.addImage({
            data: kycSeqData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}

// =============================================================================
// SLIDE 20 — DIAGRAMME DE SÉQUENCE : EMBARQUEMENT BIOMÉTRIQUE (UML) (Affichage Diagramme Unique)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Diagramme de Séquence : Embarquement", "Conception & Modélisation", "20/33");

    const checkinSeqData = toBase64(IMG_SEQ_CHECKIN);
    if (checkinSeqData) {
        s.addImage({
            data: checkinSeqData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}

// =============================================================================
// SLIDE 21 — TOPOLOGIE DU DÉPLOIMENT CLOUD (Split layout avec diagramme réel)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Topologie du Déploiement Cloud", "Conception & Modélisation", "21/33");

    const deployImgData = toBase64(IMG_DEPLOYMENT_CLOUD);
    if (deployImgData) {
        s.addImage({
            data: deployImgData,
            x: 0.5, y: 1.25,
            w: 9.0, h: 3.9,
            sizing: { type: "contain", w: 9.0, h: 3.9 }
        });
    }
}
// =============================================================================
// SLIDE 19 — SECTION DIVIDER 05 (Résultats & Validation)
// =============================================================================
buildSectionDivider("05", "Résultats & Validation", "Qu'avons-nous obtenu ?", "05");

// =============================================================================
// SLIDE 21 — 100% DE RÉUSSITE AUX TESTS AUTOMATISÉS (Split layout avec capture réelle)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "100% de Réussite aux Tests Automatisés", "Résultats & Validation", "23/33");

    // Left Column: Donut, Progress bars and Coverage stats
    const lx = 0.5;
    const ly = 1.3;
    const lw = 4.2;
    const lh = 3.9;
    drawCard(s, lx, ly, lw, lh);
    s.addShape("rect", { x: lx, y: ly + 0.1, w: 0.08, h: lh - 0.2, fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 } });

    s.addText("Couverture des Tests", {
        x: lx + 0.2, y: ly + 0.15, w: lw - 0.3, h: 0.35,
        fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "left"
    });

    // Donut stat simulation
    const donutX = lx + 0.15;
    const donutY = ly + 0.6;
    const donutD = 1.0;
    s.addShape("ellipse", {
        x: donutX, y: donutY, w: donutD, h: donutD,
        fill: { color: COLOR_BG_LIGHT },
        line: { color: COLOR_PRIMARY_DARK, width: 3.5 }
    });
    s.addText("373/373", {
        x: donutX, y: donutY + 0.25, w: donutD, h: 0.3,
        fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "center", valign: "middle"
    });
    s.addText("100% Succès", {
        x: donutX, y: donutY + 0.55, w: donutD, h: 0.2,
        fontSize: 8, bold: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "center"
    });

    // Right of Donut: 4 Progress bars
    const barX = lx + 1.35;
    const barW = lw - 1.5;
    const barH = 0.06;

    const progressBars = [
        { label: "Statements (75.4%)", val: 0.754, color: COLOR_PRIMARY_DARK },
        { label: "Functions (77.8%)", val: 0.778, color: COLOR_PRIMARY_DARK },
        { label: "Lines (77.3%)", val: 0.773, color: COLOR_PRIMARY_DARK },
        { label: "Branches (63.9%)", val: 0.639, color: COLOR_GOLD }
    ];

    progressBars.forEach((bar, i) => {
        const by = ly + 0.55 + i * 0.58;

        s.addText(bar.label, { x: barX, y: by, w: barW, h: 0.18, fontSize: 8.5, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_BODY });
        s.addShape("roundRect", { x: barX, y: by + 0.2, w: barW, h: barH, fill: { color: COLOR_LIGHT_GRAY }, line: { width: 0 }, rectRadius: 0.5 });
        s.addShape("roundRect", { x: barX, y: by + 0.2, w: barW * bar.val, h: barH, fill: { color: bar.color }, line: { width: 0 }, rectRadius: 0.5 });
    });

    // Stats Grid inside left card
    const cats = [
        { label: "Routes API", val: "98.7% Couv." },
        { label: "Middlewares", val: "89.0% Couv." },
        { label: "Classes Utils", val: "84.6% Couv." },
        { label: "Contrôleurs", val: "71.8% Couv." }
    ];

    cats.forEach((cat, i) => {
        const cx = lx + 0.15 + (i % 2) * 2.0;
        const cy = ly + 2.9 + Math.floor(i / 2) * 0.45;
        const cw = 1.9;
        const ch = 0.4;

        s.addShape("roundRect", { x: cx, y: cy, w: cw, h: ch, fill: { color: COLOR_BG_LIGHT }, line: { color: "E7ECE8", width: 0.5 }, rectRadius: 0.05 });
        s.addText(`${cat.label} : ${cat.val}`, {
            x: cx + 0.05, y: cy, w: cw - 0.1, h: ch,
            fontSize: 8, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });
    });

    // Campaign execution stat footer on left card
    s.addShape("roundRect", { x: lx + 0.15, y: ly + 3.45, w: lw - 0.3, h: 0.35, fill: { color: "E8F5EE" }, line: { width: 0 }, rectRadius: 0.04 });
    s.addText("Campagne de 31 suites exécutée avec succès en 50.439s", {
        x: lx + 0.15, y: ly + 3.45, w: lw - 0.3, h: 0.35,
        fontSize: 8.5, italic: true, color: COLOR_MEDIUM_GREEN,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });

    // Right Column: Terminal screenshot card
    const rx = 4.9;
    const rw = 4.6;
    drawCard(s, rx, ly, rw, lh);
    s.addText("Console de Rendu Jest (Terminal)", {
        x: rx + 0.2, y: ly + 0.15, w: rw - 0.4, h: 0.3,
        fontSize: 12, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "left"
    });

    const jestData = toBase64(IMG_JEST_RESULTS);
    if (jestData) {
        s.addImage({ data: jestData, x: rx + 0.15, y: ly + 0.55, w: rw - 0.3, h: lh - 0.75 });
    }
}

// =============================================================================
// SLIDE 21 — SCORE SUS DE 71.6 : USABILITÉ VALIDÉE
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Score SUS de 71.6 : Usabilité Validée", "Résultats & Validation", "24/33");

    // Score Circle Card (Left)
    drawCard(s, 0.6, 1.4, 2.2, 2.5);

    s.addShape("ellipse", {
        x: 0.8, y: 1.6, w: 1.8, h: 1.8,
        fill: { color: COLOR_BG_LIGHT },
        line: { color: COLOR_PRIMARY_DARK, width: 3 }
    });

    s.addText("71.6", {
        x: 0.8, y: 1.9, w: 1.8, h: 0.5,
        fontSize: 32, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "center", valign: "middle"
    });
    s.addText("/ 100", {
        x: 0.8, y: 2.3, w: 1.8, h: 0.3,
        fontSize: 12, bold: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "center"
    });

    // GOOD Badge pill
    s.addShape("roundRect", {
        x: 1.1, y: 3.4, w: 1.2, h: 0.32,
        fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 }, rectRadius: 0.2
    });
    s.addText("GOOD", {
        x: 1.1, y: 3.4, w: 1.2, h: 0.32,
        fontSize: 10, bold: true, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });

    // SUS Scale Bar (Right Top)
    const rightX = 3.1;
    s.addText("Échelle de Score d'Usabilité (System Usability Scale)", {
        x: rightX, y: 1.2, w: 5.5, h: 0.25,
        fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY
    });

    const scaleY = 1.55;
    const scaleW = 5.8;
    const scaleH = 0.2;

    // Scale Segments
    // 0-50 (Poor - Red)
    s.addShape("rect", { x: rightX, y: scaleY, w: scaleW * 0.5, h: scaleH, fill: { color: COLOR_RED }, line: { width: 0 } });
    // 50-68 (OK - Orange)
    s.addShape("rect", { x: rightX + scaleW * 0.5, y: scaleY, w: scaleW * 0.18, h: scaleH, fill: { color: COLOR_GOLD }, line: { width: 0 } });
    // 68-80 (Good - Light Green)
    s.addShape("rect", { x: rightX + scaleW * 0.68, y: scaleY, w: scaleW * 0.12, h: scaleH, fill: { color: COLOR_LIGHT_GREEN }, line: { width: 0 } });
    // 80-100 (Excellent - Dark Green)
    s.addShape("rect", { x: rightX + scaleW * 0.8, y: scaleY, w: scaleW * 0.2, h: scaleH, fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 } });

    // Indicator Marker at 71.6%
    const markerX = rightX + scaleW * 0.716;
    s.addShape("rect", { x: markerX - 0.03, y: scaleY - 0.15, w: 0.06, h: 0.5, fill: { color: COLOR_DARK_TEXT }, line: { width: 0 } });
    s.addText("▲\n71.6", {
        x: markerX - 0.4, y: scaleY - 0.55, w: 0.8, h: 0.45,
        fontSize: 10, bold: true, color: COLOR_DARK_TEXT,
        fontFace: FONT_BODY, align: "center"
    });

    // Profile Bars
    s.addText("Score Passagers (N=10)", { x: rightX, y: 2.3, w: 3.0, h: 0.22, fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_BODY });
    s.addShape("roundRect", { x: rightX, y: 2.55, w: 5.8, h: 0.2, fill: { color: COLOR_LIGHT_GRAY }, line: { width: 0 }, rectRadius: 0.5 });
    s.addShape("roundRect", { x: rightX, y: 2.55, w: 5.8 * 0.742, h: 0.2, fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 }, rectRadius: 0.5 });
    s.addText("74.2 / 100", { x: rightX + 4.6, y: 2.3, w: 1.2, h: 0.22, fontSize: 11, bold: true, color: COLOR_MEDIUM_GREEN, fontFace: FONT_BODY, align: "right" });

    s.addText("Score Conducteurs (N=8)", { x: rightX, y: 2.9, w: 3.0, h: 0.22, fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_BODY });
    s.addShape("roundRect", { x: rightX, y: 3.15, w: 5.8, h: 0.2, fill: { color: COLOR_LIGHT_GRAY }, line: { width: 0 }, rectRadius: 0.5 });
    s.addShape("roundRect", { x: rightX, y: 3.15, w: 5.8 * 0.689, h: 0.2, fill: { color: COLOR_LIGHT_GREEN }, line: { width: 0 }, rectRadius: 0.5 });
    s.addText("68.9 / 100", { x: rightX + 4.6, y: 2.9, w: 1.2, h: 0.22, fontSize: 11, bold: true, color: COLOR_LIGHT_GREEN, fontFace: FONT_BODY, align: "right" });

    // Footer note
    const fy = 4.35;
    s.addShape("roundRect", {
        x: 0.5, y: fy, w: 9.0, h: 0.6,
        fill: { color: "E8F5EE" },
        line: { color: COLOR_MEDIUM_GREEN, width: 0.5 },
        rectRadius: 0.08
    });
    s.addText("→ Écart Passager / Conducteur (5.3 pts) : le flux KYC conducteur (Documents + Véhicule) est plus complexe • Priorité UX pour la v2", {
        x: 0.5, y: fy, w: 9.0, h: 0.6,
        fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 22 — VALIDATION DES TROIS HYPOTHÈSES
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Validation des Trois Hypothèses", "Résultats & Validation", "25/33");

    const hyps = [
        {
            num: "H1",
            title: "KYC Biométrique",
            crit: "Seuil de décision discriminants par rôle",
            res: "Non quantifié sur bases de référence",
            badge: "⚠ PARTIELLE",
            color: COLOR_GOLD
        },
        {
            num: "H2",
            title: "Usabilité SUS",
            crit: "Score SUS global ≥ 68 / 100",
            res: "Empirique: 71.6 / 100 (Good)",
            badge: "✓ VALIDÉE",
            color: COLOR_MEDIUM_GREEN
        },
        {
            num: "H3",
            title: "API Résilience",
            crit: "Taux d'exécution transactions API ≥ 99%",
            res: "99.9% de succès (simulation de charge)",
            badge: "✓ VALIDÉE",
            color: COLOR_MEDIUM_GREEN
        }
    ];

    const cardW = 5.5;
    const cardH = 1.05;

    hyps.forEach((h, i) => {
        const y = 1.25 + i * 1.2;
        drawCard(s, 0.5, y, cardW, cardH);

        // H-number badge
        s.addShape("roundRect", {
            x: 0.7, y: y + 0.22, w: 0.5, h: 0.5,
            fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 }, rectRadius: 0.1
        });
        s.addText(h.num, {
            x: 0.7, y: y + 0.22, w: 0.5, h: 0.5,
            fontSize: 16, bold: true, color: COLOR_GOLD,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Title
        s.addText(h.title, {
            x: 1.35, y: y + 0.1, w: 2.5, h: 0.3,
            fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE
        });

        // Criterion & Results
        s.addText(`Critère: ${h.crit}\nRésultat: ${h.res}`, {
            x: 1.35, y: y + 0.4, w: 3.9, h: 0.55,
            fontSize: 10.5, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "left"
        });

        // Status Badge Pill
        s.addShape("roundRect", {
            x: 4.4, y: y + 0.15, w: 1.4, h: 0.26,
            fill: { color: h.color }, line: { width: 0 }, rectRadius: 0.2
        });
        s.addText(h.badge, {
            x: 4.4, y: y + 0.15, w: 1.4, h: 0.26,
            fontSize: 9, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });
    });

    // Verdict box on the right
    const vy = 1.25;
    const vw = 3.2;
    const vh = 3.45;
    drawCard(s, 6.3, vy, vw, vh, {
        fill: COLOR_BG_LIGHT,
        line: { color: COLOR_PRIMARY_DARK, width: 1.5 }
    });

    s.addText("VERDICT FINAL", {
        x: 6.3, y: vy + 0.2, w: vw, h: 0.35,
        fontSize: 16, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE, align: "center"
    });

    s.addText("• H1 (KYC) : ⚠ PARTIELLE\n  Manque d'évaluation quantitative sur base de référence.\n\n• H2 (SUS) : ✓ VALIDÉE\n  Usabilité validée comme \"Good\" (71.6/100).\n\n• H3 (API) : ✓ VALIDÉE\n  99.9% de succès transactionnel sous charge intermittente.", {
        x: 6.5, y: vy + 0.7, w: vw - 0.4, h: 2.5,
        fontSize: 11.5, bold: true, color: COLOR_DARK_TEXT,
        fontFace: FONT_BODY, align: "left"
    });

    // Footer
    s.addText("H2 & H3 pleinement validées  •  H1 partiellement validée (évaluation quantitative de l'IA à approfondir)", {
        x: 0.5, y: 4.85, w: 9.0, h: 0.3,
        fontSize: 11, italic: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "center"
    });
}

// =============================================================================
// SLIDE 23 — DÉMONSTRATION VIDÉO DU PROTOTYPE (Redesign Premium)
// =============================================================================
{
    const s = pres.addSlide();
    s.background = { color: COLOR_BG_DARK };

    // ── HEADER ────────────────────────────────────────────────────────────────
    s.addText("RÉSULTATS & VALIDATION", {
        x: 0.5, y: 0.18, w: 7.0, h: 0.22,
        fontSize: 9, bold: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "left"
    });
    s.addText("26/33", {
        x: 8.5, y: 0.18, w: 1.0, h: 0.22,
        fontSize: 10, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "right"
    });

    // Nouveau titre principal
    s.addText("Démonstration Vidéo du Prototype", {
        x: 0.4, y: 0.38, w: 9.2, h: 0.52,
        fontSize: 28, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "left"
    });

    // Sous-titre
    s.addText("Parcours complet : de l'inscription à la réservation sécurisée", {
        x: 0.4, y: 0.90, w: 9.2, h: 0.28,
        fontSize: 13, italic: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "left"
    });

    // Ligne séparatrice sous le header
    s.addShape("rect", {
        x: 0.4, y: 1.22, w: 9.2, h: 0.025,
        fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 }
    });

    // ── 4 BADGES FONCTIONNALITÉS (bande au-dessus de la vidéo) ───────────────
    const badges = [
        { icon: "APP", label: "Application Mobile" },
        { icon: "KYC", label: "KYC Biometrique"    },
        { icon: "PAY", label: "Chargily Pay V2"  },
        { icon: "GPS", label: "GPS Temps Reel"      },
    ];
    const badgeW   = 1.52;
    const badgeGap = 0.12;
    const badgeStartX = 0.4;

    badges.forEach((b, i) => {
        const bx = badgeStartX + i * (badgeW + badgeGap);
        // Fond du badge
        s.addShape("roundRect", {
            x: bx, y: 1.30, w: badgeW, h: 0.38,
            fill: { color: COLOR_MEDIUM_GREEN },
            line: { color: COLOR_LIGHT_GREEN, width: 0.5 },
            rectRadius: 0.06
        });
        // Icône abrégée
        s.addShape("roundRect", {
            x: bx + 0.07, y: 1.35, w: 0.32, h: 0.27,
            fill: { color: COLOR_PRIMARY_DARK },
            line: { width: 0 }, rectRadius: 0.04
        });
        s.addText(b.icon, {
            x: bx + 0.07, y: 1.35, w: 0.32, h: 0.27,
            fontSize: 7, bold: true, color: COLOR_GOLD,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });
        // Label
        s.addText(b.label, {
            x: bx + 0.44, y: 1.35, w: badgeW - 0.50, h: 0.27,
            fontSize: 10, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "left", valign: "middle"
        });
    });

    // ── ZONE VIDÉO GAUCHE (70 % ≈ 6.8 inch) ──────────────────────────────────
    const VX  = 0.4;
    const VY  = 1.74;
    const VW2 = 6.6;
    const VH2 = 3.45;

    // Fond de la zone vidéo (carte sombre)
    s.addShape("roundRect", {
        x: VX, y: VY, w: VW2, h: VH2,
        fill: { color: "0A1F14" },
        line: { color: COLOR_MEDIUM_GREEN, width: 1.2 },
        rectRadius: 0.14,
        shadow: makeShadow()
    });

    // Embed de la vidéo si disponible
    const VIDEO_PATH = path.join(__dirname, "RohWinBghit_Demo.mp4");
    const SPLASH_FALLBACK = fs.existsSync(IMG_SPLASH)
        ? IMG_SPLASH
        : path.join(__dirname, "Screen_Mobile", "01_landing_onboarding.png"); // Real fallback
    if (fs.existsSync(VIDEO_PATH)) {
        s.addMedia({
            type: "video",
            path: VIDEO_PATH,
            x: VX + 0.08,
            y: VY + 0.08,
            w: VW2 - 0.16,
            h: VH2 - 0.16,
        });
    } else {
        // Fallback : mockup iPhone premium avec splash screen
        // Fond interne simulant l'écran du téléphone
        s.addShape("roundRect", {
            x: VX + 0.08, y: VY + 0.08, w: VW2 - 0.16, h: VH2 - 0.16,
            fill: { color: "0F2D1E" }, line: { width: 0 }, rectRadius: 0.10
        });

        // Cadre iPhone centré
        const iPhW = 2.1, iPhH = 3.1;
        const iPhX = VX + (VW2 - iPhW) / 2;
        const iPhY = VY + (VH2 - iPhH) / 2;

        s.addShape("roundRect", {
            x: iPhX - 0.12, y: iPhY - 0.15, w: iPhW + 0.24, h: iPhH + 0.30,
            fill: { color: "0A2518" },
            line: { color: COLOR_MEDIUM_GREEN, width: 1.8 },
            rectRadius: 0.28
        });
        // Notch
        s.addShape("roundRect", {
            x: iPhX + iPhW / 2 - 0.25, y: iPhY - 0.14, w: 0.50, h: 0.14,
            fill: { color: "050F0A" }, line: { width: 0 }, rectRadius: 0.07
        });
        // Écran : splash screen
        const splashData = toBase64(SPLASH_FALLBACK);
        if (splashData) {
            s.addImage({ data: splashData, x: iPhX, y: iPhY, w: iPhW, h: iPhH });
        } else {
            s.addShape("roundRect", {
                x: iPhX, y: iPhY, w: iPhW, h: iPhH,
                fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 }, rectRadius: 0.18
            });
        }
        // Bouton Play 3D (cercle double)
        const playX = VX + VW2 / 2 - 0.38;
        const playY = VY + VH2 / 2 - 0.38;
        s.addShape("ellipse", {
            x: playX - 0.08, y: playY - 0.08, w: 0.92, h: 0.92,
            fill: { color: COLOR_WHITE }, line: { color: COLOR_LIGHT_GREEN, width: 0 },
            shadow: { type: "outer", color: "000000", blur: 12, offset: 4, angle: 45, opacity: 0.5 }
        });
        s.addShape("ellipse", {
            x: playX, y: playY, w: 0.76, h: 0.76,
            fill: { color: COLOR_PRIMARY_DARK }, line: { color: COLOR_LIGHT_GREEN, width: 1.5 }
        });
        s.addText("▶", {
            x: playX + 0.04, y: playY, w: 0.76, h: 0.76,
            fontSize: 22, bold: true, color: COLOR_WHITE,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });
    }

    // Barre de progression verte sous la vidéo
    s.addShape("roundRect", {
        x: VX, y: VY + VH2 + 0.08, w: VW2, h: 0.10,
        fill: { color: "0D2B1C" }, line: { width: 0 }, rectRadius: 0.05
    });
    s.addShape("roundRect", {
        x: VX, y: VY + VH2 + 0.08, w: VW2 * 0.62, h: 0.10,
        fill: { color: "1C4933" }, line: { width: 0 }, rectRadius: 0.05
    });
    s.addShape("ellipse", {
        x: VX + VW2 * 0.62 - 0.09, y: VY + VH2 + 0.035, w: 0.18, h: 0.18,
        fill: { color: COLOR_LIGHT_GREEN }, line: { width: 0 }
    });

    // Durée de la vidéo
    s.addText("01:19", {
        x: VX + VW2 - 0.65, y: VY + VH2 + 0.22, w: 0.6, h: 0.22,
        fontSize: 9, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY, align: "right"
    });

    // ── PANNEAU DROIT (30 %) — Fonctionnalités clés ───────────────────────────
    const RX  = VX + VW2 + 0.18;
    const RW  = SW - RX - 0.15;

    // Titre panneau
    s.addText("Fonctionnalités Clés", {
        x: RX, y: VY - 0.02, w: RW, h: 0.30,
        fontSize: 13, bold: true, color: COLOR_GOLD,
        fontFace: FONT_TITLE, align: "left"
    });

    // Liste des fonctionnalités
    const features = [
        { check: true,  label: "Inscription & Auth",    sub: "OTP + Rôle (Passager/Chauffeur)" },
        { check: true,  label: "KYC Biometrique",       sub: "OCR + ArcFace + Vivacite" },
        { check: true,  label: "Publication Trajet",    sub: "Carte interactive + prix auto" },
        { check: true,  label: "Reservation Passager",  sub: "Recherche + detail + paiement" },
        { check: true,  label: "Billet & QR Code",      sub: "Boarding Pass + scan depart" },
        { check: true,  label: "Suivi GPS Temps Reel",  sub: "WebSocket + Mapbox SDK" },
    ];

    const fY0 = VY + 0.32;
    const fRowH = 0.52;

    features.forEach((f, i) => {
        const fy2 = fY0 + i * fRowH;

        // Fond de ligne alternée
        if (i % 2 === 0) {
            s.addShape("roundRect", {
                x: RX - 0.06, y: fy2 - 0.04, w: RW + 0.06, h: fRowH - 0.04,
                fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 },
                rectRadius: 0.06, transparency: 75
            });
        }

        // Check circle
        s.addShape("ellipse", {
            x: RX, y: fy2 + 0.04, w: 0.24, h: 0.24,
            fill: { color: f.check ? COLOR_LIGHT_GREEN : "555555" },
            line: { width: 0 }
        });
        s.addText(f.check ? "✓" : "–", {
            x: RX, y: fy2 + 0.04, w: 0.24, h: 0.24,
            fontSize: 9, bold: true, color: COLOR_BG_DARK,
            fontFace: FONT_BODY, align: "center", valign: "middle"
        });

        // Label
        s.addText(f.label, {
            x: RX + 0.30, y: fy2 + 0.01, w: RW - 0.32, h: 0.20,
            fontSize: 11, bold: true, color: COLOR_WHITE,
            fontFace: FONT_TITLE, align: "left"
        });

        // Sous-label
        s.addText(f.sub, {
            x: RX + 0.30, y: fy2 + 0.22, w: RW - 0.32, h: 0.18,
            fontSize: 9, italic: true, color: COLOR_MUTED_TEXT,
            fontFace: FONT_BODY, align: "left"
        });
    });

    // Bloc script de présentation
    const scriptY = fY0 + features.length * fRowH + 0.08;
    s.addShape("roundRect", {
        x: RX - 0.06, y: scriptY, w: RW + 0.06, h: 1.00,
        fill: { color: COLOR_PRIMARY_DARK },
        line: { color: COLOR_MEDIUM_GREEN, width: 0.75 },
        rectRadius: 0.08
    });
    s.addText("Script :", {
        x: RX + 0.06, y: scriptY + 0.06, w: RW - 0.08, h: 0.18,
        fontSize: 8.5, bold: true, color: COLOR_GOLD,
        fontFace: FONT_BODY, align: "left"
    });
    s.addText("\"Cette video presente le parcours complet : de l'inscription et la verification KYC jusqu'a la reservation et le QR Code de depart.\"", {
        x: RX + 0.06, y: scriptY + 0.24, w: RW - 0.08, h: 0.70,
        fontSize: 8.5, italic: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "left", valign: "top"
    });

    // ── FOOTER ────────────────────────────────────────────────────────────────
    s.addShape("rect", {
        x: 0, y: 5.30, w: SW, h: 0.325,
        fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 }
    });
    s.addText("RohWinBghit  •  Parcours Passager Complet  •  Verification KYC + Chargily Pay V2 + Suivi GPS", {
        x: 0, y: 5.30, w: SW, h: 0.325,
        fontSize: 10, bold: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 24 — SECTION DIVIDER 06 (Modèle Économique)
// =============================================================================
buildSectionDivider("06", "Modèle Économique", "Comment rentabiliser le projet ?", "06");

// =============================================================================
// SLIDE 25 — BUSINESS MODEL CANVAS (CONDENSÉ)
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Business Model Canvas (condensé)", "Modèle Économique", "28/33");

    s.addText("Business Model & Financials", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.3,
        fontSize: 13, italic: true, color: COLOR_MUTED_TEXT,
        fontFace: FONT_BODY
    });

    // 6-block BMC Grid
    const by = 1.35;
    const bh = 1.95;

    // Partenaires Clés
    drawCard(s, 0.5, by, 1.7, bh);
    s.addText("Partenaires Clés", { x: 0.5, y: by + 0.1, w: 1.7, h: 0.22, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Chargily Pay V2\n• Universités DZ\n• Assurances locales\n• Hébergeurs Cloud", { x: 0.6, y: by + 0.35, w: 1.5, h: bh - 0.45, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Activités & Ressources
    drawCard(s, 2.3, by, 1.7, 0.95);
    s.addText("Activités Clés", { x: 2.3, y: by + 0.08, w: 1.7, h: 0.2, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Matching trajets\n• Pipeline KYC\n• Séquestre Chargily", { x: 2.4, y: by + 0.3, w: 1.5, h: 0.6, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    drawCard(s, 2.3, by + 1.05, 1.7, 0.9);
    s.addText("Ressources Clés", { x: 2.3, y: by + 1.1, w: 1.7, h: 0.2, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Code proprietary\n• Serveur VPS local\n• Modèles IA", { x: 2.4, y: by + 1.32, w: 1.5, h: 0.6, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Propositions de Valeur
    drawCard(s, 4.1, by, 1.7, bh, { fill: "E8F5EE", line: { color: COLOR_PRIMARY_DARK, width: 1 } });
    s.addText("Propositions de Valeur", { x: 4.1, y: by + 0.1, w: 1.7, h: 0.22, fontSize: 10.5, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Covoiturage KYC\n• 69 wilayas\n• Paiement DZD\n• Tarifs réduits\n• Billet QR Code", { x: 4.2, y: by + 0.35, w: 1.5, h: bh - 0.45, fontSize: 9.5, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_BODY });

    // Relations & Canaux
    drawCard(s, 5.9, by, 1.7, 0.95);
    s.addText("Relations Clients", { x: 5.9, y: by + 0.08, w: 1.7, h: 0.2, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Support client\n• Notation mutuelle\n• Notifications push", { x: 6.0, y: by + 0.3, w: 1.5, h: 0.6, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    drawCard(s, 5.9, by + 1.05, 1.7, 0.9);
    s.addText("Canaux", { x: 5.9, y: by + 1.1, w: 1.7, h: 0.2, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Stores d'applications\n• Campus univ.\n• Réseaux sociaux", { x: 6.0, y: by + 1.32, w: 1.5, h: 0.6, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Segments Clients
    drawCard(s, 7.7, by, 1.8, bh);
    s.addText("Segments Clients", { x: 7.7, y: by + 0.1, w: 1.8, h: 0.22, fontSize: 10, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Étudiants (68%)\n• Pendulaires réguliers\n• Chauffeurs (frais)\n• Familles", { x: 7.8, y: by + 0.35, w: 1.6, h: bh - 0.45, fontSize: 9.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Metric groups (Below BMC)
    const my = 3.45;
    const mw = 2.8;
    const mh = 1.35;

    // Modèle de Revenus
    drawCard(s, 0.5, my, mw, mh);
    s.addText("Modèle de Revenus", { x: 0.5, y: my + 0.1, w: mw, h: 0.25, fontSize: 12, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Commission 12% sur billets\n• Premium Driver : 500 DA / mois\n• Partenariats B2B (Publicités)", { x: 0.65, y: my + 0.38, w: mw - 0.3, h: 0.85, fontSize: 10, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Année 1 Chiffres clés
    drawCard(s, 3.6, my, mw, mh);
    s.addText("Chiffres Clés (Année 1)", { x: 3.6, y: my + 0.1, w: mw, h: 0.25, fontSize: 12, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Volume transactionnel : 15.5M DA\n• Chiffre d'Affaires Net : 6.7M DA\n• Point mort atteint à 18 mois", { x: 3.75, y: my + 0.38, w: mw - 0.3, h: 0.85, fontSize: 10, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Métriques clés
    drawCard(s, 6.7, my, mw, mh);
    s.addText("Métriques de Traction", { x: 6.7, y: my + 0.1, w: mw, h: 0.25, fontSize: 12, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE, align: "center" });
    s.addText("• Utilisateurs visés : 50 000+\n• Taux de conversion client : 15%\n• Score de satisfaction : 4.2 / 5", { x: 6.85, y: my + 0.38, w: mw - 0.3, h: 0.85, fontSize: 10, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });
}

// =============================================================================
// SLIDE 26 — MODÈLE DE REVENUS & PRÉVISIONS FINANCIÈRES
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Modèle de Revenus & Prévisions Financières", "Modèle Économique", "29/33");

    // Left column: 3 cards
    const cardW = 4.0;
    const cardH = 1.1;

    // Card 1
    drawCard(s, 0.5, 1.2, cardW, cardH);
    s.addText("Commission 12%", { x: 0.7, y: 1.25, w: 2.2, h: 0.28, fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE });
    // Green Pill
    s.addShape("roundRect", { x: 2.9, y: 1.25, w: 1.4, h: 0.26, fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 }, rectRadius: 0.2 });
    s.addText("PRINCIPAL", { x: 2.9, y: 1.25, w: 1.4, h: 0.26, fontSize: 9, bold: true, color: COLOR_WHITE, fontFace: FONT_BODY, align: "center", valign: "middle" });
    s.addText("Prélevée sur chaque trajet longue distance réservé et payé en ligne via notre passerelle sécurisée.", { x: 0.7, y: 1.58, w: 3.6, h: 0.65, fontSize: 10.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Card 2
    drawCard(s, 0.5, 2.45, cardW, cardH);
    s.addText("Premium Driver", { x: 0.7, y: 2.5, w: 2.2, h: 0.28, fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE });
    // Teal Pill
    s.addShape("roundRect", { x: 2.9, y: 2.5, w: 1.4, h: 0.26, fill: { color: "028090" }, line: { width: 0 }, rectRadius: 0.2 });
    s.addText("500 DA/MOIS", { x: 2.9, y: 2.5, w: 1.4, h: 0.26, fontSize: 9, bold: true, color: COLOR_WHITE, fontFace: FONT_BODY, align: "center", valign: "middle" });
    s.addText("Abonnement mensuel offrant aux chauffeurs une visibilité accrue et un support prioritaire.", { x: 0.7, y: 2.83, w: 3.6, h: 0.65, fontSize: 10.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Card 3
    drawCard(s, 0.5, 3.7, cardW, cardH);
    s.addText("Partenariats B2B", { x: 0.7, y: 3.75, w: 2.2, h: 0.28, fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE });
    // Gold Pill
    s.addShape("roundRect", { x: 2.9, y: 3.75, w: 1.4, h: 0.26, fill: { color: COLOR_GOLD }, line: { width: 0 }, rectRadius: 0.2 });
    s.addText("PARTENARIAT", { x: 2.9, y: 3.75, w: 1.4, h: 0.26, fontSize: 9, bold: true, color: COLOR_BG_DARK, fontFace: FONT_BODY, align: "center", valign: "middle" });
    s.addText("Publicités ciblées pour stations-services, assurances automobiles et opérateurs télécoms.", { x: 0.7, y: 4.08, w: 3.6, h: 0.65, fontSize: 10.5, color: COLOR_MUTED_TEXT, fontFace: FONT_BODY });

    // Right side native table: Projections on 3 Years
    s.addText("Prévisions de croissance sur 3 ans", { x: 4.8, y: 1.2, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: COLOR_PRIMARY_DARK, fontFace: FONT_TITLE });

    const forecastRows = [
        [
            { text: "Indicateur", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Année 1", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Année 2", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } },
            { text: "Année 3", options: { bold: true, color: COLOR_WHITE, fill: { color: COLOR_PRIMARY_DARK } } }
        ],
        [
            { text: "Utilisateurs actifs" },
            { text: "5 000" },
            { text: "25 000" },
            { text: "80 000" }
        ],
        [
            { text: "Trajets / mois" },
            { text: "2 000" },
            { text: "12 000" },
            { text: "40 000" }
        ],
        [
            { text: "CA Estimé (DA)" },
            { text: "3.6 Millions" },
            { text: "21.6 Millions" },
            { text: "72.0 Millions" }
        ],
        [
            { text: "Seuil de Rentabilité" },
            { text: "—" },
            { text: "✓ Atteint", options: { color: COLOR_PRIMARY_DARK, bold: true } },
            { text: "—" }
        ]
    ];

    s.addTable(forecastRows, {
        x: 4.8, y: 1.65, w: 4.7, h: 2.8,
        border: { type: "solid", color: "E7ECE8", size: 1 },
        align: "center",
        valign: "middle",
        fontFace: FONT_BODY,
        fontSize: 11
    });

    // Summary Footer note
    s.addText("→ Seuil de rentabilité estimé à 18 mois  •  ROI positif dès l'Année 2 de commercialisation", {
        x: 4.8, y: 4.6, w: 4.7, h: 0.35,
        fontSize: 11, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_BODY, align: "center"
    });
}

// =============================================================================
// SLIDE 27 — SECTION DIVIDER 07 (Conclusion & Perspectives)
// =============================================================================
buildSectionDivider("07", "Conclusion & Perspectives", "Quel bilan et quelles évolutions ?", "07");

// =============================================================================
// SLIDE 28 — BILAN & CONTRIBUTIONS ACADÉMIQUES
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Bilan & Contributions Académiques", "Conclusion & Perspectives", "31/33");

    const contributions = [
        { title: "Infrastructure de Confiance", desc: "Premier pipeline KYC biométrique d'identification (OCR + ArcFace + Liveness) adapté aux spécificités administratives algériennes." },
        { title: "IA Souveraine", desc: "Modèles d'IA exécutant les inférences localement sur serveurs nationaux, respectant la Loi 25-11 sur la non-transfrontalisation." },
        { title: "Monétique Locale", desc: "Première intégration de séquestre monétaire via Chargily Pay V2 (CIB/Edahabia) dédiée à la sécurisation des covoiturages inter-wilayas." },
        { title: "Validation Empirique", desc: "97.4% de réussite sur la suite de tests Jest automatisés et score SUS de 71.6/100 validant la fluidité de l'expérience utilisateur." }
    ];

    const cardW = 4.35;
    const cardH = 1.6;

    contributions.forEach((c, i) => {
        const row = Math.floor(i / 2);
        const col = i % 2;
        const x = 0.5 + col * (cardW + 0.3);
        const y = 1.4 + row * (cardH + 0.3);

        drawCard(s, x, y, cardW, cardH);

        // Accent Circle
        s.addShape("ellipse", {
            x: x + 0.2, y: y + 0.15, w: 0.35, h: 0.35,
            fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 }
        });
        s.addText(String(i + 1), {
            x: x + 0.2, y: y + 0.15, w: 0.35, h: 0.35,
            fontSize: 10, bold: true, color: COLOR_GOLD,
            fontFace: FONT_TITLE, align: "center", valign: "middle"
        });

        // Title
        s.addText(c.title, {
            x: x + 0.65, y: y + 0.15, w: cardW - 0.8, h: 0.3,
            fontSize: 14, bold: true, color: COLOR_PRIMARY_DARK,
            fontFace: FONT_TITLE, align: "left"
        });

        // Description
        s.addText(c.desc, {
            x: x + 0.2, y: y + 0.55, w: cardW - 0.4, h: 0.9,
            fontSize: 11, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, align: "left", valign: "top"
        });
    });
}

// =============================================================================
// =============================================================================
// SLIDE 30 — LIMITES & PERSPECTIVES D'ÉVOLUTION
// =============================================================================
{
    const s = pres.addSlide();
    applyContentSlideTemplate(s, "Limites & Perspectives d'Évolution", "Conclusion & Perspectives", "32/33");

    const colW = 4.35;
    const colH = 3.1;
    const yVal = 1.3;

    // Left: Limites
    drawCard(s, 0.5, yVal, colW, colH);
    // Amber Left Accent Indicator Line
    s.addShape("rect", { x: 0.5, y: yVal + 0.1, w: 0.08, h: colH - 0.2, fill: { color: COLOR_GOLD }, line: { width: 0 } });

    s.addText("Limites Identifiées", {
        x: 0.7, y: yVal + 0.15, w: colW - 0.3, h: 0.35,
        fontSize: 16, bold: true, color: COLOR_RED,
        fontFace: FONT_TITLE
    });

    const limits = [
        "Paiement monétique simulé en mode sandbox (non connecté à la production réelle).",
        "Échantillon d'usabilité SUS modeste (N=18 utilisateurs testés en interne).",
        "Tests de charge et d'accès concurrents massifs non réalisés en conditions réelles.",
        "Pipeline KYC entraîné et testé sur un dataset d'images nationaux restreint."
    ];

    limits.forEach((l, idx) => {
        const py = yVal + 0.6 + idx * 0.6;
        s.addShape("ellipse", { x: 0.8, y: py + 0.05, w: 0.08, h: 0.08, fill: { color: COLOR_RED }, line: { width: 0 } });
        s.addText(l, {
            x: 0.95, y: py, w: colW - 0.5, h: 0.55,
            fontSize: 11, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, valign: "top"
        });
    });

    // Right: Perspectives
    drawCard(s, 5.15, yVal, colW, colH);
    // Green Left Accent Indicator Line
    s.addShape("rect", { x: 5.15, y: yVal + 0.1, w: 0.08, h: colH - 0.2, fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 } });

    s.addText("Perspectives Futures", {
        x: 5.35, y: yVal + 0.15, w: colW - 0.3, h: 0.35,
        fontSize: 16, bold: true, color: COLOR_PRIMARY_DARK,
        fontFace: FONT_TITLE
    });

    const perspectives = [
        "Tarification dynamique de type Surge Pricing par Machine Learning (demande spatio-temporelle).",
        "Déploiement d'une phase pilote expérimentale : Tlemcen → Oran → Alger.",
        "Intégration en production avec protocoles Chargily Pay V2 / Edahabia réels.",
        "Extension du modèle de service vers le transport collaboratif de colis (C2C)."
    ];

    perspectives.forEach((p, idx) => {
        const py = yVal + 0.6 + idx * 0.6;
        s.addShape("ellipse", { x: 5.45, y: py + 0.05, w: 0.08, h: 0.08, fill: { color: COLOR_MEDIUM_GREEN }, line: { width: 0 } });
        s.addText(p, {
            x: 5.6, y: py, w: colW - 0.5, h: 0.55,
            fontSize: 11, color: COLOR_DARK_TEXT,
            fontFace: FONT_BODY, valign: "top"
        });
    });

    // Vision Footer Card
    const fy = 4.55;
    s.addShape("roundRect", {
        x: 0.5, y: fy, w: 9.0, h: 0.65,
        fill: { color: COLOR_BG_DARK },
        line: { width: 0 },
        rectRadius: 0.08
    });
    s.addText("Vision : Devenir la référence algérienne du covoiturage sécurisé inter-wilayas, en démocratisant la mobilité partagée à travers les 69 wilayas du pays.", {
        x: 0.5, y: fy, w: 9.0, h: 0.65,
        fontSize: 12, bold: true, color: COLOR_GOLD,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// =============================================================================
// SLIDE 32 — CLOSING / THANK YOU SLIDE (Dark Background)
// =============================================================================
{
    const s = pres.addSlide();
    s.background = { color: COLOR_BG_DARK };

    // Thank you text
    s.addText("Merci pour Votre Attention", {
        x: 0.5, y: 1.2, w: 9.0, h: 0.6,
        fontSize: 32, italic: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center"
    });

    s.addText("RohWinBghit", {
        x: 0.5, y: 1.7, w: 9.0, h: 0.5,
        fontSize: 38, bold: true, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center"
    });

    s.addText("روح وين بغيت", {
        x: 0.5, y: 2.2, w: 9.0, h: 0.4,
        fontSize: 26, color: COLOR_WHITE,
        fontFace: FONT_TITLE, align: "center"
    });

    // Description
    s.addText("Plateforme mobile multiplateforme intelligente de covoiturage inter-wilayas sécurisée\nadaptée au contexte algérien", {
        x: 0.5, y: 2.65, w: 9.0, h: 0.5,
        fontSize: 13, italic: true, color: COLOR_LIGHT_GREEN,
        fontFace: FONT_BODY, align: "center"
    });

    // Presented / Supervised columns
    s.addText("Présenté par :\nAHMED BACHA Djamel Eddine\nBELHORMA Sidi Mohammed Reduane\n\nSoutenu le : 20 juin 2026", {
        x: 0.5, y: 3.3, w: 4.25, h: 1.1,
        fontSize: 10.5, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center"
    });

    s.addText("Jury de Soutenance :\nPrésidente : Mme BENMAHDI Meriem\nEncadrante : Mme BENLEDGHEM Rafika\nExaminatrice : Mme BENMANSOUR Fazilet\nExperte I2E : Mme SELADJI Yassamine", {
        x: 5.25, y: 3.3, w: 4.25, h: 1.1,
        fontSize: 10.5, color: COLOR_WHITE,
        fontFace: FONT_BODY, align: "center"
    });

    // Tagline / Date Footer
    s.addShape("rect", { x: 0, y: 4.6, w: SW, h: 0.725, fill: { color: COLOR_PRIMARY_DARK }, line: { width: 0 } });
    s.addText("Année Universitaire : 2025/2026   •   SIMPLE  •  SÛR  •  ABORDABLE", {
        x: 0, y: 4.6, w: SW, h: 0.725,
        fontSize: 13, bold: true, color: COLOR_GOLD,
        fontFace: FONT_BODY, align: "center", valign: "middle"
    });
}

// ─── SAVE PRESENTATION ──────────────────────────────────────────────────────
const outPath = path.join(__dirname, "RohWinBghit_Presentation_v2.pptx");
const outPathOriginal = path.join(__dirname, "RohWinBghit_Presentation.pptx");
pres.writeFile({ fileName: outPath })
    .then(() => {
        console.log("✅ Success: Presentation v2 generated at:", outPath);
        return pres.writeFile({ fileName: outPathOriginal });
    })
    .then(() => {
        console.log("✅ Success: Original Presentation generated at:", outPathOriginal);
        process.exit(0);
    })
    .catch(err => {
        console.error("❌ Error generating presentation:", err);
        process.exit(1);
    });
