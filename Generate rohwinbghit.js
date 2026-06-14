/**
 * RohWinBghit — Présentation PPTX (Structure Optimisée 37 Slides)
 * Génère la soutenance complète inspirée du modèle MOBIPOS.
 * Usage: node "Generate rohwinbghit.js"
 */

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const {
    FaMoneyBillWave, FaNetworkWired, FaShieldHalved, FaMapLocationDot,
    FaCreditCard, FaMobileScreen, FaComments, FaWallet, FaChartLine,
    FaUserCheck, FaIdCardClip, FaBrain, FaStopwatch, FaLocationCrosshairs,
    FaCodeBranch, FaTicket, FaStar, FaHandshake, FaChartPie,
    FaMoneyBillTrendUp, FaAward, FaSitemap, FaFlagCheckered, FaRoute, FaMap,
    FaBoltLightning, FaVialCircleCheck, FaDatabase, FaShield,
    FaMicrochip, FaQuoteRight, FaCarSide, FaBriefcase, FaUsers, FaClock,
    FaShieldVirus, FaCompress, FaKey, FaArrowTrendUp,
    FaUserLarge, FaMagnifyingGlassLocation, FaCalendarCheck, FaCirclePlus, FaUsersGear,
    FaHourglassHalf, FaCircleCheck
} = require("react-icons/fa6");
const { FaNodeJs, FaReact, FaServer, FaGraduationCap } = require("react-icons/fa");
const { TbSteeringWheel } = require("react-icons/tb");

// ─── COLOR PALETTE ───────────────────────────────────────────────────────────
const C = {
    darkGreen: "0F3A23",
    midGreen: "166E3A",
    accentGreen: "4CC080",
    deepBlack: "0A1F14",
    cream: "F5F1E8",
    white: "FFFFFF",
    cardBorder: "E7ECE8",
    mutedText: "5B6A61",
    gold: "D4A017",
    lightGreenBg: "E8F5E9",
};

// ─── ICON HELPER ─────────────────────────────────────────────────────────────
async function iconPng(IconComp, color = "#4CC080", size = 256) {
    const svg = ReactDOMServer.renderToStaticMarkup(
        React.createElement(IconComp, { color, size: String(size) })
    );
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
}

// ─── SHARED LAYOUT HELPERS ───────────────────────────────────────────────────
const SW = 10, SH = 5.625;

function addSectionTag(slide, label, dark = false) {
    // ── Barre latérale gauche (style MOBIPOS) ──
    slide.addShape("rect", {
        x: 0, y: 0, w: 0.07, h: SH,
        fill: { color: C.midGreen },
        line: { color: C.midGreen, width: 0 }
    });
    // Accent line
    slide.addShape("rect", {
        x: 0.5, y: 0.45, w: 0.32, h: 0.035,
        fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 }
    });
    slide.addText(label, {
        x: 0.88, y: 0.38, w: 6.0, h: 0.2,
        fontSize: 11, bold: true, color: dark ? C.accentGreen : C.midGreen,
        fontFace: "Courier New", charSpacing: 2
    });
}

function addPillCounter(slide, label, dark = false) {
    if (label && label.includes('/')) {
        const current = label.split('/')[0];
        label = `${current}/37`;
    }
    slide.addShape("roundRect", {
        x: 8.7, y: 0.38, w: 0.78, h: 0.3,
        fill: { color: dark ? C.deepBlack : C.darkGreen },
        line: { color: C.midGreen, width: 1 },
        rectRadius: 0.15
    });
    slide.addText(label, {
        x: 8.7, y: 0.38, w: 0.78, h: 0.3,
        fontSize: 9, bold: true, color: C.accentGreen,
        fontFace: "Courier New", align: "center", valign: "middle"
    });
}

function addWatermark(slide, num, dark = false) {
    slide.addText(String(num).padStart(2, "0"), {
        x: 5.2, y: 0.4, w: 4.5, h: 4.5,
        fontSize: 220, bold: true,
        color: dark ? "163F29" : "ECEAE0",
        fontFace: "Arial Black",
        align: "left", valign: "top",
        wrap: false
    });
}

// ─── PROGRESS BAR HELPER (style Genspark) ───────────────────────────────────────────
function addProgressBar(slide, x, y, w, h, pct, barColor, label, sublabel) {
    slide.addText(label, {
        x, y: y - 0.3, w: w - 0.65, h: 0.26,
        fontSize: 10.5, bold: true, color: C.deepBlack, fontFace: "Calibri"
    });
    if (sublabel) {
        slide.addText(sublabel, {
            x: x + w - 0.65, y: y - 0.3, w: 0.65, h: 0.26,
            fontSize: 9, color: C.mutedText, fontFace: "Calibri", align: "right"
        });
    }
    slide.addShape("roundRect", {
        x, y, w, h,
        fill: { color: "ECEAE0" }, line: { color: C.cardBorder, width: 0.3 }, rectRadius: h / 2
    });
    const fillW = Math.max(h, (w * pct) / 100);
    slide.addShape("roundRect", {
        x, y, w: fillW, h,
        fill: { color: barColor }, line: { color: barColor, width: 0 }, rectRadius: h / 2
    });
    slide.addText(`${pct}%`, {
        x: x + w + 0.1, y: y - 0.04, w: 0.55, h: h + 0.08,
        fontSize: 12, bold: true, color: barColor, fontFace: "Calibri", valign: "middle"
    });
}

function addIconCard(slide, icon, title, body, x, y, w = 3.75, h = 0.86) {
    slide.addShape("roundRect", {
        x, y, w, h,
        fill: { color: C.white },
        line: { color: C.cardBorder, width: 0.5 },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
        rectRadius: 0.1
    });
    slide.addShape("roundRect", {
        x: x + 0.12, y: y + (h - 0.38) / 2, w: 0.38, h: 0.38,
        fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.06
    });
    if (icon) slide.addImage({ data: icon, x: x + 0.17, y: y + (h - 0.28) / 2, w: 0.28, h: 0.28 });
    slide.addText(title, {
        x: x + 0.6, y: y + 0.1, w: w - 0.72, h: 0.22,
        fontSize: 11.5, bold: true, color: C.deepBlack, fontFace: "Calibri"
    });
    slide.addText(body, {
        x: x + 0.6, y: y + 0.33, w: w - 0.72, h: h - 0.38,
        fontSize: 9, color: C.mutedText, fontFace: "Calibri"
    });
}

// Helper used in slide 01 (section tag on dark background)
function slide01_sectionTag(s) {
    s.addShape("rect", { x: 0.5, y: 0.45, w: 0.32, h: 0.035, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
    s.addText("DÉBUT", { x: 0.88, y: 0.38, w: 3, h: 0.2, fontSize: 11, bold: true, color: C.accentGreen, fontFace: "Courier New", charSpacing: 2 });
}

async function buildPresentation() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane";
    pres.title = "RohWinBghit — Présentation de Soutenance Master";

    // Pre-render icons
    const icons = {
        money: await iconPng(FaMoneyBillWave, "#166E3A"),
        network: await iconPng(FaNetworkWired, "#166E3A"),
        shield: await iconPng(FaShieldHalved, "#166E3A"),
        map: await iconPng(FaMapLocationDot, "#166E3A"),
        card: await iconPng(FaCreditCard, "#166E3A"),
        mobile: await iconPng(FaMobileScreen, "#166E3A"),
        chat: await iconPng(FaComments, "#166E3A"),
        wallet: await iconPng(FaWallet, "#166E3A"),
        chart: await iconPng(FaChartLine, "#166E3A"),
        userCheck: await iconPng(FaUserCheck, "#166E3A"),
        idCard: await iconPng(FaIdCardClip, "#4CC080"),
        brain: await iconPng(FaBrain, "#4CC080"),
        stopwatch: await iconPng(FaStopwatch, "#4CC080"),
        crosshair: await iconPng(FaLocationCrosshairs, "#4CC080"),
        code: await iconPng(FaCodeBranch, "#4CC080"),
        ticket: await iconPng(FaTicket, "#166E3A"),
        star: await iconPng(FaStar, "#D4A017"),
        handshake: await iconPng(FaHandshake, "#166E3A"),
        pie: await iconPng(FaChartPie, "#166E3A"),
        trending: await iconPng(FaMoneyBillTrendUp, "#166E3A"),
        award: await iconPng(FaAward, "#166E3A"),
        sitemap: await iconPng(FaSitemap, "#166E3A"),
        flag: await iconPng(FaFlagCheckered, "#0A1F14"),
        flagGreen: await iconPng(FaFlagCheckered, "#4CC080"),
        route: await iconPng(FaRoute, "#4CC080"),
        mapIcon: await iconPng(FaMap, "#4CC080"),
        bolt: await iconPng(FaBoltLightning, "#166E3A"),
        vial: await iconPng(FaVialCircleCheck, "#166E3A"),
        database: await iconPng(FaDatabase, "#166E3A"),
        shieldChk: await iconPng(FaShield, "#4CC080"),
        micro: await iconPng(FaMicrochip, "#0A1F14"),
        quote: await iconPng(FaQuoteRight, "#0A1F14"),
        car: await iconPng(FaCarSide, "#0A1F14"),
        carCream: await iconPng(FaCarSide, "#F5F1E8"),
        brief: await iconPng(FaBriefcase, "#0A1F14"),
        users: await iconPng(FaUsers, "#4CC080"),
        clock: await iconPng(FaClock, "#4CC080"),
        node: await iconPng(FaNodeJs, "#166E3A"),
        react: await iconPng(FaReact, "#166E3A"),
        idCardG: await iconPng(FaIdCardClip, "#4CC080"),
        shieldG: await iconPng(FaShieldHalved, "#4CC080"),
        shieldH: await iconPng(FaShieldHalved, "#166E3A"),
        mapG: await iconPng(FaMapLocationDot, "#166E3A"),
        trend: await iconPng(FaArrowTrendUp, "#4CC080"),
        userLarge: await iconPng(FaUserLarge, "#166E3A"),
        steeringWheel: await iconPng(TbSteeringWheel, "#166E3A"),
        magnifyingGlass: await iconPng(FaMagnifyingGlassLocation, "#4CC080"),
        calendarCheck: await iconPng(FaCalendarCheck, "#4CC080"),
        cardG: await iconPng(FaCreditCard, "#4CC080"),
        routeG: await iconPng(FaRoute, "#4CC080"),
        circlePlus: await iconPng(FaCirclePlus, "#4CC080"),
        usersGear: await iconPng(FaUsersGear, "#4CC080"),
        chatG: await iconPng(FaComments, "#4CC080"),
        compress: await iconPng(FaCompress, "#4CC080"),
        key: await iconPng(FaKey, "#4CC080"),
        shieldVirus: await iconPng(FaShieldVirus, "#4CC080"),
        hourglass: await iconPng(FaHourglassHalf, "#166E3A"),
        circleCheck: await iconPng(FaCircleCheck, "#4CC080")
    };

    // ─── DYNAMIC SLIDE GENERATION ────────────────────────────────────────────
    let currentSlideNum = 0;

    function createSlide(sectionLabel, dark = false) {
        currentSlideNum++;
        const s = pres.addSlide();
        s.background = { color: dark ? C.darkGreen : C.cream };
        addWatermark(s, currentSlideNum, dark);
        addSectionTag(s, sectionLabel, dark);
        addPillCounter(s, `${String(currentSlideNum).padStart(2, "0")}/37`, dark);
        return s;
    }

    function createTransitionSlide(sectionLabel, sectionTitle, questionGuide) {
        const s = createSlide(sectionLabel, true); // dark green background

        // Large circle for the section number (inspired by MOBIPOS style)
        s.addShape("ellipse", {
            x: 4.6, y: 1.4, w: 0.8, h: 0.8,
            fill: { color: C.accentGreen },
            line: { color: C.accentGreen, width: 0 }
        });
        s.addText(String(currentSlideNum).padStart(2, "0"), {
            x: 4.6, y: 1.4, w: 0.8, h: 0.8,
            fontSize: 22, bold: true, color: C.deepBlack,
            fontFace: "Courier New", align: "center", valign: "middle"
        });

        // Section Title
        s.addText(sectionTitle, {
            x: 0.5, y: 2.4, w: 9.0, h: 0.8,
            fontSize: 28, bold: true, color: C.cream,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Accent line
        s.addShape("rect", {
            x: 4.0, y: 3.3, w: 2.0, h: 0.03,
            fill: { color: C.gold }, line: { color: C.gold, width: 0 }
        });

        // Question-guide
        s.addText(questionGuide, {
            x: 0.5, y: 3.5, w: 9.0, h: 0.5,
            fontSize: 16, italic: true, color: C.gold,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
        return s;
    }

    // =========================================================================
    // SLIDE 01 — Couverture académique
    // =========================================================================
    {
        currentSlideNum = 1;
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 1, true);
        slide01_sectionTag(s);
        addPillCounter(s, "01/37", true);

        // University Logo
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "univ_logo.png"), x: 4.61, y: 0.74, w: 0.78, h: 0.78 });

        // RohWinBghit Logo Bus & wordmark
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "wordmark.png"), x: 2.19, y: 1.83, w: 0.78, h: 0.78 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo_bus.png"), x: 2.19, y: 1.83, w: 0.78, h: 0.78 });

        // RohWinBghit Wordmark
        s.addText("RohWinBghit — روح وين بغيت", {
            x: 3.13, y: 1.75, w: 5.5, h: 0.94,
            fontSize: 40, bold: true, color: C.cream,
            fontFace: "Calibri", valign: "middle"
        });

        // Title
        s.addText("Plateforme mobile multiplateforme intelligente de covoiturage\ninter-wilayas sécurisée adaptée au contexte algérien", {
            x: 0.5, y: 2.8, w: 9.0, h: 1.0,
            fontSize: 22, bold: true, color: C.cream,
            align: "center", fontFace: "Calibri"
        });

        // Subtitle & Authors
        s.addText("Mémoire de Master en Génie Logiciel — Arrêté 1275\nUniversité Abou Bekr Belkaïd — Tlemcen", {
            x: 0.5, y: 3.8, w: 9.0, h: 0.45,
            fontSize: 13, color: C.accentGreen, align: "center", fontFace: "Calibri"
        });

        s.addText("Présenté par : AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane\nEncadré par : Mme BENLEDGHEM Rafika", {
            x: 0.5, y: 4.35, w: 9.0, h: 0.5,
            fontSize: 12.5, color: C.cream, align: "center", fontFace: "Calibri"
        });

        // Bottom banner
        s.addShape("rect", { x: 0, y: 5.12, w: SW, h: 0.5, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });
        s.addText("SIMPLE   •   SÛR   •   ABORDABLE", {
            x: 0, y: 5.12, w: SW, h: 0.5,
            fontSize: 14, bold: true, color: C.cream,
            align: "center", valign: "middle", charSpacing: 3, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 02 — Plan de la présentation
    // =========================================================================
    {
        currentSlideNum = 2;
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 2, false);
        addSectionTag(s, "PLAN DE LA PRÉSENTATION");
        addPillCounter(s, "02/37");

        s.addText("Plan de la présentation", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const plans = [
            ["01", "Introduction & Objectifs", "Où ? Quand ? Dans quel contexte ?"],
            ["02", "Problématique & Enjeux", "Pourquoi ? Pour qui ? Quel intérêt ?"],
            ["03", "Cadre du Projet & Marché", "Quoi faire ?"],
            ["04", "Analyse des Besoins", "Quels sont les besoins ?"],
            ["05", "Conception & Architecture", "Quels sont les concepts ?"],
            ["06", "Réalisation & Sécurité", "Comment le faire ?"],
            ["07", "Démonstration du Prototype", "À quoi ça ressemble ?"],
            ["08", "Validation & Business", "Quels résultats obtenus ?"],
            ["09", "Conclusion & Perspectives", "Apports et perspectives ?"]
        ];

        const cols = [0.45, 3.6, 6.75];
        const rows = [1.45, 2.58, 3.71];

        plans.forEach(([num, title, guide], i) => {
            const col = i % 3;
            const row = Math.floor(i / 3);
            const x = cols[col], y = rows[row];
            s.addShape("roundRect", { x, y, w: 2.8, h: 1.0, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("ellipse", { x: x + 0.12, y: y + 0.28, w: 0.4, h: 0.4, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addText(num, { x: x + 0.12, y: y + 0.28, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: C.midGreen, fontFace: "Courier New", align: "center", valign: "middle", margin: 0 });
            s.addText(title, { x: x + 0.65, y: y + 0.08, w: 2.05, h: 0.4, fontSize: 11, bold: true, color: C.deepBlack, fontFace: "Calibri", valign: "middle" });
            s.addText(guide, { x: x + 0.65, y: y + 0.62, w: 2.05, h: 0.28, fontSize: 8.5, italic: true, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 03 — Transition Section 1 : Introduction & Objectifs
    // =========================================================================
    createTransitionSlide("INTRODUCTION", "1. Introduction & Objectifs", "Où ? Quand ? Dans quel contexte ?");

    // =========================================================================
    // SLIDE 04 — Introduction - But du Projet
    // =========================================================================
    {
        const s = createSlide("INTRODUCTION");

        s.addText("But du projet : Pourquoi RohWinBghit ?", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const goals = [
            [icons.network, "Rapprochement Offre-Demande", "Lutter contre la fragmentation du transport interurbain et l'asymétrie massive de l'information entre usagers."],
            [icons.money, "Mutualisation des Coûts", "Offrir une alternative abordable (-40% sur le trajet individuel) face à l'inflation et au coût d'entretien des véhicules."],
            [icons.shield, "Confiance Numérique", "Substituer les groupes informels sur les réseaux sociaux par un écosystème d'identité traçable et souverain."]
        ];

        goals.forEach(([icon, title, body], i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.5, w: 2.9, h: 3.4, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("ellipse", { x: x + 1.1, y: 1.8, w: 0.7, h: 0.7, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: x + 1.2, y: 1.9, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.1, y: 2.7, w: 2.7, h: 0.5, fontSize: 13, bold: true, color: C.deepBlack, align: "center", fontFace: "Calibri" });
            s.addText(body, { x: x + 0.15, y: 3.3, w: 2.6, h: 1.4, fontSize: 10, color: C.mutedText, align: "center", fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 05 — Introduction - Contexte & Objectifs SMART (Fusionné)
    // =========================================================================
    {
        const s = createSlide("INTRODUCTION");

        s.addText("Contexte national & Objectifs SMART", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Contexte national
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Le Contexte Algérien", { x: 0.7, y: 1.55, w: 4.0, h: 0.3, fontSize: 15, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• Territoire de 2.38 Millions km² divisé en 69 wilayas.\n• Mobilité universitaire intense générant un flux de voyageurs étudiants massif et captif.\n• Arrêté Ministériel n° 1275 : cadre national d'incubation « Un diplôme, une Startup » à l'Université de Tlemcen.\n• Transition d'un prototype académique vers une entité créatrice de valeur économique.", {
            x: 0.7, y: 1.95, w: 4.0, h: 2.9, fontSize: 10.5, color: C.mutedText, fontFace: "Calibri"
        });

        // Right Panel: 4 small cards for SMART
        const smarts = [
            ["OG", "Prototype", "Prototype mobile + admin web avec SUS >= 70 sous 9 mois."],
            ["OS1", "Modélisation", "Spécifications juridiques, UML 2.5 et schéma 3NF en 4 semaines."],
            ["OS2", "Ingénierie", "Client React Native, API Express et Mapbox en 16 semaines."],
            ["OS3", "IA / KYC", "Microservice IA FastAPI (InsightFace) en 6 semaines."]
        ];

        smarts.forEach(([tag, title, desc], i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const cx = 5.1 + col * 2.25;
            const cy = 1.4 + row * 1.8;
            s.addShape("roundRect", { x: cx, y: cy, w: 2.15, h: 1.7, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("roundRect", { x: cx + 0.12, y: cy + 0.12, w: 0.5, h: 0.22, fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.04 });
            s.addText(tag, { x: cx + 0.12, y: cy + 0.12, w: 0.5, h: 0.22, fontSize: 8.5, bold: true, color: C.accentGreen, align: "center", valign: "middle", fontFace: "Courier New" });
            s.addText(title, { x: cx + 0.12, y: cy + 0.38, w: 1.9, h: 0.22, fontSize: 10.5, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(desc, { x: cx + 0.12, y: cy + 0.64, w: 1.9, h: 0.95, fontSize: 8.2, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 06 — Transition Section 2 : Problématique & Enjeux
    // =========================================================================
    createTransitionSlide("PROBLÉMATIQUE", "2. Problématique & Enjeux", "Pourquoi ? Pour qui ? Quel intérêt ?");

    // =========================================================================
    // SLIDE 07 — Problématique & Enjeux (Fusionné)
    // =========================================================================
    {
        const s = createSlide("PROBLÉMATIQUE");

        s.addText("Question centrale & Enjeux du projet", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Central Question (Quote Box)
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1.5 }, rectRadius: 0.1 });
        s.addShape("ellipse", { x: 2.42, y: 1.55, w: 0.56, h: 0.56, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addImage({ data: icons.quote, x: 2.52, y: 1.65, w: 0.36, h: 0.36 });

        s.addText("Comment concevoir, structurer et déployer une architecture logicielle distribuée et tolérante aux pannes réseau pour le covoiturage inter-wilayas en Algérie, capable d'instaurer de la confiance par un pipeline KYC biométrique et d'intégrer les paiements nationaux ?", {
            x: 0.7, y: 2.15, w: 4.0, h: 2.3,
            fontSize: 11.5, italic: true, color: C.cream,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
        s.addText("Question centrale de R&D", {
            x: 0.5, y: 4.5, w: 4.4, h: 0.22,
            fontSize: 8, bold: true, color: C.accentGreen, align: "center", fontFace: "Courier New"
        });

        // Right Panel: Target Audience & Market Opportunity
        // Target Audience Card
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 1.7, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
        s.addImage({ data: icons.users, x: 5.3, y: 1.55, w: 0.3, h: 0.3 });
        s.addText("Pour qui ? (Enjeux Cibles)", { x: 5.7, y: 1.57, w: 3.6, h: 0.25, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• 1,9 Million d'étudiants ayant des besoins de navette hebdomadaires.\n• Travailleurs pendulaires inter-wilayas réguliers.\n• Conducteurs cherchant à amortir leurs frais de trajet.", {
            x: 5.3, y: 1.95, w: 4.0, h: 1.0, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri"
        });

        // Market Opportunity Card
        s.addShape("roundRect", { x: 5.1, y: 3.3, w: 4.4, h: 1.7, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
        s.addImage({ data: icons.clock, x: 5.3, y: 3.45, w: 0.3, h: 0.3 });
        s.addText("Quel intérêt ? (Opportunités)", { x: 5.7, y: 3.47, w: 3.6, h: 0.25, fontSize: 13, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("• Plateau de l'infrastructure 4G nationale et smartphone.\n• Croissance exponentielle du e-paiement (CIB & Edahabia).\n• Transition sociologique vers l'économie collaborative mobile.", {
            x: 5.3, y: 3.85, w: 4.0, h: 1.0, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 08 — Transition Section 3 : Cadre du Projet
    // =========================================================================
    createTransitionSlide("CADRE DU PROJET", "3. Cadre du Projet & Marché", "Quoi faire ?");

    // =========================================================================
    // SLIDE 09 — Cadre du Projet - Arrêté 1275 & Innovation
    // =========================================================================
    {
        const s = createSlide("CADRE DU PROJET");

        s.addText("Contexte Arrêté 1275 et Innovation", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const innovations = [
            [icons.award, "Label Startup", "Projet conçu spécifiquement sous le mécanisme d'incubation universitaire visant l'obtention du label Startup d'État."],
            [icons.brain, "Souveraineté Technologique", "Création de modèles IA propriétaires d'anti-spoofing biométrique locaux adaptés au format des documents nationaux."],
            [icons.brief, "Propriété Intellectuelle", "Enregistrements actifs du code source auprès de l'ONDA et de la marque auprès de l'INAPI."]
        ];

        innovations.forEach(([icon, title, body], i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.5, w: 2.9, h: 3.4, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("ellipse", { x: x + 1.1, y: 1.8, w: 0.7, h: 0.7, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: x + 1.2, y: 1.9, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.1, y: 2.7, w: 2.7, h: 0.5, fontSize: 13, bold: true, color: C.deepBlack, align: "center", fontFace: "Calibri" });
            s.addText(body, { x: x + 0.15, y: 3.3, w: 2.6, h: 1.4, fontSize: 10, color: C.mutedText, align: "center", fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 10 — Cadre du Projet - Marché & Concurrence (Fusionné)
    // =========================================================================
    {
        const s = createSlide("MARCHÉ");

        s.addText("Le marché numérique & limites de la concurrence", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Market Stats
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Indicateurs Marché Algérien", { x: 0.7, y: 1.55, w: 4.0, h: 0.3, fontSize: 15, bold: true, color: C.darkGreen, fontFace: "Calibri" });

        const mstats = [
            ["48M+", "Accès Internet Mobile"],
            ["75%+", "Adoption Smartphone"],
            ["7M+", "Transactions CIB/Edahabia"],
            ["1.9M", "Étudiants Actifs"]
        ];

        mstats.forEach(([num, label], i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const sx = 0.75 + col * 2.0;
            const sy = 1.95 + row * 1.3;
            s.addShape("roundRect", { x: sx, y: sy, w: 1.8, h: 1.1, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.08 });
            s.addText(num, { x: sx + 0.05, y: sy + 0.15, w: 1.7, h: 0.4, fontSize: 22, bold: true, color: C.accentGreen, align: "center", valign: "middle", fontFace: "Calibri" });
            s.addText(label, { x: sx + 0.05, y: sy + 0.6, w: 1.7, h: 0.4, fontSize: 8.5, bold: true, color: C.mutedText, align: "center", fontFace: "Calibri" });
        });

        // Right Panel: Competitors
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Limites des Solutions Concurrentes", { x: 5.3, y: 1.55, w: 4.0, h: 0.3, fontSize: 15, bold: true, color: C.gold, fontFace: "Calibri" });

        const comps = [
            ["BlaBlaCar", "Pas de DZD ni de KYC local algérien."],
            ["Yassir", "Limité à l'urbain; tarifs de VTC élevés."],
            ["Nroho", "KYC déclaratif (risques); évasion de commissions."],
            ["inDrive", "Pas de séquestre SATIM (espèces uniquement)."]
        ];

        comps.forEach(([name, limit], i) => {
            const cy = 1.95 + i * 0.72;
            s.addShape("roundRect", { x: 5.3, y: cy, w: 1.2, h: 0.55, fill: { color: "FDF2F2" }, line: { color: "F8D7DA", width: 0.5 }, rectRadius: 0.06 });
            s.addText(name, { x: 5.3, y: cy, w: 1.2, h: 0.55, fontSize: 9.5, bold: true, color: "A12015", align: "center", valign: "middle", fontFace: "Calibri" });
            s.addText(limit, { x: 6.6, y: cy, w: 2.7, h: 0.55, fontSize: 8.8, color: C.mutedText, fontFace: "Calibri", valign: "middle" });
        });
    }

    // =========================================================================
    // SLIDE 11 — Cadre du Projet - Positionnement RohWinBghit
    // =========================================================================
    {
        const s = createSlide("POSITIONNEMENT");

        s.addText("Positionnement et avantages concurrentiels", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addShape("roundRect", {
            x: 0.5, y: 1.35, w: 9, h: 3.95,
            fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.04 },
            rectRadius: 0.1
        });

        const colX = [0.75, 3.35, 5.65, 7.5];
        const colW = [2.4, 2.1, 1.8, 2.25];
        const headers = ["Critères d'évaluation", "International\n(BlaBlaCar)", "Maghreb / Algérie\n(Yassir, Nroho)", "RohWinBghit"];
        headers.forEach((h, i) => {
            const isRWB = i === 3;
            if (isRWB) {
                s.addShape("roundRect", { x: colX[i], y: 1.4, w: colW[i], h: 0.65, fill: { color: C.accentGreen, transparency: 85 }, line: { color: C.accentGreen, width: 1 }, rectRadius: 0.08 });
            }
            s.addText(h, {
                x: colX[i] + 0.05, y: 1.42, w: colW[i] - 0.1, h: 0.6,
                fontSize: isRWB ? 13 : 9.5, bold: true,
                color: isRWB ? C.darkGreen : (i === 0 ? C.mutedText : C.deepBlack),
                fontFace: "Calibri", align: isRWB ? "center" : "left", valign: "middle"
            });
        });

        s.addShape("rect", { x: 0.6, y: 2.07, w: 8.8, h: 0.02, fill: { color: C.cardBorder }, line: { color: C.cardBorder, width: 0 } });

        const rows2 = [
            ["KYC Biométrique Active (CNN)", "✗  Aucun", "✗  Manuel différé", "✓  FastAPI + InsightFace"],
            ["Séquestre Monétique SATIM", "✗  CB internationale", "✗  Espèces uniquement", "✓  Espèces + CIB/Edahabia"],
            ["Conformité Données (Loi 25-11)", "✗  Non documenté", "✗  Non documenté", "✓  Chiffrement AES-256-GCM"],
            ["Tarification Surge Pricing", "✗  Fixe par siège", "✓  Dynamique urbain", "✓  Dynamique inter-wilayas"]
        ];

        rows2.forEach(([crit, intl, magh, rwb], ri) => {
            const ry = 2.12 + ri * 0.82;
            const isLast = ri === rows2.length - 1;
            if (ri < rows2.length - 1) {
                s.addShape("rect", { x: 0.6, y: ry + 0.78, w: 8.8, h: 0.01, fill: { color: C.cardBorder }, line: { color: C.cardBorder, width: 0 } });
            }
            s.addText(crit, { x: colX[0] + 0.05, y: ry + 0.2, w: colW[0] - 0.1, h: 0.4, fontSize: 10, bold: true, color: C.deepBlack, fontFace: "Calibri" });

            [[intl, 1], [magh, 2]].forEach(([txt, ci]) => {
                s.addText(txt, { x: colX[ci] + 0.05, y: ry + 0.2, w: colW[ci] - 0.1, h: 0.4, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
            });

            s.addShape("rect", { x: colX[3], y: ry + 0.02, w: colW[3], h: isLast ? 0.75 : 0.78, fill: { color: C.accentGreen, transparency: 85 }, line: { color: C.accentGreen, width: 0.5 } });
            s.addText(rwb, { x: colX[3] + 0.05, y: ry + 0.2, w: colW[3] - 0.1, h: 0.4, fontSize: 10, bold: true, color: C.midGreen, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 12 — Cadre du Projet - Méthodologie UP/Scrum & Enquête
    // =========================================================================
    {
        const s = createSlide("MÉTHODOLOGIE");

        s.addText("Méthodologie UP + Scrum & Enquête", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Card: Méthodologie Scrum
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.25, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Méthode Scrum — 5 Sprints de 2 semaines", { x: 0.7, y: 1.55, w: 3.85, h: 0.28, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("Processus Unifié (UP) + Scrum agile", { x: 0.7, y: 1.84, w: 3.85, h: 0.2, fontSize: 9.5, italic: true, color: C.mutedText, fontFace: "Calibri" });

        const scrumSteps = [
            { num: "01", label: "Backlog du produit", desc: "User stories priorisées par valeur métier" },
            { num: "02", label: "Planification Sprint", desc: "Sélection des tâches + estimation Story Points" },
            { num: "03", label: "Sprint (2 semaines)", desc: "Développement itératif → livrable fonctionnel" },
            { num: "04", label: "Revue & Démonstration", desc: "Démo à l’encadrante + feedback intégré" },
            { num: "05", label: "Rétrospective", desc: "Amélioration du processus à chaque cycle" },
        ];
        scrumSteps.forEach(({ num, label, desc }, i) => {
            const sy = 2.1 + i * 0.55;
            if (i < scrumSteps.length - 1) {
                s.addShape("rect", { x: 0.87, y: sy + 0.34, w: 0.04, h: 0.22, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
            }
            s.addShape("ellipse", { x: 0.7, y: sy + 0.01, w: 0.38, h: 0.38, fill: { color: num === "03" ? C.accentGreen : C.darkGreen }, line: { color: num === "03" ? C.accentGreen : C.darkGreen, width: 0 } });
            s.addText(num, { x: 0.7, y: sy + 0.01, w: 0.38, h: 0.38, fontSize: 9, bold: true, color: C.cream, align: "center", valign: "middle", fontFace: "Courier New" });
            s.addText(label, { x: 1.2, y: sy + 0.01, w: 3.3, h: 0.2, fontSize: 10.5, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(desc, { x: 1.2, y: sy + 0.22, w: 3.3, h: 0.18, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Right Card: Enquête terrain
        s.addShape("roundRect", { x: 5.25, y: 1.4, w: 4.25, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Enquête Terrain (N = 150)", { x: 5.45, y: 1.6, w: 3.85, h: 0.35, fontSize: 16, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("• 68% d'étudiants | 72% de 18-25 ans.\n• Critère décisionnel prioritaire : Sécurité & Confiance (62%) face aux coûts (24%).\n• 71% prêts à fournir des données biométriques pour garantir leur sécurité.\n• 57% réclament un paiement par carte locale (CIB/Edahabia).", {
            x: 5.45, y: 2.1, w: 3.85, h: 2.6, fontSize: 11.5, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 13 — Cadre du Projet - Stack Technique
    // =========================================================================
    {
        const s = createSlide("STACK TECHNIQUE");

        s.addText("Environnement de travail et stack technologique", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const stacks = [
            [icons.react, "React Native", "Expo SDK 54 / TypeScript pour le client mobile unique cross-platform iOS & Android."],
            [icons.node, "Node.js & Express", "Serveur API REST modulaire configuré en couches découplées (SRP, OCP)."],
            [icons.micro, "FastAPI / Python", "Microservice IA isolé pour le traitement asynchrone d'InsightFace, ArcFace et liveness."],
            [icons.database, "PostgreSQL 16", "SGBD relationnel robuste gérant la concurrence d'accès par locks pessimistes FOR UPDATE."],
            [icons.bolt, "Redis 7 & Socket.io", "Redis (cache sessions et queues BullMQ), Socket.io (suivi GPS et messagerie temps réel)."],
            [icons.code, "Docker & EAS", "Dockerisation complète des services backend + Expo Application Services (compilations)."]
        ];

        const row1y = 1.35, row2y = 2.55, row3y = 3.75;
        const col1x = 0.5, col2x = 5.25;
        const positions = [
            [col1x, row1y], [col1x, row2y], [col1x, row3y],
            [col2x, row1y], [col2x, row2y], [col2x, row3y]
        ];

        stacks.forEach(([icon, title, body], i) => {
            const [x, y] = positions[i];
            s.addShape("roundRect", { x, y, w: 4.25, h: 1.05, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("roundRect", { x: x + 0.12, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.06 });
            s.addImage({ data: icon, x: x + 0.17, y: y + 0.27, w: 0.4, h: 0.4 });
            s.addText(title, { x: x + 0.75, y: y + 0.12, w: 3.4, h: 0.22, fontSize: 12, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: x + 0.75, y: y + 0.38, w: 3.4, h: 0.62, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 14 — Transition Section 4 : Spécification des Besoins
    // =========================================================================
    createTransitionSlide("BESOINS", "4. Spécification des Besoins", "Quels sont les besoins ?");

    // =========================================================================
    // SLIDE 15 — Spécification des Besoins - Acteurs
    // =========================================================================
    {
        const s = createSlide("BESOINS");

        s.addText("Identification des acteurs du système", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Acteurs humains (3 colonnes)
        const actors = [
            [icons.userLarge, "Le Passager", "Recherche les trajets inter-wilayas, effectue les réservations, paye en ligne (CIB/Edahabia) et valide l'embarquement."],
            [icons.steeringWheel, "Le Conducteur", "Enregistre son véhicule, soumet ses documents, publie ses trajets et transporte les passagers."],
            [icons.usersGear, "L'Administrateur", "Supervise l'activité, résout les litiges financiers, et valide manuellement les KYC en attente de revue."]
        ];

        actors.forEach(([icon, title, body], i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.4, w: 2.9, h: 2.2, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("ellipse", { x: x + 0.15, y: 1.55, w: 0.5, h: 0.5, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: x + 0.21, y: 1.61, w: 0.38, h: 0.38 });
            s.addText(title, { x: x + 0.75, y: 1.65, w: 2.05, h: 0.3, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });
            s.addText(body, { x: x + 0.15, y: 2.15, w: 2.6, h: 1.3, fontSize: 9, color: C.mutedText, fontFace: "Calibri" });
        });

        // Acteurs Systèmes (2 colonnes en bas)
        s.addShape("roundRect", { x: 0.5, y: 3.8, w: 4.4, h: 1.2, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
        s.addShape("ellipse", { x: 0.65, y: 3.95, w: 0.45, h: 0.45, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
        s.addImage({ data: icons.brain, x: 0.71, y: 4.01, w: 0.33, h: 0.33 });
        s.addText("Service KYC (Interne)", { x: 1.2, y: 4.0, w: 3.5, h: 0.22, fontSize: 12, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("Microservice IA FastAPI chargé des inférences asynchrones d'anti-spoofing et de comparaison faciale.", { x: 1.2, y: 4.25, w: 3.5, h: 0.65, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });

        s.addShape("roundRect", { x: 5.1, y: 3.8, w: 4.4, h: 1.2, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
        s.addShape("ellipse", { x: 5.25, y: 3.95, w: 0.45, h: 0.45, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
        s.addImage({ data: icons.card, x: 5.31, y: 4.01, w: 0.33, h: 0.33 });
        s.addText("SATIM / Algérie Poste (Externe)", { x: 5.8, y: 4.0, w: 3.5, h: 0.22, fontSize: 12, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("Passerelle monétique nationale traitant les flux de paiement sécurisés CIB et Edahabia.", { x: 5.8, y: 4.25, w: 3.5, h: 0.65, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });
    }

    // =========================================================================
    // SLIDE 16 — Spécification des Besoins - Cahier des Charges (Besoins F/NF)
    // =========================================================================
    {
        const s = createSlide("BESOINS");

        s.addText("Cahier des charges : Spécifications et contraintes", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Fonctionnelles
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.code, x: 0.7, y: 1.55, w: 0.35, h: 0.35 });
        s.addText("Spécifications Fonctionnelles", { x: 1.2, y: 1.57, w: 3.5, h: 0.3, fontSize: 15, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• Authentification SMS OTP : Inscription instantanée sécurisée sans mot de passe.\n• Pipeline KYC Biométrique : Vérification d'identité et de vivacité active par IA.\n• Moteur de Recherche Géographique : Indexation spatiale des trajets sur 69 wilayas.\n• Séquestre Monétique : Verrouillage financier CIB/Edahabia jusqu'à validation.\n• Tracking GPS & Chat temps réel : Socket.io pour la coordination du covoiturage.", {
            x: 0.7, y: 2.0, w: 4.0, h: 2.8, fontSize: 10.2, color: C.mutedText, fontFace: "Calibri"
        });

        // Right Panel: Non-Fonctionnelles & Sécurité
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.shieldH, x: 5.3, y: 1.55, w: 0.35, h: 0.35 });
        s.addText("Contraintes Techniques & Qualité", { x: 5.8, y: 1.57, w: 3.5, h: 0.3, fontSize: 15, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("• Sécurité & Confidentialité : Chiffrement AES-256-GCM des pièces d'identité et conformité stricte à la Loi 25-11.\n• Performance & Latence : Temps de réponse API p95 < 200 ms et inférence du KYC biométrique < 3 secondes.\n• Robustesse Réseau : Cache persistant SQLite/AsyncStorage pour tolérer les zones blanches 4G sur autoroute.\n• Scalabilité : Services sans état conteneurisés Docker et files d'attente Redis / BullMQ.", {
            x: 5.3, y: 2.0, w: 4.0, h: 2.8, fontSize: 10.2, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 17 — Spécification des Besoins - Diagramme Use Case
    // =========================================================================
    {
        const s = createSlide("BESOINS (UML)");

        s.addText("Diagramme de cas d'utilisation global (Consolidé)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Container Card
        s.addShape("roundRect", {
            x: 1.2, y: 1.3, w: 7.6, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // Use Case Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_use_case.png"),
            x: 1.3, y: 1.4, w: 7.4, h: 3.8,
            sizing: { type: "contain", w: 7.4, h: 3.8 }
        });
    }

    // =========================================================================
    // SLIDE 18 — Transition Section 5 : Conception
    // =========================================================================
    createTransitionSlide("CONCEPTION", "5. Conception & Architecture", "Quels sont les concepts ?");

    // =========================================================================
    // SLIDE 19 — Conception - Architecture & Déploiement (Fusionné)
    // =========================================================================
    {
        const s = createSlide("CONCEPTION");

        s.addText("Architecture globale & déploiement physique", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Architecture (4 layers)
        const archBox = (x, y, w, h, title, items, borderClr, bgAlpha = 80) => {
            s.addShape("roundRect", { x, y, w, h, fill: { color: C.deepBlack, transparency: bgAlpha }, line: { color: borderClr, width: 1.2 }, rectRadius: 0.06 });
            s.addText(title, { x: x + 0.08, y: y + 0.05, w: w - 0.16, h: 0.18, fontSize: 7.5, bold: true, color: borderClr, fontFace: "Courier New" });
            items.forEach((item, i) => {
                s.addText(item, { x: x + 0.1, y: y + 0.22 + i * 0.16, w: w - 0.2, h: 0.18, fontSize: 7.5, color: C.cream, fontFace: "Calibri" });
            });
        };

        archBox(0.5, 1.4, 4.4, 0.75, "📱 CLIENTS", ["App Mobile (React Native / Expo SDK 54)", "Dashboard Web d'Administration (React / Vite)"], C.accentGreen, 65);
        archBox(0.5, 2.22, 4.4, 1.05, "⚙️ BACKEND principal (Express.js)", ["Passerelle API, routage géographique, surge pricing", "PaymentFactory (CIB, Edahabia, Cash)", "Queues asynchrones avec BullMQ 5.71"], C.accentGreen, 70);
        archBox(0.5, 3.35, 4.4, 0.75, "🧠 SERVICE IA (FastAPI)", ["Inférence d'InsightFace (antelopev2)", "Vivacité active & détection anti-spoofing"], C.accentGreen, 65);
        archBox(0.5, 4.18, 4.4, 0.8, "💾 PERSISTANCE & CACHE", ["PostgreSQL 16 (verrous ACID FOR UPDATE)", "Redis 7 (cache de session et BullMQ queues)"], C.gold, 70);

        // Right Panel: Deployment Diagram Container
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Déploiement Physique Conteneurisé", { x: 5.3, y: 1.55, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addImage({
            path: path.join(__dirname, "figures", "fig_deployment.png"),
            x: 5.2, y: 1.95, w: 4.2, h: 2.9,
            sizing: { type: "contain", w: 4.2, h: 2.9 }
        });
    }

    // =========================================================================
    // SLIDE 20 — Conception - Pipeline KYC Biométrique
    // =========================================================================
    {
        const s = createSlide("CONCEPTION (IA)");

        s.addText("Pipeline mathématique et décisionnel du KYC", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const steps = [
            [icons.idCard, "1. Extraction OCR de la CIN", "Validation de la présence de la carte d'identité, contrôle de format et extraction du nom/prénom."],
            [icons.brain, "2. Comparaison faciale", "Extraction d'embeddings faciaux à 512 dimensions par ArcFace (antelopev2) et calcul de la distance cosinus."],
            [icons.stopwatch, "3. Détection de vivacité", "Vérification active de la présence physique (anti-spoofing) par analyse du mouvement (flux vidéo)."]
        ];

        steps.forEach(([icon, title, body], i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.4, w: 2.9, h: 2.3, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("ellipse", { x: x + 0.15, y: 1.55, w: 0.45, h: 0.45, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: x + 0.20, y: 1.60, w: 0.35, h: 0.35 });
            s.addText(title, { x: x + 0.15, y: 2.1, w: 2.6, h: 0.3, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });
            s.addText(body, { x: x + 0.15, y: 2.45, w: 2.6, h: 1.15, fontSize: 9, color: C.mutedText, fontFace: "Calibri" });
        });

        s.addShape("roundRect", { x: 0.5, y: 3.9, w: 9.0, h: 1.1, fill: { color: C.darkGreen }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.1 });
        s.addText("STRATÉGIE MULTI-SEUILS BIOMÉTRIQUES DIFFÉRENCIÉE PAR RÔLE", { x: 0.7, y: 4.0, w: 8.6, h: 0.22, fontSize: 8.5, bold: true, color: C.accentGreen, fontFace: "Courier New" });
        s.addText("• Profil Conducteur (Criticité Haute) : Identité >= 0.65 | Vivacité >= 0.70 | Anti-spoofing >= 0.75\n• Profil Passager (Criticité Moyenne) : Identité >= 0.60 | Vivacité >= 0.60 | Anti-spoofing >= 0.60", {
            x: 0.7, y: 4.25, w: 8.6, h: 0.65, fontSize: 10.5, color: C.cream, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 21 — Conception - Séquence Inscription OTP
    // =========================================================================
    {
        const s = createSlide("CONCEPTION (UML)");

        s.addText("Séquence Inscription & Validation OTP", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 20, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addShape("roundRect", {
            x: 1.5, y: 1.3, w: 7.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_sequence_inscription.png"),
            x: 1.6, y: 1.4, w: 6.8, h: 3.8,
            sizing: { type: "contain", w: 6.8, h: 3.8 }
        });
    }

    // =========================================================================
    // SLIDE 22 — Conception - Séquence Réservation & Paiement
    // =========================================================================
    {
        const s = createSlide("CONCEPTION (UML)");

        s.addText("Séquence Réservation avec Séquestre Monétique", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 20, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addShape("roundRect", {
            x: 0.5, y: 1.3, w: 9.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_sequence_booking.png"),
            x: 0.6, y: 1.4, w: 8.8, h: 3.8,
            sizing: { type: "contain", w: 8.8, h: 3.8 }
        });
    }

    // =========================================================================
    // SLIDE 23 — Transition Section 6 : Réalisation
    // =========================================================================
    createTransitionSlide("RÉALISATION", "6. Réalisation & Sécurité", "Comment le faire ?");

    // =========================================================================
    // SLIDE 24 — Réalisation - Sprints & Difficultés/Solutions (Fusionné)
    // =========================================================================
    {
        const s = createSlide("RÉALISATION");

        s.addText("Calendrier Scrum & Résolution des Défis", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Sprints Timeline
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Calendrier de Réalisation (5 Sprints)", { x: 0.7, y: 1.55, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: C.darkGreen, fontFace: "Calibri" });

        const timelines = [
            ["S1", "DB & Setup : Monorepo, schéma 3NF, API REST."],
            ["S2", "Métier : Publication, recherche wilayas, chat."],
            ["S3", "IA : Pipeline FastAPI, ArcFace, vivacité."],
            ["S4", "Monétique : Sandbox SATIM, PaymentFactory, surge."],
            ["S5", "Supervision : Console admin, 151 tests Jest."]
        ];

        timelines.forEach(([step, desc], i) => {
            const sy = 1.95 + i * 0.58;
            if (i < timelines.length - 1) {
                s.addShape("rect", { x: 0.88, y: sy + 0.32, w: 0.03, h: 0.28, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
            }
            s.addShape("ellipse", { x: 0.72, y: sy, w: 0.35, h: 0.35, fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 } });
            s.addText(step, { x: 0.72, y: sy, w: 0.35, h: 0.35, fontSize: 8, bold: true, color: C.cream, align: "center", valign: "middle", fontFace: "Courier New" });
            s.addText(desc, { x: 1.2, y: sy, w: 3.5, h: 0.45, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Right Panel: Difficultés & Solutions
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Défis Rencontrés & Solutions Injectées", { x: 5.3, y: 1.55, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: C.gold, fontFace: "Calibri" });

        const challenges = [
            ["Déconnexions 4G", "Caching SQLite & AsyncStorage + requêtes HTTP idempotentes."],
            ["Double Réservation", "Verrou pessimiste SELECT FOR UPDATE ACID PostgreSQL."],
            ["Évasion Commission", "Mécanisme de séquestre monétique SATIM bloquant les fonds."]
        ];

        challenges.forEach(([def, sol], i) => {
            const cy = 1.95 + i * 1.0;
            s.addShape("roundRect", { x: 5.3, y: cy, w: 4.0, h: 0.88, fill: { color: C.cream }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addText("⚠️ " + def, { x: 5.42, y: cy + 0.05, w: 3.76, h: 0.22, fontSize: 10, bold: true, color: "A12015", fontFace: "Calibri" });
            s.addText("✓ " + sol, { x: 5.42, y: cy + 0.28, w: 3.76, h: 0.55, fontSize: 8.8, color: C.darkGreen, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 25 — Réalisation - Interfaces Mobile (Screenshots)
    // =========================================================================
    {
        const s = createSlide("RÉALISATION (MOBILE)");

        s.addText("Interfaces de l'application mobile (Passager)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Phone Mockup 1
        s.addShape("roundRect", { x: 1.09, y: 1.52, w: 1.72, h: 3.48, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Screen1.jpg"), x: 1.14, y: 1.57, w: 1.62, h: 3.38 });
        s.addShape("roundRect", { x: 1.65, y: 1.63, w: 0.6, h: 0.15, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        // Phone Mockup 2
        s.addShape("roundRect", { x: 3.13, y: 1.52, w: 1.72, h: 3.48, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Screen2.jpg"), x: 3.18, y: 1.57, w: 1.62, h: 3.38 });
        s.addShape("roundRect", { x: 3.69, y: 1.63, w: 0.6, h: 0.15, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        const props = [
            [icons.idCardG, "Authentification OTP & KYC", "Processus d'accueil fluide sans mot de passe, suivi d'un contrôle d'identité de l'utilisateur par scan et vivacité faciale."],
            [icons.routeG, "Publication & Matching", "Appariement dynamique de l'offre et de la demande inter-wilayas avec des calculs d'itinéraires et d'ETA exacts."]
        ];

        props.forEach(([icon, title, body], i) => {
            const y = 1.7 + i * 1.6;
            s.addShape("roundRect", { x: 5.31, y, w: 3.91, h: 1.35, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("ellipse", { x: 5.48, y: y + 0.15, w: 0.45, h: 0.45, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: 5.53, y: y + 0.20, w: 0.35, h: 0.35 });
            s.addText(title, { x: 6.05, y: y + 0.15, w: 3.0, h: 0.25, fontSize: 13, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: 5.48, y: y + 0.7, w: 3.6, h: 0.6, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 26 — Réalisation - Dashboard Administrateur
    // =========================================================================
    {
        const s = createSlide("RÉALISATION (WEB)");

        s.addText("Console d'Administration Web (React 18 / Vite)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: UI Capture Mockup
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.8, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 1 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "figures", "v2_landing_page.jpg"), x: 0.6, y: 1.5, w: 4.6, h: 3.4, sizing: { type: "contain", w: 4.6, h: 3.4 } });

        // Right Panel: Admin capabilities
        s.addShape("roundRect", { x: 5.6, y: 1.4, w: 3.9, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Fonctionnalités de supervision", { x: 5.8, y: 1.6, w: 3.5, h: 0.3, fontSize: 16, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• Modération KYC : File d'attente manuelle pour dossiers litigieux avec affichage comparé de la CIN, selfie et scores d'IA.\n• Gestion financière : Suivi des portefeuilles et libération du séquestre de paiement SATIM.\n• Administration du trafic : Modération des publications de trajets et exclusion des comptes signalés.", {
            x: 5.8, y: 2.1, w: 3.5, h: 2.7, fontSize: 11, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 27 — Réalisation - Sécurité OWASP & Conformité
    // =========================================================================
    {
        const s = createSlide("RÉALISATION (SÉCURITÉ)");

        s.addText("Sécurisation et conformité réglementaire", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const secs = [
            [icons.shieldChk, "Couverture OWASP Top 10", "Mitigation systématique : injections (Knex.js ORM), XSS (JSX escaping), failles d'authentification (tokens JWT signés et chiffrés)."],
            [icons.compress, "Chiffrement AES-256-GCM", "Conformité Loi 25-11 relative à la protection des données personnelles. Les images de CIN et embeddings d'identités sont chiffrés au repos."],
            [icons.shieldVirus, "Détection Active de Vivacité", "Contrôle d'authentification contre les attaques par rejeu de photographies ou d'écrans vidéo lors de la connexion."]
        ];

        secs.forEach(([icon, title, body], i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.5, w: 2.9, h: 3.4, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("ellipse", { x: x + 1.1, y: 1.8, w: 0.7, h: 0.7, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addImage({ data: icon, x: x + 1.2, y: 1.9, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.1, y: 2.7, w: 2.7, h: 0.5, fontSize: 13, bold: true, color: C.deepBlack, align: "center", fontFace: "Calibri" });
            s.addText(body, { x: x + 0.15, y: 3.3, w: 2.6, h: 1.4, fontSize: 9.8, color: C.mutedText, align: "center", fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 28 — Transition Section 7 : Démonstration
    // =========================================================================
    createTransitionSlide("DÉMONSTRATION", "7. Démonstration du Prototype", "À quoi ça ressemble ?");

    // =========================================================================
    // SLIDE 29 — Démonstration - Scénario de Réservation (Screenshots)
    // =========================================================================
    {
        const s = createSlide("DÉMONSTRATION");

        s.addText("Scénario de réservation en direct", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const phones = [
            { path: "Passager réserve.jpg", label: "1. Réservation", desc: "Le passager sélectionne son trajet et valide le mode de paiement." },
            { path: "Trajet publié.jpg", label: "2. Notification", desc: "Le conducteur reçoit une alerte push contenant la demande d'embarquement." },
            { path: "QR scanné au départ.jpg", label: "3. Scan & Départ", desc: "Le conducteur scanne le billet QR cryptographique du passager au départ." }
        ];

        phones.forEach((ph, i) => {
            const x = 0.5 + i * 3.1;
            s.addShape("roundRect", { x, y: 1.4, w: 2.9, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("roundRect", { x: x + 0.65, y: 1.5, w: 1.6, h: 2.2, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.05 });
            s.addImage({ path: path.join(__dirname, "Screen_mobile", ph.path), x: x + 0.68, y: 1.53, w: 1.54, h: 2.14 });
            s.addText(ph.label, { x: x + 0.1, y: 3.8, w: 2.7, h: 0.22, fontSize: 12, bold: true, color: C.darkGreen, fontFace: "Calibri" });
            s.addText(ph.desc, { x: x + 0.1, y: 4.05, w: 2.7, h: 0.85, fontSize: 9.2, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 30 — Transition Section 8 : Validation & Business
    // =========================================================================
    createTransitionSlide("VALIDATION", "8. Validation & Business", "Quels résultats obtenus ?");

    // =========================================================================
    // SLIDE 31 — Validation - Tests & Score Usabilité (SUS)
    // =========================================================================
    {
        const s = createSlide("VALIDATION");

        s.addText("Tests unitaires, d'intégration et d'utilisabilité (SUS)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Jest suite
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.25, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.vial, x: 0.7, y: 1.6, w: 0.35, h: 0.35 });
        s.addText("Tests Automatisés Jest", { x: 1.2, y: 1.62, w: 3.3, h: 0.3, fontSize: 16, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• 151 Cas de tests exécutés avec succès.\n• 107 Tests Unitaires isolés validant la logique métier à 100% de réussite.\n• 44 Tests d'Intégration validant les intéractions DB/Redis/BullMQ (90.9% de réussite).\n• Taux de réussite global de la suite de tests : 97.4 %.", {
            x: 0.7, y: 2.1, w: 3.85, h: 2.7, fontSize: 11, color: C.mutedText, fontFace: "Calibri"
        });

        // Right Panel: SUS Usability
        s.addShape("roundRect", { x: 5.25, y: 1.4, w: 4.25, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.userCheck, x: 5.45, y: 1.6, w: 0.35, h: 0.35 });
        s.addText("Évaluation d'Utilisabilité (SUS)", { x: 5.95, y: 1.62, w: 3.3, h: 0.3, fontSize: 16, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("• Menée auprès de 18 usagers réels (étudiants et conducteurs).\n• Score Global SUS de 71.6 / 100.\n• Bangor classification : Usabilité qualifiée de « Bonne » (seuil >= 68).\n• Score Passagers : 74.2 / 100.\n• Score Conducteurs : 68.9 / 100 (friction KYC justifiant l'ajout de tutoriels).", {
            x: 5.45, y: 2.1, w: 3.85, h: 2.7, fontSize: 11, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 32 — Business - Business Model Canvas (BMC)
    // =========================================================================
    {
        const s = createSlide("BUSINESS");

        s.addText("Business Model Canvas (Synthèse 9 Blocs)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const cell = (x, y, w, h, title, body) => {
            s.addShape("roundRect", { x, y, w, h, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.05 });
            s.addText(title, { x: x + 0.08, y: y + 0.08, w: w - 0.16, h: 0.18, fontSize: 8.5, bold: true, color: C.darkGreen, fontFace: "Courier New" });
            s.addText(body, { x: x + 0.08, y: y + 0.28, w: w - 0.16, h: h - 0.34, fontSize: 7.2, color: C.mutedText, fontFace: "Calibri" });
        };

        // Col 1
        cell(0.5, 1.35, 1.76, 2.5, "PARTENAIRES CLÉS", "• SATIM & GIE Monétique.\n• Incubateur Université Tlemcen.\n• Hetzner (hébergement).\n• Médias universitaires.");
        // Col 2
        cell(2.31, 1.35, 1.76, 1.22, "ACTIVITÉS CLÉS", "• Dev et maintenance API.\n• Revue manuelle KYC.\n• Support & médiation.");
        cell(2.31, 2.63, 1.76, 1.22, "RESSOURCES CLÉS", "• Pipeline IA InsightFace.\n• Serveur PostgreSQL/Redis.\n• Marque INAPI & ONDA.");
        // Col 3
        cell(4.12, 1.35, 1.76, 2.5, "VALEUR AJOUTÉE", "• Covoiturage inter-wilayas sécurisé.\n• KYC biométrique active.\n• Intégration CIB/Edahabia.\n• Indemnisations.");
        // Col 4
        cell(5.93, 1.35, 1.76, 1.22, "RELATIONS CLIENTS", "• Support client in-app.\n• Système d'évaluation.\n• Processus KYC autonome.");
        cell(5.93, 2.63, 1.76, 1.22, "CANAUX", "• App Store & Play Store.\n• Groupes universitaires.\n• Réseaux sociaux.");
        // Col 5
        cell(7.74, 1.35, 1.76, 2.5, "SEGMENTS CLIENTS", "• Passagers : étudiants, travailleurs pendulaires.\n• Conducteurs : navetteurs réguliers cherchant à amortir leurs frais.");

        // Bottom row
        cell(0.5, 3.9, 4.47, 1.2, "STRUCTURE DE COÛTS", "• Consommation API cartographique Mapbox.\n• Serveur Cloud & base de données. Consommation API SMS OTP.\n• Masse salariale et budget marketing.");
        cell(5.03, 3.9, 4.47, 1.2, "SOURCES DE REVENUS", "• Commission logicielle 12% par siège réservé (8% en bêta).\n• Abonnement Premium conducteurs (1 500 DZD/mois).\n• Publicité locale ciblée in-app (CPM 2 000 DZD).\n• Abonnements B2B.");
    }

    // =========================================================================
    // SLIDE 33 — Business - Modèle de revenus
    // =========================================================================
    {
        const s = createSlide("MONÉTISATION");

        s.addText("Modèle de revenus et monétisation", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const revCards = [
            { icon: icons.ticket, title: "Commission sur trajets", body: "Frais de service de 12% par siège réservé (8% en phase bêta pour stimuler l'adoption initiale).", x: 0.3, y: 1.28 },
            { icon: icons.star, title: "Abonnement Premium", body: "1 500 DZD / mois pour les conducteurs : visibilité accrue des trajets, badges de confiance et statistiques avancées.", x: 5.15, y: 1.28 },
            { icon: icons.handshake, title: "Partenariats B2B", body: "Intégration avec les universités, entreprises et hôtels locaux pour proposer des solutions de covoiturage groupé.", x: 0.3, y: 3.25 },
            { icon: icons.pie, title: "Publicité in-app", body: "Publicité ciblée pour les commerces locaux le long des itinéraires, basée sur un modèle de CPM à 2 000 DZD.", x: 5.15, y: 3.25 }
        ];

        revCards.forEach(({ icon, title, body, x, y }) => {
            s.addShape("roundRect", { x, y, w: 4.55, h: 1.8, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
            s.addShape("roundRect", { x: x + 0.18, y: y + 0.55, w: 0.6, h: 0.6, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
            s.addImage({ data: icon, x: x + 0.23, y: y + 0.6, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.2, y: y + 0.12, w: 4.2, h: 0.3, fontSize: 14, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: x + 0.95, y: y + 0.55, w: 3.4, h: 1.15, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 34 — Business - Prévisions Financières & Conformité 1275 (Fusionné)
    // =========================================================================
    {
        const s = createSlide("FINANCES");

        s.addText("Prévisions financières & conformité Arrêté 1275", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Projections Financières
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.trending, x: 0.7, y: 1.55, w: 0.35, h: 0.35 });
        s.addText("Viabilité Financière", { x: 1.2, y: 1.57, w: 3.5, h: 0.3, fontSize: 14, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• Coûts Année 1 : 15 455 000 DZD (équipe principale 9.6M, acquisition marketing 2.4M, cloud/SMS 1.6M).\n• Projections de revenus : A1 = 6.7M DZD | A3 = 92M DZD.\n• Seuil de rentabilité (Break-even) : Atteint début Q2 de l'Année 2.\n• Lifetime Value / CAC : Ratio LTV/CAC exceptionnel de 19x.\n• Bénéfice net cumulé sur 36 mois : > 52M DZD.", {
            x: 0.7, y: 2.0, w: 4.0, h: 2.8, fontSize: 9.8, color: C.mutedText, fontFace: "Calibri"
        });

        // Right Panel: 1275 Conformity
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addImage({ data: icons.award, x: 5.3, y: 1.55, w: 0.35, h: 0.35 });
        s.addText("Cadre Arrêté 1275", { x: 5.8, y: 1.57, w: 3.5, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Calibri" });
        s.addText("• Labellisation Startup : Projet conçu pour l'accréditation nationale « Un diplôme, une Startup » par l'incubateur de Tlemcen.\n• Propriété Intellectuelle (INAPI) : Dépôt officiel de la marque 'RohWinBghit' finalisé.\n• Protection Logique (ONDA) : Code source Express/FastAPI et modèles convolutifs d'anti-spoofing enregistrés et protégés.\n• Structure Commerciale : SARL en cours de constitution pour exploiter le prototype.", {
            x: 5.3, y: 2.0, w: 4.0, h: 2.8, fontSize: 9.8, color: C.mutedText, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 35 — Transition Section 9 : Conclusion
    // =========================================================================
    createTransitionSlide("CONCLUSION", "9. Conclusion & Perspectives", "Apports et perspectives ?");

    // =========================================================================
    // SLIDE 36 — Conclusion - Bilan & Perspectives (Fusionné)
    // =========================================================================
    {
        const s = createSlide("CONCLUSION");

        s.addText("Bilan général et perspectives futures", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Left Panel: Bilan/Contributions
        s.addShape("roundRect", { x: 0.5, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Bilan & Contributions", { x: 0.7, y: 1.55, w: 4.0, h: 0.3, fontSize: 15, bold: true, color: C.darkGreen, fontFace: "Calibri" });
        s.addText("• Génie Logiciel : Architecture distribuée résiliente aux pertes 4G (taux de succès de 99.9% sous instabilité réseau).\n• Sécurité Biométrique : Modèles d'IA active d'anti-spoofing éliminant le vol d'identité en ligne.\n• Faisabilité Commerciale : Viabilité économique validée par business plan Arrêté 1275 et score SUS de 71.6.", {
            x: 0.7, y: 2.0, w: 4.0, h: 2.8, fontSize: 10.2, color: C.mutedText, fontFace: "Calibri"
        });

        // Right Panel: Perspectives
        s.addShape("roundRect", { x: 5.1, y: 1.4, w: 4.4, h: 3.6, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addText("Perspectives d'Évolution", { x: 5.3, y: 1.55, w: 4.0, h: 0.3, fontSize: 15, bold: true, color: C.gold, fontFace: "Calibri" });

        const roadmaps = [
            ["Court Terme", "Pilote sur l'axe Tlemcen-Oran-Alger & labellisation Startup d'État."],
            ["Moyen Terme", "IA de Surge Pricing avancée et partenariats B2B avec les universités."],
            ["Long Terme", "Expansion nationale sur les 58 wilayas, puis à l'échelle de la zone MENA."]
        ];

        roadmaps.forEach(([phase, desc], i) => {
            const cy = 1.95 + i * 1.0;
            s.addShape("roundRect", { x: 5.3, y: cy, w: 4.0, h: 0.88, fill: { color: C.cream }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addText("🎯 " + phase, { x: 5.42, y: cy + 0.05, w: 3.76, h: 0.22, fontSize: 10, bold: true, color: C.darkGreen, fontFace: "Calibri" });
            s.addText(desc, { x: 5.42, y: cy + 0.28, w: 3.76, h: 0.55, fontSize: 8.8, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // =========================================================================
    // SLIDE 37 — Conclusion - Remerciements & Questions
    // =========================================================================
    {
        currentSlideNum = 37;
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 37, true);
        addSectionTag(s, "CONCLUSION", true);
        addPillCounter(s, "37/37", true);

        // RohWinBghit Logo Main
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo.png"), x: 3.28, y: 1.0, w: 1.1, h: 1.1 });

        // RohWinBghit Wordmark & Bus
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo_bus.png"), x: 4.69, y: 1.38, w: 0.55, h: 0.55 });
        s.addText("RohWinBghit", {
            x: 5.31, y: 1.38, w: 2.03, h: 0.55,
            fontSize: 32, bold: true, color: C.cream,
            fontFace: "Calibri", valign: "middle"
        });

        // Title "Merci"
        s.addText("Merci", {
            x: 0, y: 2.35, w: 10.0, h: 0.7,
            fontSize: 54, bold: true, color: C.cream, align: "center", fontFace: "Calibri"
        });
        s.addText("Merci pour votre attention. Place aux questions / réponses.", {
            x: 0, y: 3.05, w: 10.0, h: 0.35,
            fontSize: 18, bold: false, color: C.accentGreen, align: "center", fontFace: "Calibri"
        });

        // Authors Card (Left Side)
        s.addText("Présenté par :", { x: 1.4, y: 3.65, w: 3.28, h: 0.22, fontSize: 11, bold: true, color: C.accentGreen, align: "right", fontFace: "Calibri" });
        s.addText("AHMED BACHA Djamel Eddine", { x: 1.4, y: 3.91, w: 3.28, h: 0.26, fontSize: 13, bold: false, color: C.cream, align: "right", fontFace: "Calibri" });
        s.addText("BELHORMA Sidi Mohammed Reduane", { x: 1.4, y: 4.19, w: 3.28, h: 0.26, fontSize: 13, color: C.cream, align: "right", fontFace: "Calibri" });

        // Divider
        s.addShape("rect", { x: 4.99, y: 3.75, w: 0.01, h: 0.55, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });

        // Supervisor Card (Right Side)
        s.addText("Encadré par :", { x: 5.31, y: 3.65, w: 3.28, h: 0.22, fontSize: 11, bold: true, color: C.accentGreen, align: "left", fontFace: "Calibri" });
        s.addText("Mme BENLEDGHEM Rafika", { x: 5.31, y: 3.91, w: 3.28, h: 0.26, fontSize: 13, color: C.cream, align: "left", fontFace: "Calibri" });
        s.addText("Département d'Informatique\nUniversité de Tlemcen", { x: 5.31, y: 4.19, w: 3.28, h: 0.36, fontSize: 10.5, color: C.cream, transparency: 20, align: "left", fontFace: "Calibri" });

        // CTA Banner
        s.addShape("rect", { x: 0, y: 4.84, w: 10.0, h: 0.78, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addText("Demandez le mémoire complet ou une démonstration en direct", {
            x: 0, y: 4.84, w: 10.0, h: 0.78,
            fontSize: 18, bold: true, color: C.deepBlack,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ─── WRITE FILE ──────────────────────────────────────────────────────────
    const outPath = path.join(__dirname, "RohWinBghit_Presentation.pptx");
    await pres.writeFile({ fileName: outPath });
    console.log("✅  Fichier généré :", outPath);
}

buildPresentation().catch(err => { console.error("❌ Erreur :", err); process.exit(1); });