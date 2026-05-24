/**
 * RohWinBghit — Présentation PPTX
 * Génère la présentation complète à partir des slides HTML
 * Usage: node generate_rohwinbghit.js
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
    FaUserLarge, FaMagnifyingGlassLocation, FaCalendarCheck, FaCirclePlus, FaUsersGear
} = require("react-icons/fa6");
const { FaNodeJs, FaReact } = require("react-icons/fa");
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
// Slide dimensions: 10" × 5.625" (16x9)
const SW = 10, SH = 5.625;

function addSectionTag(slide, label, dark = false) {
    // Accent line
    slide.addShape("rect", {
        x: 0.5, y: 0.45, w: 0.32, h: 0.035,
        fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 }
    });
    slide.addText(label, {
        x: 0.88, y: 0.38, w: 3, h: 0.2,
        fontSize: 11, bold: true, color: dark ? C.accentGreen : C.midGreen,
        fontFace: "Courier New", charSpacing: 2
    });
}

function addPillCounter(slide, label, dark = false) {
    if (label && label.includes('/')) {
        const current = label.split('/')[0];
        label = `${current}/28`;
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

function addIconCard(slide, icon, title, body, x, y, w = 3.75, h = 0.86) {
    slide.addShape("roundRect", {
        x, y, w, h,
        fill: { color: C.white },
        line: { color: C.cardBorder, width: 0.5 },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
        rectRadius: 0.1
    });
    // Icon circle
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

// ─── SLIDES ──────────────────────────────────────────────────────────────────

async function buildPresentation() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane";
    pres.title = "RohWinBghit — Plateforme de covoiturage";

    // Pre-render frequently used icons
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
        route: await iconPng(FaRoute, "#4CC080"),
        mapIcon: await iconPng(FaMap, "#4CC080"),
        bolt: await iconPng(FaBoltLightning, "#166E3A"),
        vial: await iconPng(FaVialCircleCheck, "#166E3A"),
        database: await iconPng(FaDatabase, "#166E3A"),
        shieldChk: await iconPng(FaShield, "#4CC080"),
        micro: await iconPng(FaMicrochip, "#0A1F14"),
        quote: await iconPng(FaQuoteRight, "#0A1F14"),
        car: await iconPng(FaCarSide, "#0A1F14"),
        brief: await iconPng(FaBriefcase, "#0A1F14"),
        users: await iconPng(FaUsers, "#4CC080"),
        clock: await iconPng(FaClock, "#4CC080"),
        node: await iconPng(FaNodeJs, "#166E3A"),
        react: await iconPng(FaReact, "#166E3A"),
        font: await iconPng(FaMoneyBillWave, "#166E3A"),
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
    };

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 01 — COVER
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 1, true);
        slide01_sectionTag(s);
        addPillCounter(s, "01/23", true);

        // University Logo
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "univ_logo.png"), x: 4.61, y: 0.94, w: 0.78, h: 0.78 });

        // RohWinBghit Logo Bus
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo_bus.png"), x: 2.19, y: 2.03, w: 0.78, h: 0.78 });

        // RohWinBghit Wordmark
        s.addText("RohWinBghit", {
            x: 3.13, y: 1.95, w: 4.69, h: 0.94,
            fontSize: 54, bold: true, color: C.cream,
            fontFace: "Calibri", valign: "middle"
        });

        // Title
        s.addText("Plateforme de covoiturage intelligente et sécurisée\npour l'Algérie inter-wilayas", {
            x: 0.5, y: 3.2, w: 9.0, h: 1.25,
            fontSize: 28, bold: true, color: C.cream,
            align: "center", fontFace: "Calibri"
        });

        // Subtitle
        s.addText("Conception et réalisation (Master GL)", {
            x: 0.5, y: 4.5, w: 9.0, h: 0.35,
            fontSize: 18, color: C.accentGreen, align: "center", fontFace: "Calibri"
        });

        // Bottom banner
        s.addShape("rect", { x: 0, y: 5.12, w: SW, h: 0.5, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });
        s.addText("SIMPLE   •   SÛR   •   ABORDABLE", {
            x: 0, y: 5.12, w: SW, h: 0.5,
            fontSize: 14, bold: true, color: C.cream,
            align: "center", valign: "middle", charSpacing: 3, fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 02 — TABLE DES MATIÈRES
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 2, false);
        addSectionTag(s, "APERÇU");
        addPillCounter(s, "02/23");

        s.addText("Table des matières", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const chapters = [
            ["01", "Contexte & Marché"],
            ["02", "La Plateforme"],
            ["03", "Technologie"],
            ["04", "Sécurité & Validation"],
            ["05", "Modèle d'Affaires"],
            ["06", "Feuille de Route"],
        ];
        const cols = [[0.45, 1.42], [5.15, 6.12]];
        const rows = [1.5, 2.4, 3.3];
        chapters.forEach(([num, title], i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const x = cols[col][0], y = rows[row];
            s.addShape("roundRect", { x, y, w: 4.4, h: 0.78, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addShape("ellipse", { x: x + 0.12, y: y + 0.19, w: 0.4, h: 0.4, fill: { color: C.cream }, line: { color: C.cream, width: 0 } });
            s.addText(num, { x: x + 0.12, y: y + 0.19, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: C.midGreen, fontFace: "Courier New", align: "center", valign: "middle", margin: 0, wrap: false });
            s.addText(title, { x: x + 0.65, y: y + 0.25, w: 3.6, h: 0.3, fontSize: 13.5, bold: true, color: C.deepBlack, fontFace: "Calibri", valign: "middle" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 03 — PROBLÈME
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 3, false);
        addSectionTag(s, "CONTEXTE & MARCHÉ");
        addPillCounter(s, "03/23");

        s.addText("Problème — Défis du transport inter-wilayas", {
            x: 0.5, y: 0.75, w: 9, h: 0.5,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const problems = [
            [icons.money, "Coût & Disponibilité", "Tarifs élevés, disponibilité limitée et pics saisonniers perturbent les trajets longue distance."],
            [icons.network, "Friction de Réservation", "Les réseaux informels fragmentés créent une friction massive lors de la réservation."],
            [icons.shield, "Confiance & Sécurité", "Des déficits de confiance et l'absence d'identités vérifiées favorisent la méfiance entre usagers."],
            [icons.map, "Coordination Complexe", "La gestion des trajets à travers 69 wilayas est très difficile sans outils numériques adaptés."],
            [icons.card, "Limites de Paiement", "L'adoption limitée du paiement électronique restreint fortement les transactions fluides."],
        ];

        const positions = [
            [0.5, 1.38], [3.9, 1.38], [7.3, 1.38],
            [1.9, 2.7], [5.3, 2.7],
        ];
        problems.forEach(([icon, title, body], i) => {
            const [x, y] = positions[i];
            addIconCard(s, icon, title, body, x, y, 3.15, 1.15);
        });

        s.addText("Source : Introduction / Résumé", {
            x: 0.5, y: 5.35, w: 5, h: 0.2,
            fontSize: 8, color: C.mutedText, fontFace: "Courier New"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 04 — SOLUTION
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 4, false);
        addSectionTag(s, "SOLUTION");
        addPillCounter(s, "04/23");

        s.addText("Solution — Qu'est-ce que RohWinBghit ?", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const solutions = [
            [icons.mobile, "Plateforme mobile", "Solution dédiée au covoiturage inter-wilayas parfaitement adaptée au marché algérien."],
            [icons.map, "Réservation temps réel & Suivi", "Réservations instantanées avec suivi GPS en direct via l'intégration Mapbox."],
            [icons.idCard, "Identités vérifiées (KYC)", "Vérification biométrique IA avec détection de vivacité et anti-spoofing."],
            [icons.chat, "Communication instantanée", "Coordination fluide conducteur-passager via chat in-app (Socket.IO) et notifications."],
            [icons.wallet, "Paiements multi-méthodes", "Espèces opérationnel, avec CIB et Edahabia intégrés en sandbox (en attente SATIM)."],
            [icons.chart, "Économie intelligente & Sécurité", "Algorithmes de tarification dynamique couplés à une détection proactive de fraude."],
        ];

        const row1y = 1.38, row2y = 2.55, row3y = 3.72;
        const col1x = 0.5, col2x = 5.3;
        const positions = [
            [col1x, row1y], [col1x, row2y], [col1x, row3y],
            [col2x, row1y], [col2x, row2y], [col2x, row3y],
        ];
        solutions.forEach(([icon, title, body], i) => {
            const [x, y] = positions[i];
            addIconCard(s, icon, title, body, x, y, 4.55, 0.95);
        });

        // Bottom dark banner
        s.addShape("rect", { x: 0, y: 4.82, w: SW, h: 0.8, fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 } });
        s.addText("SIMPLE  •  SÛR  •  ÉCONOMIQUE", {
            x: 0, y: 4.82, w: SW, h: 0.8,
            fontSize: 16, bold: true, color: C.accentGreen,
            align: "center", valign: "middle", charSpacing: 5, fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 05 — SECTION DIVIDER (CONTEXTE)
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 5, true);
        addSectionTag(s, "CONTEXTE & MARCHÉ", true);
        addPillCounter(s, "05/23", true);

        // Icon circle
        s.addShape("ellipse", { x: 4.72, y: 1.55, w: 0.55, h: 0.55, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addImage({ data: icons.quote, x: 4.82, y: 1.65, w: 0.35, h: 0.35 });

        // Quote card
        s.addShape("roundRect", { x: 1.4, y: 1.95, w: 7.2, h: 1.6, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.15 });
        s.addText('"Absence d\'une solution numérique sécurisée et de confiance\npour les déplacements longue distance en Algérie."', {
            x: 1.6, y: 2.1, w: 6.8, h: 1.3,
            fontSize: 16, italic: true, color: C.cream,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 06 — OPPORTUNITÉ
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 6, false);
        addSectionTag(s, "MARCHÉ");
        addPillCounter(s, "06/23");

        s.addText("Opportunité en Algérie", {
            x: 0.5, y: 0.75, w: 9, h: 0.5,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Stat circle 1
        s.addShape("ellipse", { x: 0.5, y: 1.5, w: 1.6, h: 1.6, fill: { color: C.darkGreen }, line: { color: C.accentGreen, width: 3 } });
        s.addText("69\nWILAYAS", {
            x: 0.5, y: 1.5, w: 1.6, h: 1.6,
            fontSize: 22, bold: true, color: C.accentGreen,
            align: "center", valign: "middle", fontFace: "Calibri", margin: 0, wrap: false
        });
        s.addText("Couverture nationale ciblée", { x: 2.4, y: 1.7, w: 7, h: 0.3, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        s.addText("Résolution des problèmes majeurs liés au coût, à la disponibilité et à la confiance pour le transport entre les 69 wilayas.", {
            x: 2.4, y: 2.05, w: 7, h: 0.7,
            fontSize: 11.5, color: C.mutedText, fontFace: "Calibri"
        });

        // Stat circle 2
        s.addShape("ellipse", { x: 0.5, y: 3.4, w: 1.6, h: 1.6, fill: { color: C.darkGreen }, line: { color: C.accentGreen, width: 3 } });
        s.addImage({ data: icons.trend, x: 0.9, y: 3.7, w: 0.8, h: 0.7 });
        s.addText("ADOPTION", { x: 0.5, y: 4.4, w: 1.6, h: 0.3, fontSize: 9, bold: true, color: C.cream, align: "center", fontFace: "Calibri" });
        s.addText("Adoption numérique en hausse", { x: 2.4, y: 3.6, w: 7, h: 0.3, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        s.addText("Authentification OTP pour faciliter l'onboarding. Englobe toutes les parties prenantes : passagers, conducteurs, administrateurs et régulateurs.", {
            x: 2.4, y: 3.95, w: 7, h: 0.7,
            fontSize: 11.5, color: C.mutedText, fontFace: "Calibri"
        });

        s.addText("Réf. : Chapitre 1 (Contexte, opportunité)", { x: 0.5, y: 5.38, w: 6, h: 0.2, fontSize: 8, color: C.mutedText, fontFace: "Courier New" });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 07 — ANALYSE CONCURRENTIELLE
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 7, false);
        addSectionTag(s, "CONCURRENCE");
        addPillCounter(s, "07/23");

        s.addText("Solutions existantes — Vue comparative", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Table card background
        s.addShape("roundRect", {
            x: 0.5, y: 1.35, w: 9, h: 3.95,
            fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.04 },
            rectRadius: 0.1
        });

        // Column headers
        const colX = [0.75, 3.35, 5.65, 7.5];
        const colW = [2.4, 2.1, 1.8, 2.25];
        const headers = ["Critères d'évaluation", "International\n(Limité paiements/KYC)", "Maghreb/Algérie\n(Fragmenté, déficits)", "RohWinBghit"];
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

        // Header separator
        s.addShape("rect", { x: 0.6, y: 2.07, w: 8.8, h: 0.02, fill: { color: C.cardBorder }, line: { color: C.cardBorder, width: 0 } });

        // Rows
        const rows2 = [
            ["Confiance & Identité (KYC)", "✗  Pas d'ID local", "~  Basique / Manuel", "✓  Vérification IA Biométrique"],
            ["Intégration Paiement Local", "✗  International uniquement", "~  Partiel / Surtout espèces", "✓  Espèces + CIB/Edahabia"],
            ["Couverture Inter-Wilayas", "✓  Étendue (Globale)", "~  Surtout centres urbains", "✓  Dédiée 69 Wilayas"],
            ["UX Temps Réel & Suivi GPS", "✓  Applications avancées", "✗  Forums/Groupes", "✓  Mapbox + Sockets en direct"],
        ];
        const rowColors = ["D9534F", "F0AD4E", C.accentGreen];
        rows2.forEach(([crit, intl, magh, rwb], ri) => {
            const ry = 2.12 + ri * 0.82;
            const isLast = ri === rows2.length - 1;
            if (ri < rows2.length - 1) {
                s.addShape("rect", { x: 0.6, y: ry + 0.78, w: 8.8, h: 0.01, fill: { color: C.cardBorder }, line: { color: C.cardBorder, width: 0 } });
            }
            s.addText(crit, { x: colX[0] + 0.05, y: ry + 0.2, w: colW[0] - 0.1, h: 0.4, fontSize: 10, bold: true, color: C.deepBlack, fontFace: "Calibri" });

            [[intl, 1], [magh, 2]].forEach(([txt, ci]) => {
                const sym = txt[0];
                const clr = sym === "✗" ? "D9534F" : sym === "~" ? "F0AD4E" : C.accentGreen;
                s.addText(txt, { x: colX[ci] + 0.05, y: ry + 0.2, w: colW[ci] - 0.1, h: 0.4, fontSize: 10, bold: false, color: C.mutedText, fontFace: "Calibri" });
            });

            // RWB cell (highlighted)
            s.addShape("rect", { x: colX[3], y: ry + 0.02, w: colW[3], h: isLast ? 0.75 : 0.78, fill: { color: C.accentGreen, transparency: 85 }, line: { color: C.accentGreen, width: 0.5 } });
            s.addText(rwb, { x: colX[3] + 0.05, y: ry + 0.2, w: colW[3] - 0.1, h: 0.4, fontSize: 10, bold: true, color: C.midGreen, fontFace: "Calibri" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 08 — SECTION DIVIDER (LA PLATEFORME)
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 8, true);
        addSectionTag(s, "LA PLATEFORME", true);
        addPillCounter(s, "08/23", true);

        s.addShape("ellipse", { x: 4.72, y: 1.55, w: 0.55, h: 0.55, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addImage({ data: icons.car, x: 4.81, y: 1.63, w: 0.37, h: 0.37 });

        s.addShape("roundRect", { x: 1.4, y: 1.95, w: 7.2, h: 1.6, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.15 });
        s.addText('"Votre trajet, notre priorité.\nLa solution de covoiturage simple, sûre et économique en Algérie."', {
            x: 1.6, y: 2.1, w: 6.8, h: 1.3,
            fontSize: 15, italic: false, bold: false, color: C.cream,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 09 — FONCTIONNALITÉS
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 9, false);
        addSectionTag(s, "FONCTIONNALITÉS");
        addPillCounter(s, "09/23");

        s.addText("Fonctionnalités principales (Vue d'ensemble de l'application)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Phone Mockup 1 (Welcome Screen)
        s.addShape("roundRect", { x: 1.09, y: 1.72, w: 1.72, h: 3.28, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Screen1.jpg"), x: 1.14, y: 1.77, w: 1.62, h: 3.18 });
        s.addShape("roundRect", { x: 1.65, y: 1.83, w: 0.6, h: 0.15, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        // Phone Mockup 2 (Value Props)
        s.addShape("roundRect", { x: 3.13, y: 1.72, w: 1.72, h: 3.28, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Screen2.jpg"), x: 3.18, y: 1.77, w: 1.62, h: 3.18 });
        s.addShape("roundRect", { x: 3.69, y: 1.83, w: 0.6, h: 0.15, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        // Feature List Title
        s.addText("Conçu pour la confiance et la rapidité", {
            x: 5.31, y: 1.88, w: 3.91, h: 0.45,
            fontSize: 18, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Feature Cards
        const feats = [
            [icons.userCheck, "Vérification d'identité IA", "La reconnaissance faciale par IA garantit que chaque utilisateur est vérifié et légitime."],
            [icons.map, "Réservation temps réel et suivi GPS", "Trouvez, réservez en quelques secondes et suivez votre trajet en direct sur la carte."],
            [icons.chat, "Messagerie et paiements multiples", "Échangez via la messagerie instantanée et payez en espèces ou par carte (CIB, Edahabia)."],
        ];

        feats.forEach(([icon, title, body], i) => {
            const y = 2.50 + i * 0.86;
            s.addShape("roundRect", { x: 5.31, y, w: 3.91, h: 0.70, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.08 });
            s.addShape("roundRect", { x: 5.47, y: y + 0.15, w: 0.39, h: 0.39, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.06 });
            s.addImage({ data: icon, x: 5.52, y: y + 0.20, w: 0.29, h: 0.29 });
            s.addText(title, { x: 5.98, y: y + 0.08, w: 3.1, h: 0.22, fontSize: 11.5, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: 5.98, y: y + 0.32, w: 3.1, h: 0.35, fontSize: 8.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Bottom banner
        s.addShape("rect", { x: 0, y: 5.12, w: SW, h: 0.5, fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 } });
        s.addText("24/7 ASSISTANCE  •  CONDUCTEURS VÉRIFIÉS  •  RÉSERVATION INSTANTANÉE", {
            x: 0, y: 5.12, w: SW, h: 0.5,
            fontSize: 11, bold: true, color: C.accentGreen,
            align: "center", valign: "middle", charSpacing: 2, fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 10 — RÔLES
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 10, false);
        addSectionTag(s, "RÔLES");
        addPillCounter(s, "10/23");

        s.addText("Rôles utilisateurs — Passager vs Conducteur", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Passenger Card Background
        s.addShape("roundRect", { x: 0.47, y: 1.72, w: 3.28, h: 3.44, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        // Passenger Icon & Title
        s.addShape("roundRect", { x: 0.70, y: 1.95, w: 0.5, h: 0.5, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.userLarge, x: 0.76, y: 2.01, w: 0.38, h: 0.38 });
        s.addText("Passager", { x: 1.36, y: 2.07, w: 2.0, h: 0.3, fontSize: 16, bold: true, color: C.deepBlack, fontFace: "Calibri" });

        // Passenger Bullets
        const pItems = [
            [icons.magnifyingGlass, "Rechercher et trouver des trajets facilement à travers les wilayas."],
            [icons.calendarCheck, "Réserver une place avec confirmation instantanée."],
            [icons.cardG, "Payer en toute sécurité (Espèces, CIB, Edahabia)."],
            [icons.routeG, "Suivre le trajet en temps réel via l'intégration GPS."]
        ];
        pItems.forEach(([icon, text], i) => {
            const by = 2.65 + i * 0.56;
            s.addImage({ data: icon, x: 0.70, y: by + 0.05, w: 0.22, h: 0.22 });
            s.addText(text, { x: 1.05, y: by, w: 2.5, h: 0.45, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Center Phone Mockup
        s.addShape("roundRect", { x: 4.06, y: 1.41, w: 1.88, h: 3.91, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.15 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "inscription.jpg"), x: 4.11, y: 1.46, w: 1.78, h: 3.81 });
        s.addShape("roundRect", { x: 4.70, y: 1.51, w: 0.6, h: 0.15, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        // Driver Card Background
        s.addShape("roundRect", { x: 6.25, y: 1.72, w: 3.28, h: 3.44, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        // Driver Icon & Title
        s.addShape("roundRect", { x: 6.48, y: 1.95, w: 0.5, h: 0.5, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.steeringWheel, x: 6.54, y: 2.01, w: 0.38, h: 0.38 });
        s.addText("Conducteur", { x: 7.14, y: 2.07, w: 2.0, h: 0.3, fontSize: 16, bold: true, color: C.deepBlack, fontFace: "Calibri" });

        // Driver Bullets
        const cItems = [
            [icons.circlePlus, "Publier des trajets et planifier des itinéraires."],
            [icons.usersGear, "Gérer les réservations et les demandes dynamiquement."],
            [icons.shieldChk, "Vérifier le véhicule et les documents de manière sécurisée."],
            [icons.chatG, "Recevoir les paiements et communiquer avec les passagers."]
        ];
        cItems.forEach(([icon, text], i) => {
            const by = 2.65 + i * 0.56;
            s.addImage({ data: icon, x: 6.48, y: by + 0.05, w: 0.22, h: 0.22 });
            s.addText(text, { x: 6.83, y: by, w: 2.5, h: 0.45, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 11 — DIAGRAMME CAS D'UTILISATION
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 11, false);
        addSectionTag(s, "ANALYSE");
        addPillCounter(s, "11/23");

        s.addText("Diagramme de Cas d'Utilisation", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Styled Container Card for Diagram
        s.addShape("roundRect", {
            x: 0.5, y: 1.3, w: 9.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // PlantUML Use Case Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_use_case.png"),
            x: 0.6, y: 1.4, w: 8.8, h: 3.8,
            sizing: { type: "contain", w: 8.8, h: 3.8 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 12 — DIAGRAMME DE CLASSES
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 12, true);
        addSectionTag(s, "CONCEPTION", true);
        addPillCounter(s, "12/23", true);

        s.addText("Diagramme de Classes du Domaine", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.cream, fontFace: "Calibri"
        });

        // Styled Container Card for Class Diagram
        s.addShape("roundRect", {
            x: 0.5, y: 1.3, w: 9.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.midGreen, width: 1.5 },
            shadow: { type: "outer", blur: 10, offset: 3, angle: 135, color: "000000", opacity: 0.2 },
            rectRadius: 0.1
        });

        // PlantUML Class Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_class_diagram.png"),
            x: 0.6, y: 1.4, w: 8.8, h: 3.8,
            sizing: { type: "contain", w: 8.8, h: 3.8 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 13 — SÉQUENCE INSCRIPTION OTP
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 13, false);
        addSectionTag(s, "SÉQUENCE");
        addPillCounter(s, "13/23");

        s.addText("Diagramme de Séquence — Inscription (OTP)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 20, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Centered Card for Diagram
        s.addShape("roundRect", {
            x: 1.5, y: 1.3, w: 7.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // PlantUML Sequence Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_sequence_inscription.png"),
            x: 1.6, y: 1.4, w: 6.8, h: 3.8,
            sizing: { type: "contain", w: 6.8, h: 3.8 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 14 — SÉQUENCE RÉSERVATION
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 14, false);
        addSectionTag(s, "SÉQUENCE");
        addPillCounter(s, "14/23");

        s.addText("Séquence — Réservation avec Paiement Multi-méthodes", {
            x: 0.3, y: 0.72, w: 9.4, h: 0.4,
            fontSize: 18, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Styled Container Card for Diagram
        s.addShape("roundRect", {
            x: 0.5, y: 1.3, w: 9.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // PlantUML Sequence Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_sequence_booking.png"),
            x: 0.6, y: 1.4, w: 8.8, h: 3.8,
            sizing: { type: "contain", w: 8.8, h: 3.8 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 15 — FLUX D'ACTIVITÉ
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 15, false);
        addSectionTag(s, "ACTIVITÉ");
        addPillCounter(s, "15/23");

        s.addText("Flux complet d'une réservation (Diagramme d'Activité)", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 20, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Centered Card for Diagram
        s.addShape("roundRect", {
            x: 1.5, y: 1.3, w: 7.0, h: 4.0,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // PlantUML Activity Diagram
        s.addImage({
            path: path.join(__dirname, "plantuml", "output", "fig_activity_main.png"),
            x: 1.6, y: 1.4, w: 6.8, h: 3.8,
            sizing: { type: "contain", w: 6.8, h: 3.8 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 16 — KYC BIOMÉTRIQUE
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 16, true);
        addSectionTag(s, "SÉCURITÉ", true);
        addPillCounter(s, "16/23", true);

        s.addText("Identité sécurisée — KYC biométrique", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.cream, fontFace: "Calibri"
        });

        // Left Cards
        const kycItems = [
            [icons.brain, "Vérification IA (FastAPI)", "Détection de vivacité en temps réel et anti-spoofing avancé pour prévenir la fraude."],
            [icons.idCardG, "Document & Visage", "Scan de document et correspondance faciale fluide avec le selfie de l'utilisateur."],
            [icons.shieldG, "Conformité Légale", "Traitement sécurisé avec consentement explicite, conforme aux Lois 18-07 & 25-11."],
        ];
        kycItems.forEach(([icon, title, body], i) => {
            const y = 1.45 + i * 1.2;
            s.addShape("roundRect", { x: 0.5, y, w: 3.8, h: 1.1, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.2 }, rectRadius: 0.1 });
            s.addShape("roundRect", { x: 0.65, y: y + 0.32, w: 0.46, h: 0.46, fill: { color: C.accentGreen, transparency: 90 }, line: { color: C.midGreen, width: 0.5 }, rectRadius: 0.08 });
            s.addImage({ data: icon, x: 0.70, y: y + 0.37, w: 0.36, h: 0.36 });
            s.addText(title, { x: 1.25, y: y + 0.1, w: 2.9, h: 0.25, fontSize: 12, bold: true, color: C.cream, fontFace: "Calibri" });
            s.addText(body, { x: 1.25, y: y + 0.38, w: 2.9, h: 0.65, fontSize: 9.5, color: C.cream, transparency: 30, fontFace: "Calibri" });
        });

        // Phone 1 (Camera Permission)
        s.addShape("roundRect", { x: 4.60, y: 1.45, w: 1.5, h: 3.5, fill: { color: "000000" }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.08 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "inscription.jpg"), x: 4.64, y: 1.49, w: 1.42, h: 3.42 });
        s.addShape("roundRect", { x: 5.07, y: 1.55, w: 0.56, h: 0.12, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.06 });

        // Phone 2 (Identity Step)
        s.addShape("roundRect", { x: 6.20, y: 1.45, w: 1.5, h: 3.5, fill: { color: "000000" }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.08 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Screen1.jpg"), x: 6.24, y: 1.49, w: 1.42, h: 3.42 });
        s.addShape("roundRect", { x: 6.67, y: 1.55, w: 0.56, h: 0.12, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.06 });

        // Diagram KYC
        s.addShape("roundRect", { x: 7.80, y: 1.45, w: 1.7, h: 3.5, fill: { color: C.white }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.08 });
        s.addImage({
            path: "https://www.genspark.ai/api/files/s/RWCKoQSQ",
            x: 7.85, y: 1.50, w: 1.6, h: 3.40,
            sizing: { type: "contain", w: 1.6, h: 3.40 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 17 — GPS & RÉSERVATION TEMPS RÉEL
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 17, false);
        addSectionTag(s, "EXPÉRIENCE");
        addPillCounter(s, "17/23");

        s.addText("Réservation temps réel et suivi GPS", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const gpsFeats = [
            [icons.mapG, "Aperçu de l'itinéraire en direct", "Propulsé par Mapbox pour une cartographie inter-wilayas ultra-précise, offrant des distances dynamiques et des ETA exacts."],
            [icons.stopwatch, "Réservation instantanée", "Recherchez, comparez et réservez des places disponibles en quelques secondes grâce à un algorithme optimisé en temps réel."],
            [icons.crosshair, "Suivi GPS en temps réel", "Suivez l'intégralité du trajet en direct sur la carte, du départ à l'arrivée, garantissant transparence et sécurité pour tous."],
        ];
        gpsFeats.forEach(([icon, title, body], i) => {
            const y = 1.3 + i * 1.2;
            s.addShape("roundRect", { x: 0.3, y: y + 0.12, w: 0.55, h: 0.55, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
            s.addImage({ data: icon, x: 0.38, y: y + 0.2, w: 0.39, h: 0.39 });
            s.addText(title, { x: 1.0, y, w: 4.35, h: 0.35, fontSize: 14, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: 1.0, y: y + 0.37, w: 4.35, h: 0.72, fontSize: 10.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Phone Mockup 1 (Trip Publication)
        s.addShape("roundRect", { x: 5.16, y: 1.56, w: 1.88, h: 3.59, fill: { color: "000000" }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Trajet publié.jpg"), x: 5.22, y: 1.63, w: 1.75, h: 3.47 });
        s.addShape("roundRect", { x: 5.80, y: 1.70, w: 0.60, h: 0.16, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });

        // Phone Mockup 2 (Results List)
        s.addShape("roundRect", { x: 7.50, y: 1.56, w: 1.88, h: 3.59, fill: { color: "000000" }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "QR scanné au départ.jpg"), x: 7.56, y: 1.63, w: 1.75, h: 3.47 });
        s.addShape("roundRect", { x: 8.14, y: 1.70, w: 0.60, h: 0.16, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 18 — PAIEMENTS
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 18, false);
        addSectionTag(s, "PAIEMENTS");
        addPillCounter(s, "18/23");

        s.addText("Méthodes de paiement et architecture", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Strategy card
        s.addShape("roundRect", { x: 0.3, y: 1.28, w: 6.6, h: 1.1, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: 0.48, y: 1.43, w: 0.7, h: 0.7, fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.code, x: 0.55, y: 1.5, w: 0.56, h: 0.56 });
        s.addText("Motif Strategy / Factory", { x: 1.3, y: 1.35, w: 5.45, h: 0.35, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        s.addText("Moteur de paiement multi-méthodes extensible conçu pour une intégration transparente de nouvelles passerelles.", {
            x: 1.3, y: 1.72, w: 5.45, h: 0.55, fontSize: 10.5, color: C.mutedText, fontFace: "Calibri"
        });

        // Cash card (OPÉRATIONNEL)
        s.addShape("roundRect", { x: 0.3, y: 2.55, w: 6.6, h: 1.1, fill: { color: C.white }, line: { color: C.accentGreen, width: 1.5 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: 0.48, y: 2.7, w: 0.7, h: 0.7, fill: { color: C.accentGreen, transparency: 85 }, line: { color: C.accentGreen, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.money, x: 0.55, y: 2.77, w: 0.56, h: 0.56 });
        s.addText("Espèces", { x: 1.3, y: 2.62, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        s.addShape("roundRect", { x: 2.8, y: 2.66, w: 1.1, h: 0.25, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 }, rectRadius: 0.12 });
        s.addText("Opérationnel", { x: 2.8, y: 2.66, w: 1.1, h: 0.25, fontSize: 8, bold: true, color: C.darkGreen, align: "center", valign: "middle", fontFace: "Calibri" });
        s.addText("Paiements en espèces entièrement fonctionnels pour un déploiement et une adoption immédiats sur le marché.", {
            x: 1.3, y: 2.99, w: 5.45, h: 0.55, fontSize: 10.5, color: C.mutedText, fontFace: "Calibri"
        });

        // CIB & Edahabia (SANDBOX)
        s.addShape("roundRect", { x: 0.3, y: 3.82, w: 6.6, h: 1.1, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: 0.48, y: 3.97, w: 0.7, h: 0.7, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.card, x: 0.55, y: 4.04, w: 0.56, h: 0.56 });
        s.addText("CIB & Edahabia", { x: 1.3, y: 3.89, w: 3.5, h: 0.35, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        s.addShape("roundRect", { x: 3.55, y: 3.93, w: 0.9, h: 0.25, fill: { color: C.cream }, line: { color: C.gold, width: 1 }, rectRadius: 0.12 });
        s.addText("Sandbox", { x: 3.55, y: 3.93, w: 0.9, h: 0.25, fontSize: 8, bold: true, color: C.gold, align: "center", valign: "middle", fontFace: "Calibri" });
        s.addText("Intégrés et testés en environnement sandbox. En attente d'homologation SATIM pour la mise en production.", {
            x: 1.3, y: 4.26, w: 5.45, h: 0.55, fontSize: 10.5, color: C.mutedText, fontFace: "Calibri"
        });

        // Phone Mockup (Ticket Screen)
        s.addShape("roundRect", { x: 7.20, y: 1.15, w: 2.2, h: 4.15, fill: { color: "000000" }, line: { color: C.midGreen, width: 2 }, rectRadius: 0.1 });
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "Passager réserve.jpg"), x: 7.25, y: 1.20, w: 2.10, h: 4.05 });
        s.addShape("roundRect", { x: 7.90, y: 1.28, w: 0.80, h: 0.16, fill: { color: "000000" }, line: { color: "000000", width: 0 }, rectRadius: 0.08 });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 19 — SECTION DIVIDER (TECHNOLOGIE)
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 19, true);
        addSectionTag(s, "TECHNOLOGIE", true);
        addPillCounter(s, "19/23", true);

        s.addShape("ellipse", { x: 4.72, y: 1.55, w: 0.55, h: 0.55, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addImage({ data: icons.quote, x: 4.82, y: 1.63, w: 0.35, h: 0.35 });
        s.addShape("roundRect", { x: 1.4, y: 1.95, w: 7.2, h: 1.6, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.15 });
        s.addText('"Architecture multiplateforme robuste\navec microservice IA isolé."', {
            x: 1.6, y: 2.1, w: 6.8, h: 1.3,
            fontSize: 18, italic: true, color: C.cream,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 20 — ARCHITECTURE SYSTÈME
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 20, true);
        addSectionTag(s, "ARCHITECTURE", true);
        addPillCounter(s, "20/23", true);

        s.addText("Architecture système détaillée", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.cream, fontFace: "Calibri"
        });

        const archBox = (x, y, w, h, title, items, borderClr, bgAlpha = 80) => {
            s.addShape("roundRect", { x, y, w, h, fill: { color: C.deepBlack, transparency: bgAlpha }, line: { color: borderClr, width: 1.5 }, rectRadius: 0.08 });
            s.addText(title, { x: x + 0.1, y: y + 0.08, w: w - 0.2, h: 0.22, fontSize: 9, bold: true, color: borderClr, fontFace: "Courier New" });
            items.forEach((item, i) => {
                s.addText(item, { x: x + 0.12, y: y + 0.35 + i * 0.22, w: w - 0.24, h: 0.22, fontSize: 8.5, color: C.cream, fontFace: "Calibri" });
            });
        };

        // Clients
        archBox(0.3, 1.25, 4.5, 0.95, "📱 CLIENTS",
            ["React Native 0.81 / Expo SDK 54 — App Mobile", "React / Vite — Dashboard Admin Web"], C.accentGreen, 60);

        // External services
        archBox(5.1, 1.25, 4.5, 0.95, "☁️ SERVICES EXTERNES",
            ["Mapbox (GPS)  |  FCM (Notifications)", "SATIM (CIB/Edahabia)  |  Sentry (Monitoring)"], C.gold, 60);

        // Arrow
        s.addText("↕  HTTPS REST + WebSocket", { x: 0.3, y: 2.28, w: 9.3, h: 0.22, fontSize: 9, color: C.accentGreen, align: "center", fontFace: "Calibri" });

        // Backend
        archBox(0.3, 2.55, 6.5, 1.85, "⚙️ BACKEND — Node.js / Express 4.18.2 :8080",
            ["API Gateway → Auth JWT → 19 Routes → 17 Controllers",
                "Services : Surge Pricing, Matching, Fare, Fraud, Heatmap, Notifications, Score",
                "PaymentFactory : Stratégies CIB | Edahabia | Cash",
                "BullMQ 5.71.0 — Workers async & réconciliation"], C.accentGreen, 70);

        // AI microservice
        archBox(7.05, 2.55, 2.7, 1.85, "🧠 MICROSERVICE IA\nFastAPI :8001",
            ["InsightFace buffalo_l (512D)", "LivenessDetector", "AntiSpoofDetector", "CardValidator"], C.accentGreen, 60);

        // Data layer
        archBox(0.3, 4.52, 9.4, 0.95, "💾 COUCHE DONNÉES",
            ["PostgreSQL 16 — 29 tables, Knex 3.1.0   |   Redis 7 — Cache Surge, Sessions TTL=60s   |   Multer (/uploads/)"], C.gold, 70);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 21 — STACK MOBILE
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 21, false);
        addSectionTag(s, "MOBILE");
        addPillCounter(s, "21/23");

        s.addText("Stack de l'application mobile", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const mobileCards = [
            [icons.react, "Framework Principal", "React Native 0.81 accéléré par Expo SDK 54 — développement cross-platform efficace."],
            [icons.node, "Backend Communication", "API REST performante sur Node.js/Express. Knex.js pour requêtes PostgreSQL flexibles."],
            [icons.shield, "Authentification", "Flux OTP fluide garantissant une vérification d'identité sécurisée pour chaque utilisateur."],
            [icons.vial, "Validation Rigoureuse", "62 écrans validés manuellement. 151 tests Jest (107 unitaires + 100 intégration) — 100% réussite."],
        ];

        const positions = [[0.3, 1.28], [5.15, 1.28], [0.3, 3.12], [5.15, 3.12]];
        mobileCards.forEach(([icon, title, body], i) => {
            const [x, y] = positions[i];
            s.addShape("roundRect", { x, y, w: 4.55, h: 1.65, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
            s.addShape("roundRect", { x: x + 0.18, y: y + 0.55, w: 0.6, h: 0.6, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
            s.addImage({ data: icon, x: x + 0.23, y: y + 0.6, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.2, y: y + 0.12, w: 4.2, h: 0.3, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });
            s.addText(body, { x: x + 0.2, y: y + 1.18, w: 4.2, h: 0.4, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 22 — BACKEND & DONNÉES
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 22, false);
        addSectionTag(s, "BACKEND");
        addPillCounter(s, "22/23");

        s.addText("Infrastructure backend et données", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const backendCards = [
            { icon: icons.node, title: "Moteur API Express.js", body: "API REST performante sur Node.js, utilisant Knex.js pour des requêtes flexibles et sécurisées.", x: 0.3, y: 1.28 },
            { icon: icons.database, title: "Modèle relationnel PostgreSQL", body: "Architecture robuste avec 5 entités principales :", x: 5.15, y: 1.28, isDb: true },
            { icon: icons.bolt, title: "Cache Redis & Temps Réel", body: "Implémentation de Redis pour les sessions/queues, couplé à Socket.IO pour le chat et les notifications.", x: 0.3, y: 3.12 },
            { icon: icons.mapG, title: "Intégrations Mapbox", body: "Intégration de Mapbox pour le suivi précis inter-wilayas, le routage dynamique et les calculs d'ETA.", x: 5.15, y: 3.12 },
        ];

        backendCards.forEach(({ icon, title, body, x, y, isDb }) => {
            s.addShape("roundRect", {
                x, y, w: 4.55, h: 1.65,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
                rectRadius: 0.1
            });
            s.addShape("roundRect", {
                x: x + 0.18, y: y + 0.55, w: 0.6, h: 0.6,
                fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1
            });
            s.addImage({ data: icon, x: x + 0.23, y: y + 0.6, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.2, y: y + 0.12, w: 4.2, h: 0.3, fontSize: 15, bold: true, color: C.deepBlack, fontFace: "Calibri" });

            if (isDb) {
                s.addText(body, { x: x + 0.2, y: y + 1.18, w: 4.2, h: 0.2, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
                const badges = ["User", "Trip", "Booking", "Payment", "Vehicle"];
                badges.forEach((badge, idx) => {
                    const bx = x + 0.2 + idx * 0.72;
                    s.addShape("roundRect", {
                        x: bx, y: y + 1.35, w: 0.65, h: 0.22,
                        fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.08
                    });
                    s.addText(badge, {
                        x: bx, y: y + 1.35, w: 0.65, h: 0.22,
                        fontSize: 7.5, bold: true, color: C.accentGreen,
                        align: "center", valign: "middle", fontFace: "Courier New"
                    });
                });
            } else {
                s.addText(body, { x: x + 0.2, y: y + 1.18, w: 4.2, h: 0.4, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
            }
        });

        // Banner: Testing & Validation
        s.addShape("roundRect", {
            x: 0.3, y: 4.95, w: 9.4, h: 0.58,
            fill: { color: C.darkGreen },
            line: { color: C.midGreen, width: 1.5 },
            rectRadius: 0.08
        });

        // Banner icon
        s.addShape("roundRect", {
            x: 0.48, y: 5.04, w: 0.4, h: 0.4,
            fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 }, rectRadius: 0.08
        });
        s.addImage({ data: icons.shieldG, x: 0.53, y: 5.09, w: 0.3, h: 0.3 });

        // Banner text
        s.addText("Assurance Qualité Rigoureuse", { x: 1.0, y: 5.0, w: 3.5, h: 0.22, fontSize: 12, bold: true, color: C.cream, fontFace: "Calibri" });
        s.addText("Pipeline de tests automatisés validant la logique métier et l'intégration continue (HTTP/Socket.IO).", { x: 1.0, y: 5.22, w: 3.5, h: 0.26, fontSize: 8.5, color: C.accentGreen, fontFace: "Calibri" });

        // Stats columns
        s.addText("107", { x: 4.8, y: 5.0, w: 1.1, h: 0.25, fontSize: 16, bold: true, color: C.cream, fontFace: "Courier New", align: "center" });
        s.addText("TESTS UNITAIRES", { x: 4.8, y: 5.25, w: 1.1, h: 0.2, fontSize: 7.5, bold: true, color: C.accentGreen, fontFace: "Calibri", align: "center" });

        s.addShape("rect", { x: 6.0, y: 5.04, w: 0.01, h: 0.4, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });

        s.addText("100", { x: 6.2, y: 5.0, w: 1.3, h: 0.25, fontSize: 16, bold: true, color: C.cream, fontFace: "Courier New", align: "center" });
        s.addText("TESTS D'INTÉGRATION", { x: 6.2, y: 5.25, w: 1.3, h: 0.2, fontSize: 7.5, bold: true, color: C.accentGreen, fontFace: "Calibri", align: "center" });

        s.addShape("rect", { x: 7.6, y: 5.04, w: 0.01, h: 0.4, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });

        s.addText("100%", { x: 7.8, y: 5.0, w: 1.1, h: 0.25, fontSize: 16, bold: true, color: C.accentGreen, fontFace: "Courier New", align: "center" });
        s.addText("DE RÉUSSITE", { x: 7.8, y: 5.25, w: 1.1, h: 0.2, fontSize: 7.5, bold: true, color: C.cream, fontFace: "Calibri", align: "center" });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 23 — MICROSERVICE IA & SÉCURITÉ
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 23, true);
        addSectionTag(s, "IA KYC", true);
        addPillCounter(s, "23/28", true);

        s.addText("Microservice IA pour l'identité", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.cream, fontFace: "Calibri"
        });

        // Card 1: FastAPI Service
        const x1 = 0.5, y1 = 1.35, w1 = 3.9, h1 = 0.86;
        s.addShape("roundRect", { x: x1, y: y1, w: w1, h: h1, fill: { color: C.deepBlack }, line: { color: C.accentGreen, width: 2 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x1 + 0.12, y: y1 + 0.18, w: 0.5, h: 0.5, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.micro, x: x1 + 0.17, y: y1 + 0.23, w: 0.4, h: 0.4 });
        s.addText("Service FastAPI", { x: x1 + 0.7, y: y1 + 0.1, w: w1 - 0.8, h: 0.22, fontSize: 13, bold: true, color: C.cream, fontFace: "Calibri" });
        s.addText("Moteur de vérification avec détection de vivacité et anti-spoofing.", { x: x1 + 0.7, y: y1 + 0.35, w: w1 - 0.8, h: 0.45, fontSize: 9.5, color: "A3B1A8", fontFace: "Calibri" });

        // Card 2: Decision States
        const x2 = 0.5, y2 = 2.35, w2 = 3.9, h2 = 0.86;
        s.addShape("roundRect", { x: x2, y: y2, w: w2, h: h2, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x2 + 0.12, y: y2 + 0.18, w: 0.5, h: 0.5, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.1 });
        s.addImage({ data: icons.code, x: x2 + 0.17, y: y2 + 0.23, w: 0.4, h: 0.4 });
        s.addText("États de décision", { x: x2 + 0.7, y: y2 + 0.1, w: w2 - 0.8, h: 0.22, fontSize: 13, bold: true, color: C.cream, fontFace: "Calibri" });

        // Badges: ACCEPTÉ, RÉESSAYER, REJETÉ
        const badges = [
            { text: "ACCEPTÉ", bg: "0F3A23", fg: C.accentGreen, x: x2 + 0.7 },
            { text: "RÉESSAYER", bg: "4D380E", fg: C.gold, x: x2 + 1.55 },
            { text: "REJETÉ", bg: "4D1714", fg: "FF6B6B", x: x2 + 2.45 }
        ];
        badges.forEach(b => {
            s.addShape("roundRect", { x: b.x, y: y2 + 0.42, w: 0.75, h: 0.26, fill: { color: b.bg }, line: { color: b.bg, width: 0 }, rectRadius: 0.08 });
            s.addText(b.text, { x: b.x, y: y2 + 0.42, w: 0.75, h: 0.26, fontSize: 8, bold: true, color: b.fg, align: "center", valign: "middle", fontFace: "Courier New" });
        });

        // Card 3: Security & Privacy
        const x3 = 0.5, y3 = 3.35, w3 = 3.9, h3 = 1.35;
        s.addShape("roundRect", { x: x3, y: y3, w: w3, h: h3, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x3 + 0.12, y: y3 + 0.18, w: 0.5, h: 0.5, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.1 });
        s.addImage({ data: icons.shieldG, x: x3 + 0.17, y: y3 + 0.23, w: 0.4, h: 0.4 });
        s.addText("Sécurité & Confidentialité", { x: x3 + 0.7, y: y3 + 0.12, w: w3 - 0.8, h: 0.25, fontSize: 13, bold: true, color: C.cream, fontFace: "Calibri" });

        const secBullets = [
            { icon: icons.compress, text: "Minimisation des données et transferts tokenisés" },
            { icon: icons.key, text: "Chiffrement des plaques d'immatriculation au repos" },
            { icon: icons.shieldVirus, text: "Limitation de débit et prévention du spam OTP" }
        ];
        secBullets.forEach((b, idx) => {
            const by = y3 + 0.45 + idx * 0.28;
            s.addImage({ data: b.icon, x: x3 + 0.7, y: by + 0.03, w: 0.18, h: 0.18 });
            s.addText(b.text, { x: x3 + 0.95, y: by, w: w3 - 1.05, h: 0.22, fontSize: 8.5, color: "A3B1A8", fontFace: "Calibri" });
        });

        // Right side: Diagram Image
        const dx = 4.65, dy = 1.35, dw = 4.85, dh = 3.35;
        s.addShape("roundRect", { x: dx, y: dy, w: dw, h: dh, fill: { color: C.white }, line: { color: C.midGreen, width: 3 }, rectRadius: 0.12 });
        s.addImage({
            path: "https://www.genspark.ai/api/files/s/uHpzXJiL",
            x: dx + 0.05, y: dy + 0.05, w: dw - 0.1, h: dh - 0.1,
            sizing: { type: "contain", w: dw - 0.1, h: dh - 0.1 }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 24 — SECTION DIVIDER (AFFAIRES)
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 24, true);
        addSectionTag(s, "AFFAIRES", true);
        addPillCounter(s, "24/28", true);

        // Quote card shape
        s.addShape("roundRect", { x: 1.5, y: 1.8, w: 7.0, h: 1.8, fill: { color: C.deepBlack }, line: { color: C.midGreen, width: 1 }, rectRadius: 0.15 });

        // Quote Icon
        s.addShape("ellipse", { x: 4.72, y: 1.5, w: 0.56, h: 0.56, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addImage({ data: icons.brief, x: 4.82, y: 1.6, w: 0.36, h: 0.36 });

        // Quote Text
        s.addText("Étude technico-économique et faisabilité du projet", {
            x: 1.6, y: 2.1, w: 6.8, h: 0.4,
            fontSize: 18, bold: true, color: C.cream, align: "center", fontFace: "Calibri"
        });
        s.addText("Mécanisme « Un diplôme, une Startup » — Arrêté interministériel n° 1275", {
            x: 1.6, y: 2.6, w: 6.8, h: 0.3,
            fontSize: 11, bold: true, color: C.accentGreen, align: "center", fontFace: "Courier New", charSpacing: 1
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 25 — MODÈLE DE REVENUS
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 25, false);
        addSectionTag(s, "MONÉTISATION");
        addPillCounter(s, "25/28");

        s.addText("Modèle de revenus et monétisation", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const revCards = [
            { icon: icons.ticket, title: "Flux de revenus principal", body: "Des frais de service par réservation appliqués pour chaque siège.\n\nPourcentage à finaliser après l'homologation officielle SATIM.", x: 0.3, y: 1.28 },
            { icon: icons.star, title: "Fonctionnalités optionnelles", body: "Introduction de fonctionnalités premium (à déterminer) pour la visibilité.\n\nOptions de priorité et badges étendus de vérification de profil.", x: 5.15, y: 1.28, isStar: true },
            { icon: icons.handshake, title: "Partenariats (futur)", body: "Intégration avec les centres de transport et gares routières.\n\nCollaboration avec les initiatives de mobilité des wilayas.", x: 0.3, y: 3.25 },
            { icon: icons.pie, title: "Facteurs de coûts", body: "", x: 5.15, y: 3.25, isCosts: true }
        ];

        revCards.forEach(({ icon, title, body, x, y, isStar, isCosts }) => {
            s.addShape("roundRect", { x, y, w: 4.55, h: 1.8, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
            s.addShape("roundRect", { x: x + 0.18, y: y + 0.55, w: 0.6, h: 0.6, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
            s.addImage({ data: icon, x: x + 0.23, y: y + 0.6, w: 0.5, h: 0.5 });
            s.addText(title, { x: x + 0.2, y: y + 0.12, w: 4.2, h: 0.3, fontSize: 14, bold: true, color: C.deepBlack, fontFace: "Calibri" });

            if (isCosts) {
                const costBullets = [
                    "Infrastructure, cloud et base de données",
                    "Consommation de l'API KYC biométrique",
                    "Support client et assistance",
                    "Conformité réglementaire et légale"
                ];
                costBullets.forEach((bullet, idx) => {
                    const by = y + 0.55 + idx * 0.26;
                    s.addShape("ellipse", { x: x + 0.9, y: by + 0.06, w: 0.06, h: 0.06, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
                    s.addText(bullet, { x: x + 1.05, y: by, w: 3.2, h: 0.24, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
                });
            } else {
                s.addText(body, { x: x + 0.2, y: y + 1.18, w: 4.2, h: 0.52, fontSize: 10, color: C.mutedText, fontFace: "Calibri" });
            }
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 26 — FINANCES (APÈRCU FINANCIER)
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.cream };
        addWatermark(s, 26, false);
        addSectionTag(s, "FINANCES");
        addPillCounter(s, "26/28");

        s.addText("Aperçu financier et vision d'affaires", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 22, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Cost Structure Card
        const x1 = 0.3, y1 = 1.28, w1 = 4.55, h1 = 1.85;
        s.addShape("roundRect", { x: x1, y: y1, w: w1, h: h1, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x1 + 0.18, y: y1 + 0.3, w: 0.48, h: 0.48, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.pie, x: x1 + 0.23, y: y1 + 0.35, w: 0.38, h: 0.38 });
        s.addText("Structure des coûts année 1", { x: x1 + 0.8, y: y1 + 0.12, w: 3.5, h: 0.25, fontSize: 13, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        const costBullets = [
            "Infrastructure cloud et serveurs",
            "Frais KYC IA & passerelles de paiement",
            "Équipe de développement principale",
            "Marketing et acquisition d'utilisateurs"
        ];
        costBullets.forEach((bullet, idx) => {
            const by = y1 + 0.45 + idx * 0.28;
            s.addText("✓", { x: x1 + 0.8, y: by, w: 0.2, h: 0.22, fontSize: 10, bold: true, color: C.accentGreen, fontFace: "Calibri" });
            s.addText(bullet, { x: x1 + 1.05, y: by, w: 3.2, h: 0.22, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Projections Card
        const x2 = 5.15, y2 = 1.28, w2 = 4.55, h2 = 1.85;
        s.addShape("roundRect", { x: x2, y: y2, w: w2, h: h2, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x2 + 0.18, y: y2 + 0.3, w: 0.48, h: 0.48, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.trending, x: x2 + 0.23, y: y2 + 0.35, w: 0.38, h: 0.38 });
        s.addText("Projections et seuil de rentabilité", { x: x2 + 0.8, y: y2 + 0.12, w: 3.5, h: 0.25, fontSize: 13, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        const projectionParas = [
            "Économie unitaire optimisée établissant une voie claire vers une rentabilité précoce.",
            "La tarification dynamique atténue les risques liés aux coûts opérationnels.",
            "Détails disponibles en annexe (à détailler)."
        ];
        projectionParas.forEach((p, idx) => {
            const py = y2 + 0.45 + idx * 0.38;
            s.addText(p, { x: x2 + 0.8, y: py, w: 3.5, h: 0.35, fontSize: 9, color: C.mutedText, fontFace: "Calibri" });
        });

        // Label Startup Card
        const x3 = 0.3, y3 = 3.35, w3 = 4.55, h3 = 1.85;
        s.addShape("roundRect", { x: x3, y: y3, w: w3, h: h3, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x3 + 0.18, y: y3 + 0.3, w: 0.48, h: 0.48, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.award, x: x3 + 0.23, y: y3 + 0.35, w: 0.38, h: 0.38 });
        s.addText("Label startup (Arrêté 1275)", { x: x3 + 0.8, y: y3 + 0.12, w: 3.5, h: 0.25, fontSize: 13, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        const labelParas = [
            "Ancré dans le mécanisme « Un diplôme, une Startup ».",
            "Plan d'incubation structuré guidant la transition du projet à l'entreprise.",
            "Ciblant l'accréditation officielle du label Startup."
        ];
        labelParas.forEach((p, idx) => {
            const py = y3 + 0.45 + idx * 0.38;
            s.addText(p, { x: x3 + 0.8, y: py, w: 3.5, h: 0.35, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });

        // Target Organisation Card
        const x4 = 5.15, y4 = 3.35, w4 = 4.55, h4 = 1.85;
        s.addShape("roundRect", { x: x4, y: y4, w: w4, h: h4, fill: { color: C.white }, line: { color: C.cardBorder, width: 0.5 }, shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 }, rectRadius: 0.1 });
        s.addShape("roundRect", { x: x4 + 0.18, y: y4 + 0.3, w: 0.48, h: 0.48, fill: { color: C.cream }, line: { color: C.cream, width: 0 }, rectRadius: 0.1 });
        s.addImage({ data: icons.sitemap, x: x4 + 0.23, y: y4 + 0.35, w: 0.38, h: 0.38 });
        s.addText("Organisation cible à 18 mois", { x: x4 + 0.8, y: y4 + 0.12, w: 3.5, h: 0.25, fontSize: 13, bold: true, color: C.deepBlack, fontFace: "Calibri" });
        const targetBullets = [
            "Équipe d'ingénierie évolutive pour le déploiement continu.",
            "Équipe de support et de conformité dédiée 24/7.",
            "Responsables des opérations régionales à l'échelle nationale."
        ];
        targetBullets.forEach((bullet, idx) => {
            const by = y4 + 0.45 + idx * 0.38;
            s.addText("→", { x: x4 + 0.8, y: by, w: 0.2, h: 0.24, fontSize: 10, bold: true, color: C.accentGreen, fontFace: "Calibri" });
            s.addText(bullet, { x: x4 + 1.05, y: by, w: 3.2, h: 0.34, fontSize: 9.5, color: C.mutedText, fontFace: "Calibri" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 27 — ROADMAP
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 27, true);
        addSectionTag(s, "FEUILLE DE ROUTE", true);
        addPillCounter(s, "27/28", true);

        s.addText("Calendrier stratégique 18 mois", {
            x: 0.5, y: 0.75, w: 9, h: 0.45,
            fontSize: 24, bold: true, color: C.cream, fontFace: "Calibri"
        });

        // Timeline track
        s.addShape("rect", { x: 0.8, y: 2.12, w: 8.4, h: 0.04, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });

        const phases = [
            { icon: icons.flag, period: "Q1 — Q2", title: "Lancement et pilote", bullets: ["Pilote dans les corridors clés", "Homologation SATIM"], x: 0.8, active: true },
            { icon: icons.route, period: "Q3 — Q4", title: "Expansion", bullets: ["Passage à 30+ wilayas", "Incitations pour conducteurs"], x: 3.75, active: false },
            { icon: icons.mapIcon, period: "Q5 — Q6", title: "Couverture nationale", bullets: ["Couverture nationale complète", "Tarification avancée, partenariats"], x: 6.7, active: false }
        ];

        phases.forEach(({ icon, period, title, bullets, x, active }) => {
            const nx = x + 1.25 - 0.22;
            s.addShape("ellipse", {
                x: nx, y: 1.9, w: 0.44, h: 0.44,
                fill: { color: active ? C.accentGreen : C.deepBlack },
                line: { color: active ? C.accentGreen : C.midGreen, width: 2 }
            });
            if (icon) s.addImage({ data: icon, x: nx + 0.07, y: 1.97, w: 0.3, h: 0.3 });

            // Card
            s.addShape("roundRect", {
                x, y: 2.6, w: 2.5, h: 1.6,
                fill: { color: C.deepBlack },
                line: { color: C.midGreen, width: 1 },
                shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.4 },
                rectRadius: 0.1
            });
            s.addText(period, { x: x + 0.15, y: 2.7, w: 2.2, h: 0.22, fontSize: 11, bold: true, color: C.accentGreen, fontFace: "Courier New" });
            s.addText(title, { x: x + 0.15, y: 2.92, w: 2.2, h: 0.3, fontSize: 14, bold: true, color: C.cream, fontFace: "Calibri" });
            bullets.forEach((b, i) => {
                s.addShape("ellipse", { x: x + 0.18, y: 3.32 + i * 0.24, w: 0.06, h: 0.06, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
                s.addText(b, { x: x + 0.32, y: 3.26 + i * 0.24, w: 2.05, h: 0.32, fontSize: 9.5, color: "A3B1A8", fontFace: "Calibri" });
            });
        });

        // KPI Banner
        s.addShape("roundRect", {
            x: 0.8, y: 4.4, w: 8.4, h: 0.8,
            fill: { color: C.accentGreen, transparency: 90 },
            line: { color: C.midGreen, width: 1, dashType: "dash" },
            rectRadius: 0.08
        });

        const kpis = [
            { icon: icons.users, cx: 1.0, title: "KPI TARGET", label: "MAU" },
            { icon: icons.trend, cx: 3.1, title: "KPI TARGET", label: "Conversion" },
            { icon: icons.shieldG, cx: 5.2, title: "KPI TARGET", label: "Score de sécurité" },
            { icon: icons.clock, cx: 7.3, title: "KPI TARGET", label: "Taux de ponctualité" }
        ];

        kpis.forEach(({ icon, cx, title, label }) => {
            s.addImage({ data: icon, x: cx, y: 4.62, w: 0.32, h: 0.32 });
            s.addText(title, { x: cx + 0.4, y: 4.48, w: 1.5, h: 0.18, fontSize: 7, bold: true, color: C.accentGreen, fontFace: "Courier New" });
            s.addText(label, { x: cx + 0.4, y: 4.66, w: 1.5, h: 0.32, fontSize: 9.5, bold: true, color: C.cream, fontFace: "Calibri" });
        });
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SLIDE 28 — CLÔTURE
    // ══════════════════════════════════════════════════════════════════════════
    {
        const s = pres.addSlide();
        s.background = { color: C.darkGreen };
        addWatermark(s, 28, true);
        addSectionTag(s, "CLÔTURE", true);
        addPillCounter(s, "28/28", true);

        // RohWinBghit Logo Main
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo.png"), x: 3.28, y: 1.3, w: 1.1, h: 1.1 });

        // RohWinBghit Logo Bus
        s.addImage({ path: path.join(__dirname, "Screen_mobile", "logo_bus.png"), x: 4.69, y: 1.68, w: 0.55, h: 0.55 });

        // RohWinBghit Wordmark
        s.addText("RohWinBghit", {
            x: 5.31, y: 1.68, w: 2.03, h: 0.55,
            fontSize: 32, bold: true, color: C.cream,
            fontFace: "Calibri", valign: "middle"
        });

        // Title "Merci"
        s.addText("Merci", {
            x: 0, y: 2.65, w: 10.0, h: 0.7,
            fontSize: 64, bold: true, color: C.cream, align: "center", fontFace: "Calibri"
        });
        s.addText("Merci pour votre temps et votre attention durant cette présentation.", {
            x: 0, y: 3.35, w: 10.0, h: 0.35,
            fontSize: 20, bold: false, color: C.accentGreen, align: "center", fontFace: "Calibri"
        });

        // Authors Card (Left Side)
        s.addText("Réalisé par :", { x: 1.4, y: 3.82, w: 3.28, h: 0.22, fontSize: 11, bold: true, color: C.accentGreen, align: "right", fontFace: "Calibri" });
        s.addText("AHMED BACHA Djamel Eddine", { x: 1.4, y: 4.08, w: 3.28, h: 0.26, fontSize: 13, bold: false, color: C.cream, align: "right", fontFace: "Calibri" });
        s.addText("BELHORMA Sidi Mohammed Reduane", { x: 1.4, y: 4.36, w: 3.28, h: 0.26, fontSize: 13, color: C.cream, align: "right", fontFace: "Calibri" });

        // Divider
        s.addShape("rect", { x: 4.99, y: 3.9, w: 0.01, h: 0.55, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });

        // Supervisor Card (Right Side)
        s.addText("Encadré par :", { x: 5.31, y: 3.82, w: 3.28, h: 0.22, fontSize: 11, bold: true, color: C.accentGreen, align: "left", fontFace: "Calibri" });
        s.addText("Mme BENLEDGHEM Rafika", { x: 5.31, y: 4.08, w: 3.28, h: 0.26, fontSize: 13, color: C.cream, align: "left", fontFace: "Calibri" });
        s.addText("Université Abou Bekr Belkaïd – Tlemcen\nDépartement d'Informatique", { x: 5.31, y: 4.36, w: 3.28, h: 0.36, fontSize: 10.5, color: C.cream, transparency: 20, align: "left", fontFace: "Calibri" });

        // CTA Banner
        s.addShape("rect", { x: 0, y: 4.84, w: 10.0, h: 0.78, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addText("Demandez le mémoire complet ou une démo en direct", {
            x: 0, y: 4.84, w: 10.0, h: 0.78,
            fontSize: 20, bold: true, color: C.deepBlack,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ─── WRITE FILE ──────────────────────────────────────────────────────────
    const outPath = path.join(__dirname, "RohWinBghit_Presentation.pptx");
    await pres.writeFile({ fileName: outPath });
    console.log("✅  Fichier généré :", outPath);
}

// Helper used in slide 01 (section tag on dark background)
function slide01_sectionTag(s) {
    s.addShape("rect", { x: 0.5, y: 0.45, w: 0.32, h: 0.035, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
    s.addText("DÉBUT", { x: 0.88, y: 0.38, w: 3, h: 0.2, fontSize: 11, bold: true, color: C.accentGreen, fontFace: "Courier New", charSpacing: 2 });
}

buildPresentation().catch(err => { console.error("❌ Erreur :", err); process.exit(1); });