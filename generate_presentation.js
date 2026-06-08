/**
 * RohWinBghit — Présentation PPTX (Structure Optimisée 31 Slides)
 * 24 slides de contenu + 7 slides de séparation de section
 * Génère la soutenance Master 2 Génie Logiciel — Arrêté 1275
 * Université Abou Bekr Belkaïd — Tlemcen
 * Usage: node "generate_presentation.js"
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
    FaHourglassHalf, FaCircleCheck, FaFlask, FaGavel, FaFileShield
} = require("react-icons/fa6");
const { FaNodeJs, FaReact, FaDocker, FaPython, FaGlobe } = require("react-icons/fa");
const { TbSteeringWheel } = require("react-icons/tb");
const {
    SiPostgresql, SiRedis, SiExpo, SiExpress, SiNginx
} = require("react-icons/si");

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
    red: "D44017",
    orange: "E8A838",
    blue: "2196F3",
};

// ─── ICON HELPER ─────────────────────────────────────────────────────────────
async function iconPng(IconComp, color = "#4CC080", size = 256) {
    try {
        const svg = ReactDOMServer.renderToStaticMarkup(
            React.createElement(IconComp, { color, size: String(size) })
        );
        const buf = await sharp(Buffer.from(svg)).png().toBuffer();
        return "image/png;base64," + buf.toString("base64");
    } catch {
        return null;
    }
}

// ─── SHARED LAYOUT HELPERS ───────────────────────────────────────────────────
const SW = 10, SH = 5.625;
const TOTAL_SLIDES = 31;

function addSectionTag(slide, label, dark = false) {
    slide.addShape("rect", {
        x: 0, y: 0, w: 0.07, h: SH,
        fill: { color: C.midGreen },
        line: { color: C.midGreen, width: 0 }
    });
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

/** Adds a small colored pill badge */
function addBadge(slide, text, x, y, bgColor, textColor = C.white, w = 0.9) {
    slide.addShape("roundRect", {
        x, y, w, h: 0.22,
        fill: { color: bgColor },
        line: { color: bgColor, width: 0 },
        rectRadius: 0.11
    });
    slide.addText(text, {
        x, y, w, h: 0.22,
        fontSize: 8, bold: true, color: textColor,
        fontFace: "Calibri", align: "center", valign: "middle"
    });
}

/** Adds a table row-like block for data presentation */
function addTableRow(slide, cols, y, h = 0.38, headerRow = false) {
    const totalW = 8.5;
    const startX = 0.75;
    const colW = totalW / cols.length;
    cols.forEach((text, i) => {
        slide.addShape("rect", {
            x: startX + i * colW, y, w: colW, h,
            fill: { color: headerRow ? C.darkGreen : C.white },
            line: { color: C.cardBorder, width: 0.5 }
        });
        slide.addText(text, {
            x: startX + i * colW, y, w: colW, h,
            fontSize: headerRow ? 10 : 9.5,
            bold: headerRow,
            color: headerRow ? C.cream : C.deepBlack,
            fontFace: "Calibri",
            align: "center", valign: "middle"
        });
    });
}

// ─── IMAGE PATH HELPER ──────────────────────────────────────────────────────
const IMG = {
    logo: path.join(__dirname, "Screen_mobile", "logo_bus.png"),
    univLogo: path.join(__dirname, "Screen_mobile", "univ_logo.png"),
    useCaseDiag: path.join(__dirname, "figures", "fig_use_case.png"),
    classDiag: path.join(__dirname, "figures", "fig_class_diagram.png"),
    deployDiag: path.join(__dirname, "figures", "fig_deployment2.png"),
    archDiag: path.join(__dirname, "figures", "fig_architecture2.png"),
    kycMobile: path.join(__dirname, "figures", "fig_kyc_mobile.png"),
    kycReview: path.join(__dirname, "figures", "fig_kyc_review.png"),
    // Mobile screens
    landing: path.join(__dirname, "figures", "v2_landing_page.jpg"),
    splash: path.join(__dirname, "figures", "v2_welcome_splash.jpg"),
    roleSelect: path.join(__dirname, "figures", "v2_role_selection.jpg"),
    publishTrip: path.join(__dirname, "figures", "v2_publish_trip.jpg"),
    searchResults: path.join(__dirname, "figures", "v2_search_results.jpg"),
    ticketPending: path.join(__dirname, "figures", "v2_ticket_pending.jpg"),
    inscription: path.join(__dirname, "Screen_mobile", "inscription.jpg"),
    passagerReserve: path.join(__dirname, "Screen_mobile", "Passager réserve.jpg"),
    trajetPublie: path.join(__dirname, "Screen_mobile", "Trajet publié.jpg"),
    qrScan: path.join(__dirname, "Screen_mobile", "QR scanné au départ.jpg"),
    tripDetails: path.join(__dirname, "Screen_mobile", "p_09_trip_details.png"),
    // Admin
    adminDash: path.join(__dirname, "assets", "admin_dashboard.png"),
    adminKyc: path.join(__dirname, "assets", "admin_kyc.png"),
    adminUsers: path.join(__dirname, "assets", "admin_users.png"),
    // Driver
    driverDash: path.join(__dirname, "assets", "driver_dashboard.png"),
};

// ─── BUILD ──────────────────────────────────────────────────────────────────
async function buildPresentation() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane";
    pres.title = "RohWinBghit — Présentation de Soutenance Master 2 GL";

    // Pré-rendu des icônes
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
        circleCheck: await iconPng(FaCircleCheck, "#4CC080"),
        flask: await iconPng(FaFlask, "#4CC080"),
        gavel: await iconPng(FaGavel, "#166E3A"),
        fileShield: await iconPng(FaFileShield, "#166E3A"),
        docker: await iconPng(FaDocker, "#166E3A"),
        python: await iconPng(FaPython, "#166E3A"),
        globe: await iconPng(FaGlobe, "#166E3A"),
    };

    let currentSlideNum = 0;

    function createSlide(sectionLabel, dark = false) {
        currentSlideNum++;
        const s = pres.addSlide();
        s.background = { color: dark ? C.darkGreen : C.cream };
        addWatermark(s, currentSlideNum, dark);
        addSectionTag(s, sectionLabel, dark);
        addPillCounter(s, `${String(currentSlideNum).padStart(2, "0")}/${TOTAL_SLIDES}`, dark);
        return s;
    }

    function createTransitionSlide(sectionNum, sectionTitle, questionGuide) {
        const s = createSlide(sectionTitle.toUpperCase(), true);

        // Grand cercle de section
        s.addShape("ellipse", {
            x: 1.8, y: 1.6, w: 2.0, h: 2.0,
            fill: { color: C.accentGreen },
            line: { color: C.accentGreen, width: 0 },
            shadow: { type: "outer", blur: 12, offset: 3, angle: 180, color: "000000", opacity: 0.25 }
        });
        s.addText(String(sectionNum), {
            x: 1.8, y: 1.6, w: 2.0, h: 2.0,
            fontSize: 72, bold: true, color: C.deepBlack,
            fontFace: "Arial Black", align: "center", valign: "middle"
        });

        // Titre de section
        s.addText(sectionTitle, {
            x: 4.2, y: 1.8, w: 5.3, h: 1.0,
            fontSize: 32, bold: true, color: C.cream,
            fontFace: "Calibri", align: "left", valign: "middle"
        });

        // Ligne d'accentuation dorée
        s.addShape("rect", {
            x: 4.2, y: 3.0, w: 3.0, h: 0.04,
            fill: { color: C.gold }, line: { color: C.gold, width: 0 }
        });

        // Question-guide
        s.addText(questionGuide, {
            x: 4.2, y: 3.25, w: 5.3, h: 0.5,
            fontSize: 16, italic: true, color: C.gold,
            fontFace: "Calibri", align: "left", valign: "middle"
        });

        return s;
    }

    // =========================================================================
    // SLIDE 01 — Couverture académique
    // =========================================================================
    {
        currentSlideNum = 0;
        const s = createSlide("DÉBUT", true);

        s.addImage({ path: IMG.univLogo, x: 4.61, y: 0.74, w: 0.78, h: 0.78 });
        s.addImage({ path: IMG.logo, x: 2.19, y: 1.83, w: 0.78, h: 0.78 });

        s.addText("RohWinBghit — روح وين بغيت", {
            x: 3.13, y: 1.75, w: 5.5, h: 0.94,
            fontSize: 40, bold: true, color: C.cream,
            fontFace: "Calibri", valign: "middle"
        });

        s.addText("Plateforme mobile multiplateforme intelligente de covoiturage\ninter-wilayas sécurisée adaptée au contexte algérien", {
            x: 0.5, y: 2.8, w: 9.0, h: 1.0,
            fontSize: 22, bold: true, color: C.cream,
            align: "center", fontFace: "Calibri"
        });

        s.addText("Mémoire de Master en Génie Logiciel — Arrêté 1275\nUniversité Abou Bekr Belkaïd — Tlemcen — Année 2024/2025", {
            x: 0.5, y: 3.8, w: 9.0, h: 0.45,
            fontSize: 13, color: C.accentGreen, align: "center", fontFace: "Calibri"
        });

        s.addText("Présenté par : AHMED BACHA Djamel Eddine & BELHORMA Sidi Mohammed Reduane\nEncadré par : Mme BENLEDGHEM Rafika", {
            x: 0.5, y: 4.35, w: 9.0, h: 0.5,
            fontSize: 12.5, color: C.cream, align: "center", fontFace: "Calibri"
        });

        s.addShape("rect", { x: 0, y: 5.12, w: SW, h: 0.5, fill: { color: C.midGreen }, line: { color: C.midGreen, width: 0 } });
        s.addText("SIMPLE   •   SÛR   •   ABORDABLE", {
            x: 0, y: 5.12, w: SW, h: 0.5,
            fontSize: 14, bold: true, color: C.cream,
            align: "center", valign: "middle", charSpacing: 3, fontFace: "Calibri"
        });
    }

    // =========================================================================
    // SLIDE 02 — Plan de la Présentation (4+3 Grid with Premium Styling)
    // =========================================================================
    {
        const s = createSlide("PLAN DE LA PRÉSENTATION");

        s.addText("Plan de la Présentation", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const plans = [
            ["01", "Introduction & Contexte", "Mobilité inter-wilayas • Chiffres clés", icons.route],
            ["02", "Problématique & Objectifs", "Confiance • KYC • Hypothèses", icons.shieldG],
            ["03", "Étude de l'Existant", "Concurrence • Gaps • Loi 25-11", icons.magnifyingGlass],
            ["04", "Analyse & Conception", "Use Case • UML • Classes", icons.sitemap],
            ["05", "Architecture & Technologies", "Stack • KYC • Microservices", icons.network],
            ["06", "Réalisation & Démonstration", "Prototype • Flux passager", icons.mobile],
            ["07", "Résultats & Conclusion", "Tests Jest • SUS 71.6 • Perspectives", icons.trend],
        ];

        // Layout: 4 cards on top row, 3 cards on bottom row (centered)
        const gridStartY = 1.55;
        const colW = 2.1, rowH = 1.6, gap = 0.15;

        plans.forEach(([num, title, sub, icon], i) => {
            let x, y;
            if (i < 4) {
                // Top row: 4 cards
                x = 0.45 + i * (colW + gap);
                y = gridStartY;
            } else {
                // Bottom row: 3 cards, centered
                const offset = (4 * (colW + gap) - 3 * (colW + gap)) / 2;
                x = 0.45 + offset + (i - 4) * (colW + gap);
                y = gridStartY + rowH + gap;
            }

            const isActive = (num === "01");

            s.addShape("roundRect", {
                x, y, w: colW, h: rowH,
                fill: { color: C.white },
                line: { color: isActive ? C.accentGreen : C.cardBorder, width: isActive ? 1.5 : 0.5 },
                shadow: { type: "outer", blur: 6, offset: 1, angle: 135, color: "0A1F14", opacity: 0.04 },
                rectRadius: 0.08
            });

            if (isActive) {
                // Bottom green accent line
                s.addShape("rect", {
                    x: x + 0.05, y: y + rowH - 0.04, w: colW - 0.1, h: 0.04,
                    fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 }
                });
            }

            // Number badge (dark green pill)
            s.addShape("roundRect", {
                x: x + (colW - 0.45) / 2, y: y + 0.08, w: 0.45, h: 0.22,
                fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.11
            });
            s.addText(num, {
                x: x + (colW - 0.45) / 2, y: y + 0.08, w: 0.45, h: 0.22,
                fontSize: 9.5, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle"
            });

            // Icon Container (light green square)
            s.addShape("roundRect", {
                x: x + (colW - 0.38) / 2, y: y + 0.36, w: 0.38, h: 0.38,
                fill: { color: "F0F7F2" }, line: { color: "F0F7F2", width: 0 }, rectRadius: 0.06
            });
            if (icon) {
                s.addImage({
                    data: icon,
                    x: x + (colW - 0.24) / 2, y: y + 0.43,
                    w: 0.24, h: 0.24
                });
            }

            // Title (Bold, 12pt, Calibri)
            s.addText(title, {
                x: x + 0.05, y: y + 0.78, w: colW - 0.1, h: 0.42,
                fontSize: 12, bold: true, color: C.deepBlack, fontFace: "Calibri", align: "center", valign: "middle"
            });

            // Subtitle / Description (8.5pt, gray-green)
            s.addText(sub, {
                x: x + 0.05, y: y + 1.25, w: colW - 0.1, h: 0.3,
                fontSize: 8.5, color: C.mutedText, fontFace: "Calibri", align: "center", valign: "top"
            });
        });
    }

    // ─── SÉPARATION : Introduction & Problématique ────────────────────────────
    createTransitionSlide(1, "Introduction & Problématique", "Où ? Quand ? Dans quel contexte ?");

    // =========================================================================
    // SLIDE — La Mobilité Inter-Wilayas en Algérie
    // =========================================================================
    {
        const s = createSlide("INTRODUCTION");

        s.addText("La Mobilité Inter-Wilayas en Algérie", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addText("Un réseau routier de 112 696 km reliant 69 wilayas,\nmais aucune plateforme numérique dédiée au covoiturage longue distance.", {
            x: 0.5, y: 1.4, w: 9.0, h: 0.6,
            fontSize: 14, italic: true, color: C.mutedText, fontFace: "Calibri", align: "center"
        });

        // Stat cards row
        const stats = [
            [icons.users, "6.3M", "Étudiants se déplaçant\nentre wilayas chaque année"],
            [icons.car, "80%", "Déplacements inter-wilayas\nvia transports informels"],
            [icons.shieldG, "0", "Plateformes KYC dédiées\naux trajets longue distance"],
        ];
        stats.forEach(([icon, num, desc], i) => {
            const x = 0.75 + i * 3.1;
            s.addShape("roundRect", {
                x, y: 2.25, w: 2.8, h: 1.6,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
                rectRadius: 0.1
            });
            if (icon) s.addImage({ data: icon, x: x + 1.1, y: 2.4, w: 0.5, h: 0.5 });
            s.addText(num, {
                x, y: 2.95, w: 2.8, h: 0.45,
                fontSize: 32, bold: true, color: C.darkGreen,
                fontFace: "Calibri", align: "center", valign: "middle"
            });
            s.addText(desc, {
                x, y: 3.42, w: 2.8, h: 0.38,
                fontSize: 9.5, color: C.mutedText,
                fontFace: "Calibri", align: "center", valign: "top"
            });
        });

        // Problem framing
        s.addShape("roundRect", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.6,
            fill: { color: C.darkGreen },
            line: { color: C.midGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("⚠  Risques : Aucune vérification d'identité • Paiement en espèces uniquement • Groupes Facebook non sécurisés", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.6,
            fontSize: 11, bold: true, color: C.gold,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 04 — Le Problème de Confiance & Sécurité
    // =========================================================================
    {
        const s = createSlide("PROBLÉMATIQUE");

        s.addText("Le Problème de Confiance & Sécurité", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Central research question box
        s.addShape("roundRect", {
            x: 0.75, y: 1.5, w: 8.5, h: 1.2,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 1 },
            rectRadius: 0.1
        });
        s.addText("QUESTION CENTRALE DE RECHERCHE", {
            x: 0.75, y: 1.55, w: 8.5, h: 0.3,
            fontSize: 10, bold: true, color: C.accentGreen,
            fontFace: "Courier New", align: "center", charSpacing: 2
        });
        s.addText("Comment concevoir une plateforme de covoiturage inter-wilayas\nintégrant un pipeline biométrique KYC pour garantir\nla confiance et la sécurité des utilisateurs ?", {
            x: 1.0, y: 1.85, w: 8.0, h: 0.75,
            fontSize: 16, bold: true, color: C.cream,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Sub-problems as cards
        const subProbs = [
            [icons.idCard, "Identité", "Vérification de l'identité réelle\ndes conducteurs et passagers"],
            [icons.card, "Paiement", "Intégration du paiement local\n(SATIM, Edahabia, CIB)"],
            [icons.shieldG, "Résilience", "Architecture API tolérante\naux pannes (99%+ uptime)"],
            [icons.mapG, "Couverture", "Desservir les 69 wilayas\navec une UX fluide"],
        ];
        subProbs.forEach(([icon, title, body], i) => {
            const x = 0.55 + i * 2.3;
            addIconCard(s, icon, title, body, x, 3.0, 2.15, 1.5);
        });
    }

    // =========================================================================
    // SLIDE 05 — Trois Hypothèses de Travail
    // =========================================================================
    {
        const s = createSlide("HYPOTHÈSES");

        s.addText("Trois Hypothèses de Travail", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const hypotheses = [
            {
                id: "H1", title: "KYC Biométrique",
                desc: "Un pipeline d'IA combinant OCR, ArcFace et détection de vivacité peut vérifier l'identité des utilisateurs avec un seuil de confiance ≥ 0.65 (conducteurs) et ≥ 0.60 (passagers).",
                icon: icons.brain, badge: "INNOVATION", badgeColor: C.midGreen
            },
            {
                id: "H2", title: "Usabilité Mobile",
                desc: "L'application mobile obtient un score SUS ≥ 68/100, validant une expérience utilisateur de qualité « Good » selon les standards ISO.",
                icon: icons.mobile, badge: "PROTOTYPAGE", badgeColor: C.blue
            },
            {
                id: "H3", title: "Résilience API",
                desc: "L'architecture microservices avec circuit breaker, retry patterns et health checks maintient un taux de disponibilité ≥ 99%.",
                icon: icons.network, badge: "SOLVABILITÉ", badgeColor: C.gold
            }
        ];

        hypotheses.forEach((h, i) => {
            const y = 1.55 + i * 1.25;
            // Card
            s.addShape("roundRect", {
                x: 0.75, y, w: 8.5, h: 1.1,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                shadow: { type: "outer", blur: 6, offset: 1, angle: 135, color: "0A1F14", opacity: 0.04 },
                rectRadius: 0.1
            });
            // H-number badge
            s.addShape("roundRect", {
                x: 0.95, y: y + 0.15, w: 0.55, h: 0.55,
                fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.08
            });
            s.addText(h.id, {
                x: 0.95, y: y + 0.15, w: 0.55, h: 0.55,
                fontSize: 18, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle"
            });
            // Icon
            if (h.icon) s.addImage({ data: h.icon, x: 1.65, y: y + 0.2, w: 0.45, h: 0.45 });
            // Title + badge
            s.addText(h.title, {
                x: 2.25, y: y + 0.12, w: 3.0, h: 0.3,
                fontSize: 14, bold: true, color: C.deepBlack, fontFace: "Calibri"
            });
            addBadge(s, h.badge, 5.3, y + 0.15, h.badgeColor);
            // Description
            s.addText(h.desc, {
                x: 2.25, y: y + 0.45, w: 6.7, h: 0.55,
                fontSize: 10, color: C.mutedText, fontFace: "Calibri"
            });
        });
    }

    // ─── SÉPARATION : Contexte & Cadre Légal ─────────────────────────────────
    createTransitionSlide(2, "Contexte & Cadre Légal", "Quel marché ? Quelle réglementation ?");

    // =========================================================================
    // SLIDE — Limites de la Concurrence & Gaps
    // =========================================================================
    {
        const s = createSlide("CONTEXTE");

        s.addText("Analyse Concurrentielle : Les Gaps du Marché", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Comparison table
        const headers = ["Critère", "BlaBlaCar", "Yassir", "Nroho", "inDrive", "RohWinBghit"];
        const rows = [
            ["KYC Biométrique", "✗", "✗", "✗", "✗", "✓  ArcFace"],
            ["Paiement Local", "✗  (CB EU)", "Partiel", "✗", "Espèces", "✓  SATIM/CIB"],
            ["Inter-Wilayas", "✗", "✗  (Urbain)", "✓", "✗", "✓  69 wilayas"],
            ["Évaluation Mutuelle", "✓", "✓", "✗", "✓", "✓  +QR Code"],
            ["Open Source", "✗", "✗", "✗", "✗", "✓"],
        ];

        const colW = [1.8, 1.3, 1.1, 1.1, 1.1, 2.1];
        const startX = 0.6;
        const startY = 1.5;
        const rowH = 0.5;

        // Header
        let cx = startX;
        headers.forEach((h, i) => {
            s.addShape("rect", {
                x: cx, y: startY, w: colW[i], h: rowH,
                fill: { color: C.darkGreen },
                line: { color: C.midGreen, width: 0.5 }
            });
            s.addText(h, {
                x: cx, y: startY, w: colW[i], h: rowH,
                fontSize: 10, bold: true, color: C.cream,
                fontFace: "Calibri", align: "center", valign: "middle"
            });
            cx += colW[i];
        });

        // Data rows
        rows.forEach((row, ri) => {
            const y = startY + (ri + 1) * rowH;
            cx = startX;
            row.forEach((cell, ci) => {
                const isLast = ci === row.length - 1;
                const hasCheck = cell.startsWith("✓");
                const hasCross = cell.startsWith("✗");
                s.addShape("rect", {
                    x: cx, y, w: colW[ci], h: rowH,
                    fill: { color: isLast ? C.lightGreenBg : C.white },
                    line: { color: C.cardBorder, width: 0.5 }
                });
                s.addText(cell, {
                    x: cx, y, w: colW[ci], h: rowH,
                    fontSize: 9.5, bold: isLast,
                    color: hasCheck ? C.midGreen : hasCross ? C.red : C.deepBlack,
                    fontFace: "Calibri", align: "center", valign: "middle"
                });
                cx += colW[ci];
            });
        });

        // Key insight
        s.addShape("roundRect", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("→  RohWinBghit est la seule plateforme combinant KYC biométrique + paiement local + couverture 69 wilayas", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fontSize: 12, bold: true, color: C.accentGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 07 — Cadre Réglementaire : Loi 25-11
    // =========================================================================
    {
        const s = createSlide("CONTEXTE");

        s.addText("Cadre Réglementaire : Loi 25-11", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addText("Protection des données personnelles dans le contexte algérien", {
            x: 0.5, y: 1.35, w: 9.0, h: 0.35,
            fontSize: 13, italic: true, color: C.mutedText, fontFace: "Calibri", align: "center"
        });

        const legalCards = [
            [icons.fileShield, "Consentement Explicite", "Collecte de données uniquement\naprès accord utilisateur (Art. 7)"],
            [icons.key, "Chiffrement AES-256", "Données biométriques stockées\nchiffrées côté serveur (Art. 42)"],
            [icons.gavel, "Droit à l'Effacement", "Suppression des données sur\ndemande de l'utilisateur (Art. 34)"],
            [icons.shieldVirus, "Traitement Local", "Aucun transfert transfrontalier\nsans cadre légal (Art. 44)"],
        ];
        legalCards.forEach(([icon, title, body], i) => {
            const col = i % 2, row = Math.floor(i / 2);
            const x = 0.75 + col * 4.3;
            const y = 1.95 + row * 1.15;
            addIconCard(s, icon, title, body, x, y, 4.0, 0.95);
        });

        // Compliance badge
        s.addShape("roundRect", {
            x: 0.75, y: 4.4, w: 8.5, h: 0.55,
            fill: { color: C.lightGreenBg },
            line: { color: C.midGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("✓  RohWinBghit est conçu en conformité totale avec la Loi 25-11 et les principes GDPR", {
            x: 0.75, y: 4.4, w: 8.5, h: 0.55,
            fontSize: 12, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // ─── SÉPARATION : Méthodologie & Architecture ────────────────────────────
    createTransitionSlide(3, "Méthodologie & Architecture", "Comment ? Avec quelles technologies ?");

    // =========================================================================
    // SLIDE — Gestion Agile : 5 Sprints en 10 Semaines
    // =========================================================================
    {
        const s = createSlide("MÉTHODOLOGIE");

        s.addText("Gestion Agile : 5 Sprints en 10 Semaines", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const sprints = [
            { id: "S1", label: "Auth & Profils", weeks: "Sem 1-2", pts: "21 SP", color: C.midGreen },
            { id: "S2", label: "Trajets & Réservation", weeks: "Sem 3-4", pts: "34 SP", color: C.midGreen },
            { id: "S3", label: "Pipeline KYC", weeks: "Sem 5-6", pts: "28 SP", color: C.darkGreen },
            { id: "S4", label: "Paiement & QR", weeks: "Sem 7-8", pts: "25 SP", color: C.midGreen },
            { id: "S5", label: "Admin & Tests", weeks: "Sem 9-10", pts: "22 SP", color: C.midGreen },
        ];

        // Timeline bar
        const barY = 1.65, barH = 0.08;
        s.addShape("rect", {
            x: 0.75, y: barY, w: 8.5, h: barH,
            fill: { color: C.accentGreen },
            line: { color: C.accentGreen, width: 0 }
        });

        sprints.forEach((sp, i) => {
            const segW = 8.5 / 5;
            const x = 0.75 + i * segW;
            const nodeY = barY - 0.12;
            // Node circle
            s.addShape("ellipse", {
                x: x + segW / 2 - 0.15, y: nodeY, w: 0.3, h: 0.3,
                fill: { color: sp.color },
                line: { color: C.accentGreen, width: 1 }
            });
            s.addText(sp.id, {
                x: x + segW / 2 - 0.15, y: nodeY, w: 0.3, h: 0.3,
                fontSize: 9, bold: true, color: C.cream,
                fontFace: "Courier New", align: "center", valign: "middle"
            });
            // Sprint card below
            const cardY = 2.1;
            s.addShape("roundRect", {
                x: x + 0.05, y: cardY, w: segW - 0.1, h: 1.8,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                rectRadius: 0.08
            });
            s.addText(sp.label, {
                x: x + 0.05, y: cardY + 0.1, w: segW - 0.1, h: 0.35,
                fontSize: 10.5, bold: true, color: C.deepBlack,
                fontFace: "Calibri", align: "center"
            });
            s.addText(sp.weeks, {
                x: x + 0.05, y: cardY + 0.48, w: segW - 0.1, h: 0.25,
                fontSize: 9, color: C.mutedText,
                fontFace: "Calibri", align: "center"
            });
            // Story points badge
            addBadge(s, sp.pts, x + (segW - 0.9) / 2, cardY + 0.85, C.darkGreen, C.accentGreen);
        });

        // Total velocity
        s.addShape("roundRect", {
            x: 0.75, y: 4.3, w: 8.5, h: 0.55,
            fill: { color: C.lightGreenBg },
            line: { color: C.midGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("Vélocité totale : 130 Story Points  •  Moyenne : 26 SP/Sprint  •  Rétrospectives après chaque sprint", {
            x: 0.75, y: 4.3, w: 8.5, h: 0.55,
            fontSize: 11, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 09 — Stack Technique Moderne
    // =========================================================================
    {
        const s = createSlide("MÉTHODOLOGIE");

        s.addText("Stack Technique Moderne", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const stackItems = [
            {
                layer: "Frontend", items: [
                    [icons.react, "React Native (Expo)", "UI cross-platform iOS/Android"],
                ]
            },
            {
                layer: "Backend", items: [
                    [icons.node, "Node.js + Express", "API REST principale"],
                    [icons.python, "FastAPI (Python)", "Pipeline KYC/IA"],
                ]
            },
            {
                layer: "Data & Infra", items: [
                    [icons.database, "PostgreSQL", "Base de données relationnelle"],
                    [icons.bolt, "Redis", "Cache & sessions"],
                    [icons.docker, "Docker + Nginx", "Conteneurisation & reverse proxy"],
                ]
            },
        ];

        let yPos = 1.5;
        stackItems.forEach(({ layer, items }) => {
            // Layer label
            s.addShape("roundRect", {
                x: 0.55, y: yPos, w: 1.4, h: 0.3,
                fill: { color: C.darkGreen },
                line: { color: C.darkGreen, width: 0 },
                rectRadius: 0.06
            });
            s.addText(layer.toUpperCase(), {
                x: 0.55, y: yPos, w: 1.4, h: 0.3,
                fontSize: 9, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle"
            });
            // Items in row
            const itemW = (8.5 - 1.6) / items.length;
            items.forEach(([icon, name, desc], ii) => {
                const x = 2.1 + ii * itemW;
                s.addShape("roundRect", {
                    x, y: yPos - 0.1, w: itemW - 0.1, h: 0.5,
                    fill: { color: C.white },
                    line: { color: C.cardBorder, width: 0.5 },
                    rectRadius: 0.06
                });
                if (icon) s.addImage({ data: icon, x: x + 0.08, y: yPos - 0.02, w: 0.3, h: 0.3 });
                s.addText(name, {
                    x: x + 0.45, y: yPos - 0.08, w: itemW - 0.65, h: 0.22,
                    fontSize: 10.5, bold: true, color: C.deepBlack, fontFace: "Calibri"
                });
                s.addText(desc, {
                    x: x + 0.45, y: yPos + 0.12, w: itemW - 0.65, h: 0.2,
                    fontSize: 8.5, color: C.mutedText, fontFace: "Calibri"
                });
            });
            yPos += 0.75;
        });

        // Key differentiators
        const diffCards = [
            [icons.shieldG, "Sécurité", "JWT + Bcrypt + AES-256\nChiffrement de bout en bout"],
            [icons.compress, "Performance", "Redis caching, requêtes < 200ms\nOptimistic locking PostgreSQL"],
            [icons.circleCheck, "Qualité", "ESLint + Prettier\n147/151 tests Jest (97.4%)"],
        ];
        diffCards.forEach(([icon, title, body], i) => {
            addIconCard(s, icon, title, body, 0.55 + i * 3.15, 3.6, 2.95, 1.1);
        });
    }

    // =========================================================================
    // SLIDE 10 — Architecture Multi-Services Distribuée
    // =========================================================================
    {
        const s = createSlide("ARCHITECTURE");

        s.addText("Architecture Multi-Services Distribuée", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Architecture diagram
        s.addImage({
            path: IMG.archDiag,
            x: 0.5, y: 1.5, w: 9.0, h: 3.7,
            sizing: { type: "contain", w: 9.0, h: 3.7 }
        });
    }

    // =========================================================================
    // SLIDE 11 — Pipeline Biométrique KYC d'IA
    // =========================================================================
    {
        const s = createSlide("INNOVATION");

        s.addText("Pipeline Biométrique KYC d'IA", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // 3-stage pipeline
        const stages = [
            { num: "1", title: "OCR Document", desc: "Extraction des données\nd'identité depuis la CNI\n(nom, prénom, n° carte)", icon: icons.idCard },
            { num: "2", title: "ArcFace Match", desc: "Comparaison faciale\nphoto CNI vs selfie\nSeuil : 0.65 (conducteur)", icon: icons.brain },
            { num: "3", title: "Liveness Detection", desc: "Détection de vivacité\nanti-spoofing temps réel\nRejet des photos/vidéos", icon: icons.shieldG },
        ];

        stages.forEach((st, i) => {
            const x = 0.55 + i * 3.15;
            // Card
            s.addShape("roundRect", {
                x, y: 1.5, w: 2.95, h: 2.4,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
                rectRadius: 0.1
            });
            // Step number circle
            s.addShape("ellipse", {
                x: x + 1.15, y: 1.65, w: 0.55, h: 0.55,
                fill: { color: C.darkGreen },
                line: { color: C.accentGreen, width: 1.5 }
            });
            s.addText(st.num, {
                x: x + 1.15, y: 1.65, w: 0.55, h: 0.55,
                fontSize: 20, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle"
            });
            // Icon
            if (st.icon) s.addImage({ data: st.icon, x: x + 1.2, y: 2.3, w: 0.5, h: 0.5 });
            // Title
            s.addText(st.title, {
                x, y: 2.9, w: 2.95, h: 0.3,
                fontSize: 13, bold: true, color: C.deepBlack,
                fontFace: "Calibri", align: "center"
            });
            // Description
            s.addText(st.desc, {
                x: x + 0.15, y: 3.2, w: 2.65, h: 0.6,
                fontSize: 9.5, color: C.mutedText,
                fontFace: "Calibri", align: "center"
            });

            // Arrow between stages
            if (i < 2) {
                s.addText("→", {
                    x: x + 2.85, y: 2.4, w: 0.4, h: 0.4,
                    fontSize: 24, bold: true, color: C.accentGreen,
                    fontFace: "Calibri", align: "center", valign: "middle"
                });
            }
        });

        // Threshold info bar
        s.addShape("roundRect", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.6,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("Seuils de confiance :  Conducteur ≥ 0.65  |  Passager ≥ 0.60  |  Temps de traitement < 3 secondes", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.6,
            fontSize: 11, bold: true, color: C.accentGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // ─── SÉPARATION : Conception UML ─────────────────────────────────────────
    createTransitionSlide(4, "Conception & Modélisation", "Quels sont les modèles UML ?");

    // =========================================================================
    // SLIDE — Cas d'Utilisation Global (UML)
    // =========================================================================
    {
        const s = createSlide("CONCEPTION");

        s.addText("Cas d'Utilisation Global", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Actors legend
        const actors = [
            [icons.userLarge, "Passager", "Recherche, réserve, évalue"],
            [icons.steeringWheel, "Conducteur", "Publie, gère, KYC"],
            [icons.usersGear, "Admin", "Vérifie, modère, analytics"],
        ];
        actors.forEach(([icon, title, desc], i) => {
            const x = 0.55 + i * 1.55;
            if (icon) s.addImage({ data: icon, x: x + 0.05, y: 1.45, w: 0.28, h: 0.28 });
            s.addText(title, {
                x: x + 0.38, y: 1.42, w: 1.1, h: 0.18,
                fontSize: 9.5, bold: true, color: C.deepBlack, fontFace: "Calibri"
            });
            s.addText(desc, {
                x: x + 0.38, y: 1.6, w: 1.1, h: 0.18,
                fontSize: 7.5, color: C.mutedText, fontFace: "Calibri"
            });
        });

        // Use case diagram
        s.addImage({
            path: IMG.useCaseDiag,
            x: 0.5, y: 1.9, w: 9.0, h: 3.3,
            sizing: { type: "contain", w: 9.0, h: 3.3 }
        });
    }

    // =========================================================================
    // SLIDE 13 — Schéma Relationnel de Données (UML)
    // =========================================================================
    {
        const s = createSlide("CONCEPTION");

        s.addText("Schéma Relationnel de Données", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Core entities legend
        const entities = [
            "User  •  Trip  •  Booking  •  Identity  •  Payment  •  Vehicle  •  Review"
        ];
        s.addText(entities[0], {
            x: 0.5, y: 1.35, w: 9.0, h: 0.3,
            fontSize: 11, bold: true, color: C.midGreen,
            fontFace: "Courier New", align: "center", charSpacing: 1
        });

        // Class diagram
        s.addImage({
            path: IMG.classDiag,
            x: 0.5, y: 1.7, w: 9.0, h: 3.5,
            sizing: { type: "contain", w: 9.0, h: 3.5 }
        });
    }

    // =========================================================================
    // SLIDE 14 — Topologie du Déploiement Cloud (UML)
    // =========================================================================
    {
        const s = createSlide("CONCEPTION");

        s.addText("Topologie du Déploiement Cloud", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        s.addText("Docker • Nginx Reverse Proxy • PostgreSQL • Redis • FastAPI • Node.js", {
            x: 0.5, y: 1.35, w: 9.0, h: 0.3,
            fontSize: 10, bold: true, color: C.midGreen,
            fontFace: "Courier New", align: "center", charSpacing: 1
        });

        // Deployment diagram
        s.addImage({
            path: IMG.deployDiag,
            x: 0.5, y: 1.7, w: 9.0, h: 3.5,
            sizing: { type: "contain", w: 9.0, h: 3.5 }
        });
    }

    // ─── SÉPARATION : Résultats & Validation ─────────────────────────────────
    createTransitionSlide(5, "Résultats & Validation", "Qu'avons-nous obtenu ?");

    // =========================================================================
    // SLIDE — 97.4% de Réussite aux Tests (Jest)
    // =========================================================================
    {
        const s = createSlide("RÉSULTATS");

        s.addText("97.4% de Réussite aux Tests Automatisés", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Big stat
        s.addText("147/151", {
            x: 0.75, y: 1.5, w: 3.5, h: 0.8,
            fontSize: 48, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
        s.addText("Tests réussis", {
            x: 0.75, y: 2.25, w: 3.5, h: 0.3,
            fontSize: 14, color: C.mutedText,
            fontFace: "Calibri", align: "center"
        });

        // Bar chart simulation — Unit Tests
        s.addText("Tests Unitaires", {
            x: 4.8, y: 1.55, w: 4.5, h: 0.25,
            fontSize: 10, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });
        s.addShape("roundRect", {
            x: 4.8, y: 1.85, w: 4.5, h: 0.35,
            fill: { color: C.lightGreenBg },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.05
        });
        s.addShape("roundRect", {
            x: 4.8, y: 1.85, w: 4.5, h: 0.35,
            fill: { color: C.midGreen },
            line: { color: C.midGreen, width: 0 },
            rectRadius: 0.05
        });
        s.addText("107 / 107  (100%)", {
            x: 4.8, y: 1.85, w: 4.5, h: 0.35,
            fontSize: 10, bold: true, color: C.white,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Bar chart simulation — Integration Tests
        s.addText("Tests d'Intégration", {
            x: 4.8, y: 2.4, w: 4.5, h: 0.25,
            fontSize: 10, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });
        s.addShape("roundRect", {
            x: 4.8, y: 2.7, w: 4.5, h: 0.35,
            fill: { color: C.lightGreenBg },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.05
        });
        const integW = 4.5 * (40 / 44);
        s.addShape("roundRect", {
            x: 4.8, y: 2.7, w: integW, h: 0.35,
            fill: { color: C.accentGreen },
            line: { color: C.accentGreen, width: 0 },
            rectRadius: 0.05
        });
        s.addText("40 / 44  (90.9%)", {
            x: 4.8, y: 2.7, w: 4.5, h: 0.35,
            fontSize: 10, bold: true, color: C.deepBlack,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Explanation of 4 failures
        s.addShape("roundRect", {
            x: 0.75, y: 3.35, w: 8.5, h: 0.85,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("Les 4 échecs : Timeouts réseau sur les tests d'intégration API", {
            x: 0.75, y: 3.4, w: 8.5, h: 0.25,
            fontSize: 10, bold: true, color: C.deepBlack, fontFace: "Calibri", align: "center"
        });
        s.addText("• 2 timeouts sur les tests de réservation concurrente (race condition en CI)\n• 2 timeouts sur les tests de notification push (dépendance réseau externe)\n→ Aucun bug fonctionnel : tous les échecs sont liés à l'environnement CI, non au code.", {
            x: 1.0, y: 3.65, w: 8.0, h: 0.5,
            fontSize: 9, color: C.mutedText, fontFace: "Calibri", align: "center"
        });

        // Test coverage categories
        const testCats = [
            [icons.code, "Auth & Profils", "28 tests"],
            [icons.route, "Trajets & Réservation", "35 tests"],
            [icons.brain, "Pipeline KYC", "24 tests"],
            [icons.card, "Paiement", "22 tests"],
            [icons.usersGear, "Admin", "38 tests"],
        ];
        testCats.forEach(([icon, title, count], i) => {
            const x = 0.55 + i * 1.85;
            s.addShape("roundRect", {
                x, y: 4.45, w: 1.7, h: 0.65,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                rectRadius: 0.06
            });
            if (icon) s.addImage({ data: icon, x: x + 0.05, y: 4.55, w: 0.22, h: 0.22 });
            s.addText(title, {
                x: x + 0.3, y: 4.5, w: 1.35, h: 0.22,
                fontSize: 8.5, bold: true, color: C.deepBlack, fontFace: "Calibri"
            });
            s.addText(count, {
                x: x + 0.3, y: 4.72, w: 1.35, h: 0.22,
                fontSize: 8, color: C.mutedText, fontFace: "Calibri"
            });
        });
    }

    // =========================================================================
    // SLIDE 16 — Score SUS de 71.6 : Usabilité Validée
    // =========================================================================
    {
        const s = createSlide("RÉSULTATS");

        s.addText("Score SUS de 71.6 : Usabilité Validée", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // SUS Gauge visual
        s.addShape("roundRect", {
            x: 0.75, y: 1.5, w: 3.5, h: 2.5,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "0A1F14", opacity: 0.05 },
            rectRadius: 0.1
        });

        // Score circle
        s.addShape("ellipse", {
            x: 1.4, y: 1.7, w: 1.7, h: 1.7,
            fill: { color: C.lightGreenBg },
            line: { color: C.midGreen, width: 3 }
        });
        s.addText("71.6", {
            x: 1.4, y: 1.7, w: 1.7, h: 1.2,
            fontSize: 38, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
        s.addText("/ 100", {
            x: 1.4, y: 2.65, w: 1.7, h: 0.35,
            fontSize: 14, color: C.mutedText,
            fontFace: "Calibri", align: "center"
        });
        addBadge(s, "GOOD", 1.85, 3.55, C.midGreen, C.white, 0.8);

        // Breakdown
        s.addText("Résultats par profil", {
            x: 4.6, y: 1.55, w: 4.5, h: 0.3,
            fontSize: 12, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Passenger bar
        s.addText("Passagers (N=10)", {
            x: 4.6, y: 1.95, w: 2.5, h: 0.22,
            fontSize: 9.5, color: C.deepBlack, fontFace: "Calibri"
        });
        s.addShape("roundRect", {
            x: 4.6, y: 2.2, w: 4.5, h: 0.3,
            fill: { color: C.lightGreenBg },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.05
        });
        s.addShape("roundRect", {
            x: 4.6, y: 2.2, w: 4.5 * 0.742, h: 0.3,
            fill: { color: C.midGreen },
            line: { color: C.midGreen, width: 0 },
            rectRadius: 0.05
        });
        s.addText("74.2", {
            x: 4.6, y: 2.2, w: 4.5 * 0.742, h: 0.3,
            fontSize: 10, bold: true, color: C.white,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Driver bar
        s.addText("Conducteurs (N=8)", {
            x: 4.6, y: 2.65, w: 2.5, h: 0.22,
            fontSize: 9.5, color: C.deepBlack, fontFace: "Calibri"
        });
        s.addShape("roundRect", {
            x: 4.6, y: 2.9, w: 4.5, h: 0.3,
            fill: { color: C.lightGreenBg },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.05
        });
        s.addShape("roundRect", {
            x: 4.6, y: 2.9, w: 4.5 * 0.689, h: 0.3,
            fill: { color: C.accentGreen },
            line: { color: C.accentGreen, width: 0 },
            rectRadius: 0.05
        });
        s.addText("68.9", {
            x: 4.6, y: 2.9, w: 4.5 * 0.689, h: 0.3,
            fontSize: 10, bold: true, color: C.deepBlack,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // SUS scale reference
        s.addShape("roundRect", {
            x: 4.6, y: 3.4, w: 4.5, h: 0.5,
            fill: { color: C.white },
            line: { color: C.cardBorder, width: 0.5 },
            rectRadius: 0.06
        });
        s.addText("Échelle SUS : 0-50 Poor | 50-68 OK | 68-80 Good | 80+ Excellent", {
            x: 4.6, y: 3.4, w: 4.5, h: 0.5,
            fontSize: 8.5, color: C.mutedText,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        // Key insight
        s.addShape("roundRect", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.6,
            fill: { color: C.lightGreenBg },
            line: { color: C.midGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("→  Écart Passager/Conducteur (5.3 pts) : le flux KYC conducteur est plus complexe — priorité UX pour v2", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.6,
            fontSize: 11, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 17 — Validation des Trois Hypothèses
    // =========================================================================
    {
        const s = createSlide("RÉSULTATS");

        s.addText("Validation des Trois Hypothèses", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const validations = [
            {
                id: "H1", title: "KYC Biométrique",
                criterion: "Seuil ≥ 0.65 (conducteur), ≥ 0.60 (passager)",
                result: "Conducteur : 0.72  |  Passager : 0.68",
                status: "✓ VALIDÉE", statusColor: C.midGreen
            },
            {
                id: "H2", title: "Usabilité SUS",
                criterion: "Score SUS ≥ 68/100",
                result: "Score global : 71.6/100 (Good)",
                status: "✓ VALIDÉE", statusColor: C.midGreen
            },
            {
                id: "H3", title: "Résilience API",
                criterion: "Disponibilité ≥ 99%",
                result: "97.4% tests Jest + circuit breaker actif",
                status: "⚠ PARTIELLE", statusColor: C.orange
            },
        ];

        // Table header
        addTableRow(s, ["Hypothèse", "Critère de Validation", "Résultat Empirique", "Statut"], 1.5, 0.4, true);

        validations.forEach((v, i) => {
            const y = 1.9 + i * 0.75;

            // Row background
            s.addShape("rect", {
                x: 0.75, y, w: 8.5, h: 0.7,
                fill: { color: i % 2 === 0 ? C.white : C.lightGreenBg },
                line: { color: C.cardBorder, width: 0.5 }
            });

            // H-id
            s.addShape("roundRect", {
                x: 0.9, y: y + 0.1, w: 0.45, h: 0.45,
                fill: { color: C.darkGreen }, line: { color: C.darkGreen, width: 0 }, rectRadius: 0.06
            });
            s.addText(v.id, {
                x: 0.9, y: y + 0.1, w: 0.45, h: 0.45,
                fontSize: 14, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle"
            });
            s.addText(v.title, {
                x: 1.45, y: y + 0.1, w: 1.15, h: 0.45,
                fontSize: 10, bold: true, color: C.deepBlack,
                fontFace: "Calibri", valign: "middle"
            });

            // Criterion
            s.addText(v.criterion, {
                x: 2.875, y: y + 0.05, w: 2.45, h: 0.6,
                fontSize: 9.5, color: C.mutedText, fontFace: "Calibri", valign: "middle", align: "center"
            });

            // Result
            s.addText(v.result, {
                x: 5.325, y: y + 0.05, w: 2.2, h: 0.6,
                fontSize: 9.5, bold: true, color: C.deepBlack, fontFace: "Calibri", valign: "middle", align: "center"
            });

            // Status badge
            addBadge(s, v.status, 7.75, y + 0.22, v.statusColor, C.white, 1.3);
        });

        // Summary insight
        s.addShape("roundRect", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("H1 & H2 pleinement validées  •  H3 partiellement validée (environnement sandbox)\n→ Validation complète requiert un déploiement en production avec monitoring APM", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fontSize: 10.5, bold: true, color: C.accentGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 18 — Démonstration en Direct du Prototype
    // =========================================================================
    {
        const s = createSlide("DÉMONSTRATION", true);

        s.addText("Démonstration du Prototype", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.cream, fontFace: "Calibri"
        });

        // Composite screenshot layout — 4 key screens
        const screens = [
            { path: IMG.splash, label: "Accueil" },
            { path: IMG.roleSelect, label: "Inscription" },
            { path: IMG.searchResults, label: "Réservation" },
            { path: IMG.qrScan, label: "QR Code Départ" },
        ];

        screens.forEach((scr, i) => {
            const x = 0.55 + i * 2.35;
            // Phone frame
            s.addShape("roundRect", {
                x, y: 1.5, w: 2.1, h: 3.2,
                fill: { color: C.deepBlack },
                line: { color: C.midGreen, width: 1 },
                rectRadius: 0.15
            });
            s.addImage({
                path: scr.path,
                x: x + 0.08, y: 1.6, w: 1.94, h: 2.85,
                rounding: true,
                sizing: { type: "cover", w: 1.94, h: 2.85 }
            });
            // Label
            s.addText(scr.label, {
                x, y: 4.75, w: 2.1, h: 0.28,
                fontSize: 10, bold: true, color: C.accentGreen,
                fontFace: "Calibri", align: "center"
            });
        });

        // Flow arrows
        for (let i = 0; i < 3; i++) {
            const x = 0.55 + (i + 1) * 2.35 - 0.3;
            s.addText("→", {
                x, y: 2.8, w: 0.4, h: 0.4,
                fontSize: 20, bold: true, color: C.gold,
                fontFace: "Calibri", align: "center", valign: "middle"
            });
        }

        // CTA
        s.addShape("rect", { x: 0, y: 5.12, w: SW, h: 0.5, fill: { color: C.accentGreen }, line: { color: C.accentGreen, width: 0 } });
        s.addText("▶  DÉMONSTRATION EN DIRECT  —  Flux Passager Complet", {
            x: 0, y: 5.12, w: SW, h: 0.5,
            fontSize: 14, bold: true, color: C.deepBlack,
            align: "center", valign: "middle", fontFace: "Calibri"
        });
    }

    // ─── SÉPARATION : Modèle Économique ──────────────────────────────────────
    createTransitionSlide(6, "Modèle Économique", "Comment rentabiliser le projet ?");

    // =========================================================================
    // SLIDE — Business Model Canvas (9 Blocs)
    // =========================================================================
    {
        const s = createSlide("MODÈLE ÉCONOMIQUE");

        s.addText("Business Model Canvas", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // BMC Grid — simplified 3x3 approach
        const bmcBlocks = [
            { title: "Partenaires Clés", items: "• SATIM/GIE Monétique\n• Opérateurs télécoms\n• Universités", x: 0.55, y: 1.5, w: 2.0, h: 1.65 },
            { title: "Activités Clés", items: "• Matching conducteur/passager\n• Vérification KYC\n• Paiement sécurisé", x: 2.65, y: 1.5, w: 2.0, h: 0.8 },
            { title: "Proposition de Valeur", items: "• Covoiturage sécurisé KYC\n• 69 wilayas couvertes\n• Paiement local intégré\n• Tarif transparent", x: 4.75, y: 1.5, w: 2.0, h: 1.65 },
            { title: "Relation Client", items: "• Support in-app\n• Notation mutuelle\n• Notifications push", x: 6.85, y: 1.5, w: 2.6, h: 0.8 },
            { title: "Ressources Clés", items: "• Équipe dev (2 devs)\n• Serveur VPS\n• APIs IA (ArcFace)", x: 2.65, y: 2.4, w: 2.0, h: 0.75 },
            { title: "Canaux", items: "• App Store / Google Play\n• Réseaux sociaux\n• Campus universitaires", x: 6.85, y: 2.4, w: 2.6, h: 0.75 },
            { title: "Segments Clients", items: "• Étudiants inter-wilayas\n• Professionnels\n• Familles voyageuses", x: 6.85, y: 1.5, w: 2.6, h: 0.0 },
        ];

        // Draw only the main 5-column BMC layout
        const bmcData = [
            { title: "Partenaires", content: "SATIM • Télécoms\nUniversités", x: 0.45, y: 1.45, w: 1.7, h: 1.85 },
            { title: "Activités", content: "Matching\nKYC\nPaiement", x: 2.2, y: 1.45, w: 1.55, h: 0.88 },
            { title: "Ressources", content: "2 devs • VPS\nAPIs IA", x: 2.2, y: 2.38, w: 1.55, h: 0.92 },
            { title: "Valeur", content: "Covoiturage KYC\n69 wilayas\nPaiement local\nTarif transparent", x: 3.8, y: 1.45, w: 1.9, h: 1.85 },
            { title: "Relations", content: "Support in-app\nNotation mutuelle", x: 5.75, y: 1.45, w: 1.55, h: 0.88 },
            { title: "Canaux", content: "Stores mobiles\nRéseaux sociaux\nCampus", x: 5.75, y: 2.38, w: 1.55, h: 0.92 },
            { title: "Segments", content: "Étudiants\nProfessionnels\nFamilles", x: 7.35, y: 1.45, w: 2.1, h: 1.85 },
            { title: "Coûts", content: "Hébergement VPS • API tiers • Marketing digital", x: 0.45, y: 3.4, w: 4.2, h: 0.7 },
            { title: "Revenus", content: "Commission 12% • Abonnement Premium • Partenariats", x: 4.7, y: 3.4, w: 4.75, h: 0.7 },
        ];

        bmcData.forEach(block => {
            s.addShape("roundRect", {
                x: block.x, y: block.y, w: block.w, h: block.h,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                rectRadius: 0.06
            });
            // Title
            s.addShape("rect", {
                x: block.x, y: block.y, w: block.w, h: 0.28,
                fill: { color: C.darkGreen },
                line: { color: C.darkGreen, width: 0 }
            });
            s.addText(block.title.toUpperCase(), {
                x: block.x, y: block.y, w: block.w, h: 0.28,
                fontSize: 8, bold: true, color: C.accentGreen,
                fontFace: "Courier New", align: "center", valign: "middle", charSpacing: 1
            });
            // Content
            s.addText(block.content, {
                x: block.x + 0.08, y: block.y + 0.32, w: block.w - 0.16, h: block.h - 0.38,
                fontSize: 8, color: C.mutedText, fontFace: "Calibri", valign: "top"
            });
        });

        // Highlight bar
        s.addShape("roundRect", {
            x: 0.75, y: 4.35, w: 8.5, h: 0.5,
            fill: { color: C.lightGreenBg },
            line: { color: C.midGreen, width: 0.5 },
            rectRadius: 0.06
        });
        s.addText("Modèle inspiré du BMC de BlaBlaCar, adapté au contexte algérien (paiement local + KYC souverain)", {
            x: 0.75, y: 4.35, w: 8.5, h: 0.5,
            fontSize: 10, bold: true, color: C.midGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 20 — Modèle de Revenus & Prévisions
    // =========================================================================
    {
        const s = createSlide("MODÈLE ÉCONOMIQUE");

        s.addText("Modèle de Revenus & Prévisions Financières", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Revenue model cards
        const revModels = [
            [icons.trending, "Commission 12%", "Sur chaque trajet réservé\nvia la plateforme"],
            [icons.star, "Premium Driver", "Abonnement mensuel 500 DA\nVisibilité prioritaire"],
            [icons.handshake, "Partenariats", "Stations-service, assurances,\nopérateurs télécoms"],
        ];
        revModels.forEach(([icon, title, desc], i) => {
            addIconCard(s, icon, title, desc, 0.55 + i * 3.15, 1.5, 2.95, 0.95);
        });

        // 3-year forecast table
        s.addText("Prévisions sur 3 ans", {
            x: 0.55, y: 2.7, w: 9.0, h: 0.3,
            fontSize: 12, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        addTableRow(s, ["Indicateur", "Année 1", "Année 2", "Année 3"], 3.05, 0.35, true);
        addTableRow(s, ["Utilisateurs actifs", "5 000", "25 000", "80 000"], 3.4);
        addTableRow(s, ["Trajets/mois", "2 000", "12 000", "40 000"], 3.78);
        addTableRow(s, ["CA estimé (DA)", "3.6M", "21.6M", "72M"], 4.16);
        addTableRow(s, ["Seuil rentabilité", "—", "✓ Atteint", "—"], 4.54);
    }

    // =========================================================================
    // SLIDE 21 — Propriété Intellectuelle (IP)
    // =========================================================================
    {
        const s = createSlide("MODÈLE ÉCONOMIQUE");

        s.addText("Propriété Intellectuelle & Protection Légale", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 28, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Two main IP cards
        const ipCards = [
            {
                icon: icons.award, title: "ONDA — Enregistrement du Code Source",
                desc: "Dépôt officiel du code source complet de la plateforme RohWinBghit auprès de l'Office National des Droits d'Auteur et Droits Voisins.",
                status: "Certificat Obtenu", statusColor: C.midGreen
            },
            {
                icon: icons.flag, title: "INAPI — Dépôt de Marque Commerciale",
                desc: "Enregistrement de la marque « RohWinBghit » et du logo auprès de l'Institut National Algérien de la Propriété Industrielle.",
                status: "En Cours", statusColor: C.orange
            },
        ];

        ipCards.forEach((ip, i) => {
            const y = 1.55 + i * 1.5;
            s.addShape("roundRect", {
                x: 0.75, y, w: 8.5, h: 1.25,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                shadow: { type: "outer", blur: 6, offset: 1, angle: 135, color: "0A1F14", opacity: 0.04 },
                rectRadius: 0.1
            });
            // Icon circle
            s.addShape("ellipse", {
                x: 1.0, y: y + 0.22, w: 0.7, h: 0.7,
                fill: { color: C.lightGreenBg },
                line: { color: C.midGreen, width: 1 }
            });
            if (ip.icon) s.addImage({ data: ip.icon, x: 1.12, y: y + 0.34, w: 0.46, h: 0.46 });
            // Title
            s.addText(ip.title, {
                x: 1.95, y: y + 0.15, w: 5.5, h: 0.3,
                fontSize: 14, bold: true, color: C.deepBlack, fontFace: "Calibri"
            });
            // Status badge
            addBadge(s, ip.status, 7.8, y + 0.2, ip.statusColor, C.white, 1.3);
            // Description
            s.addText(ip.desc, {
                x: 1.95, y: y + 0.5, w: 7.0, h: 0.6,
                fontSize: 10.5, color: C.mutedText, fontFace: "Calibri"
            });
        });

        // Protection strategy summary
        s.addShape("roundRect", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("Stratégie IP : Code protégé (ONDA) + Marque déposée (INAPI) + Licence MIT pour la communauté open source", {
            x: 0.75, y: 4.25, w: 8.5, h: 0.65,
            fontSize: 11, bold: true, color: C.accentGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // ─── SÉPARATION : Conclusion ─────────────────────────────────────────────
    createTransitionSlide(7, "Conclusion & Perspectives", "Quel bilan ? Quelles perspectives ?");

    // =========================================================================
    // SLIDE — Bilan & Contributions Académiques
    // =========================================================================
    {
        const s = createSlide("CONCLUSION");

        s.addText("Bilan & Contributions Académiques", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        const contributions = [
            [icons.brain, "Infrastructure de Confiance", "Premier pipeline KYC biométrique (OCR + ArcFace + Liveness) adapté au marché algérien de la mobilité."],
            [icons.shieldG, "IA Souveraine", "Modèles d'IA traitant les données localement en conformité avec la Loi 25-11, sans transfert transfrontalier."],
            [icons.card, "Monétique Locale", "Intégration SATIM/Edahabia/CIB, première plateforme de covoiturage avec paiement électronique algérien."],
            [icons.circleCheck, "Validation Empirique", "97.4% tests Jest réussis + SUS 71.6/100 validant l'approche technique et l'expérience utilisateur."],
        ];

        contributions.forEach(([icon, title, desc], i) => {
            const y = 1.5 + i * 0.95;
            s.addShape("roundRect", {
                x: 0.75, y, w: 8.5, h: 0.8,
                fill: { color: i % 2 === 0 ? C.white : C.lightGreenBg },
                line: { color: C.cardBorder, width: 0.5 },
                rectRadius: 0.08
            });
            if (icon) s.addImage({ data: icon, x: 0.95, y: y + 0.18, w: 0.4, h: 0.4 });
            s.addText(title, {
                x: 1.5, y: y + 0.08, w: 3.0, h: 0.28,
                fontSize: 12, bold: true, color: C.deepBlack, fontFace: "Calibri"
            });
            s.addText(desc, {
                x: 1.5, y: y + 0.38, w: 7.5, h: 0.38,
                fontSize: 9.5, color: C.mutedText, fontFace: "Calibri"
            });
        });
    }

    // =========================================================================
    // SLIDE 23 — Limites & Perspectives d'Évolution
    // =========================================================================
    {
        const s = createSlide("CONCLUSION");

        s.addText("Limites & Perspectives d'Évolution", {
            x: 0.5, y: 0.75, w: 9, h: 0.55,
            fontSize: 30, bold: true, color: C.deepBlack, fontFace: "Calibri"
        });

        // Limits (left column)
        s.addText("LIMITES IDENTIFIÉES", {
            x: 0.55, y: 1.45, w: 4.2, h: 0.25,
            fontSize: 10, bold: true, color: C.red,
            fontFace: "Courier New", charSpacing: 1
        });

        const limits = [
            "Paiement en mode sandbox (non connecté à la production SATIM)",
            "Échantillon SUS limité (N=18 utilisateurs)",
            "Tests de charge non réalisés en conditions réelles",
            "Pipeline KYC testé sur dataset restreint",
        ];
        limits.forEach((lim, i) => {
            const y = 1.8 + i * 0.55;
            s.addShape("roundRect", {
                x: 0.55, y, w: 4.2, h: 0.45,
                fill: { color: C.white },
                line: { color: C.cardBorder, width: 0.5 },
                rectRadius: 0.06
            });
            s.addText("⚠", {
                x: 0.65, y, w: 0.3, h: 0.45,
                fontSize: 12, color: C.orange,
                fontFace: "Calibri", align: "center", valign: "middle"
            });
            s.addText(lim, {
                x: 1.0, y, w: 3.7, h: 0.45,
                fontSize: 9, color: C.deepBlack,
                fontFace: "Calibri", valign: "middle"
            });
        });

        // Perspectives (right column)
        s.addText("PERSPECTIVES FUTURES", {
            x: 5.1, y: 1.45, w: 4.35, h: 0.25,
            fontSize: 10, bold: true, color: C.midGreen,
            fontFace: "Courier New", charSpacing: 1
        });

        const perspectives = [
            "Tarification dynamique (surge pricing) par ML",
            "Phase pilote : Tlemcen → Oran → Alger",
            "Intégration paiement production SATIM/CIB",
            "Extension vers le transport de colis",
        ];
        perspectives.forEach((persp, i) => {
            const y = 1.8 + i * 0.55;
            s.addShape("roundRect", {
                x: 5.1, y, w: 4.35, h: 0.45,
                fill: { color: C.lightGreenBg },
                line: { color: C.midGreen, width: 0.5 },
                rectRadius: 0.06
            });
            s.addText("→", {
                x: 5.2, y, w: 0.3, h: 0.45,
                fontSize: 12, bold: true, color: C.midGreen,
                fontFace: "Calibri", align: "center", valign: "middle"
            });
            s.addText(persp, {
                x: 5.55, y, w: 3.8, h: 0.45,
                fontSize: 9, color: C.deepBlack,
                fontFace: "Calibri", valign: "middle"
            });
        });

        // Bottom vision statement
        s.addShape("roundRect", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.7,
            fill: { color: C.darkGreen },
            line: { color: C.accentGreen, width: 0.5 },
            rectRadius: 0.08
        });
        s.addText("Vision : Devenir la référence algérienne du covoiturage sécurisé inter-wilayas,\nen démocratisant la mobilité partagée à travers les 69 wilayas du pays.", {
            x: 0.75, y: 4.15, w: 8.5, h: 0.7,
            fontSize: 11, bold: true, color: C.accentGreen,
            fontFace: "Calibri", align: "center", valign: "middle"
        });
    }

    // =========================================================================
    // SLIDE 24 — Merci pour Votre Attention
    // =========================================================================
    {
        const s = createSlide("FIN", true);

        // Central thank you
        s.addImage({ path: IMG.logo, x: 4.3, y: 1.2, w: 1.4, h: 1.4 });

        s.addText("Merci pour Votre Attention", {
            x: 0.5, y: 2.7, w: 9.0, h: 0.7,
            fontSize: 34, bold: true, color: C.cream,
            fontFace: "Calibri", align: "center", valign: "middle"
        });

        s.addShape("rect", {
            x: 4.0, y: 3.45, w: 2.0, h: 0.03,
            fill: { color: C.gold }, line: { color: C.gold, width: 0 }
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
    console.log("✅  Présentation générée avec succès :", outPath);
    console.log(`📊  Total : ${TOTAL_SLIDES} slides (optimisé pour 20 min de soutenance)`);
}

buildPresentation().catch(err => { console.error("❌ Erreur :", err); process.exit(1); });