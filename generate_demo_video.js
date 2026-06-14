#!/usr/bin/env node
/**
 * ═══════════════════════════════════════════════════════════════════════
 *   RohWinBghit — Générateur Vidéo Cinématique v3
 *   Soutenance Master 2 · Génie Logiciel · Tlemcen 2025-2026
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  Nouveautés v3 :
 *    • Multi-frame animation (30 fps — motion réel par scène)
 *    • Ken Burns zoom sur tous les écrans téléphone
 *    • Fondu enchaîné (fade in/out) entre les scènes
 *    • Effet scan KYC (ligne animée + marqueurs de coin)
 *    • Pulsation QR Code (halo doré)
 *    • Carte Algérie avec villes et routes animées (Section 1)
 *    • Infographie Problématique avec cards dynamiques (Section 2)
 *    • Dashboard Admin SVG généré (Section 9)
 *    • Architecture technique animée (Section 10)
 *    • KPI Counter-up animé (Section 11)
 *    • Outro cinématique avec cercles pulsants (Section 12)
 *    • Rendu parallèle (8 threads simultanés)
 *    • Vignette cinématique FFmpeg
 *
 *  Usage : node generate_demo_video.js
 *  Sortie : RohWinBghit_Demo.mp4  (~2 min 15 s, Full HD 1920×1080)
 */

'use strict';

const sharp      = require('sharp');
const ffmpeg     = require('fluent-ffmpeg');
const ffmpegPath = require('@ffmpeg-installer/ffmpeg').path;
const path       = require('path');
const fs         = require('fs');

ffmpeg.setFfmpegPath(ffmpegPath);

// ─── CHEMINS ─────────────────────────────────────────────────────────────────
const BASE_DIR    = __dirname;
const SCREEN_DIR  = path.join(BASE_DIR, 'Screen_mobile');
const MOBILE_DIR  = path.join(BASE_DIR, 'mobile-all');
const TEMP_DIR    = path.join(BASE_DIR, 'video_temp_v3');
const OUTPUT_FILE = path.join(BASE_DIR, 'RohWinBghit_Demo.mp4');
const CONCAT_FILE = path.join(TEMP_DIR, 'concat_list.txt');

if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

// ─── DIMENSIONS VIDÉO ────────────────────────────────────────────────────────
const VW  = 1920;
const VH  = 1080;
const FPS = 30;

// Zone téléphone (côté droit, centré verticalement)
const PH_W       = 388;
const PH_H       = 776;
const PH_BEZEL   = 24;
const PH_CORNER  = 46;
const PH_TOTAL_W = PH_W + PH_BEZEL * 2;
const PH_TOTAL_H = PH_H + PH_BEZEL * 2 + 80;
const PH_LEFT    = VW - PH_TOTAL_W - 90;
const PH_TOP     = Math.floor((VH - PH_TOTAL_H) / 2);
const PH_SCR_L   = PH_LEFT + PH_BEZEL;
const PH_SCR_T   = PH_TOP  + PH_BEZEL + 40;

// ─── PALETTE FOREST & MOSS ───────────────────────────────────────────────────
const C = {
  bgDark:  '#0B2E1E',
  bgMid:   '#1C4933',
  primary: '#2E7D52',
  accent:  '#4CAF72',
  gold:    '#C9A84C',
  white:   '#FFFFFF',
  muted:   '#A8C5B5',
  phone:   '#091F14',
  teal:    '#028090',
  red:     '#C0392B',
  cyan:    '#00C9D4',
  purple:  '#7C3AED',
};

const rgb = hex => ({
  r: parseInt(hex.slice(1, 3), 16),
  g: parseInt(hex.slice(3, 5), 16),
  b: parseInt(hex.slice(5, 7), 16),
});

/** Échapper les caractères XML + supprimer les emojis (évite crash Pango) */
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u{1F000}-\u{1FFFF}]/gu, '')
    .replace(/[\u2700-\u27BF]/g, '')
    .replace(/[\u2600-\u26FF]/g, '');
}

// ─── FONCTIONS D'ANIMATION ───────────────────────────────────────────────────
const ease = {
  inOut:  t => t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t,
  out:    t => 1 - (1 - t) * (1 - t),
  in:     t => t * t,
  linear: t => t,
};

const lerp = (a, b, t) => a + (b - a) * t;

/** Calcule l'opacité globale d'une frame (fade in + fade out) */
function fadeAlpha(t, fadeIn = 0.18, fadeOut = 0.15) {
  if (t < fadeIn) return ease.out(t / fadeIn);
  if (t > 1 - fadeOut) return ease.out((1 - t) / fadeOut);
  return 1.0;
}

// ─── LIMITEUR DE CONCURRENCE ─────────────────────────────────────────────────
function createLimiter(concurrency) {
  let running = 0;
  const queue = [];
  return fn => new Promise((resolve, reject) => {
    const run = async () => {
      running++;
      try { resolve(await fn()); }
      catch (e) { reject(e); }
      finally {
        running--;
        if (queue.length > 0) queue.shift()();
      }
    };
    if (running < concurrency) run();
    else queue.push(run);
  });
}

// ─── RÉSOLUTION D'IMAGE ──────────────────────────────────────────────────────
function resolveImg(filename) {
  if (!filename) return null;
  const variants = [filename, filename.normalize('NFC'), filename.normalize('NFD')];
  for (const dir of [SCREEN_DIR, MOBILE_DIR]) {
    if (!fs.existsSync(dir)) continue;
    for (const v of variants) {
      const p = path.join(dir, v);
      if (fs.existsSync(p)) return p;
    }
    // Scan pour problèmes d'encodage
    for (const f of fs.readdirSync(dir)) {
      if (f.normalize('NFC') === filename.normalize('NFC')) return path.join(dir, f);
    }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════════════════
//  GÉNÉRATEURS SVG — COMPOSANTS RÉUTILISABLES
// ════════════════════════════════════════════════════════════════════════════

/** Cadre téléphone avec ombre portée et glow de couleur */
function phoneBezelSVG(glowColor = C.accent, glowOpacity = 0.18) {
  const W = PH_TOTAL_W, H = PH_TOTAL_H, R = PH_CORNER;
  const sX = PH_BEZEL, sY = PH_BEZEL + 40;
  return Buffer.from(`<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <filter id="ph-shadow">
        <feDropShadow dx="0" dy="12" stdDeviation="28"
                      flood-color="${glowColor}" flood-opacity="${glowOpacity}"/>
        <feDropShadow dx="0" dy="3" stdDeviation="8"
                      flood-color="#000000" flood-opacity="0.55"/>
      </filter>
    </defs>
    <!-- Corps téléphone -->
    <rect x="3" y="3" width="${W - 6}" height="${H - 6}" rx="${R}" ry="${R}"
          fill="${C.phone}" stroke="${C.primary}" stroke-width="1.5"
          filter="url(#ph-shadow)"/>
    <!-- Bordure glow subtile -->
    <rect x="3" y="3" width="${W - 6}" height="${H - 6}" rx="${R}" ry="${R}"
          fill="none" stroke="${glowColor}" stroke-width="1" opacity="${glowOpacity * 2.5}"/>
    <!-- Zone écran -->
    <rect x="${sX}" y="${sY}" width="${PH_W}" height="${PH_H}"
          rx="${R - 14}" ry="${R - 14}" fill="white"/>
    <!-- Notch caméra -->
    <rect x="${W / 2 - 28}" y="${sY - 28}" width="56" height="14" rx="7" fill="${C.bgDark}"/>
    <circle cx="${W / 2}" cy="${sY - 21}" r="4" fill="${C.bgDark}"/>
    <!-- Indicateur home -->
    <rect x="${W / 2 - 48}" y="${H - 22}" width="96" height="5" rx="3"
          fill="${C.primary}" opacity="0.55"/>
    <!-- Boutons latéraux -->
    <rect x="-3" y="${H * 0.27}" width="5" height="46" rx="3" fill="${C.bgMid}"/>
    <rect x="-3" y="${H * 0.40}" width="5" height="30" rx="3" fill="${C.bgMid}"/>
    <rect x="${W - 2}" y="${H * 0.33}" width="5" height="62" rx="3" fill="${C.bgMid}"/>
  </svg>`);
}

/** Fond dégradé gauche→droite avec lueur radiale colorée */
function backgroundSVG(tagColor = C.accent) {
  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
      <radialGradient id="hl" cx="75%" cy="50%" r="40%">
        <stop offset="0%"   stop-color="${tagColor}" stop-opacity="0.10"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#hl)"/>
    <rect x="0" y="0"        width="${VW}" height="5" fill="${tagColor}" opacity="0.55"/>
    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${tagColor}" opacity="0.18"/>
  </svg>`);
}

/** Panneau texte gauche avec badge, titre, narration, barre de progression */
function leftPanelSVG(opts) {
  const { tag, tagColor, title, narration, sceneLabel, actNum, progress, alpha = 1.0 } = opts;
  const titleLines = title.split('\n');
  const TY = 265, TH = 72;
  const narY = TY + titleLines.length * TH + 52;
  const NAR_H = 34;
  const panelW = PH_LEFT - 20;
  const actStr = String(actNum).padStart(2, '0');
  const progW  = Math.max(0, Math.round(progress * (panelW - 180)));

  const titleSVG = titleLines.map((l, i) =>
    `<text x="80" y="${TY + i * TH}"
       font-family="Arial,Helvetica,sans-serif"
       font-size="58" font-weight="bold"
       fill="${C.white}" opacity="${alpha}">${esc(l)}</text>`
  ).join('\n');

  const narSVG = narration.filter(l => l !== undefined).map((l, i) => {
    if (!l) return '';
    return `<text x="80" y="${narY + i * NAR_H}"
      font-family="Arial,Helvetica,sans-serif"
      font-size="20" fill="${C.muted}" opacity="${alpha}">${esc(l)}</text>`;
  }).join('\n');

  return Buffer.from(`<svg width="${panelW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="pgbg" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="${tagColor}" stop-opacity="0.13"/>
        <stop offset="55%"  stop-color="transparent" stop-opacity="0"/>
      </linearGradient>
    </defs>
    <rect width="${panelW}" height="${VH}" fill="${C.bgDark}" opacity="0.96"/>
    <rect width="${panelW}" height="${VH}" fill="url(#pgbg)"/>

    <!-- Numéro d'acte watermark -->
    <text x="18" y="${VH - 28}" font-family="Arial,Helvetica,sans-serif"
          font-size="270" font-weight="bold" fill="${C.white}" opacity="0.025">${actStr}</text>

    <!-- Barre décorative verticale -->
    <rect x="55" y="145" width="3" height="65" rx="2" fill="${tagColor}" opacity="${alpha * 0.8}"/>

    <!-- Badge tag -->
    <rect x="78" y="152" width="${tag.length * 10.8 + 36}" height="32" rx="16"
          fill="${tagColor}" opacity="${alpha * 0.9}"/>
    <text x="97" y="174" font-family="Arial,Helvetica,sans-serif"
          font-size="11" font-weight="bold" fill="${C.phone}"
          opacity="${alpha}">${esc(tag)}</text>

    <!-- Titre principal -->
    ${titleSVG}

    <!-- Ligne d'accent -->
    <rect x="80" y="${narY - 20}" width="125" height="3" rx="2"
          fill="${tagColor}" opacity="${alpha * 0.85}"/>

    <!-- Narration -->
    ${narSVG}

    <!-- Label scène (bas) -->
    <rect x="80" y="${VH - 70}" width="${Math.min(sceneLabel.length * 7.8 + 28, panelW - 110)}" height="30"
          rx="8" fill="${C.phone}" opacity="${alpha * 0.88}"/>
    <text x="96" y="${VH - 49}" font-family="Arial,Helvetica,sans-serif"
          font-size="13" fill="${C.muted}" opacity="${alpha}">${esc(sceneLabel.substring(0, 72))}</text>

    <!-- Barre de progression -->
    <rect x="80" y="${VH - 17}" width="${panelW - 180}" height="4" rx="2" fill="${C.bgMid}"/>
    <rect x="80" y="${VH - 17}" width="${progW}"          height="4" rx="2" fill="${tagColor}"/>
  </svg>`);
}

// ════════════════════════════════════════════════════════════════════════════
//  GÉNÉRATEURS SVG — SCÈNES ANIMÉES
// ════════════════════════════════════════════════════════════════════════════

/** Section 1 — Intro : carte Algérie + logo animé */
function introSVG(t) {
  const textA   = Math.min(1, ease.out(t * 2.5));
  const mapA    = Math.min(1, ease.out(Math.max(0, (t - 0.15) * 2)));
  const routeP  = Math.min(1, ease.out(Math.max(0, (t - 0.25) * 2.5)));
  const dotPulse = 0.72 + 0.28 * Math.sin(t * Math.PI * 5);

  // Villes principales Algérie (coordonnées relatives à la zone carte)
  const cities = [
    { n: 'Alger',      x: 720, y: 298, major: true  },
    { n: 'Oran',       x: 385, y: 296, major: true  },
    { n: 'Annaba',     x: 940, y: 258, major: true  },
    { n: 'Tlemcen',   x: 316, y: 332, major: true  },
    { n: 'Bejaia',    x: 790, y: 268, major: false },
    { n: 'Setif',     x: 820, y: 296, major: false },
    { n: 'Batna',     x: 870, y: 342, major: false },
    { n: 'Tiaret',    x: 540, y: 352, major: false },
    { n: 'Ghardaia',  x: 740, y: 538, major: false },
    { n: 'Ouargla',   x: 880, y: 580, major: false },
    { n: 'Tamanrasset', x: 730, y: 790, major: false },
    { n: 'Blida',     x: 700, y: 332, major: false },
  ];
  const routes = [[0,1],[0,2],[0,3],[0,4],[0,11],[1,7],[2,6],[4,5],[8,9],[8,10]];

  const routeSVG = routes.map(([a, b], i) => {
    const rp = Math.max(0, routeP - i * 0.07);
    const ca = cities[a], cb = cities[b];
    const ex = ca.x + (cb.x - ca.x) * rp;
    const ey = ca.y + (cb.y - ca.y) * rp;
    return `<line x1="${ca.x}" y1="${ca.y}" x2="${ex}" y2="${ey}"
              stroke="${C.accent}" stroke-width="2" opacity="${mapA * 0.65}"
              stroke-dasharray="9 5"/>`;
  }).join('\n');

  const dotsSVG = cities.map((c, i) => {
    const da = Math.min(1, Math.max(0, mapA - i * 0.055));
    const r = c.major ? 8 : 5;
    return `
      <circle cx="${c.x}" cy="${c.y}" r="${c.major ? 22 : 14}"
              fill="${C.accent}" opacity="${da * 0.14 * dotPulse}"/>
      <circle cx="${c.x}" cy="${c.y}" r="${r}"
              fill="${C.accent}" opacity="${da * 0.9}"/>
      <circle cx="${c.x}" cy="${c.y}" r="${r * 0.45}"
              fill="${C.white}" opacity="${da * 0.85}"/>`;
  }).join('\n');

  const pillA = Math.min(1, Math.max(0, (t - 0.55) * 4));

  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
      <radialGradient id="glow" cx="60%" cy="46%" r="42%">
        <stop offset="0%" stop-color="${C.accent}" stop-opacity="0.14"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#glow)"/>
    <rect x="0" y="0" width="${VW}" height="6" fill="${C.gold}" opacity="0.75"/>

    <!-- Carte Algérie stylisée (gauche) -->
    <g transform="translate(55, 60) scale(0.88)" opacity="${mapA}">
      <!-- Contour simplifié Algérie -->
      <path d="M 100,50 L 680,45 L 760,95 L 800,230 L 775,390 L 720,560 L 670,760 L 580,880 L 390,900 L 210,820 L 80,660 L 35,460 L 40,260 L 70,140 Z"
            fill="none" stroke="${C.primary}" stroke-width="2.5"
            opacity="${mapA * 0.4}" stroke-dasharray="14 7"/>
      ${routeSVG}
      ${dotsSVG}
    </g>

    <!-- Titre principal -->
    <text x="${VW - 580}" y="${VH / 2 - 95}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="94" font-weight="bold"
          fill="${C.white}" opacity="${textA}">RohWinBghit</text>

    <!-- Sous-titre arabe -->
    <text x="${VW - 580}" y="${VH / 2 - 12}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="40"
          fill="${C.gold}" opacity="${textA}">&#x631;&#x648;&#x62D; &#x648;&#x64A;&#x646; &#x628;&#x63A;&#x64A;&#x62A;</text>

    <!-- Slogan -->
    <text x="${VW - 580}" y="${VH / 2 + 58}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="21"
          fill="${C.accent}" opacity="${textA}">Plateforme intelligente de covoiturage inter-wilayas</text>

    <!-- Séparateur -->
    <rect x="${VW - 790}" y="${VH / 2 + 90}" width="420" height="2" rx="1"
          fill="${C.gold}" opacity="${textA * 0.55}"/>

    <!-- Pills fonctionnalités -->
    <rect x="${VW - 820}" y="${VH / 2 + 115}" width="190" height="34" rx="17"
          fill="${C.primary}" opacity="${pillA * 0.88}"/>
    <text x="${VW - 725}" y="${VH / 2 + 138}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="13"
          fill="${C.white}" opacity="${pillA}">KYC Biometrique</text>

    <rect x="${VW - 615}" y="${VH / 2 + 115}" width="190" height="34" rx="17"
          fill="${C.primary}" opacity="${pillA * 0.88}"/>
    <text x="${VW - 520}" y="${VH / 2 + 138}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="13"
          fill="${C.white}" opacity="${pillA}">CIB / Edahabia</text>

    <rect x="${VW - 410}" y="${VH / 2 + 115}" width="160" height="34" rx="17"
          fill="${C.primary}" opacity="${pillA * 0.88}"/>
    <text x="${VW - 330}" y="${VH / 2 + 138}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="13"
          fill="${C.white}" opacity="${pillA}">69 Wilayas</text>

    <!-- Auteurs -->
    <text x="${VW - 580}" y="${VH - 72}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="16"
          fill="${C.muted}" opacity="${textA}">
      AHMED BACHA Djamel Eddine  x  BELHORMA Sidi Mohammed Reduane</text>
    <text x="${VW - 580}" y="${VH - 46}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="14"
          fill="${C.primary}" opacity="${textA}">
      Encadrante : Mme BENLEDGHEM Rafika  x  2025/2026</text>

    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.gold}" opacity="0.4"/>
  </svg>`;
}

/** Section 2 — Problématique avec cards animées */
function problemSVG(t) {
  const titleA = Math.min(1, ease.out(t * 5));
  const problems = [
    { stat: '47M+',  unit: 'Algeriens',      label: 'Mobilite precaire',    color: C.red    },
    { stat: '3-4',   unit: 'places vides',   label: 'Sièges perdus/trajet', color: C.gold   },
    { stat: '2x',    unit: 'prix taxi',      label: 'Cout eleve mobilite',  color: C.teal   },
    { stat: '0',     unit: 'verification',   label: 'Insecurite trajets',   color: C.accent },
  ];
  const cW = 375, cH = 265;
  const totalW = problems.length * cW + 3 * 38;
  const startX = (VW - totalW) / 2;
  const cY = (VH - cH) / 2 + 45;

  const cards = problems.map((p, i) => {
    const ca = Math.min(1, ease.out(Math.max(0, (t - i * 0.14) * 3)));
    const slideY = lerp(45, 0, ca);
    const cx = startX + i * (cW + 38);
    return `
      <g transform="translate(0, ${slideY})" opacity="${ca}">
        <rect x="${cx}" y="${cY}" width="${cW}" height="${cH}" rx="20"
              fill="${C.bgMid}" opacity="0.88" stroke="${p.color}" stroke-width="1.5"/>
        <rect x="${cx}" y="${cY}" width="${cW}" height="6" rx="3" fill="${p.color}"/>
        <text x="${cx + cW / 2}" y="${cY + 90}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="74" font-weight="bold"
              fill="${p.color}">${esc(p.stat)}</text>
        <text x="${cx + cW / 2}" y="${cY + 130}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="18" fill="${C.muted}">${esc(p.unit)}</text>
        <rect x="${cx + 40}" y="${cY + 152}" width="${cW - 80}" height="2" fill="${p.color}" opacity="0.45"/>
        <text x="${cx + cW / 2}" y="${cY + 195}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="16" font-weight="bold"
              fill="${C.white}">${esc(p.label)}</text>
      </g>`;
  }).join('');

  const ctaA = Math.min(1, Math.max(0, (t - 0.72) * 5));
  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect x="0" y="0" width="${VW}" height="5" fill="${C.red}" opacity="0.65"/>

    <rect x="${VW / 2 - 145}" y="42" width="290" height="34" rx="17"
          fill="${C.red}" opacity="${titleA * 0.88}"/>
    <text x="${VW / 2}" y="65" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="12" font-weight="bold"
          fill="${C.white}" opacity="${titleA}">PROBLEMATIQUE</text>

    <text x="${VW / 2}" y="136" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="46" font-weight="bold"
          fill="${C.white}" opacity="${titleA}">Un besoin croissant de mobilite securisee</text>

    ${cards}

    <text x="${VW / 2}" y="${cY + cH + 70}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="22"
          fill="${C.accent}" opacity="${ctaA}">
      RohWinBghit repond a ces defis avec la technologie.</text>

    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.red}" opacity="0.28"/>
  </svg>`;
}

/** Section 9 — Dashboard Admin SVG (avec mouvement continu) */
function adminDashSVG(t) {
  const alpha = Math.min(1, ease.out(t * 3.5));

  // ── Mouvement continu tout au long de la scène ───────────────────────────
  const livePulse  = 0.55 + 0.45 * Math.abs(Math.sin(t * Math.PI * 7));
  const activeRow  = Math.floor(t * 14) % 5;          // rangée surlignée cyclique
  const scanBarW   = Math.round((0.35 + 0.65 * Math.abs(Math.sin(t * Math.PI * 3.5))) * 120);
  const clockSec   = Math.floor(t * 60);              // horloge factice
  const clockStr   = `12:${String(clockSec % 60).padStart(2,'0')}`;
  // ────────────────────────────────────────────────────────────────────────

  const statsData = [
    { label: 'Utilisateurs',    val: '1 247',  color: C.accent },
    { label: 'Trajets actifs',  val: '89',     color: C.gold   },
    { label: 'KYC en attente',  val: '12',     color: C.teal   },
    { label: 'Revenus (DA)',    val: '84 520', color: C.cyan   },
  ];
  const sW = 340, sH = 118, sY = 135;
  const sStartX = 95;

  const statCards = statsData.map((s, i) => {
    const sa = Math.min(1, ease.out(Math.max(0, (t - i * 0.09) * 4)));
    const sx = sStartX + i * (sW + 28);
    // Barre de progression animée dans chaque carte
    const barFill = Math.round((0.4 + 0.6 * Math.abs(Math.sin(t * Math.PI * 2.5 + i * 1.1))) * (sW - 32));
    return `
      <rect x="${sx}" y="${sY}" width="${sW}" height="${sH}" rx="14"
            fill="${C.bgDark}" opacity="${sa * 0.92}" stroke="${s.color}" stroke-width="1.2"/>
      <rect x="${sx}" y="${sY}" width="${sW}" height="5" rx="3" fill="${s.color}" opacity="${sa}"/>
      <text x="${sx + 20}" y="${sY + 45}" font-family="Arial,Helvetica,sans-serif"
            font-size="13" fill="${C.muted}" opacity="${sa}">${esc(s.label)}</text>
      <text x="${sx + 20}" y="${sY + 90}" font-family="Arial,Helvetica,sans-serif"
            font-size="36" font-weight="bold" fill="${s.color}" opacity="${sa}">${esc(s.val)}</text>
      <!-- mini barre animée -->
      <rect x="${sx + 16}" y="${sY + 104}" width="${sW - 32}" height="4" rx="2"
            fill="${C.bgMid}" opacity="${sa * 0.6}"/>
      <rect x="${sx + 16}" y="${sY + 104}" width="${barFill}" height="4" rx="2"
            fill="${s.color}" opacity="${sa * 0.75}"/>`;
  }).join('');

  const rows = [
    { id: 'USR-001', name: 'Ahmed Bacha D.',     role: 'Chauffeur', kyc: 'VALIDE',      c: C.accent },
    { id: 'USR-002', name: 'Fatima Z. Mansouri', role: 'Passager',  kyc: 'VALIDE',      c: C.accent },
    { id: 'USR-003', name: 'Karim Benali',       role: 'Chauffeur', kyc: 'EN ATTENTE',  c: C.gold   },
    { id: 'USR-004', name: 'Amina Belkaid',      role: 'Passager',  kyc: 'VALIDE',      c: C.accent },
    { id: 'USR-005', name: 'Youcef Hadj M.',     role: 'Chauffeur', kyc: 'VALIDE',      c: C.accent },
  ];

  const tableY = sY + sH + 58;
  const rowH = 52;

  const tableRows = rows.map((r, i) => {
    const ra = Math.min(1, ease.out(Math.max(0, (t - 0.22 - i * 0.07) * 4)));
    const ry = tableY + rowH + i * rowH + 4;
    const isActive = i === activeRow;   // surlignage cyclique
    return `
      <!-- Surlignage rangée active -->
      ${isActive ? `<rect x="95" y="${ry - 8}" width="${VW - 270}" height="${rowH - 4}" rx="8"
            fill="${C.accent}" opacity="${ra * 0.12}"/>` : ''}
      <rect x="95" y="${ry - 8}" width="${VW - 270}" height="${rowH - 4}" rx="8"
            fill="${C.bgDark}" opacity="${ra * (isActive ? 0.55 : 0.7)}"/>
      <text x="125" y="${ry + 26}" font-family="Arial,Helvetica,sans-serif"
            font-size="14" fill="${C.muted}" opacity="${ra}">${esc(r.id)}</text>
      <text x="270" y="${ry + 26}" font-family="Arial,Helvetica,sans-serif"
            font-size="14" fill="${isActive ? C.white : C.white}" opacity="${ra}">${esc(r.name)}</text>
      <text x="640" y="${ry + 26}" font-family="Arial,Helvetica,sans-serif"
            font-size="14" fill="${C.muted}" opacity="${ra}">${esc(r.role)}</text>
      <rect x="900" y="${ry + 9}" width="115" height="24" rx="12" fill="${r.c}" opacity="${ra * 0.82}"/>
      <text x="957" y="${ry + 26}" text-anchor="middle"
            font-family="Arial,Helvetica,sans-serif" font-size="11" font-weight="bold"
            fill="${C.phone}" opacity="${ra}">${esc(r.kyc)}</text>`;
  }).join('');

  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <!-- Sidebar -->
    <rect x="0" y="0" width="68" height="${VH}" fill="${C.bgDark}" opacity="0.95"/>
    <rect x="18" y="38" width="28" height="28" rx="7" fill="${C.accent}" opacity="${alpha * 0.9}"/>
    <rect x="16" y="98"  width="36" height="4" rx="2" fill="${C.primary}" opacity="${alpha * 0.7}"/>
    <rect x="16" y="118" width="36" height="4" rx="2" fill="${C.primary}" opacity="${alpha * 0.5}"/>
    <rect x="16" y="138" width="36" height="4" rx="2" fill="${C.primary}" opacity="${alpha * 0.4}"/>
    <rect x="0" y="0" width="${VW}" height="5" fill="${C.primary}" opacity="0.7"/>
    <!-- Header -->
    <rect x="68" y="0" width="${VW - 68}" height="78" fill="${C.bgMid}" opacity="${alpha * 0.5}"/>
    <text x="108" y="50" font-family="Arial,Helvetica,sans-serif"
          font-size="24" font-weight="bold" fill="${C.white}" opacity="${alpha}">Interface d'Administration</text>
    <!-- Horloge en direct -->
    <text x="${VW - 340}" y="48" font-family="Arial,Helvetica,sans-serif"
          font-size="14" fill="${C.muted}" opacity="${alpha}">Supervision &amp; Controle  |  ${esc(clockStr)}</text>
    <!-- Badge LIVE animé -->
    <circle cx="${VW - 140}" cy="42" r="7" fill="#FF4455" opacity="${alpha * livePulse}"/>
    <text x="${VW - 127}" y="48" font-family="Arial,Helvetica,sans-serif"
          font-size="13" font-weight="bold" fill="#FF4455" opacity="${alpha * livePulse}">LIVE</text>
    <!-- Stats -->
    ${statCards}
    <!-- Barre de scan horizontale animée sur le tableau -->
    <rect x="95" y="${tableY + rowH + activeRow * rowH}" width="${scanBarW * 12}" height="2"
          fill="${C.accent}" opacity="${alpha * 0.35}"/>
    <!-- Table header -->
    <text x="95" y="${tableY - 18}" font-family="Arial,Helvetica,sans-serif"
          font-size="18" font-weight="bold" fill="${C.white}" opacity="${alpha}">Gestion Utilisateurs</text>
    <rect x="95" y="${tableY - 10}" width="${VW - 270}" height="${rowH}" rx="10"
          fill="${C.bgMid}" opacity="${alpha * 0.55}"/>
    <text x="125" y="${tableY + 27}" font-family="Arial,Helvetica,sans-serif"
          font-size="13" font-weight="bold" fill="${C.muted}" opacity="${alpha}">ID</text>
    <text x="270" y="${tableY + 27}" font-family="Arial,Helvetica,sans-serif"
          font-size="13" font-weight="bold" fill="${C.muted}" opacity="${alpha}">Nom</text>
    <text x="640" y="${tableY + 27}" font-family="Arial,Helvetica,sans-serif"
          font-size="13" font-weight="bold" fill="${C.muted}" opacity="${alpha}">Role</text>
    <text x="957" y="${tableY + 27}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="13" font-weight="bold"
          fill="${C.muted}" opacity="${alpha}">KYC</text>
    ${tableRows}
    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.primary}" opacity="0.28"/>
  </svg>`;
}

/** Section 10 — Architecture technique animée */
function archSVG(t) {
  const titleA = Math.min(1, ease.out(t * 6));
  const layers = [
    { name: 'Flutter  (Mobile App)',    sub: 'Dart  x  Riverpod  x  Mapbox SDK',     color: C.cyan,   icon: 'F'  },
    { name: 'NestJS   (API Gateway)',   sub: 'TypeScript  x  REST  x  WebSocket',    color: C.accent, icon: 'N'  },
    { name: 'FastAPI  (Service IA)',    sub: 'Python  x  ArcFace  x  OCR  x  ONNX',  color: C.gold,   icon: 'AI' },
    { name: 'PostgreSQL  +  Redis',     sub: 'Donnees relationnelles  +  Cache',      color: C.teal,   icon: 'DB' },
    { name: 'SATIM / CIB Gateway',     sub: 'Paiement local securise',               color: C.purple, icon: '$'  },
  ];
  const boxW = 580, boxH = 76, boxX = (VW - boxW) / 2;
  const spacing = 125;
  const startY = 148;

  // Particule de données qui descend en boucle sur chaque connecteur
  const packetCycle = (t * 2.8) % 1.0;

  const layers_svg = layers.map((l, i) => {
    const la = Math.min(1, ease.out(Math.max(0, (t - i * 0.11) * 5)));
    const slideX = lerp(-90, 0, la);
    const by = startY + i * spacing;
    const arA = Math.min(1, Math.max(0, (t - (i + 0.6) * 0.11) * 6));
    // Particule animée sur la ligne de connexion
    const lineStart = by + boxH + 4;
    const lineEnd   = by + spacing - 8;
    const lineLen   = lineEnd - lineStart;
    const pOffset   = ((packetCycle + i * 0.22) % 1.0);
    const packetY   = lineStart + lineLen * pOffset;
    const packetGlow = 0.55 + 0.45 * Math.sin(t * Math.PI * 6 + i * 0.8);
    return `
      <g transform="translate(${slideX}, 0)" opacity="${la}">
        <rect x="${boxX}" y="${by}" width="${boxW}" height="${boxH}" rx="16"
              fill="${C.bgMid}" stroke="${l.color}" stroke-width="2"/>
        <rect x="${boxX}" y="${by}" width="66" height="${boxH}" rx="16"
              fill="${l.color}" opacity="0.82"/>
        <rect x="${boxX + 46}" y="${by}" width="20" height="${boxH}"
              fill="${l.color}" opacity="0.82"/>
        <text x="${boxX + 33}" y="${by + 47}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="14" font-weight="bold"
              fill="${C.white}">${esc(l.icon)}</text>
        <text x="${boxX + 90}" y="${by + 30}" font-family="Arial,Helvetica,sans-serif"
              font-size="20" font-weight="bold" fill="${C.white}">${esc(l.name)}</text>
        <text x="${boxX + 90}" y="${by + 57}" font-family="Arial,Helvetica,sans-serif"
              font-size="13" fill="${C.muted}">${esc(l.sub)}</text>
      </g>
      ${i < layers.length - 1 ? `
      <g opacity="${arA}">
        <line x1="${VW / 2}" y1="${lineStart}" x2="${VW / 2}" y2="${lineEnd}"
              stroke="${l.color}" stroke-width="2" opacity="0.6" stroke-dasharray="8 5"/>
        <polygon points="${VW / 2 - 7},${lineEnd - 2} ${VW / 2 + 7},${lineEnd - 2} ${VW / 2},${lineEnd + 10}"
                 fill="${l.color}" opacity="0.65"/>
        <!-- Particule de données mobile -->
        <circle cx="${VW / 2}" cy="${packetY}" r="6"
                fill="${l.color}" opacity="${arA * packetGlow}"/>
        <circle cx="${VW / 2}" cy="${packetY}" r="11"
                fill="${l.color}" opacity="${arA * packetGlow * 0.28}"/>
      </g>` : ''}`;
  }).join('');

  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect x="0" y="0" width="${VW}" height="5" fill="${C.cyan}" opacity="0.55"/>

    <rect x="${VW / 2 - 140}" y="42" width="280" height="34" rx="17"
          fill="${C.cyan}" opacity="${titleA * 0.18}"/>
    <text x="${VW / 2}" y="66" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="12" font-weight="bold"
          fill="${C.cyan}" opacity="${titleA}">ARCHITECTURE TECHNIQUE</text>
    <text x="${VW / 2}" y="118" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="40" font-weight="bold"
          fill="${C.white}" opacity="${titleA}">Stack Technologique</text>

    ${layers_svg}

    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.cyan}" opacity="0.28"/>
  </svg>`;
}

/** Section 11 — KPI avec counter-up animé */
function kpiSVG(t) {
  const titleA = Math.min(1, ease.out(t * 5));
  const kpis = [
    { label: 'Precision KYC',  finalVal: 94,   unit: '%',    note: 'ArcFace + Anti-spoofing', color: C.accent },
    { label: 'Disponibilite',  finalVal: 99.1, unit: '%',    note: 'API + Circuit Breaker',   color: C.gold   },
    { label: 'Score SUS',      finalVal: 71.6, unit: '/100', note: 'ISO 9241 — Bon',          color: C.teal   },
    { label: 'Wilayas',        finalVal: 69,   unit: '',     note: 'Couverture nationale',     color: C.cyan   },
  ];
  const cW = 380, cH = 298;
  const totalW = kpis.length * cW + 3 * 38;
  const startX = (VW - totalW) / 2;
  const cY = (VH - cH) / 2 + 25;
  const pulse = 0.8 + 0.2 * Math.sin(t * Math.PI * 4);

  const cards = kpis.map((k, i) => {
    const ca = Math.min(1, ease.out(Math.max(0, (t - i * 0.11) * 3.5)));
    const cv = Math.min(1, ease.out(Math.max(0, (t - i * 0.11 - 0.08) * 2.2)));
    const cur = k.finalVal * cv;
    const disp = Number.isInteger(k.finalVal) ? Math.round(cur) : cur.toFixed(1);
    const cx = startX + i * (cW + 38);
    return `
      <g opacity="${ca}">
        <rect x="${cx - 8}" y="${cY - 8}" width="${cW + 16}" height="${cH + 16}" rx="28"
              fill="none" stroke="${k.color}" stroke-width="1" opacity="${0.22 * pulse}"/>
        <rect x="${cx}" y="${cY}" width="${cW}" height="${cH}" rx="22"
              fill="${C.bgMid}" opacity="0.92" stroke="${k.color}" stroke-width="2"/>
        <rect x="${cx}" y="${cY}" width="${cW}" height="7" rx="4" fill="${k.color}"/>
        <text x="${cx + cW / 2}" y="${cY + 108}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="86" font-weight="bold"
              fill="${k.color}">${esc(String(disp))}</text>
        <text x="${cx + cW / 2}" y="${cY + 148}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="24" fill="${C.white}">${esc(k.unit)}</text>
        <rect x="${cx + 40}" y="${cY + 172}" width="${cW - 80}" height="2" fill="${k.color}" opacity="0.48"/>
        <text x="${cx + cW / 2}" y="${cY + 212}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="18" font-weight="bold"
              fill="${C.white}">${esc(k.label)}</text>
        <text x="${cx + cW / 2}" y="${cY + 248}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="14" fill="${C.muted}">${esc(k.note)}</text>
      </g>`;
  }).join('');

  const hypsA = Math.min(1, Math.max(0, (t - 0.68) * 5));
  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect x="0" y="0" width="${VW}" height="5" fill="${C.gold}" opacity="0.72"/>

    <rect x="${VW / 2 - 145}" y="42" width="290" height="34" rx="17"
          fill="${C.gold}" opacity="${titleA * 0.88}"/>
    <text x="${VW / 2}" y="66" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="12" font-weight="bold"
          fill="${C.phone}" opacity="${titleA}">RESULTATS VALIDES</text>
    <text x="${VW / 2}" y="126" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="44" font-weight="bold"
          fill="${C.white}" opacity="${titleA}">Validation Experimentale</text>

    ${cards}

    <text x="${VW * 0.18}" y="${cY + cH + 68}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="17" font-weight="bold"
          fill="${C.accent}" opacity="${hypsA}">H1 : KYC VALIDE</text>
    <text x="${VW * 0.50}" y="${cY + cH + 68}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="17" font-weight="bold"
          fill="${C.gold}" opacity="${hypsA}">H2 : Usabilite VALIDE</text>
    <text x="${VW * 0.82}" y="${cY + cH + 68}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="17" font-weight="bold"
          fill="${C.teal}" opacity="${hypsA}">H3 : Resilience VALIDE</text>

    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.gold}" opacity="0.4"/>
  </svg>`;
}

/** Section 12 — Outro cinématique avec cercles pulsants */
function conclusionSVG(t) {
  const alpha   = Math.min(1, ease.out(t * 2.8));
  const pillA   = Math.min(1, ease.out(Math.max(0, (t - 0.28) * 3.5)));
  const msgA    = Math.min(1, Math.max(0, (t - 0.58) * 5));
  const ps      = 1 + 0.06 * Math.sin(t * Math.PI * 5);

  const pillars = ['Simple', 'Sur', 'Abordable'];
  const pillarColors = [C.accent, C.gold, C.teal];

  const pillarsCards = pillars.map((p, i) => {
    const pa = Math.min(1, ease.out(Math.max(0, (t - 0.28 - i * 0.11) * 3.5)));
    const px = VW / 2 - 280 + i * 280;
    return `
      <g opacity="${pa}">
        <rect x="${px - 88}" y="${VH / 2 + 88}" width="176" height="62" rx="31"
              fill="${pillarColors[i]}" opacity="0.9"/>
        <text x="${px}" y="${VH / 2 + 128}" text-anchor="middle"
              font-family="Arial,Helvetica,sans-serif" font-size="22" font-weight="bold"
              fill="${C.phone}">${esc(p)}</text>
      </g>`;
  }).join('');

  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
      <radialGradient id="hero" cx="50%" cy="44%" r="46%">
        <stop offset="0%" stop-color="${C.accent}" stop-opacity="0.20"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#hero)"/>
    <rect x="0" y="0" width="${VW}" height="6" fill="${C.gold}" opacity="${alpha * 0.82}"/>

    <!-- Anneaux pulsants -->
    <circle cx="${VW / 2}" cy="${VH / 2 - 22}" r="${255 * ps}"
            fill="none" stroke="${C.accent}" stroke-width="1" opacity="${alpha * 0.07}"/>
    <circle cx="${VW / 2}" cy="${VH / 2 - 22}" r="${175 * ps}"
            fill="none" stroke="${C.accent}" stroke-width="1" opacity="${alpha * 0.11}"/>
    <circle cx="${VW / 2}" cy="${VH / 2 - 22}" r="${105 * ps}"
            fill="none" stroke="${C.accent}" stroke-width="2" opacity="${alpha * 0.18}"/>

    <!-- Nom application -->
    <text x="${VW / 2}" y="${VH / 2 - 42}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="98" font-weight="bold"
          fill="${C.white}" opacity="${alpha}">RohWinBghit</text>

    <!-- Arabe -->
    <text x="${VW / 2}" y="${VH / 2 + 26}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="38"
          fill="${C.gold}" opacity="${alpha}">&#x631;&#x648;&#x62D; &#x648;&#x64A;&#x646; &#x628;&#x63A;&#x64A;&#x62A;</text>

    <!-- Pilliers -->
    ${pillarsCards}

    <!-- Message final -->
    <text x="${VW / 2}" y="${VH / 2 + 210}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="20"
          fill="${C.muted}" opacity="${msgA}">
      Contribue a rendre la mobilite inter-wilayas plus sure et accessible pour tous.</text>

    <!-- Auteurs -->
    <text x="${VW / 2}" y="${VH - 72}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="16"
          fill="${C.muted}" opacity="${alpha}">
      AHMED BACHA Djamel Eddine  x  BELHORMA Sidi Mohammed Reduane</text>
    <text x="${VW / 2}" y="${VH - 46}" text-anchor="middle"
          font-family="Arial,Helvetica,sans-serif" font-size="14"
          fill="${C.primary}" opacity="${alpha}">
      Universite Abou Bekr Belkaid – Tlemcen  x  Master 2 GL  x  2025/2026</text>

    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.gold}" opacity="${alpha * 0.5}"/>
  </svg>`;
}

// ════════════════════════════════════════════════════════════════════════════
//  OVERLAYS SPÉCIAUX
// ════════════════════════════════════════════════════════════════════════════

/** Ligne de scan KYC animée (remonte de bas en haut en boucle) */
function kycScanOverlay(t) {
  const scanY = Math.round(((t * 2.2) % 1.0) * PH_H);
  return Buffer.from(`<svg width="${PH_W}" height="${PH_H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="sg" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="${C.cyan}" stop-opacity="0"/>
        <stop offset="100%" stop-color="${C.cyan}" stop-opacity="0.55"/>
      </linearGradient>
    </defs>
    <rect x="0" y="${Math.max(0, scanY - 46)}" width="${PH_W}" height="${Math.min(49, scanY)}"
          fill="url(#sg)"/>
    <rect x="0" y="${scanY}" width="${PH_W}" height="3" fill="${C.cyan}" opacity="0.92"/>
    <!-- Marqueurs de coin -->
    <rect x="28" y="28" width="44" height="4" fill="${C.cyan}"/>
    <rect x="28" y="28" width="4" height="44" fill="${C.cyan}"/>
    <rect x="${PH_W - 72}" y="28" width="44" height="4" fill="${C.cyan}"/>
    <rect x="${PH_W - 32}" y="28" width="4" height="44" fill="${C.cyan}"/>
    <rect x="28" y="${PH_H - 32}" width="44" height="4" fill="${C.cyan}"/>
    <rect x="28" y="${PH_H - 72}" width="4" height="44" fill="${C.cyan}"/>
    <rect x="${PH_W - 72}" y="${PH_H - 32}" width="44" height="4" fill="${C.cyan}"/>
    <rect x="${PH_W - 32}" y="${PH_H - 72}" width="4" height="44" fill="${C.cyan}"/>
  </svg>`);
}

/** Halo QR code pulsant */
function qrGlowOverlay(t) {
  const pulse = 0.45 + 0.55 * Math.abs(Math.sin(t * Math.PI * 2.8));
  return Buffer.from(`<svg width="${PH_W}" height="${PH_H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <radialGradient id="qg" cx="50%" cy="44%" r="36%">
        <stop offset="0%"   stop-color="${C.gold}" stop-opacity="${pulse * 0.42}"/>
        <stop offset="100%" stop-color="${C.gold}" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${PH_W}" height="${PH_H}" fill="url(#qg)"/>
    <rect x="${PH_W * 0.18}" y="${PH_H * 0.18}" width="${PH_W * 0.64}" height="${PH_H * 0.36}"
          rx="12" fill="none" stroke="${C.gold}"
          stroke-width="${1.5 + pulse * 2.5}" opacity="${0.28 + pulse * 0.52}"/>
  </svg>`);
}

// ════════════════════════════════════════════════════════════════════════════
//  COMPOSITEURS DE FRAMES
// ════════════════════════════════════════════════════════════════════════════

/** Frame SVG animée (sections entièrement générées) */
async function composeSVGFrame(svgStr, frameIdx) {
  const out = path.join(TEMP_DIR, `f${String(frameIdx).padStart(6, '0')}.png`);
  if (fs.existsSync(out)) return out;
  await sharp({ create: { width: VW, height: VH, channels: 4,
    background: { r: 0, g: 0, b: 0, alpha: 255 }}})
    .composite([{ input: Buffer.from(svgStr), top: 0, left: 0 }])
    .png({ compressionLevel: 3 })
    .toFile(out);
  return out;
}

/** Frame téléphone avec Ken Burns zoom + effets spéciaux */
async function composePhoneFrame(opts, frameIdx) {
  const { act, scene, imgBuf, t, actProgress } = opts;
  const out = path.join(TEMP_DIR, `f${String(frameIdx).padStart(6, '0')}.png`);
  if (fs.existsSync(out)) return out;

  const alpha = fadeAlpha(t, 0.18, 0.15);

  // Ken Burns : zoom progressif 1.0→1.07
  let kbBuf = imgBuf;
  try {
    const zoom = lerp(1.0, 1.07, ease.inOut(t));
    const srcW = Math.round(PH_W / zoom);
    const srcH = Math.round(PH_H / zoom);
    const srcX = Math.max(0, Math.round((PH_W - srcW) / 2));
    const srcY = Math.max(0, Math.round((PH_H - srcH) / 2));
    kbBuf = await sharp(imgBuf)
      .extract({ left: srcX, top: srcY, width: srcW, height: srcH })
      .resize(PH_W, PH_H, { fit: 'fill' })
      .png().toBuffer();
  } catch (_) { /* fallback: imgBuf non modifié */ }

  const bg        = backgroundSVG(act.tagColor);
  const leftPanel = leftPanelSVG({
    tag:        act.tag,
    tagColor:   act.tagColor,
    title:      act.title,
    narration:  act.narration,
    sceneLabel: scene.label,
    actNum:     act.act,
    progress:   actProgress / SCENARIO.length,
    alpha,
  });
  const bezel = phoneBezelSVG(act.tagColor, 0.16 * alpha);

  const composites = [
    { input: bg,        top: 0,        left: 0        },
    { input: leftPanel, top: 0,        left: 0        },
    { input: bezel,     top: PH_TOP,   left: PH_LEFT  },
    { input: kbBuf,     top: PH_SCR_T, left: PH_SCR_L },
  ];

  if (scene.effect === 'kyc_scan') {
    composites.push({ input: kycScanOverlay(t), top: PH_SCR_T, left: PH_SCR_L });
  }
  if (scene.effect === 'qr_glow') {
    composites.push({ input: qrGlowOverlay(t), top: PH_SCR_T, left: PH_SCR_L });
  }

  await sharp({ create: { width: VW, height: VH, channels: 3, background: rgb(C.bgDark) }})
    .composite(composites)
    .png({ compressionLevel: 3 })
    .toFile(out);
  return out;
}

/** Frame plein écran avec Ken Burns + overlay dégradé */
async function composeFullscreenFrame(opts, frameIdx) {
  const { act, scene, imgBuf, t } = opts;
  const out = path.join(TEMP_DIR, `f${String(frameIdx).padStart(6, '0')}.png`);
  if (fs.existsSync(out)) return out;

  const alpha = fadeAlpha(t, 0.15, 0.15);

  let fsBuf = imgBuf;
  try {
    // Zoom 1.0→1.10 : mouvement bien visible sur les 10s GPS
    const zoom = lerp(1.0, 1.10, ease.inOut(t));
    const meta = await sharp(imgBuf).metadata();
    const srcW = Math.round(meta.width / zoom);
    const srcH = Math.round(meta.height / zoom);
    const srcX = Math.max(0, Math.round((meta.width - srcW) / 2));
    const srcY = Math.max(0, Math.round((meta.height - srcH) / 2));
    fsBuf = await sharp(imgBuf)
      .extract({ left: srcX, top: srcY, width: srcW, height: srcH })
      .resize(VW, VH, { fit: 'fill' })
      .png().toBuffer();
  } catch (_) { /* fallback */ }

  const overlay = Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="ov" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="${C.bgDark}" stop-opacity="0.58"/>
        <stop offset="32%"  stop-color="transparent" stop-opacity="0"/>
        <stop offset="68%"  stop-color="transparent" stop-opacity="0"/>
        <stop offset="100%" stop-color="${C.bgDark}" stop-opacity="0.76"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#ov)"/>
    <rect x="0" y="0" width="${VW}" height="5" fill="${act.tagColor}" opacity="${alpha * 0.7}"/>
    <rect x="60" y="30" width="${act.tag.length * 11 + 36}" height="32" rx="16"
          fill="${act.tagColor}" opacity="${alpha * 0.9}"/>
    <text x="76" y="52" font-family="Arial,Helvetica,sans-serif"
          font-size="11" font-weight="bold" fill="${C.phone}" opacity="${alpha}">${esc(act.tag)}</text>
    <rect x="0" y="${VH - 7}" width="${VW}" height="7"
          fill="${act.tagColor}" opacity="${alpha * 0.28}"/>
  </svg>`);

  await sharp({ create: { width: VW, height: VH, channels: 3, background: rgb(C.bgDark) }})
    .composite([
      { input: fsBuf,   top: 0, left: 0 },
      { input: overlay, top: 0, left: 0 },
    ])
    .png({ compressionLevel: 3 })
    .toFile(out);
  return out;
}

/** Dispatch : choisit le bon compositeur selon le type de scène */
async function renderFrame(act, scene, imgBuf, t, frameIdx, actProgress) {
  if (scene.type) {
    let svg;
    switch (scene.type) {
      case 'intro_animated':      svg = introSVG(t);        break;
      case 'problem_animated':    svg = problemSVG(t);      break;
      case 'admin_animated':      svg = adminDashSVG(t);    break;
      case 'arch_animated':       svg = archSVG(t);         break;
      case 'kpi_animated':        svg = kpiSVG(t);          break;
      case 'conclusion_animated': svg = conclusionSVG(t);   break;
      default: throw new Error(`Type de scène inconnu: ${scene.type}`);
    }
    return composeSVGFrame(svg, frameIdx);
  }
  if (scene.fullscreen && imgBuf) return composeFullscreenFrame({ act, scene, imgBuf, t }, frameIdx);
  if (imgBuf)                     return composePhoneFrame({ act, scene, imgBuf, t, actProgress }, frameIdx);
  return null;
}

// ════════════════════════════════════════════════════════════════════════════
//  SCÉNARIO — 12 SECTIONS
// ════════════════════════════════════════════════════════════════════════════

const SCENARIO = [

  // ── S1 : INTRO (8 s) — carte Algérie + logo ─────────────────────────────
  {
    act: 1, tag: 'INTRODUCTION', tagColor: C.accent,
    title: 'RohWinBghit\n\"Votre trajet,\nnotre priorite.\"',
    narration: [],
    scenes: [
      { type: 'intro_animated', duration: 8, label: 'Introduction — RohWinBghit' },
    ],
  },

  // ── S1b : LANDING SCREENS (8 s) ──────────────────────────────────────────
  {
    act: 1, tag: 'DECOUVERTE', tagColor: C.accent,
    title: 'RohWinBghit\n\"Votre trajet,\nnotre priorite.\"',
    narration: ['69 Wilayas couvertes', 'Disponible 24h/24 et 7j/7', '100% Securise & Verifie'],
    scenes: [
      { file: 'Screen1.jpg', fullscreen: true, duration: 4,
        label: 'Ecran d\'accueil — Splash Screen' },
      { file: 'Screen2.jpg', fullscreen: true, duration: 4,
        label: 'Landing Page — Fonctionnalites cles' },
    ],
  },

  // ── S2 : PROBLÉMATIQUE (7 s) ─────────────────────────────────────────────
  {
    act: 2, tag: 'PROBLEMATIQUE', tagColor: C.red,
    title: 'Un besoin\ncroissant',
    narration: [],
    scenes: [
      { type: 'problem_animated', duration: 7,
        label: 'Problematique — Mobilite algerienne' },
    ],
  },

  // ── S3 : INSCRIPTION (10 s) ───────────────────────────────────────────────
  {
    act: 3, tag: 'INSCRIPTION', tagColor: C.primary,
    title: 'Inscription &\nChoix du Role',
    narration: [
      'Creation de compte rapide', '',
      'Passager — reserver des trajets',
      'Chauffeur — partager ses frais',
    ],
    scenes: [
      { file: 'inscription.jpg', duration: 10,
        label: 'Choix du role — Passager ou Chauffeur' },
    ],
  },

  // ── S4 : KYC BIOMÉTRIQUE (16 s) ──────────────────────────────────────────
  {
    act: 4, tag: 'KYC BIOMETRIQUE', tagColor: C.gold,
    title: 'Pipeline KYC\nBiometrique',
    narration: [
      'OCR : Carte Nationale d\'identite',
      'ArcFace : Correspondance faciale',
      'Liveness : Detection anti-spoofing', '',
      'Seuil conducteur >= 0,65',
    ],
    scenes: [
      { file: 'd_27_verif_intro.png',       duration: 4,
        label: '1. Introduction verification KYC' },
      { file: 'd_28_verif_id_capture.png',  duration: 4,
        label: '2. Capture CIN — Extraction OCR', effect: 'kyc_scan' },
      { file: 'd_29_verif_face_capture.png', duration: 4,
        label: '3. Capture faciale — Detection vivacite', effect: 'kyc_scan' },
      { file: 'd_30_verif_pending.png',      duration: 4,
        label: '4. Analyse IA en cours...' },
    ],
  },

  // ── S5 : CONDUCTEUR (15 s) ────────────────────────────────────────────────
  {
    act: 5, tag: 'COTE CONDUCTEUR', tagColor: C.primary,
    title: 'Publier\nun Trajet',
    narration: [
      'Tlemcen → Alger  :  ~516 km',
      'Duree estimee    :  5h33',
      'Prix recommande  :  automatique', '',
      'Gestion demandes en temps reel',
    ],
    scenes: [
      { file: 'Trajet publié.jpg',   duration: 5,
        label: 'Publication du trajet — Carte interactive' },
      { file: 'd_01_tab_dashboard.png', duration: 5,
        label: 'Tableau de bord conducteur' },
      { file: 'd_06_trip_manage.png',   duration: 5,
        label: 'Gestion des demandes de reservation' },
    ],
  },

  // ── S6 : PASSAGER (15 s) ─────────────────────────────────────────────────
  {
    act: 6, tag: 'COTE PASSAGER', tagColor: C.teal,
    title: 'Rechercher\n& Reserver',
    narration: [
      'Oran → Alger  :  2 trajets trouves',
      'Conducteurs verifies KYC', '',
      'Prix : 850 DA ou 1 000 DA',
      'Paiement CIB / Edahabia / Especes',
    ],
    scenes: [
      { file: 'Passager réserve.jpg',   duration: 5,
        label: 'Resultats de recherche — Trajets disponibles' },
      { file: 'p_08_search_results.png', duration: 5,
        label: 'Vue complete resultats de recherche' },
      { file: 'p_09_trip_details.png',   duration: 5,
        label: 'Detail du trajet — Bouton Reserver' },
    ],
  },

  // ── S7 : PAIEMENT & QR (12 s) ────────────────────────────────────────────
  {
    act: 7, tag: 'BILLET & QR CODE', tagColor: C.gold,
    title: 'Billet\n& QR Code',
    narration: [
      'Boarding Pass numerique genere', 'apres confirmation de reservation.', '',
      'QR Code unique par trajet :',
      'Scan depart = liberation paiement',
    ],
    scenes: [
      { file: 'QR scanné au départ.jpg', duration: 4,
        label: 'Boarding Pass — QR Code unique', effect: 'qr_glow' },
      { file: 'p_10_ticket.png',          duration: 4,
        label: 'Ticket passager avec QR Code', effect: 'qr_glow' },
      { file: 'd_11_wallet.png',          duration: 4,
        label: 'Wallet — Paiement SATIM / CIB / Edahabia' },
    ],
  },

  // ── S8 : GPS TEMPS RÉEL (10 s) ───────────────────────────────────────────
  {
    act: 8, tag: 'SUIVI GPS', tagColor: C.accent,
    title: 'Suivi GPS\nTemps Reel',
    narration: [
      'Lien partageable sans inscription', '',
      'WebSocket  x  Mapbox SDK',
      'La famille suit le trajet en live.',
    ],
    scenes: [
      { file: 'p_26_live_tracking.png', fullscreen: true, duration: 10,
        label: 'Suivi GPS temps reel — Carte interactive' },
    ],
  },

  // ── S9 : DASHBOARD ADMIN (10 s) ──────────────────────────────────────────
  {
    act: 9, tag: 'ADMINISTRATION', tagColor: C.primary,
    title: 'Interface\nAdmin',
    narration: [],
    scenes: [
      { type: 'admin_animated', duration: 10,
        label: 'Tableau de bord administration' },
    ],
  },

  // ── S10 : ARCHITECTURE TECHNIQUE (8 s) ───────────────────────────────────
  {
    act: 10, tag: 'ARCHITECTURE', tagColor: C.cyan,
    title: 'Stack\nTechnologique',
    narration: [],
    scenes: [
      { type: 'arch_animated', duration: 8,
        label: 'Architecture technique — Stack complet' },
    ],
  },

  // ── S11 : RÉSULTATS KPI (10 s) ───────────────────────────────────────────
  {
    act: 11, tag: 'RESULTATS', tagColor: C.gold,
    title: 'Validation\nExperimentale',
    narration: [],
    scenes: [
      { type: 'kpi_animated', duration: 10,
        label: 'Resultats experimentaux valides' },
    ],
  },

  // ── S12 : CONCLUSION (6 s) ────────────────────────────────────────────────
  {
    act: 12, tag: 'CONCLUSION', tagColor: C.accent,
    title: 'Simple • Sur\n• Abordable',
    narration: [],
    scenes: [
      { type: 'conclusion_animated', duration: 6,
        label: 'Conclusion — RohWinBghit' },
    ],
  },
];

// ════════════════════════════════════════════════════════════════════════════
//  TRAITEMENT PRINCIPAL
// ════════════════════════════════════════════════════════════════════════════

async function processAllScenes() {
  const limit = createLimiter(8);
  let frameIdx   = 0;
  let missing    = 0;
  const frameList = [];

  const totalFrames = SCENARIO.reduce((s, act) =>
    s + act.scenes.reduce((ss, sc) => ss + Math.round(sc.duration * FPS), 0), 0);
  console.log(`\n  Frames a generer : ~${totalFrames}  (${(totalFrames / FPS / 60).toFixed(1)} min)\n`);

  for (let ai = 0; ai < SCENARIO.length; ai++) {
    const act        = SCENARIO[ai];
    const actProgress = ai + 1;
    console.log(`  [S${String(act.act).padStart(2, '0')}] ${act.tag}`);

    for (const scene of act.scenes) {
      const nFrames = Math.round(scene.duration * FPS);

      // Pré-charger l'image (une fois par scène)
      let imgBuf = null;
      if (scene.file) {
        const imgPath = resolveImg(scene.file);
        if (!imgPath) {
          console.warn(`     MANQUANT: ${scene.file}`);
          missing++;
          continue;
        }
        try {
          const tW = scene.fullscreen ? VW  : PH_W;
          const tH = scene.fullscreen ? VH  : PH_H;
          imgBuf = await sharp(imgPath)
            .resize(tW, tH, { fit: 'cover', position: 'top' })
            .png()
            .toBuffer();
        } catch (e) {
          console.error(`     Erreur image : ${e.message}`);
          continue;
        }
      }

      process.stdout.write(`     └─ ${scene.label.substring(0, 54)}... `);

      const sceneStart = frameIdx;
      const tasks = Array.from({ length: nFrames }, (_, fi) => {
        const t   = nFrames > 1 ? fi / (nFrames - 1) : 0;
        const idx = sceneStart + fi;
        return limit(() => renderFrame(act, scene, imgBuf, t, idx, actProgress));
      });

      const paths = await Promise.all(tasks);
      for (const p of paths) {
        if (p) frameList.push({ path: p, duration: 1 / FPS });
      }

      frameIdx += nFrames;
      console.log(`${nFrames}f  OK`);
    }
  }

  if (missing > 0) console.warn(`\n  ${missing} image(s) manquante(s) ignoree(s)`);

  const dur = frameList.length / FPS;
  console.log(`\n  Total : ${frameList.length} frames (${dur.toFixed(0)}s = ${(dur / 60).toFixed(1)} min)`);
  return frameList;
}

// ─── CONCAT FILE & ASSEMBLAGE ────────────────────────────────────────────────

function writeConcatFile(frameList) {
  const lines = frameList.flatMap(f => [
    `file '${f.path.replace(/\\/g, '/')}'`,
    `duration ${f.duration.toFixed(6)}`,
  ]);
  // FFmpeg concat : répéter la dernière image pour fermer la séquence
  lines.push(`file '${frameList[frameList.length - 1].path.replace(/\\/g, '/')}'`);
  fs.writeFileSync(CONCAT_FILE, lines.join('\n'));
}

async function assembleVideo(frameList) {
  return new Promise((resolve, reject) => {
    writeConcatFile(frameList);
    const durSec = (frameList.length / FPS).toFixed(0);
    console.log(`\n  Assemblage MP4 (${durSec}s @ ${FPS}fps | H.264 High | crf 18)...`);

    ffmpeg()
      .input(CONCAT_FILE)
      .inputOptions(['-f', 'concat', '-safe', '0'])
      .videoFilters([
        `fps=${FPS}`,
        `scale=${VW}:${VH}:force_original_aspect_ratio=disable`,
      ])
      .outputOptions([
        '-c:v', 'libx264',
        '-preset', 'fast',
        '-crf', '18',
        '-pix_fmt', 'yuv420p',
        '-movflags', '+faststart',
        '-profile:v', 'high',
        '-level', '4.1',
      ])
      .output(OUTPUT_FILE)
      .on('start', cmd => console.log(`  FFmpeg: ${cmd.slice(0, 115)}...`))
      .on('progress', p => {
        if (p.percent != null)
          process.stdout.write(`\r  Encodage : ${Math.round(p.percent)}%   `);
      })
      .on('end', () => {
        process.stdout.write('\r');
        console.log('  Encodage termine.                    ');
        resolve();
      })
      .on('error', (err, _s, stderr) => {
        console.error('\n  Erreur FFmpeg :', err.message);
        if (stderr) console.error(stderr.slice(-800));
        reject(err);
      })
      .run();
  });
}

function cleanup() {
  try { fs.rmSync(TEMP_DIR, { recursive: true, force: true }); } catch (_) {}
}

// ─── POINT D'ENTRÉE ──────────────────────────────────────────────────────────
(async () => {
  const t0  = Date.now();
  const sep = '═'.repeat(65);

  console.log(sep);
  console.log('  RohWinBghit — Generateur Video Cinematique v3');
  console.log('  12 sections  x  Full HD 1920x1080  x  30 fps  x  H.264 High');
  console.log(sep);

  // Inventaire Screen_mobile
  const screenFiles = fs.existsSync(SCREEN_DIR)
    ? fs.readdirSync(SCREEN_DIR).filter(f => /\.(png|jpg|jpeg)$/i.test(f) && !f.startsWith('.'))
    : [];
  console.log(`\n  Captures Screen_mobile : ${screenFiles.length} fichiers`);
  screenFiles.forEach(f => console.log(`    - ${f}`));

  try {
    const frameList = await processAllScenes();
    if (!frameList.length) throw new Error('Aucune frame generee — verifiez les images.');

    await assembleVideo(frameList);

    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    const sizeMB  = (fs.statSync(OUTPUT_FILE).size / 1024 / 1024).toFixed(1);
    const dur     = (frameList.length / FPS);

    console.log('\n' + sep);
    console.log('  GENERATION TERMINEE');
    console.log(`  Fichier : ${OUTPUT_FILE}`);
    console.log(`  Duree   : ${Math.floor(dur / 60)}min ${Math.round(dur % 60)}s  (${frameList.length} frames)`);
    console.log(`  Taille  : ${sizeMB} MB`);
    console.log(`  Temps   : ${elapsed}s de traitement`);
    console.log(sep + '\n');

    cleanup();
  } catch (e) {
    console.error('\nErreur fatale :', e.message);
    if (e.stack) console.error(e.stack.split('\n').slice(1, 5).join('\n'));
    process.exit(1);
  }
})();
