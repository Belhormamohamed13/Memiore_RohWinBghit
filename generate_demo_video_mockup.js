#!/usr/bin/env node
/**
 * ═══════════════════════════════════════════════════════════════════════
 *   RohWinBghit — Générateur Vidéo Démo Mockup (Parcours Utilisateur)
 *   Soutenance Master 2 · Génie Logiciel · Tlemcen 2025-2026
 * ═══════════════════════════════════════════════════════════════════════
 */

'use strict';

const sharp      = require('sharp');
const ffmpeg     = require('fluent-ffmpeg');
const ffmpegPath = require('@ffmpeg-installer/ffmpeg').path;
const path       = require('path');
const fs         = require('fs');
const winston    = require('winston');
const cliProgress = require('cli-progress');

ffmpeg.setFfmpegPath(ffmpegPath);

// ═══════════════════════════════════════════════════════════════════════
//  CONFIGURATION GLOBALE
// ═══════════════════════════════════════════════════════════════════════

const CONFIG = {
  mode: process.argv.includes('--turbo') ? 'turbo'
      : process.argv.includes('--4k')    ? '4k'
      : 'normal',

  resolutions: {
    turbo:  { w: 1280, h: 720,  fps: 24, crf: 23, preset: 'ultrafast' },
    normal: { w: 1920, h: 1080, fps: 30, crf: 18, preset: 'fast'      },
    '4k':   { w: 3840, h: 2160, fps: 30, crf: 15, preset: 'slow'      },
  },

  baseDir:   __dirname,
  screenDir: path.join(__dirname, 'Screen_mobile'),
  tempDir:   path.join(__dirname, 'video_temp_mockup'),
  assetsDir: path.join(__dirname, 'assets'),

  threads:           8,
  cacheEnabled:      true,
  exportFormats:     ['mp4', 'webm'],
  generateGifPreview: true,

  audio: {
    enabled:           false,
    bgMusicPath:       path.join(__dirname, 'assets', 'bg_music.mp3'),
  },
};

const RES = CONFIG.resolutions[CONFIG.mode];
const VW  = RES.w;
const VH  = RES.h;
const FPS = RES.fps;

const S = VW / 1920;   // Facteur d'échelle par rapport à 1080p

const OUTPUT_BASE   = path.join(CONFIG.baseDir, `RohWinBghit_Demo_mockup_${CONFIG.mode}`);
const OUTPUT_MP4    = `${OUTPUT_BASE}.mp4`;
const OUTPUT_WEBM   = `${OUTPUT_BASE}.webm`;
const OUTPUT_GIF    = `${OUTPUT_BASE}_preview.gif`;
const CONCAT_FILE   = path.join(CONFIG.tempDir, 'concat_list.txt');
const CHAPTERS_FILE = path.join(CONFIG.baseDir, 'youtube_chapters_mockup.txt');

[CONFIG.tempDir, CONFIG.assetsDir].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// ═══════════════════════════════════════════════════════════════════════
//  LOGGER WINSTON
// ═══════════════════════════════════════════════════════════════════════

const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp({ format: 'HH:mm:ss' }),
    winston.format.printf(({ timestamp, level, message }) =>
      `${timestamp} [${level.toUpperCase()}] ${message}`)
  ),
  transports: [
    new winston.transports.File({ filename: path.join(CONFIG.baseDir, 'video_mockup.log') }),
    new winston.transports.Console({
      format: winston.format.combine(winston.format.colorize(), winston.format.simple()),
    }),
  ],
});

// ═══════════════════════════════════════════════════════════════════════
//  PALETTE DE COULEURS & BRANDING
// ═══════════════════════════════════════════════════════════════════════

const C = {
  bgDark:  '#0B2E1E',
  bgMid:   '#1C4933',
  primary: '#2E7D52',
  accent:  '#4CAF72',
  gold:    '#C9A84C',
  white:   '#FFFFFF',
  muted:   '#A8C5B5',
  phone:   '#091F14',
  cyan:    '#00C9D4',
  red:     '#C0392B',
  green:   { light: '#5FD68A', dark: '#1B4D3E', neon: '#00FF87' },
};

const rgb = hex => ({
  r: parseInt(hex.slice(1, 3), 16),
  g: parseInt(hex.slice(3, 5), 16),
  b: parseInt(hex.slice(5, 7), 16),
});

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// ═══════════════════════════════════════════════════════════════════════
//  EASING FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

const ease = {
  linear:     t => t,
  inOut:      t => t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t,
  outCubic:   t => 1 - Math.pow(1 - t, 3),
  outQuint:   t => 1 - Math.pow(1 - t, 5),
};

const lerp = (a, b, t) => a + (b - a) * t;

function fadeAlpha(t, fadeIn = 0.15, fadeOut = 0.15, curve = ease.inOut) {
  if (t < fadeIn)      return curve(t / fadeIn);
  if (t > 1 - fadeOut) return curve((1 - t) / fadeOut);
  return 1.0;
}

// ═══════════════════════════════════════════════════════════════════════
//  PARTICULES FLOTTANTES (ARRIÈRE-PLAN DYNAMIQUE)
// ═══════════════════════════════════════════════════════════════════════

function generateParticles(count, seed = 0) {
  const out = [];
  for (let i = 0; i < count; i++) {
    out.push({
      x:     (Math.sin(seed + i * 2.4) * 0.5 + 0.5) * VW,
      y:     (Math.cos(seed + i * 3.1) * 0.5 + 0.5) * VH,
      r:     2 * S + Math.abs(Math.sin(seed + i)) * 5 * S,
      speed: 0.2 + Math.abs(Math.cos(seed + i * 1.7)) * 0.5,
      phase: i * 0.3,
    });
  }
  return out;
}

// ═══════════════════════════════════════════════════════════════════════
//  CACHE
// ═══════════════════════════════════════════════════════════════════════

const CACHE_MANIFEST = path.join(CONFIG.tempDir, 'cache_manifest.json');

function loadCacheManifest() {
  if (fs.existsSync(CACHE_MANIFEST))
    return JSON.parse(fs.readFileSync(CACHE_MANIFEST, 'utf8'));
  return { frames: {}, version: 5, mode: CONFIG.mode };
}

function saveCacheManifest(manifest) {
  fs.writeFileSync(CACHE_MANIFEST, JSON.stringify(manifest, null, 2));
}

function isFrameCached(key, manifest) {
  if (!CONFIG.cacheEnabled) return false;
  const cached = manifest.frames[key];
  return cached && fs.existsSync(cached.path) && manifest.mode === CONFIG.mode;
}

function getFrameCacheKey(act, scene, frameIdx) {
  return `a${act.act}_s${scene.label.substring(0, 15).replace(/\s/g,'_')}_f${frameIdx}`;
}

// ═══════════════════════════════════════════════════════════════════════
//  LIMITATEUR DE CONCURRENCE
// ═══════════════════════════════════════════════════════════════════════

function createLimiter(n) {
  let running = 0;
  const queue = [];
  return fn => new Promise((res, rej) => {
    const run = async () => {
      running++;
      try { res(await fn()); } catch(e) { rej(e); }
      finally { running--; if (queue.length) queue.shift()(); }
    };
    if (running < n) run(); else queue.push(run);
  });
}

// ═══════════════════════════════════════════════════════════════════════
//  RÉSOLUTION DES ASSETS IMAGES
// ═══════════════════════════════════════════════════════════════════════

function resolveImg(filename) {
  if (!filename) return null;
  const variants = [filename, filename.normalize('NFC'), filename.normalize('NFD')];
  const dir = CONFIG.screenDir;
  if (!fs.existsSync(dir)) {
    logger.error(`Dossier Screen_mobile inexistant : ${dir}`);
    return null;
  }
  for (const v of variants) {
    const p = path.join(dir, v);
    if (fs.existsSync(p)) return p;
  }
  for (const f of fs.readdirSync(dir)) {
    if (f.normalize('NFC') === filename.normalize('NFC')) return path.join(dir, f);
  }
  logger.warn(`Asset non trouvé dans Screen_mobile: ${filename}`);
  return null;
}

// ═══════════════════════════════════════════════════════════════════════
//  DIMENSIONS DU MOCKUP TÉLÉPHONE (CENTRÉ)
// ═══════════════════════════════════════════════════════════════════════

const PH_W       = Math.round(360 * S);
const PH_H       = Math.round(720 * S);
const PH_BEZEL   = Math.round(20 * S);
const PH_CORNER  = Math.round(42 * S);
const PH_TOTAL_W = PH_W + PH_BEZEL * 2;
const PH_TOTAL_H = PH_H + PH_BEZEL * 2 + Math.round(50 * S);

const PH_LEFT    = Math.floor((VW - PH_TOTAL_W) / 2);
const PH_TOP     = Math.floor((VH - PH_TOTAL_H) / 2) - Math.round(15 * S); // décalé légèrement vers le haut

const PH_SCR_L   = PH_LEFT + PH_BEZEL;
const PH_SCR_T   = PH_TOP  + PH_BEZEL + Math.round(40 * S);

const FF = `Poppins, -apple-system, Segoe UI, Roboto, Arial, sans-serif`;

// ═══════════════════════════════════════════════════════════════════════
//  RENDU DES COMPOSANTS SVG
// ═══════════════════════════════════════════════════════════════════════

/** Arrière-plan dynamique avec particules */
function backgroundSVG(t = 0) {
  const particles = generateParticles(14, t * 0.4);
  const particlesSVG = particles.map(p => {
    const oy = Math.sin(t * Math.PI * p.speed + p.phase) * 35;
    const op = 0.03 + Math.abs(Math.sin(t * Math.PI * p.speed)) * 0.08;
    return `<circle cx="${p.x.toFixed(1)}" cy="${(p.y + oy).toFixed(1)}" r="${p.r.toFixed(1)}" fill="${C.accent}" opacity="${op.toFixed(3)}"/>`;
  }).join('\n');

  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%"   stop-color="${C.bgDark}"/>
        <stop offset="50%"  stop-color="${C.bgMid}"/>
        <stop offset="100%" stop-color="${C.bgDark}"/>
      </linearGradient>
      <radialGradient id="glow" cx="50%" cy="50%" r="50%">
        <stop offset="0%"   stop-color="${C.accent}" stop-opacity="0.16"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#glow)"/>
    ${particlesSVG}
  </svg>`);
}

/** Boîtier de téléphone réaliste */
function phoneBezelSVG(glowColor = C.accent, glowOpacity = 0.18) {
  const W = PH_TOTAL_W, H = PH_TOTAL_H, R = PH_CORNER;
  const sX = PH_BEZEL, sY = PH_BEZEL + Math.round(40 * S);
  return Buffer.from(`<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <filter id="ph-sh" x="-50%" y="-50%" width="200%" height="200%">
        <feGaussianBlur in="SourceAlpha" stdDeviation="14"/>
        <feOffset dx="0" dy="8" result="ob"/>
        <feFlood flood-color="${glowColor}" flood-opacity="${glowOpacity}"/>
        <feComposite in2="ob" operator="in"/>
        <feMerge><feMergeNode/><feMergeNode in="SourceGraphic"/></feMerge>
      </filter>
      <linearGradient id="metal" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="#1E2220"/>
        <stop offset="40%"  stop-color="#3A3F3C"/>
        <stop offset="60%"  stop-color="#282C29"/>
        <stop offset="100%" stop-color="#121413"/>
      </linearGradient>
      <linearGradient id="shine" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stop-color="white" stop-opacity="0"/>
        <stop offset="50%"  stop-color="white" stop-opacity="0.08"/>
        <stop offset="100%" stop-color="white" stop-opacity="0"/>
      </linearGradient>
    </defs>
    <rect x="3" y="3" width="${W-6}" height="${H-6}" rx="${R}" ry="${R}" fill="url(#metal)" stroke="${C.primary}" stroke-width="1.8" filter="url(#ph-sh)"/>
    <rect x="3" y="3" width="${W-6}" height="${H-6}" rx="${R}" ry="${R}" fill="url(#shine)" opacity="0.3"/>
    <rect x="3" y="3" width="${W-6}" height="${H-6}" rx="${R}" ry="${R}" fill="none" stroke="${glowColor}" stroke-width="1.2" opacity="${glowOpacity * 2.5}"/>
    <rect x="${sX}" y="${sY}" width="${PH_W}" height="${PH_H}" rx="${Math.round(20*S)}" ry="${Math.round(20*S)}" fill="white"/>
    <ellipse cx="${W/2}" cy="${sY-20}" rx="52" ry="13" fill="${C.bgDark}" opacity="0.95"/>
    <circle cx="${W/2-18}" cy="${sY-20}" r="4" fill="#1A1A1A"/>
    <circle cx="${W/2+18}" cy="${sY-20}" r="3" fill="#0D47A1" opacity="0.6"/>
    <rect x="${W/2-48}" y="${H-22}" width="96" height="5" rx="3" fill="${C.primary}" opacity="0.65"/>
    <rect x="-3" y="${H*0.27}" width="5" height="46" rx="3" fill="${C.bgMid}" stroke="#000" stroke-width="0.5"/>
    <rect x="-3" y="${H*0.40}" width="5" height="30" rx="3" fill="${C.bgMid}" stroke="#000" stroke-width="0.5"/>
    <rect x="${W-2}" y="${H*0.33}" width="5" height="62" rx="3" fill="${C.bgMid}" stroke="#000" stroke-width="0.5"/>
  </svg>`);
}

/** Entête de la vidéo (Wordmark & Métadonnées) */
function headerOverlaySVG() {
  const titleY = Math.round(55 * S);
  const sideMargin = Math.round(80 * S);
  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <text x="${sideMargin}" y="${titleY}" font-family="${FF}" font-size="${Math.round(20*S)}" font-weight="700" fill="${C.white}" letter-spacing="0.05em">ROHWINBGHIT</text>
    <text x="${sideMargin}" y="${titleY + Math.round(18*S)}" font-family="${FF}" font-size="${Math.round(10*S)}" font-weight="600" fill="${C.accent}" letter-spacing="0.1em">MOCKUP DÉMONSTRATION</text>
    
    <text x="${VW - sideMargin}" y="${titleY}" text-anchor="end" font-family="${FF}" font-size="${Math.round(13*S)}" font-weight="500" fill="${C.muted}">Soutenance de Master 2</text>
    <text x="${VW - sideMargin}" y="${titleY + Math.round(18*S)}" text-anchor="end" font-family="${FF}" font-size="${Math.round(10*S)}" font-weight="500" fill="${C.primary}">Tlemcen · 2025/2026</text>
    <rect x="0" y="0" width="${VW}" height="${Math.round(5*S)}" fill="${C.accent}" opacity="0.6"/>
  </svg>`);
}

/** Boîte de légende glassmorphism en bas */
function captionOverlaySVG(stepTitle, stepDesc, progress, alpha = 1.0) {
  const cardW = Math.round(840 * S);
  const cardH = Math.round(124 * S);
  const cardX = Math.round((VW - cardW) / 2);
  const cardY = VH - cardH - Math.round(45 * S);
  const progressW = Math.max(0, Math.round(progress * cardW));

  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="cardGrad" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="#0E281C" stop-opacity="0.88"/>
        <stop offset="100%" stop-color="#06140E" stop-opacity="0.96"/>
      </linearGradient>
    </defs>
    <g opacity="${alpha}">
      <rect x="${cardX}" y="${cardY}" width="${cardW}" height="${cardH}" rx="${Math.round(16*S)}" fill="url(#cardGrad)" stroke="${C.accent}" stroke-width="1.5" stroke-opacity="0.25"/>
      <text x="${cardX + Math.round(32*S)}" y="${cardY + Math.round(44*S)}" font-family="${FF}" font-size="${Math.round(18*S)}" font-weight="700" fill="${C.gold}" letter-spacing="0.04em">${esc(stepTitle.toUpperCase())}</text>
      <text x="${cardX + Math.round(32*S)}" y="${cardY + Math.round(80*S)}" font-family="${FF}" font-size="${Math.round(13.5*S)}" font-weight="400" fill="${C.white}" fill-opacity="0.88">${esc(stepDesc)}</text>
      <rect x="${cardX}" y="${cardY + cardH - 4}" width="${cardW}" height="4" rx="2" fill="${C.bgDark}" opacity="0.5"/>
      <rect x="${cardX}" y="${cardY + cardH - 4}" width="${progressW}" height="4" rx="2" fill="${C.accent}"/>
    </g>
  </svg>`);
}

// ═══════════════════════════════════════════════════════════════════════
//  EFFETS SPECIAUX DE L'APPLICATION
// ═══════════════════════════════════════════════════════════════════════

function kycScanOverlay(t) {
  const scanY = Math.round(((t * 1.5) % 1.0) * PH_H);
  return Buffer.from(`<svg width="${PH_W}" height="${PH_H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="scanGrad" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stop-color="${C.cyan}" stop-opacity="0"/>
        <stop offset="100%" stop-color="${C.cyan}" stop-opacity="0.45"/>
      </linearGradient>
    </defs>
    <rect x="0" y="${Math.max(0, scanY - Math.round(40*S))}" width="${PH_W}" height="${Math.min(40*S, scanY)}" fill="url(#scanGrad)"/>
    <rect x="0" y="${scanY}" width="${PH_W}" height="${Math.round(3*S)}" fill="${C.cyan}" opacity="0.9"/>
    
    <rect x="${Math.round(30*S)}" y="${Math.round(30*S)}" width="${Math.round(25*S)}" height="${Math.round(4*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${Math.round(30*S)}" y="${Math.round(30*S)}" width="${Math.round(4*S)}" height="${Math.round(25*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${PH_W - Math.round(55*S)}" y="${Math.round(30*S)}" width="${Math.round(25*S)}" height="${Math.round(4*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${PH_W - Math.round(34*S)}" y="${Math.round(30*S)}" width="${Math.round(4*S)}" height="${Math.round(25*S)}" fill="${C.cyan}" opacity="0.8"/>
    
    <rect x="${Math.round(30*S)}" y="${PH_H - Math.round(34*S)}" width="${Math.round(25*S)}" height="${Math.round(4*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${Math.round(30*S)}" y="${PH_H - Math.round(55*S)}" width="${Math.round(4*S)}" height="${Math.round(25*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${PH_W - Math.round(55*S)}" y="${PH_H - Math.round(34*S)}" width="${Math.round(25*S)}" height="${Math.round(4*S)}" fill="${C.cyan}" opacity="0.8"/>
    <rect x="${PH_W - Math.round(34*S)}" y="${PH_H - Math.round(55*S)}" width="${Math.round(4*S)}" height="${Math.round(25*S)}" fill="${C.cyan}" opacity="0.8"/>
  </svg>`);
}

function qrGlowOverlay(t) {
  const pulse = 0.5 + 0.5 * Math.abs(Math.sin(t * Math.PI * 3.2));
  return Buffer.from(`<svg width="${PH_W}" height="${PH_H}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <radialGradient id="qrGrad" cx="50%" cy="45%" r="35%">
        <stop offset="0%"   stop-color="${C.gold}" stop-opacity="${(pulse * 0.35).toFixed(3)}"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${PH_W}" height="${PH_H}" fill="url(#qrGrad)"/>
    <rect x="${Math.round(PH_W * 0.18)}" y="${Math.round(PH_H * 0.22)}" width="${Math.round(PH_W * 0.64)}" height="${Math.round(PH_H * 0.32)}" rx="${Math.round(12*S)}" fill="none" stroke="${C.gold}" stroke-width="${(1.5 + pulse * 2.5).toFixed(2)}" opacity="${(0.3 + pulse * 0.5).toFixed(3)}"/>
  </svg>`);
}

function gpsPulseOverlay(t) {
  const pulse = (t * 2.0) % 1.0;
  return Buffer.from(`<svg width="${PH_W}" height="${PH_H}" xmlns="http://www.w3.org/2000/svg">
    <circle cx="${Math.round(PH_W * 0.5)}" cy="${Math.round(PH_H * 0.51)}" r="${Math.round(40 * pulse * S)}" fill="none" stroke="${C.accent}" stroke-width="2" opacity="${(1.0 - pulse).toFixed(3)}"/>
    <circle cx="${Math.round(PH_W * 0.5)}" cy="${Math.round(PH_H * 0.51)}" r="${Math.round(16 * S)}" fill="${C.accent}" opacity="0.3"/>
    <circle cx="${Math.round(PH_W * 0.5)}" cy="${Math.round(PH_H * 0.51)}" r="${Math.round(6 * S)}" fill="${C.white}"/>
  </svg>`);
}

// ═══════════════════════════════════════════════════════════════════════
//  SCÈNES DYNAMIQUES HORS-MOCKUP (INTRO / OUTRO)
// ═══════════════════════════════════════════════════════════════════════

function introSVG(t) {
  const alpha  = Math.min(1, ease.outCubic(t * 2.0));
  const logoA  = Math.min(1, ease.outCubic(Math.max(0, (t - 0.2) * 2.2)));
  const scale = 0.95 + 0.05 * ease.outQuint(t);

  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%"   stop-color="${C.bgDark}"/>
        <stop offset="50%"  stop-color="${C.bgMid}"/>
        <stop offset="100%" stop-color="${C.bgDark}"/>
      </linearGradient>
      <radialGradient id="glow" cx="50%" cy="50%" r="40%">
        <stop offset="0%"   stop-color="${C.accent}" stop-opacity="0.22"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#glow)"/>
    <rect x="0" y="0" width="${VW}" height="6" fill="${C.gold}" opacity="0.8"/>
    
    <g transform="translate(${VW/2}, ${VH/2 - Math.round(50*S)}) scale(${scale})" text-anchor="middle" opacity="${alpha}">
      <!-- Animated outline rings -->
      <circle cx="0" cy="0" r="${Math.round(200*S)}" fill="none" stroke="${C.accent}" stroke-width="1.2" opacity="0.08"/>
      <circle cx="0" cy="0" r="${Math.round(130*S)}" fill="none" stroke="${C.accent}" stroke-width="1.5" opacity="0.14"/>
      
      <!-- Logo Wordmark -->
      <text x="0" y="${Math.round(-15*S)}" font-family="${FF}" font-size="${Math.round(86*S)}" font-weight="700" fill="${C.white}" letter-spacing="-0.01em">RohWinBghit</text>
      <text x="0" y="${Math.round(45*S)}" font-family="${FF}" font-size="${Math.round(36*S)}" font-weight="500" fill="${C.gold}">روح وين بغيت</text>
      <text x="0" y="${Math.round(98*S)}" font-family="${FF}" font-size="${Math.round(18*S)}" font-weight="400" fill="${C.accent}" letter-spacing="0.05em">Plateforme Intelligente de Covoiturage Inter-Wilayas</text>
    </g>

    <!-- Subtitle metadata -->
    <g text-anchor="middle" opacity="${logoA}">
      <text x="${VW/2}" y="${VH - Math.round(110*S)}" font-family="${FF}" font-size="${Math.round(14*S)}" fill="${C.muted}">Démonstration des principales fonctionnalités de l'application mobile</text>
      <text x="${VW/2}" y="${VH - Math.round(75*S)}" font-family="${FF}" font-size="${Math.round(14*S)}" font-weight="500" fill="${C.white}">Présenté par : AHMED BACHA Djamel Eddine &amp; BELHORMA Sidi Mohammed Reduane</text>
      <text x="${VW/2}" y="${VH - Math.round(52*S)}" font-family="${FF}" font-size="${Math.round(12*S)}" fill="${C.primary}">Université de Tlemcen · Master 2 Génie Logiciel · 2025/2026</text>
    </g>
    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.gold}" opacity="0.5"/>
  </svg>`);
}

function conclusionSVG(t) {
  const alpha  = Math.min(1, ease.outCubic(t * 2.0));
  const textA  = Math.min(1, ease.outCubic(Math.max(0, (t - 0.25) * 2.0)));
  const scale = 0.95 + 0.05 * ease.outQuint(t);

  return Buffer.from(`<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%"   stop-color="${C.bgDark}"/>
        <stop offset="50%"  stop-color="${C.bgMid}"/>
        <stop offset="100%" stop-color="${C.bgDark}"/>
      </linearGradient>
      <radialGradient id="glow" cx="50%" cy="50%" r="40%">
        <stop offset="0%"   stop-color="${C.accent}" stop-opacity="0.22"/>
        <stop offset="100%" stop-color="transparent" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect width="${VW}" height="${VH}" fill="url(#glow)"/>
    <rect x="0" y="0" width="${VW}" height="6" fill="${C.gold}" opacity="0.8"/>
    
    <g transform="translate(${VW/2}, ${VH/2 - Math.round(40*S)}) scale(${scale})" text-anchor="middle" opacity="${alpha}">
      <text x="0" y="${Math.round(-15*S)}" font-family="${FF}" font-size="${Math.round(76*S)}" font-weight="700" fill="${C.white}" letter-spacing="-0.01em">RohWinBghit</text>
      <text x="0" y="${Math.round(45*S)}" font-family="${FF}" font-size="${Math.round(30*S)}" font-weight="500" fill="${C.gold}">Votre trajet, notre priorité.</text>
      <text x="0" y="${Math.round(92*S)}" font-family="${FF}" font-size="${Math.round(18*S)}" font-weight="400" fill="${C.accent}">Simple • Sûr • Abordable</text>
    </g>

    <g text-anchor="middle" opacity="${textA}">
      <text x="${VW/2}" y="${VH - Math.round(110*S)}" font-family="${FF}" font-size="${Math.round(15*S)}" font-weight="600" fill="${C.white}">Merci pour votre attention !</text>
      <text x="${VW/2}" y="${VH - Math.round(75*S)}" font-family="${FF}" font-size="${Math.round(13*S)}" fill="${C.muted}">Projet de fin d'études Master 2 GL · Université de Tlemcen</text>
    </g>
    <rect x="0" y="${VH - 7}" width="${VW}" height="7" fill="${C.gold}" opacity="0.5"/>
  </svg>`);
}

// ═══════════════════════════════════════════════════════════════════════
//  COMPOSITEURS DE FRAMES
// ═══════════════════════════════════════════════════════════════════════

async function composeSVGFrame(svgStr, frameIdx, cacheKey, manifest) {
  const out = path.join(CONFIG.tempDir, `f${String(frameIdx).padStart(6, '0')}.png`);
  if (isFrameCached(cacheKey, manifest)) return manifest.frames[cacheKey].path;
  if (fs.existsSync(out)) {
    manifest.frames[cacheKey] = { path: out, ts: Date.now() };
    return out;
  }
  await sharp({ create: { width: VW, height: VH, channels: 4, background: { r:0, g:0, b:0, alpha:255 } } })
    .composite([{ input: svgStr, top: 0, left: 0 }])
    .png({ compressionLevel: 3 })
    .toFile(out);
  manifest.frames[cacheKey] = { path: out, ts: Date.now() };
  return out;
}

async function composePhoneFrame(opts, frameIdx, cacheKey, manifest) {
  const { act, scene, imgBuf, t, progress } = opts;
  const out = path.join(CONFIG.tempDir, `f${String(frameIdx).padStart(6, '0')}.png`);
  if (isFrameCached(cacheKey, manifest)) return manifest.frames[cacheKey].path;
  if (fs.existsSync(out)) {
    manifest.frames[cacheKey] = { path: out, ts: Date.now() };
    return out;
  }

  const alpha = fadeAlpha(t, 0.18, 0.15);

  // Apply Ken Burns Zoom
  let kbBuf = imgBuf;
  try {
    const zoom = lerp(1.0, 1.07, ease.inOut(t));
    const meta = await sharp(imgBuf).metadata();
    const srcW = Math.round(meta.width / zoom);
    const srcH = Math.round(meta.height / zoom);
    
    // Extract and zoom
    let zoomed = await sharp(imgBuf)
      .extract({
        left:   Math.max(0, Math.round((meta.width - srcW) / 2)),
        top:    Math.max(0, Math.round((meta.height - srcH) / 2)),
        width:  srcW,
        height: srcH,
      })
      .resize(PH_W, PH_H, { fit: 'fill' })
      .png().toBuffer();

    // Mask with rounded corners to fit inside the bezel perfectly
    const roundedCornerMask = Buffer.from(
      `<svg width="${PH_W}" height="${PH_H}"><rect x="0" y="0" width="${PH_W}" height="${PH_H}" rx="${Math.round(20*S)}" ry="${Math.round(20*S)}" fill="#fff"/></svg>`
    );
    kbBuf = await sharp(zoomed)
      .composite([{ input: roundedCornerMask, blend: 'dest-in' }])
      .png().toBuffer();
  } catch (e) {
    logger.error(`Ken Burns error composition on frame ${frameIdx}: ${e.message}`);
  }

  // Generate frame layers
  const bg      = backgroundSVG(t);
  const header  = headerOverlaySVG();
  const bezel   = phoneBezelSVG(C.accent, 0.18 * alpha);
  const caption = captionOverlaySVG(act.title, act.desc, progress, alpha);

  const composites = [
    { input: bg,      top: 0,        left: 0 },
    { input: header,  top: 0,        left: 0 },
    { input: bezel,   top: PH_TOP,   left: PH_LEFT },
    { input: kbBuf,    top: PH_SCR_T, left: PH_SCR_L },
  ];

  // Visual interactive overlays on the phone screen
  if (scene.effect === 'kyc_scan') composites.push({ input: kycScanOverlay(t), top: PH_SCR_T, left: PH_SCR_L });
  if (scene.effect === 'qr_glow')  composites.push({ input: qrGlowOverlay(t),  top: PH_SCR_T, left: PH_SCR_L });
  if (scene.effect === 'gps_pulse') composites.push({ input: gpsPulseOverlay(t), top: PH_SCR_T, left: PH_SCR_L });

  // Glassmorphism caption overlay at the very bottom
  composites.push({ input: caption, top: 0, left: 0 });

  await sharp({ create: { width: VW, height: VH, channels: 3, background: rgb(C.bgDark) } })
    .composite(composites)
    .png({ compressionLevel: 3 })
    .toFile(out);
  manifest.frames[cacheKey] = { path: out, ts: Date.now() };
  return out;
}

async function renderFrame(act, scene, imgBuf, t, frameIdx, progress, cacheKey, manifest) {
  if (scene.type === 'intro_animated') {
    return composeSVGFrame(introSVG(t), frameIdx, cacheKey, manifest);
  }
  if (scene.type === 'conclusion_animated') {
    return composeSVGFrame(conclusionSVG(t), frameIdx, cacheKey, manifest);
  }
  if (imgBuf) {
    return composePhoneFrame({ act, scene, imgBuf, t, progress }, frameIdx, cacheKey, manifest);
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════
//  SCÉNARIO COMPLET — DÉMO MOCKUP (PARCOURS UTILISATEUR EN 10 ÉTAPES)
// ═══════════════════════════════════════════════════════════════════════

const SCENARIO = [
  // INTRO (5s)
  {
    act: 0, tag: 'Introduction', title: 'Introduction', desc: '',
    scenes: [{ type: 'intro_animated', duration: 5, label: 'Intro' }]
  },
  // 1. Écran d'accueil (8s)
  {
    act: 1, tag: "Écran d'accueil",
    title: "1. Écran d'accueil",
    desc: "Découvrez l'interface de bienvenue et le carrousel des fonctionnalités clés.",
    scenes: [
      { file: 'Screen1.jpg', duration: 4, label: "Écran de démarrage (Splash Screen)" },
      { file: 'Screen2.jpg', duration: 4, label: "Onboarding & fonctionnalités clés" }
    ]
  },
  // 2. Inscription & Connexion (8s)
  {
    act: 2, tag: "Inscription & Connexion",
    title: "2. Inscription & Connexion",
    desc: "Créez votre compte rapidement et choisissez votre rôle : Passager ou Conducteur.",
    scenes: [
      { file: 'inscription.jpg', duration: 8, label: "Sélection du rôle et inscription utilisateur" }
    ]
  },
  // 3. Vérification KYC (16s)
  {
    act: 3, tag: "Vérification KYC",
    title: "3. Vérification KYC Biométrique",
    desc: "Un système de sécurité avancé avec extraction OCR de la CIN et reconnaissance faciale par IA.",
    scenes: [
      { file: 'd_27_verif_intro.png', duration: 4, label: "Introduction de la vérification d'identité" },
      { file: 'd_28_verif_id_capture.png', duration: 4, label: "Capture de la pièce d'identité et OCR", effect: 'kyc_scan' },
      { file: 'd_29_verif_face_capture.png', duration: 4, label: "Scan facial & détection de vivacité", effect: 'kyc_scan' },
      { file: 'd_30_verif_pending.png', duration: 4, label: "Traitement intelligent et validation d'identité" }
    ]
  },
  // 4. Recherche de trajet (10s)
  {
    act: 4, tag: "Recherche de trajet",
    title: "4. Recherche de Trajet",
    desc: "Saisissez votre destination et trouvez instantanément les trajets disponibles.",
    scenes: [
      { file: 'Passager réserve.jpg', duration: 5, label: "Saisie de l'itinéraire et de la date" },
      { file: 'p_08_search_results.png', duration: 5, label: "Recherche de trajets et liste des conducteurs vérifiés" }
    ]
  },
  // 5. Publication de trajet (12s)
  {
    act: 5, tag: "Publication de trajet",
    title: "5. Publication d'un Trajet",
    desc: "Conducteurs : publiez votre itinéraire en quelques clics et partagez vos frais.",
    scenes: [
      { file: 'd_03_tab_publish.png', duration: 6, label: "Configuration du trajet (départ, arrivée, prix)" },
      { file: 'Trajet publié.jpg', duration: 6, label: "Trajet enregistré et publié sur la carte" }
    ]
  },
  // 6. Réservation de trajet (8s)
  {
    act: 6, tag: "Réservation de trajet",
    title: "6. Réservation d'un Trajet",
    desc: "Sélectionnez votre trajet idéal, examinez les détails et réservez instantanément.",
    scenes: [
      { file: 'p_09_trip_details.png', duration: 8, label: "Détails du trajet et réservation passager" }
    ]
  },
  // 7. Paiement sécurisé (8s)
  {
    act: 7, tag: "Paiement sécurisé",
    title: "7. Paiement Sécurisé",
    desc: "Réglez vos réservations de manière sécurisée via CIB, Edahabia, ou en espèces.",
    scenes: [
      { file: 'd_11_wallet.png', duration: 8, label: "Portefeuille numérique et passerelle de paiement" }
    ]
  },
  // 8. Génération du QR Code (8s)
  {
    act: 8, tag: "QR Code unique",
    title: "8. Billet & QR Code Unique",
    desc: "Votre ticket d'embarquement numérique est généré. Scannez au départ pour valider le trajet.",
    scenes: [
      { file: 'p_10_ticket.png', duration: 4, label: "Billet d'embarquement avec QR Code sécurisé", effect: 'qr_glow' },
      { file: 'QR scanné au départ.jpg', duration: 4, label: "Validation de la présence par scan de sécurité au départ", effect: 'qr_glow' }
    ]
  },
  // 9. Suivi en temps réel (10s)
  {
    act: 9, tag: "Suivi en temps réel",
    title: "9. Suivi GPS en Temps Réel",
    desc: "Suivez le déplacement du véhicule en temps réel sur la carte interactive.",
    scenes: [
      { file: 'p_26_live_tracking.png', duration: 10, label: "Suivi en temps réel de la position du véhicule", effect: 'gps_pulse' }
    ]
  },
  // 10. Profil utilisateur (8s)
  {
    act: 10, tag: "Profil utilisateur",
    title: "10. Profil Utilisateur",
    desc: "Gérez vos informations personnelles, vos évaluations et vos statistiques de confiance.",
    scenes: [
      { file: 'p_05_tab_profile.png', duration: 8, label: "Profil utilisateur complet et niveau de confiance" }
    ]
  },
  // OUTRO (4s)
  {
    act: 11, tag: 'Conclusion', title: 'Conclusion', desc: '',
    scenes: [{ type: 'conclusion_animated', duration: 4, label: 'Outro' }]
  }
];

// ═══════════════════════════════════════════════════════════════════════
//  TRAITEMENT DE RENDU DE FRAMES
// ═══════════════════════════════════════════════════════════════════════

async function processAllScenes() {
  const limit    = createLimiter(CONFIG.threads);
  const manifest = loadCacheManifest();

  let frameIdx = 0, missing = 0;
  const frameList  = [];

  const totalFrames = SCENARIO.reduce((s, act) =>
    s + act.scenes.reduce((ss, sc) => ss + Math.round(sc.duration * FPS), 0), 0);

  logger.info(`Mode : ${CONFIG.mode.toUpperCase()} (${VW}x${VH} @ ${FPS}fps)`);
  logger.info(`Frames à générer : ~${totalFrames}  (${(totalFrames / FPS).toFixed(0)} secondes)`);

  const bar = new cliProgress.SingleBar({
    format: '  Rendu |{bar}| {percentage}% | {value}/{total} frames | ETA: {eta}s',
    barCompleteChar: '\u2588',
    barIncompleteChar: '\u2591',
    hideCursor: true,
  }, cliProgress.Presets.shades_classic);
  bar.start(totalFrames, 0);

  // Calcul du nombre d'actes pertinents pour la barre de progression (excluant intro/outro)
  const numActs = SCENARIO.filter(a => a.act > 0 && a.act <= 10).length;

  for (let ai = 0; ai < SCENARIO.length; ai++) {
    const act = SCENARIO[ai];
    
    // Calcul de la progression globale sur la frise (0.0 à 1.0)
    let actProgress = 0;
    if (act.act > 0 && act.act <= 10) {
      actProgress = act.act / numActs;
    } else if (act.act > 10) {
      actProgress = 1.0;
    }

    logger.info(`[Act ${String(act.act).padStart(2,'0')}] ${act.tag}`);

    for (const scene of act.scenes) {
      const nFrames = Math.round(scene.duration * FPS);
      let imgBuf = null;

      if (scene.file) {
        const imgPath = resolveImg(scene.file);
        if (!imgPath) {
          missing++;
          bar.increment(nFrames);
          continue;
        }
        try {
          // Pre-load screenshot at higher resolution to preserve quality during zooms
          imgBuf = await sharp(imgPath)
            .resize(PH_W * 2, PH_H * 2, { fit: 'cover', position: 'top' })
            .png().toBuffer();
        } catch (e) {
          logger.error(`Erreur chargement image ${scene.file}: ${e.message}`);
          missing++;
          bar.increment(nFrames);
          continue;
        }
      }

      const sceneStart = frameIdx;
      const tasks = Array.from({ length: nFrames }, (_, fi) => {
        const t        = nFrames > 1 ? fi / (nFrames - 1) : 0;
        const idx      = sceneStart + fi;
        const cacheKey = getFrameCacheKey(act, scene, idx);
        
        return limit(async () => {
          const p = await renderFrame(act, scene, imgBuf, t, idx, actProgress, cacheKey, manifest);
          bar.increment();
          return p;
        });
      });

      const paths = await Promise.all(tasks);
      for (const p of paths) if (p) frameList.push({ path: p, duration: 1 / FPS });
      frameIdx += nFrames;
    }
  }

  bar.stop();
  saveCacheManifest(manifest);
  if (missing > 0) logger.warn(`${missing} image(s) manquante(s) ignorée(s)`);
  const dur = frameList.length / FPS;
  logger.info(`Total : ${frameList.length} frames (${dur.toFixed(0)}s = ${(dur / 60).toFixed(1)} min)`);
  return frameList;
}

// ═══════════════════════════════════════════════════════════════════════
//  ASSEMBLAGE FFMEPG
// ═══════════════════════════════════════════════════════════════════════

function writeConcatFile(frameList) {
  const lines = frameList.flatMap(f => [
    `file '${f.path.replace(/\\/g, '/')}'`,
    `duration ${f.duration.toFixed(6)}`,
  ]);
  // FFmpeg concat bug workaround (repeat last file without duration)
  lines.push(`file '${frameList[frameList.length - 1].path.replace(/\\/g, '/')}'`);
  fs.writeFileSync(CONCAT_FILE, lines.join('\n'));
}

async function assembleVideo(frameList) {
  return new Promise((resolve, reject) => {
    writeConcatFile(frameList);
    const durSec = (frameList.length / FPS).toFixed(0);
    logger.info(`Assemblage MP4 (${durSec}s @ ${FPS}fps | CRF ${RES.crf})`);

    let cmd = ffmpeg()
      .input(CONCAT_FILE)
      .inputOptions(['-f', 'concat', '-safe', '0'])
      .videoFilters([
        `fps=${FPS}`,
        `scale=${VW}:${VH}:force_original_aspect_ratio=disable`,
      ])
      .outputOptions([
        '-c:v', 'libx264',
        '-preset', RES.preset,
        '-crf', String(RES.crf),
        '-pix_fmt', 'yuv420p',
        '-movflags', '+faststart',
        '-profile:v', 'high',
        '-level', '4.1',
      ]);

    if (CONFIG.audio.enabled && fs.existsSync(CONFIG.audio.bgMusicPath)) {
      cmd = cmd
        .input(CONFIG.audio.bgMusicPath)
        .outputOptions(['-c:a', 'aac', '-b:a', '192k', '-shortest']);
      logger.info('Audio intégré : musique de fond');
    }

    cmd
      .output(OUTPUT_MP4)
      .on('start', c => logger.info(`FFmpeg CMD : ${c.slice(0, 120)}...`))
      .on('progress', p => {
        if (p.percent != null) process.stdout.write(`\r  Encodage : ${Math.round(p.percent)}%   `);
      })
      .on('end', () => { process.stdout.write('\r'); logger.info('MP4 finalisé.'); resolve(); })
      .on('error', (err, _s, stderr) => {
        logger.error(`FFmpeg Error : ${err.message}`);
        if (stderr) console.error(stderr.slice(-500));
        reject(err);
      })
      .run();
  });
}

async function exportWebM(frameList) {
  if (!CONFIG.exportFormats.includes('webm')) return;
  return new Promise((resolve, reject) => {
    logger.info('Export WebM (VP9)...');
    writeConcatFile(frameList);
    ffmpeg()
      .input(CONCAT_FILE)
      .inputOptions(['-f', 'concat', '-safe', '0'])
      .videoFilters([`fps=${FPS}`, `scale=${VW}:${VH}:force_original_aspect_ratio=disable`])
      .outputOptions(['-c:v', 'libvpx-vp9', '-b:v', '0', '-crf', '30',
                      '-pix_fmt', 'yuv420p', '-deadline', 'good', '-cpu-used', '4'])
      .output(OUTPUT_WEBM)
      .on('end', () => { logger.info('WebM finalisé.'); resolve(); })
      .on('error', e => { logger.warn(`WebM échoué : ${e.message} (ignoré)`); resolve(); })
      .run();
  });
}

async function exportGifPreview() {
  if (!CONFIG.generateGifPreview) return;
  return new Promise((resolve, reject) => {
    logger.info('Génération GIF de prévisualisation (10s, 480p)...');
    ffmpeg()
      .input(OUTPUT_MP4)
      .inputOptions(['-t', '10'])
      .outputOptions([
        '-vf', 'fps=10,scale=854:480:flags=lanczos,split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse',
        '-loop', '0',
      ])
      .output(OUTPUT_GIF)
      .on('end', () => { logger.info('GIF de prévisualisation généré.'); resolve(); })
      .on('error', e => { logger.warn(`GIF échoué : ${e.message} (ignoré)`); resolve(); })
      .run();
  });
}

function generateYouTubeChapters() {
  let totalTime = 0;
  const chapters = [];
  for (const act of SCENARIO) {
    const dur  = act.scenes.reduce((s, sc) => s + sc.duration, 0);
    const mins = Math.floor(totalTime / 60);
    const secs = String(Math.round(totalTime % 60)).padStart(2, '0');
    chapters.push(`${mins}:${secs} ${act.tag}`);
    totalTime += dur;
  }
  fs.writeFileSync(CHAPTERS_FILE, chapters.join('\n'));
  logger.info(`Chapitres YouTube générés : ${CHAPTERS_FILE}`);
}

function cleanup() {
  // Désactivé pour préserver le cache et éviter les suppressions prématurées
  /*
  if (!CONFIG.cacheEnabled) {
    try { fs.rmSync(CONFIG.tempDir, { recursive: true, force: true }); } catch (_) {}
  }
  */
}

// ═══════════════════════════════════════════════════════════════════════
//  POINT D'ENTRÉE PRINCIPAL
// ═══════════════════════════════════════════════════════════════════════

(async () => {
  const t0  = Date.now();
  const sep = '═'.repeat(78);

  console.log(sep);
  console.log('  RohWinBghit - Générateur Vidéo Démo Mockup');
  console.log(`  Mode : ${CONFIG.mode.toUpperCase()} | ${VW}x${VH} @ ${FPS}fps | CRF ${RES.crf} | ${RES.preset}`);
  console.log('  Mockup smartphone centré • Légende en bas • Sans slides techniques');
  console.log(sep);

  const screenFiles = fs.existsSync(CONFIG.screenDir)
    ? fs.readdirSync(CONFIG.screenDir).filter(f => /\.(png|jpg|jpeg)$/i.test(f) && !f.startsWith('.'))
    : [];
  logger.info(`Captures disponibles dans Screen_mobile: ${screenFiles.length} fichiers`);

  try {
    const frameList = await processAllScenes();
    if (!frameList.length) throw new Error('Aucune frame générée — vérifiez les assets.');

    await assembleVideo(frameList);
    await exportWebM(frameList);
    await exportGifPreview();
    generateYouTubeChapters();

    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    const sizeMB  = (fs.statSync(OUTPUT_MP4).size / 1024 / 1024).toFixed(1);
    const dur     = frameList.length / FPS;

    console.log('\n' + sep);
    console.log('  GÉNÉRATION TERMINÉE');
    console.log(`  MP4    : ${OUTPUT_MP4}`);
    console.log(`  WebM   : ${CONFIG.exportFormats.includes('webm') ? OUTPUT_WEBM : 'Non généré'}`);
    console.log(`  GIF    : ${CONFIG.generateGifPreview ? OUTPUT_GIF : 'Non généré'}`);
    console.log(`  Durée  : ${Math.floor(dur/60)}min ${Math.round(dur%60)}s  (${frameList.length} frames)`);
    console.log(`  Taille : ${sizeMB} MB`);
    console.log(`  Temps  : ${elapsed}s`);
    console.log(sep + '\n');

    cleanup();
  } catch (e) {
    logger.error(`Erreur fatale : ${e.message}`);
    if (e.stack) console.error(e.stack.split('\n').slice(0, 5).join('\n'));
    process.exit(1);
  }
})();
