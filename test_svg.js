'use strict';
const sharp = require('sharp');
const VW=1920, VH=1080, SX=1, SY=1;
const C = {
  bgDark:'#0B2E1E', bgMid:'#1C4933', primary:'#2E7D52', accent:'#4CAF72',
  gold:'#C9A84C', white:'#FFFFFF', muted:'#A8C5B5', phone:'#091F14',
  green:{light:'#5FD68A'}, gray:{}, status:{}, teal:'#028090',
  red:'#C0392B', cyan:'#00C9D4', purple:'#7C3AED',
};
const FF = 'Poppins, -apple-system, Segoe UI, Roboto, Arial, sans-serif';
const ease = {
  out:      t => 1-(1-t)*(1-t),
  outCubic: t => 1-Math.pow(1-t,3),
  inOut:    t => t < 0.5 ? 2*t*t : -1+(4-2*t)*t,
};
const lerp = (a,b,t) => a+(b-a)*t;
function esc(s){
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;')
    .replace(/[\u{1F000}-\u{1FFFF}]/gu,'')
    .replace(/[\u2700-\u27BF]/g,'')
    .replace(/[\u2600-\u26FF]/g,'');
}

function generateParticles(count, seed=0){
  const out=[];
  for(let i=0;i<count;i++){
    out.push({ x:(Math.sin(seed+i*2.4)*0.5+0.5)*VW, y:(Math.cos(seed+i*3.1)*0.5+0.5)*VH,
               r:2+Math.abs(Math.sin(seed+i))*4, speed:0.3+Math.abs(Math.cos(seed+i*1.7))*0.7, phase:i*0.3 });
  }
  return out;
}

function introSVG(t) {
  const textA    = Math.min(1, ease.outCubic(t * 2.5));
  const mapA     = Math.min(1, ease.outCubic(Math.max(0, (t - 0.15) * 2)));
  const routeP   = Math.min(1, ease.out(Math.max(0, (t - 0.25) * 2.5)));
  const dotPulse = 0.72 + 0.28 * Math.sin(t * Math.PI * 5);
  const parallaxMap  = Math.sin(t * Math.PI * 0.3) * 8;
  const parallaxText = Math.sin(t * Math.PI * 0.2) * -5;
  const pillA = Math.min(1, Math.max(0, (t - 0.55) * 4));

  const cities = [
    { n:'Alger', x:Math.round(720*SX), y:Math.round(298*SY), major:true },
    { n:'Oran',  x:Math.round(385*SX), y:Math.round(296*SY), major:true },
    { n:'Tlemcen',x:Math.round(316*SX),y:Math.round(332*SY), major:true },
  ];
  const routes = [[0,1],[0,2]];
  const routeSVG = routes.map(([a,b],i) => {
    const rp = Math.max(0, routeP - i*0.07);
    const ca=cities[a], cb=cities[b];
    return `<line x1="${ca.x}" y1="${ca.y}" x2="${(ca.x+(cb.x-ca.x)*rp).toFixed(1)}" y2="${(ca.y+(cb.y-ca.y)*rp).toFixed(1)}" stroke="${C.green.light}" stroke-width="2" opacity="${(mapA*0.65).toFixed(3)}" stroke-dasharray="9 5"/>`;
  }).join('\n');
  const dotsSVG = cities.map((c,i) => {
    const da=Math.min(1,Math.max(0,mapA-i*0.055));
    const r=c.major?8:5, pr=c.major?22:14;
    return `<circle cx="${c.x}" cy="${c.y}" r="${pr}" fill="${C.accent}" opacity="${(da*0.14*dotPulse).toFixed(3)}"/><circle cx="${c.x}" cy="${c.y}" r="${r}" fill="${C.accent}" opacity="${(da*0.9).toFixed(3)}"/>`;
  }).join('\n');

  return `<svg width="${VW}" height="${VH}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${C.bgDark}"/>
        <stop offset="100%" stop-color="${C.bgMid}"/>
      </linearGradient>
    </defs>
    <rect width="${VW}" height="${VH}" fill="url(#bg)"/>
    <rect x="0" y="0" width="${VW}" height="6" fill="${C.gold}" opacity="0.75"/>
    <g transform="translate(${(Math.round(55*SX)+parallaxMap).toFixed(1)}, ${Math.round(60*SY)}) scale(${(0.88*SX).toFixed(4)})" opacity="${mapA.toFixed(3)}">
      <path d="M100,50 L680,45 L760,95 L800,230 L775,390 L720,560 L670,760 L580,880 L390,900 L210,820 L80,660 L35,460 L40,260 L70,140 Z" fill="none" stroke="${C.primary}" stroke-width="2.5" opacity="${(mapA*0.4).toFixed(3)}" stroke-dasharray="14 7"/>
      ${routeSVG}
      ${dotsSVG}
    </g>
    <g transform="translate(${parallaxText.toFixed(1)}, 0)">
      <text x="${VW-Math.round(580*SX)}" y="${VH/2-Math.round(95*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(94*SX)}" font-weight="700" fill="${C.white}" opacity="${textA.toFixed(3)}">RohWinBghit</text>
      <text x="${VW-Math.round(580*SX)}" y="${VH/2-Math.round(12*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(40*SX)}" font-weight="500" fill="${C.gold}" opacity="${textA.toFixed(3)}">Roh Win Bghit</text>
      <text x="${VW-Math.round(580*SX)}" y="${VH/2+Math.round(58*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(21*SX)}" font-weight="400" fill="${C.accent}" opacity="${textA.toFixed(3)}">Plateforme intelligente de covoiturage inter-wilayas</text>
    </g>
    <rect x="${VW-Math.round(790*SX)}" y="${VH/2+Math.round(90*SY)}" width="${Math.round(420*SX)}" height="2" rx="1" fill="${C.gold}" opacity="${(textA*0.55).toFixed(3)}"/>
    <rect x="${VW-Math.round(820*SX)}" y="${VH/2+Math.round(115*SY)}" width="${Math.round(190*SX)}" height="${Math.round(34*SY)}" rx="17" fill="${C.primary}" opacity="${(pillA*0.88).toFixed(3)}"/>
    <text x="${VW-Math.round(725*SX)}" y="${VH/2+Math.round(138*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(13*SX)}" font-weight="600" fill="${C.white}" opacity="${pillA.toFixed(3)}">KYC Biometrique</text>
    <text x="${VW-Math.round(580*SX)}" y="${VH-Math.round(72*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(16*SX)}" fill="${C.muted}" opacity="${textA.toFixed(3)}">AHMED BACHA Djamel Eddine x BELHORMA Sidi Mohammed Reduane</text>
    <text x="${VW-Math.round(580*SX)}" y="${VH-Math.round(46*SY)}" text-anchor="middle" font-family="${FF}" font-size="${Math.round(14*SX)}" fill="${C.primary}" opacity="${textA.toFixed(3)}">Encadrante : Mme BENLEDGHEM Rafika x 2025/2026</text>
    <rect x="0" y="${VH-7}" width="${VW}" height="7" fill="${C.gold}" opacity="0.4"/>
  </svg>`;
}

(async () => {
  for (const t of [0, 0.3, 0.7, 1.0]) {
    const svg = introSVG(t);
    const lines = svg.split('\n');
    console.log(`t=${t}: SVG has ${lines.length} lines`);
    // Check for problematic patterns
    let inTag = false;
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      if (l.includes('</svg>') && i !== lines.length - 1) {
        console.log(`  WARNING: </svg> not at last line (line ${i+1})`);
      }
    }
    try {
      await sharp({ create: {width:VW, height:VH, channels:4, background:{r:0,g:0,b:0,alpha:255}} })
        .composite([{input: Buffer.from(svg), top:0, left:0}])
        .png().toBuffer();
      console.log(`  t=${t}: OK`);
    } catch(e) {
      console.error(`  t=${t}: FAIL — ${e.message.slice(0,120)}`);
      // Print SVG around line 113
      console.log('  Lines around 110-115:');
      lines.slice(108,116).forEach((l,i) => console.log(`    ${i+109}: ${l}`));
    }
  }
  console.log('Test done');
})();
