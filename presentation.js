#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const http = require('http');
const { exec } = require('child_process');

// Locate HTML file (case-insensitive)
const possiblePaths = [
    path.join(__dirname, 'Presentatione.html'),
    path.join(__dirname, 'presentatione.html')
];
let HTML_FILE = '';
for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
        HTML_FILE = p;
        break;
    }
}

if (!HTML_FILE) {
    console.error('Error: Presentation HTML file (Presentatione.html) not found in the script directory.');
    process.exit(1);
}

const htmlContent = fs.readFileSync(HTML_FILE, 'utf8');

// Parse the slides
const slidesRaw = htmlContent.split(/<!DOCTYPE html>/i).slice(1);
const slides = [];

// Scopes CSS rules to a specific container prefix to prevent leakages
function scopeCss(css, prefix) {
    if (!css) return '';
    return css.replace(/([^\r\n,{}]+)(?=\s*\{)/g, (match) => {
        return match.split(',').map(selector => {
            selector = selector.trim();
            if (selector === 'body') {
                return prefix;
            }
            if (selector.startsWith('@')) {
                return selector;
            }
            return `${prefix} ${selector}`;
        }).join(', ');
    });
}

for (let i = 0; i < slidesRaw.length; i++) {
    const slideHtml = slidesRaw[i];
    
    // Extract title
    const titleMatch = slideHtml.match(/<title>([\s\S]*?)<\/title>/i);
    const title = titleMatch ? titleMatch[1].trim() : `Slide ${i + 1}`;
    
    // Extract style
    const styleMatch = slideHtml.match(/<style>([\s\S]*?)<\/style>/i);
    const style = styleMatch ? styleMatch[1].trim() : '';
    
    // Extract body content
    const bodyMatch = slideHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    const body = bodyMatch ? bodyMatch[1].trim() : '';
    
    // Extract body background color
    const bgMatch = slideHtml.match(/body\s*\{[^}]*background-color\s*:\s*([^;}]+)/i);
    const bgColor = bgMatch ? bgMatch[1].trim() : '#F5F1E8';
    
    slides.push({
        index: i,
        title,
        style: scopeCss(style, `#slide-frame-${i}`),
        body,
        bgColor
    });
}

// Generate the unified HTML presentation
const parentHtml = `
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RohWinBghit - Présentation</title>
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;800&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        * {
            box-sizing: border-box;
        }
        body {
            margin: 0;
            padding: 0;
            overflow: hidden;
            width: 100vw;
            height: 100vh;
            background-color: #0F3A23;
            transition: background-color 0.5s ease;
            font-family: 'Manrope', sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        #slide-container {
            width: 100vw;
            height: 100vh;
            position: relative;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        #slide-frame {
            width: 1280px;
            height: 720px;
            position: absolute;
            top: 50%;
            left: 50%;
            transform-origin: center center;
            transform: translate(-50%, -50%) scale(1);
            box-shadow: 0 20px 50px rgba(0,0,0,0.5);
            border-radius: 12px;
            overflow: hidden;
            background-color: #FFFFFF;
        }
        .slide {
            width: 100%;
            height: 100%;
            display: none;
            position: absolute;
            top: 0;
            left: 0;
        }
        .slide.active {
            display: block;
        }
        
        /* Scoped Slide Styles */
        ${slides.map(s => s.style).join('\n')}

        /* UI Overlays */
        #controls {
            position: absolute;
            bottom: 24px;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            align-items: center;
            gap: 16px;
            background: rgba(0, 0, 0, 0.75);
            backdrop-filter: blur(10px);
            padding: 8px 20px;
            border-radius: 30px;
            z-index: 1000;
            transition: opacity 0.3s ease;
            border: 1px solid rgba(255,255,255,0.15);
        }
        #controls.hidden {
            opacity: 0;
            pointer-events: none;
        }
        .btn {
            background: transparent;
            border: none;
            color: #FFFFFF;
            font-size: 18px;
            cursor: pointer;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            transition: background 0.2s;
        }
        .btn:hover {
            background: rgba(255, 255, 255, 0.2);
        }
        #slide-select {
            background: transparent;
            border: none;
            color: #FFFFFF;
            font-family: 'JetBrains Mono', monospace;
            font-size: 14px;
            font-weight: bold;
            outline: none;
            cursor: pointer;
            padding: 4px 8px;
        }
        #slide-select option {
            background: #181818;
            color: #FFF;
        }
        #progress-bar-container {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
            background: rgba(255, 255, 255, 0.1);
            z-index: 1001;
        }
        #progress-bar {
            height: 100%;
            width: 0%;
            background: #4CC080;
            transition: width 0.3s ease;
        }
        #fullscreen-btn {
            position: absolute;
            top: 24px;
            right: 24px;
            z-index: 1000;
            background: rgba(0, 0, 0, 0.6);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.1);
        }
    </style>
</head>
<body>
    <div id="progress-bar-container">
        <div id="progress-bar"></div>
    </div>

    <button id="fullscreen-btn" class="btn" onclick="toggleFullscreen()">
        <i class="fa-solid fa-expand"></i>
    </button>

    <div id="slide-container">
        <div id="slide-frame">
            ${slides.map((s, idx) => `
                <div id="slide-frame-${idx}" class="slide ${idx === 0 ? 'active' : ''}">
                    ${s.body}
                </div>
            `).join('\n')}
        </div>
    </div>

    <div id="controls">
        <button class="btn" onclick="prevSlide()"><i class="fa-solid fa-chevron-left"></i></button>
        <select id="slide-select" onchange="goToSlide(parseInt(this.value))">
            ${slides.map((s, idx) => `<option value="${idx}">${(idx+1).toString().padStart(2, '0')}/${slides.length} — ${s.title}</option>`).join('\n')}
        </select>
        <button class="btn" onclick="nextSlide()"><i class="fa-solid fa-chevron-right"></i></button>
    </div>

    <script>
        const slides = ${JSON.stringify(slides.map(s => ({ bgColor: s.bgColor })))};
        let currentSlide = 0;
        let controlsTimeout;

        function showSlide(index) {
            if (index < 0 || index >= slides.length) return;
            
            document.querySelectorAll('.slide').forEach((el, idx) => {
                if (idx === index) {
                    el.classList.add('active');
                } else {
                    el.classList.remove('active');
                }
            });

            currentSlide = index;
            document.getElementById('slide-select').value = index;
            
            // Set body background color matching the slide context
            document.body.style.backgroundColor = slides[index].bgColor;

            // Update progress bar
            const progress = ((index + 1) / slides.length) * 100;
            document.getElementById('progress-bar').style.width = progress + '%';
        }

        function nextSlide() {
            if (currentSlide < slides.length - 1) {
                showSlide(currentSlide + 1);
            }
        }

        function prevSlide() {
            if (currentSlide > 0) {
                showSlide(currentSlide - 1);
            }
        }

        function goToSlide(index) {
            showSlide(index);
        }

        // Auto scale slide frame to fit viewport perfectly
        function resize() {
            const frame = document.getElementById('slide-frame');
            const width = window.innerWidth;
            const height = window.innerHeight;
            const scale = Math.min(width / 1280, height / 720);
            frame.style.transform = 'translate(-50%, -50%) scale(' + scale + ')';
        }

        window.addEventListener('resize', resize);
        window.addEventListener('load', resize);

        // Keyboard Controls
        window.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown' || e.key === 'Enter') {
                nextSlide();
            } else if (e.key === 'ArrowLeft' || e.key === 'Backspace' || e.key === 'PageUp') {
                prevSlide();
            } else if (e.key === 'Home') {
                showSlide(0);
            } else if (e.key === 'End') {
                showSlide(slides.length - 1);
            } else if (e.key.toLowerCase() === 'f') {
                toggleFullscreen();
            }
        });

        // Fullscreen Mode
        function toggleFullscreen() {
            if (!document.fullscreenElement) {
                document.documentElement.requestFullscreen().catch(err => {
                    console.error('Error attempting to enable full-screen mode:', err.message);
                });
                document.getElementById('fullscreen-btn').innerHTML = '<i class="fa-solid fa-compress"></i>';
            } else {
                document.exitFullscreen();
                document.getElementById('fullscreen-btn').innerHTML = '<i class="fa-solid fa-expand"></i>';
            }
        }

        // Show/Hide controls overlay on mouse move
        window.addEventListener('mousemove', () => {
            const controls = document.getElementById('controls');
            controls.classList.remove('hidden');
            clearTimeout(controlsTimeout);
            controlsTimeout = setTimeout(() => {
                controls.classList.add('hidden');
            }, 3000);
        });

        // Initialize first slide background
        showSlide(0);
    </script>
</body>
</html>
`;

// Start http server
const server = http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(parentHtml);
});

const PORT = 3000;
server.listen(PORT, () => {
    const url = `http://localhost:${PORT}`;
    console.log(`\n======================================================`);
    console.log(`Slideshow loaded successfully (${slides.length} slides found).`);
    console.log(`Presentation is running at: ${url}`);
    console.log(`Press Ctrl+C to stop the server.`);
    console.log(`======================================================\n`);
    
    // Automatically open browser depending on OS
    const startCmd = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start' : 'xdg-open';
    exec(`${startCmd} ${url}`);
});
