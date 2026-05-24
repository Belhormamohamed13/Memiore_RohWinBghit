# RohWinBghit — Mémoire de Master & Présentation

<p align="center">
  <img src="Screen_mobile/logo_bus.png" width="80" alt="RohWinBghit Logo"/>
</p>

<p align="center">
  <b>Plateforme de covoiturage intelligente et sécurisée pour l'Algérie inter-wilayas</b><br/>
  Mémoire de fin d'études — Master 2 Génie Logiciel<br/>
  Université de Tlemcen · 2025–2026
</p>

<p align="center">
  <img src="https://img.shields.io/badge/LaTeX-XeLaTeX%20%2B%20Biber-blue?logo=latex" alt="LaTeX"/>
  <img src="https://img.shields.io/badge/Presentation-PPTX%20(pptxgenjs)-green" alt="PPTX"/>
  <img src="https://img.shields.io/badge/Diagrams-PlantUML-orange" alt="PlantUML"/>
  <img src="https://img.shields.io/badge/Node.js-v20-brightgreen?logo=node.js" alt="Node"/>
</p>

---

## 📁 Structure du projet

```
RohWinBghit_LaTeX_Complet/
│
├── main.tex                        # Fichier maître LaTeX (à compiler)
├── biblio.bib                      # Bibliographie IEEE
├── memoire_rohwinbghit.tex         # Version mémoire consolidée
├── annexes_standalone.tex          # Annexes autonomes
│
├── frontmatter/                    # Pages liminaires
│   ├── page_garde.tex
│   ├── dedicaces.tex
│   ├── remerciements.tex
│   ├── resumes.tex                 # Résumés AR / FR / EN
│   └── acronymes.tex
│
├── chapters/                       # Chapitres du mémoire
│   ├── introduction_generale.tex
│   ├── chap1_etat_art.tex
│   ├── chap2_conception.tex        # Modélisation UML complète
│   ├── chap3_realisation.tex       # Captures d'écran & implémentation
│   ├── chap4_business_plan.tex     # Business Model Canvas
│   ├── conclusion_generale.tex
│   └── annexes.tex
│
├── figures/                        # Diagrammes UML et images
├── assets/                         # Ressources graphiques
├── plantuml/                       # Sources PlantUML + output PNG
│   └── output/                     # Diagrammes exportés (PNG)
│       ├── fig_use_case.png
│       ├── fig_class_diagram.png
│       ├── fig_sequence_inscription.png
│       ├── fig_sequence_booking.png
│       └── fig_activity_main.png
│
├── Screen_mobile/                  # Captures d'écran de l'application mobile
│   ├── icons/                      # Icônes SVG (43 icônes)
│   ├── logo.png / logo_bus.png
│   ├── univ_logo.png
│   └── *.jpg                       # Écrans de l'app
│
├── Generate rohwinbghit.js         # Générateur PPTX (pptxgenjs)
├── RohWinBghit_Presentation.pptx   # Présentation générée (28 slides)
├── presentatione.html              # Présentation HTML interactive
└── package.json                    # Dépendances Node.js
```

---

## 🚀 Compilation du mémoire LaTeX

```bash
cd RohWinBghit_LaTeX_Complet

# Compilation complète (avec bibliographie et glossaire)
xelatex main.tex
biber main
makeglossaries main
xelatex main.tex
xelatex main.tex
```

> **Prérequis :** TeX Live ≥ 2023 (ou MacTeX), police **Amiri** (`brew install --cask font-amiri`)

---

## 📊 Génération de la présentation PPTX

```bash
npm install           # Installer les dépendances Node.js
node "Generate rohwinbghit.js"
# → Génère : RohWinBghit_Presentation.pptx
```

---

## 🧩 Diagrammes UML (PlantUML)

Les sources PlantUML se trouvent dans `plantuml/` et les PNG exportés dans `plantuml/output/` :

| Fichier | Description |
|--------|-------------|
| `fig_use_case.png` | Diagramme de Cas d'Utilisation global |
| `fig_class_diagram.png` | Diagramme de Classes du domaine (PostgreSQL 3NF) |
| `fig_sequence_inscription.png` | Séquence Inscription avec OTP |
| `fig_sequence_booking.png` | Séquence Réservation multi-paiements |
| `fig_activity_main.png` | Flux d'Activité complet (Conducteur/Passager) |

---

## 👥 Auteurs

| Nom | Rôle |
|-----|------|
| **AHMED BACHA Djamel Eddine** | Conception, Réalisation, Rédaction |
| **BELHORMA Sidi Mohammed Reduane** | Conception, Réalisation, Rédaction |

> Encadrant : **[Nom de l'encadrant]**  
> Établissement : Université De Tlemcen

---

## 📦 Packages LaTeX requis

`polyglossia`, `biblatex`, `glossaries-extra`, `tcolorbox`, `tabularx`, `titlesec`, `fontspec`, `geometry`, `hyperref`, `graphicx`, `xcolor`, `arabxetex`

---

## 📄 Licence

Ce projet est déposé à titre académique. Tous droits réservés © 2025–2026 AHMED BACHA & BELHORMA.
