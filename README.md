# Personal Developer Portfolio — ITST 301 Final Project

[![HTML5](https://img.shields.io/badge/HTML5-Semantic%20Markup-E34F26?style=for-the-badge&logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Glossary/HTML5)
[![CSS3](https://img.shields.io/badge/CSS3-Modern%20Design-1572B6?style=for-the-badge&logo=css3&logoColor=white)](https://www.w3.org/Style/CSS/)
[![Bootstrap 5](https://img.shields.io/badge/Bootstrap-5.3.3-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)](https://getbootstrap.com/)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![GitHub Pages](https://img.shields.io/badge/Deployment-GitHub%20Pages-22c55e?style=for-the-badge&logo=githubpages&logoColor=white)](https://hiddengemofbell.github.io/basic-portfolio/)

---

## 📌 Project & Academic Information

| Item | Details |
| :--- | :--- |
| **Student Developer** | **Maribel Bravo** |
| **Student ID Number** | **0224-0865** |
| **Course & Section** | **BSIT 3A WMAD** (Web & Mobile Application Development) |
| **Institution** | **Laguna State Polytechnic University** (Siniloan Campus) |
| **College** | **College of Computer Studies** |
| **Course Code & Title**| **ITST 301 Principles of Web Design** / **ITEL 304 • IPT2 Final Project** |
| **Course Instructor** | **Prof. Julius A. Derla** |
| **GitHub Repository** | [https://github.com/hiddengemofbell/basic-portfolio](https://github.com/hiddengemofbell/basic-portfolio) |
| **Live Deployed Site**| [https://hiddengemofbell.github.io/basic-portfolio/](https://hiddengemofbell.github.io/basic-portfolio/) |
| **Figma Design File** | [Figma Project Board Link](https://www.figma.com/@hiddengemofbell) |

---

## 📖 Executive Summary & Website Purpose

This personal developer portfolio serves as the culminating performance-based assessment for **ITST 301 (Principles of Web Design)**. The website highlights the developer's background, technical competencies, core software projects, certifications, and academic progress in Web and Mobile Application Development.

The design philosophy balances a futuristic dark aesthetic with strict adherence to **W3C Semantic HTML5 standards**, **responsive grid layouts (Bootstrap 5.3.3 & CSS Custom Properties)**, **WCAG 2.1 accessibility benchmarks**, and **vanilla ES6+ JavaScript micro-interactions**.

---

## 🎨 Visual Identity & Design System

### 1. Color Palette

The site implements a coherent cyberpunk/deep-space dark theme with luminous neon accents:

| Token Name | HEX Code | Preview / Role | Applied Location |
| :--- | :--- | :--- | :--- |
| `--bg-primary` | `#060610` | ⬛ Deep Midnight Black | Page background, canvas base |
| `--bg-secondary`| `#0d0d1f` | 🌌 Obsidian Navy | Section background alternations |
| `--bg-card` | `rgba(18, 18, 38, 0.65)` | 🪟 Glassmorphic Surface | Cards, project tiles, form inputs |
| `--purple` | `#7c3aed` | 🟪 Royal Violet | Primary buttons, brand mark, gradient start |
| `--cyan` | `#06b6d4` | 🔷 Electric Cyan | Accent glow, focus rings, reading progress bar |
| `--text-primary`| `#f8fafc` | ⚪ Ghost White | Main headings, bold copy |
| `--text-secondary`|`#94a3b8` | 🔘 Slate Gray | Body text, captions, secondary tags |

### 2. Typography & Hierarchy
- **Primary Typeface**: `Inter` (Google Fonts: 300, 400, 500, 600, 700, 800, 900)
- **Fallback Stack**: `-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`
- **Heading Hierarchy**:
  - `h1`: Display title for personal name (hero section)
  - `h2`: Section titles (`.sec-title`) with glowing visual indicators
  - `h3`: Card titles (projects, certificates, highlights)

---

## 🛠️ System Architecture & File Structure

```
basic-portfolio/
├── assets/
│   ├── profile.jpg             # Optimized developer portrait (68 KB, web-ready)
│   ├── profile.JPG             # Original high-resolution photograph
│   ├── maribel-bravo-cv.pdf    # Direct CV download document
│   └── resume (2).pdf          # Backup resume document
├── css/
│   └── style.css               # Unified styling, CSS variables, keyframe animations
├── js/
│   └── main.js                 # Vanilla ES6+ interactivity & API handlers
├── index.html                  # Semantic HTML5 single-page application
└── README.md                   # Complete repository documentation & manual
```

---

## 🚀 Website Requirements & Implementation Evidence

### **WEB 1: Content and Semantic HTML5**
- **Semantic Structure**: Proper utilization of `<header>`, `<nav>`, `<main id="main">`, `<section>`, `<article>`, and `<footer>`.
- **Logical Flow**: Sequential heading hierarchy (`h1` → `h2` → `h3`) with no skipped heading ranks.
- **Sections Included**:
  1. Header with brand monogram and sticky navigation.
  2. Hero with introductory pitch, status indicator, and call-to-action buttons.
  3. About Me with verified academic credentials and specialty tags.
  4. Technical Skills categorized by Languages, Frontend, Tools, and AI environments.
  5. Featured Projects showcasing real-world applications.
  6. Certificates & Credentials verifying completed technical coursework.
  7. Contact Section with interactive messaging form and direct contact points.
  8. Footer with social links and university copyright metadata.

### **WEB 2: CSS, Bootstrap 5, and Responsive Layout**
- **Framework Integration**: Bootstrap 5.3.3 grid system, utility classes, and responsive breakpoints.
- **Custom Enhancements**: CSS Custom Properties (`:root`), glassmorphism backdrop filters (`backdrop-filter: blur(20px)`), and radial orb ambient gradients.
- **Responsive Verification**:
  - **Mobile (~375px)**: Single-column stack, collapsible hamburger menu, touch-friendly touch targets.
  - **Tablet (~768px)**: Balanced 2-column grids for project cards and highlights.
  - **Desktop (~1280px+)**: Multi-column layout with fixed max-widths and no unintended horizontal overflow.

### **WEB 3: JavaScript Interaction and Animation**
The site uses vanilla JavaScript without heavy third-party runtime overhead:
1. **Typewriter Effect**:
   - *Event*: `DOMContentLoaded` triggering looped recursive `setTimeout` updates.
   - *Result*: Smoothly cycles through role descriptions (*"Aspiring Developer"*, *"Problem Solver"*, *"Tech Enthusiast"*, *"Lifelong Learner"*).
2. **Dynamic Reading Progress Bar**:
   - *Event*: `window.addEventListener("scroll")`.
   - *Result*: Top-anchored `#scroll-progress` bar expands from 0% to 100% width relative to document scroll depth.
3. **ScrollSpy Navigation**:
   - *Event*: `IntersectionObserver` observing sections.
   - *Result*: Automatically adds `.active` styling to navbar links as sections scroll into viewport.
4. **Scroll Reveal Animations**:
   - *Event*: `IntersectionObserver` triggering on `.reveal` elements at 12% threshold.
   - *Result*: Cards and content smoothly fade up into view with hardware-accelerated CSS transitions.
5. **Asynchronous Contact Form**:
   - *Event*: `form.addEventListener("submit")`.
   - *Result*: Dispatches asynchronous `fetch()` POST request to Formspree with button loading state and live inline confirmation message.

### **WEB 4: Accessible Content and Controls**
- **Keyboard Navigation**: Skip-to-content anchor (`.skip-to-content`) allows keyboard-only users to bypass navigation directly to `#main`.
- **Visible Focus States**: High-contrast outline (`2px solid var(--cyan)`) applied to `:focus-visible` states across all interactive buttons, links, and inputs.
- **Alternative Text**: Meaningful descriptive `alt` tags on all informative images (e.g. `alt="Professional portrait of Maribel Bravo, Web Developer and IT Student"`).
- **Color Contrast**: Compliant text contrast ratios meeting WCAG 2.1 AA standards.

### **WEB 5: GitHub and Deployment**
- **Repository**: Public repository with clean commit messages, organized assets, and zero dead code.
- **Deployment**: Live hosted on GitHub Pages with HTTPS enabled.

### **WEB 6: Existing Project Samples**
1. **Cake Ordering & Bakery Platform**:
   - *Role*: Full-stack development lead.
   - *Tools*: Python (Flask), PostgreSQL, Tailwind CSS, Cloudinary.
   - *Contribution*: Database schema, authentication, shopping cart, and order pipeline.
2. **Real-Time Queue Ordering System**:
   - *Role*: Backend & frontend engineer.
   - *Tools*: Node.js, Server-Sent Events (SSE), HTML5, JavaScript.
   - *Contribution*: Concurrent queue processor streaming real-time status updates without client polling.
3. **EduFlow: Academic Task Manager (PWA)**:
   - *Role*: Frontend & API integration developer.
   - *Tools*: React, Vite, Tailwind CSS, Supabase (PostgreSQL).
   - *Contribution*: Offline LocalStorage caching, RESTful Supabase backend, and deadline tracking.

---

## 🧪 Acceptance Testing Matrix (T01 – T06)

| Test ID | Test Description & Action | Expected Result | Actual Result | Status |
| :---: | :--- | :--- | :--- | :---: |
| **T01** | Open deployed URL in private browser session without active logins. | Page loads completely with all styles, fonts, and assets appearing without authentication barriers. | Loaded instantly; all CDN and local assets resolved cleanly. | **PASS** |
| **T02** | Click through all navigation items, CTA buttons, and project links. | Each navigation link smoothly scrolls to its section; social and project links open proper destinations. | All anchor scrolls and external links functioned as labeled. | **PASS** |
| **T03** | Test responsive breakpoints at 375px (Mobile), 768px (Tablet), and 1280px (Desktop). | Layout re-flows seamlessly; no overlapping text or horizontal scrollbars at any resolution. | Responsive grid adapted cleanly across all tested viewport widths. | **PASS** |
| **T04** | Trigger JavaScript interactive components (Typewriter, ScrollSpy, Progress bar, Contact form). | Animations trigger smoothly and repeatedly without script errors in the developer console. | Zero console errors; animations and asynchronous form submit functioned reliably. | **PASS** |
| **T05** | Navigate entire interface using only `Tab`, `Shift+Tab`, and `Enter`. | Keyboard focus ring is visibly highlighted on all interactive elements; skip link functions. | Cyan outline indicator clearly visible; skip link successfully routed to main content. | **PASS** |
| **T06** | Inspect GitHub repository, file organization, Figma links, and documentation booklet. | Code matches submission guidelines; README is comprehensive; booklet formatted to spec. | All required documentation and file structures verified against ITST 301 criteria. | **PASS** |

---

## 💡 Learning Reflection

During the development of this personal portfolio, the primary technical challenge was achieving a high-end visual aesthetic (deep dark palette, glowing orb gradients, and glassmorphic card surfaces) while maintaining **strict WCAG 2.1 AA accessibility and sub-second page performance**. 

Initially, heavy media assets and unoptimized backdrop filters created rendering lag on simulated mobile devices. To address this, I resized and compressed the primary portrait image by over 99% (from 6.6 MB to 68 KB), implemented hardware-accelerated CSS transforms (`will-change`, `translateY`), and replaced intrusive focus suppressions with a dedicated `:focus-visible` electric-cyan outline. 

This taught me that great web design is never just about visual flair; true craftsmanship requires equal dedication to semantic structure, keyboard accessibility, and optimized performance.

---

## 📄 License & Attribution

- **Developed By**: Maribel Bravo (BSIT 3A WMAD)
- **Institution**: Laguna State Polytechnic University — Siniloan Campus
- **Icons**: [Bootstrap Icons](https://icons.getbootstrap.com/)
- **Fonts**: [Inter via Google Fonts](https://fonts.google.com/specimen/Inter)
- **Academic Submission**: ITST 301 Principles of Web Design Final Culminating Project (2026)
