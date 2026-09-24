🌿 CozySuite
The offline, distraction-free student operating system.

Traditional office suites are bloated with corporate features you never use, and modern web apps lock your notes behind mandatory internet connections and subscriptions. CozySuite is a beautiful, local-first productivity suite engineered entirely around the modern student workflow.

It combines word processing, spreadsheets, native PDF reading, and infinite whiteboarding into a single, cohesive desktop environment. Zero telemetry. Zero forced cloud sync. 100% offline.

⚠️ A Note from the Developer
This is my very first software development project! While I am incredibly proud of what CozySuite is becoming, it is still a major learning experience. Because of this, you will likely encounter bugs, quirky UI behaviors, or unoptimized features.

If you find something broken (or if you just have ideas on how to make it better), please feel free to open an issue or submit a pull request. Feedback, patience, and contributions are hugely appreciated as I continue to learn and improve the app!

✨ Features
🛠️ The Core Suite
📝 Docs: A clean, rich-text editor with Markdown shortcuts (#, -), real-time word counting, and a "Clean Paste" utility that automatically strips web formatting.

📊 Sheets: A fully functional spreadsheet grid engine supporting standard arithmetic formulas and local CSV/XLSX imports/exports.

📽️ Slides: A minimalist presentation deck builder with drag-and-drop image support and distraction-free presenting.

📑 PDF Reader: A dedicated drag-and-drop PDF canvas. No more juggling Adobe Acrobat and Word in separate windows.

🎨 Board: An infinite Excalidraw-powered whiteboard for sketching mind maps, diagramming concepts, and solving visual problems.

🧠 Study: An integrated flashcard system. Highlight text in your notes or PDFs and instantly send them to your study deck for spaced repetition.

🎓 Student Superpowers
Split-Screen Study Mode: Hit one button to instantly dock a PDF on the left and your notebook on the right.

STEM Ready: Native support for rendering beautiful LaTeX math equations (via KaTeX) and syntax-highlighted code blocks (via Highlight.js).

Time Machine (Local Version History): Accidentally deleted your essay? CozySuite takes a silent local snapshot every 5 minutes. Use the timeline slider to restore previous versions instantly.

Focus Environment: A built-in Pomodoro timer in the status bar and an ambient lo-fi soundscape player (Rain, Cafe, White Noise) to keep you locked in.

♿ Accessibility First
Voice Typing: Hands-free dictation using the native Web Speech API.

Read Aloud: Built-in text-to-speech for proofreading essays or resting your eyes.

Neurodivergent Typography: Native integration of Lexend and OpenDyslexic fonts designed to reduce reading errors and eye strain.

High-Contrast Dark Mode: A carefully curated pastel dark mode palette that prevents text haloing while being gentle on the eyes late at night.

💻 Tech Stack
Frontend UI: React, Vite, Tailwind CSS

Desktop Wrapper: Tauri / Electron

Rich Text Engine: Quill.js

Spreadsheet Engine: x-data-spreadsheet

PDF Parsing: pdfjs-dist

Whiteboard: @excalidraw/excalidraw

🚀 Installation & Setup
Prerequisites
Make sure you have Node.js installed on your machine. If compiling for Windows via Tauri, ensure you have the Rust toolchain and MSVC build tools installed.

Local Development
Clone the repository:

Bash
git clone https://github.com/yourusername/cozysuite.git
cd cozysuite
Install all dependencies:

Bash
npm install
Run the local development server:

Bash
npm run dev
The app will be available in your browser at http://localhost:5173.

Building the Desktop Executable (.exe)
To package CozySuite into a standalone, installable Windows application:

Bash
# If using Electron Builder
npm run build
npx electron-builder

# If using Tauri
npm run tauri build
Once the compilation finishes, navigate to the dist/ or src-tauri/target/release/ directory to find your .exe setup file.

🔒 Privacy & Data
CozySuite respects your data. Everything you write, calculate, or draw is saved strictly to your local machine's storage. There are no tracking scripts, no telemetry, and no accounts required.

Created by Zaid Khan. Built during B.Tech Computer Science and Engineering studies at Pranveer Singh Institute of Technology.
