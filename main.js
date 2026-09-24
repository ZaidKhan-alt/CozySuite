import Quill from 'quill';
import html2pdf from 'html2pdf.js';
import Spreadsheet from 'x-data-spreadsheet';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerSrc from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import mammoth from 'mammoth/mammoth.browser';
import 'x-data-spreadsheet/dist/xspreadsheet.css';

import React from 'react';
import { createRoot } from 'react-dom/client';
import { Excalidraw } from '@excalidraw/excalidraw';
import hljs from 'highlight.js';
import katex from 'katex';
import QuillBlotFormatter from 'quill-blot-formatter';

window.katex = katex;

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerSrc;

document.addEventListener('DOMContentLoaded', () => {
  // --- Theme Toggle Logic ---
  const themeToggleBtn = document.getElementById('theme-toggle');
  const themeIcon = document.getElementById('theme-icon');
  
  const savedTheme = localStorage.getItem('cozy-theme') || 'light';
  if (savedTheme === 'dark') {
    document.documentElement.classList.add('dark');
    themeIcon.classList.replace('ph-moon', 'ph-sun');
  }

  themeToggleBtn.addEventListener('click', () => {
    document.documentElement.classList.toggle('dark');
    const isDark = document.documentElement.classList.contains('dark');
    localStorage.setItem('cozy-theme', isDark ? 'dark' : 'light');
    if (isDark) {
      themeIcon.classList.replace('ph-moon', 'ph-sun');
    } else {
      themeIcon.classList.replace('ph-sun', 'ph-moon');
    }
  });

  // --- Status Bar UI Elements ---
  const statusLeftContent = document.getElementById('status-left-content');
  
  const updateStatusBar = (tabId) => {
    if (tabId === 'write') {
      statusLeftContent.innerHTML = `
        <div class="status-item" title="Word Count"><i class="ph ph-text-aa"></i> <span id="sb-words">0</span> words</div>
        <div class="status-divider"></div>
        <div class="status-item" title="Character Count"><i class="ph ph-character-recognition"></i> <span id="sb-chars">0</span> chars</div>
        <div class="status-divider"></div>
        <div class="status-item" title="Estimated Reading Time"><i class="ph ph-clock"></i> <span id="sb-read">0</span> min read</div>
      `;
      updateWordCount();
    } else if (tabId === 'calculate') {
      statusLeftContent.innerHTML = `
        <div class="status-item"><i class="ph ph-table"></i> Cell: <span id="sb-cell">A1</span></div>
      `;
    } else if (tabId === 'present') {
      statusLeftContent.innerHTML = `
        <div class="status-item"><i class="ph ph-presentation-chart"></i> Slide <span id="sb-current-slide">1</span> of <span id="sb-total-slides">1</span></div>
      `;
      updateSlideCount();
    } else if (tabId === 'read') {
      statusLeftContent.innerHTML = `
        <div class="status-item"><i class="ph ph-file-pdf"></i> <span id="sb-pdf-name">No document loaded</span></div>
      `;
      updatePdfStatus();
    }
  };

  const updateGlobalTitleExtension = (tabId) => {
    const titleInput = document.getElementById('global-doc-title');
    let baseName = titleInput.value.replace(/\.(docx|xlsx|pptx|pdf|csv)$/i, '') || 'Untitled Document';
    let ext = '';
    if (tabId === 'write') ext = '.docx';
    else if (tabId === 'calculate') ext = '.xlsx';
    else if (tabId === 'present') ext = '.pptx';
    else if (tabId === 'read') ext = '.pdf';
    
    titleInput.value = baseName + ext;
  };

  // --- Tab Switching Logic ---
  const navItems = document.querySelectorAll('.top-tabs .tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');

  navItems.forEach(item => {
    item.addEventListener('click', () => {
      navItems.forEach(nav => nav.classList.remove('active'));
      tabContents.forEach(tab => tab.classList.remove('active'));

      item.classList.add('active');
      const tabId = item.getAttribute('data-tab');
      const targetContent = document.getElementById(`tab-${tabId}`);
      if (targetContent) targetContent.classList.add('active');
      
      updateStatusBar(tabId);
      updateGlobalTitleExtension(tabId);
      
      // Trigger resize for spreadsheet if it becomes visible
      if (tabId === 'calculate' && window.excelGrid) {
        // Trigger resize inside requestAnimationFrame
        requestAnimationFrame(() => {
          window.excelGrid.resize();
        });
      }
    });
  });

  // --- Global Actions & Dropdown ---
  const fileMenuBtn = document.getElementById('file-menu-btn');
  const fileDropdownMenu = document.getElementById('file-dropdown-menu');
  if (fileMenuBtn && fileDropdownMenu) {
    fileMenuBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      fileDropdownMenu.style.display = fileDropdownMenu.style.display === 'none' ? 'flex' : 'none';
    });
    document.addEventListener('click', (e) => {
      if (!fileMenuBtn.contains(e.target) && !fileDropdownMenu.contains(e.target)) {
        fileDropdownMenu.style.display = 'none';
      }
    });
  }
  document.getElementById('global-save-btn').addEventListener('click', () => {
    // Mock save action
    const badge = document.querySelector('.save-status');
    const originalHtml = badge.innerHTML;
    badge.innerHTML = `<i class="ph ph-spinner ph-spin"></i> Saving...`;
    setTimeout(() => {
      badge.innerHTML = originalHtml;
    }, 1000);
  });

  // Export dynamically based on active tab
  document.getElementById('global-export-btn').addEventListener('click', () => {
    const activeTab = document.querySelector('.top-tabs .tab-btn.active').getAttribute('data-tab');
    if (activeTab === 'write') {
      const element = document.createElement('div');
      element.innerHTML = quill.root.innerHTML;
      const title = document.getElementById('global-doc-title').value || 'Document';
      html2pdf().set({
        margin: 1, filename: `${title}.pdf`, image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
      }).from(element).save();
    } else if (activeTab === 'calculate') {
      document.getElementById('excel-export-xlsx-btn').click();
    } else {
      alert('Export is only supported for Docs and Sheets currently.');
    }
  });


  // --- DOCS TAB (Quill) ---
  const Font = Quill.import('formats/font');
  const fontWhitelist = ['nunito', 'arial', 'roboto', 'open-sans', 'times-new-roman', 'courier-new', 'lexend', 'opendyslexic'];
  Font.whitelist = fontWhitelist;
  Quill.register(Font, true);
  Quill.register('modules/blotFormatter', QuillBlotFormatter);

  const quill = new Quill('#editor-container', {
    theme: 'snow',
    modules: {
      toolbar: [
        [{ 'font': fontWhitelist }],
        [{ 'header': [1, 2, 3, false] }],
        ['bold', 'italic', 'underline', 'strike'],
        [{ 'list': 'ordered'}, { 'list': 'bullet' }],
        ['image', 'code-block'],
        ['clean']
      ],
      syntax: { hljs },
      blotFormatter: {}
    },
    placeholder: 'Start typing something cozy...'
  });

  // Move Quill toolbar directly beneath the custom secondary toolbar
  const quillToolbar = document.querySelector('#tab-write .ql-toolbar');
  const docsTab = document.getElementById('tab-write');
  const scrollArea = document.getElementById('docs-scroll-area');
  if (quillToolbar && docsTab && scrollArea) {
    docsTab.insertBefore(quillToolbar, scrollArea);
  }

  const updateWordCount = () => {
    const text = quill.getText().trim();
    const chars = text.length;
    const words = chars > 0 ? text.split(/\\s+/).length : 0;
    const readTime = Math.ceil(words / 200);
    
    const wordsEl = document.getElementById('sb-words');
    const charsEl = document.getElementById('sb-chars');
    const readEl = document.getElementById('sb-read');
    
    if (wordsEl) wordsEl.textContent = words;
    if (charsEl) charsEl.textContent = chars;
    if (readEl) readEl.textContent = readTime;
  };

  quill.on('text-change', (delta, oldDelta, source) => {
    updateWordCount();
    if (source === 'user') {
      const selection = quill.getSelection();
      if (!selection) return;
      const [line, offset] = quill.getLine(selection.index);
      if (!line) return;
      const text = line.domNode.textContent;
      
      // Auto-format markdown shortcuts
      if (text === '# ' && selection.index >= 2) {
        quill.deleteText(selection.index - 2, 2);
        quill.formatLine(selection.index - 2, 1, 'header', 1);
      } else if (text === '## ' && selection.index >= 3) {
        quill.deleteText(selection.index - 3, 3);
        quill.formatLine(selection.index - 3, 1, 'header', 2);
      } else if (text === '- ' && selection.index >= 2) {
        quill.deleteText(selection.index - 2, 2);
        quill.formatLine(selection.index - 2, 1, 'list', 'bullet');
      }
    }
  });

  // Clean Paste Utility
  let cleanPasteEnabled = false;
  const cleanPasteBtn = document.getElementById('clean-paste-toggle');
  cleanPasteBtn.addEventListener('click', () => {
    cleanPasteEnabled = !cleanPasteEnabled;
    cleanPasteBtn.title = `Clean Paste: ${cleanPasteEnabled ? 'On' : 'Off'}`;
    if (cleanPasteEnabled) {
      cleanPasteBtn.classList.add('primary');
      cleanPasteBtn.classList.remove('secondary');
    } else {
      cleanPasteBtn.classList.remove('primary');
      cleanPasteBtn.classList.add('secondary');
    }
  });
  
  quill.root.addEventListener('paste', (e) => {
    if (cleanPasteEnabled) {
      e.preventDefault();
      const text = (e.originalEvent || e).clipboardData.getData('text/plain');
      document.execCommand('insertText', false, text);
    }
  });

  // Voice Typing (Dictation)
  const dictationBtn = document.getElementById('dictation-btn');
  let recognition = null;
  let isRecording = false;
  if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = false;
    
    recognition.onresult = (event) => {
      let finalTranscript = '';
      for (let i = event.resultIndex; i < event.results.length; ++i) {
        if (event.results[i].isFinal) {
          finalTranscript += event.results[i][0].transcript;
        }
      }
      if (finalTranscript) {
        const selection = quill.getSelection(true);
        quill.insertText(selection.index, finalTranscript + ' ');
        quill.setSelection(selection.index + finalTranscript.length + 1);
      }
    };
    
    recognition.onend = () => {
      isRecording = false;
      dictationBtn.classList.remove('recording');
    };
  }
  
  dictationBtn.addEventListener('click', () => {
    if (!recognition) {
      alert('Speech Recognition is not supported in this browser.');
      return;
    }
    if (isRecording) {
      recognition.stop();
    } else {
      recognition.start();
      isRecording = true;
      dictationBtn.classList.add('recording');
    }
  });

  // Read Aloud (TTS)
  const ttsBtn = document.getElementById('tts-btn');
  let isPlayingTTS = false;
  
  ttsBtn.addEventListener('click', () => {
    if (isPlayingTTS) {
      window.speechSynthesis.cancel();
      isPlayingTTS = false;
      ttsBtn.classList.remove('playing');
      return;
    }
    
    let textToRead = '';
    const selection = quill.getSelection();
    if (selection && selection.length > 0) {
      textToRead = quill.getText(selection.index, selection.length);
    } else {
      // Read current paragraph or all if none
      const [line, offset] = quill.getLine(selection ? selection.index : 0);
      if (line) textToRead = line.domNode.textContent;
      if (!textToRead) textToRead = quill.getText();
    }
    
    if (textToRead.trim()) {
      const utterance = new SpeechSynthesisUtterance(textToRead);
      utterance.onend = () => {
        isPlayingTTS = false;
        ttsBtn.classList.remove('playing');
      };
      window.speechSynthesis.speak(utterance);
      isPlayingTTS = true;
      ttsBtn.classList.add('playing');
    }
  });

  // Line Spacing
  const lineHeights = {
    '1.0': 'lh-1-0',
    '1.15': 'lh-1-15',
    '1.5': 'lh-1-5',
    '2.0': 'lh-2-0'
  };
  document.querySelector('.ql-lineHeight').addEventListener('change', (e) => {
    const val = e.target.value;
    const editor = document.querySelector('.ql-editor');
    // Remove existing line height classes
    Object.values(lineHeights).forEach(cls => editor.classList.remove(cls));
    // Add new one
    if (lineHeights[val]) editor.classList.add(lineHeights[val]);
  });

  // Insert Table Modal
  const tableModal = document.getElementById('insert-table-modal');
  document.getElementById('docs-insert-table-btn').addEventListener('click', () => {
    tableModal.style.display = 'flex';
  });
  document.getElementById('table-cancel').addEventListener('click', () => {
    tableModal.style.display = 'none';
  });
  document.getElementById('table-insert-confirm').addEventListener('click', () => {
    const rows = parseInt(document.getElementById('table-rows-input').value, 10);
    const cols = parseInt(document.getElementById('table-cols-input').value, 10);
    
    if (rows > 0 && cols > 0) {
      let html = '<table style="width: 100%; border-collapse: collapse;"><tbody>';
      for (let r = 0; r < rows; r++) {
        html += '<tr>';
        for (let c = 0; c < cols; c++) {
          html += '<td style="border: 1px solid var(--border-soft); padding: 8px;"><br></td>';
        }
        html += '</tr>';
      }
      html += '</tbody></table><p><br></p>';
      
      const selection = quill.getSelection(true);
      quill.clipboard.dangerouslyPasteHTML(selection.index, html);
    }
    tableModal.style.display = 'none';
  });

  // Time Machine (Auto-save & History)
  const HISTORY_KEY = 'cozy-doc-history';
  setInterval(() => {
    let history = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
    history.push({ timestamp: Date.now(), content: quill.root.innerHTML });
    if (history.length > 20) history.shift();
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
  }, 5 * 60 * 1000);

  const historyToggleBtn = document.getElementById('history-toggle-btn');
  const historyDrawer = document.getElementById('time-machine-drawer');
  const historySlider = document.getElementById('history-slider');
  const historyTimestampDisplay = document.getElementById('history-timestamp');
  
  let currentHistory = [];

  historyToggleBtn.addEventListener('click', () => {
    currentHistory = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
    if (currentHistory.length === 0) {
      alert("No history available yet.");
      return;
    }
    historyDrawer.style.display = 'flex';
    setTimeout(() => historyDrawer.style.transform = 'translateY(0)', 10);
    
    historySlider.max = currentHistory.length - 1;
    historySlider.value = currentHistory.length - 1;
    historyTimestampDisplay.textContent = new Date(currentHistory[currentHistory.length - 1].timestamp).toLocaleString();
  });

  historySlider.addEventListener('input', (e) => {
    const val = parseInt(e.target.value, 10);
    if (currentHistory[val]) {
      quill.root.innerHTML = currentHistory[val].content;
      historyTimestampDisplay.textContent = new Date(currentHistory[val].timestamp).toLocaleString();
    }
  });

  document.getElementById('restore-history-btn').addEventListener('click', () => {
    historyDrawer.style.transform = 'translateY(100%)';
    setTimeout(() => historyDrawer.style.display = 'none', 300);
  });
  document.getElementById('close-history-btn').addEventListener('click', () => {
    historyDrawer.style.transform = 'translateY(100%)';
    setTimeout(() => historyDrawer.style.display = 'none', 300);
  });


  // --- SHEETS TAB (Spreadsheet) ---
  const excelGrid = new Spreadsheet('#spreadsheet-container', {
    view: {
      height: () => {
        const container = document.getElementById('spreadsheet-container');
        return container ? container.clientHeight : 500;
      },
      width: () => {
        const container = document.getElementById('spreadsheet-container');
        return container ? container.clientWidth : 800;
      }
    }
  }).loadData({});
  window.excelGrid = excelGrid;

  // Trigger initial resize using requestAnimationFrame for proper bounding rect calculations
  setTimeout(() => {
    requestAnimationFrame(() => {
      excelGrid.resize();
    });
  }, 150);

  // Listen for cell selection
  excelGrid.on('cell-selected', (cell, ri, ci) => {
    const cellEl = document.getElementById('sb-cell');
    if (cellEl) {
      // Convert ci to letter (0 -> A, 1 -> B)
      const colLetter = String.fromCharCode(65 + ci);
      cellEl.textContent = `${colLetter}${ri + 1}`;
    }
  });

  const stox = (wb) => {
    const out = [];
    wb.SheetNames.forEach(function (name) {
      const o = { name: name, rows: {} };
      const ws = wb.Sheets[name];
      if(!ws || !ws['!ref']) return;
      const range = XLSX.utils.decode_range(ws['!ref']);
      let rows = { len: range.e.r + 1 };
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const row = { cells: {} };
        let hasContent = false;
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = ws[XLSX.utils.encode_cell({ c: C, r: R })];
          if (cell && cell.t !== 'z') {
            row.cells[C] = { text: String(cell.v) };
            hasContent = true;
          }
        }
        if (hasContent) rows[R] = row;
      }
      o.rows = rows;
      out.push(o);
    });
    return out;
  };

  const xtos = (sdata) => {
    const out = XLSX.utils.book_new();
    sdata.forEach(function (xws) {
      const aoa = [[]];
      const rowobj = xws.rows;
      for (let ri = 0; ri < rowobj.len; ++ri) {
        const row = rowobj[ri];
        if (!row) continue;
        aoa[ri] = [];
        Object.keys(row.cells).forEach(function (k) {
          const idx = +k;
          if (isNaN(idx)) return;
          aoa[ri][idx] = row.cells[k].text;
        });
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(out, ws, xws.name);
    });
    return out;
  };

  document.getElementById('excel-import-btn').addEventListener('click', () => {
    document.getElementById('excel-import-input').click();
  });

  document.getElementById('excel-import-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = e.target.result;
      const wb = XLSX.read(data, {type: 'binary'});
      excelGrid.loadData(stox(wb));
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; // reset
  });

  document.getElementById('excel-export-csv-btn').addEventListener('click', () => {
    const title = document.getElementById('global-doc-title').value || 'Spreadsheet';
    const wb = xtos(excelGrid.getData());
    XLSX.writeFile(wb, `${title}.csv`, { bookType: 'csv' });
  });

  document.getElementById('excel-export-xlsx-btn').addEventListener('click', () => {
    const title = document.getElementById('global-doc-title').value || 'Spreadsheet';
    const wb = xtos(excelGrid.getData());
    XLSX.writeFile(wb, `${title}.xlsx`);
  });


  // --- SLIDES TAB ---
  let slides = [];
  let activeSlideIndex = -1;

  const titleEl = document.querySelector('.slide-title');
  const subtitleEl = document.querySelector('.slide-subtitle');
  const bulletsEl = document.querySelector('.slide-bullets');
  const bgColorEl = document.getElementById('slide-bg-color');
  const thumbnailsContainer = document.getElementById('slides-thumbnails');
  const activeSlideContainer = document.getElementById('active-slide');

  const updateSlideCount = () => {
    const currEl = document.getElementById('sb-current-slide');
    const totEl = document.getElementById('sb-total-slides');
    if (currEl) currEl.textContent = activeSlideIndex + 1;
    if (totEl) totEl.textContent = slides.length || 1;
  };

  const updateSlideState = () => {
    if (activeSlideIndex === -1 || !slides[activeSlideIndex]) return;
    
    // Save images and links state if necessary, but we can just save innerHTML of the canvas
    // Wait, the slide bg is separate. The content is inside activeSlideContainer.
    slides[activeSlideIndex] = {
      title: titleEl.innerHTML,
      subtitle: subtitleEl.innerHTML,
      bullets: bulletsEl.innerHTML,
      bg: bgColorEl.value,
      extraHtml: Array.from(activeSlideContainer.querySelectorAll('.slide-image-wrapper')).map(el => el.outerHTML).join('')
    };
    renderThumbnails();
  };

  titleEl.addEventListener('input', updateSlideState);
  subtitleEl.addEventListener('input', updateSlideState);
  bulletsEl.addEventListener('input', updateSlideState);
  bgColorEl.addEventListener('input', (e) => {
    activeSlideContainer.style.backgroundColor = e.target.value;
    updateSlideState();
  });

  const renderThumbnails = () => {
    thumbnailsContainer.innerHTML = '';
    slides.forEach((slide, index) => {
      const wrapper = document.createElement('div');
      wrapper.className = 'slide-thumbnail-wrapper';
      
      const thumb = document.createElement('div');
      thumb.className = `slide-thumbnail ${index === activeSlideIndex ? 'active' : ''}`;
      thumb.style.backgroundColor = slide.bg;
      
      const content = document.createElement('div');
      content.className = 'slide-thumbnail-content';
      content.innerHTML = `<h1>${slide.title || 'Slide ' + (index+1)}</h1>`;
      thumb.appendChild(content);

      const delBtn = document.createElement('button');
      delBtn.className = 'delete-slide-btn';
      delBtn.innerHTML = '<i class="ph ph-x"></i>';
      delBtn.onclick = (e) => {
        e.stopPropagation();
        deleteSlide(index);
      };

      thumb.onclick = () => selectSlide(index);
      
      wrapper.appendChild(thumb);
      wrapper.appendChild(delBtn);
      thumbnailsContainer.appendChild(wrapper);
    });
  };

  const selectSlide = (index) => {
    if (index < 0 || index >= slides.length) {
      activeSlideIndex = -1;
      activeSlideContainer.style.display = 'none';
      updateSlideCount();
      return;
    }
    activeSlideIndex = index;
    const slide = slides[index];
    titleEl.innerHTML = slide.title || '';
    subtitleEl.innerHTML = slide.subtitle || '';
    bulletsEl.innerHTML = slide.bullets || '';
    bgColorEl.value = slide.bg || '#ffffff';
    activeSlideContainer.style.backgroundColor = slide.bg || '#ffffff';
    
    // Restore images
    const existingImages = activeSlideContainer.querySelectorAll('.slide-image-wrapper');
    existingImages.forEach(img => img.remove());
    if (slide.extraHtml) {
      activeSlideContainer.insertAdjacentHTML('beforeend', slide.extraHtml);
      // Reattach drag/resize events
      activeSlideContainer.querySelectorAll('.slide-image-wrapper').forEach(attachImageEvents);
    }
    
    activeSlideContainer.style.display = 'flex';
    renderThumbnails();
    updateSlideCount();
  };

  const addSlide = () => {
    slides.push({ title: '', subtitle: '', bullets: '<li></li>', bg: '#ffffff', extraHtml: '' });
    selectSlide(slides.length - 1);
  };

  const deleteSlide = (index) => {
    slides.splice(index, 1);
    if (slides.length === 0) {
      addSlide();
    } else {
      selectSlide(Math.max(0, index - 1));
    }
  };

  document.getElementById('slides-add-btn').addEventListener('click', addSlide);
  if (slides.length === 0) addSlide();

  // Presentation Mode
  const presentationContainer = document.getElementById('presentation-container');
  const presentationSlide = document.getElementById('presentation-slide');
  let currentPresentIndex = 0;

  const renderPresentationSlide = () => {
    if (!slides[currentPresentIndex]) return;
    const slide = slides[currentPresentIndex];
    presentationSlide.innerHTML = `
      <div class="slide-canvas" style="background-color: ${slide.bg || '#ffffff'}; width: 80vw; height: 45vw; max-width: 1200px; max-height: 675px; position: relative; border-radius: var(--radius-md);">
        <div class="slide-content" style="color: #000; z-index: 1;">
          <h1 class="slide-title" style="font-size: 4rem;">${slide.title}</h1>
          <h2 class="slide-subtitle" style="font-size: 2rem; color: #444;">${slide.subtitle}</h2>
          <ul class="slide-bullets" style="font-size: 1.5rem; width: 80%; list-style-position: inside;">${slide.bullets}</ul>
        </div>
        ${slide.extraHtml || ''}
      </div>
    `;
    
    // Add link handlers in presentation
    presentationSlide.querySelectorAll('.slide-link').forEach(link => {
      link.addEventListener('click', (e) => {
        e.stopPropagation();
        const target = link.getAttribute('data-target');
        if (target.startsWith('#')) {
          const slideNum = parseInt(target.replace('#', ''), 10);
          if (!isNaN(slideNum) && slideNum > 0 && slideNum <= slides.length) {
            currentPresentIndex = slideNum - 1;
            renderPresentationSlide();
          }
        } else {
          let url = target;
          if (!url.startsWith('http://') && !url.startsWith('https://')) {
            url = 'https://' + url;
          }
          window.open(url, '_blank');
        }
      });
    });
  };

  document.getElementById('slides-present-btn').addEventListener('click', async () => {
    currentPresentIndex = Math.max(0, activeSlideIndex);
    renderPresentationSlide();
    presentationContainer.style.display = 'flex';
    if (presentationContainer.requestFullscreen) {
      await presentationContainer.requestFullscreen();
    }
  });

  document.addEventListener('keydown', (e) => {
    if (presentationContainer.style.display === 'flex') {
      if (e.key === 'ArrowRight' || e.key === 'Space') {
        if (currentPresentIndex < slides.length - 1) {
          currentPresentIndex++;
          renderPresentationSlide();
        }
      } else if (e.key === 'ArrowLeft') {
        if (currentPresentIndex > 0) {
          currentPresentIndex--;
          renderPresentationSlide();
        }
      } else if (e.key === 'Escape') {
        presentationContainer.style.display = 'none';
        if (document.fullscreenElement) {
          document.exitFullscreen();
        }
      }
    }
  });
  
  document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement) {
      presentationContainer.style.display = 'none';
    }
  });

  // --- SLIDES IMAGE DRAG/DROP/RESIZE ---
  let isDraggingImg = false;
  let currentImg = null;
  let startX, startY, initialLeft, initialTop;
  let isResizing = false;
  let initialWidth, initialHeight;

  const attachImageEvents = (wrapper) => {
    wrapper.addEventListener('mousedown', (e) => {
      if (e.target.classList.contains('resize-handle')) {
        isResizing = true;
        currentImg = wrapper;
        startX = e.clientX;
        startY = e.clientY;
        initialWidth = wrapper.offsetWidth;
        initialHeight = wrapper.offsetHeight;
        e.stopPropagation();
      } else {
        isDraggingImg = true;
        currentImg = wrapper;
        startX = e.clientX;
        startY = e.clientY;
        initialLeft = wrapper.offsetLeft;
        initialTop = wrapper.offsetTop;
        
        // Deselect others
        activeSlideContainer.querySelectorAll('.slide-image-wrapper').forEach(el => el.classList.remove('selected'));
        wrapper.classList.add('selected');
      }
    });
  };

  document.addEventListener('mousemove', (e) => {
    if (isDraggingImg && currentImg) {
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      // Convert to percentages for responsive scaling
      const containerRect = activeSlideContainer.getBoundingClientRect();
      let newLeft = ((initialLeft + dx) / containerRect.width) * 100;
      let newTop = ((initialTop + dy) / containerRect.height) * 100;
      currentImg.style.left = `${newLeft}%`;
      currentImg.style.top = `${newTop}%`;
    } else if (isResizing && currentImg) {
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const containerRect = activeSlideContainer.getBoundingClientRect();
      let newWidth = ((initialWidth + dx) / containerRect.width) * 100;
      let newHeight = ((initialHeight + dy) / containerRect.height) * 100;
      currentImg.style.width = `${newWidth}%`;
      currentImg.style.height = `${newHeight}%`;
    }
  });

  document.addEventListener('mouseup', () => {
    if (isDraggingImg || isResizing) {
      isDraggingImg = false;
      isResizing = false;
      currentImg = null;
      updateSlideState();
    }
  });
  
  // Deselect on click outside
  activeSlideContainer.addEventListener('mousedown', (e) => {
    if (e.target === activeSlideContainer || e.target.classList.contains('slide-content')) {
      activeSlideContainer.querySelectorAll('.slide-image-wrapper').forEach(el => el.classList.remove('selected'));
    }
  });

  const addImageToSlide = (dataUrl) => {
    if (activeSlideIndex === -1) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'slide-image-wrapper';
    wrapper.style.left = '10%';
    wrapper.style.top = '10%';
    wrapper.style.width = '30%';
    wrapper.style.height = '30%';
    wrapper.style.zIndex = '2';
    
    wrapper.innerHTML = `
      <img src="${dataUrl}" />
      <div class="resize-handle br"></div>
    `;
    
    activeSlideContainer.appendChild(wrapper);
    attachImageEvents(wrapper);
    updateSlideState();
  };

  document.getElementById('slides-insert-img-btn').addEventListener('click', () => {
    document.getElementById('slides-img-input').click();
  });

  document.getElementById('slides-img-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = (evt) => addImageToSlide(evt.target.result);
      reader.readAsDataURL(file);
    }
    e.target.value = '';
  });

  activeSlideContainer.addEventListener('dragover', (e) => e.preventDefault());
  activeSlideContainer.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0] && e.dataTransfer.files[0].type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = (evt) => addImageToSlide(evt.target.result);
      reader.readAsDataURL(e.dataTransfer.files[0]);
    }
  });

  document.addEventListener('paste', (e) => {
    // Only paste image if in present tab and focus is on canvas (not editing text directly)
    if (document.getElementById('tab-present').classList.contains('active')) {
      const items = (e.clipboardData || e.originalEvent.clipboardData).items;
      for (let index in items) {
        const item = items[index];
        if (item.kind === 'file' && item.type.startsWith('image/')) {
          const blob = item.getAsFile();
          const reader = new FileReader();
          reader.onload = (evt) => addImageToSlide(evt.target.result);
          reader.readAsDataURL(blob);
          e.preventDefault();
        }
      }
    }
  });

  // --- SLIDES LINK SUPPORT ---
  const linkModal = document.getElementById('slide-link-modal');
  const linkUrlInput = document.getElementById('slide-link-url');
  let savedSelectionRange = null;

  document.getElementById('slides-insert-link-btn').addEventListener('click', () => {
    const selection = window.getSelection();
    if (selection.rangeCount > 0 && !selection.isCollapsed) {
      savedSelectionRange = selection.getRangeAt(0);
      // Ensure selection is inside active slide
      if (activeSlideContainer.contains(savedSelectionRange.commonAncestorContainer)) {
        linkModal.style.display = 'flex';
        linkUrlInput.value = '';
        linkUrlInput.focus();
      } else {
        alert('Please select text inside the active slide.');
      }
    } else {
      alert('Please select some text to link.');
    }
  });

  document.getElementById('slide-link-cancel').addEventListener('click', () => {
    linkModal.style.display = 'none';
  });

  document.getElementById('slide-link-save').addEventListener('click', () => {
    const url = linkUrlInput.value.trim();
    if (url && savedSelectionRange) {
      const a = document.createElement('span');
      a.className = 'slide-link';
      a.setAttribute('data-target', url);
      a.appendChild(savedSelectionRange.extractContents());
      savedSelectionRange.insertNode(a);
      updateSlideState();
    }
    linkModal.style.display = 'none';
  });




  // --- PDF TAB ---
  let pdfDoc = null;
  let pageNum = 1;
  let pageRendering = false;
  let pageNumPending = null;
  let pdfScale = 1.0;
  let loadedFileName = '';
  const canvas = document.getElementById('pdf-canvas');
  const ctx = canvas.getContext('2d');
  
  const dropzone = document.getElementById('pdf-dropzone');
  const viewerContainer = document.getElementById('pdf-viewer-container');
  const toolbar = document.getElementById('pdf-toolbar');

  const updatePdfStatus = () => {
    const el = document.getElementById('sb-pdf-name');
    if (el) el.textContent = loadedFileName || 'No document loaded';
  };

  const renderPage = (num) => {
    pageRendering = true;
    pdfDoc.getPage(num).then((page) => {
      const viewport = page.getViewport({ scale: pdfScale });
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      const renderContext = { canvasContext: ctx, viewport: viewport };
      const renderTask = page.render(renderContext);
      
      renderTask.promise.then(() => {
        pageRendering = false;
        if (pageNumPending !== null) {
          renderPage(pageNumPending);
          pageNumPending = null;
        }
      });
    });
  };

  const queueRenderPage = (num) => {
    if (pageRendering) {
      pageNumPending = num;
    } else {
      renderPage(num);
    }
  };

  const onPrevPage = () => {
    if (pageNum <= 1) return;
    pageNum--;
    queueRenderPage(pageNum);
  };

  const onNextPage = () => {
    if (pageNum >= pdfDoc.numPages) return;
    pageNum++;
    queueRenderPage(pageNum);
  };

  document.getElementById('pdf-prev-btn').addEventListener('click', onPrevPage);
  document.getElementById('pdf-next-btn').addEventListener('click', onNextPage);
  
  document.getElementById('pdf-zoom-in-btn').addEventListener('click', () => {
    pdfScale += 0.2;
    queueRenderPage(pageNum);
  });
  
  document.getElementById('pdf-zoom-out-btn').addEventListener('click', () => {
    if (pdfScale <= 0.4) return;
    pdfScale -= 0.2;
    queueRenderPage(pageNum);
  });

  const loadPdf = (arrayBuffer, filename) => {
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    loadingTask.promise.then((pdf) => {
      pdfDoc = pdf;
      pageNum = 1;
      loadedFileName = filename;
      
      dropzone.style.display = 'none';
      viewerContainer.style.display = 'flex';
      toolbar.style.display = 'flex';
      
      renderPage(pageNum);
      
      // Update global title
      document.getElementById('global-doc-title').value = filename;
      
      const el = document.getElementById('sb-pdf-name');
      if (el) el.textContent = filename;
    }).catch(err => {
      console.error('Error loading PDF:', err);
      alert('Failed to load PDF file.');
    });
  };

  dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
  dropzone.addEventListener('dragleave', () => { dropzone.classList.remove('dragover'); });
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault(); dropzone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
      const reader = new FileReader();
      reader.onload = (evt) => loadPdf(new Uint8Array(evt.target.result), file.name);
      reader.readAsArrayBuffer(file);
    } else {
      alert('Please drop a valid PDF file.');
    }
  });

  document.getElementById('pdf-browse-btn').addEventListener('click', () => { document.getElementById('pdf-file-input').click(); });
  document.getElementById('pdf-file-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file && file.type === 'application/pdf') {
      const reader = new FileReader();
      reader.onload = (evt) => loadPdf(new Uint8Array(evt.target.result), file.name);
      reader.readAsArrayBuffer(file);
    }
  });

  // --- Window Controls ---
  const minimizeBtn = document.getElementById('minimize-btn');
  const maximizeBtn = document.getElementById('maximize-btn');
  const closeBtn = document.getElementById('close-btn');

  if (window.electronAPI) {
    if (minimizeBtn) minimizeBtn.addEventListener('click', () => window.electronAPI.minimize());
    if (maximizeBtn) maximizeBtn.addEventListener('click', () => window.electronAPI.maximize());
    if (closeBtn) closeBtn.addEventListener('click', () => window.electronAPI.close());
  }
  
  // --- FOCUS MODE ---
  const focusBtn = document.getElementById('focus-mode-btn');
  const focusExitBtn = document.getElementById('focus-exit-btn');
  
  focusBtn.addEventListener('click', () => {
    document.body.classList.add('focus-mode');
    focusExitBtn.style.display = 'flex';
  });
  
  focusExitBtn.addEventListener('click', () => {
    document.body.classList.remove('focus-mode');
    focusExitBtn.style.display = 'none';
  });
  
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && document.body.classList.contains('focus-mode')) {
      document.body.classList.remove('focus-mode');
      focusExitBtn.style.display = 'none';
    }
  });

  // --- KEYBOARD SHORTCUTS MODAL ---
  const shortcutsModal = document.getElementById('shortcuts-modal');
  const closeShortcutsBtn = document.getElementById('close-shortcuts-btn');
  closeShortcutsBtn.addEventListener('click', () => {
    shortcutsModal.style.display = 'none';
  });

  // --- GLOBAL KEYBOARD SHORTCUTS ---
  document.addEventListener('keydown', (e) => {
    // Focus Mode Exit
    if (e.key === 'Escape' && document.body.classList.contains('focus-mode')) {
      document.body.classList.remove('focus-mode');
      focusExitBtn.style.display = 'none';
    }
    
    // Shortcuts Modal
    if (e.key === '?' && e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA' && !document.activeElement.isContentEditable) {
      e.preventDefault();
      shortcutsModal.style.display = 'flex';
    }

    // Tabs navigation
    if (e.ctrlKey || e.metaKey) {
      if (e.key === '1') { e.preventDefault(); document.querySelector('.tab-btn[data-tab="write"]').click(); }
      if (e.key === '2') { e.preventDefault(); document.querySelector('.tab-btn[data-tab="calculate"]').click(); }
      if (e.key === '3') { e.preventDefault(); document.querySelector('.tab-btn[data-tab="present"]').click(); }
      if (e.key === '4') { e.preventDefault(); document.querySelector('.tab-btn[data-tab="read"]').click(); }
    }
  });

  // --- SPLIT VIEW ---
  const splitViewBtn = document.getElementById('split-view-btn');
  splitViewBtn.addEventListener('click', () => {
    document.body.classList.toggle('split-view-active');
    if (document.body.classList.contains('split-view-active')) {
      splitViewBtn.style.color = 'var(--text-main)';
      splitViewBtn.style.backgroundColor = 'var(--bg-ribbon)';
    } else {
      splitViewBtn.style.color = '';
      splitViewBtn.style.backgroundColor = '';
      // Retrigger active tab to clean up display styles from CSS
      const activeTab = document.querySelector('.top-tabs .tab-btn.active');
      if (activeTab) activeTab.click();
    }
  });

  // --- POMODORO TIMER ---
  const pomodoroBtn = document.getElementById('pomodoro-btn');
  const pomodoroTime = document.getElementById('pomodoro-time');
  const pomodoroContainer = document.getElementById('pomodoro-container');
  let pomodoroInterval = null;
  let pomodoroSeconds = 25 * 60; // 25 mins

  const updatePomodoroDisplay = () => {
    const mins = Math.floor(pomodoroSeconds / 60);
    const secs = pomodoroSeconds % 60;
    pomodoroTime.textContent = `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  };

  const playPomodoroChime = () => {
    // Simple beep using Web Audio API
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.type = 'bell';
      osc.frequency.setValueAtTime(880, ctx.currentTime);
      gain.gain.setValueAtTime(0.5, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 1);
      osc.start(ctx.currentTime);
      osc.stop(ctx.currentTime + 1);
    } catch (e) {
      console.error('AudioContext not supported', e);
    }
  };

  pomodoroBtn.addEventListener('click', () => {
    if (pomodoroInterval) {
      // Stop
      clearInterval(pomodoroInterval);
      pomodoroInterval = null;
      pomodoroBtn.innerHTML = '<i class="ph ph-play"></i>';
      pomodoroContainer.classList.remove('running');
    } else {
      // Start
      pomodoroContainer.classList.add('running');
      pomodoroBtn.innerHTML = '<i class="ph ph-pause"></i>';
      pomodoroInterval = setInterval(() => {
        pomodoroSeconds--;
        updatePomodoroDisplay();
        if (pomodoroSeconds <= 0) {
          clearInterval(pomodoroInterval);
          pomodoroInterval = null;
          pomodoroSeconds = 25 * 60; // Reset
          pomodoroBtn.innerHTML = '<i class="ph ph-play"></i>';
          pomodoroContainer.classList.remove('running');
          updatePomodoroDisplay();
          playPomodoroChime();
        }
      }, 1000);
    }
  });

  // --- SCRATCHPAD ---
  const scratchpadToggle = document.getElementById('scratchpad-toggle');
  const scratchpadDrawer = document.getElementById('scratchpad-drawer');
  const closeScratchpadBtn = document.getElementById('close-scratchpad');
  const scratchpadTextarea = document.getElementById('scratchpad-textarea');

  // Load saved content
  const savedScratchpad = localStorage.getItem('cozy-scratchpad') || '';
  scratchpadTextarea.value = savedScratchpad;

  scratchpadToggle.addEventListener('click', () => {
    scratchpadDrawer.classList.toggle('open');
  });

  closeScratchpadBtn.addEventListener('click', () => {
    scratchpadDrawer.classList.remove('open');
  });

  scratchpadTextarea.addEventListener('input', (e) => {
    localStorage.setItem('cozy-scratchpad', e.target.value);
  });

  // --- FIND AND REPLACE WIDGET ---
  const frWidget = document.getElementById('find-replace-widget');
  const frFindInput = document.getElementById('fr-find-input');
  const frReplaceInput = document.getElementById('fr-replace-input');
  const frCloseBtn = document.getElementById('close-fr-btn');
  let searchCursor = null;

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'h')) {
      e.preventDefault();
      frWidget.style.display = 'flex';
      frFindInput.focus();
    }
  });

  frCloseBtn.addEventListener('click', () => {
    frWidget.style.display = 'none';
    searchCursor = null;
    quill.formatText(0, quill.getLength(), 'background', false); // clear highlights in quill
  });

  document.getElementById('fr-next-btn').addEventListener('click', () => {
    const term = frFindInput.value;
    if (!term) return;
    const activeTab = document.querySelector('.top-tabs .tab-btn.active').getAttribute('data-tab');
    if (activeTab === 'write') {
      const text = quill.getText();
      const index = text.indexOf(term, searchCursor ? searchCursor + 1 : 0);
      if (index !== -1) {
        quill.setSelection(index, term.length);
        searchCursor = index;
      } else {
        searchCursor = null; // loop around next time
      }
    }
  });

  document.getElementById('fr-replace-btn').addEventListener('click', () => {
    const term = frFindInput.value;
    const replaceTerm = frReplaceInput.value;
    if (!term) return;
    const activeTab = document.querySelector('.top-tabs .tab-btn.active').getAttribute('data-tab');
    
    if (activeTab === 'write') {
      const selection = quill.getSelection();
      if (selection && selection.length > 0 && quill.getText(selection.index, selection.length) === term) {
        quill.deleteText(selection.index, selection.length);
        quill.insertText(selection.index, replaceTerm);
        quill.setSelection(selection.index, replaceTerm.length);
      } else {
        document.getElementById('fr-next-btn').click();
      }
    }
  });

  document.getElementById('fr-replace-all-btn').addEventListener('click', () => {
    const term = frFindInput.value;
    const replaceTerm = frReplaceInput.value;
    if (!term) return;
    const activeTab = document.querySelector('.top-tabs .tab-btn.active').getAttribute('data-tab');
    
    if (activeTab === 'write') {
      let text = quill.getText();
      let index = text.indexOf(term);
      let offset = 0;
      while (index !== -1) {
        quill.deleteText(index, term.length);
        quill.insertText(index, replaceTerm);
        text = quill.getText();
        index = text.indexOf(term, index + replaceTerm.length);
      }
    }
  });

  // --- NATIVE FILE IPC ---
  if (window.electronAPI) {
    document.getElementById('global-open-btn').addEventListener('click', () => {
      window.electronAPI.openFileDialog();
    });

    window.electronAPI.onOpenFile((filePath, fileData, ext) => {
      document.getElementById('global-doc-title').value = filePath.split(/[/\\]/).pop();
      
      if (ext === '.docx') {
        document.querySelector('.tab-btn[data-tab="write"]').click();
        mammoth.convertToHtml({ arrayBuffer: fileData }).then(result => {
          quill.clipboard.dangerouslyPasteHTML(result.value);
        }).catch(err => {
          console.error('Error parsing DOCX:', err);
          alert('Could not parse DOCX file.');
        });
      } else if (ext === '.xlsx' || ext === '.csv') {
        document.querySelector('.tab-btn[data-tab="calculate"]').click();
        const wb = XLSX.read(fileData, { type: 'array' });
        excelGrid.loadData(stox(wb));
        setTimeout(() => window.excelGrid.resize(), 100);
      } else if (ext === '.pdf') {
        document.querySelector('.tab-btn[data-tab="read"]').click();
        loadPdf(fileData, filePath.split(/[/\\]/).pop());
      } else if (ext === '.pptx') {
        document.querySelector('.tab-btn[data-tab="present"]').click();
        alert('PPTX parsing is not natively implemented yet. Opening a blank presentation.');
      }
    });
  }

  // --- AMBIENT AUDIO ---
  const ambientAudioBtn = document.getElementById('ambient-audio-btn');
  const ambientPopup = document.getElementById('ambient-audio-popup');
  
  ambientAudioBtn.addEventListener('click', () => {
    ambientPopup.style.display = ambientPopup.style.display === 'none' ? 'flex' : 'none';
  });
  
  let audioCtx;
  let rainNodes = [], noiseNode = null;
  
  document.getElementById('audio-rain').addEventListener('change', (e) => {
    if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (e.target.checked) {
      const bufferSize = audioCtx.sampleRate * 2;
      const buffer = audioCtx.createBuffer(1, bufferSize, audioCtx.sampleRate);
      const data = buffer.getChannelData(0);
      for (let i = 0; i < bufferSize; i++) data[i] = Math.random() * 2 - 1;
      const noise = audioCtx.createBufferSource();
      noise.buffer = buffer;
      noise.loop = true;
      const filter = audioCtx.createBiquadFilter();
      filter.type = 'lowpass';
      filter.frequency.value = 1000;
      const gain = audioCtx.createGain();
      gain.gain.value = 0.5;
      noise.connect(filter);
      filter.connect(gain);
      gain.connect(audioCtx.destination);
      noise.start();
      rainNodes.push({ noise, gain });
    } else {
      rainNodes.forEach(n => { n.noise.stop(); n.noise.disconnect(); });
      rainNodes = [];
    }
  });

  document.getElementById('audio-noise').addEventListener('change', (e) => {
    if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (e.target.checked) {
      const bufferSize = audioCtx.sampleRate * 2;
      const buffer = audioCtx.createBuffer(1, bufferSize, audioCtx.sampleRate);
      const data = buffer.getChannelData(0);
      for (let i = 0; i < bufferSize; i++) data[i] = Math.random() * 2 - 1;
      const noise = audioCtx.createBufferSource();
      noise.buffer = buffer;
      noise.loop = true;
      const gain = audioCtx.createGain();
      gain.gain.value = 0.1;
      noise.connect(gain);
      gain.connect(audioCtx.destination);
      noise.start();
      noiseNode = { noise, gain };
    } else {
      if (noiseNode) { noiseNode.noise.stop(); noiseNode.noise.disconnect(); }
    }
  });

  // --- STUDY TAB (Flashcards) ---
  const STUDY_KEY = 'cozy-flashcards';
  let flashcards = JSON.parse(localStorage.getItem(STUDY_KEY) || '[]');
  
  if (flashcards.length === 0) {
    flashcards = [
      { front: 'Mitochondria', back: 'Powerhouse of the cell' },
      { front: 'O(n log n)', back: 'Average time complexity of merge sort' },
      { front: 'Photosynthesis', back: 'Process by which plants use sunlight to synthesize foods' }
    ];
    localStorage.setItem(STUDY_KEY, JSON.stringify(flashcards));
  }
  
  const studyGrid = document.getElementById('study-grid');

  const renderFlashcards = () => {
    studyGrid.innerHTML = '';
    flashcards.forEach((card, idx) => {
      const cardEl = document.createElement('div');
      cardEl.className = 'flashcard';
      cardEl.innerHTML = `
        <div class="front">${card.front}</div>
        <div class="back">${card.back}</div>
        <div class="actions">
          <button class="icon-btn" onclick="deleteCard(${idx})"><i class="ph ph-trash"></i></button>
        </div>
      `;
      studyGrid.appendChild(cardEl);
    });
  };

  window.deleteCard = (idx) => {
    flashcards.splice(idx, 1);
    localStorage.setItem(STUDY_KEY, JSON.stringify(flashcards));
    renderFlashcards();
  };

  const floatingAddBtn = document.getElementById('floating-add-flashcard-btn');
  const addModal = document.getElementById('add-flashcard-modal');
  const frontPreview = document.getElementById('flashcard-front-preview');
  const backInput = document.getElementById('flashcard-back-input');
  
  document.addEventListener('mouseup', () => {
    const activeTab = document.querySelector('.top-tabs .tab-btn.active')?.getAttribute('data-tab');
    if (activeTab === 'write' || activeTab === 'read') {
      const selection = window.getSelection();
      if (selection.toString().trim().length > 0) {
        const range = selection.getRangeAt(0);
        const rect = range.getBoundingClientRect();
        floatingAddBtn.style.top = `${rect.bottom + 10}px`;
        floatingAddBtn.style.left = `${rect.left}px`;
        floatingAddBtn.style.display = 'flex';
      } else {
        floatingAddBtn.style.display = 'none';
      }
    } else {
      floatingAddBtn.style.display = 'none';
    }
  });

  floatingAddBtn.addEventListener('click', () => {
    frontPreview.textContent = window.getSelection().toString();
    backInput.value = '';
    addModal.style.display = 'flex';
    floatingAddBtn.style.display = 'none';
    backInput.focus();
  });

  document.getElementById('flashcard-cancel-btn').addEventListener('click', () => {
    addModal.style.display = 'none';
  });

  document.getElementById('flashcard-save-btn').addEventListener('click', () => {
    flashcards.push({ front: frontPreview.textContent, back: backInput.value });
    localStorage.setItem(STUDY_KEY, JSON.stringify(flashcards));
    addModal.style.display = 'none';
    renderFlashcards();
  });
  
  document.getElementById('study-add-btn').addEventListener('click', () => {
    frontPreview.textContent = 'Manual Card Front (Double click to edit)';
    frontPreview.contentEditable = true;
    backInput.value = '';
    addModal.style.display = 'flex';
  });

  // Review Flashcards
  const reviewModal = document.getElementById('review-flashcards-modal');
  const reviewCardText = document.getElementById('review-card-text');
  let currentReviewIdx = 0;
  let showingFront = true;

  document.getElementById('study-review-btn').addEventListener('click', () => {
    if (flashcards.length === 0) return alert("No flashcards to review.");
    currentReviewIdx = 0;
    showingFront = true;
    updateReviewCard();
    reviewModal.style.display = 'flex';
  });

  const updateReviewCard = () => {
    const card = flashcards[currentReviewIdx];
    reviewCardText.textContent = showingFront ? card.front : card.back;
  };

  document.getElementById('review-card').addEventListener('click', () => {
    showingFront = !showingFront;
    updateReviewCard();
  });

  document.getElementById('review-next-btn').addEventListener('click', () => {
    currentReviewIdx = (currentReviewIdx + 1) % flashcards.length;
    showingFront = true;
    updateReviewCard();
  });

  document.getElementById('review-prev-btn').addEventListener('click', () => {
    currentReviewIdx = (currentReviewIdx - 1 + flashcards.length) % flashcards.length;
    showingFront = true;
    updateReviewCard();
  });

  document.getElementById('close-review-btn').addEventListener('click', () => {
    reviewModal.style.display = 'none';
  });

  renderFlashcards();

  // --- EXCALIDRAW (Board Tab) ---
  const excalidrawContainer = document.getElementById('excalidraw-container');
  if (excalidrawContainer) {
    const root = createRoot(excalidrawContainer);
    const savedBoard = localStorage.getItem('cozy-board');
    const initialData = savedBoard ? JSON.parse(savedBoard) : null;
    
    let boardTimeout;
    const onBoardChange = (elements, appState, files) => {
      clearTimeout(boardTimeout);
      boardTimeout = setTimeout(() => {
        localStorage.setItem('cozy-board', JSON.stringify({ elements }));
      }, 1000);
    };

    root.render(React.createElement(
      'div',
      { style: { height: "calc(100vh - 100px)", width: "100%", position: "relative" } },
      React.createElement(Excalidraw, {
        initialData: initialData,
        onChange: onBoardChange,
        theme: document.body.classList.contains('dark') ? 'dark' : 'light'
      })
    ));
  }

  // Initialize initial status bar for Docs tab
  updateStatusBar('write');
});
