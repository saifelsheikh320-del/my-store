const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const { Document, Packer, Paragraph, TextRun, AlignmentType, SectionType, PageBreak, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');
const Tesseract = require('tesseract.js');
const Jimp = require('jimp');

// Configure PDF.js for Node environment
const PDFJS_DIST_PATH = path.dirname(require.resolve('pdfjs-dist/package.json'));
const CMAP_URL = path.join(PDFJS_DIST_PATH, 'cmaps' + path.sep);
const STANDARD_FONT_DATA_URL = path.join(PDFJS_DIST_PATH, 'standard_fonts' + path.sep);

class PDFToWordConverter {
    constructor() {
        this.pageMargins = { top: 720, right: 720, bottom: 720, left: 720 };
    }

    async convert(pdfPath, outputPath, onProgress) {
        console.log(`[ENGINE] Starting Advanced Hybrid Conversion: ${pdfPath}`);

        // Added Hungarian (hun) to support medical reports provided by user
        this.worker = await Tesseract.createWorker('ara+eng+hun', 1, {
            logger: m => {
                if (m.status === 'recognizing text' && onProgress) {
                    onProgress(Math.round(m.progress * 100));
                }
            }
        });

        await this.worker.setParameters({
            tessedit_pageseg_mode: Tesseract.PSM.AUTO,
            preserve_interword_spaces: '1',
        });

        const data = new Uint8Array(fs.readFileSync(pdfPath));
        const loadingTask = pdfjsLib.getDocument({
            data,
            cMapUrl: CMAP_URL,
            cMapPacked: true,
            standardFontDataUrl: STANDARD_FONT_DATA_URL,
            useSystemFonts: true,
            disableFontFace: false
        });

        const pdfDocument = await loadingTask.promise;
        const numPages = pdfDocument.numPages;
        const sections = [];

        const headerFooterMap = await this.analyzeHeadersFooters(pdfDocument);

        for (let i = 1; i <= numPages; i++) {
            const pageProgress = Math.round(((i - 1) / numPages) * 100);
            if (onProgress) onProgress(pageProgress);

            console.log(`[ENGINE] Analyzing Page ${i}/${numPages}...`);
            const page = await pdfDocument.getPage(i);
            const content = await this.processHybridPage(page, headerFooterMap);

            sections.push({
                properties: {
                    type: SectionType.NEXT_PAGE,
                    page: { margin: this.pageMargins },
                },
                children: content,
            });
        }

        const doc = new Document({ sections });
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(outputPath, buffer);

        if (this.worker) {
            await this.worker.terminate();
            this.worker = null;
        }
        if (onProgress) onProgress(100);
        console.log(`[ENGINE] Process Completed Successfully.`);
    }

    async analyzeHeadersFooters(pdfDocument) {
        const pageCount = pdfDocument.numPages;
        if (pageCount < 2) return new Set();
        const positionMap = new Map();
        const sampleSize = Math.min(pageCount, 5);

        for (let i = 1; i <= sampleSize; i++) {
            const page = await pdfDocument.getPage(i);
            const textContent = await page.getTextContent();
            const viewport = page.getViewport({ scale: 1.0 });

            textContent.items.forEach(item => {
                if (!item.str) return;
                const y = Math.round(viewport.height - item.transform[5]);
                if (y < viewport.height * 0.12 || y > viewport.height * 0.88) {
                    const key = `${Math.round(item.transform[4])}_${y}_${item.str.trim()}`;
                    if (item.str.trim().length > 2) {
                        positionMap.set(key, (positionMap.get(key) || 0) + 1);
                    }
                }
            });
        }
        const repeatable = new Set();
        positionMap.forEach((count, key) => {
            if (count >= sampleSize - 1) repeatable.add(key);
        });
        return repeatable;
    }

    async processHybridPage(page, headerFooterMap) {
        const textContent = await page.getTextContent();
        const viewport = page.getViewport({ scale: 1.0 });
        const { width, height } = viewport;

        // 1. Collect Digital Elements
        let digitalItems = textContent.items.map(item => {
            const tx = item.transform[4];
            const ty = item.transform[5];
            const y = height - ty;
            const x = tx;
            const key = `${Math.round(x)}_${Math.round(y)}_${(item.str || '').trim()}`;

            return {
                text: item.str || '',
                x: x, y: y,
                width: item.width || 0,
                height: item.height || 0,
                fontSize: Math.sqrt(item.transform[0] ** 2 + item.transform[1] ** 2),
                isHeaderFooter: headerFooterMap.has(key)
            };
        }).filter(it => !it.isHeaderFooter && it.text.trim().length > 0);

        // 2. Collect OCR Elements (Mandatory check for images)
        const ocrItems = await this.performDeepOCR(page, width, height);

        // 3. Combine and Sort Hybrid Items
        const allItems = [...digitalItems, ...ocrItems];

        if (allItems.length === 0) return [new Paragraph({ children: [] })];

        // 4. Layout Reconstruction Engine
        return this.reconstructLayout(allItems, width);
    }

    async performDeepOCR(page, viewportWidth, viewportHeight) {
        try {
            const ops = await page.getOperatorList();
            const imageObjects = [];

            // To track image positions, we'd need to simulate the CTM, 
            // but for medical dossiers, images are usually full-page or large blocks.
            // We'll treat OCR results as middle-of-page content if mixed.
            for (let i = 0; i < ops.fnArray.length; i++) {
                if (ops.fnArray[i] === pdfjsLib.OPS.paintImageXObject || ops.fnArray[i] === pdfjsLib.OPS.paintInlineImageXObject) {
                    const img = await page.objs.get(ops.argsArray[i][0]);
                    if (img) imageObjects.push(img);
                }
            }

            if (imageObjects.length === 0) return [];

            console.log(`[ENGINE] Embedded images found: ${imageObjects.length}. Starting Deep Analysis...`);
            const sortedImages = imageObjects.sort((a, b) => (b.width * b.height) - (a.width * a.height));
            const items = [];

            // Process up to 3 largest images to catch all medical reports
            for (const img of sortedImages.slice(0, 3)) {
                if (img.width < 100 || img.height < 100) continue;

                const jimg = new Jimp(img.width, img.height);
                // Correct pixel data mapping
                if (img.kind === 1) { // RGB
                    for (let i = 0, idx = 0; i < img.data.length; i += 3, idx += 4) {
                        jimg.bitmap.data[idx] = img.data[i];
                        jimg.bitmap.data[idx + 1] = img.data[i + 1];
                        jimg.bitmap.data[idx + 2] = img.data[i + 2];
                        jimg.bitmap.data[idx + 3] = 255;
                    }
                } else if (img.kind === 3) { // Grayscale
                    for (let i = 0, idx = 0; i < img.data.length; i++, idx += 4) {
                        jimg.bitmap.data[idx] = img.data[i];
                        jimg.bitmap.data[idx + 1] = img.data[i];
                        jimg.bitmap.data[idx + 2] = img.data[i];
                        jimg.bitmap.data[idx + 3] = 255;
                    }
                } else {
                    jimg.bitmap.data = Buffer.from(img.data);
                }

                // AI Enhancement: Scale and Normalize for faint/scan text
                if (img.width < 2200) await jimg.scale(2.5);
                await jimg.grayscale().contrast(0.2).normalize();

                const buffer = await jimg.getBufferAsync(Jimp.MIME_PNG);
                const { data: { blocks } } = await this.worker.recognize(buffer);

                if (blocks) {
                    blocks.forEach(block => {
                        // Map block bbox back to PDF coordinates (vague mapping for embedded images)
                        // Assume images are centered if they are large
                        items.push({
                            text: block.text.trim(),
                            x: (viewportWidth * 0.1), // Estimated margin
                            y: (viewportHeight * 0.5), // Estimated middle
                            width: viewportWidth * 0.8,
                            height: 20,
                            fontSize: 11,
                            isOCR: true
                        });
                    });
                }
            }
            return items;
        } catch (e) {
            console.error("[OCR ERROR]", e);
            return [];
        }
    }

    reconstructLayout(items, pageWidth) {
        // Precise Y sorting
        items.sort((a, b) => Math.abs(a.y - b.y) < 4 ? a.x - b.x : a.y - b.y);

        const lines = [];
        let currentLine = { items: [items[0]], y: items[0].y, fontSize: items[0].fontSize || 12 };

        for (let i = 1; i < items.length; i++) {
            const it = items[i];
            if (Math.abs(it.y - currentLine.y) < (currentLine.fontSize * 0.6)) {
                currentLine.items.push(it);
            } else {
                lines.push(this.finalizeLine(currentLine));
                currentLine = { items: [it], y: it.y, fontSize: it.fontSize || 12 };
            }
        }
        lines.push(this.finalizeLine(currentLine));

        // Detect Tables or Paragraphs
        const blocks = this.detectStructure(lines, pageWidth);
        const elements = [];

        for (const block of blocks) {
            if (block.type === 'table') {
                elements.push(this.buildTable(block));
            } else {
                elements.push(this.buildParagraph(block, pageWidth));
            }
        }
        return elements;
    }

    finalizeLine(line) {
        line.items.sort((a, b) => a.x - b.x);
        let text = '';
        let minX = 9999, maxX = -9999;
        let totalSize = 0;

        line.items.forEach((it, i) => {
            text += it.text;
            totalSize += it.fontSize;
            minX = Math.min(minX, it.x);
            maxX = Math.max(maxX, it.x + it.width);

            const next = line.items[i + 1];
            if (next) {
                const gap = next.x - (it.x + it.width);
                if (gap > (it.fontSize * 0.25)) text += ' ';
            }
        });

        return {
            text: text.trim(),
            y: line.y,
            minX, maxX, width: maxX - minX,
            avgFontSize: totalSize / line.items.length,
            items: line.items
        };
    }

    detectStructure(lines, pageWidth) {
        const blocks = [];
        let currentTableRows = [];

        for (const line of lines) {
            // Check if line has multiple isolated columns (table-like)
            const gapCount = this.countLargeGaps(line.items);
            if (gapCount >= 1 && line.text.length > 5) {
                currentTableRows.push(line);
            } else {
                if (currentTableRows.length > 0) {
                    blocks.push({ type: 'table', rows: currentTableRows });
                    currentTableRows = [];
                }
                blocks.push({ type: 'text', lines: [line] });
            }
        }
        if (currentTableRows.length > 0) blocks.push({ type: 'table', rows: currentTableRows });

        // Merge adjacent text blocks
        const merged = [];
        for (const block of blocks) {
            const last = merged[merged.length - 1];
            if (block.type === 'text' && last && last.type === 'text') {
                const prevLine = last.lines[last.lines.length - 1];
                const currLine = block.lines[0];
                const vGap = currLine.y - prevLine.y;
                if (vGap < (prevLine.avgFontSize * 2) && Math.abs(currLine.minX - prevLine.minX) < 50) {
                    last.lines.push(...block.lines);
                    continue;
                }
            }
            merged.push(block);
        }
        return merged;
    }

    countLargeGaps(items) {
        let count = 0;
        for (let i = 0; i < items.length - 1; i++) {
            const gap = items[i + 1].x - (items[i].x + items[i].width);
            if (gap > 60) count++;
        }
        return count;
    }

    buildParagraph(block, pageWidth) {
        const text = block.lines.map(l => l.text).join(' ');
        const isRTL = this.containsRTL(text);

        return new Paragraph({
            children: block.lines.map(line => new TextRun({
                text: line.text + ' ',
                size: Math.round(line.avgFontSize * 2),
                rightToLeft: isRTL,
                font: isRTL ? "Traditional Arabic" : "Arial"
            })),
            alignment: this.detectAlignment(block.lines[0], pageWidth),
            bidirectional: isRTL,
            spacing: { before: 100, after: 100, line: 300 }
        });
    }

    buildTable(block) {
        const rows = block.rows.map(rowLine => {
            // Heuristic column split
            const cells = [];
            let currentCellItems = [rowLine.items[0]];
            for (let i = 1; i < rowLine.items.length; i++) {
                const it = rowLine.items[i];
                const prev = rowLine.items[i - 1];
                if (it.x - (prev.x + prev.width) > 50) {
                    cells.push(this.createCell(currentCellItems, rowLine.avgFontSize));
                    currentCellItems = [it];
                } else {
                    currentCellItems.push(it);
                }
            }
            cells.push(this.createCell(currentCellItems, rowLine.avgFontSize));
            return new TableRow({ children: cells });
        });

        return new Table({
            rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideHorizontal: { style: BorderStyle.NONE },
                insideVertical: { style: BorderStyle.NONE },
            }
        });
    }

    createCell(items, fontSize) {
        const text = items.map(it => it.text).join(' ').trim();
        return new TableCell({
            children: [new Paragraph({
                children: [new TextRun({ text, size: Math.round(fontSize * 2) })],
            })]
        });
    }

    detectAlignment(line, pageWidth) {
        const center = (line.minX + line.maxX) / 2;
        if (Math.abs(center - pageWidth / 2) < 40 && line.width < pageWidth * 0.6) return AlignmentType.CENTER;
        if (this.containsRTL(line.text)) return AlignmentType.RIGHT;
        return AlignmentType.LEFT;
    }

    containsRTL(text) {
        return /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(text);
    }
}

module.exports = PDFToWordConverter;
