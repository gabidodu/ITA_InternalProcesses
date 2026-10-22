import PDFDocument from "pdfkit";
import path from "path";
import fs from "fs";
import { formatDate, formatDateTime } from "../../shared/lib/dates";

// Ships alongside the compiled function code as a plain file (not a TS
// module) — see api/assets/logo.png, fetched once from the company's real
// SharePoint-hosted logo. dist/src/leave/lib -> ../../../../ reaches the api folder's
// own root (Azure Functions' wwwroot), where "assets" sits as a sibling of
// "dist"/"node_modules".
const LOGO_PATH = path.join(__dirname, "..", "..", "..", "..", "assets", "logo.png");

// PNG stores its pixel dimensions as two big-endian uint32s at a fixed offset
// in the IHDR chunk (byte 16 = width, byte 20 = height) — reading them
// directly avoids pulling in an image-metadata dependency just to keep the
// logo's aspect ratio correct when scaling it to a fixed width.
function getPngDimensions(filePath: string): { width: number; height: number } {
    const buffer = fs.readFileSync(filePath);
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

export type ApprovalSignature = {
    actorName: string;
    timestamp: string;
};

export type LeaveRequestPdfData = {
    requestId: string;
    requestDate: string;
    employeeFullName: string;
    department?: string;
    jobTitle?: string;
    companyName?: string;
    actionTypeName: string;
    actionTypeNameEn?: string;
    startDate: string;
    endDate: string;
    takenDays: number;
    requestSubject: string;
    comments?: string;
    standInName?: string;
    signatures: ApprovalSignature[];
};

// pdfkit's built-in Standard-14 fonts (Helvetica) only cover WinAnsi/Latin-1 —
// none of Romanian's diacritics (ă/â/î/ș/ț and legacy ş/ţ) render without
// embedding a custom Unicode font. Rather than bundle one for this single
// generated form, diacritics are stripped to their base Latin letter — a
// common, perfectly legible convention for Romanian text in constrained fonts.
const DIACRITIC_MAP: Record<string, string> = {
    ă: "a", Ă: "A",
    â: "a", Â: "A",
    î: "i", Î: "I",
    ș: "s", Ș: "S",
    ş: "s", Ş: "S",
    ț: "t", Ț: "T",
    ţ: "t", Ţ: "T",
};

function stripDiacritics(text: string): string {
    return text.replace(/[ăĂâÂîÎșȘşŞțȚţŢ]/g, (ch) => DIACRITIC_MAP[ch] ?? ch);
}

type Segment = { text: string; bold?: boolean };

// Renders a run of regular/bold segments as one flowing paragraph — pdfkit's
// `continued: true` chains each call onto the same line(s) instead of
// starting a new one, which is how a sentence can have specific words in
// bold without breaking it into separate blocks. Text options (align,
// lineGap, ...) only need to be set on the first chunk — pdfkit carries them
// through the rest of a continued chain. NOT safe to combine with
// `align: "justify"` — see renderJustifiedRun below.
function renderRun(doc: PDFKit.PDFDocument, segments: Segment[], options?: PDFKit.Mixins.TextOptions): void {
    segments.forEach((seg, i) => {
        doc.font(seg.bold ? "Helvetica-Bold" : "Helvetica");
        const isLast = i === segments.length - 1;
        doc.text(stripDiacritics(seg.text), { continued: !isLast, ...(i === 0 ? options : {}) });
    });
}

type Word = { text: string; bold: boolean };

// pdfkit's built-in `align: "justify"` mis-measures lines that contain a
// `continued: true` font switch — confirmed by reproducing it twice, with two
// different data sets: a short bold token right at a line-wrap boundary gets
// glued to the next word, followed by an oversized gap earlier in the same
// line. Rather than fight that, this does justify by hand: split the run
// into words (tagging each with its bold/regular font), greedily wrap them
// into lines using measured widths, then render every line except the last
// with its inter-word gap stretched to fill the column exactly. `firstLineIndent`
// reserves that much space at the start of line one only (a traditional
// paragraph tab), narrowing just that line's wrap width and start position.
function renderJustifiedRun(doc: PDFKit.PDFDocument, segments: Segment[], lineGap: number, firstLineIndent = 0): void {
    const words: Word[] = [];
    segments.forEach((seg) => {
        stripDiacritics(seg.text)
            .split(" ")
            .filter((w) => w.length > 0)
            .forEach((w) => words.push({ text: w, bold: Boolean(seg.bold) }));
    });

    const startX = doc.page.margins.left;
    const fullWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

    function widthOf(word: Word): number {
        doc.font(word.bold ? "Helvetica-Bold" : "Helvetica");
        return doc.widthOfString(word.text);
    }
    function spaceWidth(bold: boolean): number {
        doc.font(bold ? "Helvetica-Bold" : "Helvetica");
        return doc.widthOfString(" ");
    }

    const lines: Word[][] = [];
    let current: Word[] = [];
    let currentWidth = 0;
    words.forEach((w) => {
        const wWidth = widthOf(w);
        const gapWidth = current.length > 0 ? spaceWidth(current[current.length - 1].bold) : 0;
        const lineMaxWidth = lines.length === 0 ? fullWidth - firstLineIndent : fullWidth;
        if (current.length > 0 && currentWidth + gapWidth + wWidth > lineMaxWidth) {
            lines.push(current);
            current = [w];
            currentWidth = wWidth;
        } else {
            if (current.length > 0) currentWidth += gapWidth;
            current.push(w);
            currentWidth += wWidth;
        }
    });
    if (current.length > 0) lines.push(current);

    const lineHeight = doc.currentLineHeight();
    let y = doc.y;

    lines.forEach((line, lineIndex) => {
        const isFirstLine = lineIndex === 0;
        const isLastLine = lineIndex === lines.length - 1;
        const lineStartX = isFirstLine ? startX + firstLineIndent : startX;
        const lineMaxWidth = isFirstLine ? fullWidth - firstLineIndent : fullWidth;
        const wordsWidth = line.reduce((sum, w) => sum + widthOf(w), 0);
        const gapCount = line.length - 1;
        const gap =
            !isLastLine && gapCount > 0 ? (lineMaxWidth - wordsWidth) / gapCount : spaceWidth(line[line.length - 1]?.bold ?? false);
        let x = lineStartX;
        line.forEach((w, i) => {
            doc.font(w.bold ? "Helvetica-Bold" : "Helvetica");
            doc.text(w.text, x, y, { lineBreak: false });
            x += widthOf(w);
            if (i < line.length - 1) x += gap;
        });
        y += lineHeight + lineGap;
    });

    doc.x = startX;
    doc.y = y;
}

/**
 * Renders the "cerere de concediu" as a PDF modeled on the company's own
 * legacy paper form (logo block, centered title, "Nr. X din Y" register
 * line, a declarative paragraph with the key facts in bold, a stand-in line,
 * a signature line, and one "De acord," block per approval step) rather than
 * a generic table layout. Generated only on final approval — see
 * RespondToLeaveRequest.ts / RespondToLeaveRequestAction.ts. Returns the
 * finished PDF as a Buffer, suitable for a Graph sendMail file attachment.
 */
export function generateLeaveRequestPdf(data: LeaveRequestPdfData): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument({ size: "A4", margin: 42 });
        const chunks: Buffer[] = [];
        doc.on("data", (chunk) => chunks.push(chunk));
        doc.on("end", () => resolve(Buffer.concat(chunks)));
        doc.on("error", reject);

        // Logo block: the real company logo (see LOGO_PATH) + tagline underneath.
        const logoX = doc.page.margins.left;
        const logoY = doc.page.margins.top;
        const logoWidth = 90;
        const logoDims = getPngDimensions(LOGO_PATH);
        const logoHeight = logoWidth * (logoDims.height / logoDims.width);
        doc.image(LOGO_PATH, logoX, logoY, { width: logoWidth });
        doc.fillColor("#999999").font("Helvetica-Oblique").fontSize(7).text("Departamentul tau IT", logoX, logoY + logoHeight + 4);
        doc.fillColor("#000000");
        doc.x = doc.page.margins.left;
        doc.y = logoY + logoHeight + 18;

        doc.moveDown(3.2);
        doc.font("Helvetica-Bold").fontSize(19).text("CERERE DE CONCEDIU", { align: "center" });
        doc.moveDown(1.5);

        doc.font("Helvetica").fontSize(11.5).text(`Nr. ${data.requestId} din ${formatDate(data.requestDate)}`, { align: "right" });
        doc.moveDown(3);

        doc.fontSize(13);
        renderJustifiedRun(
            doc,
            [
                { text: "Subsemnatul/a " },
                { text: data.employeeFullName, bold: true },
                { text: " angajat/a la societatea " },
                { text: data.companyName || "—", bold: true },
                { text: " pe postul de " },
                { text: data.jobTitle || "—", bold: true },
                { text: " in departamentul " },
                { text: `${data.department || "—"},`, bold: true },
                { text: " va rog sa aprobati efectuarea a " },
                { text: String(data.takenDays), bold: true },
                { text: " zile de concediu in perioada " },
                { text: `${formatDate(data.startDate)} - ${formatDate(data.endDate)}`, bold: true },
                { text: " reprezentand:" },
            ],
            6,
            24
        );

        doc.moveDown(1.8);
        doc.font("Helvetica").text(
            stripDiacritics(data.actionTypeNameEn ? `${data.actionTypeName} – ${data.actionTypeNameEn}` : data.actionTypeName),
            { indent: 20 }
        );

        if (data.comments) {
            doc.moveDown(1);
            doc.font("Helvetica-Oblique").fontSize(11.5).text(stripDiacritics(`Observatii: ${data.comments}`), { indent: 20 });
            doc.fontSize(13);
        }

        if (data.standInName) {
            doc.moveDown(2.2);
            renderRun(doc, [
                { text: "Persoana care imi va tine locul: " },
                { text: data.standInName, bold: true },
                { text: " ." },
            ]);
        }

        doc.moveDown(3);
        renderRun(doc, [{ text: "Data: " }, { text: formatDate(data.requestDate), bold: true }]);
        doc.moveDown(2.2);
        // Same visual pattern as the "De acord," approval blocks below — label,
        // then the name on its own line, then a smaller/greyed date+time under it.
        doc.font("Helvetica").fontSize(13).text("Semnatura angajat:");
        doc.font("Helvetica-Bold").text(`   ${stripDiacritics(data.employeeFullName)}`);
        doc.font("Helvetica").fontSize(9).fillColor("#888888").text(`   ${formatDateTime(data.requestDate)}`);
        doc.fillColor("#000000").fontSize(13);

        doc.moveDown(1.8);
        if (data.signatures.length === 0) {
            doc.font("Helvetica-Oblique").text("(niciun pas de aprobare inregistrat)");
        } else {
            data.signatures.forEach((sig) => {
                doc.font("Helvetica-Oblique").fontSize(13).text("De acord,");
                doc.font("Helvetica-Bold").text(`   ${stripDiacritics(sig.actorName)}`);
                doc.font("Helvetica").fontSize(9).fillColor("#888888").text(`   ${formatDateTime(sig.timestamp)}`);
                doc.fillColor("#000000").fontSize(13);
                doc.moveDown(1.7);
            });
        }

        doc.moveDown(1.8);
        doc.fontSize(8.5).fillColor("#888888").font("Helvetica").text(`Document generat automat la ${formatDateTime(new Date().toISOString())}.`);

        doc.end();
    });
}
