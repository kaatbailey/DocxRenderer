package com.docx.docxrenderer.service;

import com.docx.docxrenderer.util.ImageUtil;
import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.*;
import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfContentByte;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class PdfConversionStrategy {

    /**
     * Converts an XWPFDocument to a PDF byte array using OpenPDF.
     * Returns the raw bytes which the controller streams back as application/pdf.
     */
    public byte[] convert(XWPFDocument document) throws IOException, DocumentException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        Document pdf = new Document(PageSize.A4);
        PdfWriter.getInstance(pdf, out);
        pdf.open();

        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                Element pdfElement = convertParagraph(paragraph);
                if (pdfElement != null) {
                    pdf.add(pdfElement);
                }
            } else if (element instanceof XWPFTable table) {
                pdf.add(convertTable(table));
            }
        }

        pdf.close();
        return out.toByteArray();
    }

    // ─── Paragraph ───────────────────────────────────────────────────────────

    private Element convertParagraph(XWPFParagraph paragraph) throws IOException {
        // Check for embedded images first
        List<XWPFPicture> pictures = getAllPictures(paragraph);
        if (!pictures.isEmpty()) {
            // Return the first image found in this paragraph
            byte[] raw = pictures.get(0).getPictureData().getData();
            byte[] compressed = ImageUtil.compress(raw);
            Image image = Image.getInstance(compressed);
            image.scaleToFit(500, 700);
            return image;
        }

        // Build a Phrase from the runs
        Paragraph pdfParagraph = new Paragraph();
        pdfParagraph.setSpacingAfter(6f);

        // Apply heading size if applicable
        String style = paragraph.getStyle();
        float fontSize = resolveHeadingSize(style);
        boolean isHeading = fontSize > 12f;

        if (isHeading) {
            pdfParagraph.setSpacingBefore(10f);
            pdfParagraph.setSpacingAfter(4f);
        }

        // Alignment
        if (paragraph.getAlignment() != null) {
            switch (paragraph.getAlignment()) {
                case CENTER -> pdfParagraph.setAlignment(Element.ALIGN_CENTER);
                case RIGHT  -> pdfParagraph.setAlignment(Element.ALIGN_RIGHT);
                case BOTH   -> pdfParagraph.setAlignment(Element.ALIGN_JUSTIFIED);
                default     -> pdfParagraph.setAlignment(Element.ALIGN_LEFT);
            }
        }

        for (XWPFRun run : paragraph.getRuns()) {
            Chunk chunk = convertRun(run, isHeading ? fontSize : 0);
            if (chunk != null) {
                pdfParagraph.add(chunk);
            }
        }

        // Return null for completely empty paragraphs to avoid blank pages
        if (pdfParagraph.isEmpty()) return null;

        return pdfParagraph;
    }

    // ─── Run ─────────────────────────────────────────────────────────────────

    private Chunk convertRun(XWPFRun run, float overrideFontSize) {
        String text = run.getText(0);
        if (text == null || text.isEmpty()) return null;

        int style = Font.NORMAL;
        if (run.isBold() && run.isItalic()) style = Font.BOLDITALIC;
        else if (run.isBold())              style = Font.BOLD;
        else if (run.isItalic())            style = Font.ITALIC;

        float size = overrideFontSize > 0
                ? overrideFontSize
                : resolveFontSize(run);

        Font font = new Font(Font.HELVETICA, size, style);

        // Font color
        String hex = run.getColor();
        if (hex != null && !hex.equalsIgnoreCase("auto")) {
            try {
                font.setColor(Color.decode("#" + hex));
            } catch (NumberFormatException ignored) {}
        }

        Chunk chunk = new Chunk(text, font);

        if (run.isStrikeThrough()) {
            chunk.setTextRenderMode(PdfContentByte.TEXT_RENDER_MODE_FILL, 0.5f, Color.BLACK);
        }

        if (run.getUnderline() != UnderlinePatterns.NONE) {
            chunk.setUnderline(0.5f, -1.5f);
        }

        return chunk;
    }

    // ─── Table ───────────────────────────────────────────────────────────────

    private PdfPTable convertTable(XWPFTable table) throws IOException {
        int numCols = table.getRows().get(0).getTableCells().size();
        PdfPTable pdfTable = new PdfPTable(numCols);
        pdfTable.setWidthPercentage(100);
        pdfTable.setSpacingBefore(8f);
        pdfTable.setSpacingAfter(8f);

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                PdfPCell pdfCell = new PdfPCell();
                pdfCell.setPadding(5f);

                Phrase phrase = new Phrase();
                for (XWPFParagraph p : cell.getParagraphs()) {
                    for (XWPFRun run : p.getRuns()) {
                        Chunk chunk = convertRun(run, 0);
                        if (chunk != null) phrase.add(chunk);
                    }
                }
                pdfCell.setPhrase(phrase);
                pdfTable.addCell(pdfCell);
            }
        }

        return pdfTable;
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    private float resolveHeadingSize(String style) {
        if (style == null) return 12f;
        return switch (style) {
            case "Heading1" -> 22f;
            case "Heading2" -> 18f;
            case "Heading3" -> 15f;
            case "Heading4" -> 13f;
            default         -> 12f;
        };
    }

    private float resolveFontSize(XWPFRun run) {
        int size = run.getFontSize();
        return size > 0 ? (float) size : 12f;
    }

    private List<XWPFPicture> getAllPictures(XWPFParagraph paragraph) {
        return paragraph.getRuns().stream()
                .flatMap(r -> r.getEmbeddedPictures().stream())
                .toList();
    }
}