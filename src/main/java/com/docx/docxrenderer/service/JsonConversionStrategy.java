package com.docx.docxrenderer.service;

import com.docx.docxrenderer.model.ConvertResponse;
import com.docx.docxrenderer.util.ImageUtil;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class JsonConversionStrategy {

    /**
     * Converts an XWPFDocument to a ConvertResponse DTO containing:
     * - html: the document body as an HTML fragment (no <html> or <head> wrapper)
     * - css: base stylesheet as a string the client can inject
     * - metadata: title, word count, paragraph count
     * - images: list of base64 compressed images with IDs
     *
     * The HTML fragment references images by ID (<img src="img_0"> etc.)
     * The client resolves those IDs against the images array.
     */
    public ConvertResponse convert(XWPFDocument document) throws IOException {
        StringBuilder bodyHtml = new StringBuilder();
        List<ConvertResponse.ImageData> images = new ArrayList<>();
        int[] imageCounter = {0}; // array so the lambda can mutate it

        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                bodyHtml.append(convertParagraph(paragraph, images, imageCounter));
            } else if (element instanceof XWPFTable table) {
                bodyHtml.append(convertTable(table, images, imageCounter));
            }
        }

        ConvertResponse.Metadata metadata = buildMetadata(document);
        String css = buildCss();

        return new ConvertResponse(bodyHtml.toString(), css, metadata, images);
    }

    // ─── Paragraph ───────────────────────────────────────────────────────────

    private String convertParagraph(
            XWPFParagraph paragraph,
            List<ConvertResponse.ImageData> images,
            int[] imageCounter) throws IOException {

        String style = paragraph.getStyle();
        String tag = resolveTag(style);

        StringBuilder sb = new StringBuilder();
        sb.append("<").append(tag).append(">");

        for (XWPFRun run : paragraph.getRuns()) {
            sb.append(convertRun(run));
        }

        // Images — referenced by ID, added to the images list
        for (XWPFPicture picture : getAllPictures(paragraph)) {
            String id = "img_" + imageCounter[0]++;
            byte[] raw = picture.getPictureData().getData();
            String base64 = ImageUtil.compressToBase64(raw);
            String mimeType = "image/jpeg";

            images.add(new ConvertResponse.ImageData(id, mimeType, base64));
            sb.append("<img src=\"").append(id).append("\" style=\"max-width:100%\">");
        }

        sb.append("</").append(tag).append(">\n");
        return sb.toString();
    }

    // ─── Run ─────────────────────────────────────────────────────────────────

    private String convertRun(XWPFRun run) {
        String text = run.getText(0);
        if (text == null || text.isEmpty()) return "";

        text = text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");

        StringBuilder sb = new StringBuilder();

        if (run.isBold())   sb.append("<strong>");
        if (run.isItalic()) sb.append("<em>");
        if (run.isStrikeThrough()) sb.append("<s>");
        if (run.getUnderline() != UnderlinePatterns.NONE) sb.append("<u>");

        String color = run.getColor();
        if (color != null && !color.equalsIgnoreCase("auto")) {
            sb.append("<span style=\"color:#").append(color).append("\">");
        }

        sb.append(text);

        if (color != null && !color.equalsIgnoreCase("auto")) sb.append("</span>");
        if (run.getUnderline() != UnderlinePatterns.NONE) sb.append("</u>");
        if (run.isStrikeThrough()) sb.append("</s>");
        if (run.isItalic()) sb.append("</em>");
        if (run.isBold())   sb.append("</strong>");

        return sb.toString();
    }

    // ─── Table ───────────────────────────────────────────────────────────────

    private String convertTable(
            XWPFTable table,
            List<ConvertResponse.ImageData> images,
            int[] imageCounter) throws IOException {

        StringBuilder sb = new StringBuilder();
        sb.append("<table>\n");

        for (XWPFTableRow row : table.getRows()) {
            sb.append("<tr>\n");
            for (XWPFTableCell cell : row.getTableCells()) {
                sb.append("<td>");
                for (XWPFParagraph p : cell.getParagraphs()) {
                    sb.append(convertParagraph(p, images, imageCounter));
                }
                sb.append("</td>\n");
            }
            sb.append("</tr>\n");
        }

        sb.append("</table>\n");
        return sb.toString();
    }

    // ─── Metadata ────────────────────────────────────────────────────────────

    private ConvertResponse.Metadata buildMetadata(XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        // Use the first non-empty paragraph as the title
        String title = paragraphs.stream()
                .map(p -> p.getText().trim())
                .filter(t -> !t.isEmpty())
                .findFirst()
                .orElse("Untitled");

        int wordCount = paragraphs.stream()
                .map(XWPFParagraph::getText)
                .filter(t -> t != null && !t.isBlank())
                .mapToInt(t -> t.trim().split("\\s+").length)
                .sum();

        return new ConvertResponse.Metadata(title, wordCount, paragraphs.size());
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    private String resolveTag(String style) {
        if (style == null) return "p";
        return switch (style) {
            case "Heading1" -> "h1";
            case "Heading2" -> "h2";
            case "Heading3" -> "h3";
            case "Heading4" -> "h4";
            case "Heading5" -> "h5";
            case "Heading6" -> "h6";
            default         -> "p";
        };
    }

    private List<XWPFPicture> getAllPictures(XWPFParagraph paragraph) {
        return paragraph.getRuns().stream()
                .flatMap(run -> run.getEmbeddedPictures().stream())
                .toList();
    }

    private String buildCss() {
        return """
                body {
                    font-family: Arial, sans-serif;
                    font-size: 12pt;
                    line-height: 1.5;
                    max-width: 860px;
                    margin: 40px auto;
                    color: #222;
                }
                h1, h2, h3, h4, h5, h6 {
                    margin-top: 1.2em;
                    margin-bottom: 0.4em;
                }
                p { margin: 0.4em 0; }
                table {
                    border-collapse: collapse;
                    width: 100%;
                    margin: 1em 0;
                }
                td, th {
                    border: 1px solid #ccc;
                    padding: 6px 10px;
                }
                img { max-width: 100%; height: auto; }
                """;
    }
}