package com.docx.docxrenderer.service;

import com.docx.docxrenderer.util.ImageUtil;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.util.Base64;
import java.util.List;

public class HtmlConversionStrategy {

    /**
     * Converts an XWPFDocument to a self-contained HTML string.
     * Images are base64 inlined, CSS is embedded in a <style> tag.
     * The result can be returned directly to a browser with no extra assets.
     */
    public String convert(XWPFDocument document) throws IOException {
        StringBuilder html = new StringBuilder();
        StringBuilder body = new StringBuilder();

        // ─── Walk every block-level element in the document ──────────────────
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                body.append(convertParagraph(paragraph));
            } else if (element instanceof XWPFTable table) {
                body.append(convertTable(table));
            }
        }

        // ─── Assemble the full HTML document ─────────────────────────────────
        html.append("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n");
        html.append("<style>\n").append(buildCss()).append("</style>\n");
        html.append("</head>\n<body>\n");
        html.append(body);
        html.append("</body>\n</html>");

        return html.toString();
    }

    // ─── Paragraph ───────────────────────────────────────────────────────────

    private String convertParagraph(XWPFParagraph paragraph) throws IOException {
        String style = paragraph.getStyle();
        String tag = resolveTag(style);

        StringBuilder sb = new StringBuilder();
        sb.append("<").append(tag);

        // Text alignment
        if (paragraph.getAlignment() != null) {
            switch (paragraph.getAlignment()) {
                case CENTER -> sb.append(" style=\"text-align:center\"");
                case RIGHT  -> sb.append(" style=\"text-align:right\"");
                case BOTH   -> sb.append(" style=\"text-align:justify\"");
                default     -> {}
            }
        }

        sb.append(">");

        // ─── Runs inside the paragraph ────────────────────────────────────────
        for (XWPFRun run : paragraph.getRuns()) {
            sb.append(convertRun(run));
        }

        // ─── Pictures embedded in the paragraph ──────────────────────────────
        for (XWPFPicture picture : getAllPictures(paragraph)) {
            byte[] rawBytes = picture.getPictureData().getData();
            byte[] compressed = ImageUtil.compress(rawBytes);
            String base64 = Base64.getEncoder().encodeToString(compressed);
            sb.append("<img src=\"data:image/jpeg;base64,")
                    .append(base64)
                    .append("\" style=\"max-width:100%\">");
        }

        sb.append("</").append(tag).append(">\n");
        return sb.toString();
    }

    // ─── Run (inline text with formatting) ───────────────────────────────────

    private String convertRun(XWPFRun run) {
        String text = run.getText(0);
        if (text == null || text.isEmpty()) return "";

        // Escape HTML special characters
        text = text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");

        StringBuilder sb = new StringBuilder();

        if (run.isBold())   sb.append("<strong>");
        if (run.isItalic()) sb.append("<em>");
        if (run.isStrikeThrough()) sb.append("<s>");
        if (run.getUnderline() != UnderlinePatterns.NONE) sb.append("<u>");

        // Font color
        String color = run.getColor();
        if (color != null && !color.equalsIgnoreCase("auto")) {
            sb.append("<span style=\"color:#").append(color).append("\">");
        }

        sb.append(text);

        // Close in reverse order
        if (color != null && !color.equalsIgnoreCase("auto")) sb.append("</span>");
        if (run.getUnderline() != UnderlinePatterns.NONE) sb.append("</u>");
        if (run.isStrikeThrough()) sb.append("</s>");
        if (run.isItalic()) sb.append("</em>");
        if (run.isBold())   sb.append("</strong>");

        return sb.toString();
    }

    // ─── Table ───────────────────────────────────────────────────────────────

    private String convertTable(XWPFTable table) throws IOException {
        StringBuilder sb = new StringBuilder();
        sb.append("<table>\n");

        for (XWPFTableRow row : table.getRows()) {
            sb.append("<tr>\n");
            for (XWPFTableCell cell : row.getTableCells()) {
                sb.append("<td>");
                for (XWPFParagraph p : cell.getParagraphs()) {
                    sb.append(convertParagraph(p));
                }
                sb.append("</td>\n");
            }
            sb.append("</tr>\n");
        }

        sb.append("</table>\n");
        return sb.toString();
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    /**
     * Maps Word paragraph styles to HTML tags.
     * Word uses style names like "Heading1", "Heading2" etc.
     * Everything else becomes a <p>.
     */
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

    /**
     * Base CSS applied to every HTML output document.
     * Keeps things readable without being opinionated about design.
     */
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
