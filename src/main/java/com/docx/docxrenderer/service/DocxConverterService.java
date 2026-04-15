package com.docx.docxrenderer.service;

import com.docx.docxrenderer.model.ConvertResponse;
import com.lowagie.text.DocumentException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Service
public class DocxConverterService {

    private final HtmlConversionStrategy htmlStrategy;
    private final PdfConversionStrategy pdfStrategy;
    private final JsonConversionStrategy jsonStrategy;

    public DocxConverterService() {
        this.htmlStrategy = new HtmlConversionStrategy();
        this.pdfStrategy = new PdfConversionStrategy();
        this.jsonStrategy = new JsonConversionStrategy();
    }

    /**
     * Converts the uploaded .docx to a PDF byte array.
     */
    public byte[] toPdf(MultipartFile file) {
        try (XWPFDocument document = new XWPFDocument(file.getInputStream())) {
            return pdfStrategy.convert(document);
        } catch (IOException | DocumentException e) {
            throw new RuntimeException("Failed to convert document to PDF: " + e.getMessage(), e);
        }
    }

    /**
     * Converts the uploaded .docx to a self-contained HTML string.
     */
    public String toHtml(MultipartFile file) {
        try (XWPFDocument document = new XWPFDocument(file.getInputStream())) {
            return htmlStrategy.convert(document);
        } catch (IOException e) {
            throw new RuntimeException("Failed to convert document to HTML: " + e.getMessage(), e);
        }
    }

    /**
     * Converts the uploaded .docx to a ConvertResponse DTO for JSON delivery.
     */
    public ConvertResponse toJson(MultipartFile file) {
        try (XWPFDocument document = new XWPFDocument(file.getInputStream())) {
            return jsonStrategy.convert(document);
        } catch (IOException e) {
            throw new RuntimeException("Failed to convert document to JSON: " + e.getMessage(), e);
        }
    }
}