package com.docx.docxrenderer.controller;

import com.docx.docxrenderer.model.ConvertResponse;
import com.docx.docxrenderer.service.DocxConverterService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api")
public class ConvertController {

    private final DocxConverterService converterService;

    public ConvertController(DocxConverterService converterService) {
        this.converterService = converterService;
    }

    /**
     * POST /api/convert?format=pdf   → returns application/pdf as byte stream
     * POST /api/convert?format=html  → returns text/html as a self-contained document
     * POST /api/convert?format=json  → returns application/json with html, css, images, metadata
     */
    @PostMapping(value = "/convert", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> convert(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "format", defaultValue = "pdf") String format
    ) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().body("No file provided.");
        }

        String filename = file.getOriginalFilename();
        if (filename == null || !filename.toLowerCase().endsWith(".docx")) {
            return ResponseEntity.badRequest().body("Only .docx files are supported.");
        }

        return switch (format.toLowerCase()) {
            case "pdf" -> {
                byte[] pdf = converterService.toPdf(file);
                yield ResponseEntity.ok()
                        .header(HttpHeaders.CONTENT_DISPOSITION,
                                "attachment; filename=\"" + stripExtension(filename) + ".pdf\"")
                        .contentType(MediaType.APPLICATION_PDF)
                        .body(pdf);
            }
            case "html" -> {
                String html = converterService.toHtml(file);
                yield ResponseEntity.ok()
                        .contentType(MediaType.TEXT_HTML)
                        .body(html);
            }
            case "json" -> {
                ConvertResponse response = converterService.toJson(file);
                yield ResponseEntity.ok(response);
            }
            default -> ResponseEntity.badRequest()
                    .body("Unknown format '" + format + "'. Use pdf, html, or json.");
        };
    }

    private String stripExtension(String filename) {
        int dot = filename.lastIndexOf('.');
        return dot > 0 ? filename.substring(0, dot) : filename;
    }
}