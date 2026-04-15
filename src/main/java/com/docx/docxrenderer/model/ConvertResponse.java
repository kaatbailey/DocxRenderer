package com.docx.docxrenderer.model;

import java.util.List;

public class ConvertResponse {

    private String html;
    private String css;
    private Metadata metadata;
    private List<ImageData> images;

    // ─── Constructors ────────────────────────────────────────────────────────

    public ConvertResponse() {}

    public ConvertResponse(String html, String css, Metadata metadata, List<ImageData> images) {
        this.html = html;
        this.css = css;
        this.metadata = metadata;
        this.images = images;
    }

    // ─── Getters & Setters ───────────────────────────────────────────────────

    public String getHtml() { return html; }
    public void setHtml(String html) { this.html = html; }

    public String getCss() { return css; }
    public void setCss(String css) { this.css = css; }

    public Metadata getMetadata() { return metadata; }
    public void setMetadata(Metadata metadata) { this.metadata = metadata; }

    public List<ImageData> getImages() { return images; }
    public void setImages(List<ImageData> images) { this.images = images; }

    // ─── Nested Classes ──────────────────────────────────────────────────────

    public static class Metadata {
        private String title;
        private int wordCount;
        private int paragraphCount;

        public Metadata() {}

        public Metadata(String title, int wordCount, int paragraphCount) {
            this.title = title;
            this.wordCount = wordCount;
            this.paragraphCount = paragraphCount;
        }

        public String getTitle() { return title; }
        public void setTitle(String title) { this.title = title; }

        public int getWordCount() { return wordCount; }
        public void setWordCount(int wordCount) { this.wordCount = wordCount; }

        public int getParagraphCount() { return paragraphCount; }
        public void setParagraphCount(int paragraphCount) { this.paragraphCount = paragraphCount; }
    }

    public static class ImageData {
        private String id;
        private String mimeType;
        private String data; // base64 encoded

        public ImageData() {}

        public ImageData(String id, String mimeType, String data) {
            this.id = id;
            this.mimeType = mimeType;
            this.data = data;
        }

        public String getId() { return id; }
        public void setId(String id) { this.id = id; }

        public String getMimeType() { return mimeType; }
        public void setMimeType(String mimeType) { this.mimeType = mimeType; }

        public String getData() { return data; }
        public void setData(String data) { this.data = data; }
    }
}