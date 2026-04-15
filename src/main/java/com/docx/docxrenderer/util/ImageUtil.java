package com.docx.docxrenderer.util;

import net.coobird.thumbnailator.Thumbnails;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;

public class ImageUtil {

    // Maximum dimension for any image — width or height
    private static final int MAX_DIMENSION = 1200;

    // JPEG quality — 0.75 is a good balance between size and clarity
    private static final double JPEG_QUALITY = 0.75;

    /**
     * Takes raw image bytes from a .docx picture part and returns
     * compressed bytes capped at MAX_DIMENSION on either axis.
     *
     * Used by both PdfConversionStrategy (embeds bytes directly)
     * and JsonConversionStrategy (base64 encodes the result).
     */
    public static byte[] compress(byte[] rawBytes) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        Thumbnails.of(new ByteArrayInputStream(rawBytes))
                .size(MAX_DIMENSION, MAX_DIMENSION)
                .keepAspectRatio(true)
                .outputQuality(JPEG_QUALITY)
                .outputFormat("jpg")
                .toOutputStream(out);

        return out.toByteArray();
    }

    /**
     * Convenience method for the JSON strategy —
     * compresses and returns a base64 encoded string
     * ready to drop into the ImageData DTO.
     */
    public static String compressToBase64(byte[] rawBytes) throws IOException {
        return Base64.getEncoder().encodeToString(compress(rawBytes));
    }
}