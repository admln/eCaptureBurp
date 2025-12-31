package com.ecapture.burp.export;

import com.ecapture.burp.event.CapturedEvent;
import com.ecapture.burp.event.MatchedHttpPair;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Base64;
import java.util.List;

/**
 * Minimal HAR serializer producing HAR 1.2 compatible JSON.
 */
public class HarSerializer {

    private static final Gson GSON = new GsonBuilder().setPrettyPrinting().create();

    public static void writeHar(OutputStream out, List<MatchedHttpPair> pairs, List<String> columns) throws Exception {
        JsonObject root = new JsonObject();
        JsonObject log = new JsonObject();
        log.addProperty("version", "1.2");
        JsonObject creator = new JsonObject();
        creator.addProperty("name", "eCaptureBurp");
        creator.addProperty("version", "1.0");
        log.add("creator", creator);

        JsonArray entries = new JsonArray();

        DateTimeFormatter fmt = DateTimeFormatter.ISO_INSTANT.withZone(ZoneOffset.UTC);

        for (MatchedHttpPair pair : pairs) {
            JsonObject entry = new JsonObject();
            long startedMillis = pair.getCreatedAt();
            entry.addProperty("startedDateTime", fmt.format(Instant.ofEpochMilli(startedMillis)));
            entry.addProperty("time", 0);

            // request
            JsonObject request = new JsonObject();
            request.addProperty("method", pair.getMethod());
            request.addProperty("url", pair.getUrl());
            request.addProperty("httpVersion", "HTTP/1.1");
            // headers - best effort: include Host
            JsonArray reqHeaders = new JsonArray();
            JsonObject hostHeader = new JsonObject();
            hostHeader.addProperty("name", "Host");
            hostHeader.addProperty("value", pair.getHost());
            reqHeaders.add(hostHeader);
            request.add("headers", reqHeaders);
            request.addProperty("headersSize", -1);
            request.addProperty("bodySize", pair.getRequestLength());

            // postData if available
            if (pair.getRequest() != null && pair.getRequest().getPayload() != null) {
                byte[] payload = pair.getRequest().getPayload();
                String text = null;
                JsonObject postData = new JsonObject();
                boolean binary = false;
                try {
                    text = new String(payload, StandardCharsets.UTF_8);
                    // crude check for non-text
                    if (!isMostlyText(payload)) {
                        binary = true;
                    }
                } catch (Exception e) {
                    binary = true;
                }
                if (binary) {
                    postData.addProperty("mimeType", "application/octet-stream");
                    postData.addProperty("text", Base64.getEncoder().encodeToString(payload));
                    postData.addProperty("encoding", "base64");
                } else {
                    postData.addProperty("mimeType", "text/plain");
                    postData.addProperty("text", text);
                }
                request.add("postData", postData);
            }

            // response
            JsonObject response = new JsonObject();
            response.addProperty("status", parseStatus(pair.getStatusCode()));
            response.addProperty("statusText", "");
            response.addProperty("httpVersion", "HTTP/1.1");
            response.add("headers", new JsonArray());
            response.addProperty("headersSize", -1);
            response.addProperty("bodySize", pair.getResponseLength());
            if (pair.getResponse() != null && pair.getResponse().getPayload() != null) {
                byte[] payload = pair.getResponse().getPayload();
                boolean binary = !isMostlyText(payload);
                JsonObject content = new JsonObject();
                content.addProperty("size", payload.length);
                content.addProperty("mimeType", "application/octet-stream");
                if (binary) {
                    content.addProperty("text", Base64.getEncoder().encodeToString(payload));
                    content.addProperty("encoding", "base64");
                } else {
                    content.addProperty("text", new String(payload, StandardCharsets.UTF_8));
                }
                response.add("content", content);
            }

            entry.add("request", request);
            entry.add("response", response);
            entries.add(entry);
        }

        log.add("entries", entries);
        root.add("log", log);

        String json = GSON.toJson(root);
        out.write(json.getBytes(StandardCharsets.UTF_8));
    }

    private static boolean isMostlyText(byte[] data) {
        int printable = 0;
        int total = Math.min(data.length, 200);
        for (int i = 0; i < total; i++) {
            int b = data[i] & 0xff;
            if (b >= 32 && b <= 126) printable++;
            if (b == '\n' || b == '\r' || b == '\t') printable++;
        }
        return total == 0 || ((double) printable / total) > 0.8;
    }

    private static int parseStatus(String statusStr) {
        try {
            return Integer.parseInt(statusStr);
        } catch (Exception e) {
            return 0;
        }
    }
}

