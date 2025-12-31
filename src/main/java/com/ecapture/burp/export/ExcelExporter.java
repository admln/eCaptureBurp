package com.ecapture.burp.export;

import com.ecapture.burp.event.MatchedHttpPair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * Simple Excel exporter using Apache POI.
 */
public class ExcelExporter {

    public static void writeXlsx(OutputStream out, List<MatchedHttpPair> pairs, List<String> columns) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            CreationHelper createHelper = wb.getCreationHelper();
            Sheet sheet = wb.createSheet("eCapture");

            int rowIdx = 0;
            // Header
            Row header = sheet.createRow(rowIdx++);
            for (int c = 0; c < columns.size(); c++) {
                Cell cell = header.createCell(c);
                cell.setCellValue(columns.get(c));
            }

            // Rows
            for (MatchedHttpPair pair : pairs) {
                Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < columns.size(); c++) {
                    String col = columns.get(c);
                    Cell cell = r.createCell(c);
                    switch (col) {
                        case "#":
                            cell.setCellValue(pair.getUuid());
                            break;
                        case "Time":
                            cell.setCellValue(pair.getTimestamp());
                            break;
                        case "Method":
                            cell.setCellValue(pair.getMethod());
                            break;
                        case "Host":
                            cell.setCellValue(pair.getHost());
                            break;
                        case "URL":
                            cell.setCellValue(pair.getUrl());
                            break;
                        case "Status":
                            cell.setCellValue(pair.getStatusCode());
                            break;
                        case "Req Len":
                            cell.setCellValue(pair.getRequestLength());
                            break;
                        case "Resp Len":
                            cell.setCellValue(pair.getResponseLength());
                            break;
                        case "Process":
                            cell.setCellValue(pair.getProcessInfo());
                            break;
                        case "Complete":
                            cell.setCellValue(pair.isComplete() ? "✓" : "...");
                            break;
                        case "Request Body":
                            if (pair.getRequest() != null && pair.getRequest().getPayload() != null) {
                                byte[] p = pair.getRequest().getPayload();
                                cell.setCellValue(new String(p, StandardCharsets.UTF_8));
                            } else {
                                cell.setCellValue("");
                            }
                            break;
                        case "Response Body":
                            if (pair.getResponse() != null && pair.getResponse().getPayload() != null) {
                                byte[] p = pair.getResponse().getPayload();
                                cell.setCellValue(new String(p, StandardCharsets.UTF_8));
                            } else {
                                cell.setCellValue("");
                            }
                            break;
                        default:
                            cell.setCellValue("");
                    }
                }
            }

            // Auto-size first few columns
            for (int c = 0; c < Math.min(10, columns.size()); c++) {
                sheet.autoSizeColumn(c);
            }

            wb.write(out);
        }
    }
}

