package org.example;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePoiItextPdfExample {

  @SneakyThrows
  public static void main(String[] args) {
    String excelFilePath = "in.xlsx";
    String pdfFilePath = "out.pdf";
    String fontPath = "DejaVuSans.ttf";
    BaseFont baseFont = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
    Font font = new Font(baseFont, 6);

    try (FileInputStream excelFile = new FileInputStream(excelFilePath);
         Workbook workbook = new XSSFWorkbook(excelFile);
         FileOutputStream pdfFile = new FileOutputStream(pdfFilePath)) {

      Document document = new Document();
      PdfWriter.getInstance(document, pdfFile);
      document.open();

      Sheet sheet = workbook.getSheetAt(0);

      for (Row row : sheet) {
        boolean isTableRow = false;

        for (Cell cell : row) {
          CellStyle style = cell.getCellStyle();
          if (style.getBorderTop() != BorderStyle.NONE || style.getBorderBottom() != BorderStyle.NONE ||
                  style.getBorderLeft() != BorderStyle.NONE || style.getBorderRight() != BorderStyle.NONE) {
            isTableRow = true;
            break;
          }
        }

        if (isTableRow) {
          int firstNonEmptyCol = 0;
          int lastNonEmptyCol = row.getLastCellNum() - 1;

          // Найдем первый непустой столбец
          while (firstNonEmptyCol <= lastNonEmptyCol &&
                  (row.getCell(firstNonEmptyCol) == null || row.getCell(firstNonEmptyCol).getCellType() == CellType.BLANK)) {
            firstNonEmptyCol++;
          }

          while (lastNonEmptyCol >= firstNonEmptyCol &&
                  (row.getCell(lastNonEmptyCol) == null || row.getCell(lastNonEmptyCol).getCellType() == CellType.BLANK)) {
            lastNonEmptyCol--;
          }

          if (firstNonEmptyCol <= lastNonEmptyCol) {
            PdfPTable table = new PdfPTable(lastNonEmptyCol - firstNonEmptyCol + 1);

            for (int colIndex = firstNonEmptyCol; colIndex <= lastNonEmptyCol; colIndex++) {
              Cell cell = row.getCell(colIndex);
              if (cell != null) {
                PdfPCell pdfCell = new PdfPCell(new Paragraph(cell.toString(), font));

                for (int j = 0; j < sheet.getNumMergedRegions(); j++) {
                  CellRangeAddress region = sheet.getMergedRegion(j);
                  if (region.isInRange(cell.getRowIndex(), colIndex)) {
                    int colspan = region.getLastColumn() - region.getFirstColumn() + 1;
                    pdfCell.setColspan(colspan);
                    colIndex += colspan - 1;
                    break;
                  }
                }

                table.addCell(pdfCell);
              } else {
                table.addCell(new PdfPCell(new Paragraph("", font)));
              }
            }

            document.add(table);
          }
        } else {
          StringBuilder rowData = new StringBuilder();
          for (Cell cell : row) {
            if (!cell.getCellType().equals(CellType.BLANK)) {
              rowData.append(cell.toString()).append(" ");
            }
          }
          if (rowData.length() > 0) {
            document.add(new Paragraph(rowData.toString(), font));
          }
        }
      }

      document.close();
      System.out.println("PDF успешно создан!");

    } catch (IOException | DocumentException e) {
      e.printStackTrace();
    }
  }
}


