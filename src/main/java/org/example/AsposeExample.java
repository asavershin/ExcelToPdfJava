package org.example;

import com.aspose.cells.PdfCompliance;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import lombok.SneakyThrows;

public class AsposeExample {
  @SneakyThrows
  public static void main(String[] args) {
    // Create Workbook to load Excel file
    Workbook workbook = new Workbook("dfs.xlsx");

// Create PDF options
    PdfSaveOptions options = new PdfSaveOptions();
    options.setCompliance(PdfCompliance.PDF_A_1_A);

// Save the document in PDF format
    workbook.save("out.pdf", options);
  }
}
