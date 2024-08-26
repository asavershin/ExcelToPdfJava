package org.example;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;

//Нужно переключать зависимость в pom, есть free версия и не free. Free у меня не завелась.
public class SpireXlsxExample {

  public static void main(String[] args) {

    //Create a Workbook instance and load an Excel file
    Workbook workbook = new Workbook();
    workbook.loadFromFile("in.xlsx");

    //Set worksheets to fit to page when converting
    workbook.getConverterSetting().setSheetFitToPage(true);

    //Save the resulting document to a specified path
    workbook.saveToFile("out.pdf", FileFormat.PDF);
  }
}
