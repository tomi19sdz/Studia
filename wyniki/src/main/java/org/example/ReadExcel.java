package org.example;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.Tomasz_malicki_123742;

public class ReadExcel {
  public static void readExcelFile(Tomasz_malicki_123742 aObject) {
    String excelFilePath = aObject.getSelectedFile(); // Uzyskanie dostępu do zmiennej selectedFile z klasy A
      try (FileInputStream fis = new FileInputStream(excelFilePath);
           Workbook workbook = new XSSFWorkbook(fis)) {

          // Otwórz pierwszy arkusz w pliku Excel
          Sheet sheet = workbook.getSheetAt(0);

          // Wyświetlanie nagłówków kolumn
          Row headerRow = sheet.getRow(0);
          for (Cell cell : headerRow) {
              System.out.print(formatCell(cell) + "\t");
          }
          System.out.println();

          // Iteracja przez wiersze w arkuszu, zaczynając od drugiego wiersza (dane)
          for (int i = 1; i <= sheet.getLastRowNum(); i++) {
              Row row = sheet.getRow(i);
              if (row != null) {
                  for (Cell cell : row) {
                      System.out.print(formatCell(cell) + "\t");
                  }
                  System.out.println();
              }
          }
      } catch (IOException e) {
          e.printStackTrace();
      }
  }

  // Pomocnicza metoda do formatowania komórek
  private static String formatCell(Cell cell) {
      switch (cell.getCellType()) {
          case STRING:
              return cell.getStringCellValue();
          case NUMERIC:
              if (DateUtil.isCellDateFormatted(cell)) {
                  return cell.getDateCellValue().toString();
              } else {
                  return String.valueOf(cell.getNumericCellValue());
              }
          case BOOLEAN:
              return String.valueOf(cell.getBooleanCellValue());
          case FORMULA:
              return cell.getCellFormula();
          default:
              return "";
      }
  }
}
