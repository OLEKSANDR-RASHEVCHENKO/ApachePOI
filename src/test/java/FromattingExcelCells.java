import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class FromattingExcelCells {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("SheetOne");
        String[][] data = {{"Name", "Location"},
                {"Arun", "Osi"},
                {"Vanat", "Banana"},
                {"Sasha", "fjdfjd"}};
        int rows = data.length;
        int cols = data[0].length;
        XSSFCellStyle style1 = workbook.createCellStyle();
        style1.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        style1.setFillPattern(FillPatternType.DIAMONDS);
        style1.setBorderLeft(BorderStyle.THICK);
        style1.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        style1.setBorderRight(BorderStyle.THICK);
        style1.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style1.setBorderTop(BorderStyle.THICK);
        style1.setTopBorderColor(IndexedColors.GREEN.getIndex());
        style1.setBorderBottom(BorderStyle.THICK);
        style1.setTopBorderColor(IndexedColors.GREEN.getIndex());

        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setBorderLeft(BorderStyle.THICK);
        style2.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        style2.setBorderRight(BorderStyle.THICK);
        style2.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style2.setBorderTop(BorderStyle.THICK);
        style2.setTopBorderColor(IndexedColors.GREEN.getIndex());
        style2.setBorderBottom(BorderStyle.THICK);
        style2.setTopBorderColor(IndexedColors.GREEN.getIndex());


        for (int r = 0;r<rows;r++){
            XSSFRow row = sheet.createRow(r);
            for (int c = 0;c<cols;c++){
                XSSFCell cell = row.createCell(c);
                String cellValue=data[r][c];
                cell.setCellValue(cellValue);
                if (r==0) {
                    cell.setCellStyle(style1);
                }else{
                    cell.setCellStyle(style2);
                }
            }
        }
        File file = new File("src/test/Files/sheetsBaground.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        workbook.close();
        System.out.println("Completed");
    }
}
