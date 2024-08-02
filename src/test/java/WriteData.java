import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteData {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("SheetOne");
        Object[][] data = {{"Name", "Location", "Experience"},
                {"Arun", "Osi", 18},
                {"Vanat", "Banana", 40},
                {"Sasha", "fjdfjd", 14}};

        int rows = data.length;
        int colums = data[0].length;

        for (int r =0;r<rows;r++){
            XSSFRow row = sheet.createRow(r);
            for (int c = 0;c<colums;c++){
                XSSFCell cell = row.createCell(c);
                Object cellValue=data[r][c];
                if (cellValue instanceof String){
                    cell.setCellValue((String) cellValue);
                } else if (cellValue instanceof Integer) {
                    cell.setCellValue((Integer)cellValue);
                } else if (cellValue instanceof  Boolean) {
                    cell.setCellValue((Boolean) cellValue);
                }
            }
        }
        File file = new File("src/test/Files/sheets.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        workbook.close();
        System.out.println("Completed");

    }
}
