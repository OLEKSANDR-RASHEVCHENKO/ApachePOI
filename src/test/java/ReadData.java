import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadData {
    public static void main(String[] args) throws IOException {
//            String excelFilePath = System.getProperty("src/test/Files/employyes.xlsx");
        File excelFile = new File("src/test/Files/employyes.xlsx");
        FileInputStream fis = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();

        // first

        while (iterator.hasNext()) {
            Row row = iterator.next();
            Iterator<Cell> cellIterator = row.iterator();
            while (cellIterator.hasNext()) {
                Cell cell=cellIterator.next();
                CellType cellType = cell.getCellType();
                switch (cellType){
                    case STRING :
                        System.out.print(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                }
                System.out.print("  ");

            }
            System.out.println();
            workbook.close();
        }
//second
            /*
            int rows=sheet.getLastRowNum();
            int cols= sheet.createRow(1).getLastCellNum();
            for (int r=0;r<rows;r++){
                XSSFRow row = sheet.getRow(r);
                for ( int c = 0;c<cols;c++){
                    XSSFCell cell = row.getCell(c);
                    CellType cellType = cell.getCellType();
                    switch (cellType){
                        case STRING :
                            System.out.println(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                    }
                }
            */
    }
}

