import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class ExcelSheet {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(new FileInputStream(new File("gt.xlsx")));

        // An output stream accepts output bytes and sends them to sink.

        OutputStream fileOut = new FileOutputStream("gt.xlsx");

        Sheet sheet1 = wb.getSheetAt(0);
        int x = 10050;
        String s = "50LT";

     /*   Row sample = sheet1.createRow(4);
        Cell ssss= sample.createCell(0);
        ssss.setCellValue("HELlo");*/


        int rowCount = sheet1.getLastRowNum();
        for(int i=1;i<=50;i++){
            Row row = sheet1.createRow(++rowCount) ;
            Cell cell = row.createCell(0);

            cell.setCellValue(s+x);
            x++;

        }

        wb.write(fileOut);
        System.out.println("CREATED");
    }
}
