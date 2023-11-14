package ApachePOI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Write {

    String pathName;
    FileOutputStream outputStream;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFRow row;
    XSSFCell cell;
    List<Object[]> items;

    @Test
    public void WriteIntoExcel() {

        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("employees");

        items = new LinkedList<>();

        items.add(new Object[] { "EmpId", "Ename", "Job" });
        items.add(new Object[] { 101, "Scott", "Analyst" });
        items.add(new Object[] { 102, "David", "Engineer", });
        items.add(new Object[] { 103, "Smith", "Manager" });

        int rowNumber = 0;

        for (Object[] objects : items) {
            row = sheet.createRow(rowNumber++);
            int columnNumber = 0;
            for (Object object : objects) {
                cell = row.createCell(columnNumber++);

                if (object instanceof String) {
                    cell.setCellValue(object.toString());
                }
                if (object instanceof Integer) {
                    cell.setCellValue(Integer.parseInt(object.toString()));
                }
                if (object instanceof Boolean) {
                    cell.setCellValue(Boolean.parseBoolean(object.toString()));
                }
            }

        }

        pathName = ".//Data//Employee.xlsx";
        try {
            outputStream = new FileOutputStream(pathName);
            workbook.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(" Unable to write the data in the excel");
        }

    }
}
