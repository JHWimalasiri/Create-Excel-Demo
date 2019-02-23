import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.util.*;

import java.io.FileOutputStream;

public class Worksheet {

    public void createWorksheet(ArrayList data) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Report");


        Iterator<JSONObject> iterator = data.iterator();

        // create header
        Row hrow = sheet.createRow(0);
        JSONObject hobj = (JSONObject) data.get(0);
        Iterator<String> hkeys = hobj.keys();
        int hcolNum = 0;

        while (hkeys.hasNext()) {
            Cell cell = hrow.createCell(hcolNum++);
            String hkey = hkeys.next();
            cell.setCellValue(hkey);
        }

        int rowNum = 1;
        // create rows
        for (Object obj: data) {
            int colNum = 0;
            Row row = sheet.createRow(rowNum++);
            JSONObject object = (JSONObject) obj;
            Iterator<String> keys = object.keys();

            while (keys.hasNext()) {
                Cell cell = row.createCell(colNum++);
                String key = keys.next();
                if (object.get(key) instanceof String) {
                    cell.setCellValue((String) object.get(key));
                } else if (object.get(key) instanceof Number) {
                    cell.setCellValue((Integer) object.get(key));
                } else if (object.get(key) == null) {
                    cell.setCellValue("");
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream("Reports/output.xlsx");
            workbook.write(outputStream);
            workbook.close();
        }

        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
