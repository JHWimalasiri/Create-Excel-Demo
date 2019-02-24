import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.util.*;

import java.io.FileOutputStream;

import static org.apache.poi.ss.usermodel.Font.COLOR_RED;

public class Worksheet {

    public static void main (String[] args) {

        String testData = "[{\n" +
                "  \"id\": 1,\n" +
                "  \"first_name\": \"Jeanette\",\n" +
                "  \"last_name\": \"Penddreth\",\n" +
                "  \"email\": \"jpenddreth0@census.gov\",\n" +
                "  \"gender\": \"Female\",\n" +
                "  \"ip_address\": \"26.58.193.2\"\n" +
                "}, {\n" +
                "  \"id\": 2,\n" +
                "  \"first_name\": \"Giavani\",\n" +
                "  \"last_name\": \"Frediani\",\n" +
                "  \"email\": \"gfrediani1@senate.gov\",\n" +
                "  \"gender\": \"Male\",\n" +
                "  \"ip_address\": \"229.179.4.212\"\n" +
                "}, {\n" +
                "  \"id\": 3,\n" +
                "  \"first_name\": \"Noell\",\n" +
                "  \"last_name\": \"Bea\",\n" +
                "  \"email\": \"nbea2@imageshack.us\",\n" +
                "  \"gender\": \"Female\",\n" +
                "  \"ip_address\": \"180.66.162.255\"\n" +
                "}, {\n" +
                "  \"id\": 4,\n" +
                "  \"first_name\": \"Willard\",\n" +
                "  \"last_name\": \"Valek\",\n" +
                "  \"email\": \"wvalek3@vk.com\",\n" +
                "  \"gender\": \"Male\",\n" +
                "  \"ip_address\": \"67.76.188.26\"\n" +
                "}]\n" +
                "\n";

        ArrayList output = convertToArrayList(testData);

        Worksheet wk = new Worksheet();
        wk.createWorksheet(output);

    }

    public static ArrayList convertToArrayList(String inputJson) {
        JSONArray output = new JSONArray(inputJson);
        ArrayList<JSONObject> list = new ArrayList<JSONObject>();
        for (int i = 0; i < output.length();list.add(output.getJSONObject(i++)));
        System.out.println(list.toString());
        return list;
    }


    public void createWorksheet(ArrayList data) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Report");

        // create header
        Row hrow = sheet.createRow(0);
        JSONObject hobj = (JSONObject) data.get(0);
        Iterator<String> hkeys = hobj.keys();
        int hcolNum = 0;

        // Add background color to the header
        CellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(COLOR_RED);
        style.setFillPattern(FillPatternType.BIG_SPOTS);

        while (hkeys.hasNext()) {
            Cell cell = hrow.createCell(hcolNum++);
            String hkey = hkeys.next();
            cell.setCellValue(hkey);
            cell.setCellStyle(style);
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
