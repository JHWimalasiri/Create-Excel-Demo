import org.json.JSONArray;
import org.json.JSONObject;

import java.util.ArrayList;

public class Main {




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

}
