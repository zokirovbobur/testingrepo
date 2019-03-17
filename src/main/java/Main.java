import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        ExcelReader excelReader = new ExcelReader();
        excelReader.detectBook("Books.xlsx");
    }
}
