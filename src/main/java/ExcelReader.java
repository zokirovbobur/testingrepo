import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

public class ExcelReader {
    public static final String PROFESSORS_XLSX_FILE_PATH = "Professors.xlsx";

    public void formatFullName(String fullName){
        String fName = fullName.substring(0,fullName.indexOf(" "));
        fName = fName.substring(1).toLowerCase();
        fullName=fullName.replace(fName,"");
        String lName = fullName.substring(1,fullName.length());
        lName = lName.substring(1).toLowerCase();
        System.out.println("formatted string: " + fName +" " + lName);
    }

    public void detectBook(String path) throws IOException, InvalidFormatException {

        Workbook workbook = WorkbookFactory.create(new File(path));

        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            Book book = new Book();
            book.setAuthor(dataFormatter.formatCellValue(row.getCell(2)));
            book.setBookPicUrl(dataFormatter.formatCellValue(row.getCell(7)));
            book.setSubject(dataFormatter.formatCellValue(row.getCell(0)));
            book.setBookUrl(dataFormatter.formatCellValue(row.getCell(5)));


            System.out.println(book.toString());
        });

        // Closing the workbook
        workbook.close();
    }


}
