import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

import java.time.DayOfWeek;
import java.time.LocalDateTime;
import java.util.Random;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "TemplateUse.xlsx");

        // Get template sheet.
        ExcelWorksheet templateSheet = ef.getWorksheet(0);

        // Copy template sheet.
        for (int i = 0; i < 4; i++)
            ef.getWorksheets().addCopy("Invoice " + (i + 1), templateSheet);

        // Delete template sheet.
        ef.getWorksheets().remove(0);

        LocalDateTime startTime = LocalDateTime.now();

        // Go to the first Monday from today.
        while (startTime.getDayOfWeek() != DayOfWeek.MONDAY)
            startTime = startTime.plusDays(1);

        Random rnd = new Random();

        // For each sheet.
        for (int i = 0; i < 4; i++) {
            // Get sheet.
            ExcelWorksheet ws = ef.getWorksheet(i);

            // Set some fields.
            ws.getCell("J5").setValue(14 + i);
            ws.getCell("J6").setValue(LocalDateTime.now());
            ws.getCell("J6").getStyle().setNumberFormat("m/dd/yyyy");

            ws.getCell("D12").setValue("ACME Corp");
            ws.getCell("D13").setValue("240 Old Country Road, Springfield, IL");
            ws.getCell("D14").setValue("USA");
            ws.getCell("D15").setValue("Joe Smith");

            ws.getCell("E18").setValue(startTime.toLocalDate().toString() + " until " + startTime.plusDays(11).toLocalDate().toString());

            for (int j = 0; j < 10; j++) {
                ws.getCell(21 + j, 1).setValue(startTime); // Set date.
                ws.getCell(21 + j, 1).getStyle().setNumberFormat("dddd, mmmm dd, yyyy");
                ws.getCell(21 + j, 4).setValue(rnd.nextInt(3) + 6); // Work hours.

                // Skip Saturday and Sunday.
                startTime = startTime.plusDays(j == 4 ? 3 : 1);
            }

            // Skip Saturday and Sunday.
            startTime = startTime.plusDays(2);

            ws.getCell("B36").setValue("Payment via check.");
        }

        ef.save("Sheet Copying Deleting.xlsx");
    }
}