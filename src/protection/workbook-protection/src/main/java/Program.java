import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Sheet Protection");

        ws.getCell(0, 2).setValue("Only cells from A1 to A10 are editable.");

        for (int i = 0; i < 10; i++) {
            ExcelCell cell = ws.getCell(i, 0);
            cell.setValue(i);
            cell.getStyle().setLocked(false);
        }

        ws.setProtected(true);

        // ProtectionSettings class is supported only for XLSX file format.
        ws.getCell(2, 2).setValue("Inserting columns is allowed (only supported for XLSX file format).");
        WorksheetProtection protectionSettings = ws.getProtectionSettings();
        protectionSettings.setAllowInsertingColumns(true);

        ws.getCell(3, 2).setValue("Sheet password is 123 (only supported for XLSX file format).");
        protectionSettings.setPassword("123");

        ef.save("Sheet Protection.xlsx");
    }
}