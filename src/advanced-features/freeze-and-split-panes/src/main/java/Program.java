import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        // Frozen Rows (first 2 rows are frozen)
        ExcelWorksheet ws1 = ef.addWorksheet("Frozen rows");
        ws1.setPanes(new WorksheetPanes(PanesState.FROZEN, 0, 2, "A3", PanePosition.BOTTOM_LEFT));

        // Frozen Columns (first column is frozen)
        ExcelWorksheet ws2 = ef.addWorksheet("Frozen columns");
        ws2.setPanes(new WorksheetPanes(PanesState.FROZEN, 1, 0, "B1", PanePosition.TOP_RIGHT));

        // Frozen Rows and Columns (first 2 rows and first 3 columns are frozen)
        ExcelWorksheet ws3 = ef.addWorksheet("Frozen rows and columns");
        ws3.setPanes(new WorksheetPanes(PanesState.FROZEN, 3, 2, "E5", PanePosition.BOTTOM_RIGHT));

        // Split pane
        ExcelWorksheet ws4 = ef.addWorksheet("Split pane");
        ws4.setPanes(new WorksheetPanes(PanesState.SPLIT, 2310, 1500, "D7", PanePosition.BOTTOM_RIGHT));

        ef.save("Freeze or Split Panes.xlsx");
    }
}