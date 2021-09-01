import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "Excel 2010.xlsx");

        // Modify all values in column C. Set them to some random value between -10 and 10.
        CellRangeIterator readEnumerator = ef.getWorksheet(0).getColumn("C").getCells().getReadIterator();

        java.util.Random rnd = new java.util.Random();
        while (readEnumerator.hasNext()) {
            ExcelCell cell = readEnumerator.next();
            if (cell.getValueType() == CellValueType.INT)
                cell.setValue(rnd.nextInt(20) - 10);
        }

        ef.save("Excel 2010_2013 Features.xlsx");
    }
}