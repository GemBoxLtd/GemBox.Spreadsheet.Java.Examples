import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Header and Footer");

        SheetHeaderFooter headerFooter = ws.getHeadersFooters();

        // Show title only on the first page
        headerFooter.getFirstPage().getHeader().getCenterSection().setContent("Title on the first page");

        // Show logo
        headerFooter.getFirstPage().getHeader().getLeftSection().appendPicture("Dices.png", 40, 40);
        headerFooter.getDefaultPage().getHeader().setLeftSection(headerFooter.getFirstPage().getHeader().getLeftSection());

        // "Page number" of "Number of pages"
        headerFooter.getFirstPage().getFooter().getRightSection().append("Page ").append(HeaderFooterFieldType.PAGE_NUMBER).append(" of ").append(HeaderFooterFieldType.NUMBER_OF_PAGES);
        headerFooter.getDefaultPage().setFooter(headerFooter.getFirstPage().getFooter());

        // Fill Sheet1 with some data
        for (int i = 0; i < 140; i++)
            for (int j = 0; j < 9; j++)
                ws.getCell(i, j).setValue(i + j);

        ef.save("Header and Footer.xlsx");
    }
}