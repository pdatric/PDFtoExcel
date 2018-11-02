import com.smartxls.WorkBook;

public class MultiSheetTestSample
{

    public static void main(String args[])
    {
        try
        {
        WorkBook m_book = new WorkBook();
        m_book.setSheetName(0, "sheet1");
        m_book.setText(1, 1, "sheet1");

        m_book.insertSheets(1, 1);
        m_book.setSheetName(1, "sheet2");
        m_book.setText(2, 2, "sheet2");
        m_book.insertSheets(2, 1);
        m_book.setSheetName(2, "sheet3");
        m_book.setText(3, 3, "sheet3");
        m_book.setText(1, 3, "sheet3");

        m_book.setSheet(2);
        m_book.setText(3, 3, "sheet1");
        m_book.setSheet(1);
        m_book.setText(3, 3, "sheet2");

        //m_book.write(".\\multisheet.xls");
        m_book.writeXLSX(".\\multisheet.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}