import com.smartxls.WorkBook;

public class RowHeightandColumnWidthSample
{

    public static void main(String args[])
    {
        try
        {
        WorkBook m_book = new WorkBook();
        //m_book.read("..\\template\\book.xls");
        m_book.readXLSX("..\\template\\book.xlsx");

        m_book.setRowHeight(1, 25 * 20);
        m_book.setColWidth(1, 25 * 256);

        //m_book.write(".\\rowheightcolwidth.xls");
        m_book.writeXLSX(".\\rowheightcolwidth.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    RowHeightandColumnWidthSample()
    {
    }

    void run()
            throws Exception
    {
    }
}