import com.smartxls.WorkBook;

public class PageBreakSample
{

    public static void main(String args[])
    {
        try
        {
        WorkBook m_book = new WorkBook();
        //m_book.read("..\\template\\book.xls");
        m_book.readXLSX("..\\template\\book.xlsx");
        m_book.copyRange(8, 9, 14, 13, 1, 1, 7, 5);
        m_book.addRowPageBreak(2);
        m_book.addColPageBreak(2);
        //m_book.write(".\\pagebreak.xls");
        m_book.writeXLSX(".\\pagebreak.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}