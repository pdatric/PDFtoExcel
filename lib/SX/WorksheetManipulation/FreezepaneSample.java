import com.smartxls.WorkBook;

public class FreezepaneSample
{

    public static void main(String args[])
    {
        try
        {
        WorkBook m_book = new WorkBook();
        //m_book.read("..\\template\\book.xls");
        m_book.readXLSX("..\\template\\book.xlsx");
        m_book.copyRange(8, 9, 14, 13, 1, 1, 7, 5);
        m_book.freezePanes(0, 0, 2, 2, true);
        //m_book.write(".\\freezepane.xls");
        m_book.writeXLSX(".\\freezepane.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}