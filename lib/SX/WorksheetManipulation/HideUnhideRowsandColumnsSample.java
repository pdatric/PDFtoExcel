import com.smartxls.WorkBook;

public class HideUnhideRowsandColumnsSample
{

    public static void main(String args[])
    {
        try
        {
        WorkBook m_book = new WorkBook();
        //m_book.read("..\\template\\book.xls");
        m_book.readXLSX("..\\template\\book.xlsx");
        m_book.setRowHidden(1,true);
        m_book.setColHidden(1,true);
        //m_book.write(".\\rowcolhide.xls");
        m_book.writeXLSX(".\\rowcolhide.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}