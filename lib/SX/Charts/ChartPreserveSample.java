import com.smartxls.WorkBook;

public class ChartPreserveSample
{

    public static void main(String args[])
    {
        try
        {
            WorkBook m_book = new WorkBook();
            //m_book.read("..\\template\\ChartPropertiesTemplate.xls");
            m_book.readXLSX("..\\template\\ChartPropertiesTemplate.xlsx");
            m_book.copyRange(8, 9, 14, 13, 1, 1, 7, 5);
            m_book.addRowPageBreak(2);
            m_book.addColPageBreak(2);
            //m_book.write(".\\chartProp.xls");
            m_book.writeXLSX(".\\chartProp.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}