import com.smartxls.HyperLink;
import com.smartxls.WorkBook;

public class HyperlinkWriteSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            //add a url link to F6
            workBook.addHyperlink(5,5,5,5,"http://www.smartxls.com/", HyperLink.kURLAbs,"Hello,web url hyperlink!");

            //add a file link to F7
            workBook.addHyperlink(6,5,6,5,"c:\\",HyperLink.kFileAbs,"file link");

            //workBook.write(".\\link.xls");
            workBook.writeXLSX(".\\link.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}