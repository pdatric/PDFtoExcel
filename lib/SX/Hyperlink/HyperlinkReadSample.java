import com.smartxls.HyperLink;
import com.smartxls.WorkBook;

public class HyperlinkReadSample
{

    public static void main(String args[])
    {

        WorkBook workBook = new WorkBook();
        String version = workBook.getVersionString();
        System.out.println("Ver:" + version);
        try
        {
            //workBook.read("..\\template\\book.xls");
            workBook.readXLSX("..\\template\\book.xlsx");

            // get the first index from the current sheet
            HyperLink hyperLink = workBook.getHyperlink(0);
            if(hyperLink != null)
            {
                System.out.println(hyperLink.getType());
                System.out.println(hyperLink.getLinkURLString());
                System.out.println(hyperLink.getLinkShowString());
                System.out.println(hyperLink.getToolTipString());
                System.out.println(hyperLink.getRange());
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}