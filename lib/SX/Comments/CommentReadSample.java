import com.smartxls.CommentShape;
import com.smartxls.WorkBook;

public class CommentReadSample
{

    public static void main(String args[])
    {

        WorkBook workBook = new WorkBook();
        try
        {
            //workBook.read("..\\template\\book.xlsx");
            workBook.readXLSX("..\\template\\book.xlsx");

            // get the first index from the current sheet
            CommentShape commentShape = workBook.getComment(1, 1);
            if(commentShape != null)
            {
                System.out.println("comment text:" + commentShape.getText());
                System.out.println("comment author:" + commentShape.getAuthor());
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}