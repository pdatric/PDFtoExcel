import com.smartxls.WorkBook;
import com.smartxls.CommentShape;
import com.smartxls.RangeStyle;
import java.awt.*;

public class CommentWriteSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            String commentText = "comment text here!";

            //add a comment to B2
            workBook.addComment(1, 1, commentText, "author name here!");

            CommentShape commentShape = workBook.getComment(1, 1);

            //set the text to Rich-Text-Formatting
            RangeStyle rs = workBook.getRangeStyle();
            rs.setFontName("Arial");
            rs.setFontSize((short)360);
            rs.setFontBold(true);
            rs.setFontColor(Color.RED.getRGB());
            commentShape.setTextSelection(rs, 0, 7);

            rs.resetFormat();
            rs.setFontName("Cambria");
            rs.setFontSize((short)240);
            rs.setFontColor(Color.BLUE.getRGB());
            commentShape.setTextSelection(rs, 8, commentText.length());

            //workBook.write(".\\comment.xls");
            workBook.writeXLSX(".\\comment.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}