import com.smartxls.RangeStyle;
import com.smartxls.WorkBook;

public class TextFormattingSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            //workBook.read("..\\template\\book.xls");
            workBook.readXLSX("..\\template\\book.xlsx");
 
            //a double value representing the number of days since 1/1/1900.Fractional days represent hours,minutes,and seconds.
            double dd = workBook.getNumber(12,2);
            System.out.println("days since 1900:" + dd);
            String ds = workBook.getText(12,2);
            //the formatted text as it is showed in Excel.
            String dfs = workBook.getFormattedText(12,2);
            System.out.println("Formatted text:" + dfs);

            //set cell value with formatted text(mm/dd/yy).
            workBook.setEntry(13,2,"08/08/2009");
            //set the cell with number value.
            workBook.setNumber(14,2,40033.0);
            //formatting the value to date 'yyyy/mm/dd'
            RangeStyle rs = workBook.getRangeStyle();
            rs.setCustomFormat("yyyy/mm/dd");
            workBook.setRangeStyle(rs, 14,2,14,2);

            //set richText data
            workBook.setText(0,0,"Hello, you are welcome!");
            //select the cell which it's text will be formatted.
            workBook.setSelection(0,0,0,0);

            //text orientation
            RangeStyle rangeStyle = workBook.getRangeStyle();
            rangeStyle.setOrientation((short)35);
            workBook.setRangeStyle(rangeStyle);

            //multi text selection format
            workBook.setTextSelection(0, 6);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.setFontBold(true);
            rangeStyle.setFontColor(java.awt.Color.blue.getRGB());
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(7, 10);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.setFontItalic(true);
            rangeStyle.setFontColor(java.awt.Color.magenta.getRGB());
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(11, 14);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.setFontUnderline(RangeStyle.UnderlineSingle);
            rangeStyle.setFontColor(java.awt.Color.green.getRGB());
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(15, 22);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.setFontSize(14*20);
            rangeStyle.setFontColor(java.awt.Color.red.getRGB());
            rangeStyle.setFontSuperscript(true);
            workBook.setRangeStyle(rangeStyle);

            //workBook.write("./TextFormatting.xls");
            workBook.writeXLSX("./TextFormatting.xlsx");
		} 
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}