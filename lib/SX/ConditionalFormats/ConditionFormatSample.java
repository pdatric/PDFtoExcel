import com.smartxls.ConditionFormat;
import com.smartxls.RangeStyle;
import com.smartxls.WorkBook;

public class ConditionFormatSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            ConditionFormat condfmt[]=new ConditionFormat[3];
            condfmt[0]=workBook.CreateConditionFormat();
            condfmt[1]=workBook.CreateConditionFormat();
            condfmt[2]=workBook.CreateConditionFormat();

            // Condition #1
            RangeStyle cf=condfmt[0].getRangeStyle();
            condfmt[0].setType(ConditionFormat.eTypeFormula);
            condfmt[0].setFormula1("and(iseven(row()), $E1 > 1000)", 0, 0);
            cf.setFontColor(0x00ff00);
            cf.setPattern(RangeStyle.PatternSolid);
            cf.setPatternFG(0xcc99ff);
            condfmt[0].setRangeStyle(cf);

            // Condition #2
            condfmt[1].setType(ConditionFormat.eTypeFormula);
            condfmt[1].setFormula1("iseven($B1)", 0, 0);
            cf.setFontColor(0xffffff);
            condfmt[1].setRangeStyle(cf);

            // Condition #3
            condfmt[2].setType(ConditionFormat.eTypeCell);
            condfmt[2].setFormula1("500", 0, 0);
            condfmt[2].setOperator(ConditionFormat.eOperatorGreaterThan);
            cf=condfmt[2].getRangeStyle();
            cf.setFontColor(0xff0000);
            condfmt[2].setRangeStyle(cf);

            // Select the range and apply conditional formatting
            workBook.setSelection(1, 1, 39, 4);
            workBook.setConditionalFormats(condfmt);

            workBook.setText(0, 1, "iseven");
            workBook.setText(0, 4, ">1000");

            workBook.setNumber(5, 1, 3);
            workBook.setNumber(4, 4, 1200);
            workBook.setSelection(5, 1, 5, 1);
        
            //workBook.write("./conditionFormats.xls");
            workBook.writeXLSX("./conditionFormats.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}