import com.smartxls.DataValidation;
import com.smartxls.WorkBook;
import com.smartxls.RangeStyle;

public class DataValidationSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            workBook.setText(1, 1, "Apple");
            workBook.setText(1, 2, "Orange");
            workBook.setText(1, 3, "Banana");

            DataValidation dataValidation = workBook.CreateDataValidation();
            dataValidation.setType(DataValidation.eUser);
            dataValidation.setFormula1("\"APPLE\0IBM\0ORACLE\"");
            workBook.setSelection("B9:D10");
            workBook.setDataValidation(dataValidation);

            dataValidation = workBook.CreateDataValidation();
            dataValidation.setType(DataValidation.eUser);
            dataValidation.setFormula1("$B$2:$D$2");
            workBook.setSelection("B2:D6");
            workBook.setDataValidation(dataValidation);

            workBook.setText(0, 1, "Data validation source from B2:D6");
            RangeStyle rs = workBook.getRangeStyle();
            rs.setPattern(RangeStyle.PatternSolid);
            rs.setPatternFG(java.awt.Color.magenta.getRGB());
            workBook.setRangeStyle(rs, 1,1,5,3);

            workBook.setText(7, 1, "Data validation source from text");
            rs = workBook.getRangeStyle();
            rs.setPattern(RangeStyle.PatternSolid);
            rs.setPatternFG(java.awt.Color.blue.getRGB());
            workBook.setRangeStyle(rs, 8,1,9,3);


            //workBook.write(".\\datavalidation.xls");
            workBook.writeXLSX(".\\datavalidation.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}