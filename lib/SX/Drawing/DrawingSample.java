import com.smartxls.AutoShape;
import com.smartxls.FormControlShape;
import com.smartxls.WorkBook;

public class DrawingSample
{

    public static void main(String args[])
    {

        WorkBook workBook = new WorkBook();
        try
        {
            workBook.setText(0,0,"aa");
            workBook.setText(1,0,"bb");
            workBook.setText(2,0,"cc");

            FormControlShape checkBoxShape = workBook.addFormControl(3.0, 1.0, 5.0, 2.1, FormControlShape.CheckBox);
            checkBoxShape.setCellRange("A1:A3");
            checkBoxShape.setCellLink("B2");

            checkBoxShape.setText("checkbox1");

            FormControlShape comBoxShape1 = workBook.addFormControl(3.0, 3.0, 4.1, 4.1, FormControlShape.CombBox);
            comBoxShape1.setCellRange("A1:A3");
            comBoxShape1.setCellLink("B4");
     
       FormControlShape formControlShape2 = workBook.addFormControl(11.0, 2.0, 13.4, 5.5, FormControlShape.ListBox);
            formControlShape2.setCellRange("A1:A3");
            formControlShape2.setCellLink("B4");

            AutoShape rectangleShape = workBook.addAutoShape(1.0, 9.0, 3.0, 11.0, AutoShape.Rectangle);
            AutoShape textBoxShape = workBook.addAutoShape(3.0, 5.0, 5.0, 7.0, AutoShape.TextBox);
            textBoxShape.setText("textBox ddd");
            AutoShape ovalShape = workBook.addAutoShape(3.5, 9.0, 5.0, 11.0, AutoShape.Oval);
            AutoShape lineShape = workBook.addAutoShape(6.0, 9.0, 8.0, 11.0, AutoShape.Line);

            //workBook.write("./drawings.xls");
            workBook.writeXLSX("./drawings.xlsx");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}