import com.smartxls.WorkBook;

public class AutofiltersSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            //set data
            workBook.setText(0,1,"Jan");
            workBook.setText(0,2,"Feb");
            workBook.setText(0,3,"Mar");
            workBook.setText(0,4,"Apr");
            workBook.setText(0,5,"Jun");

            workBook.setText(1,0,"Comfrey");
            workBook.setText(2,0,"Bananas");
            workBook.setText(3,0,"Papaya");
            workBook.setText(4,0,"Mango");
            workBook.setText(5,0,"Lilikoi");
            for(int col = 1; col <= 5; col++)
                for(int row = 1; row <= 5; row++)
                    workBook.setFormula(row, col, "RAND()");
            workBook.setText(6, 0, "Total");
            workBook.setFormula(6, 1, "SUM(B2:B6)");
            workBook.setSelection("B7:F7");
            //auto fill the range with the first cell's formula or data
            workBook.editCopyRight();

            //select range A1:F7
            workBook.setSelection(0,0,6,5);
            //Creating an AutoFilter
            workBook.autoFilter();

            //Counting the auto filtered value in the cell "E11"
            workBook.setFormula(10, 4, "SUBTOTAL(2,B1:B7)");

            //workBook.write("./autofilter.xls");
            workBook.writeXLSX("./autofilter.xlsx");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}