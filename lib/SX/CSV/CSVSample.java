import com.smartxls.WorkBook;

public class CSVSample
{

    public static void main(String args[])
    {

        WorkBook workBook = new WorkBook();
        try
        {
            workBook.read("..\\template\\book.csv");

            //workBook.write(".\\out.xls");
            workBook.writeXLSX(".\\out.xlsx");
            
            workBook.setText(0,0,"aa");
            workBook.setText(0,1,"bb");
            workBook.setText(0,2,"cc");

            workBook.writeCSV("./out.csv");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}