import com.smartxls.*;

public class PivotSample
{

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            //workBook.read("..\\template\\pivotTemplate.xls");
            workBook.readXLSX("..\\template\\pivotTemplate.xlsx");

//            BookPivotRange pivotRange = workBook.addPivotRange("A1:D27", "F7");
            BookPivotRange pivotRange = workBook.addPivotRange(0, 0, 0, 26, 3, 0, 7, 5);
            BookPivotArea rowArea = pivotRange.getArea(BookPivotRange.row);
            BookPivotArea columnArea = pivotRange.getArea(BookPivotRange.column);
            BookPivotArea dataArea = pivotRange.getArea(BookPivotRange.data);
            BookPivotArea pageArea = pivotRange.getArea(BookPivotRange.page);

            BookPivotField rowField = pivotRange.addField("Who", rowArea);
            BookPivotField dataField = pivotRange.addField("Amount", dataArea);
            BookPivotField columnField = pivotRange.addField("What", columnArea);
            BookPivotField pageField = pivotRange.addField("Week", pageArea);


            //workBook.write(".\\PivotTable.xls");
            workBook.writeXLSX(".\\PivotTable.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}