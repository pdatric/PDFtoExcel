import com.smartxls.*;
import com.smartxls.*;
import com.smartxls.enums.TableBuiltInStyles;

public class TableSample
{

    public static void main(String args[])
    {
        try
        {

            String tableDataInCSV = "Sales Person\tRegion\tSales Amount\t% Commission\n" +
            "Joe\tNorth\t260\t10%\n" +
            "Robert\tSouth\t660\t15%\n" +
            "Michelle\tEast\t940\t15%\n" +
            "Erich\tWest\t410\t12%\n" +
            "Dafna\tNorth\t800\t15%\n" +
            "Rob\tSouth\t900\t15%";
            WorkBook workbook = new WorkBook();
//import table data
            workbook.setCSVString(tableDataInCSV);

//add table to range A1:D7
            Table table = workbook.addTable("DeptSales", 0, 0, 6, 3, TableBuiltInStyles.TableStyleMedium2);
//banded row
            table.setRowStripes(true);
//enable total row
            table.setTotalRow(true, "Total");

//add new column(with structured references formula) to the table
            table.addCalculatedColumn("Commission Amount", "DeptSales[[#This Row],[Sales Amount]]*DeptSales[[#This Row],[% Commission]]");
//set the column's total func(1-average 2-count 3-countNums 4-max 5-min 6-sum 7-stdDev 8-var)
            table.setColumnTotalsFunc("Commission Amount", 6);

            workbook.setText(11,1,"Sales Total");
//using structured references formula in sheet 
            workbook.setFormula(11,2,"SUM(DeptSales[Sales Amount])");

            for(int i=0;i<5;i++)
            workbook.setColWidthAutoSize(i, true);

            workbook.writeXLSX("Table.xlsx");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}