import com.smartxls.ChartShape;
import com.smartxls.WorkBook;

public class ChartsEditSample
{
    public ChartsEditSample()
    {
    }

    public static void main(String args[])
    {
        WorkBook workBook = new WorkBook();
        try
        {
            //read in
            //workBook.read("..\\template\\chartTemplate.xls");
            workBook.readXLSX("..\\template\\chartTemplate.xlsx");

            //get chartshape from sheet 1
            ChartShape chartShape = workBook.getChart(0);

            chartShape.setChartType(ChartShape.Bar);
            chartShape.setTitle("Chart 1");
            //change 3D chart to 2D
            chartShape.set3Dimensional(false);

            //select sheet 2
            workBook.setSheet(1);
            //get chartshape in the sheet
            chartShape = workBook.getChart(0);
            //change chart type to step
            chartShape.setChartType(ChartShape.Step);
            //set axis title
            chartShape.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
            chartShape.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
            chartShape.setTitle("Chart 2");
            //change chart to 3D
            chartShape.set3Dimensional(true);

            //write out
            //workBook.write("./ChartEdit.xls");
            workBook.writeXLSX("./ChartEdit.xlsx");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}