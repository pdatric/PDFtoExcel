import com.smartxls.ChartFormat;
import com.smartxls.ChartShape;
import com.smartxls.WorkBook;

import java.awt.*;

public class ChartSample
{
    public ChartSample()
    {
    }

    public static void main(String args[])
    {
        WorkBook m_workbook = new WorkBook();
        try
        {
            //set data
            m_workbook.setText(0,1,"Jan");
            m_workbook.setText(0,2,"Feb");
            m_workbook.setText(0,3,"Mar");
            m_workbook.setText(0,4,"Apr");
            m_workbook.setText(0,5,"Jun");

            m_workbook.setText(1,0,"Comfrey");
            m_workbook.setText(2,0,"Bananas");
            m_workbook.setText(3,0,"Papaya");
            m_workbook.setText(4,0,"Mango");
            m_workbook.setText(5,0,"Lilikoi");
            for(int col = 1; col <= 5; col++)
                for(int row = 1; row <= 5; row++)
                    m_workbook.setFormula(row, col, "RAND()");
            m_workbook.setText(6, 0, "Total");
            m_workbook.setFormula(6, 1, "SUM(B2:B6)");
            m_workbook.setSelection("B7:F7");
            //auto fill the range with the first cell's formula or data
            m_workbook.editCopyRight();

			int left = 1;
			int top = 7;
			int right =13;
			int bottom = 31;

            //create chart with it's location
            ChartShape chart = m_workbook.addChart(left,top,right,bottom);
			chart.setChartType(ChartShape.Column);
            //link data source, link each series to columns(true to rows).
            chart.setLinkRange("Sheet1!$a$1:$F$6",false);
            //set axis title
            chart.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
			chart.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
            //set series name
            chart.setSeriesName(0, "My Series number 1");
			chart.setSeriesName(1, "My Series number 2");
			chart.setSeriesName(2, "My Series number 3");
			chart.setSeriesName(3, "My Series number 4");
			chart.setSeriesName(4, "My Series number 5");
			chart.setTitle("My Chart");

            //set plot area's color to darkgray
            ChartFormat chartFormat = chart.getPlotFormat();
            chartFormat.setSolid();
            chartFormat.setForeColor(Color.DARK_GRAY.getRGB());
            chart.setPlotFormat(chartFormat);

            //set series 0's color to blue
            ChartFormat seriesformat = chart.getSeriesFormat(0);
            seriesformat.setSolid();
            seriesformat.setForeColor(Color.BLUE.getRGB());
            chart.setSeriesFormat(0, seriesformat);

            //set series 1's color to red
            seriesformat = chart.getSeriesFormat(1);
            seriesformat.setSolid();
            seriesformat.setForeColor(Color.RED.getRGB());
            chart.setSeriesFormat(1, seriesformat);

            //set chart title's font property
            ChartFormat titleformat = chart.getTitleFormat();
            titleformat.setFontSize(14*20);
            titleformat.setFontUnderline(true);
            chart.setTitleFormat(titleformat);
            
            //m_workbook.write("./Chart.xls");
            m_workbook.writeXLSX("./Chart.xlsx");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}