import com.smartxls.RangeStyle;
import com.smartxls.WorkBook;

import java.awt.*;

public class RangeStyleSample
{

    public static void main(String args[])
    {
        try
        {
            RangeStyleSample rsSample = new RangeStyleSample();
            rsSample.run();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    int StartRow, StartCol, EndRow, EndCol;
    String StartRange, eRange, hdrRange, colRange, ftrRange, bodyRange;
    WorkBook m_WorkBook;
    RangeStyle m_RangeStyle;

    public RangeStyleSample()
    {
    }

    void run()
            throws Exception
    {
        m_WorkBook = new WorkBook();
        m_WorkBook.setNumSheets(14);
        m_WorkBook.setSheetName(0, "Simple Format");
        m_WorkBook.setSheetName(1, "Classic1 Format");
        m_WorkBook.setSheetName(2, "Classic2 Format");
        m_WorkBook.setSheetName(3, "Classic3 Format");
        m_WorkBook.setSheetName(4, "Accounting1 Format");
        m_WorkBook.setSheetName(5, "Accounting2 Format");
        m_WorkBook.setSheetName(6, "Accounting3 Format");
        m_WorkBook.setSheetName(7, "Effects3D1 Format");
        m_WorkBook.setSheetName(8, "Colorful1 Format");
        m_WorkBook.setSheetName(9, "Colorful2 Format");
        m_WorkBook.setSheetName(10, "Colorful3 Format");
        m_WorkBook.setSheetName(11, "List1 Format");
        m_WorkBook.setSheetName(12, "List2 Format");
        m_WorkBook.setSheetName(13, "List3 Format");

        m_WorkBook.setSheet(0);
        setData();
        simpleFormat();

        m_WorkBook.setSheet(1);
        setData();
        Classic1();

        m_WorkBook.setSheet(2);
        setData();
        Classic2();

        m_WorkBook.setSheet(3);
        setData();
        Classic3();

        m_WorkBook.setSheet(4);
        setData();
        Accounting1();

        m_WorkBook.setSheet(5);
        setData();
        Accounting2();

        m_WorkBook.setSheet(6);
        setData();
        Accounting3();

        m_WorkBook.setSheet(7);
        setData();
        Effects3D1();

        m_WorkBook.setSheet(8);
        setData();
        Colorful1();

        m_WorkBook.setSheet(9);
        setData();
        Colorful2();

        m_WorkBook.setSheet(10);
        setData();
        Colorful3();

        m_WorkBook.setSheet(11);
        setData();
        List1();

        m_WorkBook.setSheet(12);
        setData();
        List2();

        m_WorkBook.setSheet(13);
        setData();
        List3();

        //m_WorkBook.write(".\\RangeStyle.xls");
        m_WorkBook.writeXLSX(".\\RangeStyle.xlsx");
    }

    private void setData()
            throws Exception
    {
        m_WorkBook.setText(1, 2, "Jan");
        m_WorkBook.setText(1, 3, "Feb");
        m_WorkBook.setText(1, 4, "Mar");
        m_WorkBook.setText(1, 5, "Apr");
        m_WorkBook.setText(2, 1, "Bananas");
        m_WorkBook.setText(3, 1, "Papaya");
        m_WorkBook.setText(4, 1, "Mango");
        m_WorkBook.setText(5, 1, "Lilikoi");
        m_WorkBook.setText(6, 1, "Comfrey");
        m_WorkBook.setText(7, 1, "Total");
        m_WorkBook.setFormula(2, 2, "RAND()*100");
        m_WorkBook.setSelection(2, 2, 2, 5);
        m_WorkBook.editCopyRight();
        m_WorkBook.setSelection(2, 2, 6, 5);
        m_WorkBook.editCopyDown();
        m_WorkBook.setFormula(7, 2, "SUM(C3:C7)");
        m_WorkBook.setSelection("C8:F8");
        m_WorkBook.editCopyRight();

        StartRow = 1;
        StartCol = 1;
        EndRow = 7;
        EndCol = 5;
        StartRange = m_WorkBook.formatRCNr(StartRow, StartCol, false);
        eRange = m_WorkBook.formatRCNr(StartRow, EndCol, false);
        hdrRange = StartRange + ":" + eRange;

        eRange = m_WorkBook.formatRCNr(EndRow, StartCol, false);
        colRange = StartRange + ":" + eRange;

        StartRange = m_WorkBook.formatRCNr(EndRow, StartCol, false);
        eRange = m_WorkBook.formatRCNr(EndRow, EndCol, false);
        ftrRange = StartRange + ":" + eRange;

        StartRange = m_WorkBook.formatRCNr(StartRow + 1, StartCol + 1, false);
        eRange = m_WorkBook.formatRCNr(EndRow - 1, EndCol, false);
        bodyRange = StartRange + ":" + eRange;

        m_RangeStyle = m_WorkBook.getRangeStyle();
    }

    private void simpleFormat()
            throws Exception
    {
        m_WorkBook.setSelection(colRange);
        AdjustFont(Color.BLACK.getRGB(), true, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_WorkBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);

        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentRight);
        m_RangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentBottom);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_WorkBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);
    }

    private void Classic1()
            throws Exception
    {
        m_WorkBook.setSelection(colRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderNone);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderNone);
        m_RangeStyle.setRightBorder(RangeStyle.BorderThin);
        AdjustFont(Color.BLACK.getRGB(), true, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setVerticalInsideBorder(RangeStyle.BorderNone);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        AdjustFont(Color.black.getRGB(), false, true, false);
        m_RangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentRight);
        m_RangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentBottom);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Classic2()
            throws Exception
    {
        m_WorkBook.setSelection(colRange);
        short nPattern;
        AdjustFont(Color.BLACK.getRGB(), true, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        AdjustFont(Color.BLACK.getRGB(), false, false, false);
        nPattern = 1;
        m_RangeStyle.setPattern(nPattern);
        m_RangeStyle.setPatternFG(Color.MAGENTA.getRGB());
        m_RangeStyle.setPatternBG(Color.MAGENTA.getRGB());
        m_RangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentRight);
        m_RangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentBottom);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Classic3()
            throws Exception
    {
        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        AdjustFont(Color.WHITE.getRGB(), true, true, false);
        AlignRight();
        SetSolidPattern(m_WorkBook.getPaletteEntry(11), Color.BLACK.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.BLACK.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.BLACK.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(colRange);
        SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.BLACK.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Accounting1()
            throws Exception
    {
        m_WorkBook.setSelection(hdrRange);
        String numberFormat;
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        AdjustFont(Color.MAGENTA.getRGB(), true, true, false);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderNone);
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderDouble);
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Accounting2()
            throws Exception 
    {
        String numberFormat;

        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThick);
        m_RangeStyle.setTopBorderColor(Color.LIGHT_GRAY.getRGB());
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        m_RangeStyle.setBottomBorderColor(Color.LIGHT_GRAY.getRGB());
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderNone);
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThick);
        m_RangeStyle.setBottomBorderColor(Color.LIGHT_GRAY.getRGB());
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setTopBorderColor(Color.LIGHT_GRAY.getRGB());
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Accounting3()
            throws Exception
    {
        String numberFormat;
        m_WorkBook.setSelection(colRange);
        AdjustFont(Color.BLACK.getRGB(), false, true, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        numberFormat = "#,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setTopBorder(RangeStyle.BorderNone);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderNone);
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderDouble);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(bodyRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderNone);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorderColor(Color.GREEN.getRGB());
        AdjustFont(m_WorkBook.getPaletteEntry(16), false, true, false);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Effects3D1()
            throws Exception
    {
        m_WorkBook.setSelection(1, 1, 7, 5);
        SetSolidPattern(Color.LIGHT_GRAY.getRGB(), 0);

        Set3DBorder(2, 2,
                      6, 5,
                      Color.LIGHT_GRAY.getRGB(), Color.DARK_GRAY.getRGB(), Color.LIGHT_GRAY.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        AdjustFont(Color.MAGENTA.getRGB(), true, false, false);
        AlignCenter();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(colRange);
        Set3DBorder(1, 1, 7, 1, Color.LIGHT_GRAY.getRGB(), Color.LIGHT_GRAY.getRGB(), Color.DARK_GRAY.getRGB());
        AdjustFont(Color.BLACK.getRGB(), true, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        Set3DBorder(7, 1, 7, 5, Color.LIGHT_GRAY.getRGB(), Color.DARK_GRAY.getRGB(), Color.DARK_GRAY.getRGB());
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Colorful1()
            throws Exception
    {
        m_WorkBook.setSelection(1, 1, 7, 5);
        int color = Color.RED.getRGB();
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        m_RangeStyle.setBottomBorderColor(Color.RED.getRGB());
        SetSolidPattern(Color.DARK_GRAY.getRGB(), Color.BLACK.getRGB());

        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setLeftBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setRightBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorderColor(Color.RED.getRGB());
        m_RangeStyle.setLeftBorderColor(Color.RED.getRGB());
        m_RangeStyle.setRightBorderColor(Color.RED.getRGB());

        AdjustFont(color, false, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        SetSolidPattern(Color.BLACK.getRGB(), Color.BLACK.getRGB());
        AdjustFont(color, true, true, false);
        AlignCenter();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(colRange);
        SetSolidPattern(m_WorkBook.getPaletteEntry(11), Color.BLACK.getRGB());
        AdjustFont(color, true, true, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Colorful2() throws Exception
    {
        int color = m_WorkBook.getPaletteEntry(14);
        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        SetSolidPattern(m_WorkBook.getPaletteEntry(9), Color.BLACK.getRGB());
        AdjustFont(color, true, true, false);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(colRange);
        AdjustFont(Color.BLACK.getRGB(), true, true, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(1, 1, 7, 5);
        SetHatchPattern(m_WorkBook.getPaletteEntry(16), Color.RED.getRGB());
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Colorful3()
            throws Exception
    {
        m_WorkBook.setSelection(1, 1, 7, 5);
        SetSolidPattern(Color.BLACK.getRGB(), Color.BLACK.getRGB());
        AdjustFont(Color.WHITE.getRGB(), false, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        AdjustFont(Color.GREEN.getRGB(), true, true, false);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(colRange);
        AdjustFont(Color.MAGENTA.getRGB(), true, true, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void List1()
            throws Exception
    {
        m_WorkBook.setSelection(1, 1, 7, 5);
        String newSelection, numberFormat;
        int fcolor, bcolor;
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setLeftBorder(RangeStyle.BorderThin);
        m_RangeStyle.setRightBorder(RangeStyle.BorderThin);
        m_RangeStyle.setTopBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setLeftBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setRightBorderColor(m_WorkBook.getPaletteEntry(16));
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        SetSolidPattern(m_WorkBook.getPaletteEntry(14), Color.BLACK.getRGB());
        AdjustFont(Color.BLUE.getRGB(), true, true, false);
        AlignCenter();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        SetSolidPattern(m_WorkBook.getPaletteEntry(14), Color.BLACK.getRGB());
        AdjustFont(Color.BLUE.getRGB(), true, false, false);
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        fcolor = m_WorkBook.getPaletteEntry(19);
        bcolor = Color.RED.getRGB();
        for (int i = StartRow + 1; i <= EndRow - 1; i = i + 2)
        {
            newSelection = m_WorkBook.formatRCNr(i, StartCol, false) +
                            ":" + m_WorkBook.formatRCNr(i, EndCol, false);
            m_WorkBook.setSelection(newSelection);
            SetHatchPattern(fcolor, bcolor);
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        fcolor = Color.WHITE.getRGB();
        bcolor = m_WorkBook.getPaletteEntry(15);
        for (int i = StartRow + 2; i < EndRow - 1; i = i + 2)
        {
            newSelection = m_WorkBook.formatRCNr(i, StartCol, false) +
                            ":" + m_WorkBook.formatRCNr(i, EndCol, false);
            m_WorkBook.setSelection(newSelection);
            SetHatchPattern(fcolor, bcolor);
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }
    }

    private void List2()
            throws Exception
    {
        m_WorkBook.setSelection(1, 1, 7, 5);
        String numberFormat, newSelection;
        int fcolor, bcolor;
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        m_RangeStyle.setLeftBorder(RangeStyle.BorderThin);
        m_RangeStyle.setRightBorder(RangeStyle.BorderThin);
        m_RangeStyle.setTopBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setLeftBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setRightBorderColor(m_WorkBook.getPaletteEntry(16));
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderThick);
        m_RangeStyle.setTopBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThin);
        SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.BLACK.getRGB());
        AlignCenter();
        AdjustFont(Color.RED.getRGB(), true, true, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderThick);
        m_RangeStyle.setBottomBorderColor(m_WorkBook.getPaletteEntry(16));
        m_RangeStyle.setTopBorder(RangeStyle.BorderThin);
        numberFormat = "$ #,##0.00_);(#,##0.00)";
        m_RangeStyle.setCustomFormat(numberFormat);
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);

        fcolor = Color.ORANGE.getRGB();
        bcolor = Color.WHITE.getRGB();
        for (int i = StartRow + 1; i < EndRow; i = i + 2)
        {
            newSelection = m_WorkBook.formatRCNr(i, StartCol, false) +
                            ":" + m_WorkBook.formatRCNr(i, EndCol, false);

            m_WorkBook.setSelection(newSelection);
            SetHatchPattern(fcolor, bcolor);
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }
    }

    private void List3()
            throws Exception
    {
        m_WorkBook.setSelection(colRange);
        m_WorkBook.setRangeStyle(m_RangeStyle);
        m_WorkBook.setSelection(hdrRange);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorderColor(Color.DARK_GRAY.getRGB());
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorderColor(Color.DARK_GRAY.getRGB());
        AlignCenter();
        AdjustFont(m_WorkBook.getPaletteEntry(11), true, false, false);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(ftrRange);
        m_RangeStyle = m_WorkBook.getRangeStyle(EndRow, StartCol, EndRow, EndCol);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorderColor(Color.DARK_GRAY.getRGB());
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorderColor(Color.DARK_GRAY.getRGB());
        AlignRight();
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void Set3DBorder(
            int row1,
            int col1,
            int row2,
            int col2,
            int outlineColor,
            int rightColor,
            int bottomColor)
            throws Exception
    {
        m_WorkBook.setSelection(row1, col1, row2, col2);
        m_RangeStyle.setTopBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setLeftBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setRightBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setTopBorderColor(outlineColor);
        m_RangeStyle.setBottomBorderColor(outlineColor);
        m_RangeStyle.setLeftBorderColor(outlineColor);
        m_RangeStyle.setRightBorderColor(outlineColor);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(row1, col2, row2, col2);
        m_RangeStyle.setRightBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setRightBorderColor(rightColor);
        m_WorkBook.setRangeStyle(m_RangeStyle);

        m_WorkBook.setSelection(row2, col1, row2, col2);
        m_RangeStyle.setBottomBorder(RangeStyle.BorderMedium);
        m_RangeStyle.setBottomBorderColor(bottomColor);
        m_WorkBook.setRangeStyle(m_RangeStyle);
    }

    private void SetHatchPattern(
            int fcolor,
            int bcolor)
    {
        m_RangeStyle.setPattern((short)4);
        m_RangeStyle.setPatternFG(fcolor);
        m_RangeStyle.setPatternBG(bcolor);
    }

    private void AlignCenter()
    {
        m_RangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentCenter);
        m_RangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentBottom);
        m_RangeStyle.setWordWrap(false);
    }

    private void AdjustFont(int color, boolean bold, boolean italic, boolean underline)
    {
        m_RangeStyle.setFontBold(bold);
        m_RangeStyle.setFontItalic(italic);
        m_RangeStyle.setFontUnderline(RangeStyle.UnderlineSingle);
        m_RangeStyle.setFontColor(color);
    }

    private void AlignRight()
    {
        m_RangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentRight);
        m_RangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentBottom);
        m_RangeStyle.setWordWrap(false);
    }

    private void SetSolidPattern(int fcolor, int bcolor)
    {
        short nPattern;
        nPattern = 1;
        m_RangeStyle.setPattern(nPattern);
        m_RangeStyle.setPatternFG(fcolor);
        m_RangeStyle.setPatternBG(bcolor);
    }

}
