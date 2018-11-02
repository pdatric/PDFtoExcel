import com.smartxls.WorkBook;

import java.io.FileOutputStream;

public class ReadImageSample
{

    public static void main(String args[])
    {
        try
        {
            WorkBook m_book = new WorkBook();

            //open the workbook
            //m_book.read("..\\template\\book.xls");
            m_book.readXLSX("..\\template\\book.xlsx");

            String filename = "img";
            com.smartxls.PictureShape pic = m_book.getPictureShape(0);
            int type = pic.getPictureType();
            if(type == -1)
                filename += ".gif";
            else if(type == 5)
                filename += ".jpg";
            else if(type == 6)
                filename += ".png";
            else if(type == 7)
                filename += ".bmp";

            byte[] imagedata = pic.getPictureData();
            
            FileOutputStream fos = new FileOutputStream(filename);
            fos.write(imagedata);
            fos.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}