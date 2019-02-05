import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.util.Units;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ImageDoc 
{
    public static void main(String[] args) throws IOException, InvalidFormatException 
    {
        XWPFDocument docx = new XWPFDocument();
        XWPFParagraph par = docx.createParagraph();
        XWPFRun run = par.createRun();
        run.setText("Hello, World. This is my first java generated docx-file. Have fun.");
        run.setFontSize(13);
        InputStream pic = new FileInputStream("D:\\a.jpg");
        //byte [] picbytes = IOUtils.toByteArray(pic);
        //run.addPicture(picbytes, Document.PICTURE_TYPE_JPEG);
        run.addPicture(pic, Document.PICTURE_TYPE_JPEG, "a", Units.toEMU(200), Units.toEMU(200));

        FileOutputStream out = new FileOutputStream("D:\\demo.docx"); 
        docx.write(out); 
        out.close(); 
        pic.close();
    }
}