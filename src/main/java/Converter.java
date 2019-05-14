
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.Iterator;
import java.util.logging.Logger;

import static org.apache.poi.xwpf.converter.pdf.PdfOptions.*;

public class Converter {

    private static Logger logger = Logger.getLogger(Converter.class.getName());

    public static void ConvertDocxToPDF(String documentPath, String pdfPath){
        try{
            InputStream inputStream = new FileInputStream(new File(documentPath));
            XWPFDocument document = new XWPFDocument(inputStream);
            PdfOptions pdfOptions = create();
            OutputStream outputStream = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(document, outputStream, pdfOptions);
            logger.info(documentPath + " to " + pdfPath + " Conversion: successful");
        }catch (FileNotFoundException e){
            logger.warning(e.getMessage());
        }catch (IOException e){
            logger.warning(e.getMessage());
        }
    }

    public static void ConvertXlsxToPDF(String documentPath, String pdfPath){
        try {
            InputStream inputStream = new FileInputStream(new File(documentPath));
            XSSFWorkbook document = new XSSFWorkbook(inputStream);
            XSSFSheet workSheet = document.getSheetAt(0);
            Iterator<Row> rowIterator = workSheet.iterator();
            com.lowagie.text.Document pdfDocument = new com.lowagie.text.Document();
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(pdfPath));
            pdfDocument.open();
            //PdfTable table = new PdfTable(2);

        }catch (FileNotFoundException e){
            logger.warning(e.getMessage());
        }catch (IOException e){
            logger.warning(e.getMessage());
        } catch (DocumentException e) {
            logger.warning(e.getMessage());
        }
    }

    public static void ConvertPptxToPDF(String documentPath, String pdfPath){


    }
}
