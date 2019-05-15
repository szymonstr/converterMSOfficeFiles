

import com.lowagie.text.*;
import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
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
            FileInputStream fileInputStream = new FileInputStream(new File(documentPath));
            XSSFWorkbook document = new XSSFWorkbook(fileInputStream);
            XSSFSheet workSheet = document.getSheetAt(0);
            Iterator<Row> rowIterator = workSheet.iterator();
            com.lowagie.text.Document pdfDocument = new com.lowagie.text.Document();
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(pdfPath));
            pdfDocument.open();
            com.lowagie.text.pdf.PdfPTable pdfTable  = new com.lowagie.text.pdf.PdfPTable((int) workSheet.getRow(2).getLastCellNum());
            com.lowagie.text.pdf.PdfPCell tableCell;
            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    org.apache.poi.ss.usermodel.Cell cell = cellIterator.next();
                   //if (cell.getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING){
                        tableCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        pdfTable.addCell(tableCell);
                    //}
                }
            }
            pdfDocument.add(pdfTable);
            pdfDocument.close();
            fileInputStream.close();
            logger.info(documentPath + " to " + pdfPath + " Conversion: successful");

        }catch (FileNotFoundException e){
            logger.warning(e.getMessage());
        }catch (IOException e){
            logger.warning(e.getMessage());
        } catch (DocumentException e) {
            logger.warning(e.getMessage());
        }
    }

    public static void ConvertPptxToPDF(String documentPath, String pdfPath){
        try{
            FileInputStream fileInputStream = new FileInputStream(documentPath);
            XMLSlideShow pptx = new XMLSlideShow(OPCPackage.open(fileInputStream));
            fileInputStream.close();
            Dimension pageSize = pptx.getPageSize();
            float scale = 1;
            int width = (int) (pageSize.width * scale);
            int height = (int) (pageSize.height * scale);
            int iterator = 1;
            int numberOfSlides = pptx.getSlides().length;

            for (XSLFSlide slide : pptx.getSlides()){
                BufferedImage image = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_BGR);
                Graphics2D graphics = image.createGraphics();
                graphics.setPaint(Color.white);
                graphics.fill( new Rectangle2D.Float(0, 0, pageSize.width, pageSize.height));
                graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
                graphics.setColor(Color.white);
                graphics.clearRect(0, 0, width, height);
                graphics.scale(scale, scale);
                slide.draw(graphics);
                FileOutputStream fileOutputStream = new FileOutputStream("C:\\temp\\" + iterator + ".png");
                ImageIO.write(image, "png", fileOutputStream);
                fileOutputStream.close();
                iterator++;
            }

            Document pdfDocument = new Document();
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(pdfPath));
            PdfPTable pdfTable  = new PdfPTable(1);

            for (int i = 1 ; i <= numberOfSlides; i++){
                Image slideImage = Image.getInstance("C:\\temp\\" + i + ".png");
                pdfDocument.setPageSize(new Rectangle(slideImage.getWidth(), slideImage.getHeight()));
                pdfDocument.open();
                slideImage.setAbsolutePosition(0,0);
                pdfTable.addCell(new com.lowagie.text.pdf.PdfPCell(slideImage, true));
            }
            pdfDocument.add(pdfTable);
            pdfDocument.close();
            logger.info(documentPath + " to " + pdfPath + " Conversion: successful");
        }catch (FileNotFoundException e){
            logger.warning(e.getMessage());
        }catch (InvalidFormatException e){
            logger.warning(e.getMessage());
        }catch (IOException e){
            logger.warning(e.getMessage());
        }catch (DocumentException e){
            logger.warning(e.getMessage());
        }


    }
}
