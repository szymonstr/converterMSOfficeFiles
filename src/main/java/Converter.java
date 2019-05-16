import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.lowagie.text.*;
import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.logging.Logger;


public class Converter {

    private static Logger logger = Logger.getLogger(Converter.class.getName());

    public static void ConvertDocx(String documentPath, OutputFileSettings outputFileSettings){
        String pdfPath = outputFileSettings.getPath() + outputFileSettings.getFileName() + ".pdf";

        DocxToPDF(documentPath, pdfPath);

        if (outputFileSettings.isPng()){
            ConvertPDFToPNG(pdfPath, outputFileSettings);
        }

        if (!outputFileSettings.isPdf()){
            DeleteFile(pdfPath);
        }
    }

    public static void ConvertPptx(String documentPath, OutputFileSettings outputFileSettings){

        ArrayList<String> tempFilesPaths = PptxToTempPNG(documentPath);

        if (outputFileSettings.isPdf()) {
            String pdfPath = outputFileSettings.getPath() + outputFileSettings.getFileName() + ".pdf";
            TempPNGToPDF(tempFilesPaths, pdfPath);
        }

        if (outputFileSettings.isPng()){
            PreparePNGFromTempPNG(outputFileSettings, tempFilesPaths);
        }

        DeleteRedundantFiles(tempFilesPaths);
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
            com.lowagie.text.pdf.PdfPTable pdfTable  = new com.lowagie.text.pdf.PdfPTable(4);
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

    public static void ConvertPDFToPNG(String pdfPath, OutputFileSettings outputFileSettings){
        String imagePath = outputFileSettings.getPath() + outputFileSettings.getFileName() + "_page";
        try{
            PDDocument document = PDDocument.load(new File(pdfPath));
            PDFRenderer pdfRenderer = new PDFRenderer(document);
            for (int page : outputFileSettings.getNumPages()){
                if (page < document.getNumberOfPages()) {
                    BufferedImage image = pdfRenderer.renderImageWithDPI(page-1, 340, ImageType.RGB);
                    String imageName = imagePath + page + ".png";
                    ImageIOUtil.writeImage(image, imageName, 340);
                }
            }
            document.close();
            logger.info(pdfPath + " to PNG files. Conversion: successful");
        }catch (IOException e){
            logger.warning(e.getMessage());
        }

    }

    private static void DocxToPDF(String documentPath, String pdfPath) {
        try {
            InputStream docxInputStream = new FileInputStream(documentPath);
            OutputStream outputStream = new FileOutputStream(pdfPath);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();

        }
    }

    private static void PreparePNGFromTempPNG(OutputFileSettings outputFileSettings, ArrayList<String> tempFileNames) {
        for (int slide : outputFileSettings.getNumPages()){
            int index = slide - 1;
            String src = tempFileNames.get(index);
            String dst = outputFileSettings.getPath() + outputFileSettings.getFileName() + "_slide" + slide + ".png";
            Path srcPath = Paths.get(src);
            Path dstPath = Paths.get(dst);
            try {
                Files.copy(srcPath, dstPath, StandardCopyOption.REPLACE_EXISTING);
            }catch (IOException e){
                logger.warning(e.getMessage());
            }
        }
    }

    private static void TempPNGToPDF(ArrayList<String> tempFileNames, String pdfPath) {
        try {
            Document pdfDocument = new Document();
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(pdfPath));
            PdfPTable pdfTable = new PdfPTable(1);

            for (String tempFile : tempFileNames) {
                Image slideImage = Image.getInstance(tempFile);
                pdfDocument.setPageSize(new Rectangle(slideImage.getWidth(), slideImage.getHeight()));
                pdfDocument.setMargins(0, 0, 0, 0);
                pdfDocument.open();
                slideImage.setAbsolutePosition(0, 0);
                pdfTable.addCell(new PdfPCell(slideImage, true));
            }
            pdfDocument.add(pdfTable);
            pdfDocument.close();
        }catch (IOException e) {
            logger.warning(e.getMessage());
        }catch (DocumentException e){
            logger.warning(e.getMessage());
        }
    }

    private static ArrayList<String> PptxToTempPNG(String documentPath) {
        ArrayList<String> fileNames = new ArrayList<>();
        try{
            FileInputStream fileInputStream = new FileInputStream(documentPath);
            XMLSlideShow pptx = new XMLSlideShow(OPCPackage.open(fileInputStream));
            fileInputStream.close();
            Dimension pageSize = pptx.getPageSize();
            float scale = 1;
            int width = (int) (pageSize.width * scale);
            int height = (int) (pageSize.height * scale);
            int iterator = 1;
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
                String fileName = "C:\\temp\\" + iterator + ".png";
                FileOutputStream fileOutputStream = new FileOutputStream(fileName);
                ImageIO.write(image, "png", fileOutputStream);
                fileOutputStream.close();
                fileNames.add(fileName);
                iterator++;
            }
        }catch (FileNotFoundException e){
            logger.warning(e.getMessage());
        }catch (InvalidFormatException e){
            logger.warning(e.getMessage());
        }catch (IOException e) {
            logger.warning(e.getMessage());
        }
        return fileNames;
    }

    private static void DeleteRedundantFiles(ArrayList<String> tempFileNames) {
        for (String tempFileName : tempFileNames){
            DeleteFile(tempFileName);
        }
    }

    private static void DeleteFile(String path){
        File file = new File(path);
        if (file.exists()){
            file.delete();
        }
    }
}
