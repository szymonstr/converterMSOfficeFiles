import java.rmi.server.UID;
import java.util.UUID;

public class TestConverter {

    private static Converter converter;

    private static String docx = "C:\\testFiles\\DocTest.docx";
    private static String pdfDocx = "C:\\testFiles\\DocTest.pdf";

    private static String xlsx = "C:\\testFiles\\ExcelTest.xlsx";
    private static String pdfXlsx = "C:\\testFiles\\ExcelTest.pdf";

    private static String pptx = "C:\\testFiles\\PptxTest.pptx";
    private static String pdfPptx = "C:\\testFiles\\PptxTest.pdf";

    public static void main(String[] args){

        System.out.println("Start conversion...");
        ///*
        System.out.println("Start: docx -> pdf");
        //converter.ConvertToPDF(docx, pdfDocx);
        System.out.println("Finish: docx -> pdf");
        //*/
        /*
        System.out.println("Start: xlsx -> pdf");
        converter.ConvertToPDF(xlsx, pdfXlsx);
        System.out.println("Finish: xlsx -> pdf");

         */
        /*
        System.out.println("Start: pptx -> pdf");
        converter.ConvertToPDF(pptx, pdfPptx);
        System.out.println("Finish: pptx -> pdf");
        */
        System.out.println("Conversion Ended");

    }

}
