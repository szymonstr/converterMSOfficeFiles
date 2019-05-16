

import java.rmi.server.UID;
import java.util.UUID;

public class TestConverter {

    private static Converter converter;

    private static String docx = "C:\\testFiles\\DocTest1.docx";
    private static String pdfDocxPath = "C:\\testFiles\\";
    private static String pdfDocxName = "DocTest";
    private static String images = "C:\\testFiles\\images\\";
    private static int[] pageNumbers = {1,2,5};

    private static String xlsx = "C:\\testFiles\\ExcelTest.xlsx";
    private static String pdfXlsx = "C:\\testFiles\\ExcelTest.pdf";

    private static String pptx = "C:\\testFiles\\PptxTest1.pptx";
    private static String pdfPptxPath = "C:\\testFiles\\";
    private static String pdfPptxName = "PptxTest";

    private static String pdf = "C:\\testFiles\\pdfTest.pdf";
    private static String pdfPath = "C:\\testFiles\\";
    private static String pdfName = "PdfTest";

    public static void main(String[] args){

        OutputFileSettings outputPptxSettings = new OutputFileSettings(true, true, pageNumbers, pdfPptxPath, pdfPptxName);
        OutputFileSettings outputDocxSettings = new OutputFileSettings(true, true, pageNumbers, pdfDocxPath, pdfDocxName);
        OutputFileSettings outputPDFSettings = new OutputFileSettings(true, true, pageNumbers, pdfPath, pdfName);

        System.out.println("Start conversion...");
        ///*
        System.out.println("Start: docx -> pdf");
        converter.ConvertDocx(docx, outputDocxSettings);
        System.out.println("Finish: docx -> pdf");
        //*/
        /*
        System.out.println("Start: xlsx -> pdf");
        converter.ConvertXlsxToPDF(xlsx, pdfXlsx);
        System.out.println("Finish: xlsx -> pdf");
        // */
        ///*
        System.out.println("Start: pptx -> pdf");
        converter.ConvertPptx(pptx, outputPptxSettings);
        System.out.println("Finish: pptx -> pdf");
        //*/
        ///*
        System.out.println("Start: pdf -> png");
        converter.ConvertPDFToPNG(pdf, outputPDFSettings );
        System.out.println("Finish: pdf -> png");
        //*/

        System.out.println("Conversion Ended");

    }

}
