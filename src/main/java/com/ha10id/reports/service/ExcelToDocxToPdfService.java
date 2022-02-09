package com.ha10id.reports.service;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import javax.xml.bind.JAXBException;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.AltChunkType;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

/**
 * @author ha10id
 */
@Service
public class ExcelToDocxToPdfService {

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static final String PDF_FILE = "pdf";
    private static final String DOC_FILE = "docx";

    private static void convertHtmlToDocx(InputStream inputStream, String outputFilePath) throws IOException, InvalidFormatException, Docx4JException, JAXBException {
        String inputData = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
        ndp.unmarshalDefaultNumbering();
        AlternativeFormatInputPart inputPart = new AlternativeFormatInputPart(AltChunkType.Xhtml);
        inputPart.setContentType(new ContentType("text/html"));
        inputPart.setBinaryData(inputData.getBytes());
        Relationship altChunkRel = wordMLPackage.getMainDocumentPart().addTargetPart(inputPart);
// .. the bit in document body
        CTAltChunk ac = Context.getWmlObjectFactory().createCTAltChunk();
        ac.setId(altChunkRel.getId());
        wordMLPackage.getMainDocumentPart().addObject(ac);
// .. content type
        wordMLPackage.getContentTypeManager().addDefaultContentType("html", "text/html");
// .. Saving the Document
        wordMLPackage.save(new java.io.File(outputFilePath));
    }

    public Boolean convertDocxToPdf(String src, String dst) throws FileNotFoundException, IOException {
        InputStream inputStream = new FileInputStream(src);
        OutputStream out;
        try ( XWPFDocument document = new XWPFDocument(inputStream)) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("This is new Text in this document.");
            PdfOptions options = PdfOptions.create();
            out = new FileOutputStream(String.format("%s.pdf", dst));
            PdfConverter.getInstance().convert(document, out, options);
        }
        out.close();
        return null;
    }

    private ByteArrayOutputStream convertXlsxToHtml(String src) throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException {
        InputStream inputStream = new FileInputStream(src);
        HSSFWorkbook excelBook = new HSSFWorkbook();
        Document htmlDocument = null;
        //Judge the Excel file to convert the 07+ version to the 03 version
        if (src.endsWith(EXCEL_XLS)) {  //Excel 2003
            excelBook = new HSSFWorkbook(inputStream);
        } else if (src.endsWith(EXCEL_XLSX)) {  // Excel 2007/2010
            Transform xls = new Transform();
            XSSFWorkbook workbookOld = new XSSFWorkbook(inputStream);
            xls.transformXSSF(workbookOld, excelBook);
        }

        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        //Remove the Excel header line
        excelToHtmlConverter.setOutputColumnHeaders(false);
        //Remove the Excel line number
        excelToHtmlConverter.setOutputRowNumbers(false);
        excelToHtmlConverter.processWorkbook(excelBook);

        htmlDocument = excelToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        //convert xls stream to html - w3 document
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();

        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);

        return outStream;
    }

    public Boolean convert(String src, String dst) throws IOException, Exception {
//        byte[] pdfBytes = null;
        ByteArrayOutputStream pdfOutStream = new ByteArrayOutputStream();

        InputStream htmlInputStream = new ByteArrayInputStream(convertXlsxToHtml(src).toByteArray());
        //convert html inputstream to pdf out stream
        //ConverterProperties converterProperties = new ConverterProperties();
        //HtmlConverter.convertToPdf(htmlInputStream, pdfOutStream, converterProperties);
        //convert html inputstream to doc out stream
//        convertHtmlToDocx(htmlInputStream, String.format("%s.docx", dst));
//        htmlInputStream.reset();
//        pdfBytes = pdfOutStream.toByteArray();
//        System.out.println(Arrays.toString(pdfBytes));
        PdfDocument pdfDoc = new PdfDocument(new PdfWriter(String.format("%s.pdf", dst)));
        pdfDoc.setDefaultPageSize(new PageSize(1700, 842));
//        write to physical pdf file
        HtmlConverter.convertToPdf(htmlInputStream, pdfDoc);
        return null;
    }
}
