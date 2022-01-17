/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ha10id.reports.service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import java.io.ByteArrayInputStream;
import java.util.Arrays;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.w3c.dom.Document;

import org.springframework.stereotype.Service;

/**
 * https://www.programmerall.com/article/2047471202/
 *
 * @author ha10id
 */
@Service
public class DocToPdfService {

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static final String WORD_DOC = "doc";
    private static final String WORD_DOCX = "docx";

    public Boolean convert(String src, String dst) throws FileNotFoundException, IOException, Docx4JException {
        byte[] pdfBytes = null;
        try {
            ByteArrayOutputStream pdfOutStream = new ByteArrayOutputStream();
            String fileName = "./test.doc";
            InputStream inputStream = new FileInputStream(fileName);
            HSSFWorkbook excelBook = new HSSFWorkbook();
            Document htmlDocument = null;
            //Judge the Excel file to convert the 07+ version to the 03 version
            if (fileName.endsWith(EXCEL_XLS)) {  //Excel 2003 
                excelBook = new HSSFWorkbook(inputStream);
            } else if (fileName.endsWith(EXCEL_XLSX)) {  // Excel 2007/2010 
                Transform xls = new Transform();
                XSSFWorkbook workbookOld = new XSSFWorkbook(inputStream);
                //excelBook = new HSSFWorkbook(inputStream);
                xls.transformXSSF(workbookOld, excelBook);
            } else if (fileName.endsWith(WORD_DOC)) {  //Word 2003 
                HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(inputStream);

                WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                        DocumentBuilderFactory.newInstance().newDocumentBuilder()
                                .newDocument());
                wordToHtmlConverter.processDocument(wordDocument);
                htmlDocument = wordToHtmlConverter.getDocument();
            } else if (fileName.endsWith(WORD_DOCX)) {
                WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                        .load(inputStream);
                
//                // 2) Prepare Pdf settings
//                PdfSettings pdfSettings = new PdfSettings();
//
//                // 3) Convert WordprocessingMLPackage to Pdf
//                OutputStream out = new FileOutputStream(new File(
//                        "pdf/HelloWorld.pdf"));
//                PdfConversion converter = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(
//                        wordMLPackage);
//                converter.output(out, pdfSettings);
            }

            if (fileName.endsWith(EXCEL_XLS) || fileName.endsWith(EXCEL_XLSX)) {
                ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
                //Remove the Excel header line 
                excelToHtmlConverter.setOutputColumnHeaders(false);
                //Remove the Excel line number 
                excelToHtmlConverter.setOutputRowNumbers(false);
                excelToHtmlConverter.processWorkbook(excelBook);

                htmlDocument = excelToHtmlConverter.getDocument();
            }
            try ( ByteArrayOutputStream outStream = new ByteArrayOutputStream()) {
                //convert xls stream to html - w3 document
                DOMSource domSource = new DOMSource(htmlDocument);
                StreamResult streamResult = new StreamResult(outStream);
                TransformerFactory tf = TransformerFactory.newInstance();
                Transformer serializer = tf.newTransformer();

                serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
                serializer.setOutputProperty(OutputKeys.INDENT, "yes");
                serializer.setOutputProperty(OutputKeys.METHOD, "html");
                serializer.transform(domSource, streamResult);

                InputStream htmlInputStream = new ByteArrayInputStream(outStream.toByteArray());
                //convert html inputstream to pdf out stream
                ConverterProperties converterProperties = new ConverterProperties();
                HtmlConverter.convertToPdf(htmlInputStream, pdfOutStream, converterProperties);

                pdfBytes = pdfOutStream.toByteArray();
                System.out.println(Arrays.toString(pdfBytes));
                try ( //write to physical pdf file
                         OutputStream out = new FileOutputStream("./toPDF.pdf")) {
                    out.write(pdfBytes);
                }
            }
        } catch (IOException | ParserConfigurationException | TransformerException e) {
            System.out.println(e);
            // log exception details and throw custom exception
        }
        return true;
    }
}
