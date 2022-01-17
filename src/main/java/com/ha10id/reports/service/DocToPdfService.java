/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ha10id.reports.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import java.io.ByteArrayInputStream;
import java.util.Arrays;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;

import org.springframework.stereotype.Service;

/**
 *
 * @author ha10id
 */
@Service
public class DocToPdfService {

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    public Boolean convert(String src, String dst) throws IOException {
        try {
            ByteArrayOutputStream pdfOutStream = new ByteArrayOutputStream();
            InputStream inputStream = new FileInputStream(src);
            HSSFWorkbook excelBook = new HSSFWorkbook();
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

            Document htmlDocument = excelToHtmlConverter.getDocument();
            try (ByteArrayOutputStream outStream = new ByteArrayOutputStream()) {
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
                //convert html input stream to pdf out stream
                ConverterProperties converterProperties = new ConverterProperties();
                HtmlConverter.convertToPdf(htmlInputStream, pdfOutStream, converterProperties);

                byte[] pdfBytes = pdfOutStream.toByteArray();
                System.out.println(Arrays.toString(pdfBytes));
                try ( //write to physical pdf file
                        OutputStream out = new FileOutputStream(dst)) {
                    out.write(pdfBytes);
                }
            }
        } catch (IOException | IllegalArgumentException | ParserConfigurationException | TransformerException e) {
            System.out.println(e.getMessage());
            return false;
            // log exception details and throw custom exception
        }
        return true;
    }
}
