package com.ha10id.reports;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

@SpringBootApplication
public class ReportApplication {
//    public static PDFGenerator pDFGenerator;
    public static void main(String[] args) {
        ApplicationContext ac = SpringApplication.run(ReportApplication.class, args);
//        pDFGenerator = ac.getBean("pdfGenerator",PDFGenerator.class);
    }
}
