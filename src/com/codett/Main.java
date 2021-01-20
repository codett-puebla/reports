package com.codett;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JsonDataSource;
import net.sf.jasperreports.engine.export.JRXlsExporter;
import net.sf.jasperreports.engine.util.JRLoader;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import net.sf.jasperreports.export.SimpleXlsExporterConfiguration;
import net.sf.jasperreports.export.SimpleXlsReportConfiguration;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class Main {

    public static String pathReport;
    public static String pathSubReport;
    public static String data;
    public static String output;

    //order parameters pathReport pathSubReport output data

    public static void main(String[] args) {
        if (args.length != 0 && setData(args)) {
            try {
                try {
                    JasperReport report = (JasperReport) JRLoader.loadObject(new File(pathReport));
                    Map parameters = new HashMap();
                    parameters.put("subreportsPath", pathSubReport);
                    JRXlsExporter xlsExporter = new JRXlsExporter();
                    xlsExporter.setExporterInput(new SimpleExporterInput(JasperFillManager.fillReport(report, parameters, new JsonDataSource(new ByteArrayInputStream(data.getBytes())))));
                    xlsExporter.setExporterOutput(new SimpleOutputStreamExporterOutput(output));
                    SimpleXlsReportConfiguration xlsReportConfiguration = new SimpleXlsReportConfiguration();
                    SimpleXlsExporterConfiguration xlsExporterConfiguration = new SimpleXlsExporterConfiguration();
                    xlsReportConfiguration.setOnePagePerSheet(true);
                    xlsReportConfiguration.setRemoveEmptySpaceBetweenRows(false);
                    xlsReportConfiguration.setDetectCellType(true);
                    xlsReportConfiguration.setWhitePageBackground(false);
                    xlsExporter.setConfiguration(xlsReportConfiguration);
                    xlsExporter.exportReport();
                } catch (JRException ex) {
                    ex.printStackTrace();
                    throw new RuntimeException(ex);
                } catch (Exception ex) {
                    ex.printStackTrace();
                    throw new RuntimeException(ex);
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                throw new RuntimeException(ex);
            }
        } else {
            System.out.println("Arguments not found");
        }
    }

    private static boolean setData(String[] args) {
        boolean valid = false;
        if (args.length == 4) {
            valid = true;
            pathReport = args[0];
            pathSubReport = args[1];
            output = args[2];
            data = args[3];
        }
        return valid;
    }
}
