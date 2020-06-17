
package com.globalhitss.scanPC;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.net.NetworkInterface;
import java.net.SocketException;
import java.util.Collections;
import java.util.Enumeration;
import java.util.Scanner;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

/**
 * ScanPC
 */

@Component
public class ScanPC {

    @Value("${app.ip.prefix}")
    String prefixIp;

    public  void scan() {

        System.out.println("============================== STARTING TO SCAN ==============================");
        
        String biosSerialNumber = getBiosSerialNumber();
        
        String[] getIPAndMac = getIPAndMac();
        String ipLocal      = getIPAndMac[0];
        String macAdress    = getIPAndMac[1];
        
        System.err.println("########## RESULT ##########");
        System.err.println(" - biosSerialNumber : " + biosSerialNumber);
        System.err.println(" - ipLocal : " + ipLocal);
        System.err.println(" - macAdress : " + macAdress);

        createExcelFile(biosSerialNumber, ipLocal, macAdress);

        System.out.println("============================== END OF SCAN ==============================");
    }

    private  String getBiosSerialNumber() {

        String biosSerialNumber = null;

        try {
            // wmic command for diskdrive id: wmic DISKDRIVE GET SerialNumber
            // wmic command for cpu id : wmic cpu get ProcessorId
            Process process = Runtime.getRuntime().exec(new String[] { "wmic", "bios", "get", "serialnumber" });
            process.getOutputStream().close();
            Scanner sc = new Scanner(process.getInputStream());
            String property = sc.next();
            String serial = sc.next();
            System.out.println(property + ": " + serial);

            biosSerialNumber = serial;

        } catch (IOException ex) {
            biosSerialNumber = "N/A";
            ex.printStackTrace();
        }

        return biosSerialNumber;
    }

    private   String[] getIPAndMac() {

        System.err.println("LOOKING FOR AN IP WITH THE PREFIX : " + prefixIp);

        String[] response = new String[2];
        String ipLocal = null;
        String macAdress = null;

        try {
            // get all network interfaces of the current system
            Enumeration<NetworkInterface> networkInterface = NetworkInterface.getNetworkInterfaces();
            // iterate over all interfaces
            boolean isLocalIPValid;

            while (networkInterface.hasMoreElements()) {

                isLocalIPValid = false;

                // get an interface
                NetworkInterface network = networkInterface.nextElement();
                // get its hardware or mac address
                byte[] macAddressBytes = network.getHardwareAddress();
                if (macAddressBytes != null) {

                    Enumeration<InetAddress> inetAddresses = network.getInetAddresses();
                    for (InetAddress inetAddress : Collections.list(inetAddresses)) {
                        if (inetAddress.getHostAddress().startsWith( prefixIp )) {
                            isLocalIPValid = true;
                            ipLocal = inetAddress.getHostAddress();
                        }
                    }

                    if (!isLocalIPValid) {
                        continue;
                    }

                    // initialize a string builder to hold mac address
                    StringBuilder macAddressStr = new StringBuilder();
                    // iterate over the bytes of mac address
                    for (int i = 0; i < macAddressBytes.length; i++) {
                        // convert byte to string in hexadecimal form
                        macAddressStr.append(String.format("%02X", macAddressBytes[i]));
                        // check if there are more bytes, then add a "-" to make it more readable
                        if (i < macAddressBytes.length - 1) {
                            macAddressStr.append("-");
                        }
                    }

                    macAdress = macAddressStr.toString();
                    break;

                }

            }

            System.err.println("ipLocal : " + ipLocal);
            System.err.println("macAdress : " + macAdress);

        } catch (SocketException e) {
            e.printStackTrace();
            ipLocal     = "N/A";
            macAdress   = "N/A";
        }

        response[0] = ipLocal;
        response[1] = macAdress;

        return response;
    }

    private  void createExcelFile(String biosSerialNumber, String ipLocal, String macAdress ) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("PC Scan Result");
         
        Object[][] bookData = {
            {biosSerialNumber, ipLocal, macAdress}
        };
 

         // HEADER
         // FONT HEADER
         CellStyle theaderStyle = workbook.createCellStyle();
        theaderStyle.setBorderBottom(BorderStyle.MEDIUM);
        theaderStyle.setBorderTop(BorderStyle.MEDIUM);
        theaderStyle.setBorderRight(BorderStyle.MEDIUM);
        theaderStyle.setBorderLeft(BorderStyle.MEDIUM);

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setFontName("Arial");
        font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        font.setBold(true);
        
        theaderStyle.setFont(font);
        theaderStyle.setAlignment(HorizontalAlignment.CENTER);
        theaderStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.index);
        theaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         Row row = sheet.createRow(0);
         row.createCell(0).setCellValue( "BIOS SERIAL NUMBER" );
         row.createCell(1).setCellValue( "IP LOCAL" );
         row.createCell(2).setCellValue( "MAC ADRESS" );
 
         row.getCell(0).setCellStyle(theaderStyle);
         row.getCell(1).setCellStyle(theaderStyle);
         row.getCell(2).setCellStyle(theaderStyle);

        int rowCount = 1;
         
        for (Object[] aBook : bookData) {

            row = sheet.createRow(rowCount++);
             
            int columnCount = 0;
             
            for (Object field : aBook) {
                Cell cell = row.createCell(columnCount++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
             
        }
         
        try {
            
            FileOutputStream outputStream = new FileOutputStream("PC_Scan_Result.xlsx");
            workbook.write(outputStream);

        } catch (Exception e) {
            System.err.println("ERROR WHILE GENERATIN EXCEL FILE");
            e.printStackTrace();
        }
         
    }

}