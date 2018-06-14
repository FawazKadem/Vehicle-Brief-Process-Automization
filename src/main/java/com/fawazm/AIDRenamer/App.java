package com.fawazm.AIDRenamer;
/**
 * Autogenerate Author Signature:
 * Author: Fawaz Mohammad
 * fawazm.mohammad@gmail.com
 * fmoham26@uwo.ca
 * <p>
 * Created for Autodata Solutions London
 */

/**
 * Created for Autodata Solutions
 * Needed as argument: Directory to folder containing original SAL file
 */

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.IOFileFilter;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.poi.ss.formula.functions.Index;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.*;
import java.util.concurrent.TimeUnit;

// Program to rename assetIDs in sal files to equal their


/**
 * Hello world!
 *
 *
 *
 */
public class App {
    static NodeList partNodes, partRefNodes, paintNodes, paintRefNodes, packageNodes, packageRefNodes;
    static Integer amountOfParts, amountOfPartRefs, amountOfPaints, amountOfPaintRefs, amountOfPackages, amountOfPackageRefs;


    public static void main(String[] args) throws InterruptedException {


        try {

            DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
            documentBuilderFactory.setValidating(false);
            DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();

            //String dirPath = "C:\\\\Users\\\\fawaz.mohammad\\\\AID Renamer Test Environment\\\\Cascada\\\\Ext\\\\";


            String dirPath = args[0].replace("//","////");
            System.out.println(dirPath);
            File dir = new File(dirPath);

            Collection<File> dotSalFiles = FileUtils.listFiles(dir,new WildcardFileFilter("*.sal"), null);

            File salFile = dotSalFiles.iterator().next();

            System.out.println(salFile.getPath());


            String outputFileDir = dirPath + "\\NewSalAndAB\\";
            String outputFile = outputFileDir + "New SAL.xml";


            Document inProgSal = documentBuilder.parse(new FileInputStream(salFile));
            System.out.println("!!!!!!!!!!PARSING!!!!!!!!!!!!");


            partNodes = inProgSal.getElementsByTagName("Part");
            amountOfParts = partNodes.getLength();

            partRefNodes = inProgSal.getElementsByTagName("PartRef");
            amountOfPartRefs = partRefNodes.getLength();

            paintNodes = inProgSal.getElementsByTagName("Paint");
            amountOfPaints = paintNodes.getLength();

            paintRefNodes = inProgSal.getElementsByTagName("PaintRef");
            amountOfPaintRefs = paintRefNodes.getLength();

            packageNodes = inProgSal.getElementsByTagName("Package");
            amountOfPackages = packageNodes.getLength();

            packageRefNodes = inProgSal.getElementsByTagName("PackageRef");
            amountOfPackageRefs = packageRefNodes.getLength();


            Node partNode;
            Node refNode;

            String partNodeAssetID;
            String partNodeCode;

            String refNodeAssetRef;

            System.out.println("REPLACING PARTS");
            for (int i = 0; i < amountOfPartRefs; i++) {
                refNode = partRefNodes.item(i);
                refNodeAssetRef = refNode.getAttributes().item(0).getNodeValue();

                for (int j = 0; j < amountOfParts; j++) {
                    partNode = partNodes.item(j);

                    partNodeAssetID = partNode.getAttributes().item(0).getNodeValue();
                    partNodeCode = partNode.getAttributes().item(2).getNodeValue();


                    if (Objects.equals(refNodeAssetRef, partNodeAssetID)) {
                        System.out.print("beepboopBOOPBEEPboopBOOPBEEPbeepBOOPbeepBEEFboopbeepboopBEEP");
                        refNode.getAttributes().item(0).setNodeValue(partNodeCode);
                        break;
                    }
                }
            }


            System.out.println("REPLACING PAINTS");
            for (int i = 0; i < amountOfPaintRefs; i++) {
                refNode = paintRefNodes.item(i);
                refNodeAssetRef = refNode.getAttributes().item(0).getNodeValue();

                for (int j = 0; j < amountOfPaints; j++) {
                    partNode = paintNodes.item(j);

                    partNodeAssetID = partNode.getAttributes().item(0).getNodeValue();
                    partNodeCode = partNode.getAttributes().getNamedItem("code").getNodeValue();


                    if (Objects.equals(refNodeAssetRef, partNodeAssetID)) {
                        System.out.println(".....!!");
                        refNode.getAttributes().item(0).setNodeValue(partNodeCode);
                        break;
                    }
                }
            }


            System.out.println("REPLACING PACKAGES");
            for (int i = 0; i < amountOfPackageRefs; i++) {
                refNode = packageRefNodes.item(i);
                refNodeAssetRef = refNode.getAttributes().item(0).getNodeValue();

                for (int j = 0; j < amountOfPackages; j++) {
                    partNode = packageNodes.item(j);

                    partNodeAssetID = partNode.getAttributes().item(0).getNodeValue();
                    partNodeCode = partNode.getAttributes().item(1).getNodeValue();



                    if (Objects.equals(refNodeAssetRef, partNodeAssetID)) {
                        System.out.print("beepboopBOOPBEEPboopBOOPBEEPbeepBOOPbeepBEEFboopbeepboopBEEP");
                        refNode.getAttributes().item(0).setNodeValue(partNodeCode);
                        break;
                    }
                }

            }


            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();

            DOMSource source = new DOMSource(inProgSal);

            new File(outputFileDir).mkdirs();
            StreamResult target = new StreamResult(outputFile);

            transformer.transform(source, target);


            String workingFilePath = outputFileDir + '\\';
            String[] workingFilePathList = workingFilePath.split("\\\\");
            System.out.println(workingFilePathList[workingFilePathList.length-1]);

            String exportFileName = (

                    workingFilePathList[workingFilePathList.length-5]
                    + workingFilePathList[workingFilePathList.length-4]
                    + workingFilePathList[workingFilePathList.length-3]
                    + workingFilePathList[workingFilePathList.length-2]
                    + "_NewAssetBrief.xlsx"

                    );



            writeNewSalToExcel(inProgSal, outputFileDir + exportFileName);

            System.out.println(workingFilePathList[0]);
            System.out.println(workingFilePath);
            System.out.println(outputFileDir.toString());

        } catch (Exception e) {
            e.printStackTrace();
        }


        TimeUnit.SECONDS.sleep(3);
        System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! FINISHED.......! ");
        System.out.println("Created by Fawaz Mohammad for Autodata Solutions");
        System.out.println("Please contact if anything is broken. This program has NOT been extensively tested.");
        System.out.println("fawaz.mohammad@autodata.net");
        System.out.println("fawazm.mohammad@gmail.com");
    }

    private static void writeNewSalToExcel(Document finishedSal, String filepath) throws InterruptedException {


        NodeList trimSetNodes = finishedSal.getElementsByTagName("TrimSet");
        Integer amountOfTrimSets = trimSetNodes.getLength();

        NodeList trimLevelList = finishedSal.getElementsByTagName("TrimLevel");
        Integer amountOfTrimLevels = trimLevelList.getLength();


        List threeHeaderColumns = generateColumns(trimLevelList, amountOfTrimLevels);
        ArrayList firstHeaderColumn = (ArrayList) threeHeaderColumns.get(0);
        ArrayList secondHeaderColumn = (ArrayList) threeHeaderColumns.get(1);
        ArrayList thirdHeaderColumn = (ArrayList) threeHeaderColumns.get(2);

        XSSFWorkbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();

        XSSFSheet sheet = workbook.createSheet("Hype New Asset Brief");


        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        Font normalFont = workbook.createFont();
        normalFont.setBold(false);
        normalFont.setFontHeightInPoints((short) 10);
        normalFont.setColor(IndexedColors.BLACK.getIndex());

        Font boldNormalFont = workbook.createFont();
        normalFont.setBold(true);
        normalFont.setFontHeightInPoints((short) 10);
        normalFont.setColor(IndexedColors.BLACK.getIndex());

        Font labelFont = workbook.createFont();
        labelFont.setBold(true);
        labelFont.setFontHeightInPoints((short) 12);
        labelFont.setColor((IndexedColors.WHITE.getIndex()));


        XSSFCellStyle darkGreyFillStyle = workbook.createCellStyle();
        darkGreyFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        darkGreyFillStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        darkGreyFillStyle.setFont(labelFont);
        darkGreyFillStyle.setWrapText(true);


        XSSFCellStyle lightGreyFillStyle = workbook.createCellStyle();
        lightGreyFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        lightGreyFillStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        lightGreyFillStyle.setWrapText(true);

        CellStyle normalCellStyle = workbook.createCellStyle();
        normalCellStyle.setFont(normalFont);
        normalCellStyle.setWrapText(true);

        XSSFCellStyle boldNormalCellStyle = workbook.createCellStyle();
        boldNormalCellStyle.setFont(boldNormalFont);
        boldNormalCellStyle.setWrapText(true);


        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setWrapText(true);


        Row topGreyRow = sheet.createRow(0);

        Row headerRow1 = sheet.createRow(1);
        Row headerRow2 = sheet.createRow(2);
        Row headerRow3 = sheet.createRow(3);

        Row fillerRow1 = sheet.createRow(4);
        int rowNum = 5;

        for (int i = 0; i < firstHeaderColumn.size(); i++) {
            Cell cell = topGreyRow.createCell(i);
            cell.setCellValue(" ");
            cell.setCellStyle(darkGreyFillStyle);

            cell = headerRow1.createCell(i);
            cell.setCellValue((String) firstHeaderColumn.get(i));

            cell = headerRow2.createCell(i);
            cell.setCellStyle(boldNormalCellStyle);
            cell.setCellValue((String) secondHeaderColumn.get(i));

            cell = headerRow3.createCell(i);
            cell.setCellValue((String) thirdHeaderColumn.get(i));

            cell = fillerRow1.createCell(i);
            cell.setCellValue(" ");
            cell.setCellStyle(darkGreyFillStyle);


        }


        for (int i = 0; i < amountOfParts; i++) {
            Row row = sheet.createRow(rowNum++);
            String partCode = partNodes.item(i).getAttributes().getNamedItem("code").getNodeValue();
            String partDesc = partNodes.item(i).getAttributes().getNamedItem("desc").getNodeValue();

            row.createCell(0).setCellValue(partCode);
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue(partDesc);
            row.getCell(1).setCellStyle(normalCellStyle);
        }

        Row paintLabel = sheet.createRow(rowNum++);
        sheet.addMergedRegion(new CellRangeAddress((rowNum - 1), (rowNum - 1), 0, firstHeaderColumn.size() - 1));
        paintLabel.createCell(0).setCellValue("PAINTS");
        paintLabel.getCell(0).setCellStyle(darkGreyFillStyle);


        for (int i = 0; i < amountOfPaints; i++) {
            Row row = sheet.createRow(rowNum++);
            String paintCode = paintNodes.item(i).getAttributes().getNamedItem("code").getNodeValue();
            String paintDesc = paintNodes.item(i).getAttributes().getNamedItem("name").getNodeValue();

            row.createCell(0).setCellValue(paintCode);
            row.getCell(0).setCellStyle(headerCellStyle);

            row.createCell(1).setCellValue(paintDesc);
            row.getCell(1).setCellStyle(normalCellStyle);
        }

        Row packageLabel = sheet.createRow(rowNum++);
        sheet.addMergedRegion(new CellRangeAddress((rowNum - 1), (rowNum - 1), 0, firstHeaderColumn.size() - 1));
        packageLabel.createCell(0).setCellValue("PACKAGES");
        packageLabel.getCell(0).setCellStyle(darkGreyFillStyle);

        for (int i = 0; i < amountOfPackages; i++) {
            Row row = sheet.createRow(rowNum++);

            String packageCode = packageNodes.item(i).getAttributes().getNamedItem("code").getNodeValue();
            String packageDesc = packageNodes.item(i).getAttributes().getNamedItem("desc").getNodeValue();

            row.createCell(0).setCellValue(packageCode);
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue(packageDesc);
            row.getCell(1).setCellStyle(normalCellStyle);
            row.setHeight((short) 750);
        }


        ArrayList<XSSFCellStyle> colorChooser = colorOptions(workbook);
        int colorChooserIndex = 0;
        XSSFCellStyle currentStyle = colorChooser.get(0);



        for (int j = 2; j < firstHeaderColumn.size();j++){

            System.out.println();
            if (j > 2){



                if (sheet.getRow(2).getCell(j).getStringCellValue().equals(sheet.getRow(2).getCell(j-1).getStringCellValue())){

                    System.out.print("fill fill fill ");

                    currentStyle = colorChooser.get(colorChooserIndex);

                } else {

                    colorChooserIndex = (colorChooserIndex + 1) % 6;
                    currentStyle = colorChooser.get(colorChooserIndex);

                }




            }

            currentStyle.setBorderBottom(BorderStyle.THIN);
            currentStyle.setBorderTop(BorderStyle.THIN);
            currentStyle.setBorderLeft(BorderStyle.THIN);
            currentStyle.setBorderRight(BorderStyle.THIN);
            currentStyle.setWrapText(true);

            TimeUnit.MILLISECONDS.sleep(200);
            for (int i = 0; i < 3;i++){

                Cell cell = sheet.getRow(i+1).getCell(j);
                cell.setCellStyle(currentStyle);


            }

            for (int i = 0; i < amountOfParts; i++){
                Cell cell = sheet.getRow(i+5).createCell(j);
                cell.setCellStyle(currentStyle);
                System.out.print("%");
            }


            for (int i = 0; i < amountOfPaints;i++){
                Cell cell = sheet.getRow(i+5 + amountOfParts + 1).createCell(j);
                cell.setCellStyle(currentStyle);
                System.out.print("^");
            }

            for (int i = 0; i < amountOfPackages;i++){
                Cell cell = sheet.getRow(i+5 + amountOfParts + amountOfPaints + 2).createCell(j);
                cell.setCellStyle(currentStyle);

                System.out.print("!");
                System.out.print("----------");

            }



        }


        TimeUnit.SECONDS.sleep(2);

        for (int i = 0; i < amountOfPartRefs; i++) {

            System.out.println("!!!CONF !!!REFERENCES!!!");

            Node currentNode = partRefNodes.item(i);
            Node currentParentNode = currentNode.getParentNode();
            Node currentGPNode = currentParentNode.getParentNode();

            String currentPart = currentNode.getAttributes().getNamedItem("assetRef").getNodeValue();
            String currentTrimSet = currentGPNode.getAttributes().getNamedItem("code").getNodeValue();
            String currentTrimLevel = currentParentNode.getAttributes().getNamedItem("code").getNodeValue();

            for (int j = 2; j < firstHeaderColumn.size(); j++) {

                String trimSetToMatch = sheet.getRow(2).getCell(j).getStringCellValue();
                String trimLevelToMatch = sheet.getRow(3).getCell(j).getStringCellValue();

                System.out.print("Processing....");

                if (currentTrimSet.equals(trimSetToMatch) && currentTrimLevel.equals(trimLevelToMatch)) {


                    for (int k = 0; k < amountOfParts; k++) {

                        System.out.print(",,,");

                        if (sheet.getRow(k + 5).getCell(0).getStringCellValue().equals(currentPart)) {

                            System.out.println("()");
                            sheet.getRow(k + 5).getCell(j).setCellValue
                                    (currentNode.getAttributes().getNamedItem("text").getNodeValue());
                        }
                    }
                }
            }
        }


        for (int i = 0; i < amountOfPaintRefs; i++) {

            Node currentNode = paintRefNodes.item(i);
            Node currentParentNode = currentNode.getParentNode();
            Node currentGPNode = currentParentNode.getParentNode();

            String currentPart = currentNode.getAttributes().getNamedItem("assetRef").getNodeValue();
            String currentTrimSet = currentGPNode.getAttributes().getNamedItem("code").getNodeValue();
            String currentTrimLevel = currentParentNode.getAttributes().getNamedItem("code").getNodeValue();

            System.out.println();
            for (int j = 2; j < firstHeaderColumn.size(); j++) {

                String trimSetToMatch = sheet.getRow(2).getCell(j).getStringCellValue();
                String trimLevelToMatch = sheet.getRow(3).getCell(j).getStringCellValue();
                System.out.print("+++++++++++++++++++++++++++++++");

                if (currentTrimSet.equals(trimSetToMatch) && currentTrimLevel.equals(trimLevelToMatch)) {


                    for (int k = 0; k < amountOfPaints; k++) {

                        if (sheet.getRow(k + 5 + amountOfParts + 1).getCell(0).getStringCellValue().equals(currentPart)) {

                            sheet.getRow(k + 5 + amountOfParts + 1).getCell(j).setCellValue
                                    (currentNode.getAttributes().getNamedItem("text").getNodeValue());
                        }
                    }
                }
            }
        }



        for (int i = 0; i < amountOfPackageRefs; i++) {

            Node currentNode = packageRefNodes.item(i);
            System.out.print("CHILD...");
            Node currentParentNode = currentNode.getParentNode();
            System.out.print("PARENT...");
            Node currentGPNode = currentParentNode.getParentNode();
            System.out.print("PARENT...TWO...");


            String currentPart = currentNode.getAttributes().getNamedItem("assetRef").getNodeValue();
            String currentTrimSet = currentGPNode.getAttributes().getNamedItem("code").getNodeValue();
            String currentTrimLevel = currentParentNode.getAttributes().getNamedItem("code").getNodeValue();

            for (int j = 2; j < firstHeaderColumn.size(); j++) {

                String trimSetToMatch = sheet.getRow(2).getCell(j).getStringCellValue();
                String trimLevelToMatch = sheet.getRow(3).getCell(j).getStringCellValue();
                System.out.println("!!PACK!!AGE!! ");

                if (currentTrimSet.equals(trimSetToMatch) && currentTrimLevel.equals(trimLevelToMatch)) {


                    for (int k = 0; k < amountOfPackages; k++) {

                        if (sheet.getRow(k + 5 + amountOfParts + amountOfPaints + 1 + 1).getCell(0).getStringCellValue().equals(currentPart)) {

                            sheet.getRow(k + 5 + amountOfParts + amountOfPaints + 1 + 1).getCell(j).setCellValue
                                    (currentNode.getAttributes().getNamedItem("text").getNodeValue());
                        }
                    }
                }
            }
        }





        TimeUnit.SECONDS.sleep(3);



        Row signature = sheet.createRow(rowNum++ + 2);
        Row signature2 = sheet.createRow(rowNum++ + 2);

        signature.createCell(0).setCellValue("Please contact Fawaz M. if anything is broken");
        signature2.createCell(0).setCellValue("fawaz.mohammad@autodata.net fawazm.mohammad@gmail.com");





        sheet.setDefaultColumnWidth(11);
        sheet.setColumnWidth(0, 5000);
        sheet.setColumnWidth(1, 12000);

        sheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));


        try {
            FileOutputStream fileOut = new FileOutputStream(filepath);
            workbook.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    private static ArrayList<XSSFCellStyle> colorOptions(XSSFWorkbook workbook){

        ArrayList<XSSFCellStyle> result = new ArrayList<>();

        XSSFCellStyle greenS,blueS,purpleS,redS,orangeS,yellowS;
        XSSFColor green,blue,purple,red,orange,yellow;

        XSSFFont legendFont = workbook.createFont();
        legendFont.setBold(true);
        legendFont.setFontHeightInPoints((short) 10);
        legendFont.setColor(IndexedColors.BLACK.getIndex());




        green = new XSSFColor();
        greenS = workbook.createCellStyle();
        greenS.setAlignment(HorizontalAlignment.CENTER);
        greenS.setFont(legendFont);
        green.setIndexed(IndexedColors.LIGHT_GREEN.getIndex());
        greenS.setFillForegroundColor(green);
        greenS.setFillBackgroundColor(green);
        greenS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(greenS);

        blue = new XSSFColor();
        blueS = workbook.createCellStyle();
        blueS.setAlignment(HorizontalAlignment.CENTER);
        blueS.setFont(legendFont);
        blue.setIndexed(IndexedColors.LIGHT_BLUE.getIndex());
        blueS.setFillForegroundColor(blue);
        blueS.setFillBackgroundColor(blue);
        blueS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(blueS);

        purple = new XSSFColor();
        purpleS = workbook.createCellStyle();
        purpleS.setFont(legendFont);
        purpleS.setAlignment(HorizontalAlignment.CENTER);
        purple.setIndexed(IndexedColors.LAVENDER.getIndex());
        purpleS.setFillForegroundColor(purple);
        purpleS.setFillBackgroundColor(purple);
        purpleS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(purpleS);

        red = new XSSFColor();
        redS = workbook.createCellStyle();
        redS.setFont(legendFont);
        redS.setAlignment(HorizontalAlignment.CENTER);
        red.setIndexed(IndexedColors.RED1.getIndex());
        redS.setFillForegroundColor(red);
        redS.setFillBackgroundColor(red);
        redS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(redS);

        orange = new XSSFColor();
        orangeS = workbook.createCellStyle();
        orangeS.setFont(legendFont);
        orangeS.setAlignment(HorizontalAlignment.CENTER);
        orange.setIndexed(IndexedColors.LIGHT_ORANGE.getIndex());
        orangeS.setFillForegroundColor(orange);
        orangeS.setFillBackgroundColor(orange);
        orangeS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(orangeS);

        yellow = new XSSFColor();
        yellowS = workbook.createCellStyle();
        yellowS.setFont(legendFont);
        yellowS.setAlignment(HorizontalAlignment.CENTER);
        yellow.setIndexed(IndexedColors.LIGHT_YELLOW.getIndex());
        yellowS.setFillForegroundColor(yellow);
        yellowS.setFillBackgroundColor(yellow);
        yellowS.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.add(yellowS);


        return result;

    }

    private static List generateColumns(NodeList trimLevelList, Integer amountOfTrimLevels) {
        List result = new ArrayList();

        List<String> firstColumnHeaders = new ArrayList();
        List<String> secondColumnHeaders = new ArrayList();
        List<String> thirdColumnHeaders = new ArrayList();


        firstColumnHeaders.addAll(Arrays.asList("Part Code", "Description"));
        secondColumnHeaders.addAll(Arrays.asList("-", "-"));
        thirdColumnHeaders.addAll(Arrays.asList("-", "-"));



        Node currentTrimLevel;

        for (int i = 0; i < amountOfTrimLevels; i++) {
            currentTrimLevel = trimLevelList.item(i);

            firstColumnHeaders.add(currentTrimLevel.getParentNode().getAttributes().getNamedItem("name").getNodeValue());

            secondColumnHeaders.add(currentTrimLevel.getParentNode().getAttributes().getNamedItem("code").getNodeValue());

            thirdColumnHeaders.add(currentTrimLevel.getAttributes().getNamedItem("code").getNodeValue());

        }

        //test

        System.out.println(firstColumnHeaders.toString());
        System.out.println(secondColumnHeaders.toString());
        System.out.println(thirdColumnHeaders.toString());

        result.add(firstColumnHeaders);
        result.add(secondColumnHeaders);
        result.add(thirdColumnHeaders);


        return result;


    }



}
