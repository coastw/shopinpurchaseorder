/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.coast.controler;

import com.coast.model.Product;
import com.coast.model.ResultMSG;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Coast
 */
public class Controler {

    public static ResultMSG merge(String sapFile, String exportFile, String mergedFilePath) {
        ResultMSG msg = null;
        try {
            ArrayList<Product> products;
            //要导入到模板的SAP文件
            products = readProductsFromMyExcel(sapFile);
            //从上品网站导出的模板文件
            String inFile = exportFile;
            //最后上传到上品网站的文件
            int lastSlash = sapFile.lastIndexOf(File.separator);
            String outFileName = sapFile.substring(lastSlash+1, sapFile.length() - 5) + "_merged.xls";
            String outFile = mergedFilePath + File.separator + outFileName;
            //执行
            msg = writeProductsToExcel(products, inFile, outFile);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return msg;
        }
    }

    public static ResultMSG writeProductsToExcel(ArrayList<Product> products, String inFile, String outFile) throws Exception {
        ResultMSG resultMSG = new ResultMSG();
        int sum = 0;
        InputStream is = null;
        OutputStream os = null;
        try {
            File f = new File(outFile);
            f.delete();
            is = new FileInputStream(new File(inFile));
            HSSFWorkbook wb = new HSSFWorkbook(is);
            Sheet sheet = wb.getSheetAt(0);

            Iterator<Product> iter = products.iterator();
            while (iter.hasNext()) {
                Product product = iter.next();
                int thatRowNum = getRowNum(sheet, product.getSn(), product.getColor(), product.getSize());
                if (thatRowNum == 0) {
                    System.out.println("没有找到对应的SAP！sn=" + product.getSn() + " color=" + product.getColor() + " size=" + product.getSize() + " amount=" + product.getAmount());
                    String msg = "没有找到对应的SAP！sn=" + product.getSn() + " color=" + product.getColor() + " size=" + product.getSize() + " amount=" + product.getAmount() + "\n";
                    resultMSG.setMsg(msg);
                } else {
                    sheet.getRow(thatRowNum).createCell(6).setCellValue((double) product.getAmount());
                    sum += product.getAmount();
                }
            }

            //
            os = new FileOutputStream(new File(outFile));
            wb.write(os);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            is.close();
            os.close();
            resultMSG.setSum(sum);
            return resultMSG;
        }
    }

    public static ArrayList<Product> readProductsFromMyExcel(String file) throws Exception {
        ArrayList<Product> products = new ArrayList<Product>();
        InputStream is = null;
        int sum = 0;
        int row = 4;
        try {
            File f = new File(file);
            is = new FileInputStream(f);
//            HSSFWorkbook wb = new HSSFWorkbook(is);
            XSSFWorkbook wb = new XSSFWorkbook(is);
            Sheet sheet = wb.getSheetAt(0);
            while (row < sheet.getLastRowNum()) {
                if (sheet.getRow(row).getCell(1).getRichStringCellValue().toString().toUpperCase() == "") {
                    break;
                }
                String sn = sheet.getRow(row).getCell(1).getRichStringCellValue().toString().toUpperCase();
                String color = sheet.getRow(row).getCell(3).getRichStringCellValue().toString();
//                String size = sheet.getRow(row).getCell(11).getRichStringCellValue().toString();
                String size = trimSize(sheet.getRow(row).getCell(11).getRichStringCellValue().toString());
                double amount = sheet.getRow(row).getCell(23).getNumericCellValue();
                Product product = new Product();
                product.setSn(sn);
                product.setColor(color);
                product.setSize(size);
                product.setAmount(amount);
                products.add(product);
                sum += product.getAmount();
                row++;
            }
        } catch (Exception e) {
            System.out.println("readProductsFromMyExcel出现异常" + row);
            products = null;
            e.printStackTrace();
        } finally {
            is.close();
            System.out.println("读取总数:" + sum);
            return products;
        }
    }

    public static int getRowNum(Sheet sheet, String sn, String color, String size) throws Exception {
        int lastRowNum = sheet.getLastRowNum();//excell中左后一行显示为lastRowNum+1;
        int rowNum = lastRowNum;
        while (rowNum > 0) {
            if (sheet.getRow(rowNum).getCell(3).getRichStringCellValue().toString().trim().equals(sn)
                    && sheet.getRow(rowNum).getCell(4).getRichStringCellValue().toString().trim().equals(color)
                    && sheet.getRow(rowNum).getCell(5).getRichStringCellValue().toString().trim().equals(size)) {
                break;
            }
            rowNum--;
        }
        return rowNum;
    }

    public static String trimSize(String size) {
        String editedSize = "";
        String regex = "[cm]";
        editedSize = size.replaceAll(regex, "");
        return editedSize;
    }
}
