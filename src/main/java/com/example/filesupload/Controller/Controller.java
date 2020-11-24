package com.example.filesupload.Controller;


import jxl.Cell;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.ls.LSOutput;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;


@org.springframework.stereotype.Controller
@RequestMapping
@ResponseBody
public class Controller {



    @PostMapping("/upload")
    public String upload(MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            throw new RuntimeException("文件为空");
        } else {
            System.out.println("文件大小为: "+ file.getSize());
            System.out.println("文件类型: "+file.getContentType());
            String originalFilename = file.getOriginalFilename();
            System.out.println("文档名: "+originalFilename);
            File path = new File("D:\\porject\\filesupload\\src\\main\\resources\\static\\" + originalFilename);
            File path1 = new File("D:\\porject\\filesupload\\src\\main\\resources\\static\\Back\\" + originalFilename);
            if (path.exists()||path1.exists()) {

                throw new RuntimeException("文件存在");
            } else{
                int res = originalFilename.indexOf(".");
                String substring = originalFilename.substring(res, originalFilename.length());
                if(substring.equals(".doc")&& path1.exists()==false){
                    System.out.println(originalFilename + "文档内容: ");
                    File file2 = new File("D:\\porject\\filesupload\\src\\main\\resources\\static\\Back\\"+ originalFilename);
                    file.transferTo(file2);
                    InputStream is = new FileInputStream(file2);
                    WordExtractor re = new WordExtractor(is);
                    String resullt = re.getText();
                    re.close();
                    System.out.println(resullt);

                }else if(substring.equals(".xls")){
                    System.out.println(originalFilename + "表内容: ");
                    File file2 = new File("D:\\porject\\filesupload\\src\\main\\resources\\static\\Back\\"+originalFilename);
                    file.transferTo(file2);
                    HSSFWorkbook workbook;
                    HSSFSheet sheet;
                 //   HSSFCell cell = null;
                    try {
                     workbook=new HSSFWorkbook(new POIFSFileSystem(file2));

                        sheet = workbook.getSheet("sheet1");
//                       cell = sheet.getCell(0, 0);
//                        System.out.println("标题："+cell.getContents());
                        int ii = 0;
                        int a = sheet.getLastRowNum();
                       // System.out.println(sheet.getPhysicalNumberOfRows());

                        System.out.println(a);
                     //  String empty="Null";
                        for (ii = 0; ii <=a; ii++) {
                            HSSFRow row =sheet.getRow(ii);
                            for (int j = 0; j <row.getLastCellNum(); j++) {
                                if (j == row.getLastCellNum()-1) {
                                    if(row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)==null){
                                        System.out.println("null"+"  ");
                                    }
                                    else {
                                       System.out.println(row.getCell(j));
                                    }

                                } else {
                                    if (row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)==null){
                                        System.out.print("null"+"  ");}
                                    else {
                                        System.out.print(row.getCell(j)+"  ");
                                    }
                                }

                            }
                        }

                    } catch (RuntimeException e) {
                        System.out.println("错误错误");
                    }

                }else{
                    file.transferTo(path);
                }



            }
        }



     return "上传"+file.getOriginalFilename()+"成功了" ;
       }

   }
