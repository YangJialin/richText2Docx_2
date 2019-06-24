package com.cnki.richtext2word;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;


public class Transform {
    private static final String IMAG_PATH="E:\\kmis\\auditorData\\";
    private static final String WORD_PATH="E:\\kmis\\auditorData\\word\\";
    private static final String TEMP_PATH="E:\\kmis\\auditorData\\temp\\";
    private static final String TEMPLATE_PATH="E:\\kmis\\auditorData\\template.docx";



    public void jacob_html2word(File html) throws IOException {
        File isExist = new File(WORD_PATH+html.getName().substring(0,html.getName().lastIndexOf("."))+".docx");
        if (!isExist.exists()){
            MSOfficeGeneratorUtils officeUtils = new MSOfficeGeneratorUtils(false); // 将生成过程设置为不可见
            int imgIndex = 1;
            Map<String, String> imgMap = new HashMap<String, String>(); //存放图片标识符及物理路径  {"image_1","D:\img.png"};
            //替换图片
            Document document = Jsoup.parse(html,"UTF-8");
            Elements elements = document.select("img");

            for (Element img : elements){
                img.after("<p>${image_" + imgIndex + "}</p>"); // 为img添加同级p标签，内容为<p>${image_imgIndexNumber}</p>
                String src = img.attr("src").replace("/IIS/ewebeditor/", IMAG_PATH).replace("/","\\");
                // 下载图片到本地
                //download(src,"image_"+imgIndex,"D:\\zql\\imgs\\");
                // 保存图片标识符及物理路径
                //System.out.println("src:"+src+"    "+"${image_" + imgIndex++ + "}");

                imgMap.put("${image_" + imgIndex++ + "}", src);
                // 删除Img标签
                img.remove();
            }

            FileWriter temp = new FileWriter(TEMP_PATH+html.getName());
            if (new File(TEMP_PATH+html.getName()).exists()){
                new File(TEMP_PATH+html.getName()).delete();
            }
            temp.write(document.html(), 0, document.html().length());// 写入文件
            temp.flush();
            temp.close();


            File template = new File(TEMPLATE_PATH);
            String newFileName = WORD_PATH+html.getName().substring(0,html.getName().lastIndexOf("."))+".docx";
            File newFile=new File(newFileName);
            Files.copy(template.toPath(),newFile.toPath());

            // html文件转为word
            officeUtils.html2Word(TEMP_PATH+html.getName(),newFileName);

            // 替换标识符为图片
            for (Map.Entry<String, String> entry : imgMap.entrySet()){
                System.out.println("key:"+entry.getKey()+"    value:"+entry.getValue());
                officeUtils.replaceText2Image(entry.getKey(), entry.getValue());
            }
            officeUtils.saveAs(newFileName);    // 保存
            officeUtils.close(); // 关闭Office Word创建的文档
            officeUtils.quit(); // 退出Office Word程序

            // 这里可以删除本地图片 略去
            System.out.println(newFileName);
            Thread t = Thread.currentThread();
            String name = t.getName();
            System.out.println("Thread:" + name);
            imgIndex = 1;
            imgMap.clear();
        }else{

            Thread t = Thread.currentThread();
            String name = t.getName();
            System.out.println(isExist.getName()+"已存在 "+" Thread:" + name);
        }


    }





}