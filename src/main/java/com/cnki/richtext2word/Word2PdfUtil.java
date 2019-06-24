package com.cnki.richtext2word;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

public class Word2PdfUtil {


    /**
     * html页面转word
     * wdFormatDocument = 0
     * wdFormatDocument97 = 0
     * wdFormatDocumentDefault = 16
     * wdFormatDOSText = 4
     * wdFormatDOSTextLineBreaks = 5
     * wdFormatEncodedText = 7
     * wdFormatFilteredHTML = 10
     * wdFormatFlatXML = 19
     * wdFormatFlatXMLMacroEnabled = 20
     * wdFormatFlatXMLTemplate = 21
     * wdFormatFlatXMLTemplateMacroEnabled = 22
     * wdFormatHTML = 8
     * wdFormatPDF = 17
     * wdFormatRTF = 6
     * wdFormatTemplate = 1
     * wdFormatTemplate97 = 1
     * wdFormatText = 2
     * wdFormatTextLineBreaks = 3
     * wdFormatUnicodeText = 7
     * wdFormatWebArchive = 9
     * wdFormatXML = 11
     * wdFormatXMLDocument = 12
     * wdFormatXMLDocumentMacroEnabled = 13
     * wdFormatXMLTemplate = 14
     * wdFormatXMLTemplateMacroEnabled = 15
     * wdFormatXPS = 18<br><br>over！
     * @return
     */
    public void exportPdf() throws IOException, InterruptedException {
        String path = "E:\\kmis\\pageoffice\\ofclWord";
        String pdfPath = "E:\\kmis\\pageoffice\\ofclPdf\\";
        File file = new File(path);
        List<File> fileList = new ArrayList<>();

        if (file.exists()) {
            LinkedList<File> list = new LinkedList<File>();
            File[] files = file.listFiles();

            for (File html : files) {
                String newFileName =pdfPath+html.getName().substring(0,html.getName().lastIndexOf("."))+".pdf";
                if(!new File(newFileName).exists()){
                    word2pdf(html.getAbsolutePath(),newFileName);
                }
            }
        }
    }











    private static ActiveXComponent app;
    /**
     * 单例模式
     */
    public static ActiveXComponent getWordInstance(){
        if (app == null) {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
        }
        return app;
    }

    public static void word2pdf(String source, String target) {
        int wdFormatPDF = 17;// word转PDF 格式
//        ComThread.InitSTA();
        ActiveXComponent app = null;
        try {
            app = getWordInstance();
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Dispatch doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();
            File tofile = new File(target);
            System.out.println("create PDF:"+tofile.getAbsolutePath());
//            if (tofile.exists()) {
//                tofile.delete();
//            }
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF);
            //save html
            String htmlTarget =target.replaceAll("pdf","html").replaceAll("Pdf","Html");
            if (!new File(htmlTarget).exists()){
                Dispatch.call(doc, "SaveAs", htmlTarget, 10);
            }
            Dispatch.call(doc, "Close", false);
            doc = null;
        } catch (Exception e) {
            System.out.println(e.toString());
        } /*finally {
            if (app != null) {
                app.invoke("Quit", 0);
            }
            ComThread.Release();
        }*/
    }
}
