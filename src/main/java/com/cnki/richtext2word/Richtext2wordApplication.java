package com.cnki.richtext2word;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class Richtext2wordApplication {

	public static void main(String[] args) throws IOException, InterruptedException {
		SpringApplication.run(Richtext2wordApplication.class, args);
		/*//富文本转word。
		RichText2Word richText2Word = new RichText2Word();
        try {
            richText2Word.exportWord();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }*/
        //word转pdf、富文本。
        Word2PdfUtil word2PdfUtil = new Word2PdfUtil();
        word2PdfUtil.exportPdf();
    }

}
