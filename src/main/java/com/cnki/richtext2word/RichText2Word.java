package com.cnki.richtext2word;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class RichText2Word {
    /**
     * html页面转word
     *
     * @return
     */
    public void exportWord() throws IOException, InterruptedException {
        String path = "E:\\kmis\\auditorData\\html";
        File file = new File(path);
        List<File> fileList = new ArrayList<>();

        if (file.exists()) {
            LinkedList<File> list = new LinkedList<File>();
            File[] files = file.listFiles();

            for (File html : files) {
                fileList.add(html);
            }

            int threadNum = 3;
            ExecutorService executorService = Executors.newFixedThreadPool(threadNum);
            CountDownLatch countDownLatch = new CountDownLatch(threadNum);
            int perSize = fileList.size() / threadNum;

            for (int i = 0; i < threadNum; i++) {
                MultiThread thread = new MultiThread();
                thread.setIdList(fileList.subList(i * perSize, (i + 1) * perSize));
                thread.setCountDownLatch(countDownLatch);
                executorService.submit(thread);
            }

            countDownLatch.await();
            executorService.shutdown();

        }
    }


    class MultiThread extends Thread {
        private List<File> idList;

        private CountDownLatch countDownLatch;

        public void setIdList(List<File> idList) {
            this.idList = idList;
        }

        public void setCountDownLatch(CountDownLatch countDownLatch) {
            this.countDownLatch = countDownLatch;
        }

        @Override
        public void run() {
            try {
                for (File html:idList
                     ) {
                    Transform tran = new Transform();
                    tran.jacob_html2word(html);
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (countDownLatch != null) {
                    countDownLatch.countDown();
                }
            }
        }
    }
}
