package com.phy.jacob.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.CountDownLatch;
import java.util.stream.Collectors;

/**
 * jacob工具类
 *
 * @author Ryan.Peng
 * @date 2019年5月29日
 */
public class JacobMultiUtil {

    /**
     * 各启动方式宏
     */
    public final static String MS_DOC = "Word.Application";
    public final static String MS_EXCEL = "Excel.Application";
    public final static String MS_PPT = "Powerpoint.Application";
    public final static String WPS_WPS = "KWPS.Application";
    public final static String WPS_ET = "KET.Application";
    public final static String WPS_DPS = "KWPP.Application";
    private static ConcurrentLinkedQueue<ActiveXComponent> appQueue = new ConcurrentLinkedQueue<>();
    private static ConcurrentLinkedQueue<ConvertedTarget> fileQueue = new ConcurrentLinkedQueue<>();
    /**
     * 映射对应文档的文档对象类型
     **/
    private static Map<String, String> documentMap;
    /**
     * 转换成pdf文件对应的宏值的映射
     **/
    private static Map<String, Integer> pdfMacro;

    /**
     * 初始化映射关系
     */
    static {
        documentMap = new HashMap<>();
        pdfMacro = new HashMap<>();
        documentMap.put("Word.Application", "Documents");
        documentMap.put("Excel.Application", "Workbooks");
        documentMap.put("Powerpoint.Application", "Presentations");
        documentMap.put("KWPS.Application", "Documents");
        documentMap.put("KET.Application", "Workbooks");
        documentMap.put("KWPP.Application", "Presentations");
        pdfMacro.put("Word.Application", 17);
        pdfMacro.put("Powerpoint.Application", 32);
        pdfMacro.put("KWPS.Application", 17);
        pdfMacro.put("KWPP.Application", 32);
    }

    public static void init(String type) {
        try {
            if (appQueue.isEmpty()) {
//                String cmd = "taskkill /F /IM WPS.EXE";
                String cmd = "taskkill /F /IM WINWORD.EXE";
                Runtime.getRuntime().exec(cmd);
                Thread.sleep(200);
                int processorNum = Runtime.getRuntime().availableProcessors();
                ComThread.InitMTA();
                for (int i = 0; i < processorNum; i++) {
                    new Thread(new JacobMultiUtil.JacobThread(fileQueue, appQueue, type), "Jacob转换线程-" + (i + 1)).start();
                }
            }
        } catch (Exception e) {
            System.err.println("初始化Jacob进程错误");
        }
    }

    public static void quit() {
        try {
            for (ActiveXComponent app : appQueue) {
                if (app != null) {
                    app.invoke("Quit", new Variant[]{});
                    app = null;
                }
            }
            ComThread.Release();
            String cmd = "taskkill /F /IM WPS.EXE";
            Runtime.getRuntime().exec(cmd);
        } catch (IOException e) {
            System.err.println("清理Jacob进程错误");
        }
    }

    public static void main(String[] args) throws Exception {
        JacobMultiUtil.init(WPS_WPS);
        File templateDir = new File("Z:\\template");
        List<File> files = listFiles(templateDir).stream().filter(file -> file.getName().endsWith(".doc")).collect(Collectors.toList());
        List<ConvertedTarget> convertedTargets = new ArrayList<>();
        for (File file : files) {
            ConvertedTarget ct = new ConvertedTarget(file);
            convertedTargets.add(ct);
        }
        JacobMultiUtil.fileQueue.addAll(convertedTargets);
    }

    /**
     * 遍历目录文件
     *
     * @param directory
     * @return
     */
    public static List<File> listFiles(File directory) {

        ArrayList list = new ArrayList(1000);
        File[] files = directory.listFiles();
        for (File file : files) {
            if (file.isDirectory()) {
                list.addAll(listFiles(file));
            } else {
                list.add(file);
            }
        }
        return list;
    }

    /**
     * 按指定大小，分隔集合，将集合按规定个数分为n个部分
     *
     * @param <T>
     * @param list
     * @param len
     * @return
     */
    public static <T> List<List<T>> splitList(List<T> list, int len) {

        if (list == null || list.isEmpty() || len < 1) {
            return Collections.emptyList();
        }

        List<List<T>> result = new ArrayList<>();

        int size = list.size();
        int count = (size + len - 1) / len;

        for (int i = 0; i < count; i++) {
            List<T> subList = list.subList(i * len, ((i + 1) * len > size ? size : len * (i + 1)));
            result.add(subList);
        }

        return result;
    }

    public static class ConvertedTarget {
        //标识该文件已被处理,并通知线程
        private CountDownLatch countDownLatch = new CountDownLatch(1);
        private File inputFile;
        private File outputFile;
        private String base64File;
        private byte[] byteFile;

        public CountDownLatch getCountDownLatch() {
            return countDownLatch;
        }

        public void setCountDownLatch(CountDownLatch countDownLatch) {
            this.countDownLatch = countDownLatch;
        }

        public ConvertedTarget(File inputFile) {
            this.inputFile = inputFile;
        }
        public ConvertedTarget(File inputFile,File outputFile) {
            this.inputFile = inputFile;
            this.outputFile = outputFile;
        }

        public ConvertedTarget(String base64File) {
            this.base64File = base64File;
        }

        public ConvertedTarget(byte[] byteFile) {
            this.byteFile = byteFile;
        }

        public File getInputFile() {
            return inputFile;
        }

        public void setInputFile(File inputFile) {
            this.inputFile = inputFile;
        }

        public File getOutputFile() {
            return outputFile;
        }

        public void setOutputFile(File outputFile) {
            this.outputFile = outputFile;
        }

        public String getBase64File() {
            return base64File;
        }

        public void setBase64File(String base64File) {
            this.base64File = base64File;
        }

        public byte[] getByteFile() {
            return byteFile;
        }

        public void setByteFile(byte[] byteFile) {
            this.byteFile = byteFile;
        }
    }

    /**
     * Jacob线程类,加速处理
     */
    public static class JacobThread implements Runnable {
        private ConcurrentLinkedQueue<ActiveXComponent> appPool;
        private ConcurrentLinkedQueue<ConvertedTarget> files;
        private String type;

        JacobThread(ConcurrentLinkedQueue<ConvertedTarget> files, ConcurrentLinkedQueue<ActiveXComponent> appPool, String type) {
            this.files = files;
            this.appPool = appPool;
            this.type = type;
        }

        @Override
        public void run() {
            Dispatch documents;
            String filePath = null;
            ActiveXComponent app = new ActiveXComponent(type);
            app.setProperty("Visible", new Variant(false));
            documents = app.getProperty(documentMap.get(type)).toDispatch();
            appPool.offer(app);
            while (true) {
                ConvertedTarget f = files.poll();
                try {
                    if (f != null) {
                        long st=System.currentTimeMillis();
                        filePath = f.getInputFile().getAbsolutePath();
                        Dispatch document = Dispatch.call(documents, "Open", filePath).toDispatch();
                        String outFilePath = filePath.substring(0, filePath.lastIndexOf(".")) + ".pdf";
                        Dispatch.call(document, "SaveAs", outFilePath, new Variant(pdfMacro.get(type)));
                        Dispatch.call(document, "Close", false);
                        System.out.println(Thread.currentThread().getName() + " PDF转换成功,文件位置:" + outFilePath);
                        File outputFile=new File(outFilePath);
                        if(outputFile.isFile()){
                            f.setOutputFile(outputFile);
                            f.countDownLatch.countDown();
                        }else{
                            System.err.println(Thread.currentThread().getName() + " 转换文件异常,重新进入队列,文件位置:" + filePath);
                            files.offer(f);
                        }
                        long et=System.currentTimeMillis();
                        System.out.println(Thread.currentThread().getName() +"转换耗时: "+(et-st));
                    } else {
//                        System.out.println(Thread.currentThread().getName() + " 队列中没有需要转换的文件");
                        Thread.sleep(200);
                    }
                } catch (Exception e) {
                    System.err.println(Thread.currentThread().getName() + " 转换异常,重新进入队列,文件位置:" + filePath);
                    if(f!=null){
                        files.offer(f);
                    }
                }
            }
        }
    }

    public static ConcurrentLinkedQueue<ActiveXComponent> getAppQueue() {
        return appQueue;
    }

    public static void setAppQueue(ConcurrentLinkedQueue<ActiveXComponent> appQueue) {
        JacobMultiUtil.appQueue = appQueue;
    }

    public static ConcurrentLinkedQueue<ConvertedTarget> getFileQueue() {
        return fileQueue;
    }

    public static void setFileQueue(ConcurrentLinkedQueue<ConvertedTarget> fileQueue) {
        JacobMultiUtil.fileQueue = fileQueue;
    }
}
