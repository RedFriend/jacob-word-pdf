package com.phy.jacob.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.Getter;
import lombok.Setter;
import org.springframework.util.FileCopyUtils;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.ConcurrentMap;
import java.util.concurrent.CountDownLatch;

/**
 * jacob工具类
 *
 * @author Ryan.Peng
 * @date 2019年5月29日
 */
@Getter
@Setter
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
    private static Queue<ActiveXComponent> appQueue = new ConcurrentLinkedQueue<>();
    private static ConcurrentMap<ActiveXComponent, Long> keepAliveMap = new ConcurrentHashMap<>();
    private static Queue<ConvertedTarget> fileQueue = new ConcurrentLinkedQueue<>();
    private static Queue<ConvertedTarget> garbageQueue = new ConcurrentLinkedQueue<>();
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
            createAppThread(type);
            keepAliveThread(type);
            garbageThread();
        } catch (Exception e) {
            System.err.println("初始化Jacob进程错误");
            e.printStackTrace();
        }
    }

    public static void createAppThread(String type) throws Exception {
        if (appQueue.isEmpty()) {
            System.out.println("正在初始化转换程序...");
            System.out.println("结束操作系统Offince进程");
            String cmd = "taskkill /F /IM WPS.EXE";
            String cmd2 = "taskkill /F /IM WINWORD.EXE";
            Runtime.getRuntime().exec(cmd);
            Runtime.getRuntime().exec(cmd2);
            Thread.sleep(200);
            int processorNum = Runtime.getRuntime().availableProcessors();
            System.out.println("启用多线程支持");
            ComThread.InitMTA();
            for (int i = 0; i < processorNum; i++) {
                new Thread(new JacobMultiUtil.JacobThread(fileQueue, appQueue, type), "Converter - " + (i + 1)).start();
                System.out.println("创建线程:" + "Converter - " + (i + 1));
            }
        }
    }

    public static void garbageThread() {
        System.out.println("创建临时文件清理进程");
        new Thread(() -> {
            while (true) {
                if (!garbageQueue.isEmpty()) {
                    int i = 1;
                    for (ConvertedTarget ct : garbageQueue) {
                        ct.getInputFile().delete();
                        ct.getOutputFile().delete();
                        i++;
                    }
                    System.out.println("本次清理了临时文件个数:" + 2 * i);
                }
                try {
                    Thread.sleep(30 * 60 * 1000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }).start();
    }


    public static void keepAliveThread(String type) {
        System.out.println("创建keepAlive进程");
        new Thread(() -> {
            while (true) {
                try {
                    Long now = System.currentTimeMillis();
                    for (ActiveXComponent app : keepAliveMap.keySet()) {
                        Long alive = keepAliveMap.get(app);
                        if (now - alive > 5 * 60 * 1000) {
                            System.out.println("发现僵死进程,进程上次活跃时间:" + (now - alive) + "ms");
                            //队列中移除app
                            appQueue.remove(app);

                            if (app != null) {
                                app.invoke("Quit", new Variant[]{});
                                System.out.println("结束僵死进程完毕");
                                app = null;
                                new Thread(new JacobThread(fileQueue, appQueue, type)).start();
                                System.out.println("创建新进程替换僵死进程");
                            }
                        }
                    }
                    Thread.sleep(30 * 1000);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }).start();
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
            String cmd2 = "taskkill /F /IM WINWORD.EXE";
            Runtime.getRuntime().exec(cmd);
            Runtime.getRuntime().exec(cmd2);
        } catch (IOException e) {
            System.err.println("清理Jacob进程错误");
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        JacobMultiUtil.init(WPS_WPS);
        File templateDir = new File("C:\\Users\\pengh\\Desktop\\2e39a731-7586-4263-b626-51b6e8bbabd4.doc");
        File templateDir2 = new File("C:\\Users\\pengh\\Desktop\\2e39a731-7586-4263-b626-51b6e8bbabd4.pdf");
//        List<File> files = listFiles(templateDir).stream().filter(file -> file.getName().endsWith(".doc")).collect(Collectors.toList());
        List<ConvertedTarget> convertedTargets = new ArrayList<>();
        ConvertedTarget ct = new ConvertedTarget(templateDir, templateDir2);
        convertedTargets.add(ct);
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

    public static Queue<ActiveXComponent> getAppQueue() {
        return appQueue;
    }

    public static void setAppQueue(Queue<ActiveXComponent> appQueue) {
        JacobMultiUtil.appQueue = appQueue;
    }

    public static Queue<ConvertedTarget> getFileQueue() {
        return fileQueue;
    }

    public static void setFileQueue(Queue<ConvertedTarget> fileQueue) {
        JacobMultiUtil.fileQueue = fileQueue;
    }

    public static class ConvertedTarget {
        //标识该文件已被处理,并通知线程
        private CountDownLatch countDownLatch = new CountDownLatch(1);
        private File inputFile;
        private File outputFile;
        private String base64File;
        private byte[] byteFile;
        private int tryCount;

        public ConvertedTarget(File inputFile) {
            this.inputFile = inputFile;
        }

        public ConvertedTarget(File inputFile, File outputFile) {
            this.inputFile = inputFile;
            this.outputFile = outputFile;
        }

        public ConvertedTarget(String base64File) {
            this.base64File = base64File;
        }

        public ConvertedTarget(byte[] byteFile) {
            this.byteFile = byteFile;
        }

        public int getTryCount() {
            return tryCount;
        }

        public void setTryCount(int tryCount) {
            this.tryCount = tryCount;
        }

        public CountDownLatch getCountDownLatch() {
            return countDownLatch;
        }

        public void setCountDownLatch(CountDownLatch countDownLatch) {
            this.countDownLatch = countDownLatch;
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
        private Queue<ActiveXComponent> appPool;
        private Queue<ConvertedTarget> files;
        private String type;

        JacobThread(Queue<ConvertedTarget> files, Queue<ActiveXComponent> appPool, String type) {
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
                //更新进程活跃状态
                keepAliveMap.put(app, System.currentTimeMillis());
                try {
                    if (f != null) {
                        long st = System.currentTimeMillis();
                        filePath = f.getInputFile().getAbsolutePath();
                        Dispatch document = Dispatch.call(documents, "Open", filePath).toDispatch();
                        String outFilePath = filePath.substring(0, filePath.lastIndexOf(".")) + ".pdf";
                        Dispatch.call(document, "SaveAs", outFilePath, new Variant(pdfMacro.get(type)));
                        Dispatch.call(document, "Close", false);
                        System.out.println("\n" + Thread.currentThread().getName() + " PDF转换成功,文件位置:\n" + outFilePath);
                        File outputFile = new File(outFilePath);
                        if (outputFile.isFile()) {
                            f.setOutputFile(outputFile);
                            f.countDownLatch.countDown();
                            //正常转换转入待清理队列
                            garbageQueue.add(f);
                        } else {
                            if (f.getTryCount() < 3) {
                                f.setTryCount(f.getTryCount() + 1);
                                System.err.println("\n" + Thread.currentThread().getName() + " 转换文件异常,重新进入队列,文件位置:\n" + filePath);
                                files.offer(f);
                            } else {
                                System.err.println("\n" + Thread.currentThread().getName() + " 3次尝试转换文件异常,已排除队列,文件位置:\n" + filePath);
                                FileCopyUtils.copy(f.getInputFile(), new File(System.getenv("TEMP") + File.separatorChar + f.getInputFile().getName()));
                            }
                        }
                        long et = System.currentTimeMillis();
                        System.out.println(Thread.currentThread().getName() + " 转换耗时: " + (et - st) + "ms");
                    } else {
//                        System.out.println(Thread.currentThread().getName() + " 队列中没有需要转换的文件");
                        Thread.sleep(50);
                    }
                } catch (Exception e) {
                    System.err.println("\n" + Thread.currentThread().getName() + " 转换异常,重新进入队列,文件位置:\n" + filePath);
                    e.printStackTrace();
                    if (f != null && f.getTryCount() < 3) {
                        f.setTryCount(f.getTryCount() + 1);
                        files.offer(f);
                    }
                }
            }
        }
    }
}
