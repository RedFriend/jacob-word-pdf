package com.phy.jacob.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * jacob工具类
 * @author Ryan.Peng
 * @date 2019年5月29日
 */
public class JacobUtil {
    //各启动方式宏
    public final static String MICROSOFT_DOC   = "Word.Application";
    public final static String MICROSOFT_EXCEL = "Excel.Application";
    public final static String MICROSOFT_PPT   = "Powerpoint.Application";
    public final static String WPS_WPS = "KWPS.Application";
    public final static String WPS_ET  = "KET.Application";
    public final static String WPS_DPS = "KWPP.Application";

    /** 映射对应文档的文档对象类型 **/
    private static Map<String, String> documentMap;

    /** 转换成pdf文件对应的宏值的映射 **/
    private static Map<String, Integer> pdfMacro;

    /** 应用程序对象 **/
    private ActiveXComponent app;

    /** 所有文档对象 **/
    private Dispatch documents;

    /** 当前活动文档对象 **/
    private Dispatch document;

    /** 是否可视化运行方式 **/
    private boolean visible = false;

    /** 当前应用程序类型 **/
    private String type;

    /** 文档路径 **/
    private String filePath;

    /**
     * 初始化映射关系
     */
    static {
        documentMap = new HashMap<>();
        pdfMacro    = new HashMap<>();
        documentMap.put("Word.Application"       , "Documents");
        documentMap.put("Excel.Application"      , "Workbooks");
        documentMap.put("Powerpoint.Application" , "Presentations");
        documentMap.put("KWPS.Application", "Documents");
        documentMap.put("KET.Application" , "Workbooks");
        documentMap.put("KWPP.Application", "Presentations");
        pdfMacro.put("Word.Application"       , 17);
        pdfMacro.put("Powerpoint.Application" , 32);
        pdfMacro.put("KWPS.Application", 17);
        pdfMacro.put("KWPP.Application", 32);
    }

    /**
     * 根据传入类型打开对应的应用程序 打开相应文档
     * @param filePath 文档路径
     * @param type     文档打开所用的应用程序
     * @param visible  是否可视化运行方式（默认false）
     */
    public void open( String filePath, String type, boolean visible ) {
        this.type = type;
        this.filePath = filePath;
        this.visible = visible;
        try {
            app = new ActiveXComponent(type);
            app.setProperty("Visible", new Variant(visible));
            documents = app.getProperty(documentMap.get(type)).toDispatch();
            document = Dispatch.call(documents, "Open", filePath).toDispatch();
            System.out.println("JacobUtil.class : 打开文档   " + filePath);
        } catch (Exception e) {
            System.err.println("JacobUtil.class : 打开文档失败 " + filePath);
            e.printStackTrace();
        }
    }

    /**
     * 转换成pdf文件并保存 默认当前目录
     */
    public boolean toPDF(String outFilePath) {
        boolean bResult = false;
        try {
            if ( type.equals(MICROSOFT_EXCEL) || type.equals(WPS_ET) ) {
                Dispatch.call( document, "ExportAsFixedFormat", Dispatch.Method,outFilePath, new Variant(0) );
            } else {
                Dispatch.call(document, "SaveAs", outFilePath, new Variant(pdfMacro.get(type)));
            }
            bResult = true;
        } catch (Exception e) {
            System.err.println("JacobUtil.class : 文档转换为pdf失败！");
            bResult = false;
            e.printStackTrace();
        }
        return bResult;
    }

    /**
     * 转换成pdf文件并保存 默认当前目录
     */
    public boolean toPDF() {
        boolean bResult = false;
        String outFilePath = filePath.substring(0, filePath.lastIndexOf(".")) + ".pdf";
        try {
            return toPDF(outFilePath);
        } catch (Exception e) {
            System.err.println("JacobUtil.class : 文档转换为pdf失败！");
            bResult = false;
        }
        return bResult;
    }

    /**
     * 根据传入类型打开对应的应用程序 打开相应文档
     * @param filePath 文档路径
     * @param type     文档打开所用的应用程序
     */
    public void open( String filePath, String type ) {
        try {
            open(filePath, type, visible);
        }catch (Exception e) {
            System.err.println("JacobUtil.class : 打开文档失败 " + filePath);
            e.printStackTrace();
        }
    }

    /**
     * 关闭文档以及应用程序
     */
    public void close () {
        try {
            Dispatch.call(document, "Close", false);
            if (app != null){
                app.invoke("Quit", new Variant[] {});
                app = null;
            }
//            System.out.println("JacobUtil.class : 文档关闭  " + filePath);
        } catch (Exception e) {
            System.err.println("关闭ActiveXComponent异常");
            e.printStackTrace();
        }
    }

    public boolean isVisible() {
        return visible;
    }

    public void setVisible(boolean visible) {
        this.visible = visible;
    }
    public static void main(String[] args) {
        File templateDir=new File("Z:\\template");
        List<File> files= listFiles(templateDir);
        for (File f : files) {
            JacobUtil jacob = new JacobUtil();
            jacob.open(f.getAbsolutePath(), JacobUtil.WPS_WPS);
            //默认同路径 也可自定义
            jacob.toPDF();
            jacob.close();
        }
    }

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
}
