package com.phy.jacob.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.*;

public class WmiServiceUtils {
    public static final Logger logger = LoggerFactory.getLogger(WmiServiceUtils.class);

    private static List<Map<String, Object>> getAllResult(String[] cmdStr, int flag) throws IOException {
        List<Map<String, Object>> list = new ArrayList<>();
        Integer index = 1;
        Process p = null;
        String str = null;
        String[] arrStr = new String[2];
        Map<String, Object> map = new HashMap<String, Object>();
        InputStreamReader isr = null;
        BufferedReader br = null;
        try {
            p = Runtime.getRuntime().exec(cmdStr);
            isr = new InputStreamReader(p.getInputStream());
            br = new BufferedReader(isr);
            while ((str = br.readLine()) != null) {
                if (!StringUtils.isEmpty(str)) {
                    if (index % flag == 0) {
                        list.add(map);
                        map = new HashMap<String, Object>();
                    }
                    arrStr = str.split("=");
                    str = str.endsWith("=") ? "" : arrStr[1];
                    map.put(arrStr[0], str);
                    index++;
                }
            }
        } catch (IOException e) {
            logger.error("获取进程的所有信息失败！", e);
            throw e;
        } catch (Exception e) {
            logger.error("获取执行结果失败！", e);
            throw e;
        } finally {
            try {
                if (br != null) {
                }
                br.close();
                if (isr != null) {
                    isr.close();
                }
            } catch (IOException e) {
                logger.error("", e);
            }
            if (p != null) {
                p.destroy();
            }
        }
        return list;
    }

    @SuppressWarnings("unused")
    private static String parse2Time(long milliseconds) {
        if (milliseconds == 0) {
            return "0秒";
        }
        if (milliseconds / 1000 == 0) {
            return "0." + milliseconds + "秒";
        }
        milliseconds = milliseconds / 1000;
        long day = milliseconds / (24 * 3600);
        milliseconds = milliseconds % (24 * 3600);
        if (milliseconds == 0) {
            return day + "天";
        }
        long hour = milliseconds / (3600);
        milliseconds = milliseconds % (3600);
        if (milliseconds == 0) {
            return day + "天" + hour + "小时";
        }
        long minute = milliseconds / (60);
        milliseconds = milliseconds % (60);
        if (milliseconds == 0) {
            return day + "天 " + hour + "小时 " + minute + "分钟";
        } else {
            return day + "天 " + hour + "小时 " + minute + "分钟 " + milliseconds + "秒";
        }
    }

    private static Map<String, Object> printStream(InputStream input) throws IOException { //InputStream input  final Process proc
        InputStreamReader isr = new InputStreamReader(input); //proc.getInputStream()
        BufferedReader br = new BufferedReader(isr);
        Map<String, Object> map = new HashMap<String, Object>();
        String str = null;
        String[] arrStr = null;
        try {
            while ((str = br.readLine()) != null) {
                if (!StringUtils.isEmpty(str)) {
                    if (str.contains("=")) {
                        arrStr = str.split("=");
                        str = str.endsWith("=") ? "" : arrStr[1];
                        map.put(arrStr[0], str);
                    } else {
                        map.put(str, null);
                    }
                }
            }
        } catch (IOException e) {
            logger.error("关闭文件流失败！", e);
            throw e;
        } finally {
            try {
                if (br != null) {
                    br.close();
                }
                if (isr != null) {
                    isr.close();
                }
                if (input != null) {
                    input.close();
                }
            } catch (IOException e) {
                logger.error("关闭文件流失败！", e);
                throw e;
            }
        }
        return map;
    }

    private static String printErrorStream(InputStream input) throws IOException {
        InputStreamReader reader = new InputStreamReader(input);
        BufferedReader br = new BufferedReader(reader);
        String msg = "";
        String str = "";
        try {
            while ((str = br.readLine()) != null) {
                if (!StringUtils.isEmpty(str)) {
                    msg += str + ",";
                }
            }
            if (msg.endsWith(",")) {
                msg.substring(0, msg.lastIndexOf(","));
            }
            return msg;
        } catch (IOException e) {
            logger.error("读取错误信息失败！", e);
            throw e;
        } finally {
            try {
                if (br != null) {
                    br.close();
                }
                if (reader != null) {
                    reader.close();
                }
                if (input != null) {
                    input.close();
                }
            } catch (IOException e) {
                logger.error("关闭文件流失败！", e);
                throw e;
            }
        }
    }

    private static Map<String, Object> execCommand(String[] cmdStr) throws IOException {
        Process p = null;
        Map<String, Object> map = new HashMap<>();
        try {
            p = Runtime.getRuntime().exec(cmdStr);
            logger.info("执行错误信息： " + printErrorStream(p.getErrorStream()));
            map = printStream(p.getInputStream());
        } catch (IOException e) {
            logger.error("启动服务失败！", e);
            throw e;
        } catch (Exception e) {
            logger.error("获取执行结果失败！", e);
            throw e;
        } finally {
            if (p != null) {
                p.destroy();
            }
        }
        return map;
    }

    /**
     * 启动服务
     *
     * @param serviceName 右键 指定服务项-》属性 -》服务名称
     * @return
     * @throws IOException
     */
    public static Map<String, Object> startService(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", " net start " + serviceName};//runAs /user:Administrator
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("开启服务失败！", e);
            throw e;
        }
    }

    /**
     * 关闭服务
     *
     * @param serviceName 右键 指定服务项-》属性 -》服务名称
     * @return
     * @throws IOException
     */
    public static Map<String, Object> stopService(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "net stop " + serviceName};//runAs /user:Administrator
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * 禁用服务
     *
     * @param serviceName
     * @return
     * @throws IOException
     */
    public static Map<String, Object> disableService(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "sc config " + serviceName + " start= disabled"};//runAs /user:Administrator
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * 启用服务 --自动
     *
     * @param serviceName
     * @return
     * @throws IOException
     */
    public static Map<String, Object> enableService(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "sc config " + serviceName + " start= auto"};//runAs /user:Administrator
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * 启用服务 --手动
     *
     * @param serviceName
     * @return
     * @throws IOException
     */
    public static Map<String, Object> demandService(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "sc config " + serviceName + " start= demand"};//runAs /user:Administrator
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * @param taskName 映像名称  XXXX.exe
     * @return
     * @throws IOException
     */
    public static Map<String, Object> getTaskDetail(String taskName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic process where name='" + taskName + "' list full"};
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * @param processId PID
     * @return
     * @throws IOException
     */
    public static Map<String, Object> getTaskDetail(Integer processId) throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic process where processid='" + processId + "' list full"};// get /format:value
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    public static List<Map<String, Object>> getAllTaskDetail() throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic process get /value"};
        logger.info(Arrays.toString(cmdStr));
        List<Map<String, Object>> list = null;
        try {
            list = getAllResult(cmdStr, 45);
        } catch (IOException e) {
            logger.error("获取所有进程信息失败！", e);
            throw e;
        }
        return list;
    }

    public static List<Map<String, Object>> getAllService() throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic service get /value"};
        logger.info(Arrays.toString(cmdStr));
        List<Map<String, Object>> list = null;
        try {
            list = getAllResult(cmdStr, 25);
        } catch (IOException e) {
            logger.error("获取所有服务信息失败！", e);
            throw e;
        }
        return list;
    }

    /**
     * @param serviceName 右键 指定服务项-》属性 -》服务名称
     * @return
     * @throws IOException
     */
    public static Map<String, Object> getServiceDetail(String serviceName) throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic service where name='" + serviceName + "' list full"};
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * @param processId PID
     * @return
     * @throws IOException
     */
    public static Map<String, Object> getServiceDetail(Integer processId) throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic service where processid='" + processId + "' list full"};
        logger.info(Arrays.toString(cmdStr));
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    public static Map<String, Object> createProcess(String taskpath) throws IOException {
        String[] cmdStr = {"cmd", "/C", "wmic process call create'" + taskpath + "'"};
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    public static Map<String, Object> deleteProcess(String taskname) throws IOException {
        String[] cmdStr = {"cmd", "/C", " wmic process where name='" + taskname + "' delete"};//runAs /user:Administrator
        try {
            return execCommand(cmdStr);
        } catch (IOException e) {
            logger.error("关闭服务失败！", e);
            throw e;
        }
    }

    /**
     * 计算某进程cpu使用率
     *
     * @param processName
     * @return
     * @throws Exception sysTime ：表示该时间段内总的CPU时间=CPU处于用户态和内核态CPU时间的总和，即sysTime =kerneTimel + userTime（注：这里并不包括idleTime，因为当CPU处于空闲状态时，实在内核模式下运行System Idle Process这个进程，所以kernelTime实际上已经包含了idleTime）；
     *                   idleTime：表示在该时间段内CPU处于空闲状态的时间；CPU% = 1 – idleTime / sysTime * 100
     */
    public static String getCpuRatioForWindows(String processName) throws Exception {
        String[] cmdStr = {"cmd", "/C",
                "wmic process get Caption,CommandLine,KernelModeTime,ReadOperationCount,ThreadCount,UserModeTime,WriteOperationCount /value"};
        try {
            List<Map<String, Object>> list1 = getAllResult(cmdStr, 7);
            long[] data1 = getCpuTime(list1, processName);
            Thread.sleep(1000);
            List<Map<String, Object>> list2 = getAllResult(cmdStr, 7);// get(p.getInputStream());
            long[] data2 = getCpuTime(list2, processName);
            long proctime = data2[2] - data1[2];
            long totaltime = data2[1] - data1[1]; // + data2[0] - data1[0]
            if (totaltime == 0) {
                return "0%";
            }
            return Double.valueOf(10000 * proctime * 1.0 / totaltime).intValue() / 100.00 + "%";
        } catch (Exception e) {
            logger.error("获取CPU占用率失败！", e);
            throw e;
        }
    }

    public static void main(String[] args) throws Exception{
        System.out.println(getCpuRatioForWindows("wps.exe"));
    }

    private static long[] getCpuTime(List<Map<String, Object>> list, String processName) {
        long[] data = new long[3];
        long idletime = 0;
        long kneltime = 0;
        long usertime = 0;
        long processTime = 0;
        String caption = "";
        String kmtm = "";
        String umtm = "";
        for (Map<String, Object> m : list) {
            caption = m.get("Caption").toString();
            kmtm = m.get("KernelModeTime").toString();
            umtm = m.get("UserModeTime").toString();
            if (caption.equals("System Idle Process") || caption.equals("System")) {
                if (kmtm != null && !kmtm.equals("")) {
                    idletime += Long.parseLong(kmtm);
                }
                if (umtm != null && !umtm.equals("")) {
                    idletime += Long.parseLong(umtm);
                }
            }
            if (caption.equals(processName)) {
                if (kmtm != null && !kmtm.equals("")) {
                    processTime += Long.parseLong(kmtm);
                }
                if (umtm != null && !umtm.equals("")) {
                    processTime += Long.parseLong(umtm);
                }
            }
            if (kmtm != null && !kmtm.equals("")) {
                kneltime += Long.parseLong(kmtm);
            }
            if (umtm != null && !umtm.equals("")) {
                usertime += Long.parseLong(umtm);
            }
        }
        data[0] = idletime;
        data[1] = kneltime + usertime;
        data[2] = processTime;
        return data;
    }
}
