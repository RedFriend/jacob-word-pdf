package com.phy.jacob.controller;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.phy.jacob.util.JacobMultiUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.util.ObjectUtils;
import org.springframework.web.bind.annotation.*;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

@Controller
@CrossOrigin
@RequestMapping("/api")
@Slf4j
public class OfficeConverterController {

    @Value("${convertWord2PdfTempPath:D:/}")
    private String convertWord2PdfTempPath;

    @GetMapping("/stat")
    @ResponseBody
    public JSONObject wordFileToPdf() {
        JSONObject o = new JSONObject(true);
        o.put("fileQueue", JacobMultiUtil.fileQueue);
        o.put("appPidMap", JacobMultiUtil.appPidMap);
        o.put("appQueue", JacobMultiUtil.appMap);
        o.put("garbageFileQueue", JacobMultiUtil.garbageFileQueue);
        o.put("keepAliveMap", JacobMultiUtil.keepAliveMap);
        JSONArray array=new JSONArray();
        for (Thread thread : JacobMultiUtil.findAllThread()) {
            JSONObject object=new JSONObject();
            object.put("id",thread.getId());
            object.put("name",thread.getName());
            object.put("state",thread.getState());
            array.add(object);
        }
        o.put("thread", array);
        return o;
    }

    @PostMapping("/wordBase64ToPdf")
    @ResponseBody
    public String wordFileToPdf(@RequestBody JSONObject jsonObj, HttpServletRequest request) {
        String fileName = jsonObj.getString("fileName");
        String content = jsonObj.getString("content");
        long st = System.currentTimeMillis();
        System.out.println("\n=====调用转换PDF服务开始:" + getRemoteIp(request));
        log.info("\n入参fileName:{}", fileName);
        try {
            File inputFile = new File(convertWord2PdfTempPath + File.separatorChar + UUID.randomUUID().toString() + fileName.substring(fileName.lastIndexOf(".")));
            File outputFile = new File(convertWord2PdfTempPath + File.separatorChar + UUID.randomUUID().toString() + ".pdf");

            byte[] bytes = new BASE64Decoder().decodeBuffer(content);
            FileOutputStream out = new FileOutputStream(inputFile);
            out.write(bytes);
            out.close();

            // 将文件转换为pdf
            JacobMultiUtil.ConvertedTarget ct = new JacobMultiUtil.ConvertedTarget(inputFile, outputFile);
            JacobMultiUtil.init(JacobMultiUtil.MS_DOC);
            JacobMultiUtil.fileQueue.add(ct);
            ct.getCountDownLatch().await(60, TimeUnit.SECONDS);

            // 再将pdf转成Base64
            content = encodeBase64File(ct.getOutputFile().getAbsolutePath());

            System.out.println(String.format("\n=====调用转换PDF服务结束,耗时:%s ms", (System.currentTimeMillis() - st)));
            return content;
        } catch (Exception e) {
            System.err.println("文件转换异常");
            e.printStackTrace();
            return null;
        }
    }

    @PostMapping("/wordBase64ToPdfMulti")
    @ResponseBody
    public JSONObject[] wordFileToPdf(@RequestBody JSONObject[] jsonObjects, HttpServletRequest request) {
        List<JacobMultiUtil.ConvertedTarget> list = new ArrayList<>();
        long st = System.currentTimeMillis();
        System.out.println("\n=====调用转换PDF服务开始:" + getRemoteIp(request));
        try {
            JacobMultiUtil.init(JacobMultiUtil.MS_DOC);
            for (JSONObject jsonObject : jsonObjects) {
                String fileName = jsonObject.getString("fileName");
                String content = jsonObject.getString("content");
                log.info("\n入参fileName:" + fileName);
                File inputFile = new File(convertWord2PdfTempPath + File.separatorChar + UUID.randomUUID().toString() + fileName.substring(fileName.lastIndexOf(".")));
                File outputFile = new File(convertWord2PdfTempPath + File.separatorChar + UUID.randomUUID().toString() + ".pdf");

                byte[] bytes = new BASE64Decoder().decodeBuffer(content);
                FileOutputStream out = new FileOutputStream(inputFile);
                out.write(bytes);
                out.close();
                JacobMultiUtil.ConvertedTarget ct = new JacobMultiUtil.ConvertedTarget(inputFile, outputFile);
                list.add(ct);
            }

            //将文件放入转换队列
            JacobMultiUtil.fileQueue.addAll(list);
            //等待最后一个转换完毕
            list.get(list.size() - 1).getCountDownLatch().await(60, TimeUnit.SECONDS);
            for (int i = 0; i < list.size(); i++) {
                JacobMultiUtil.ConvertedTarget ct = list.get(i);
                if (ct.getCountDownLatch().getCount() > 0) {
                    ct.getCountDownLatch().await(60, TimeUnit.SECONDS);
                }
                // 再将pdf转成Base64
                jsonObjects[i].put("content", encodeBase64File(ct.getOutputFile().getAbsolutePath()));
            }

            System.out.println(String.format("\n=====调用转换PDF服务结束,本次共转换文件数:%s,耗时:%s ms", jsonObjects.length, (System.currentTimeMillis() - st)));
            return jsonObjects;
        } catch (Exception e) {
            System.err.println("文件转换位置异常");
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 获取ip地址
     *
     * @param request
     * @return
     */
    private String getRemoteIp(HttpServletRequest request) {
        if (ObjectUtils.isEmpty(request)) {
            return null;
        }
        String forwardIp = request.getHeader("x-forwarded-for");
        String realIp = request.getHeader("X-Real-IP");
        String remoteIp = request.getRemoteAddr();

        if (!ObjectUtils.isEmpty(forwardIp)) {
            return forwardIp.split(",")[0];
        } else if (!ObjectUtils.isEmpty(realIp)) {
            return realIp;
        } else {
            return remoteIp;
        }
    }

    /**
     * 将文件转成base64 字符串
     * 文件用base64编码 方便网络传输
     *
     * @param path
     * @return
     * @throws Exception
     */
    private String encodeBase64File(String path) throws Exception {
        File file = new File(path);
        ;
        FileInputStream inputFile = new FileInputStream(file);
        byte[] buffer = new byte[(int) file.length()];
        inputFile.read(buffer);
        inputFile.close();
        return new BASE64Encoder().encode(buffer);
    }
}
