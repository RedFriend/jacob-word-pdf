package com.phy.jacob.controller;

import com.alibaba.fastjson.JSONObject;
import com.phy.jacob.util.JacobMultiUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

@Controller
@CrossOrigin
@Slf4j
@RequestMapping("/api")
public class OfficeConverterController {

    /**
     * 将文件转成base64 字符串
     * 文件用base64编码 方便网络传输
     *
     * @param path
     * @return
     * @throws Exception
     */
    public static String encodeBase64File(String path) throws Exception {
        File file = new File(path);
        ;
        FileInputStream inputFile = new FileInputStream(file);
        byte[] buffer = new byte[(int) file.length()];
        inputFile.read(buffer);
        inputFile.close();
        return new BASE64Encoder().encode(buffer);
    }

    public static void main(String[] args) throws Exception {
        System.out.println(encodeBase64File("Z:\\template\\案外人申请再审案件\\民事裁定书(案外人申请再审案件，驳回案外人再审申请用).doc"));
    }

    @PostMapping("/wordBase64ToPdf")
    @ResponseBody
    public String wordFileToPdf(@RequestBody JSONObject jsonObj) {
        String fileName = jsonObj.getString("fileName");
        String content = jsonObj.getString("content");
        log.info("WORD转PDF,文件名:{}", fileName);
        try {
            File inputFile = new File(System.getenv("TEMP") + File.separatorChar + UUID.randomUUID().toString()+fileName.substring(fileName.lastIndexOf(".")));
            File outputFile = new File(System.getenv("TEMP") + File.separatorChar + UUID.randomUUID().toString()+".pdf");

            byte[] bytes = new BASE64Decoder().decodeBuffer(content);
            FileOutputStream out = new FileOutputStream(inputFile);
            out.write(bytes);
            out.close();

            // 将文件转换为pdf
            JacobMultiUtil.ConvertedTarget ct = new JacobMultiUtil.ConvertedTarget(inputFile, outputFile);
            JacobMultiUtil.init(JacobMultiUtil.MS_DOC);
            JacobMultiUtil.getFileQueue().add(ct);
            ct.getCountDownLatch().await(60, TimeUnit.SECONDS);

            // 再将pdf转成Base64
            content = encodeBase64File(ct.getOutputFile().getAbsolutePath());
            inputFile.delete();
            outputFile.delete();

            return content;
        } catch (Exception e) {
            log.error("文件转换异常", e);
            return null;
        }
    }
}
