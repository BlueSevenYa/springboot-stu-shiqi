package com.excel.controller;

import com.excel.util.ExcelUtil;
import com.excel.vo.ResultResp;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.InputStream;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-02-22:15
 */
@RestController
@RequestMapping("/import")
public class ImportUserController {

    private static final Logger log= LoggerFactory.getLogger(ImportUserController.class);


    @PostMapping(value="/userExcel")
    public ResultResp<Object> userExcel(@RequestParam("fileUrl") String fileUrl,
                                        @RequestParam("storeId") String storeId){
        ResultResp<Object> resp=new ResultResp<>();
        log.info(fileUrl);
        log.info(storeId);
        resp=simpleCheckParam(fileUrl,storeId);
        if(resp.getCode() == 0){
            return resp;
        }
        resp= ExcelUtil.getInputStreamByUrl(fileUrl);
        if(resp.getCode() == 0){
            return resp;
        }
        InputStream inputStream= (InputStream) resp.getData();
        Workbook workbook= ExcelUtil.getWorkBookByIo(inputStream);
        if(!ExcelUtil.checkTemplateRight(workbook,ExcelUtil.notHeadExcelNum,ExcelUtil.headName)){
            resp.setCode(0);
            resp.setMessage("非法模板");
            return resp;
        }
        // 开始进行excel解析，进行数据校验，这一步基本保证了用户上传的模板是合法模板，也就是使用的提供的模板样式

        resp.setCode(1);
        resp.setData(null);
        return resp;
    }


    /**
     * 先进行简单的参数校验
     * @param fileUrl
     * @param storeId
     * @return
     */
    public ResultResp<Object> simpleCheckParam(String fileUrl,String storeId){
        ResultResp<Object> resp=new ResultResp<>();

        if(StringUtils.isBlank(fileUrl) || StringUtils.isBlank(storeId)){
            log.info("isBlank");
            resp.setCode(0);
            resp.setMessage("参数异常");
            return resp;
        }
        if(fileUrl.indexOf(".") == -1){
            log.info("indexOf");
            resp.setCode(0);
            resp.setMessage("参数异常");
            return resp;
        }
        if(!fileUrl.endsWith("xls") && !fileUrl.endsWith("xlsx")){
            log.info("endsWith");
            resp.setCode(0);
            resp.setMessage("文件格式异常");
            return resp;
        }
        resp.setCode(1);
        return resp;
    }
}
