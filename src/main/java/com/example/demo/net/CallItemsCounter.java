package com.example.demo.net;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Data;
import lombok.NoArgsConstructor;
import okhttp3.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * @author keven.liu
 * @description:
 * @date 2021/11/9 13:36
 */
public class CallItemsCounter {
    private static OkHttpClient client = new OkHttpClient().newBuilder().build();
    private static SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd_HH_mm_ss");

    public static void main(String[] args) throws ParseException, IOException {
        int i = 0;
        while (true) {
            try {
                todo();
                System.out.println(i++);
                Thread.sleep(1000 * 60 * 5);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }

    private static void todo() throws ParseException {
        String savePath = "E:\\call\\数据";

        Date start = new SimpleDateFormat("yyyyMMdd HH:mm:ss").parse("20211110 09:00:00");
        Date end = new Date();
        String g = start.getTime() + "";
        String l = end.getTime() + "";

        XSSFWorkbook book = null;
        try {
            Set<Integer> seatIds = new HashSet<Integer>();
            //获取工作簿
            book = new XSSFWorkbook();
            XSSFSheet sheet = book.createSheet();
            List<TmpModel> data = getData(g, l);

            XSSFRow rowTitle = sheet.createRow(0);
            rowTitle.createCell(0).setCellValue("坐席id");
            rowTitle.createCell(1).setCellValue("新项目数量");
            rowTitle.createCell(2).setCellValue("旧项目数量");
            rowTitle.createCell(3).setCellValue("合计");
            rowTitle.createCell(4).setCellValue("等级");
            int idx = 1;
            for (int i = 0; i < data.size(); i++) {
                //去重数据，最早出现的是最后的数据，时间倒序排列
                TmpModel tmpModel = data.get(i);
                Integer seatId = tmpModel.getSeatId();
                if (seatIds.contains(seatId)) {
                    continue;
                } else {
                    seatIds.add(seatId);
                }
                XSSFRow row = sheet.createRow(idx);
                row.createCell(0).setCellValue(seatId);
                row.createCell(1).setCellValue(tmpModel.getCallNum());
                row.createCell(2).setCellValue(tmpModel.getCallHistoryNum());
                row.createCell(3).setCellValue(tmpModel.getCallNum() + tmpModel.getCallHistoryNum());
                row.createCell(4).setCellValue(tmpModel.getGrade());

                idx = idx + 1;
            }


            book.write(new FileOutputStream(new File(savePath + "\\拨打项目统计_" + formatter.format(new Date()) + ".xlsx")));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (book != null) {

            }
        }
    }

    private static List<TmpModel> getData(String g, String l) throws IOException {
        List<TmpModel> data = new ArrayList<>();
        MediaType mediaType = MediaType.parse("application/json");
        RequestBody body = RequestBody.create(mediaType, "{\n" +
                "  \"query\":{\n" +
                "        \"bool\": {\n" +
                "            \"must\": [\n" +
                "                {\n" +
                "                    \"query_string\": {\n" +
                "                        \"query\": \"txt:\\\"计数器新增，接听坐席\\\"\",\n" +
                "                        \"analyze_wildcard\": true,\n" +
                "                        \"default_field\": \"*\"\n" +
                "                    }\n" +
                "                },\n" +
                "                {\n" +
                "                    \"range\": {\n" +
                "                        \"@timestamp\": {\n" +
                "                            \"gte\": " + g + ",\n" +
                "                            \"lte\": " + l + ",\n" +
                "                            \"format\": \"epoch_millis\"\n" +
                "                        }\n" +
                "                    }\n" +
                "                }\n" +
                "            ],\n" +
                "            \"filter\": [],\n" +
                "            \"should\": [],\n" +
                "            \"must_not\": []\n" +
                "        }\n" +
                "    },\n" +
                "    \"sort\":[{\"@timestamp\":{\"order\":\"Desc\",\"unmapped_type\":\"boolean\"}}],\n" +
                "    \"size\":100\n" +
                "}");
        Request request = new Request.Builder()
                .url("http://10.50.6.208:5601/api/console/proxy?path=_search&method=POST")
                .method("POST", body)
                .addHeader("kbn-version", "6.8.6")
                .addHeader("Content-Type", "application/json")
                .build();
        Response response = client.newCall(request).execute();
        String result = response.body().string();
        JSONObject jsonObject = JSONObject.parseObject(result);
        JSONArray jsonArray = jsonObject.getJSONObject("hits").getJSONArray("hits");
        if (jsonArray.size() > 0) {
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject jsonObject1 = (JSONObject) jsonArray.get(i);
                //这里是日期
                String ts = jsonObject1.getJSONObject("_source").getString("ts");
                String msg = jsonObject1.getJSONObject("_source").getString("txt");

                int newIdx = msg.indexOf("新项目队列：");
                int hisIdx = msg.indexOf("旧项目队列：");

                List<TmpModel> tmpModels = JSON.parseArray(msg.substring(newIdx + 6, hisIdx - 1), TmpModel.class);
                List<TmpModel> tmpModels2 = JSON.parseArray(msg.substring(hisIdx + 6), TmpModel.class);
                data.addAll(tmpModels);
                data.addAll(tmpModels2);
            }
        }
        return data;
    }

    @NoArgsConstructor
    @Data
    public static class TmpModel {

        private Integer callHistoryNum;

        private Integer callNum;

        private Integer grade;

        private Integer maxCallNum;

        private Integer score;

        private Integer seatId;

    }
}
