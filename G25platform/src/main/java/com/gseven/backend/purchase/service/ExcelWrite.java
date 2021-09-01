package com.gseven.backend.purchase.service;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.jdbc.core.JdbcTemplate;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.gseven.backend.logistics.factory.MigoFactory;
import com.gseven.backend.logistics.factory.ValetFactory;
import com.gseven.backend.purchase.factory.purchaseFactory;
import com.gseven.backend.purchase.factory.reportFactory;
import com.gseven.backend.sales.entity.PriceRelations;
import com.gseven.backend.sys.service.DateUtil;
import com.gseven.backend.sys.service.DownloadUtility;
import com.gseven.backend.sys.service.ExcelUtil;
import com.gseven.backend.werp2TT.factory.ComparisonFactory;

public class ExcelWrite {

	public ExcelWrite() {
	}
	
//PSI匯出
	public void get_psi_down(String fileName, List<Map<String, Object>> IMA) throws IOException {
		try {

			Calendar c = Calendar.getInstance();
			c.add(Calendar.MONTH, 0);
			SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
			String time0 = format.format(c.getTime());

			c.add(Calendar.MONTH, 1);
			format = new SimpleDateFormat("yyyy-MM");
			String time1 = format.format(c.getTime());

			c.add(Calendar.MONTH, 2);
			format = new SimpleDateFormat("yyyy-MM");
			String time2 = format.format(c.getTime());

			c.add(Calendar.MONTH, 3);
			format = new SimpleDateFormat("yyyy-MM");
			String time3 = format.format(c.getTime());

			c.add(Calendar.MONTH, 4);
			format = new SimpleDateFormat("yyyy-MM");
			String time4 = format.format(c.getTime());

			c.add(Calendar.MONTH, 5);
			format = new SimpleDateFormat("yyyy-MM");
			String time5 = format.format(c.getTime());

			List<String> list = new ArrayList<String>();

			list.add("企業別");
			list.add("供應商");
			list.add("商品狀態");
			list.add("採購大類");
			list.add("商品大類");
			list.add("品牌");
			list.add("品號");
			list.add("品號總售訂量");
			list.add("品號總量");
			list.add("存貨倉備貨庫存總量(備貨數量)");
			list.add("存貨倉陳列庫存總量(陳列數量)");
			list.add("存貨倉庫存總量");
			list.add("承銷倉備貨庫存總量");
			list.add("承銷倉陳列庫存總量");
			list.add("承銷倉庫存總量(承銷數量)");
			list.add("在途倉庫存總量");
			list.add("其他倉庫儲位庫存總量");
			list.add(time0 + "\n售訂未銷量");
			list.add(time0 + "\n 可再接單");
			list.add(time1 + "\n PSI");
			list.add(time1 + "\n 售訂");
			list.add(time1 + "\n 可再接單");
			list.add(time2 + "\n PSI");
			list.add(time2 + "\n 售訂");
			list.add(time2 + "\n 可再接單");
			list.add(time3 + "\n PSI");
			list.add(time3 + "\n 售訂");
			list.add(time3 + "\n 可再接單");
			list.add(time4 + "\n PSI");
			list.add(time4 + "\n 售訂");
			list.add(time4 + "\n 可再接單");
			list.add(time5 + "\n PSI");
			list.add(time5 + "\n 售訂");
			list.add(time5 + "\n 可再接單");
			list.add("備註");
			list.add("品名");
			list.add("規格");

			List<List<Object>> data = new ArrayList<>();

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add("集雅社");

					if (IMA.get(i).getOrDefault("pmc03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("pmc03", ""))) {
						item.add(IMA.get(i).getOrDefault("pmc03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("tqa02t2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqa02t2", ""))) {
						item.add(IMA.get(i).getOrDefault("tqa02t2", "").toString());
					} else {
						item.add("");
					}

					
					if (IMA.get(i).getOrDefault("imz02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("imz02", ""))) {
						item.add(IMA.get(i).getOrDefault("imz02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("tqa02t1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqa02t1", ""))) {
						item.add(IMA.get(i).getOrDefault("tqa02t1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima01", ""))) {
						item.add(IMA.get(i).getOrDefault("ima01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("a1", "") != null && !"".equals(IMA.get(i).getOrDefault("a1", ""))) {
						item.add(IMA.get(i).getOrDefault("a1", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString()));

					if (IMA.get(i).getOrDefault("d2", "") != null && !"".equals(IMA.get(i).getOrDefault("d2", ""))) {
						item.add(IMA.get(i).getOrDefault("d2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("o1", "") != null && !"".equals(IMA.get(i).getOrDefault("o1", ""))) {
						item.add(IMA.get(i).getOrDefault("o1", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString()));

					if (IMA.get(i).getOrDefault("d3", "") != null && !"".equals(IMA.get(i).getOrDefault("d3", ""))) {
						item.add(IMA.get(i).getOrDefault("d3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("o2", "") != null && !"".equals(IMA.get(i).getOrDefault("o2", ""))) {
						item.add(IMA.get(i).getOrDefault("o2", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString()));

					if (IMA.get(i).getOrDefault("d4", "") != null && !"".equals(IMA.get(i).getOrDefault("d4", ""))) {
						item.add(IMA.get(i).getOrDefault("d4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("o3", "") != null && !"".equals(IMA.get(i).getOrDefault("o3", ""))) {
						item.add(IMA.get(i).getOrDefault("o3", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d4", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o3", "").toString()));

					if (IMA.get(i).getOrDefault("d5", "") != null && !"".equals(IMA.get(i).getOrDefault("d5", ""))) {
						item.add(IMA.get(i).getOrDefault("d5", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("o4", "") != null && !"".equals(IMA.get(i).getOrDefault("o4", ""))) {
						item.add(IMA.get(i).getOrDefault("o4", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d4", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o3", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d5", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o4", "").toString()));

					if (IMA.get(i).getOrDefault("d6", "") != null && !"".equals(IMA.get(i).getOrDefault("d6", ""))) {
						item.add(IMA.get(i).getOrDefault("d6", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("o5", "") != null && !"".equals(IMA.get(i).getOrDefault("o5", ""))) {
						item.add(IMA.get(i).getOrDefault("o5", "").toString());
					} else {
						item.add("");
					}

					item.add(Integer.parseInt(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d4", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o3", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d5", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o4", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d6", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o5", "").toString()));

					if (IMA.get(i).getOrDefault("ps", "") != null && !"".equals(IMA.get(i).getOrDefault("ps", ""))) {
						item.add(IMA.get(i).getOrDefault("ps", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima02", ""))) {
						item.add(IMA.get(i).getOrDefault("ima02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima021", ""))) {
						item.add(IMA.get(i).getOrDefault("ima021", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			ExcelWriterBuilder excelwrite = ExcelUtil.write(fileName);

			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("priceRelations").doWrite(data);

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	//品號售訂
	public void get_ima_order_down(String fileName, List<Map<String, Object>> IMA) throws IOException {
		try {
			List<String> list = new ArrayList<String>();

			list.add("產品編號");
			list.add("銷售單位");
			list.add("廠商型號");
			list.add("規格");
			list.add("品牌代號");
			list.add("品牌");
			list.add("供應商");
			list.add("供應商名稱");
			list.add("採購大類代號");
			list.add("採購大類");
			list.add("商品大類代號");
			list.add("商品大類");
			list.add("出貨日期");
			list.add("出貨單號");
			list.add("出貨項次");
			list.add("出貨營運中心");
			list.add("出貨倉庫名稱");
			list.add("出貨倉庫");
			list.add("出貨儲位");
			list.add("客戶產品編號");
			list.add("實際出貨數量");
			list.add("單價");
			list.add("未稅金額");
			list.add("含稅金額");
			list.add("訂單單號");
			list.add("訂單項次");
			list.add("銷退數量 (需換貨再出貨)");
			list.add("銷退數量 (不需換貨出貨)");
			list.add("銷退單單號");
			list.add("原因碼");
			list.add("倉庫庫存量");
			list.add("品號總銷量");
			list.add("品號總銷量金額(含稅)");
			list.add("品號總銷量金額(未稅)");
			list.add("庫別總銷量");
			list.add("儲位總銷量");
			list.add("存貨倉備貨庫存總量");
			list.add("存貨倉陳列庫存總量");
			list.add("存貨倉庫存總量");
			list.add("承銷倉備貨庫存總量");
			list.add("承銷倉陳列庫存總量");
			list.add("承銷倉庫存總量");
			list.add("在途倉庫存總量");
			list.add("其他倉庫儲位庫存總量");
			list.add("品號庫存總量");
			list.add("稅別");
			list.add("稅率");
			list.add("含稅否");
			list.add("搭贈");
			list.add("預計出貨日期");
			list.add("所屬營運中心");
			list.add("營運中心名稱");
			list.add("所屬法人");
			list.add("專櫃編號");
			list.add("已簽退數量");
			list.add("簽退數量");
			list.add("備註");
			list.add("會員卡號");
			list.add("顧客姓名");
			list.add("聯系電話");
			list.add("配送方式");
			list.add("收件人");
			list.add("收件人電話");
			list.add("郵遞區號");
			list.add("送貨地址");

			List<List<Object>> data = new ArrayList<>();

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					if (IMA.get(i).getOrDefault("OGB04", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB04", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB04", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB05", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB05", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB05", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA1003", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA1003", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA1003", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA1005", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA1005", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA1005", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TQA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TQA02", ""))) {
						item.add(IMA.get(i).getOrDefault("TQA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA54", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA54", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA54", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("PMC03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("PMC03", ""))) {
						item.add(IMA.get(i).getOrDefault("PMC03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA06", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA06", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA06", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMZ02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMZ02", ""))) {
						item.add(IMA.get(i).getOrDefault("IMZ02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA131", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA131", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA131", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OBA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OBA02", ""))) {
						item.add(IMA.get(i).getOrDefault("OBA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGA02", ""))) {
						item.add(IMA.get(i).getOrDefault("OGA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB01", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB03", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB08", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMD02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMD02", ""))) {
						item.add(IMA.get(i).getOrDefault("IMD02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB09", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB091", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB091", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB091", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB11", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB11", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB11", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB12", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB12", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB12", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB13", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB13", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB13", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB14", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB14", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB14", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB14T", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB14T", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB14T", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB31", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB31", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB31", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB32", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB32", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB32", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB63", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB63", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB63", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB64", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB64", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB64", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGA1012", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGA1012", ""))) {
						item.add(IMA.get(i).getOrDefault("OGA1012", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1001", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1001", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1001", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOGB12", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOGB12", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOGB12", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOGB13", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOGB13", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOGB13", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOGB14", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOGB14", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOGB14", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOGB12OGB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOGB12OGB09", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOGB12OGB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOGB12OGB91", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOGB12OGB91", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOGB12OGB91", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1008", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1008", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1008", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1009", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1009", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1009", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1010", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1010", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1010", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1012", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1012", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1012", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB1003", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB1003", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB1003", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGBPLANT", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGBPLANT", ""))) {
						item.add(IMA.get(i).getOrDefault("OGBPLANT", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("AZW08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("AZW08", ""))) {
						item.add(IMA.get(i).getOrDefault("AZW08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGBLEGAL", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGBLEGAL", ""))) {
						item.add(IMA.get(i).getOrDefault("OGBLEGAL", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB48", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB48", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB48", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB51", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB51", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB51", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGB52", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGB52", ""))) {
						item.add(IMA.get(i).getOrDefault("OGB52", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGB01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGB01", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGB01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGA87", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGA87", ""))) {
						item.add(IMA.get(i).getOrDefault("OGA87", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGA88", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGA88", ""))) {
						item.add(IMA.get(i).getOrDefault("OGA88", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OGA89", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OGA89", ""))) {
						item.add(IMA.get(i).getOrDefault("OGA89", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGA01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGA01", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGA01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGA02", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGA03", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGA03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGA04", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGA04", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGA04", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("TA_OGA05", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("TA_OGA05", ""))) {
						item.add(IMA.get(i).getOrDefault("TA_OGA05", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			ExcelWriterBuilder excelwrite = ExcelUtil.write(fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("priceRelations").doWrite(data);

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void get_pre_sale_down(String fileName, List<Map<String, Object>> IMA) throws IOException {
		try {
			List<String> list = new ArrayList<String>();

			list.add("產品編號");
			list.add("銷售單位");
			list.add("廠商型號");
			list.add("規格");
			list.add("品牌代號");
			list.add("品牌");
			list.add("採購大類代號");
			list.add("採購大類");
			list.add("商品大類代號");
			list.add("商品大類");
			list.add("訂單單號");
			list.add("項次");
			list.add("出貨營運中心");
			list.add("出貨倉庫名稱");
			list.add("出貨倉庫");
			list.add("出貨儲位");
			list.add("受訂數量");
			list.add("單價");
			list.add("未稅金額");
			list.add("含稅金額");
			list.add("訂單日期");
			list.add("約定交貨日");
			list.add("排定交貨日");
			list.add("備置否");
			list.add("待出貨量");
			list.add("已出貨量");
			list.add("已消退貨量");
			list.add("被結案貨量");
			list.add("結案否");
			list.add("結案日期");
			list.add("己備置量");
			list.add("倉庫庫存量");
			list.add("品號總受訂量");
			list.add("庫別總受訂量");
			list.add("儲位總受訂量");
			list.add("存貨倉備貨庫存總量");
			list.add("存貨倉陳列庫存總量");
			list.add("存貨倉庫存總量");
			list.add("承銷倉備貨庫存總量");
			list.add("承銷倉陳列庫存總量");
			list.add("承銷倉庫存總量");
			list.add("在途倉庫存總量");
			list.add("其他倉庫儲位庫存總量");
			list.add("品號總量");
			list.add("已分配量");
			list.add("請購單號");
			list.add("已轉請購量");
			list.add("所屬營運中心");
			list.add("營運中心名稱");
			list.add("所屬法人");
			list.add("供應商");
			list.add("供應商名稱");

			List<List<Object>> data = new ArrayList<>();

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					if (IMA.get(i).getOrDefault("OEB04", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB04", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB04", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB05", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB05", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB05", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima1003", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima1003", ""))) {
						item.add(IMA.get(i).getOrDefault("ima1003", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima1005", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima1005", ""))) {
						item.add(IMA.get(i).getOrDefault("ima1005", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("tqa02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqa02", ""))) {
						item.add(IMA.get(i).getOrDefault("tqa02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima06", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima06", ""))) {
						item.add(IMA.get(i).getOrDefault("ima06", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("imz02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("imz02", ""))) {
						item.add(IMA.get(i).getOrDefault("imz02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima131", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima131", ""))) {
						item.add(IMA.get(i).getOrDefault("ima131", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB01", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB03", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB08", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("imd02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("imd02", ""))) {
						item.add(IMA.get(i).getOrDefault("imd02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB09", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB091", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB091", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB091", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB12", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB12", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB12", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB13", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB13", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB13", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB14", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB14", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB14", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB14T", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB14T", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB14T", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEA02", ""))) {
						item.add(IMA.get(i).getOrDefault("OEA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB15", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB15", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB15", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB16", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB16", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB16", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB19", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB19", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB19", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB23", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB23", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB23", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB24", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB24", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB24", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB25", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB25", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB25", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB26", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB26", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB26", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB70", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB70", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB70", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB70D", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB70D", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB70D", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB905", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB905", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB905", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12OEB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12OEB09", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12OEB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12OEB91", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12OEB91", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12OEB91", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB920", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB920", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB920", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB27", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB27", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB27", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB28", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB28", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB28", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEBPLANT", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEBPLANT", ""))) {
						item.add(IMA.get(i).getOrDefault("OEBPLANT", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("AZW08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("AZW08", ""))) {
						item.add(IMA.get(i).getOrDefault("AZW08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEBLEGAL", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEBLEGAL", ""))) {
						item.add(IMA.get(i).getOrDefault("OEBLEGAL", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima54", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima54", ""))) {
						item.add(IMA.get(i).getOrDefault("ima54", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("pmc03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("pmc03", ""))) {
						item.add(IMA.get(i).getOrDefault("pmc03", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			ExcelWriterBuilder excelwrite = ExcelUtil.write(fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("priceRelations").doWrite(data);

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void get_stock_down1(String fileName, List<Map<String, Object>> IMA) throws IOException {
		try {
			List<String> list = new ArrayList<String>();

			list.add("公司別");
			list.add("營運中心名稱");
			list.add("所屬營運中心");
			list.add("料件編號");
			list.add("品名");
			list.add("規格");
			list.add("廠商型號");
			list.add("供應商代碼");
			list.add("供應商");
			list.add("品牌代號");
			list.add("品牌代碼");
			list.add("產品狀態");
			list.add("產品狀態代碼");
			list.add("採購大類代號");
			list.add("採購大類");
			list.add("商品大類代號");
			list.add("商品大類");
			list.add("倉庫編號");
			list.add("倉庫名稱");
			list.add("儲位");
			list.add("庫存單位");
			list.add("儲位庫存數量");
			list.add("品號總受訂量");
			list.add("庫別總受訂量");
			list.add("儲位總受訂量");
			list.add("存貨倉備貨庫存總量");
			list.add("存貨倉陳列庫存總量");
			list.add("存貨倉庫存總量");
			list.add("承銷倉備貨庫存總量");
			list.add("承銷倉陳列庫存總量");
			list.add("承銷倉庫存總量");
			list.add("在途倉庫存總量");
			list.add("其他倉庫儲位庫存總量");
			list.add("品號總量");

			List<List<Object>> data = new ArrayList<>();

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					if (IMA.get(i).getOrDefault("IMGLEGAL", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMGLEGAL", ""))) {
						item.add(IMA.get(i).getOrDefault("IMGLEGAL", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("AZW08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("AZW08", ""))) {
						item.add(IMA.get(i).getOrDefault("AZW08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMGPLANT", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMGPLANT", ""))) {
						item.add(IMA.get(i).getOrDefault("IMGPLANT", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG01", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG01", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG01", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA02", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA02", "").toString());
					} else {
						item.add("");
					}
					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA1003", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA1003", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA1003", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA54", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA54", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA54", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("pmc03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("pmc03", ""))) {
						item.add(IMA.get(i).getOrDefault("pmc03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima1005", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima1005", ""))) {
						item.add(IMA.get(i).getOrDefault("ima1005", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("tqab", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqab", ""))) {
						item.add(IMA.get(i).getOrDefault("tqab", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima1004", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima1004", ""))) {
						item.add(IMA.get(i).getOrDefault("ima1004", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("tqaa", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqaa", ""))) {
						item.add(IMA.get(i).getOrDefault("tqaa", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima06", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima06", ""))) {
						item.add(IMA.get(i).getOrDefault("ima06", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("imz02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("imz02", ""))) {
						item.add(IMA.get(i).getOrDefault("imz02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ima131", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima131", ""))) {
						item.add(IMA.get(i).getOrDefault("ima131", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OBA02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OBA02", ""))) {
						item.add(IMA.get(i).getOrDefault("OBA02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG02", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMD02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMD02", ""))) {
						item.add(IMA.get(i).getOrDefault("IMD02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG09", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12OEB24OEB04", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB04", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB04", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12OEB24OEB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB09", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMOEB12OEB24OEB091", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB091", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMOEB12OEB24OEB091", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_1", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG03_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_2", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10IMG02_4", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("SUMIMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("SUMIMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("SUMIMG10", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			ExcelWriterBuilder excelwrite = ExcelUtil.write(fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("priceRelations").doWrite(data);

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void get_all_3(String DB, JdbcTemplate jdbcTemplate1, HttpServletResponse response1) throws IOException {

		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_all_3(DB, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("銷售單位");
			item1.add("品名規格");
			item1.add("廠商型號");
			item1.add("品牌代碼");
			item1.add("品牌");
			item1.add("採購大類代碼");

			item1.add("採購大類");
			item1.add("商品大類代碼");
			item1.add("商品大類");

			item1.add("訂單單號");
			item1.add("項次");
			item1.add("出貨營運中心");
			item1.add("出貨倉庫名稱");
			item1.add("出貨倉庫");
			item1.add("出貨儲位");
			item1.add("受訂數量");
			item1.add("單價");
			item1.add("未稅金額");
			item1.add("含稅金額");
			item1.add("約定交貨日");
			item1.add("排定交貨日");
			item1.add("備置否");
			item1.add("待出貨數量");
			item1.add("已出貨數量");
			item1.add("已銷退數量");
			item1.add("被結案數量");
			item1.add("結案否");
			item1.add("結案日期");
			item1.add("己備置量");
			item1.add("倉庫庫存量");
			item1.add("庫別總受訂量");
			item1.add("庫別可用總量");
			item1.add("可用總量");
			item1.add("已分配量");
			item1.add("請購單號");
			item1.add("已轉請購量");
			item1.add("所屬營運中心");
			item1.add("營運中心名稱");
			item1.add("所屬法人");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("OEB04", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB05", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB06", "").toString());
					item.add(IMA.get(i).getOrDefault("ima1003", "").toString());
					item.add(IMA.get(i).getOrDefault("ima1005", "").toString());

					if (IMA.get(i).getOrDefault("tqa02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("tqa02", ""))) {
						item.add(IMA.get(i).getOrDefault("tqa02", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("ima06", "").toString());
					item.add(IMA.get(i).getOrDefault("imz02", "").toString());
					item.add(IMA.get(i).getOrDefault("ima131", "").toString());

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("OEB01", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB03", "").toString());

					if (IMA.get(i).getOrDefault("OEB08", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB08", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB08", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("imd02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("imd02", ""))) {
						item.add(IMA.get(i).getOrDefault("imd02", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB09", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB09", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB09", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("OEB091", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB091", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB091", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("OEB12", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB13", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB14", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB14T", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB15", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB16", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB19", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB23", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB24", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB25", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB26", "").toString());
					item.add(IMA.get(i).getOrDefault("OEB70", "").toString());

					if (IMA.get(i).getOrDefault("OEB70D", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB70D", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB70D", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("OEB905", "").toString());
					item.add(IMA.get(i).getOrDefault("a1", "").toString());
					item.add(IMA.get(i).getOrDefault("a2", "").toString());
					item.add(IMA.get(i).getOrDefault("a3", "").toString());
					item.add(IMA.get(i).getOrDefault("a4", "").toString());

					if (IMA.get(i).getOrDefault("OEB920", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB920", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB920", "").toString());
					} else {
						item.add("");
					}
					if (IMA.get(i).getOrDefault("OEB27", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("OEB27", ""))) {
						item.add(IMA.get(i).getOrDefault("OEB27", "").toString());
					} else {
						item.add("");
					}
					item.add(IMA.get(i).getOrDefault("OEB28", "").toString());
					item.add(IMA.get(i).getOrDefault("OEBPLANT", "").toString());
					item.add(IMA.get(i).getOrDefault("azw08", "").toString());
					item.add(IMA.get(i).getOrDefault("OEBLEGAL", "").toString());

					/*
					 * if(IMA.get(i).getOrDefault("a6","")!=null &&
					 * !"".equals(IMA.get(i).getOrDefault("a6",""))) {
					 * item.add(IMA.get(i).getOrDefault("a6","").toString()); }else { item.add("");
					 * }
					 */

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "受訂量.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	public void get_all_2(String DB, JdbcTemplate jdbcTemplate1, HttpServletResponse response1) throws IOException {

		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_all_2(DB, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");
			item1.add("品牌代碼");
			item1.add("品牌");
			item1.add("採購大類代碼");

			item1.add("採購大類");
			item1.add("商品大類代碼");
			item1.add("商品大類");

			item1.add("庫存數量");
			item1.add("承銷數量");
			item1.add("承列和備貨總數量");
			item1.add("在途總數量");
			item1.add("標準進價_未稅");
			item1.add("庫存金額");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("IMA01", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA021", "").toString());

					if (IMA.get(i).getOrDefault("IMA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("ima1005", "").toString());
					item.add(IMA.get(i).getOrDefault("tqa02", "").toString());
					item.add(IMA.get(i).getOrDefault("ima06", "").toString());
					item.add(IMA.get(i).getOrDefault("imz02", "").toString());

					if (IMA.get(i).getOrDefault("ima131", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima131", ""))) {
						item.add(IMA.get(i).getOrDefault("ima131", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("a1", "").toString());
					item.add(IMA.get(i).getOrDefault("a2", "").toString());
					item.add(IMA.get(i).getOrDefault("a3", "").toString());
					item.add(IMA.get(i).getOrDefault("a4", "").toString());
					item.add(IMA.get(i).getOrDefault("a5", "").toString());

					if (IMA.get(i).getOrDefault("a6", "") != null && !"".equals(IMA.get(i).getOrDefault("a6", ""))) {
						item.add(IMA.get(i).getOrDefault("a6", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "庫存量.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	public void get_all_1(String DB, HttpServletRequest request, JdbcTemplate jdbcTemplate1,
			HttpServletResponse response1) throws IOException {

		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_all_1(DB, request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");
			item1.add("品牌代碼");
			item1.add("品牌");
			item1.add("採購大類代碼");

			item1.add("採購大類");
			item1.add("商品大類代碼");
			item1.add("商品大類");

			item1.add("銷貨數量");
			item1.add("銷貨_原幣未稅金額");
			item1.add("銷退數量");
			item1.add("銷退_原幣未稅金額");
			item1.add("實銷數量");
			item1.add("銷售毛利");
			item1.add("平均單價");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("IMA01", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA021", "").toString());

					if (IMA.get(i).getOrDefault("IMA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("ima1005", "").toString());
					item.add(IMA.get(i).getOrDefault("tqa02", "").toString());
					item.add(IMA.get(i).getOrDefault("ima06", "").toString());
					item.add(IMA.get(i).getOrDefault("imz02", "").toString());

					if (IMA.get(i).getOrDefault("ima131", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ima131", ""))) {
						item.add(IMA.get(i).getOrDefault("ima131", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("a1", "").toString());
					item.add(IMA.get(i).getOrDefault("a2", "").toString());
					item.add(IMA.get(i).getOrDefault("a3", "").toString());
					item.add(IMA.get(i).getOrDefault("a4", "").toString());
					item.add(IMA.get(i).getOrDefault("a5", "").toString());

					if (IMA.get(i).getOrDefault("a6", "") != null && !"".equals(IMA.get(i).getOrDefault("a6", ""))) {
						item.add(IMA.get(i).getOrDefault("a6", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("a7", "") != null && !"".equals(IMA.get(i).getOrDefault("a7", ""))) {
						item.add(IMA.get(i).getOrDefault("a7", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-銷售量.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	public void imd_xlsx3(String DB, JdbcTemplate jdbcTemplate1, HttpServletResponse response1) throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_imd_006(DB, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");
			item1.add("倉庫編號");
			item1.add("倉庫名稱");
			item1.add("儲位");

			item1.add("庫存單位");
			item1.add("庫存數量");
			item1.add("有效日期");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("IMG01", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA02", "").toString());

					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMD02", "").toString());

					if (IMA.get(i).getOrDefault("IMG03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG09", "").toString());

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG18", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG18", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG18", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-倉庫品號.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void imd_xlsx4(String DB, JdbcTemplate jdbcTemplate1, HttpServletResponse response1) throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_oeb_006(DB, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");
			item1.add("倉庫編號");
			item1.add("倉庫名稱");
			item1.add("儲位");

			item1.add("庫存單位");
			item1.add("庫存數量");
			item1.add("有效日期");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("IMG01", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA02", "").toString());

					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMD02", "").toString());

					if (IMA.get(i).getOrDefault("IMG03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG09", "").toString());

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG18", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG18", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG18", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-倉庫品號.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void pmc_xlsx(String DB, String pmc, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_pmc(DB, pmc, jdbcTemplate1);

//				IMG01,ima03,ima54,sum(img10)
			ArrayList item1 = new ArrayList();
			item1.add("品號");
			item1.add("舊品號");
			item1.add("規格");
			item1.add("採購大類代碼");
			item1.add("採購大類");
			item1.add("庫存");
			item1.add("存貨庫(陳列&備貨)");
			item1.add("陳列");
			item1.add("備貨");
			item1.add("承銷庫");
			item1.add("在途庫");
			item1.add("其他庫(暫存)");
			item1.add("已預訂量");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();
//ima06,imz02,
					item.add(IMA.get(i).getOrDefault("IMA01", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));
					item.add(IMA.get(i).getOrDefault("ima06", ""));
					item.add(IMA.get(i).getOrDefault("imz02", ""));
					item.add(IMA.get(i).getOrDefault("c1", ""));
					item.add(IMA.get(i).getOrDefault("c2", ""));
					item.add(IMA.get(i).getOrDefault("c2_1", ""));
					item.add(IMA.get(i).getOrDefault("c2_2", ""));
					item.add(IMA.get(i).getOrDefault("c3", ""));
					item.add(IMA.get(i).getOrDefault("c4", ""));
					item.add(IMA.get(i).getOrDefault("c6", ""));
					item.add(IMA.get(i).getOrDefault("c5", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-倉庫品號.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void imd_xlsx(String DB, String imd, String c1, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_imd(DB, imd, c1, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("料件編號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");

			item1.add("主供應商代碼");
			item1.add("主供應商");
			item1.add("本幣未稅單價");
			item1.add("標準進價(未稅)");

			item1.add("採購大類");
			item1.add("採購大類名稱");
			item1.add("產品大類");
			item1.add("產品大類名稱");
			item1.add("品牌");
			item1.add("品牌名稱");

			item1.add("產品狀態");
			item1.add("產品狀態名稱");

			item1.add("倉庫編號");
			item1.add("倉庫名稱");
			item1.add("儲位");

			item1.add("庫存單位");
			item1.add("庫存數量");
			item1.add("有效日期");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("IMG01", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA02", "").toString());

					if (IMA.get(i).getOrDefault("IMA021", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA021", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA021", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA54", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA54", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA54", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("PMC03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("PMC03", ""))) {
						item.add(IMA.get(i).getOrDefault("PMC03", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA127", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA127", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA127", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMA531", "").toString());

					item.add(IMA.get(i).getOrDefault("IMA06", "").toString());
					item.add(IMA.get(i).getOrDefault("imz02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMA131", "").toString());

					if (IMA.get(i).getOrDefault("oba02", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("oba02", ""))) {
						item.add(IMA.get(i).getOrDefault("oba02", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMA1005", "").toString());

					if (IMA.get(i).getOrDefault("s3", "") != null && !"".equals(IMA.get(i).getOrDefault("s3", ""))) {
						item.add(IMA.get(i).getOrDefault("s3", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMA1004", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMA1004", ""))) {
						item.add(IMA.get(i).getOrDefault("IMA1004", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("s1", "") != null && !"".equals(IMA.get(i).getOrDefault("s1", ""))) {
						item.add(IMA.get(i).getOrDefault("s1", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG02", "").toString());
					item.add(IMA.get(i).getOrDefault("IMD02", "").toString());

					if (IMA.get(i).getOrDefault("IMG03", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG03", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG03", "").toString());
					} else {
						item.add("");
					}

					item.add(IMA.get(i).getOrDefault("IMG09", "").toString());

					if (IMA.get(i).getOrDefault("IMG10", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG10", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG10", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("IMG18", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("IMG18", ""))) {
						item.add(IMA.get(i).getOrDefault("IMG18", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-倉庫品號.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void writeExcel_valet(HttpServletRequest request, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new purchaseFactory()).selectIMA(request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("企業別");
			item1.add("供應商");
			item1.add("採購大類");
			item1.add("品號分類3");
			item1.add("品牌");
			item1.add("品號");
			item1.add("庫存數量");
			item1.add("備貨數量");
			item1.add("陳列數量");
			item1.add("承銷數量");
			item1.add("總受訂量");
			item1.add("售訂未交");
			item1.add("可再接單");
			item1.add("PSI");
			item1.add("售訂");
			item1.add("可再接單");

			item1.add("PSI");
			item1.add("售訂");
			item1.add("可再接單");
			item1.add("PSI");
			item1.add("售訂");
			item1.add("可再接單");
			item1.add("PSI");
			item1.add("售訂");
			item1.add("可再接單");
			item1.add("PSI");
			item1.add("售訂");
			item1.add("可再接單");

			item1.add("備註");
			item1.add("品名");
			item1.add("規格");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add("集雅社");
					item.add(IMA.get(i).getOrDefault("pmc03", ""));
					item.add(IMA.get(i).getOrDefault("imz02", ""));
					item.add("");
					item.add(IMA.get(i).getOrDefault("TQA_name", ""));
					item.add(IMA.get(i).getOrDefault("ima01", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("a2", "").toString()));
					item.add(IMA.get(i).getOrDefault("a1", ""));
					item.add(IMA.get(i).getOrDefault("a2", ""));
					item.add(IMA.get(i).getOrDefault("a3", ""));
					item.add(IMA.get(i).getOrDefault("a6", ""));
					item.add(IMA.get(i).getOrDefault("a5", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString()));

					item.add(IMA.get(i).getOrDefault("d2", ""));
					item.add(IMA.get(i).getOrDefault("o1", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString()));

					item.add(IMA.get(i).getOrDefault("d3", ""));
					item.add(IMA.get(i).getOrDefault("o2", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString()));

					item.add(IMA.get(i).getOrDefault("d4", ""));
					item.add(IMA.get(i).getOrDefault("o3", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d4", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o3", "").toString()));

					item.add(IMA.get(i).getOrDefault("d5", ""));
					item.add(IMA.get(i).getOrDefault("o4", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d5", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o4", "").toString()));

					item.add(IMA.get(i).getOrDefault("d6", ""));
					item.add(IMA.get(i).getOrDefault("o5", ""));
					item.add(Integer.parseInt(IMA.get(i).getOrDefault("a1", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("a5", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d2", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o1", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d3", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o2", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d5", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o4", "").toString())
							+ Integer.parseInt(IMA.get(i).getOrDefault("d6", "").toString())
							- Integer.parseInt(IMA.get(i).getOrDefault("o5", "").toString()));

					item.add(IMA.get(i).getOrDefault("ps", ""));
					item.add(IMA.get(i).getOrDefault("ima02", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + ".xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void sony_xlsx_detal(HttpServletRequest request, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony_oga_detal(request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("出貨單");
			item1.add("品號");
			item1.add("舊品號");
			item1.add("櫃位");
			item1.add("日期");
			item1.add("客戶");
			item1.add("品牌");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("ogb01", ""));
					item.add(IMA.get(i).getOrDefault("ogb04", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("ogb08", ""));
					item.add(IMA.get(i).getOrDefault("oga02", "").toString());
					item.add(IMA.get(i).getOrDefault("oga45", ""));
					item.add(IMA.get(i).getOrDefault("TQA02", ""));
					/*
					 * item.add(IMA.get(i).getOrDefault("ogb04",""));
					 * item.add(IMA.get(i).getOrDefault("ima03",""));
					 * item.add(IMA.get(i).getOrDefault("ogb08",""));
					 * item.add(IMA.get(i).getOrDefault("oga02",""));
					 * item.add(IMA.get(i).getOrDefault("oga45",""));
					 * item.add(IMA.get(i).getOrDefault("TQA02",""));
					 */
					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-sony銷售數量明細驗證.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

//庫存
	public void sony_xlsx2(String DB, HttpServletRequest request, JdbcTemplate jdbcTemplate1,
			HttpServletResponse response1) throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony_img(DB, jdbcTemplate1);

//				IMG01,ima03,ima54,sum(img10)
			ArrayList item1 = new ArrayList();
			item1.add("品號");
			item1.add("舊品號");
			item1.add("規格");
			item1.add("採購大類代碼");
			item1.add("採購大類");
			item1.add("庫存");
			item1.add("存貨庫S(陳列&備貨)");
			item1.add("陳列");
			item1.add("備貨");
			item1.add("承銷庫C");
			item1.add("在途庫W");
			item1.add("其他庫(暫存&待處理)");
			item1.add("已預訂量");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();
//ima06,imz02,
					item.add(IMA.get(i).getOrDefault("IMA01", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));
					item.add(IMA.get(i).getOrDefault("ima06", ""));
					item.add(IMA.get(i).getOrDefault("imz02", ""));
					item.add(IMA.get(i).getOrDefault("c1", ""));
					item.add(IMA.get(i).getOrDefault("c2", ""));
					item.add(IMA.get(i).getOrDefault("c2_1", ""));
					item.add(IMA.get(i).getOrDefault("c2_2", ""));
					item.add(IMA.get(i).getOrDefault("c3", ""));
					item.add(IMA.get(i).getOrDefault("c4", ""));
					item.add(IMA.get(i).getOrDefault("c6", ""));
					item.add(IMA.get(i).getOrDefault("c5", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-集雅社sony庫存.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	// 庫存
	public void sony_xlsx21(String DB, HttpServletRequest request, JdbcTemplate jdbcTemplate1,
			HttpServletResponse response1) throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony_img21(DB, jdbcTemplate1);

//						IMG01,ima03,ima54,sum(img10)
			ArrayList item1 = new ArrayList();
			item1.add("品號");
			item1.add("舊品號");
			item1.add("規格");
			item1.add("採購大類代碼");
			item1.add("採購大類");
			item1.add("庫存");
			item1.add("存貨庫(陳列&備貨)");
			item1.add("陳列");
			item1.add("備貨");
			item1.add("承銷庫");
			item1.add("在途庫");
			item1.add("其他庫(暫存)");
			item1.add("已預訂量");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();
					// ima06,imz02,
					item.add(IMA.get(i).getOrDefault("IMA01", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));
					item.add(IMA.get(i).getOrDefault("ima06", ""));
					item.add(IMA.get(i).getOrDefault("imz02", ""));
					item.add(IMA.get(i).getOrDefault("c1", ""));
					item.add(IMA.get(i).getOrDefault("c2", ""));
					item.add(IMA.get(i).getOrDefault("c2_1", ""));
					item.add(IMA.get(i).getOrDefault("c2_2", ""));
					item.add(IMA.get(i).getOrDefault("c3", ""));
					item.add(IMA.get(i).getOrDefault("c4", ""));
					item.add(IMA.get(i).getOrDefault("c6", ""));
					item.add(IMA.get(i).getOrDefault("c5", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-集雅社sony庫存.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void sony_xlsx4(String DB, HttpServletRequest request, JdbcTemplate jdbcTemplate1,
			HttpServletResponse response1) throws IOException {

		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony_rxy(DB, request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("櫃位");
			item1.add("櫃位名稱");
			item1.add("訂單單號");

			item1.add("收款日期");
			item1.add("項次");
			item1.add("品號");
			item1.add("舊品號");
			item1.add("數量");
			item1.add("含稅金額");
			item1.add("業務員");
			item1.add("品牌代碼");
			item1.add("品牌");
			item1.add("規格");
			item1.add("供應商代碼");
			item1.add("供應商");
			item1.add("產品大類代碼");
			item1.add("產品大類");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("OEA83", ""));
					item.add(IMA.get(i).getOrDefault("gem02", ""));

					item.add(IMA.get(i).getOrDefault("OEA01", ""));

					item.add(IMA.get(i).getOrDefault("oea72", "").toString());

					item.add(IMA.get(i).getOrDefault("OEB03", ""));
					item.add(IMA.get(i).getOrDefault("OEB04", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("oeb12", ""));

					item.add(IMA.get(i).getOrDefault("OEB14t", ""));
					item.add(IMA.get(i).getOrDefault("OEA14", ""));
					item.add(IMA.get(i).getOrDefault("ima1005", ""));

					item.add(IMA.get(i).getOrDefault("tqa02", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));
					item.add(IMA.get(i).getOrDefault("ima54", ""));
					item.add(IMA.get(i).getOrDefault("pmc03", ""));

					item.add(IMA.get(i).getOrDefault("ima131", ""));
					item.add(IMA.get(i).getOrDefault("oba02", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-集雅社sony收款.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void sony_xlsx(String DB, HttpServletRequest request, JdbcTemplate jdbcTemplate1,
			HttpServletResponse response1) throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony_oga(DB, request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("品號");
			item1.add("品名");
			item1.add("規格");
			item1.add("舊品號");
			item1.add("實銷數量");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("OGB04", ""));
					item.add(IMA.get(i).getOrDefault("tqa02", ""));
					item.add(IMA.get(i).getOrDefault("ima021", ""));
					item.add(IMA.get(i).getOrDefault("ima03", ""));
					item.add(IMA.get(i).getOrDefault("a1", ""));

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-sony銷售數量.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void sony_xlsx3(HttpServletRequest request, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_sony3(request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("品號");
			item1.add("舊品號");
			item1.add("預訂交貨日");
			item1.add("數量");
			item1.add("訂單單號");
			item1.add("規格");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("oeb04", "").toString());
					item.add(IMA.get(i).getOrDefault("ima03", "").toString());
					item.add(IMA.get(i).getOrDefault("oeb16", "").toString());
					item.add(IMA.get(i).getOrDefault("oeb12", "").toString());
					item.add(IMA.get(i).getOrDefault("oeb01", "").toString());
					item.add(IMA.get(i).getOrDefault("ima021", "").toString());

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-sony受定量主檔.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public void chimei_xlsx(HttpServletRequest request, JdbcTemplate jdbcTemplate1, HttpServletResponse response1)
			throws IOException {
		try {
			Map<Integer, Integer> maxColumnWidthMap = null;
			List<String> list = new ArrayList<String>();

			List<List<Object>> data = new ArrayList<>();

			List<Map<String, Object>> IMA = (new reportFactory()).get_chimei(request, jdbcTemplate1);

			ArrayList item1 = new ArrayList();
			item1.add("請購單");
			item1.add("請購日期");
			item1.add("舊品號");
			item1.add("數量");
			item1.add("項次");
			item1.add("品號");
			item1.add("預訂交貨日");
			item1.add("規格");
			item1.add("分配量");
			item1.add("收件人");
			item1.add("電話");
			item1.add("地址");

			data.add(item1);

			if (IMA != null && IMA.size() != 0)
				for (int i = 0; i < IMA.size(); i++) {
					List<Object> item = new ArrayList<>();

					item.add(IMA.get(i).getOrDefault("pml01", "").toString());
					item.add(IMA.get(i).getOrDefault("pmk04", "").toString());
					item.add(IMA.get(i).getOrDefault("ima03", "").toString());
					item.add(IMA.get(i).getOrDefault("pml20", "").toString());
					item.add(IMA.get(i).getOrDefault("pml02", "").toString());
					item.add(IMA.get(i).getOrDefault("pml04", "").toString());
					item.add(IMA.get(i).getOrDefault("pml33", "").toString());
					item.add(IMA.get(i).getOrDefault("ima021", "").toString());
					item.add(IMA.get(i).getOrDefault("pml21", "").toString());
					if (IMA.get(i).getOrDefault("ta_oeb04", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ta_oeb04", ""))) {
						item.add(IMA.get(i).getOrDefault("ta_oeb04", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ta_oeb05", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ta_oeb05", ""))) {
						item.add(IMA.get(i).getOrDefault("ta_oeb05", "").toString());
					} else {
						item.add("");
					}

					if (IMA.get(i).getOrDefault("ta_oeb07", "") != null
							&& !"".equals(IMA.get(i).getOrDefault("ta_oeb07", ""))) {
						item.add(IMA.get(i).getOrDefault("ta_oeb07", "").toString());
					} else {
						item.add("");
					}

					data.add(item);
				}

			Date date = new Date();
			SimpleDateFormat bartDateFormat = new SimpleDateFormat("yyyyMMdd");
			String str = bartDateFormat.format(date);

			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			String fileName = str + "-chimei請購單.xlsx";

			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			ExcelWriterBuilder excelwrite = ExcelUtil.write(filePath + fileName);
			excelwrite.head(ExcelUtil.createListStringHead(list));
			excelwrite.sheet("test1").doWrite(data);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName(fileName);
			download.downloadFile(response1);

			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
}
