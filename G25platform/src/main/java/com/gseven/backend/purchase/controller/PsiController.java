/**
 * 
 */
package com.gseven.backend.purchase.controller;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.gseven.backend.purchase.factory.purchaseFactory;
import com.gseven.backend.purchase.factory.reportFactory;
import com.gseven.backend.sys.entity.Department;
import com.gseven.backend.sys.entity.Employee;
import com.gseven.backend.sys.factory.DepartmentFactory;
import com.gseven.backend.sys.service.DateUtil;
import com.gseven.backend.sys.service.DownloadUtility;
import com.gseven.backend.sys.service.MyGrantedAuthority;
import com.gseven.backend.purchase.service.ExcelWrite;

/**
 * @author g7user
 *
 */

@Controller
@RequestMapping("/purchase/psi")
public class PsiController {

	@Autowired
	@Qualifier("secondaryJdbcTemplate")
	protected JdbcTemplate jdbcTemplate2; // HRMDB

	@Autowired
	@Qualifier("primaryJdbcTemplate")
	protected JdbcTemplate jdbcTemplate1;

	@RequestMapping("psi")
	public String psi(HttpServletRequest request, Model model) {
		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, 0);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String time = format.format(c.getTime());
		model.addAttribute("month_0", time);

		for (int i = 1; i <= 5; i++) {
			c = Calendar.getInstance();
			c.add(Calendar.MONTH, i);
			format = new SimpleDateFormat("yyyy-MM");
			time = format.format(c.getTime());

			model.addAttribute("month_" + i, time);
		}

		// 取品號
		List<Map<String, Object>> IMA = new ArrayList<Map<String, Object>>();

		if (request.getParameter("find_tqa") != null || request.getParameter("find_IMA01") != null
				|| request.getParameter("order_is") != null || request.getParameter("can_get_order") != null) {
			IMA = (new purchaseFactory()).selectIMA(request, jdbcTemplate1);

		}

		model.addAttribute("IMA", IMA);

		// 取品牌
		List<Map<String, Object>> tqas = (new purchaseFactory()).selectTQA(jdbcTemplate1);
		model.addAttribute("tqas", tqas);

		// 供應商
		List<Map<String, Object>> pmcblist = (new reportFactory()).selectPMC(jdbcTemplate1);
		model.addAttribute("pmclist", pmcblist);

		// 產品狀態
		List<Map<String, Object>> tqalist = (new reportFactory()).selectTQA(jdbcTemplate1);
		model.addAttribute("tqalist", tqalist);
		// 採購大類
		List<Map<String, Object>> imzlist = (new reportFactory()).selectIMZ(jdbcTemplate1);
		model.addAttribute("imzlist", imzlist);
		// 商品大類
		List<Map<String, Object>> obalist = (new reportFactory()).selectOBA(jdbcTemplate1);
		model.addAttribute("obalist", obalist);

		String find_tqa = request.getParameter("find_tqa");
		model.addAttribute("find_tqa", find_tqa);

		String find_IMA01 = request.getParameter("find_IMA01");
		model.addAttribute("find_IMA01", find_IMA01);

		String order_is = request.getParameter("order_is");
		model.addAttribute("order_is", order_is);

		String can_get_order = request.getParameter("can_get_order");
		model.addAttribute("can_get_order", can_get_order);

		String find_pmc = request.getParameter("find_pmc");

		model.addAttribute("find_pmc", find_pmc);
		String find_ima1004 = request.getParameter("find_ima1004");
		model.addAttribute("find_ima1004", find_ima1004);
		String find_IMA06 = request.getParameter("find_IMA06");
		model.addAttribute("find_IMA06", find_IMA06);
		String find_IMA131 = request.getParameter("find_IMA131");
		model.addAttribute("find_IMA131", find_IMA131);

//取登入者部門
		String loginId = "";
		Authentication auth = SecurityContextHolder.getContext().getAuthentication();
		for (GrantedAuthority ga : auth.getAuthorities()) {
			if (ga instanceof MyGrantedAuthority) {
				MyGrantedAuthority userGrantedAuthority = (MyGrantedAuthority) ga;
				loginId = userGrantedAuthority.getLoginId();
				break;
			}
		}

		int is_purch = 0;
		List<Map<String, Object>> Deplist = (new purchaseFactory()).selectDepartment(loginId, jdbcTemplate1);

		for (Map<String, Object> str : Deplist) {
			System.out.println(str.get("CODE"));
			if ("T111".equals(str.get("CODE")) || "T110".equals(str.get("CODE")) || "T112".equals(str.get("CODE"))
					|| "V110".equals(str.get("CODE")) || "T190".equals(str.get("CODE"))) {
				is_purch = 1;
			}
		}
		model.addAttribute("is_purch", is_purch);

		return "purchase/psi";
	}

	@RequestMapping("psi_save")
	public String psi_save(HttpServletRequest request, Model model) {
		/*
		 * String d1 = request.getParameter("set_d1"); System.out.println("d1 = "+d1);
		 */
		String d2 = request.getParameter("set_d2");
		String d3 = request.getParameter("set_d3");
		String d4 = request.getParameter("set_d4");
		String d5 = request.getParameter("set_d5");
		String d6 = request.getParameter("set_d6");
		String ps = request.getParameter("ps");

		int j = 0;
		// ps
		j = (new purchaseFactory()).SelectIsValue1(request, jdbcTemplate1);
		if (j == 0) {

			j = (new purchaseFactory()).insertValue1(request, jdbcTemplate1);
		} else {

			j = (new purchaseFactory()).updateValue1(request, jdbcTemplate1);
		}

		for (int i = 2; i <= 6; i++) {
			j = (new purchaseFactory()).SelectIsValue(request, i, jdbcTemplate1);
			if (j == 0) {
				// insert

				j = (new purchaseFactory()).insertValue(request, i, jdbcTemplate1);

			} else {// update

				j = (new purchaseFactory()).updateValue(request, i, jdbcTemplate1);
			}
		}

		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, 0);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String time = format.format(c.getTime());
		model.addAttribute("month_0", time);

		for (int i = 1; i <= 5; i++) {
			c = Calendar.getInstance();
			c.add(Calendar.MONTH, i);
			format = new SimpleDateFormat("yyyy-MM");
			time = format.format(c.getTime());

			model.addAttribute("month_" + i, time);
		}

		List<Map<String, Object>> IMA = (new purchaseFactory()).selectIMA(request, jdbcTemplate1);
		model.addAttribute("IMA", IMA);

		// 取品牌
		List<Map<String, Object>> tqas = (new purchaseFactory()).selectTQA(jdbcTemplate1);

		model.addAttribute("tqas", tqas);


		// 供應商
		List<Map<String, Object>> pmcblist = (new reportFactory()).selectPMC(jdbcTemplate1);
		model.addAttribute("pmclist", pmcblist);

		// 產品狀態
		List<Map<String, Object>> tqalist = (new reportFactory()).selectTQA(jdbcTemplate1);
		model.addAttribute("tqalist", tqalist);
		// 採購大類
		List<Map<String, Object>> imzlist = (new reportFactory()).selectIMZ(jdbcTemplate1);
		model.addAttribute("imzlist", imzlist);
		// 商品大類
		List<Map<String, Object>> obalist = (new reportFactory()).selectOBA(jdbcTemplate1);
		model.addAttribute("obalist", obalist);
		
		String find_tqa = request.getParameter("find_tqa");
		model.addAttribute("find_tqa", find_tqa);

		String find_IMA01 = request.getParameter("find_IMA01");
		model.addAttribute("find_IMA01", find_IMA01);

		String order_is = request.getParameter("order_is");
		model.addAttribute("order_is", order_is);

		String can_get_order = request.getParameter("can_get_order");
		model.addAttribute("can_get_order", can_get_order);


		//取登入者部門
				String loginId = "";
				Authentication auth = SecurityContextHolder.getContext().getAuthentication();
				for (GrantedAuthority ga : auth.getAuthorities()) {
					if (ga instanceof MyGrantedAuthority) {
						MyGrantedAuthority userGrantedAuthority = (MyGrantedAuthority) ga;
						loginId = userGrantedAuthority.getLoginId();
						break;
					}
				}

				int is_purch = 0;
				List<Map<String, Object>> Deplist = (new purchaseFactory()).selectDepartment(loginId, jdbcTemplate1);

				for (Map<String, Object> str : Deplist) {
					System.out.println(str.get("CODE"));
					if ("T111".equals(str.get("CODE")) || "T110".equals(str.get("CODE")) || "T112".equals(str.get("CODE"))
							|| "V110".equals(str.get("CODE")) || "T190".equals(str.get("CODE"))) {
						is_purch = 1;
					}
				}
				model.addAttribute("is_purch", is_purch);

				
		return "purchase/psi";
	}

	@RequestMapping("psi_down")
	public void psi_down(HttpServletRequest request, Model model, HttpServletResponse response) {

		try {

			List<Map<String, Object>> IMA = new ArrayList<Map<String, Object>>();
			IMA = (new purchaseFactory()).selectIMA(request, jdbcTemplate1);

			// 設定下載檔案路徑及名稱
			String sysDate = DateUtil.getDate().replaceAll("/", "");
			String filePath = "C:/G25/downloadfile/";
			File newDirectory = new File("C:/G25/", "downloadfile");
			if (!newDirectory.exists())
				newDirectory.mkdirs();

			String fileName = "PSI.xlsx";
			// 產生暫存檔
			(new ExcelWrite()).get_psi_down(filePath + fileName, IMA);

			// 下載Excel檔案
			DownloadUtility download = new DownloadUtility(filePath, fileName);
			download.setDisplayFileName("PSI_" + sysDate + ".xlsx");
			download.downloadFile(response);

			// 刪除暫存檔
			File file = new File(filePath, fileName);
			file.delete();

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}
}
